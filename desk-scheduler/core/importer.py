"""
Reads a Kronos "Location Schedule - Monthly" export (.xlsx) and turns it
into a list of validated ShiftRequest objects.

This is NOT a simple table -- it's Kronos's actual report layout, reverse
engineered from a real export:

  * Row containing the literal text "Employee" starts a block. The row
    right below it holds the dates for that block, one non-blank cell per
    day (the report merges each date across 2-3 columns for visual width,
    but only the left-most cell of a merge actually holds a value -- so
    "which columns belong to today" is figured out by finding the *next*
    non-blank date cell, without needing to inspect Excel's merge ranges
    directly).
  * Each employee row has, for each date's column span, at most one
    non-blank cell (verified against a real export). Its text is one of:
      - blank                          -> not requesting to work
      - a shift code, e.g. "D12"       -> working that shift
      - "PTO ..."                      -> approved time off, not working
      - "RDY" / "SDY" / "VDY 12:30"    -> a soft "please don't schedule me
                                           this day if possible" marker,
                                           honored (not working) when it
                                           appears alone
      - a marker + a shift code on two lines, e.g. "RDY\\nD12" -> the soft
        request was overridden and they *are* working that shift
  * A real export can contain more than one such block stacked in the same
    sheet (one per Kronos location/cost-center code). Each block is parsed
    independently and its employee rows stop at the next block's header (or
    a "... - Location: ..." banner row) rather than running to the bottom
    of the sheet.

Kronos has no concept of desk assignment at all -- there's a "Job" column
but it reflects the employee's role/title, not a desk -- so preferred_desk
is always None for real data. Deciding which of the four desks each
working person goes on is exactly what the scheduler does next.

Design goals carried over from the original flat-table importer:
  * Never raise an unhandled exception into the GUI -- bad rows/cells are
    collected as issues and reported, the rest of the file still imports.
  * Never trust spreadsheet formulas -- cells are read with data_only=True.
  * Reject macro-enabled workbooks (.xlsm) outright.
  * Bound the amount of data read so a corrupt or hostile file can't hang
    the app.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from core.models import Employee, ShiftRequest

MAX_FILE_SIZE_MB = 20
MAX_EMPLOYEE_ROWS = 5000

# "Please don't schedule me this day if possible" markers. Case-insensitive,
# optionally followed by a time like "12:30" (seen in real exports as
# "VDY 12:30"). Honored (treated as not-working) unless a shift code is
# also present in the same cell, in which case the shift wins.
_AVOID_MARKER_RE = re.compile(r"\b(RDY|SDY|VDY)\b\s*(\d{1,2}:\d{2})?", re.IGNORECASE)
_PTO_RE = re.compile(r"^pto\b", re.IGNORECASE)

_DATE_TOKEN_RE = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")
_NAME_PATTERN = re.compile(r"^[A-Za-z][A-Za-z ,.'\-]{0,79}$")

_HEADER_LABEL = "employee"
_LOCATION_BANNER_MARKER = "location:"


class KronosImportError(Exception):
    """Raised for file-level problems that stop the import entirely."""


@dataclass
class RowIssue:
    rowNumber: int  # kept for interface parity with the rest of the app
    message: str


@dataclass
class ImportResult:
    requests: list[ShiftRequest]
    issues: list[RowIssue]

    @property
    def ok(self) -> bool:
        return len(self.requests) > 0


def _collapse_whitespace(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def _strip_all_whitespace(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value))


def _clean_name(raw: object) -> str | None:
    text = _collapse_whitespace(raw)
    if not text or text.startswith("="):
        return None
    if not _NAME_PATTERN.match(text):
        return None
    return text


def parse_day_cell(raw: object) -> tuple[bool, str | None]:
    """
    Interprets one grid cell for a single employee/day.

    Returns (working, shift_code). working=False means the person is not
    requesting/scheduled to work that day (blank cell, PTO, or an
    avoid-marker with no shift code attached).
    """
    text = _collapse_whitespace(raw)
    if not text:
        return False, None
    if _PTO_RE.match(text):
        return False, None
    remainder = _AVOID_MARKER_RE.sub("", text).strip()
    if not remainder:
        return False, None
    return True, remainder


def _parse_date_header(raw: object) -> date | None:
    text = _strip_all_whitespace(raw)
    m = _DATE_TOKEN_RE.match(text)
    if not m:
        return None
    month, day, year = (int(g) for g in m.groups())
    try:
        return date(year, month, day)
    except ValueError:
        return None


def _find_header_rows(sheet: Worksheet) -> list[tuple[int, int]]:
    """Returns (row, column) for every cell that reads exactly 'Employee'."""
    headers: list[tuple[int, int]] = []
    for row in range(1, sheet.max_row + 1):
        for col in range(1, sheet.max_column + 1):
            value = sheet.cell(row=row, column=col).value
            if _collapse_whitespace(value).lower() == _HEADER_LABEL:
                headers.append((row, col))
                break
    return headers


def _find_date_spans(sheet: Worksheet, date_row: int, max_col: int) -> list[tuple[int, int, date | None, str]]:
    """
    Finds every non-blank cell in `date_row` (each marks the start of one
    day's column span) and returns (start_col, end_col, parsed_date,
    raw_text) tuples covering the whole row.
    """
    starts: list[tuple[int, object]] = []
    for col in range(1, max_col + 1):
        value = sheet.cell(row=date_row, column=col).value
        if value not in (None, ""):
            starts.append((col, value))

    spans: list[tuple[int, int, date | None, str]] = []
    for i, (col, value) in enumerate(starts):
        end_col = starts[i + 1][0] - 1 if i + 1 < len(starts) else max_col
        spans.append((col, end_col, _parse_date_header(value), _collapse_whitespace(value)))
    return spans


def _is_location_banner(value: object) -> bool:
    return _LOCATION_BANNER_MARKER in _collapse_whitespace(value).lower()


def _parse_block(
    sheet: Worksheet,
    header_row: int,
    name_col: int,
    block_end_row: int,
    max_col: int,
    issues: list[RowIssue],
) -> list[ShiftRequest]:
    date_row = header_row + 1
    date_spans = _find_date_spans(sheet, date_row, max_col)
    if not date_spans:
        issues.append(RowIssue(header_row, "No date columns found under the 'Employee' header row -- skipped this block."))
        return []

    bad_dates = [text for (_, _, d, text) in date_spans if d is None]
    if bad_dates:
        issues.append(
            RowIssue(
                date_row,
                f"{len(bad_dates)} date header(s) could not be parsed and were skipped: {', '.join(bad_dates[:5])}"
                + (" ..." if len(bad_dates) > 5 else ""),
            )
        )

    requests: list[ShiftRequest] = []
    row = date_row + 1
    rows_seen = 0
    while row < block_end_row:
        rows_seen += 1
        if rows_seen > MAX_EMPLOYEE_ROWS:
            issues.append(RowIssue(row, f"Stopped reading after {MAX_EMPLOYEE_ROWS} employee rows (safety limit)."))
            break

        name_raw = sheet.cell(row=row, column=name_col).value
        if name_raw in (None, "") or _is_location_banner(name_raw):
            row += 1
            continue

        name = _clean_name(name_raw)
        if name is None:
            issues.append(RowIssue(row, f"Skipped row: invalid or unreadable employee name ({name_raw!r})."))
            row += 1
            continue

        for start_col, end_col, work_date, _raw_header in date_spans:
            if work_date is None:
                continue
            cell_value = None
            for col in range(start_col, end_col + 1):
                v = sheet.cell(row=row, column=col).value
                if v not in (None, ""):
                    cell_value = v
                    break
            working, shift_code = parse_day_cell(cell_value)
            if working:
                requests.append(
                    ShiftRequest(
                        employee=Employee(name=name),
                        work_date=work_date,
                        preferred_desk=None,
                        shift_code=shift_code or "",
                    )
                )
        row += 1

    return requests


def load_shift_requests(path: str | Path) -> ImportResult:
    path = Path(path)

    if not path.exists():
        raise KronosImportError(f"File not found: {path}")

    if path.suffix.lower() == ".xlsm":
        raise KronosImportError(
            "Macro-enabled workbooks (.xlsm) are not accepted for security "
            "reasons. Please save the file as a plain .xlsx and try again."
        )
    if path.suffix.lower() != ".xlsx":
        raise KronosImportError(f"Unsupported file type '{path.suffix}'. Please provide a .xlsx file.")

    size_mb = path.stat().st_size / (1024 * 1024)
    if size_mb > MAX_FILE_SIZE_MB:
        raise KronosImportError(
            f"File is {size_mb:.1f} MB, which is larger than the {MAX_FILE_SIZE_MB} MB limit. "
            "Please check the file."
        )

    try:
        workbook = openpyxl.load_workbook(path, data_only=True, keep_vba=False)
    except Exception as exc:  # noqa: BLE001 - surface as a friendly import error
        raise KronosImportError(
            "Could not open this file as an Excel workbook. It may be corrupted or password-protected."
        ) from exc

    sheet = None
    for candidate in workbook.worksheets:
        if "report" in candidate.title.lower():
            sheet = candidate
            break
    if sheet is None:
        sheet = workbook.active

    header_rows = _find_header_rows(sheet)
    if not header_rows:
        workbook.close()
        raise KronosImportError(
            "Could not find a Kronos 'Location Schedule' block in this file (no cell reading "
            "'Employee' was found). Make sure this is the report export, not a different file."
        )

    issues: list[RowIssue] = []
    requests: list[ShiftRequest] = []

    for i, (header_row, name_col) in enumerate(header_rows):
        block_end_row = header_rows[i + 1][0] if i + 1 < len(header_rows) else sheet.max_row + 1
        requests.extend(_parse_block(sheet, header_row, name_col, block_end_row, sheet.max_column, issues))

    workbook.close()
    return ImportResult(requests=requests, issues=issues)
