"""
Writes a SchedulePeriod out to a printable, editable Excel workbook:

  * One monthly calendar-grid sheet per calendar month covered by the
    period (traditional Sun-Sat boxes, one box per day, listing who's on
    each desk).
  * A "Summary" sheet totaling shifts per person, for a quick fairness
    check.
  * A "Shortfalls" sheet listing any desk/day combos that couldn't be
    filled, so the manager knows exactly what needs manual attention.

Everything is plain cell values and styling -- no formulas, no macros --
so the output is safe to open, edit, and reprint by hand.
"""

from __future__ import annotations

import calendar
from collections import defaultdict
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from core.models import Desk, SchedulePeriod

_HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
_WEEKDAY_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
_OPEN_FILL = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
_THIN = Side(style="thin", color="AAAAAA")
_BOX_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_WEEKDAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]

_DAY_ROW_HEIGHT = 78
_DAY_COL_WIDTH = 20


def _months_in_period(start: date, end: date) -> list[tuple[int, int]]:
    months = []
    year, month = start.year, start.month
    while (year, month) <= (end.year, end.month):
        months.append((year, month))
        if month == 12:
            year, month = year + 1, 1
        else:
            month += 1
    return months


def _write_month_sheet(
    wb: Workbook,
    year: int,
    month: int,
    period: SchedulePeriod,
) -> None:
    sheet_name = date(year, month, 1).strftime("%b %Y")
    ws: Worksheet = wb.create_sheet(title=sheet_name)

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    title = date(year, month, 1).strftime("%B %Y") + " - Desk Coverage Schedule"
    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = title
    title_cell.font = Font(size=16, bold=True, color="FFFFFF")
    title_cell.fill = _HEADER_FILL
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for col, name in enumerate(_WEEKDAY_NAMES, start=1):
        cell = ws.cell(row=2, column=col, value=name)
        cell.font = Font(bold=True)
        cell.fill = _WEEKDAY_FILL
        cell.alignment = Alignment(horizontal="center")
        cell.border = _BOX_BORDER

    for col in range(1, 8):
        ws.column_dimensions[get_column_letter(col)].width = _DAY_COL_WIDTH

    month_calendar = calendar.Calendar(firstweekday=6).monthdatescalendar(year, month)

    row_idx = 3
    for week in month_calendar:
        ws.row_dimensions[row_idx].height = _DAY_ROW_HEIGHT
        for col_idx, day in enumerate(week, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = _BOX_BORDER
            cell.alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=True
            )

            if day.month != month:
                cell.value = ""
                cell.fill = PatternFill(
                    start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
                )
                continue

            lines = [str(day.day)]
            day_assignments = {a.desk: a for a in period.assignments_for_day(day)}
            has_open = False
            for desk in Desk.priority_order():
                assignment = day_assignments.get(desk)
                if assignment is not None and assignment.employee is not None:
                    label = f"{desk.value}: {assignment.employee.name}"
                    if assignment.shift_code:
                        label += f" ({assignment.shift_code})"
                    lines.append(label)
                elif assignment is not None:
                    lines.append(f"{desk.value}: (OPEN)")
                    has_open = True
            cell.value = "\n".join(lines)
            if has_open:
                cell.fill = _OPEN_FILL
        row_idx += 1

    ws.freeze_panes = "A3"


def _write_summary_sheet(wb: Workbook, period: SchedulePeriod) -> None:
    ws = wb.create_sheet(title="Summary")
    ws.append(["Employee", "Total Shifts", "East", "West", "VAB", "Resource"])
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = _WEEKDAY_FILL

    per_person: dict[str, dict[str, int]] = defaultdict(
        lambda: {"total": 0, **{d.value: 0 for d in Desk}}
    )
    for a in period.assignments:
        if a.employee is None:
            continue
        stats = per_person[a.employee.name]
        stats["total"] += 1
        stats[a.desk.value] += 1

    for name in sorted(per_person):
        stats = per_person[name]
        ws.append(
            [
                name,
                stats["total"],
                stats["East"],
                stats["West"],
                stats["VAB"],
                stats["Resource"],
            ]
        )

    for col in range(1, 7):
        ws.column_dimensions[get_column_letter(col)].width = 16


def _write_shortfalls_sheet(wb: Workbook, period: SchedulePeriod) -> None:
    if not period.shortfalls:
        return
    ws = wb.create_sheet(title="Needs Attention")
    ws.append(["Date", "Desk", "Reason"])
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = _OPEN_FILL
    for flag in sorted(period.shortfalls, key=lambda f: (f.work_date, f.desk.value)):
        ws.append([flag.work_date.strftime("%Y-%m-%d"), flag.desk.value, flag.reason])
    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 24


def export_schedule(period: SchedulePeriod, output_path: str | Path) -> Path:
    output_path = Path(output_path)
    wb = Workbook()
    wb.remove(wb.active)  # drop the default blank sheet

    _write_shortfalls_sheet(wb, period)
    _write_summary_sheet(wb, period)
    for year, month in _months_in_period(period.start, period.end):
        _write_month_sheet(wb, year, month, period)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
