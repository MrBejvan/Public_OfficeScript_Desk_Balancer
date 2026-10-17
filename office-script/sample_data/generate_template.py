"""
Generates the ready-to-use Desk Coverage Scheduler workbook: a "Report"
sheet shaped like a real Kronos "Location Schedule - Monthly" export (a
wide day-by-day grid with shift codes, not a plain table -- see README.md
for the full format description), a "Settings" sheet with the schedule
period + status cell, and a "Read Me" sheet with quick setup pointers.

In practice you don't have to use this generator at all: you can instead
take a real Kronos export, add a "Settings" sheet to it (matching the
layout below), and use that directly. This script exists so there's a
working demo file without needing a real export on hand.

Upload the output of this script to OneDrive/SharePoint, then follow
README.md to attach the Office Script.

Run with:  python generate_template.py
"""

from __future__ import annotations

import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

SAMPLE_EMPLOYEES = [
    "Alvarez, Jordan",
    "Brooks, Taylor",
    "Castillo-Reyes, Morgan",
    "Dalton, Casey",
    "Ellison, Avery",
    "Foster, Quinn",
    "Grant, Riley",
    "Hayes, Jamie",
    "Irving, Drew",
]

SHIFT_CODES = ["D12", "08c12"]
SHIFT_WEIGHTS = [0.75, 0.25]

NON_WORK_MARKERS = [None, "RDY", "SDY", "VDY 12:30", "VDY 00:30", "PTO Scheduled Hours 12:00"]
NON_WORK_WEIGHTS = [0.70, 0.10, 0.08, 0.05, 0.02, 0.05]

# Sample employment statuses -- mostly Full Time (3 shifts/wk), a couple
# Part Time (2/wk), one Flex (1/wk), one 3-2 (biweekly: 3 shifts one week,
# 2 the next), mirroring a realistic small-team mix.
SAMPLE_STATUSES = {
    "Alvarez, Jordan": "Full Time",
    "Brooks, Taylor": "Full Time",
    "Castillo-Reyes, Morgan": "Part Time",
    "Dalton, Casey": "Full Time",
    "Ellison, Avery": "Flex",
    "Foster, Quinn": "3-2",
    "Grant, Riley": "Part Time",
    "Hayes, Jamie": "Full Time",
    "Irving, Drew": "Full Time",
}

LABEL_FONT = Font(bold=True)


def _weekday_header(d: date) -> str:
    return d.strftime("%a")


def _write_report_sheet(
    ws,
    start: date,
    end: date,
    employees: list[str],
    shifts_per_week: int,
    seed: int,
    include_sample_rows: bool,
) -> None:
    rng = random.Random(seed)

    all_days: list[date] = []
    day = start
    while day <= end:
        all_days.append(day)
        day += timedelta(days=1)

    weeks: dict[tuple[int, int], list[date]] = {}
    for d in all_days:
        key = d.isocalendar()[:2]
        weeks.setdefault(key, []).append(d)

    ws.cell(row=1, column=1, value="Location Schedule - Monthly").font = Font(bold=True, size=14)
    ws.cell(row=2, column=1, value="Time Period :")
    ws.cell(
        row=2,
        column=4,
        value=f"{start.month}/{start.day:02d}/{start.year} - {end.month}/{end.day:02d}/{end.year}",
    )
    ws.cell(row=5, column=1, value="0000100-X1 - Location: SAMPLE/DEMO/DATA/0000100-X1")

    header_row = 6
    date_row = 7
    ws.cell(row=header_row, column=1, value="Employee")
    ws.cell(row=header_row, column=3, value="Job")

    col_for_date: dict[date, int] = {}
    col = 5
    for d in all_days:
        ws.cell(row=header_row, column=col, value=_weekday_header(d))
        ws.cell(row=date_row, column=col, value=d.strftime("%m/%d/\n%Y"))
        col_for_date[d] = col
        col += 2

    row = date_row + 1
    for emp in employees:
        ws.cell(row=row, column=1, value=emp)
        ws.cell(row=row, column=3, value="ISS")

        working_days: set[date] = set()
        if include_sample_rows:
            for _week_key, days_in_week in weeks.items():
                pick_count = min(shifts_per_week, len(days_in_week))
                working_days.update(rng.sample(days_in_week, pick_count))

        for d in all_days:
            c = col_for_date[d]
            if d in working_days:
                shift = rng.choices(SHIFT_CODES, weights=SHIFT_WEIGHTS, k=1)[0]
                if rng.random() < 0.08:
                    marker = rng.choice(["RDY", "SDY"])
                    value = f"{marker} \n{shift}"
                else:
                    value = shift
            else:
                value = rng.choices(NON_WORK_MARKERS, weights=NON_WORK_WEIGHTS, k=1)[0]
            if value:
                ws.cell(row=row, column=c, value=value)
        row += 1

    for col_idx in range(1, col):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = 14


def generate_template(
    output_path: str | Path,
    start: date,
    end: date,
    employees: list[str] | None = None,
    shifts_per_week: int = 3,
    seed: int = 42,
    include_sample_rows: bool = True,
) -> Path:
    output_path = Path(output_path)
    employees = employees or SAMPLE_EMPLOYEES

    wb = Workbook()

    # ---- Report sheet (Kronos-grid-shaped) ----
    report_sheet = wb.active
    report_sheet.title = "Report"
    _write_report_sheet(report_sheet, start, end, employees, shifts_per_week, seed, include_sample_rows)

    # ---- Employees sheet (roster: Name -> Full Time / Part Time / Flex) ----
    employees_sheet = wb.create_sheet("Employees")
    employees_sheet.cell(row=1, column=1, value="Name").font = LABEL_FONT
    employees_sheet.cell(row=1, column=2, value="Status").font = LABEL_FONT
    for i, emp in enumerate(employees, start=2):
        employees_sheet.cell(row=i, column=1, value=emp)
        status = SAMPLE_STATUSES.get(emp, "Full Time")
        employees_sheet.cell(row=i, column=2, value=status)
    employees_sheet.column_dimensions["A"].width = 28
    employees_sheet.column_dimensions["B"].width = 14

    # ---- Settings sheet ----
    settings_sheet = wb.create_sheet("Settings")
    settings_sheet["A1"] = "Start Date"
    settings_sheet["A1"].font = LABEL_FONT
    settings_sheet["B1"] = start
    settings_sheet["B1"].number_format = "MM/DD/YYYY"

    settings_sheet["A2"] = "End Date"
    settings_sheet["A2"].font = LABEL_FONT
    settings_sheet["B2"] = end
    settings_sheet["B2"].number_format = "MM/DD/YYYY"

    settings_sheet["A4"] = "Status"
    settings_sheet["A4"].font = LABEL_FONT
    settings_sheet["B4"] = "Not run yet. Set the dates above, then run the BalanceDeskSchedule script from the Automate tab."
    settings_sheet["B4"].alignment = Alignment(wrap_text=True, vertical="top")

    settings_sheet.column_dimensions["A"].width = 14
    settings_sheet.column_dimensions["B"].width = 70
    settings_sheet.row_dimensions[4].height = 32

    # ---- Read Me sheet ----
    readme_sheet = wb.create_sheet("Read Me")
    readme_lines = [
        ("Desk Coverage Scheduler", True, 14),
        ("", False, 11),
        ("KEEP AND REUSE THIS FILE -- don't create a new workbook each period.", True, 12),
        ("Kronos gives you a new export every time, but only its data moves into this file;", False, 11),
        ("the script and setup here stay put.", False, 11),
        ("", False, 11),
        ("Employees sheet -- keep this current:", True, 11),
        ("Every person must be listed on the Employees sheet with a Status: Full Time (3/week),", False, 11),
        ("Part Time (2/week), Flex (1/week), or 3-2 (biweekly -- 3 shifts one week, 2 the next).", False, 11),
        ("This is a hard cap -- once someone hits their weekly target the script stops assigning", False, 11),
        ("them for the rest of that week, even if a desk is left open. Anyone in Kronos but NOT on", False, 11),
        ("this sheet is skipped entirely (flagged on Needs Attention) until you add them. Approved", False, 11),
        ("PTO counts toward the weekly target; RDY/SDY/VDY 'please don't schedule me' markers do not.", False, 11),
        ("", False, 11),
        ("Desk assignment is randomized, not a fixed formula: for each desk, whoever's had that", False, 11),
        ("specific desk least often this period is favored, with a random pick among anyone tied --", False, 11),
        ("so everyone rotates through all four desks, but the exact pattern differs run to run.", False, 11),
        ("", False, 11),
        ("Every scheduling period:", True, 11),
        ("1. Export 'Location Schedule - Monthly' from Kronos -- downloads a new .xlsx file.", False, 11),
        ("2. Open it, select all on its Report sheet, copy. Come back here, select all on THIS", False, 11),
        ("   workbook's Report sheet, delete, paste at A1.", False, 11),
        ("3. Settings sheet: set Start Date / End Date for the new period.", False, 11),
        ("4. Update the Employees sheet if anyone was hired, left, or changed status.", False, 11),
        ("5. Automate tab > BalanceDeskSchedule > Run. Status appears in Settings!B4.", False, 11),
        ("6. Results appear as new sheets: one printable calendar per month, plus Summary and Needs Attention.", False, 11),
        ("", False, 11),
        ("Re-running replaces the previously generated sheets -- it never piles up duplicates.", False, 11),
        ("Full setup instructions (including the optional Power Automate flow) are in README.md.", False, 11),
    ]
    for i, (text, bold, size) in enumerate(readme_lines, start=1):
        cell = readme_sheet.cell(row=i, column=1, value=text)
        cell.font = Font(bold=bold, size=size)
    readme_sheet.column_dimensions["A"].width = 110

    # Sheet order: Read Me, Settings, Employees, Report
    wb._sheets.sort(
        key=lambda s: ["Read Me", "Settings", "Employees", "Report"].index(s.title)
        if s.title in ["Read Me", "Settings", "Employees", "Report"]
        else 99
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    today = date.today()
    period_start = today.replace(day=1)
    if period_start.month == 12:
        next_month_start = period_start.replace(year=period_start.year + 1, month=1)
    else:
        next_month_start = period_start.replace(month=period_start.month + 1)
    period_end = next_month_start - timedelta(days=1)

    out = generate_template(
        Path(__file__).parent / "Desk_Schedule_Template.xlsx",
        period_start,
        period_end,
    )
    print(f"Template written to {out}")
