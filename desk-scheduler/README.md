# Desk Coverage Scheduler

A small, offline desktop app that turns a spreadsheet of employees' requested
shifts into a fair, balanced daily desk schedule -- and exports it as a
printable, editable Excel calendar.

Built for a real workplace need: turning a Kronos schedule export into
coverage for four service desks (**East, West, VAB, Resource**) without
manual spreadsheet wrangling, and without any of the data leaving the
building.

## Why this exists

The scheduling manager was doing this by hand every period: read through
everyone's requested shifts, make sure East and West (the two priority
desks) were always covered, then fill VAB, then Resource, and try to keep
it fair so the same few people weren't always stuck with the same desk.
This app automates that process end-to-end while keeping a human in the
loop for anything it can't safely decide on its own.

## What it does

1. **Import** -- reads Kronos's actual "Location Schedule - Monthly" `.xlsx`
   export directly (the wide day-by-day grid, not a reformatted table --
   see "Input format" below). Kronos has no concept of desk assignment at
   all; it only knows who's working which shift on which day. Deciding
   which of the four desks each working person goes on is what this app adds.
2. **Balance** -- for each day in the chosen period, fills East, West, VAB,
   then Resource, in that priority order, from whoever requested that day.
   People who asked for a specific desk are only ever placed there; people
   with no preference form a flexible pool that fills whatever is still
   open. If a desk can't be covered, it's left blank and flagged instead of
   guessing.
3. **Balance fairly** -- when more than one person is eligible for a slot,
   the app prefers whoever has worked the fewest shifts that week (then
   fewest overall), so coverage doesn't fall on the same people every time.
4. **Export** -- writes a monthly calendar-grid workbook (one sheet per
   month in the period), plus a Summary sheet totaling shifts per person
   and a "Needs Attention" sheet listing anything unfilled.

## Why Python + Tkinter

* **No install friction.** Tkinter ships with the Python standard library,
  so there's nothing extra to install beyond Python itself (see the
  troubleshooting note below for Linux).
* **No network code at all.** The app never opens a socket. It reads one
  local file and writes one local file. That's the entire attack surface.
* **Runs on a locked-down work laptop.** No admin rights or IT ticket
  needed to install a GUI framework.

## Security / robustness notes

This was designed to be handed to a non-technical manager with no support
on hand, so failure modes matter as much as the happy path:

* Only `.xlsx` is accepted; macro-enabled `.xlsm` files are rejected
  outright.
* Workbooks are opened with `data_only=True` so formulas are never
  evaluated -- only their last computed value (or nothing) is read.
* Every cell value is validated before use: names must match a plain
  human-name pattern, dates must parse cleanly, desk values must match a
  known desk or are treated as "no preference" with a warning. Nothing
  that looks like a formula (`=...`) is ever accepted as a name or date.
* File size and row count are capped so a corrupted or oversized file
  can't hang the app.
* Bad rows are skipped and reported individually -- one bad row never
  aborts the whole import.
* No unhandled exception reaches the user: file, scheduling, and export
  errors are caught and shown as a plain-English message, with details
  logged to `app.log` for debugging.

## Project layout

```
desk-scheduler/
├── main.py                    # entry point
├── core/
│   ├── models.py               # Employee, ShiftRequest, DeskAssignment, etc.
│   ├── importer.py             # validated .xlsx -> ShiftRequest list
│   ├── scheduler.py            # the balancing algorithm
│   └── exporter.py             # SchedulePeriod -> printable .xlsx calendar
├── gui/
│   └── app.py                  # Tkinter UI
├── sample_data/
│   └── generate_sample.py      # generates a fake Kronos export for demos/tests
├── tests/                      # pytest unit tests for importer/scheduler/exporter
└── requirements.txt
```

The core logic (`core/`) has zero dependency on the GUI, so it's fully
unit-testable without a display -- important since this environment can't
run a Tkinter window in CI.

## Running it

```bash
pip install -r requirements.txt
python main.py
```

On the first screen, click **"Use Sample Data"** to try the whole pipeline
with realistic fake data (no real spreadsheet needed), or **"Browse..."**
to load a real Kronos export. Pick a date range, click **Run Balance**, and
review the preview table -- anything highlighted red is a desk that
couldn't be covered that day. Click **Export Calendar...** to save the
final workbook.

### Input format

This was rebuilt against a real Kronos "Location Schedule - Monthly" export
(`_Report Output_Location_Schedule_Monthly_..._.xlsx`), which is not a
simple table -- it's a wide grid:

* A row containing the literal cell text **"Employee"** marks the start of
  a block. The row directly below it holds one date per day, each spread
  across 2-3 merged columns for visual width.
* Each employee gets one row. For each date's column span, there's at most
  one non-blank cell, containing one of:
  * a **shift code**, e.g. `D12` or `08c12` -- the person is working that
    shift that day (the code itself is just carried through to the output
    calendar for reference; it currently doesn't affect which desk they
    get)
  * **`PTO ...`** -- approved time off; not available that day
  * **`RDY`, `SDY`, or `VDY 12:30`** alone -- a soft "please don't schedule
    me this day if possible" request, honored (treated as not working)
  * a marker plus a shift code, e.g. `RDY \nD12` -- the soft request was
    overridden and they're working that shift anyway
  * blank -- not requesting to work that day
* A full export can contain more than one such block stacked in the sheet
  (one per Kronos location/cost-center code). The importer detects and
  parses every block it finds, independently.

The Job column (constant per employee, e.g. `ISS`) is read but not used --
it reflects role/title, not a desk. Preferred-desk parsing from the earlier
flat-table design is still supported for a manually-built request list, but
real Kronos exports won't populate it, so in practice every real request
lands in the flexible pool the priority-fill logic draws from.

If a future export doesn't match this shape (different report type, a
column layout change, etc.), `core/importer.py` is the one place to adjust
-- `_find_header_rows`, `_find_date_spans`, and `parse_day_cell` are the
three functions that encode the assumptions above.

**Open questions to confirm against a full, multi-location real pull** (this
was validated against a real single-location export, but a few things are
still assumptions rather than confirmed facts):
* Whether a full pull actually stacks multiple location blocks in one sheet
  (supported, but untested against a real multi-block file).
* Whether shift type (`D12` vs `08c12`) should ever constrain which desk
  someone can be assigned to, or whether "worked that day at all" is
  sufficient input for the balancer (currently the latter).
* The exact list of avoid-marker codes -- `RDY`/`SDY`/`VDY` were what
  appeared in the sample pull; there may be others.

## Running the tests

```bash
pip install -r requirements.txt
pytest -v
```

Tests cover: import validation against synthetic Kronos-grid workbooks
(shift codes, avoid-markers alone vs. overridden, PTO, blank cells,
multiple stacked location blocks, invalid names, macro-file rejection,
missing "Employee" header), scheduling invariants (priority fill order,
one desk per person per day, never more than 4 people per day, explicit
preferences are respected rather than overridden, fairness rotation stays
balanced across a week), and export structure (correct sheets, correct
month spans). The importer has also been run end-to-end against a real
Children's Colorado export, not just synthetic test data.

## Known limitations / next steps

* The "~3 shifts per week" target is currently just a fairness
  tie-breaker, not a hard cap -- easy to tighten in `core/scheduler.py`
  once real usage shows whether that's needed.
* No persistence between runs yet (each run is import -> balance -> export
  in one pass). A "remember last month's totals" feature would improve
  fairness across periods rather than resetting each time.
* Tkinter needs a system package on some Linux distros
  (`sudo apt install python3-tk`) -- not needed on Windows or macOS, where
  it ships with the standard Python installer.
