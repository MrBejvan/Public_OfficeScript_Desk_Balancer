# Desk Coverage Scheduler -- Office Scripts / Power Automate Edition

Same tool, different foundation: this is a port of the standalone Python/
Tkinter Desk Coverage Scheduler to Office Scripts, so it can run inside
Microsoft 365 -- infrastructure your IT department already governs --
instead of as a foreign executable on a locked-down machine.

## Why a second version

The Python version is a complete, tested, air-gapped desktop app. It's
also, realistically, not something that will run on a fully locked-down
hospital-issued Windows image: standard-user machines in that kind of
environment typically block unsigned executables and often disable
command-line access outright, regardless of how the app is packaged.

Office Scripts sidesteps that problem entirely: it's TypeScript that runs
*inside Excel itself*, triggered by clicking a button on the ribbon or by
a Power Automate flow -- no install, no admin rights, no new executable on
the endpoint. The tradeoff is that data now lives in a OneDrive/SharePoint
workbook rather than a fully local file, which gives up the strict
air-gap property of the Python version. Given the hospital already trusts
Microsoft 365 for other automation, that's a reasonable trade: it's data
staying inside infrastructure IT already governs, not leaving anywhere new.

## What's identical to the Python version

The core rules were kept the same on purpose, so both versions produce a
similar schedule from the same input, and both were validated against the
same real Kronos export:

* Four desks -- East, West, VAB, Resource -- filled in that priority order.
* Desks that can't be covered are left blank and flagged, never guessed.
* Input validation: names must look like names, dates must parse cleanly,
  bad rows are skipped and reported individually rather than aborting the
  whole import.
* The importer parses Kronos's actual "Location Schedule - Monthly" export
  grid directly -- see "Workbook layout" below for the exact shape.

> **Status (this version is ahead of Python as of this writing):** this
> Office Script version has the Employees roster, weekly-target hard cap,
> and randomized fair desk rotation described below. The Python version
> doesn't have these yet -- porting them over is the planned next step, not
> yet started.

## How desks get assigned (randomized, not a fixed formula)

Deliberately not a deterministic "auto-balance" score: for each desk, among
everyone still eligible to work that day, whoever has been assigned *that
specific desk* the fewest times so far this period is favored; if several
people are tied for fewest, one is chosen at random. That's the whole rule.

This means: no two runs produce the identical pattern (since ties are
broken randomly), but everyone still rotates through all four desks over
time, because the algorithm never lets someone get a second turn at a desk
while someone else eligible hasn't had a first turn at it yet. East and
West are still filled before VAB and Resource on a short-staffed day, same
as before -- that priority order didn't change, only *who* fills each desk
once its turn comes up.

Total shift counts aren't directly balanced by this step anymore -- that's
now purely a side effect of each person's weekly target (below). If you
want everyone to end up with a similar number of total shifts, that's a
function of giving everyone the same status/target, not something the desk
picker tries to enforce on its own.

## Weekly shift targets (Full Time / Part Time / Flex / 3-2)

Kronos has no concept of employment status, so it's tracked by hand on the
**Employees** sheet (see "Workbook layout" below): every person gets a
Status of:

* **Full Time** -- 3 shifts every week
* **Part Time** -- 2 shifts every week
* **Flex** -- 1 shift every week
* **3-2** -- a biweekly status between Full Time and Part Time: 3 shifts
  one week, 2 shifts the next, alternating for the length of the pay
  period. Week 1 of whatever date range you enter in Settings is always
  treated as the "3" week, week 2 as the "2" week, and the pattern repeats
  if the period runs longer than two weeks.

This is a **hard cap**, not just a tie-breaker: once someone's worked
shifts plus approved PTO for a given week reach their target for that
specific week, the script stops assigning them to any desk for the rest of
that week -- even if that leaves a desk unfilled. An unfilled desk is left
blank and flagged on **Needs Attention** rather than over-scheduling
someone past their target.

Approved **PTO counts toward the weekly target** (it's still "spoken for"
time). The soft "please don't schedule me" markers -- **RDY / SDY / VDY**
-- do **not** count, since those are just requests, not confirmed absences.

**Anyone in the Kronos data who isn't listed on the Employees sheet is
skipped entirely** -- none of their requests are scheduled, and they're
listed by name in a single summary note on the Needs Attention sheet. Add
them to the Employees sheet with a status, then re-run, to have them
included. If the Employees sheet is missing from the workbook altogether,
the script stops with a clear error instead of guessing everyone's status.

The **Needs Attention** sheet also lists every roster employee-week where
worked shifts + PTO fell short of their target (resolved to the correct
week-specific number for 3-2 employees), so shortfalls caused by the hard
cap (or by someone simply not requesting enough days) are visible, not
silent. The **Summary** sheet shows each person's Status and Weekly Target
(e.g. "3" or "3-2") alongside their actual shift totals for the period.

## Workbook layout

The template (`sample_data/Desk_Schedule_Template.xlsx`) has four sheets:

**Employees** -- the roster the hard cap and unlisted-employee skip are
based on. Two columns, one header row:

| Name | Status |
| --- | --- |
| (as it appears in Kronos, e.g. `Alvarez, Jordan`) | `Full Time`, `Part Time`, `Flex`, or `3-2` (case-insensitive; `FT`/`PT`/`biweekly` also accepted) |

Keep this current -- add new hires before their first run, update anyone
whose status changes, and remove anyone who's left (or just leave them,
since a name with no matching Kronos rows simply produces no assignments).

**Report** -- shaped exactly like Kronos's own "Location Schedule -
Monthly" export (this is not a plain table; see the Python version's
README for the full reverse-engineered format description). In short:

* A row with the literal cell text **"Employee"** starts a block. The row
  below it holds one date per day (Kronos visually spreads each date
  across 2-3 columns; only the left-most cell of that span actually holds
  a value).
* Each employee gets one row, with each date's column span holding at most
  one value: a shift code (`D12`, `08c12`, ...) meaning they're working
  that day, `PTO ...` meaning approved time off, `RDY`/`SDY`/`VDY 12:30`
  alone meaning a soft "please don't schedule me" request that was
  honored, or a marker plus a shift code (`RDY \nD12`) meaning the soft
  request was overridden and they're working anyway.
* A full export can stack more than one such block in one sheet (one per
  location/cost-center code) -- each is parsed independently.

Kronos has no concept of desk assignment, so every real request lands in
the flexible pool the desk-assignment logic draws from -- deciding the
desk is what this script adds. In practice, you can skip the template's sample
data entirely and paste a real Kronos export's own sheet in here; the
script looks for a sheet named containing "report" first, and falls back
to scanning every sheet for an "Employee" cell if it doesn't find one.

**Settings** -- plain cells, not a table:

* `B1` = schedule period start date
* `B2` = schedule period end date
* `B4` = status message the script writes after every run (success, or a
  plain-English error -- this is the primary way you'll know whether a run
  worked, whether triggered manually or by Power Automate)

**Read Me** -- a one-page in-workbook summary of the steps below.

Running the script deletes and regenerates its own output sheets each
time (any "Summary," "Needs Attention," or month-named sheet), so re-runs
never pile up duplicates. It never touches the Report or Settings sheets.

## One-time setup

You do this section **once**. The workbook you set up here is reused every
scheduling period after that -- Kronos will give you a new export file each
time, but you never create a new workbook for it; only the data inside
the existing one gets refreshed (see "Every scheduling period" below).

1. **Upload the template.** Put `sample_data/Desk_Schedule_Template.xlsx`
   into a OneDrive for Business or SharePoint location, and treat this as
   your one permanent working file (e.g. name it something like
   `Desk Schedule - Working File.xlsx`). Office Scripts require the file
   to be stored there -- it won't work on a purely local file.

2. **Add the script.** Open the workbook in Excel (the web version, or
   desktop Excel signed into the same Microsoft 365 account). Go to the
   **Automate** tab, choose **New Script**, delete the placeholder code,
   and paste in the full contents of `BalanceDeskSchedule.ts`. Rename the
   script to `BalanceDeskSchedule` (the pencil icon next to the script
   name in the code editor). Save.

   If your Excel version supports importing script files directly, you
   can instead use **Automate > Code Editor > "..." > Import File** and
   select `BalanceDeskSchedule.ts` (rename it to `.osts` first if the file
   picker filters by that extension).

3. **(Optional) Pin a button.** Depending on your Excel/Microsoft 365
   version, the Automate tab lets you add a workbook button that runs the
   script from the ribbon or the sheet itself, so your manager never has
   to open the Code Editor. Look for an option like "Add Button" next to
   the script in the Automate pane. If that option isn't available in your
   tenant, running it from the Automate pane's **Run** button works
   identically -- it's just one extra click.

The template comes pre-loaded with sample data on the Report sheet so you
can try the whole flow immediately. Replace that with real data the first
time you run it for an actual period, following the steps below.

## Every scheduling period (the actual routine)

This is what your manager repeats each time a new period needs scheduling.
It always starts from a fresh Kronos export and ends by updating the
*same* working workbook from setup -- never a new upload.

1. **Export from Kronos.** Run the "Location Schedule - Monthly" report in
   Kronos for the upcoming period. This downloads a new `.xlsx` file (like
   the one used to build this importer) to your computer.

2. **Copy the data over.** Open the downloaded file, click its **Report**
   sheet tab, click the small triangle in the top-left corner (or press
   Ctrl+A) to select every cell, and copy (Ctrl+C).

   Switch to the working workbook (the one from setup, with the script
   attached), go to its **Report** sheet, select all and delete the
   existing contents (Ctrl+A, then Delete) -- this matters even though
   you're about to paste over most of it, because a shorter month could
   otherwise leave stray leftover cells from the previous period past the
   end of the new data. Click cell A1, then paste (Ctrl+V).

3. **Update the period.** On the **Settings** sheet, set `B1` (start date)
   and `B2` (end date) to the new period.

4. **Update the roster if needed.** On the **Employees** sheet, add anyone
   newly hired (with their Status) and update anyone whose Full Time/Part
   Time/Flex status changed. Anyone in the Kronos data but not listed here
   is skipped and flagged, not scheduled.

5. **Run it.** Click the pinned button, or Automate pane > Run, for
   `BalanceDeskSchedule`.

6. **Check the result.** `Settings!B4` shows a plain-English summary --
   how many desk-shifts were filled, how many are still open, how many
   employee-weeks fell under their target, and how many import notes there
   were. If anything's flagged, open the **Needs Attention** sheet for
   specifics.

7. **Use the calendar.** Open the generated month sheet(s) -- that's the
   printable, editable calendar. Print it, share the workbook link, or
   hand-edit any cell directly (it's a normal Excel sheet at that point).

No Power Automate is required for any of this -- it's the full workflow on
its own. Power Automate (below) is only useful if you want steps 4-5 to
happen automatically instead of by clicking a button.

## Optional: automating it with Power Automate

Note this only automates the *balancing* step (4-5 in the routine above),
not the manual copy of the new Kronos export into the Report sheet in
step 2 -- that data refresh still has to happen first, by hand, since it
comes from a separate file each period. What Power Automate can save you
is remembering to click Run afterward:

1. In [Power Automate](https://make.powerautomate.com), create a new flow.
2. **Trigger:** "Recurrence" -- e.g., every Monday at 6 AM, or whatever
   cadence follows your data-refresh routine.
3. **Action:** search for "Excel Online (Business)" and add **Run script**.
   Pick the workbook and select `BalanceDeskSchedule` as the script. No
   parameters need to be mapped -- the script reads everything it needs
   from the Settings sheet.
4. **(Optional) Action:** add "Office 365 Outlook -- Send an email (V2)"
   afterward to notify your manager that the schedule is ready, optionally
   including a link to the workbook.
5. Save and test the flow once manually (the "Test" button in Power
   Automate) before trusting the recurrence trigger.

Because all the actual logic lives in the Office Script and not in the
flow, the flow itself stays trivial -- just "run this script, then tell
someone" -- which avoids the exact problem that made a pure-Power-Automate
version impractical (trying to express the balancing algorithm as flow
actions).

## Testing

Office Scripts only run inside the Excel/Office Scripts runtime, so they
can't be unit tested the normal way. To still get real automated coverage,
every part of the logic that doesn't touch the Excel API -- date/serial
conversion, ISO week keying, the priority-fill + random fair-tier + weekly-
target hard-cap scheduler, *and* the Kronos grid parser (header/date-span
detection, the shift-code/PTO/avoid-marker cell interpretation) -- is
mirrored in `tests/scheduler-logic.test.mjs`, a plain JavaScript file with
zero dependencies, runnable with Node's built-in test runner:

```bash
node --test tests/scheduler-logic.test.mjs
```

`buildSchedule()` takes an injectable RNG function (defaults to
`Math.random` for real runs); tests pass in either a seeded deterministic
PRNG or a fixed-value function so the "random" behavior itself is
verifiable rather than just hoped-for.

This covers (28 tests): priority fill order, one desk per person per day,
never more than 4 people per day, per-desk rotation fairness (each specific
desk's assignment count stays balanced within 1 across two weeks), that the
tied-candidate pick is genuinely driven by the injected rng and not a fixed
order, Excel-serial-date math (checked against a known reference date), the
grid parser's handling of shift codes, PTO, avoid-markers (alone and
overridden), invalid names and location-banner rows, the weekly-target hard
cap -- a Flex/Part Time/Full Time employee is capped at exactly their
target even with more requests than that, PTO consumes part of the target
alongside worked shifts, someone with no roster entry is never capped by
`buildSchedule()` itself (roster filtering happens one layer up, in
`main()`) -- the 3-2 biweekly pattern (week 1 caps at 3, week 2 at 2, and
the pattern repeats correctly past a 2-week period), and that
`computeTargetShortfalls()` correctly reports (or doesn't report)
under-target employee-weeks, including resolving the right target for each
week of a 3-2 pattern.

Beyond the synthetic tests, the grid parser's logic was also run against
the real Children's Colorado export used to design it (via a one-off Node
script, not part of the checked-in suite) and produced the same request
count as the Python version's importer run against the same file -- real
cross-checked parity, not just two independent guesses.

If you change the scheduling or parsing logic in `BalanceDeskSchedule.ts`,
mirror the change in the test file too -- Office Scripts can't import
external modules, so the two are kept in sync by hand rather than by a
shared import, which is a real (if minor) maintenance cost of this
architecture compared to the Python version's proper module structure.

`sample_data/generate_template.py` (Python + openpyxl) regenerates the
sample template workbook, including a Kronos-grid-shaped Report sheet --
useful if you want to reset the demo data or build a fresh template with
your own employee names:

```bash
pip install openpyxl
python sample_data/generate_template.py
```

## Known limitations

* **Office Scripts availability isn't universal.** It requires a
  Microsoft 365 plan that includes it, and some tenants disable it via
  admin policy. Worth a quick check with IT before assuming this path is
  open -- it's a much smaller ask than approving an executable, but it's
  not automatically available everywhere.
* **Data now lives in OneDrive/SharePoint**, not a fully local file. This
  is the deliberate tradeoff described above -- flag it if your team has
  concerns about where scheduling data (employee names/shifts, not patient
  data) is stored.
* Desk rotation fairness is tracked *within* one run of the script, not
  persisted *across* runs/periods -- each run balances fresh from whatever's
  in the Report and Employees sheets at the time, so who had which desk
  last period isn't remembered when picking this period's rotation. Over
  many periods this should still average out, but it isn't guaranteed
  period-to-period the way it's guaranteed within a single run.
* Similarly, weekly targets aren't carried across a period boundary -- a
  target partially met right at the edge of the date range you enter in
  Settings isn't remembered the next time you run the script for the
  following period.
* Because desk assignment now uses real randomness (`Math.random()` in
  production), re-running the script for the *same* period will generally
  produce a different specific assignment pattern each time, even though
  the fairness guarantees (per-desk rotation, weekly target caps) hold
  every time. If you need to review a schedule before it's final, treat
  the last run before you stop re-running as the one you're reviewing.
* Same open assumptions as the Python version (see its README): multiple
  stacked location blocks in one sheet are supported but not yet tested
  against a real multi-block export, shift type (`D12` vs `08c12`)
  currently doesn't constrain desk assignment, and `RDY`/`SDY`/`VDY` are
  the only avoid-marker codes confirmed so far.
