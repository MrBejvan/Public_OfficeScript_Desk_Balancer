/**
 * Desk Coverage Scheduler -- Office Script
 * =========================================
 *
 * Runs entirely inside this workbook (Excel Online / Excel desktop with a
 * Microsoft 365 account, via the Automate tab). No external services, no
 * network calls beyond what OneDrive/SharePoint already does to save this
 * file -- this is the Office-Scripts sibling of a standalone Python/Tkinter
 * version of the same tool, ported here because it can run inside
 * infrastructure IT already trusts (Microsoft 365) instead of as a foreign
 * executable on a locked-down machine.
 *
 * WORKBOOK LAYOUT THIS SCRIPT EXPECTS
 * ------------------------------------
 *   A sheet containing Kronos's actual "Location Schedule - Monthly" export
 *   grid (reverse engineered from a real Children's Colorado pull, NOT a
 *   simple table). Found by name (anything containing "report",
 *   case-insensitive -- what Kronos itself names it) or, failing that, by
 *   scanning every sheet for a cell reading "Employee".
 *
 *   Inside that sheet:
 *     - A row with a literal cell "Employee" starts a block. The row right
 *       below it holds the dates for that block: one non-blank cell per
 *       day. Kronos visually merges each date across 2-3 columns, but only
 *       the left-most cell of a merge holds a value -- so "which columns
 *       belong to today" is computed by finding the *next* non-blank date
 *       cell, without needing to inspect Excel's merge geometry at all.
 *     - Each employee row has, for each date's column span, at most one
 *       non-blank cell:
 *         blank                     -> not requesting to work
 *         a shift code, e.g. "D12"  -> working that shift
 *         "PTO ..."                 -> approved time off (still counts
 *                                       toward their weekly shift target,
 *                                       see below -- just doesn't need a
 *                                       desk)
 *         "RDY" / "SDY" / "VDY 12:30" -> a soft "please don't schedule me
 *                                         this day if possible" marker,
 *                                         honored when it appears alone --
 *                                         does NOT count toward their
 *                                         weekly target
 *         marker + shift on two lines, e.g. "RDY\nD12" -> the soft request
 *                                         was overridden; they're working
 *     - A real export can stack more than one such block in one sheet (one
 *       per Kronos location/cost-center code). Each is parsed independently.
 *
 *   "Employees" sheet -- the roster. Two columns (header row, then one row
 *   per person):
 *      Name   | Status
 *      ------ | ------
 *      (as it appears in Kronos, e.g. "Alvarez, Jordan")
 *             | Full Time (3 shifts/week), Part Time (2/week), Flex
 *               (1/week), or 3-2 (a biweekly status between Full Time and
 *               Part Time: 3 shifts one week, 2 the next, alternating --
 *               week 1 of whatever period you enter in Settings is always
 *               treated as the "3" week). Case-insensitive; FT/PT/3-2/
 *               biweekly spellings are all accepted.
 *   Kronos has no concept of employment status, so this is maintained by
 *   hand. Anyone in the Kronos data who isn't listed here is skipped
 *   entirely (not scheduled, flagged on Needs Attention) until added.
 *
 *   "Settings" sheet -- plain cells (not a table):
 *      B1 = schedule period start date
 *      B2 = schedule period end date
 *      B4 = status message written back by this script after each run
 *
 * WHAT THIS SCRIPT DOES
 * ----------------------
 *   1. Reads + validates the Kronos grid and the Employees roster (bad
 *      rows are skipped and reported, never crash the whole run).
 *   2. Balances coverage across four desks -- East, West, VAB, Resource --
 *      in that priority order, for every day in the Settings period. Desks
 *      that can't be covered are left blank and flagged rather than
 *      guessed.
 *   3. Assigns fairly, with randomness instead of a fixed formula: for each
 *      desk, among everyone still eligible to work that day, whoever has
 *      had *that specific desk* the fewest times so far this period is
 *      favored; if several people are tied for fewest, one is picked at
 *      random. This is deliberately not a deterministic "auto-balance"
 *      score -- it's a simple, auditable rule (least-often-on-this-desk)
 *      with a random tie-break, so no two runs produce the identical
 *      pattern, but everyone still rotates through all four desks over
 *      time.
 *   4. Enforces each person's weekly shift target as a hard cap: once
 *      their worked shifts + PTO days for the week reach their status's
 *      target for that specific week (Full Time is always 3, Part Time
 *      always 2, Flex always 1, and 3-2 alternates 3/2/3/2... starting
 *      from week 1 of the period), they're excluded from every desk for
 *      the rest of that week, even if that leaves a desk unfilled -- it
 *      gets flagged instead of over-scheduling them.
 *   5. Writes the result back into this workbook: one printable calendar
 *      sheet per month in the period (each day showing the assigned shift
 *      code alongside the name), a "Summary" sheet totaling shifts per
 *      person alongside their status/target, and a "Needs Attention" sheet
 *      listing unfilled desks, import notes, and anyone under their weekly
 *      target for a given week. Previous runs' generated sheets are
 *      replaced, not accumulated.
 *   6. Writes a plain-English status message to Settings!B4.
 *
 * This script deliberately never throws an unhandled error -- every
 * failure path is caught and turned into a status message instead.
 *
 * OPEN ASSUMPTIONS (validated against one real single-location export;
 * flag these if a fuller pull shows otherwise):
 *   - Multiple stacked location blocks in one sheet are supported but not
 *     yet tested against a real multi-block file.
 *   - Shift type (D12 vs 08c12) does not currently constrain which desk
 *     someone can be assigned to -- "worked that day at all" is treated as
 *     sufficient for the balancer.
 *   - RDY/SDY/VDY are the only avoid-marker codes seen so far; there may
 *     be others in a fuller export.
 *
 * NOTE ON FILE EXTENSION: Office Scripts are normally saved/imported with
 * a .osts extension inside Excel's Automate tab. This file is named .ts
 * only so it displays nicely outside Excel -- the contents are identical
 * to what you'd paste into a new script in the Code Editor. See README.md
 * for exact setup steps.
 */

// ============================================================
// Types
// ============================================================

type Desk = "East" | "West" | "VAB" | "Resource";
type CellValue = string | number | boolean;

const DESK_PRIORITY: Desk[] = ["East", "West", "VAB", "Resource"];

interface ShiftRequest {
  name: string;
  dateSerial: number; // whole-day Excel serial date
  preferredDesk: Desk | null; // always null for real Kronos data; kept for extensibility
  shiftCode: string; // e.g. "D12", "08c12" -- informational, shown on the calendar
  rowNumber: number; // 1-based row within the sheet's used range, for issue messages
}

interface PtoEntry {
  name: string;
  dateSerial: number;
}

interface RowIssue {
  rowNumber: number;
  message: string;
}

interface DeskAssignment {
  dateSerial: number;
  desk: Desk;
  name: string | null; // null = unfilled
  shiftCode: string;
}

interface ShortfallFlag {
  dateSerial: number;
  desk: Desk;
  reason: string;
}

interface TargetShortfall {
  name: string;
  weekStartSerial: number;
  worked: number;
  pto: number;
  target: number;
}

interface RosterEntry {
  /** Weekly shift-count pattern, indexed by week-of-period (0-based),
   * cycling: pattern[weekIndex % pattern.length]. Full Time/Part Time/Flex
   * are a single-element pattern (flat every week); the 3-2 status is
   * [3, 2] (alternates, starting with the 3-week in week 0 of the period). */
  pattern: number[];
  status: string;
}

interface SchedulePeriod {
  startSerial: number;
  endSerial: number;
  assignments: DeskAssignment[];
  shortfalls: ShortfallFlag[];
  targetShortfalls: TargetShortfall[];
}

// ============================================================
// Constants
// ============================================================

const MAX_EMPLOYEE_ROWS = 5000;

const NAME_PATTERN = /^[A-Za-z][A-Za-z ,.'-]{0,79}$/;
const PTO_PATTERN = /^pto\b/i;
const AVOID_MARKER_PATTERN = /\b(RDY|SDY|VDY)\b\s*(\d{1,2}:\d{2})?/gi;
const DATE_TOKEN_PATTERN = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;

const STATUS_PATTERN_MAP: { [key: string]: number[] } = {
  "full time": [3],
  "full-time": [3],
  ft: [3],
  "part time": [2],
  "part-time": [2],
  pt: [2],
  flex: [1],
  flexible: [1],
  "3-2": [3, 2],
  "3/2": [3, 2],
  "3 2": [3, 2],
  "32": [3, 2],
  biweekly: [3, 2],
  "bi-weekly": [3, 2],
  "bi weekly": [3, 2],
};

const MONTH_NAMES_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const MONTH_NAMES_LONG = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];
const WEEKDAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

const COLOR_HEADER = "#1F4E78";
const COLOR_WEEKDAY = "#D9E1F2";
const COLOR_OPEN = "#FCE4E4";
const COLOR_OUT_OF_MONTH = "#F2F2F2";
const COLOR_STATUS_OK = "#E2EFDA";
const COLOR_STATUS_ERROR = "#FCE4E4";

const RESERVED_SHEET_NAMES = new Set(["settings", "read me", "employees", "summary", "needs attention"]);
const MONTH_SHEET_PATTERN = /^[A-Za-z]{3} \d{4}$/;

// ============================================================
// Entry point
// ============================================================

function main(workbook: ExcelScript.Workbook) {
  const settingsSheet = workbook.getWorksheet("Settings");
  if (!settingsSheet) {
    // Nowhere to write a status message either -- this is the one failure
    // mode that can't be reported back inside the workbook.
    console.log("ERROR: This workbook is missing a 'Settings' sheet. See BalanceDeskSchedule README.");
    return;
  }

  try {
    const reportSheet = findReportSheet(workbook);
    if (!reportSheet) {
      setStatus(
        settingsSheet,
        "Error: could not find a sheet with Kronos schedule data (looked for a sheet named " +
          "'Report', or any sheet containing a cell that reads 'Employee').",
        true
      );
      return;
    }

    const roster = readRoster(workbook);
    if (!roster.sheetFound) {
      setStatus(
        settingsSheet,
        "Error: no 'Employees' sheet found. Add one with Name and Status " +
          "(Full Time / Part Time / Flex / 3-2) columns before running.",
        true
      );
      return;
    }

    const { startSerial, endSerial, error: periodError } = readPeriod(settingsSheet);
    if (periodError) {
      setStatus(settingsSheet, `Error: ${periodError}`, true);
      return;
    }

    const importResult = importRequests(reportSheet);

    const skippedNames = new Set<string>();
    const filteredRequests = importResult.requests.filter((r) => {
      if (roster.targets.has(r.name)) return true;
      skippedNames.add(r.name);
      return false;
    });
    const filteredPto = importResult.ptoEntries.filter((p) => roster.targets.has(p.name));

    const allIssues = [...importResult.issues, ...roster.issues];
    if (skippedNames.size > 0) {
      allIssues.push({
        rowNumber: 0,
        message:
          `Skipped ${skippedNames.size} employee(s) not on the Employees roster ` +
          `(their requests were not scheduled): ${[...skippedNames].sort().join(", ")}.`,
      });
    }

    if (filteredRequests.length === 0 && importResult.requests.length === 0 && allIssues.length > 0) {
      setStatus(
        settingsSheet,
        `Error: no valid requests found. First issue -- Row ${allIssues[0].rowNumber}: ${allIssues[0].message}`,
        true
      );
      return;
    }

    const period = buildSchedule(filteredRequests, filteredPto, roster.targets, startSerial, endSerial);

    removeGeneratedSheets(workbook);
    writeNeedsAttentionSheet(workbook, period, allIssues);
    writeSummarySheet(workbook, period, roster.targets);
    const months = monthsInPeriod(startSerial, endSerial);
    for (const m of months) {
      writeMonthSheet(workbook, m.year, m.month, period);
    }

    const filled = period.assignments.filter((a) => a.name !== null).length;
    const total = period.assignments.length;
    const statusParts = [
      `Balanced ${formatSerialAsMDY(startSerial)} - ${formatSerialAsMDY(endSerial)} from "${reportSheet.getName()}".`,
      `${filled}/${total} desk-shifts filled, ${period.shortfalls.length} open.`,
      `${period.targetShortfalls.length} employee-week(s) under their weekly target.`,
      `${allIssues.length} import note(s) -- see "Needs Attention" sheet.`,
      `Last run: ${new Date().toLocaleString()}`,
    ];
    setStatus(
      settingsSheet,
      statusParts.join(" "),
      period.shortfalls.length > 0 || period.targetShortfalls.length > 0
    );
  } catch (err) {
    // Belt-and-suspenders: nothing above should throw, but if something
    // unexpected does, surface it in plain language instead of failing
    // silently or dumping a stack trace on a non-technical user.
    const message = err instanceof Error ? err.message : String(err);
    setStatus(settingsSheet, `Unexpected error: ${message}`, true);
  }
}

// ============================================================
// Settings / status helpers
// ============================================================

function readPeriod(settingsSheet: ExcelScript.Worksheet): {
  startSerial: number;
  endSerial: number;
  error: string | null;
} {
  const startRaw = settingsSheet.getRange("B1").getValues()[0][0];
  const endRaw = settingsSheet.getRange("B2").getValues()[0][0];

  const startSerial = coerceToSerial(startRaw);
  const endSerial = coerceToSerial(endRaw);

  if (startSerial === null) {
    return { startSerial: 0, endSerial: 0, error: "Settings!B1 (start date) isn't a valid date." };
  }
  if (endSerial === null) {
    return { startSerial: 0, endSerial: 0, error: "Settings!B2 (end date) isn't a valid date." };
  }
  if (startSerial > endSerial) {
    return { startSerial: 0, endSerial: 0, error: "The start date (B1) must be on or before the end date (B2)." };
  }
  return { startSerial, endSerial, error: null };
}

function setStatus(settingsSheet: ExcelScript.Worksheet, message: string, isWarningOrError: boolean) {
  const cell = settingsSheet.getRange("B4");
  cell.setValues([[message]]);
  cell.getFormat().getFill().setColor(isWarningOrError ? COLOR_STATUS_ERROR : COLOR_STATUS_OK);
  cell.getFormat().setWrapText(true);
}

// ============================================================
// Date helpers (Excel serial <-> JS Date, both treated as UTC/whole days)
// ============================================================

const EXCEL_EPOCH_OFFSET_DAYS = 25569; // days between Excel's day-0 and the Unix epoch

function excelSerialToDate(serial: number): Date {
  return new Date(Math.round((serial - EXCEL_EPOCH_OFFSET_DAYS) * 86400 * 1000));
}

function dateToExcelSerial(date: Date): number {
  return Math.round(date.getTime() / 86400000) + EXCEL_EPOCH_OFFSET_DAYS;
}

function coerceToSerial(raw: CellValue): number | null {
  if (typeof raw === "number") {
    return Math.round(raw);
  }
  if (typeof raw === "string") {
    const text = raw.trim();
    if (!text) return null;
    const parsed = parseDateString(text);
    return parsed === null ? null : dateToExcelSerial(parsed);
  }
  return null;
}

function parseDateString(text: string): Date | null {
  // Accepts MM/DD/YYYY, YYYY-MM-DD, M/D/YY
  let m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(Date.UTC(Number(m[3]), Number(m[1]) - 1, Number(m[2])));

  m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));

  m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);
  if (m) return new Date(Date.UTC(2000 + Number(m[3]), Number(m[1]) - 1, Number(m[2])));

  return null;
}

/** Parses a Kronos date-header cell, e.g. a number (real date serial) or
 * text like "8/01/\n2026" (note the embedded newline seen in real exports). */
function parseDateHeaderSerial(raw: CellValue): number | null {
  if (typeof raw === "number") return Math.round(raw);
  if (typeof raw !== "string") return null;
  const compact = raw.replace(/\s+/g, "");
  const m = compact.match(DATE_TOKEN_PATTERN);
  if (!m) return null;
  const month = Number(m[1]);
  const day = Number(m[2]);
  const year = Number(m[3]);
  return dateToExcelSerial(new Date(Date.UTC(year, month - 1, day)));
}

function formatSerialAsMDY(serial: number): string {
  const d = excelSerialToDate(serial);
  return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`;
}

function isoWeekKey(serial: number): string {
  const source = excelSerialToDate(serial);
  const d = new Date(Date.UTC(source.getUTCFullYear(), source.getUTCMonth(), source.getUTCDate()));
  const dayNum = (d.getUTCDay() + 6) % 7; // Monday = 0 .. Sunday = 6
  d.setUTCDate(d.getUTCDate() - dayNum + 3); // move to the Thursday of this ISO week

  const firstThursday = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
  const firstDayNum = (firstThursday.getUTCDay() + 6) % 7;
  firstThursday.setUTCDate(firstThursday.getUTCDate() - firstDayNum + 3);

  const weekNum = 1 + Math.round((d.getTime() - firstThursday.getTime()) / (7 * 86400000));
  return `${d.getUTCFullYear()}-W${weekNum}`;
}

/** The Excel serial of the Monday that starts the ISO week containing `serial`. */
function weekStartSerial(serial: number): number {
  const source = excelSerialToDate(serial);
  const d = new Date(Date.UTC(source.getUTCFullYear(), source.getUTCMonth(), source.getUTCDate()));
  const dayNum = (d.getUTCDay() + 6) % 7; // Monday = 0 .. Sunday = 6
  d.setUTCDate(d.getUTCDate() - dayNum);
  return dateToExcelSerial(d);
}

function monthsInPeriod(startSerial: number, endSerial: number): { year: number; month: number }[] {
  const months: { year: number; month: number }[] = [];
  const start = excelSerialToDate(startSerial);
  const end = excelSerialToDate(endSerial);
  let year = start.getUTCFullYear();
  let month = start.getUTCMonth() + 1; // 1-based
  const endYear = end.getUTCFullYear();
  const endMonth = end.getUTCMonth() + 1;

  while (year < endYear || (year === endYear && month <= endMonth)) {
    months.push({ year, month });
    if (month === 12) {
      year += 1;
      month = 1;
    } else {
      month += 1;
    }
  }
  return months;
}

// ============================================================
// Sheet discovery
// ============================================================

function normalizeHeader(value: CellValue | undefined): string {
  if (value === undefined || value === null) return "";
  return String(value).trim().toLowerCase();
}

function getUsedRangeSafe(sheet: ExcelScript.Worksheet): ExcelScript.Range | undefined {
  try {
    return sheet.getUsedRange();
  } catch {
    return undefined;
  }
}

function findReportSheet(workbook: ExcelScript.Workbook): ExcelScript.Worksheet | undefined {
  const sheets = workbook.getWorksheets();

  for (const sheet of sheets) {
    if (sheet.getName().toLowerCase().includes("report")) return sheet;
  }

  // Fallback: scan any non-reserved, non-generated sheet for a literal
  // "Employee" header cell, in case the sheet wasn't named "Report".
  for (const sheet of sheets) {
    const name = sheet.getName();
    if (RESERVED_SHEET_NAMES.has(name.toLowerCase()) || MONTH_SHEET_PATTERN.test(name)) continue;
    const usedRange = getUsedRangeSafe(sheet);
    if (!usedRange) continue;
    const values = usedRange.getValues();
    for (const row of values) {
      for (const cell of row) {
        if (normalizeHeader(cell) === "employee") return sheet;
      }
    }
  }
  return undefined;
}

// ============================================================
// Roster (Employees sheet: Name -> Full Time / Part Time / Flex)
// ============================================================

function readRoster(workbook: ExcelScript.Workbook): {
  targets: Map<string, RosterEntry>;
  issues: RowIssue[];
  sheetFound: boolean;
} {
  const sheet = workbook.getWorksheets().find((s) => s.getName().toLowerCase() === "employees");
  if (!sheet) {
    return { targets: new Map(), issues: [], sheetFound: false };
  }

  const targets = new Map<string, RosterEntry>();
  const issues: RowIssue[] = [];

  const usedRange = getUsedRangeSafe(sheet);
  if (!usedRange) {
    return { targets, issues, sheetFound: true };
  }

  const values = usedRange.getValues();
  if (values.length === 0) {
    return { targets, issues, sheetFound: true };
  }

  const header = values[0];
  let nameCol = 0;
  let statusCol = 1;
  for (let c = 0; c < header.length; c++) {
    const h = normalizeHeader(header[c]);
    if (h === "name") nameCol = c;
    else if (h === "status") statusCol = c;
  }

  for (let r = 1; r < values.length; r++) {
    const nameRaw = values[r][nameCol];
    if (nameRaw === undefined || nameRaw === null || nameRaw === "") continue;

    const name = parseName(nameRaw);
    if (name === null) {
      issues.push({ rowNumber: r + 1, message: `Employees sheet: skipped row with invalid name ("${nameRaw}").` });
      continue;
    }

    const statusRaw = values[r][statusCol];
    const statusText = collapseWhitespace(statusRaw);
    const pattern = STATUS_PATTERN_MAP[statusText.toLowerCase()];
    if (pattern === undefined) {
      issues.push({
        rowNumber: r + 1,
        message: `Employees sheet: status "${statusRaw}" for ${name} was not recognized (expected Full Time / Part Time / Flex / 3-2) -- this person will be skipped until fixed.`,
      });
      continue;
    }

    targets.set(name, { pattern, status: statusText });
  }

  return { targets, issues, sheetFound: true };
}

// ============================================================
// Import + validation (Kronos "Location Schedule" grid parser)
// ============================================================

interface ImportResult {
  requests: ShiftRequest[];
  ptoEntries: PtoEntry[];
  issues: RowIssue[];
}

function collapseWhitespace(raw: CellValue | undefined): string {
  if (raw === undefined || raw === null) return "";
  return String(raw).replace(/\s+/g, " ").trim();
}

function parseName(raw: CellValue | undefined): string | null {
  const text = collapseWhitespace(raw);
  if (!text || text.startsWith("=")) return null;
  if (!NAME_PATTERN.test(text)) return null;
  return text;
}

/** Interprets one grid cell for a single employee/day. `working: false`
 * means the person is not requesting/scheduled to work that day. `isPTO`
 * is only ever true when `working` is false -- PTO days don't need a
 * desk, but do count toward the person's weekly shift target. */
function parseDayCell(raw: CellValue | undefined): { working: boolean; shiftCode: string | null; isPTO: boolean } {
  const text = collapseWhitespace(raw);
  if (!text) return { working: false, shiftCode: null, isPTO: false };
  if (PTO_PATTERN.test(text)) return { working: false, shiftCode: null, isPTO: true };
  const remainder = text.replace(AVOID_MARKER_PATTERN, "").trim();
  if (!remainder) return { working: false, shiftCode: null, isPTO: false };
  return { working: true, shiftCode: remainder, isPTO: false };
}

function isLocationBanner(raw: CellValue | undefined): boolean {
  if (raw === undefined || raw === null) return false;
  return String(raw).toLowerCase().includes("location:");
}

function findHeaderRows(values: CellValue[][], numCols: number): { rowIdx: number; colIdx: number }[] {
  const result: { rowIdx: number; colIdx: number }[] = [];
  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    for (let c = 0; c < numCols && c < row.length; c++) {
      if (normalizeHeader(row[c]) === "employee") {
        result.push({ rowIdx: r, colIdx: c });
        break;
      }
    }
  }
  return result;
}

interface DateSpan {
  startCol: number;
  endCol: number;
  serial: number | null;
  rawText: string;
}

function findDateSpans(values: CellValue[][], dateRowIdx: number, numCols: number): DateSpan[] {
  const row = values[dateRowIdx] ?? [];
  const starts: { col: number; raw: CellValue }[] = [];
  for (let c = 0; c < numCols && c < row.length; c++) {
    const v = row[c];
    if (v !== undefined && v !== null && v !== "") starts.push({ col: c, raw: v });
  }

  const spans: DateSpan[] = [];
  for (let i = 0; i < starts.length; i++) {
    const endCol = i + 1 < starts.length ? starts[i + 1].col - 1 : numCols - 1;
    spans.push({
      startCol: starts[i].col,
      endCol,
      serial: parseDateHeaderSerial(starts[i].raw),
      rawText: String(starts[i].raw),
    });
  }
  return spans;
}

function parseBlock(
  values: CellValue[][],
  headerRowIdx: number,
  nameColIdx: number,
  blockEndRowIdx: number,
  numCols: number,
  issues: RowIssue[]
): { requests: ShiftRequest[]; ptoEntries: PtoEntry[] } {
  const dateRowIdx = headerRowIdx + 1;
  const dateSpans = findDateSpans(values, dateRowIdx, numCols);
  if (dateSpans.length === 0) {
    issues.push({
      rowNumber: headerRowIdx + 1,
      message: "No date columns found under the 'Employee' header row -- skipped this block.",
    });
    return { requests: [], ptoEntries: [] };
  }

  const badDates = dateSpans.filter((s) => s.serial === null);
  if (badDates.length > 0) {
    issues.push({
      rowNumber: dateRowIdx + 1,
      message: `${badDates.length} date header(s) could not be parsed and were skipped.`,
    });
  }

  const requests: ShiftRequest[] = [];
  const ptoEntries: PtoEntry[] = [];
  let rowsSeen = 0;

  for (let r = dateRowIdx + 1; r < blockEndRowIdx && r < values.length; r++) {
    rowsSeen += 1;
    if (rowsSeen > MAX_EMPLOYEE_ROWS) {
      issues.push({ rowNumber: r + 1, message: `Stopped reading after ${MAX_EMPLOYEE_ROWS} employee rows (safety limit).` });
      break;
    }

    const nameRaw = values[r][nameColIdx];
    if (nameRaw === undefined || nameRaw === null || nameRaw === "" || isLocationBanner(nameRaw)) {
      continue;
    }

    const name = parseName(nameRaw);
    if (name === null) {
      issues.push({ rowNumber: r + 1, message: `Skipped row: invalid or unreadable employee name ("${nameRaw}").` });
      continue;
    }

    for (const span of dateSpans) {
      if (span.serial === null) continue;
      let cellValue: CellValue | undefined;
      for (let c = span.startCol; c <= span.endCol; c++) {
        const v = values[r][c];
        if (v !== undefined && v !== null && v !== "") {
          cellValue = v;
          break;
        }
      }
      const { working, shiftCode, isPTO } = parseDayCell(cellValue);
      if (working) {
        requests.push({
          name,
          dateSerial: span.serial,
          preferredDesk: null,
          shiftCode: shiftCode ?? "",
          rowNumber: r + 1,
        });
      } else if (isPTO) {
        ptoEntries.push({ name, dateSerial: span.serial });
      }
    }
  }

  return { requests, ptoEntries };
}

function importRequests(sheet: ExcelScript.Worksheet): ImportResult {
  const issues: RowIssue[] = [];
  const requests: ShiftRequest[] = [];
  const ptoEntries: PtoEntry[] = [];

  const usedRange = getUsedRangeSafe(sheet);
  if (!usedRange) {
    issues.push({ rowNumber: 0, message: `The '${sheet.getName()}' sheet appears to be empty.` });
    return { requests, ptoEntries, issues };
  }

  const values = usedRange.getValues();
  const numCols = values.length > 0 ? values[0].length : 0;

  const headerRows = findHeaderRows(values, numCols);
  if (headerRows.length === 0) {
    issues.push({
      rowNumber: 0,
      message: "Could not find a cell reading 'Employee' -- this doesn't look like a Kronos Location Schedule export.",
    });
    return { requests, ptoEntries, issues };
  }

  for (let i = 0; i < headerRows.length; i++) {
    const { rowIdx, colIdx } = headerRows[i];
    const blockEnd = i + 1 < headerRows.length ? headerRows[i + 1].rowIdx : values.length;
    const block = parseBlock(values, rowIdx, colIdx, blockEnd, numCols, issues);
    requests.push(...block.requests);
    ptoEntries.push(...block.ptoEntries);
  }

  return { requests, ptoEntries, issues };
}

// ============================================================
// Scheduler (priority fill + random fair-tier desk rotation + weekly
// target hard cap)
// ============================================================

/** Which week of the schedule period `day` falls in, 0-based, relative to
 * the ISO week containing `periodStartWeekStart`. Used to resolve a
 * roster entry's shift-count pattern for that specific week (so a 3-2
 * status alternates starting with the "3" week in week 0 of the period). */
function weekIndexInPeriod(day: number, periodStartWeekStart: number): number {
  return Math.round((weekStartSerial(day) - periodStartWeekStart) / 7);
}

function targetForWeek(entry: RosterEntry, weekIndex: number): number {
  return entry.pattern[weekIndex % entry.pattern.length];
}

function buildSchedule(
  requests: ShiftRequest[],
  ptoEntries: PtoEntry[],
  targets: Map<string, RosterEntry>,
  startSerial: number,
  endSerial: number,
  rng: () => number = Math.random
): SchedulePeriod {
  const periodStartWeekStart = weekStartSerial(startSerial);

  const byDay = new Map<number, ShiftRequest[]>();
  for (const r of requests) {
    if (r.dateSerial >= startSerial && r.dateSerial <= endSerial) {
      if (!byDay.has(r.dateSerial)) byDay.set(r.dateSerial, []);
      byDay.get(r.dateSerial)!.push(r);
    }
  }

  const ptoByWeek = new Map<string, number>();
  for (const p of ptoEntries) {
    if (p.dateSerial >= startSerial && p.dateSerial <= endSerial) {
      const key = `${p.name}|${isoWeekKey(p.dateSerial)}`;
      ptoByWeek.set(key, (ptoByWeek.get(key) ?? 0) + 1);
    }
  }

  const weekAssigned = new Map<string, number>();
  // Cumulative count of how many times each person has been assigned each
  // specific desk so far this period -- the basis for "everyone gets a
  // turn at each desk," not just a turn at some desk.
  const deskCount = new Map<string, number>();

  const assignments: DeskAssignment[] = [];
  const shortfalls: ShortfallFlag[] = [];

  const hasReachedTarget = (name: string, day: number, weekKey: string): boolean => {
    const entry = targets.get(name);
    if (!entry) return false; // no target on file -- don't cap (shouldn't normally happen, callers pre-filter to roster members)
    const target = targetForWeek(entry, weekIndexInPeriod(day, periodStartWeekStart));
    const worked = weekAssigned.get(`${name}|${weekKey}`) ?? 0;
    const pto = ptoByWeek.get(`${name}|${weekKey}`) ?? 0;
    return worked + pto >= target;
  };

  for (let day = startSerial; day <= endSerial; day++) {
    const dayRequests = byDay.get(day) ?? [];
    const assignedToday = new Set<string>();
    const weekKey = isoWeekKey(day);

    for (const desk of DESK_PRIORITY) {
      const notYetAssignedToday = dayRequests.filter((r) => !assignedToday.has(r.name));
      const available = notYetAssignedToday.filter((r) => !hasReachedTarget(r.name, day, weekKey));
      const preferred = available.filter((r) => r.preferredDesk === desk);
      const flexible = available.filter((r) => r.preferredDesk === null);
      const pool = preferred.length > 0 ? preferred : flexible;

      if (pool.length === 0) {
        const reason =
          notYetAssignedToday.length > 0 && available.length === 0
            ? "Everyone who could work this day has already reached their weekly shift target"
            : "No one available or scheduled for this desk that day";
        shortfalls.push({ dateSerial: day, desk, reason });
        assignments.push({ dateSerial: day, desk, name: null, shiftCode: "" });
        continue;
      }

      // Fair-tier random pick: among the pool, find whoever has had THIS
      // desk the fewest times so far this period, then pick uniformly at
      // random among everyone tied for that minimum. This replaces a
      // deterministic "fewest total shifts" formula with something that's
      // both provably fair (no one gets a 2nd turn at a desk before
      // everyone eligible has had a 1st) and genuinely randomized (no
      // fixed tie-break order, so the exact pattern differs run to run).
      let minCount = Infinity;
      for (const r of pool) {
        const c = deskCount.get(`${r.name}|${desk}`) ?? 0;
        if (c < minCount) minCount = c;
      }
      const tier = pool.filter((r) => (deskCount.get(`${r.name}|${desk}`) ?? 0) === minCount);
      const pickIndex = Math.min(tier.length - 1, Math.floor(rng() * tier.length));
      const chosen = tier[pickIndex];

      assignedToday.add(chosen.name);
      const dk = `${chosen.name}|${desk}`;
      deskCount.set(dk, (deskCount.get(dk) ?? 0) + 1);
      const wk = `${chosen.name}|${weekKey}`;
      weekAssigned.set(wk, (weekAssigned.get(wk) ?? 0) + 1);
      assignments.push({ dateSerial: day, desk, name: chosen.name, shiftCode: chosen.shiftCode });
    }
  }

  const targetShortfalls = computeTargetShortfalls(
    targets,
    weekAssigned,
    ptoByWeek,
    startSerial,
    endSerial,
    periodStartWeekStart
  );

  return { startSerial, endSerial, assignments, shortfalls, targetShortfalls };
}

function computeTargetShortfalls(
  targets: Map<string, RosterEntry>,
  weekAssigned: Map<string, number>,
  ptoByWeek: Map<string, number>,
  startSerial: number,
  endSerial: number,
  periodStartWeekStart: number
): TargetShortfall[] {
  const weekKeyToStart = new Map<string, number>();
  const weekKeyToIndex = new Map<string, number>();
  for (let day = startSerial; day <= endSerial; day++) {
    const wk = isoWeekKey(day);
    if (!weekKeyToStart.has(wk)) {
      weekKeyToStart.set(wk, weekStartSerial(day));
      weekKeyToIndex.set(wk, weekIndexInPeriod(day, periodStartWeekStart));
    }
  }

  const shortfalls: TargetShortfall[] = [];
  for (const [name, entry] of targets) {
    for (const [wk, wkStart] of weekKeyToStart) {
      const target = targetForWeek(entry, weekKeyToIndex.get(wk)!);
      const worked = weekAssigned.get(`${name}|${wk}`) ?? 0;
      const pto = ptoByWeek.get(`${name}|${wk}`) ?? 0;
      if (worked + pto < target) {
        shortfalls.push({ name, weekStartSerial: wkStart, worked, pto, target });
      }
    }
  }
  shortfalls.sort((a, b) => a.weekStartSerial - b.weekStartSerial || a.name.localeCompare(b.name));
  return shortfalls;
}

// ============================================================
// Export: sheet management
// ============================================================

function removeGeneratedSheets(workbook: ExcelScript.Workbook) {
  const sheets = workbook.getWorksheets();
  for (const sheet of sheets) {
    const name = sheet.getName();
    if (name === "Summary" || name === "Needs Attention" || MONTH_SHEET_PATTERN.test(name)) {
      sheet.delete();
    }
  }
}

function applyOuterAndInnerBorders(range: ExcelScript.Range) {
  const format = range.getFormat();
  const edges: ExcelScript.BorderIndex[] = [
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight,
    ExcelScript.BorderIndex.insideHorizontal,
    ExcelScript.BorderIndex.insideVertical,
  ];
  for (const edge of edges) {
    const border = format.getRangeBorder(edge);
    border.setStyle(ExcelScript.BorderLineStyle.continuous);
    border.setColor("#AAAAAA");
    border.setWeight(ExcelScript.BorderWeight.thin);
  }
}

// ============================================================
// Export: Needs Attention sheet
// ============================================================

function writeNeedsAttentionSheet(workbook: ExcelScript.Workbook, period: SchedulePeriod, issues: RowIssue[]) {
  if (period.shortfalls.length === 0 && issues.length === 0 && period.targetShortfalls.length === 0) return;

  const sheet = workbook.addWorksheet("Needs Attention");
  sheet.setPosition(0);

  let row = 0;

  if (issues.length > 0) {
    sheet.getRange(`A${row + 1}`).setValues([["Import notes"]]);
    sheet.getRange(`A${row + 1}`).getFormat().getFont().setBold(true);
    row += 1;
    const headerRange = sheet.getRange(`A${row + 1}:B${row + 1}`);
    headerRange.setValues([["Row", "Message"]]);
    headerRange.getFormat().getFont().setBold(true);
    headerRange.getFormat().getFill().setColor(COLOR_WEEKDAY);
    row += 1;
    for (const issue of issues) {
      row += 1;
      sheet.getRange(`A${row}:B${row}`).setValues([[issue.rowNumber, issue.message]]);
    }
    row += 2;
  }

  if (period.shortfalls.length > 0) {
    sheet.getRange(`A${row + 1}`).setValues([["Unfilled desk-days"]]);
    sheet.getRange(`A${row + 1}`).getFormat().getFont().setBold(true);
    row += 1;
    const headerRange = sheet.getRange(`A${row + 1}:C${row + 1}`);
    headerRange.setValues([["Date", "Desk", "Reason"]]);
    headerRange.getFormat().getFont().setBold(true);
    headerRange.getFormat().getFill().setColor(COLOR_OPEN);
    row += 1;

    const sortedShortfalls = [...period.shortfalls].sort((a, b) => {
      if (a.dateSerial !== b.dateSerial) return a.dateSerial - b.dateSerial;
      return DESK_PRIORITY.indexOf(a.desk) - DESK_PRIORITY.indexOf(b.desk);
    });
    for (const flag of sortedShortfalls) {
      row += 1;
      sheet
        .getRange(`A${row}:C${row}`)
        .setValues([[formatSerialAsMDY(flag.dateSerial), flag.desk, flag.reason]]);
    }
    row += 2;
  }

  if (period.targetShortfalls.length > 0) {
    sheet.getRange(`A${row + 1}`).setValues([["Under weekly target (worked + PTO vs. status target)"]]);
    sheet.getRange(`A${row + 1}`).getFormat().getFont().setBold(true);
    row += 1;
    const headerRange = sheet.getRange(`A${row + 1}:E${row + 1}`);
    headerRange.setValues([["Week Of", "Employee", "Worked", "PTO", "Target"]]);
    headerRange.getFormat().getFont().setBold(true);
    headerRange.getFormat().getFill().setColor(COLOR_OPEN);
    row += 1;

    for (const ts of period.targetShortfalls) {
      row += 1;
      sheet
        .getRange(`A${row}:E${row}`)
        .setValues([[formatSerialAsMDY(ts.weekStartSerial), ts.name, ts.worked, ts.pto, ts.target]]);
    }
  }

  sheet.getRange("A1:E1").getFormat().setColumnWidth(150);
  sheet.getRange("A1:A1").getFormat().setColumnWidth(90);
}

// ============================================================
// Export: Summary sheet
// ============================================================

function writeSummarySheet(workbook: ExcelScript.Workbook, period: SchedulePeriod, targets: Map<string, RosterEntry>) {
  const sheet = workbook.addWorksheet("Summary");
  sheet.setPosition(0);

  const header = ["Employee", "Status", "Weekly Target", "Total Shifts", "East", "West", "VAB", "Resource"];
  const headerRange = sheet.getRange("A1:H1");
  headerRange.setValues([header]);
  headerRange.getFormat().getFont().setBold(true);
  headerRange.getFormat().getFill().setColor(COLOR_WEEKDAY);

  const perPerson = new Map<string, { total: number; East: number; West: number; VAB: number; Resource: number }>();
  for (const a of period.assignments) {
    if (a.name === null) continue;
    if (!perPerson.has(a.name)) {
      perPerson.set(a.name, { total: 0, East: 0, West: 0, VAB: 0, Resource: 0 });
    }
    const stats = perPerson.get(a.name)!;
    stats.total += 1;
    stats[a.desk] += 1;
  }

  const names = [...perPerson.keys()].sort();
  const rows: (string | number)[][] = names.map((name) => {
    const s = perPerson.get(name)!;
    const roster = targets.get(name);
    const targetLabel = roster ? roster.pattern.join("-") : "";
    return [
      name,
      roster?.status ?? "",
      targetLabel,
      s.total,
      s.East,
      s.West,
      s.VAB,
      s.Resource,
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(`A2:H${1 + rows.length}`).setValues(rows);
  }

  sheet.getRange("A1:H1").getFormat().setColumnWidth(110);
  sheet.getRange("A1:A1").getFormat().setColumnWidth(150);
  sheet.getRange("B1:B1").getFormat().setColumnWidth(110);
}

// ============================================================
// Export: monthly calendar grid
// ============================================================

function monthGridWeeks(year: number, month1based: number): (Date | null)[][] {
  const firstOfMonth = new Date(Date.UTC(year, month1based - 1, 1));
  const startWeekday = firstOfMonth.getUTCDay(); // 0 = Sunday
  const daysInThisMonth = new Date(Date.UTC(year, month1based, 0)).getUTCDate();

  const cells: (Date | null)[] = [];
  for (let i = 0; i < startWeekday; i++) cells.push(null);
  for (let d = 1; d <= daysInThisMonth; d++) cells.push(new Date(Date.UTC(year, month1based - 1, d)));
  while (cells.length % 7 !== 0) cells.push(null);

  const weeks: (Date | null)[][] = [];
  for (let i = 0; i < cells.length; i += 7) weeks.push(cells.slice(i, i + 7));
  return weeks;
}

function writeMonthSheet(workbook: ExcelScript.Workbook, year: number, month1based: number, period: SchedulePeriod) {
  const sheetName = `${MONTH_NAMES_SHORT[month1based - 1]} ${year}`;
  const sheet = workbook.addWorksheet(sheetName);

  const title = `${MONTH_NAMES_LONG[month1based - 1]} ${year} - Desk Coverage Schedule`;
  const titleRange = sheet.getRange("A1:G1");
  titleRange.merge();
  titleRange.setValues([[title, "", "", "", "", "", ""]]);
  titleRange.getFormat().getFont().setBold(true);
  titleRange.getFormat().getFont().setSize(16);
  titleRange.getFormat().getFont().setColor("#FFFFFF");
  titleRange.getFormat().getFill().setColor(COLOR_HEADER);
  titleRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  titleRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  titleRange.getFormat().setRowHeight(28);

  const weekdayRange = sheet.getRange("A2:G2");
  weekdayRange.setValues([WEEKDAY_NAMES]);
  weekdayRange.getFormat().getFont().setBold(true);
  weekdayRange.getFormat().getFill().setColor(COLOR_WEEKDAY);
  weekdayRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  sheet.getRange("A1:G1").getFormat().setColumnWidth(150);

  const weeks = monthGridWeeks(year, month1based);
  let rowIndex = 3;
  for (const week of weeks) {
    const rowRange = sheet.getRange(`A${rowIndex}:G${rowIndex}`);
    rowRange.getFormat().setRowHeight(78);

    for (let col = 0; col < 7; col++) {
      const day = week[col];
      const colLetter = String.fromCharCode("A".charCodeAt(0) + col);
      const cell = sheet.getRange(`${colLetter}${rowIndex}`);
      cell.getFormat().setWrapText(true);
      cell.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
      cell.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

      if (day === null) {
        cell.setValues([[""]]);
        cell.getFormat().getFill().setColor(COLOR_OUT_OF_MONTH);
        continue;
      }

      const serial = dateToExcelSerial(day);
      const dayAssignments = period.assignments.filter((a) => a.dateSerial === serial);
      const lines = [String(day.getUTCDate())];
      let hasOpen = false;
      for (const desk of DESK_PRIORITY) {
        const match = dayAssignments.find((a) => a.desk === desk);
        if (!match) continue;
        if (match.name !== null) {
          const label = match.shiftCode ? `${desk}: ${match.name} (${match.shiftCode})` : `${desk}: ${match.name}`;
          lines.push(label);
        } else {
          lines.push(`${desk}: (OPEN)`);
          hasOpen = true;
        }
      }
      cell.setValues([[lines.join("\n")]]);
      if (hasOpen) {
        cell.getFormat().getFill().setColor(COLOR_OPEN);
      }
    }
    rowIndex += 1;
  }

  const gridRange = sheet.getRange(`A2:G${rowIndex - 1}`);
  applyOuterAndInnerBorders(gridRange);

  sheet.getFreezePanes().freezeRows(2);
  sheet.getPageLayout().setOrientation(ExcelScript.PageOrientation.landscape);
  sheet.getPageLayout().setZoom({ horizontalFitToPages: 1 });
}
