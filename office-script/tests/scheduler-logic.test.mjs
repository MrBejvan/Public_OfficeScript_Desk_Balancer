// Node-runnable tests for the pure logic in BalanceDeskSchedule.ts.
//
// Office Scripts are single-file and can't import external modules, so the
// functions below are a deliberate line-for-line mirror of the
// ExcelScript-independent parts of BalanceDeskSchedule.ts (date/serial
// conversion, ISO week keying, the priority-fill + fairness + weekly-target
// hard-cap scheduler, and the Kronos grid parser). They have no dependency
// on the ExcelScript API, so they can run here under plain Node with zero
// installs (uses only node:test and node:assert).
//
// IMPORTANT: if you change the scheduling logic in BalanceDeskSchedule.ts,
// mirror the change here too -- these are intentionally kept in sync by
// hand rather than by a build step, to keep the Office Script a true
// single, self-contained file (a hard requirement of the Automate tab).
//
// Run with:  node tests/scheduler-logic.test.mjs

import { test } from "node:test";
import assert from "node:assert/strict";

// ============================================================
// Mirrored pure logic (see BalanceDeskSchedule.ts for the source of truth)
// ============================================================

const DESK_PRIORITY = ["East", "West", "VAB", "Resource"];
const EXCEL_EPOCH_OFFSET_DAYS = 25569;

const STATUS_PATTERN_MAP = {
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

// Simple deterministic PRNG (mulberry32) so tests can inject reproducible
// "randomness" instead of relying on Math.random(). Production code
// (BalanceDeskSchedule.ts) defaults to Math.random() for real runs.
function makeSeededRng(seed) {
  let a = seed >>> 0;
  return function () {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function excelSerialToDate(serial) {
  return new Date(Math.round((serial - EXCEL_EPOCH_OFFSET_DAYS) * 86400 * 1000));
}

function dateToExcelSerial(date) {
  return Math.round(date.getTime() / 86400000) + EXCEL_EPOCH_OFFSET_DAYS;
}

function isoWeekKey(serial) {
  const source = excelSerialToDate(serial);
  const d = new Date(Date.UTC(source.getUTCFullYear(), source.getUTCMonth(), source.getUTCDate()));
  const dayNum = (d.getUTCDay() + 6) % 7;
  d.setUTCDate(d.getUTCDate() - dayNum + 3);

  const firstThursday = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
  const firstDayNum = (firstThursday.getUTCDay() + 6) % 7;
  firstThursday.setUTCDate(firstThursday.getUTCDate() - firstDayNum + 3);

  const weekNum = 1 + Math.round((d.getTime() - firstThursday.getTime()) / (7 * 86400000));
  return `${d.getUTCFullYear()}-W${weekNum}`;
}

function weekStartSerial(serial) {
  const source = excelSerialToDate(serial);
  const d = new Date(Date.UTC(source.getUTCFullYear(), source.getUTCMonth(), source.getUTCDate()));
  const dayNum = (d.getUTCDay() + 6) % 7;
  d.setUTCDate(d.getUTCDate() - dayNum);
  return dateToExcelSerial(d);
}

function weekIndexInPeriod(day, periodStartWeekStart) {
  return Math.round((weekStartSerial(day) - periodStartWeekStart) / 7);
}

function targetForWeek(entry, weekIndex) {
  return entry.pattern[weekIndex % entry.pattern.length];
}

function computeTargetShortfalls(targets, weekAssigned, ptoByWeek, startSerial, endSerial, periodStartWeekStart) {
  const weekKeyToStart = new Map();
  const weekKeyToIndex = new Map();
  for (let day = startSerial; day <= endSerial; day++) {
    const wk = isoWeekKey(day);
    if (!weekKeyToStart.has(wk)) {
      weekKeyToStart.set(wk, weekStartSerial(day));
      weekKeyToIndex.set(wk, weekIndexInPeriod(day, periodStartWeekStart));
    }
  }

  const shortfalls = [];
  for (const [name, entry] of targets) {
    for (const [wk, wkStart] of weekKeyToStart) {
      const target = targetForWeek(entry, weekKeyToIndex.get(wk));
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

function buildSchedule(requests, ptoEntries, targets, startSerial, endSerial, rng = Math.random) {
  const periodStartWeekStart = weekStartSerial(startSerial);

  const byDay = new Map();
  for (const r of requests) {
    if (r.dateSerial >= startSerial && r.dateSerial <= endSerial) {
      if (!byDay.has(r.dateSerial)) byDay.set(r.dateSerial, []);
      byDay.get(r.dateSerial).push(r);
    }
  }

  const ptoByWeek = new Map();
  for (const p of ptoEntries) {
    if (p.dateSerial >= startSerial && p.dateSerial <= endSerial) {
      const key = `${p.name}|${isoWeekKey(p.dateSerial)}`;
      ptoByWeek.set(key, (ptoByWeek.get(key) ?? 0) + 1);
    }
  }

  const weekAssigned = new Map();
  const deskCount = new Map(); // `${name}|${desk}` -> cumulative times assigned that desk this period
  const assignments = [];
  const shortfalls = [];

  const hasReachedTarget = (name, day, weekKey) => {
    const entry = targets.get(name);
    if (!entry) return false;
    const target = targetForWeek(entry, weekIndexInPeriod(day, periodStartWeekStart));
    const worked = weekAssigned.get(`${name}|${weekKey}`) ?? 0;
    const pto = ptoByWeek.get(`${name}|${weekKey}`) ?? 0;
    return worked + pto >= target;
  };

  for (let day = startSerial; day <= endSerial; day++) {
    const dayRequests = byDay.get(day) ?? [];
    const assignedToday = new Set();
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

      // Fair-tier random pick: least-often-on-this-desk wins the tier,
      // random choice breaks ties within it.
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
      assignments.push({ dateSerial: day, desk, name: chosen.name, shiftCode: chosen.shiftCode ?? "" });
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

// ============================================================
// Mirrored grid-parsing logic (Kronos "Location Schedule" importer)
// ============================================================

const NAME_PATTERN = /^[A-Za-z][A-Za-z ,.'-]{0,79}$/;
const PTO_PATTERN = /^pto\b/i;
const AVOID_MARKER_PATTERN = /\b(RDY|SDY|VDY)\b\s*(\d{1,2}:\d{2})?/gi;
const DATE_TOKEN_PATTERN = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;

function collapseWhitespace(raw) {
  if (raw === undefined || raw === null) return "";
  return String(raw).replace(/\s+/g, " ").trim();
}

function parseName(raw) {
  const text = collapseWhitespace(raw);
  if (!text || text.startsWith("=")) return null;
  if (!NAME_PATTERN.test(text)) return null;
  return text;
}

function parseDayCell(raw) {
  const text = collapseWhitespace(raw);
  if (!text) return { working: false, shiftCode: null, isPTO: false };
  if (PTO_PATTERN.test(text)) return { working: false, shiftCode: null, isPTO: true };
  const remainder = text.replace(AVOID_MARKER_PATTERN, "").trim();
  if (!remainder) return { working: false, shiftCode: null, isPTO: false };
  return { working: true, shiftCode: remainder, isPTO: false };
}

function isLocationBanner(raw) {
  if (raw === undefined || raw === null) return false;
  return String(raw).toLowerCase().includes("location:");
}

function normalizeHeader(value) {
  if (value === undefined || value === null) return "";
  return String(value).trim().toLowerCase();
}

function parseDateHeaderSerial(raw) {
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

function findHeaderRows(values, numCols) {
  const result = [];
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

function findDateSpans(values, dateRowIdx, numCols) {
  const row = values[dateRowIdx] ?? [];
  const starts = [];
  for (let c = 0; c < numCols && c < row.length; c++) {
    const v = row[c];
    if (v !== undefined && v !== null && v !== "") starts.push({ col: c, raw: v });
  }
  const spans = [];
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

function parseBlock(values, headerRowIdx, nameColIdx, blockEndRowIdx, numCols, issues) {
  const dateRowIdx = headerRowIdx + 1;
  const dateSpans = findDateSpans(values, dateRowIdx, numCols);
  if (dateSpans.length === 0) {
    issues.push({ rowNumber: headerRowIdx + 1, message: "No date columns found under the 'Employee' header row -- skipped this block." });
    return { requests: [], ptoEntries: [] };
  }

  const requests = [];
  const ptoEntries = [];
  for (let r = dateRowIdx + 1; r < blockEndRowIdx && r < values.length; r++) {
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
      let cellValue;
      for (let c = span.startCol; c <= span.endCol; c++) {
        const v = values[r][c];
        if (v !== undefined && v !== null && v !== "") {
          cellValue = v;
          break;
        }
      }
      const { working, shiftCode, isPTO } = parseDayCell(cellValue);
      if (working) {
        requests.push({ name, dateSerial: span.serial, preferredDesk: null, shiftCode: shiftCode ?? "", rowNumber: r + 1 });
      } else if (isPTO) {
        ptoEntries.push({ name, dateSerial: span.serial });
      }
    }
  }
  return { requests, ptoEntries };
}

// ============================================================
// Helpers for tests
// ============================================================

function req(name, date, preferredDesk = null) {
  return { name, dateSerial: dateToExcelSerial(date), preferredDesk, shiftCode: "", rowNumber: 0 };
}

function pto(name, date) {
  return { name, dateSerial: dateToExcelSerial(date) };
}

function utc(y, m, d) {
  return new Date(Date.UTC(y, m - 1, d));
}

function targetsMap(entries) {
  // entries: { name: "Full Time" | "Part Time" | "Flex" | "3-2" }
  const m = new Map();
  for (const [name, status] of Object.entries(entries)) {
    m.set(name, { pattern: STATUS_PATTERN_MAP[status.toLowerCase()], status });
  }
  return m;
}

// ============================================================
// Tests
// ============================================================

test("date <-> Excel serial round-trips", () => {
  const cases = [utc(2026, 1, 1), utc(2026, 7, 27), utc(2000, 2, 29), utc(2026, 12, 31)];
  for (const d of cases) {
    const serial = dateToExcelSerial(d);
    const back = excelSerialToDate(serial);
    assert.equal(back.getUTCFullYear(), d.getUTCFullYear());
    assert.equal(back.getUTCMonth(), d.getUTCMonth());
    assert.equal(back.getUTCDate(), d.getUTCDate());
  }
});

test("known Excel serial numbers match real Excel", () => {
  // 2026-01-01 is serial 46023 in Excel's date system (verified against a
  // real workbook). This guards against an off-by-one in the epoch offset.
  assert.equal(dateToExcelSerial(utc(2026, 1, 1)), 46023);
});

test("fills desks in priority order when short-staffed", () => {
  const day = utc(2026, 7, 1);
  const requests = [req("Alice", day), req("Bob", day)];
  const serial = dateToExcelSerial(day);
  const period = buildSchedule(requests, [], new Map(), serial, serial);

  const filled = new Set(period.assignments.filter((a) => a.name !== null).map((a) => a.desk));
  assert.deepEqual([...filled].sort(), ["East", "West"]);

  const shortDesks = new Set(period.shortfalls.map((f) => f.desk));
  assert.deepEqual([...shortDesks].sort(), ["Resource", "VAB"]);
});

test("never assigns more than one desk per person per day", () => {
  const day = utc(2026, 7, 1);
  const serial = dateToExcelSerial(day);
  const period = buildSchedule([req("Alice", day)], [], new Map(), serial, serial);
  const aliceCount = period.assignments.filter((a) => a.name === "Alice").length;
  assert.equal(aliceCount, 1);
});

test("respects explicit desk preference over higher-priority open desks", () => {
  const day = utc(2026, 7, 1);
  const serial = dateToExcelSerial(day);
  const requests = [req("Alice", day, "Resource"), req("Bob", day, "East")];
  const period = buildSchedule(requests, [], new Map(), serial, serial);
  const byDesk = Object.fromEntries(period.assignments.map((a) => [a.desk, a.name]));

  assert.equal(byDesk.East, "Bob");
  assert.equal(byDesk.Resource, "Alice");
  assert.equal(byDesk.West, null);
  assert.equal(byDesk.VAB, null);
});

test("never exceeds four people per day", () => {
  const day = utc(2026, 7, 1);
  const serial = dateToExcelSerial(day);
  const requests = ["A", "B", "C", "D", "E", "F"].map((n) => req(n, day));
  const period = buildSchedule(requests, [], new Map(), serial, serial);
  const filled = period.assignments.filter((a) => a.name !== null);
  assert.equal(filled.length, 4);
  assert.equal(new Set(filled.map((a) => a.name)).size, 4);
});

test("desk rotation fairness: each specific desk's assignment count stays balanced (max-min <= 1) across two weeks", () => {
  // This is the actual fairness guarantee the new algorithm makes:
  // "everyone gets a turn at each desk," checked per-desk, not just an
  // even total shift count (which is no longer directly optimized for --
  // it's the weekly target hard cap's job to bound total shifts, not the
  // desk-picking step).
  const start = utc(2026, 7, 6); // a Monday
  const startSerial = dateToExcelSerial(start);
  const endSerial = startSerial + 13; // two full weeks
  const names = ["A", "B", "C", "D", "E"];

  const requests = [];
  for (let i = 0; i <= 13; i++) {
    const day = excelSerialToDate(startSerial + i);
    for (const name of names) requests.push(req(name, day));
  }

  const period = buildSchedule(requests, [], new Map(), startSerial, endSerial, makeSeededRng(12345));

  for (const desk of DESK_PRIORITY) {
    const counts = new Map();
    for (const name of names) counts.set(name, 0);
    for (const a of period.assignments) {
      if (a.desk === desk && a.name !== null) counts.set(a.name, counts.get(a.name) + 1);
    }
    const values = [...counts.values()];
    assert.ok(
      Math.max(...values) - Math.min(...values) <= 1,
      `desk ${desk} counts not balanced: ${JSON.stringify([...counts])}`
    );
  }
});

test("desk assignment among tied candidates is driven by the injected rng, not a fixed order", () => {
  const day = utc(2026, 7, 1);
  const serial = dateToExcelSerial(day);
  const requests = ["Alice", "Bob", "Cara", "Dan"].map((n) => req(n, day));

  const pickFirst = buildSchedule(requests, [], new Map(), serial, serial, () => 0);
  const eastFirst = pickFirst.assignments.find((a) => a.desk === "East").name;
  assert.equal(eastFirst, "Alice"); // rng() = 0 -> first in the tied tier

  const pickLast = buildSchedule(requests, [], new Map(), serial, serial, () => 0.9999999);
  const eastLast = pickLast.assignments.find((a) => a.desk === "East").name;
  assert.equal(eastLast, "Dan"); // rng() near 1 -> last in the tied tier
});

test("empty requests flags all four desks every day", () => {
  const day = utc(2026, 7, 1);
  const serial = dateToExcelSerial(day);
  const period = buildSchedule([], [], new Map(), serial, serial);
  assert.equal(period.shortfalls.length, 4);
  assert.ok(period.assignments.every((a) => a.name === null));
});

test("ISO week key groups a Mon-Sun week together and separates the next week", () => {
  const monday = dateToExcelSerial(utc(2026, 7, 6));
  const sunday = dateToExcelSerial(utc(2026, 7, 12));
  const nextMonday = dateToExcelSerial(utc(2026, 7, 13));

  assert.equal(isoWeekKey(monday), isoWeekKey(sunday));
  assert.notEqual(isoWeekKey(monday), isoWeekKey(nextMonday));
});

test("weekStartSerial returns the Monday of the ISO week for any day in it", () => {
  const monday = dateToExcelSerial(utc(2026, 7, 6));
  const wednesday = dateToExcelSerial(utc(2026, 7, 8));
  const sunday = dateToExcelSerial(utc(2026, 7, 12));

  assert.equal(weekStartSerial(monday), monday);
  assert.equal(weekStartSerial(wednesday), monday);
  assert.equal(weekStartSerial(sunday), monday);
});

// ============================================================
// Tests: parseDayCell (the core Kronos cell-interpretation rule)
// ============================================================

test("parseDayCell interprets shift codes, PTO, and avoid-markers correctly", () => {
  assert.deepEqual(parseDayCell(undefined), { working: false, shiftCode: null, isPTO: false });
  assert.deepEqual(parseDayCell(""), { working: false, shiftCode: null, isPTO: false });
  assert.deepEqual(parseDayCell("D12"), { working: true, shiftCode: "D12", isPTO: false });
  assert.deepEqual(parseDayCell("08c12"), { working: true, shiftCode: "08c12", isPTO: false });
  assert.deepEqual(parseDayCell("PTO Scheduled Hours 12:00"), { working: false, shiftCode: null, isPTO: true });
  assert.deepEqual(parseDayCell("pto"), { working: false, shiftCode: null, isPTO: true });
  assert.deepEqual(parseDayCell("RDY"), { working: false, shiftCode: null, isPTO: false });
  assert.deepEqual(parseDayCell("SDY"), { working: false, shiftCode: null, isPTO: false });
  assert.deepEqual(parseDayCell("VDY 12:30"), { working: false, shiftCode: null, isPTO: false });
  assert.deepEqual(parseDayCell("RDY \nD12"), { working: true, shiftCode: "D12", isPTO: false });
  assert.deepEqual(parseDayCell("SDY\nD12"), { working: true, shiftCode: "D12", isPTO: false });
});

// ============================================================
// Tests: parseName (formula-injection / bad-data guard)
// ============================================================

test("parseName accepts real names and rejects junk", () => {
  assert.equal(parseName("Alvarez, Jordan"), "Alvarez, Jordan");
  assert.equal(parseName("  Alvarez,   Jordan  "), "Alvarez, Jordan");
  assert.equal(parseName("=HACK()"), null);
  assert.equal(parseName("123 Not A Name!!"), null);
  assert.equal(parseName(""), null);
  assert.equal(parseName(undefined), null);
});

// ============================================================
// Tests: the grid parser end-to-end against a synthetic Kronos-shaped sheet
// ============================================================

function buildGridValues() {
  // Mirrors a real export's shape: header row with "Employee"/"Job", a
  // date row below it (dates spread across 2-column spans, matching how
  // Kronos visually merges them), then employee rows. Row/col indices
  // below are 0-based, matching what usedRange.getValues() would return.
  const rows = [];
  for (let i = 0; i < 8; i++) rows.push([]);

  const set = (r, c, v) => {
    while (rows[r].length <= c) rows[r].push("");
    rows[r][c] = v;
  };

  // header row (index 0)
  set(0, 0, "Employee");
  set(0, 2, "Job");
  // date row (index 1): two dates, 2-column spans starting at col 4 and 6
  set(1, 4, "8/01/\n2026");
  set(1, 6, "8/02/\n2026");
  // employee rows
  set(2, 0, "Alvarez, Jordan");
  set(2, 2, "ISS");
  set(2, 5, "RDY \nD12"); // value in the *second* column of the day-1 span
  set(2, 6, "08c12"); // value in the first column of the day-2 span
  set(3, 0, "Brooks, Taylor");
  set(3, 2, "ISS");
  set(3, 4, "SDY"); // avoid-marker alone -> not working
  // (day 2 left blank for Brooks)

  const numCols = Math.max(...rows.map((r) => r.length));
  for (const r of rows) while (r.length < numCols) r.push("");
  return { values: rows, numCols };
}

test("grid parser: finds the Employee header and date spans", () => {
  const { values, numCols } = buildGridValues();
  const headers = findHeaderRows(values, numCols);
  assert.deepEqual(headers, [{ rowIdx: 0, colIdx: 0 }]);

  const spans = findDateSpans(values, 1, numCols);
  assert.equal(spans.length, 2);
  assert.equal(spans[0].startCol, 4);
  assert.equal(spans[0].endCol, 5); // ends right before the next date's column
  assert.equal(spans[1].startCol, 6);
});

test("grid parser: finds a value regardless of which column in the span it's in", () => {
  const { values, numCols } = buildGridValues();
  const issues = [];
  const { requests } = parseBlock(values, 0, 0, values.length, numCols, issues);

  const jordan = requests.filter((r) => r.name === "Alvarez, Jordan");
  assert.equal(jordan.length, 2);
  assert.equal(jordan[0].shiftCode, "D12"); // overridden RDY, found in column 5 (2nd col of span)
  assert.equal(jordan[1].shiftCode, "08c12");

  const taylor = requests.filter((r) => r.name === "Brooks, Taylor");
  assert.equal(taylor.length, 0); // SDY alone on day 1, blank on day 2 -> no requests
  assert.equal(issues.length, 0);
});

test("grid parser: a bad employee name is skipped with an issue, not silently dropped", () => {
  const { values, numCols } = buildGridValues();
  values[2][0] = "###Bad###";
  const issues = [];
  const { requests } = parseBlock(values, 0, 0, values.length, numCols, issues);
  assert.equal(requests.filter((r) => r.name === "###Bad###").length, 0);
  assert.equal(issues.length, 1);
  assert.match(issues[0].message, /invalid or unreadable employee name/);
});

test("grid parser: a location-banner row in the name column is skipped, not read as an employee", () => {
  const { values, numCols } = buildGridValues();
  values.push(["Some Location: Banner", "", "", "", "", "", ""]);
  const issues = [];
  const { requests } = parseBlock(values, 0, 0, values.length, numCols, issues);
  // Still just Jordan's 2 + nothing from the banner row.
  assert.equal(requests.length, 2);
  assert.equal(issues.length, 0);
});

test("grid parser: a PTO cell produces a PtoEntry, not a ShiftRequest", () => {
  const { values, numCols } = buildGridValues();
  values[3][6] = "PTO Scheduled Hours 12:00"; // Taylor, day 2
  const issues = [];
  const { requests, ptoEntries } = parseBlock(values, 0, 0, values.length, numCols, issues);

  assert.equal(requests.filter((r) => r.name === "Brooks, Taylor").length, 0);
  assert.equal(ptoEntries.length, 1);
  assert.equal(ptoEntries[0].name, "Brooks, Taylor");
});

// ============================================================
// Tests: weekly-target hard cap (Full Time / Part Time / Flex)
// ============================================================

test("hard cap: stops assigning someone once they hit their weekly target, desk stays open", () => {
  // A Flex employee (target 1) requests every day of a 3-day stretch, with
  // no one else available. They should only be assigned once; the other
  // days' East desk should be left open and flagged rather than double-booked.
  const days = [utc(2026, 7, 6), utc(2026, 7, 7), utc(2026, 7, 8)];
  const requests = days.map((d) => req("Alice", d));
  const startSerial = dateToExcelSerial(days[0]);
  const endSerial = dateToExcelSerial(days[2]);
  const targets = targetsMap({ Alice: "Flex" });

  const period = buildSchedule(requests, [], targets, startSerial, endSerial);
  const aliceAssignments = period.assignments.filter((a) => a.name === "Alice");
  assert.equal(aliceAssignments.length, 1);

  const eastShortfalls = period.shortfalls.filter((s) => s.desk === "East");
  assert.equal(eastShortfalls.length, 2);
  assert.match(eastShortfalls[0].reason, /reached their weekly shift target/);
});

test("hard cap: a Full Time employee can work all 3 of their target shifts, capped at the 4th", () => {
  const days = [utc(2026, 7, 6), utc(2026, 7, 7), utc(2026, 7, 8), utc(2026, 7, 9)];
  const requests = days.map((d) => req("Bob", d));
  const startSerial = dateToExcelSerial(days[0]);
  const endSerial = dateToExcelSerial(days[3]);
  const targets = targetsMap({ Bob: "Full Time" });

  const period = buildSchedule(requests, [], targets, startSerial, endSerial);
  const bobAssignments = period.assignments.filter((a) => a.name === "Bob");
  assert.equal(bobAssignments.length, 3);
});

test("hard cap: PTO counts toward the weekly target alongside worked shifts", () => {
  // Part Time target = 2. A PTO day earlier in the week already uses up
  // one slot of the target, so only one of the two requested work days
  // that same week can be filled -- the other is left open, not double
  // counted on top of the PTO day.
  const week = [utc(2026, 7, 6), utc(2026, 7, 7)]; // Mon, Tue
  const requests = [req("Cara", week[0]), req("Cara", week[1])];
  const ptoEntries = [pto("Cara", utc(2026, 7, 8))]; // Wed PTO, same week
  const startSerial = dateToExcelSerial(week[0]);
  const endSerial = dateToExcelSerial(utc(2026, 7, 8));
  const targets = targetsMap({ Cara: "Part Time" });

  const period = buildSchedule(requests, ptoEntries, targets, startSerial, endSerial);
  const caraAssignments = period.assignments.filter((a) => a.name === "Cara");
  assert.equal(caraAssignments.length, 1); // 1 worked + 1 PTO already equals the target of 2

  // And the week is NOT reported as under target, since worked + PTO meets it.
  const shortfalls = period.targetShortfalls.filter((s) => s.name === "Cara");
  assert.equal(shortfalls.length, 0);
});

test("hard cap does not exclude someone with no roster entry (never capped, but also never scheduled by main())", () => {
  const day = utc(2026, 7, 6);
  const serial = dateToExcelSerial(day);
  const requests = [req("Ghost", day)];
  const period = buildSchedule(requests, [], new Map(), serial, serial); // empty targets map
  const ghostAssignments = period.assignments.filter((a) => a.name === "Ghost");
  assert.equal(ghostAssignments.length, 1); // buildSchedule itself doesn't filter by roster membership
});

// ============================================================
// Tests: under-weekly-target reporting
// ============================================================

test("computeTargetShortfalls reports every roster employee-week under target", () => {
  const startSerial = dateToExcelSerial(utc(2026, 7, 6)); // Monday
  const endSerial = dateToExcelSerial(utc(2026, 7, 12)); // Sunday, one full week
  const targets = targetsMap({ Dana: "Full Time", Eli: "Flex" });

  const weekAssigned = new Map([
    [`Dana|${isoWeekKey(startSerial)}`, 1], // worked 1, target 3 -> shortfall
    [`Eli|${isoWeekKey(startSerial)}`, 1], // worked 1, target 1 -> met
  ]);
  const ptoByWeek = new Map();
  const periodStartWeekStart = weekStartSerial(startSerial);

  const shortfalls = computeTargetShortfalls(targets, weekAssigned, ptoByWeek, startSerial, endSerial, periodStartWeekStart);
  assert.equal(shortfalls.length, 1);
  assert.equal(shortfalls[0].name, "Dana");
  assert.equal(shortfalls[0].worked, 1);
  assert.equal(shortfalls[0].target, 3);
});

test("computeTargetShortfalls treats PTO as counting toward the target, same as worked shifts", () => {
  const startSerial = dateToExcelSerial(utc(2026, 7, 6));
  const endSerial = dateToExcelSerial(utc(2026, 7, 12));
  const targets = targetsMap({ Fay: "Part Time" });

  const weekAssigned = new Map([[`Fay|${isoWeekKey(startSerial)}`, 1]]);
  const ptoByWeek = new Map([[`Fay|${isoWeekKey(startSerial)}`, 1]]);
  const periodStartWeekStart = weekStartSerial(startSerial);

  const shortfalls = computeTargetShortfalls(targets, weekAssigned, ptoByWeek, startSerial, endSerial, periodStartWeekStart);
  assert.equal(shortfalls.length, 0); // 1 worked + 1 PTO = target of 2, met
});

test("STATUS_PATTERN_MAP recognizes Full Time / Part Time / Flex / 3-2 spellings", () => {
  assert.deepEqual(STATUS_PATTERN_MAP["full time"], [3]);
  assert.deepEqual(STATUS_PATTERN_MAP["ft"], [3]);
  assert.deepEqual(STATUS_PATTERN_MAP["part time"], [2]);
  assert.deepEqual(STATUS_PATTERN_MAP["pt"], [2]);
  assert.deepEqual(STATUS_PATTERN_MAP["flex"], [1]);
  assert.deepEqual(STATUS_PATTERN_MAP["flexible"], [1]);
  assert.deepEqual(STATUS_PATTERN_MAP["3-2"], [3, 2]);
  assert.deepEqual(STATUS_PATTERN_MAP["3/2"], [3, 2]);
  assert.deepEqual(STATUS_PATTERN_MAP["biweekly"], [3, 2]);
});

// ============================================================
// Tests: the "3-2" biweekly status pattern
// ============================================================

test("3-2 status: week 1 of the period caps at 3, week 2 caps at 2", () => {
  // A 3-2 employee requests every day of a 2-week period, with no one
  // else available. Week 1 (period week index 0) should let them work up
  // to 3 shifts; week 2 (index 1) should cap them at 2.
  const start = utc(2026, 7, 6); // Monday, week 1 of the period
  const startSerial = dateToExcelSerial(start);
  const endSerial = startSerial + 13; // through the following Sunday (2 weeks)
  const requests = [];
  for (let i = 0; i <= 13; i++) requests.push(req("Gail", excelSerialToDate(startSerial + i)));
  const targets = targetsMap({ Gail: "3-2" });

  const period = buildSchedule(requests, [], targets, startSerial, endSerial);

  const week1Start = dateToExcelSerial(start);
  const week2Start = week1Start + 7;
  const week1Count = period.assignments.filter(
    (a) => a.name === "Gail" && a.dateSerial >= week1Start && a.dateSerial < week2Start
  ).length;
  const week2Count = period.assignments.filter(
    (a) => a.name === "Gail" && a.dateSerial >= week2Start
  ).length;

  assert.equal(week1Count, 3);
  assert.equal(week2Count, 2);
});

test("3-2 status: the pattern repeats past a 2-week period (week 3 is a '3' week again)", () => {
  const start = utc(2026, 7, 6); // Monday
  const startSerial = dateToExcelSerial(start);
  const endSerial = startSerial + 20; // 3 full weeks
  const requests = [];
  for (let i = 0; i <= 20; i++) requests.push(req("Hank", excelSerialToDate(startSerial + i)));
  const targets = targetsMap({ Hank: "3-2" });

  const period = buildSchedule(requests, [], targets, startSerial, endSerial);
  const week3Start = startSerial + 14;
  const week3Count = period.assignments.filter((a) => a.name === "Hank" && a.dateSerial >= week3Start).length;
  assert.equal(week3Count, 3); // week index 2 -> pattern[2 % 2] = pattern[0] = 3
});

test("computeTargetShortfalls resolves the correct target for each week of a 3-2 pattern", () => {
  const startSerial = dateToExcelSerial(utc(2026, 7, 6)); // Monday, week 1
  const endSerial = dateToExcelSerial(utc(2026, 7, 19)); // through week 2's Sunday
  const targets = targetsMap({ Ivy: "3-2" });
  const periodStartWeekStart = weekStartSerial(startSerial);

  const week1Key = isoWeekKey(startSerial);
  const week2Key = isoWeekKey(startSerial + 7);
  const weekAssigned = new Map([
    [`Ivy|${week1Key}`, 3], // met the week-1 target of 3
    [`Ivy|${week2Key}`, 1], // short of the week-2 target of 2
  ]);
  const ptoByWeek = new Map();

  const shortfalls = computeTargetShortfalls(targets, weekAssigned, ptoByWeek, startSerial, endSerial, periodStartWeekStart);
  assert.equal(shortfalls.length, 1);
  assert.equal(shortfalls[0].worked, 1);
  assert.equal(shortfalls[0].target, 2);
});
