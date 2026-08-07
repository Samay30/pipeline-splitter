/**
 * Assembles the pipeline-meeting KPI payload for one recruiter and writes it
 * next to their recap file as `<master>_KPI.json`.
 *
 * Why here: the Claude skill that renders the PDF runs in a sandbox that can
 * reach neither SharePoint nor Ringover, and a connector read hands Claude
 * *text*, not a file its scripts can open. This job already holds the master
 * as structured JSON, so it does the arithmetic once, on a schedule, and
 * leaves behind a small file the skill can read and render verbatim.
 *
 * Everything is best-effort in the same way the phone pull is: a workbook that
 * doesn't parse yields a payload with notes and unavailable cards rather than
 * an exception. The recap file is the thing that must never fail.
 */

import type { PhoneRow } from "./ringover";
import { parseWeekTab } from "./tabs";

/** One sheet as read from the Graph workbook API. */
export interface SheetData {
  values: unknown[][];
  numberFormat: string[][];
}

export type SheetMap = Map<string, SheetData>;

const DAY_LABELS = ["Mon", "Tue", "Wed", "Thu", "Fri"] as const;
const PLACEHOLDERS = new Set(["tbd", "not sure", "n/a", "na", "none", "-", "--", "?"]);

/* -------------------------------------------------------------------------- */
/* cell helpers                                                               */
/* -------------------------------------------------------------------------- */

function norm(v: unknown): string {
  if (v === null || v === undefined) return "";
  return String(v).replace(/\s+/g, " ").trim();
}

function nameKey(s: unknown): string {
  return norm(s).toLowerCase().replace(/[^a-z ]/g, "").trim();
}

/** Excel serial -> Date. Day 0 is 1899-12-30 in the 1900 system. */
function serialToDate(serial: number): Date {
  return new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000);
}

function isDateFormat(fmt: string | undefined): boolean {
  return !!fmt && fmt !== "General" && /[dmy]/i.test(fmt);
}

/**
 * A genuine date, or null. Text placeholders ("TBD") and names typed into a
 * date column ("Sue Lonegran") are not dates and must never read as events.
 */
function asDate(value: unknown, fmt?: string): Date | null {
  if (typeof value === "number" && isDateFormat(fmt) && value > 0) {
    return serialToDate(value);
  }
  const s = norm(value);
  if (!s || PLACEHOLDERS.has(s.toLowerCase())) return null;
  const m = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (m) {
    let year = Number(m[3]);
    if (year < 100) year += year < 70 ? 2000 : 1900;
    const d = new Date(Date.UTC(year, Number(m[1]) - 1, Number(m[2])));
    return Number.isNaN(d.valueOf()) ? null : d;
  }
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(`${iso[0]}T00:00:00.000Z`);
  return null;
}

/** m/d/yy — pipeline-table style, no zero padding. */
function fmtDate(d: Date): string {
  return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${String(d.getUTCFullYear()).slice(2)}`;
}

/** mm/dd/yy — header and card-sublabel style, zero padded. */
function fmtDatePadded(d: Date): string {
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(d.getUTCDate()).padStart(2, "0");
  return `${mm}/${dd}/${String(d.getUTCFullYear()).slice(2)}`;
}

/** Date formatted, else whatever text is there, else an em dash. */
function keepText(value: unknown, fmt?: string): string {
  const d = asDate(value, fmt);
  if (d) return fmtDate(d);
  return norm(value) || "\u2014";
}

function dashOrDate(value: unknown, fmt?: string): string {
  const d = asDate(value, fmt);
  return d ? fmtDate(d) : "\u2014";
}

function toNumber(value: unknown): number | null {
  if (typeof value === "number") return value;
  const s = norm(value).replace(/[^0-9.\-]/g, "");
  if (!s) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function sameDay(a: Date | null, b: Date): boolean {
  return !!a && a.toISOString().slice(0, 10) === b.toISOString().slice(0, 10);
}

function isoDay(d: Date): string {
  return d.toISOString().slice(0, 10);
}

function cellAt(sheet: SheetData, r: number, c: number): { v: unknown; f?: string } {
  return { v: sheet.values?.[r]?.[c], f: sheet.numberFormat?.[r]?.[c] };
}

/* -------------------------------------------------------------------------- */
/* pipeline table                                                             */
/* -------------------------------------------------------------------------- */

const HEADER_ALIASES: Record<string, string[]> = {
  institution: ["institution", "company"],
  position: ["current position/hiring for", "position", "hiring for"],
  location: ["location"],
  hiring_official: ["hiring official"],
  candidate: ["candidate"],
  submitted: ["submitted"],
  first: ["first"],
  second: ["second"],
  third: ["third"],
  fourth: ["fourth"],
  fifth: ["fifth"],
  sixth: ["sixth"],
  seventh: ["seventh"],
  offer: ["offer"],
  start: ["start"],
  notes: ["notes"],
};

function mapHeader(row: unknown[]): Record<string, number> {
  const out: Record<string, number> = {};
  row.forEach((raw, i) => {
    const label = norm(raw).toLowerCase().replace(/:$/, "");
    for (const [key, aliases] of Object.entries(HEADER_ALIASES)) {
      if (aliases.includes(label) && !(key in out)) out[key] = i;
    }
  });
  return out;
}

function findPipelineHeader(sheet: SheetData): { row: number; cols: Record<string, number> } | null {
  for (let r = 0; r < sheet.values.length; r++) {
    const first = norm(sheet.values[r]?.[0]).toLowerCase();
    if (first === "institution" || first === "company") {
      const cols = mapHeader(sheet.values[r]);
      if ("candidate" in cols && "first" in cols) return { row: r, cols };
    }
  }
  return null;
}

const PAREN = /\(([^)]*)\)/g;

function splitCandidate(raw: string): { name: string; note: string } {
  const s = norm(raw);
  if (!s) return { name: "", note: "" };
  const parens = [...s.matchAll(PAREN)].map((m) => norm(m[1])).filter(Boolean);
  return { name: norm(s.replace(PAREN, "")), note: parens.join(", ") };
}

export interface PipelineRow {
  company: string;
  position: string;
  location: string;
  hiring_official: string;
  candidate: string;
  candidate_note: string;
  submitted: string;
  first: string;
  second: string;
  third: string;
  fourth: string;
  offer: string;
  start: string;
  notes: string;
}

function readPipelineRows(sheet: SheetData): { rows: PipelineRow[]; notes: string[] } {
  const header = findPipelineHeader(sheet);
  if (!header) {
    return { rows: [], notes: ["Couldn't read the Pipeline Meeting tab."] };
  }
  const { cols } = header;
  const at = (r: number, key: string) =>
    cols[key] === undefined ? { v: undefined } : cellAt(sheet, r, cols[key]);

  const rows: PipelineRow[] = [];
  for (let r = header.row + 1; r < sheet.values.length; r++) {
    const institution = norm(at(r, "institution").v);
    if (institution.toLowerCase().startsWith("holding pattern")) break;

    const candidateRaw = norm(at(r, "candidate").v);
    if (!candidateRaw) continue; // open search with no active candidate

    const { name, note } = splitCandidate(candidateRaw);

    // The layout has room for four rounds; flag any dated rounds beyond that.
    let extra = 0;
    for (const key of ["fifth", "sixth", "seventh"]) {
      const c = at(r, key);
      if (asDate(c.v, c.f)) extra++;
    }
    let rowNotes = norm(at(r, "notes").v);
    if (extra > 0) {
      const tail = `+${extra} more round${extra > 1 ? "s" : ""}`;
      rowNotes = rowNotes ? `${rowNotes} (${tail})` : `(${tail})`;
    }

    const d = (key: string) => {
      const c = at(r, key);
      return dashOrDate(c.v, c.f);
    };
    const t = (key: string) => {
      const c = at(r, key);
      return keepText(c.v, c.f);
    };

    rows.push({
      company: institution,
      position: norm(at(r, "position").v),
      location: norm(at(r, "location").v),
      hiring_official: norm(at(r, "hiring_official").v),
      candidate: name,
      candidate_note: note,
      submitted: d("submitted"),
      first: d("first"),
      second: d("second"),
      third: d("third"),
      fourth: d("fourth"),
      offer: t("offer"),
      start: t("start"),
      notes: rowNotes,
    });
  }
  return { rows, notes: rows.length ? [] : ["No candidates on the Pipeline Meeting tab."] };
}

/* -------------------------------------------------------------------------- */
/* weekly KPI grid                                                            */
/* -------------------------------------------------------------------------- */

function readKpiRow(
  sheet: SheetData,
  week: Date
): { row: Record<string, number | null>; found: boolean; notes: string[] } {
  let headerRow = -1;
  for (let r = 0; r < sheet.values.length; r++) {
    if (norm(sheet.values[r]?.[0]).toLowerCase() === "week") {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) {
    return { row: {}, found: false, notes: ["Couldn't find the weekly KPI table."] };
  }
  const labels = sheet.values[headerRow].map((x) => norm(x));

  for (let r = headerRow + 1; r < sheet.values.length; r++) {
    const label0 = norm(sheet.values[r]?.[0]).toLowerCase();
    if (label0.startsWith("quarter") || label0.startsWith("ytd")) continue;
    const c0 = cellAt(sheet, r, 0);
    if (!sameDay(asDate(c0.v, c0.f), week)) continue;

    const out: Record<string, number | null> = {};
    for (let i = 1; i < labels.length; i++) {
      if (!labels[i]) continue;
      const raw = sheet.values[r]?.[i];
      out[labels[i]] = norm(raw) === "" ? null : toNumber(raw);
    }
    return { row: out, found: true, notes: [] };
  }
  return {
    row: {},
    found: false,
    notes: [`No KPI row for the week of ${fmtDate(week)} in this workbook.`],
  };
}

/* -------------------------------------------------------------------------- */
/* billings                                                                   */
/* -------------------------------------------------------------------------- */

export interface BillingCard {
  available: boolean;
  year?: number;
  actual?: number;
  annual_goal?: number;
  annual_pct?: number;
  stretch_goal?: number;
  stretch_pct?: number;
}

/**
 * Read one recruiter's row out of the "<year> Grand Totals" block of the
 * master billings workbook. Sheet is passed in already read, since one
 * billings file serves every recruiter in the run.
 */
export function readBilling(
  sheet: SheetData | undefined,
  recruiter: string,
  year: number
): { billing: BillingCard; notes: string[] } {
  if (!sheet) return { billing: { available: false }, notes: [] };

  let blockRow = -1;
  for (let r = 0; r < sheet.values.length; r++) {
    const label = norm(sheet.values[r]?.[0]).toLowerCase();
    if (label.includes("grand total") && label.includes(String(year))) {
      blockRow = r;
      break;
    }
  }
  if (blockRow < 0) {
    return { billing: { available: false }, notes: [`No ${year} Grand Totals block in the billings workbook.`] };
  }

  let headerRow = -1;
  for (let r = blockRow + 1; r < Math.min(blockRow + 5, sheet.values.length); r++) {
    if (norm(sheet.values[r]?.[0]).toLowerCase() === "recruiter") {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) {
    return { billing: { available: false }, notes: ["Couldn't read the Grand Totals header row."] };
  }

  const labels = sheet.values[headerRow].map((x) => norm(x).toLowerCase());
  const findCol = (needles: string[], exclude: string[] = []) => {
    for (let i = 1; i < labels.length; i++) {
      const lab = labels[i];
      if (!lab) continue;
      if (needles.some((n) => lab.includes(n)) && !exclude.some((x) => lab.includes(x))) return i;
    }
    return -1;
  };

  const cGoal = findCol(["total goal", "annual goal"]);
  const cTotal = findCol(["total"], ["goal"]);
  const cStretch = findCol(["stretch goal"], ["%", "difference"]);
  if (cGoal < 0 || cTotal < 0 || cStretch < 0) {
    return { billing: { available: false }, notes: ["Couldn't identify the billings columns."] };
  }

  // Names in this block are inconsistent ("Jason " with a trailing space,
  // "Caleb Passo" in full), so match on the full name then the first name.
  const want = nameKey(recruiter);
  const wantFirst = want.split(" ")[0] ?? "";
  for (let r = headerRow + 1; r < sheet.values.length; r++) {
    const label = nameKey(sheet.values[r]?.[0]);
    if (!label) continue;
    if (/^(recruiter total|ees grand total|retained)/.test(label)) continue;
    if (label !== want && label.split(" ")[0] !== wantFirst) continue;

    const goal = toNumber(sheet.values[r][cGoal]) ?? 0;
    const total = toNumber(sheet.values[r][cTotal]) ?? 0;
    const stretch = toNumber(sheet.values[r][cStretch]) ?? 0;
    return {
      billing: {
        available: true,
        year,
        actual: total,
        annual_goal: goal,
        annual_pct: goal ? total / goal : 0,
        stretch_goal: stretch,
        stretch_pct: stretch ? total / stretch : 0,
      },
      notes: [],
    };
  }
  return {
    billing: { available: false },
    notes: [`${recruiter} isn't listed in the ${year} Grand Totals block.`],
  };
}

/* -------------------------------------------------------------------------- */
/* payload                                                                    */
/* -------------------------------------------------------------------------- */

/** "Aaron_Rider_Pipeline_2026.xlsx" or "Aaron Rider Pipeline 2026.xlsx" -> "Aaron Rider" */
export function recruiterFromFilename(masterName: string): string {
  const stem = masterName.replace(/\.xlsx$/i, "");
  const cut = stem.split(/[_ ]Pipeline[_ ]/i)[0] ?? stem;
  return cut.replace(/_/g, " ").replace(/\s+/g, " ").trim();
}

/** The date-named tab with the latest date on or before `asOf`. */
export function pickWeekTab(
  sheetNames: string[],
  asOf: Date,
  defaultYear?: number
): { tab: string; week: Date } | null {
  const dated = sheetNames
    .map((name) => ({ name, date: parseWeekTab(name, defaultYear) }))
    .filter((x): x is { name: string; date: Date } => x.date !== null)
    .sort((a, b) => a.date.getTime() - b.date.getTime());
  if (dated.length === 0) return null;
  const eligible = dated.filter((x) => x.date.getTime() <= asOf.getTime());
  const chosen = eligible.length ? eligible[eligible.length - 1] : dated[0];
  return { tab: chosen.name, week: chosen.date };
}

export interface BuildOptions {
  masterName: string;
  sheets: SheetMap;
  billingSheet?: SheetData;
  phoneRows: PhoneRow[];
  firm?: string;
  metric?: string;
  metricLabel?: string;
  phoneGoalHours?: number;
  now?: Date;
}

export function buildKpiPayload(opts: BuildOptions): Record<string, unknown> {
  const {
    masterName,
    sheets,
    billingSheet,
    phoneRows,
    firm = "Eggers Executive Search",
    metric = "Second Interviews",
    metricLabel = "Weekly 2nd Interviews",
    phoneGoalHours = 3,
  } = opts;
  const now = opts.now ?? new Date();
  const notes: string[] = [];

  const recruiter = recruiterFromFilename(masterName);
  const firstName = recruiter.split(" ")[0] ?? "";
  const yearMatch = masterName.match(/\b(20\d{2})\b/);
  const defaultYear = yearMatch ? Number(yearMatch[1]) : undefined;

  const picked = pickWeekTab([...sheets.keys()] as string[], now, defaultYear);
  if (!picked) {
    return {
      error: "No date-named weekly tabs in this workbook.",
      recruiter: { display_name: recruiter, first_name: firstName, firm },
    };
  }
  const { tab, week } = picked;

  const ageDays = Math.floor((now.getTime() - week.getTime()) / 86400000);
  if (ageDays > 10) {
    notes.push(
      `The most recent week in this workbook is ${fmtDate(week)}, about ` +
        `${Math.floor(ageDays / 7)} weeks back — this week's tab may not be filled in yet.`
    );
  }

  // Pipeline table comes off the hand-curated Pipeline Meeting tab.
  const pmName = [...sheets.keys()].find((n) => norm(n).toLowerCase() === "pipeline meeting");
  let pipeline: PipelineRow[] = [];
  if (pmName) {
    const res = readPipelineRows(sheets.get(pmName)!);
    pipeline = res.rows;
    notes.push(...res.notes);
  } else {
    notes.push("This workbook has no Pipeline Meeting tab.");
  }

  const weekSheet = sheets.get(tab);
  const kpi = weekSheet
    ? readKpiRow(weekSheet, week)
    : { row: {}, found: false, notes: ["This week's tab couldn't be read."] };
  notes.push(...kpi.notes);

  // A blank cell in a KPI row that was found is a real 0.0. A row that wasn't
  // found is missing data — on a handout those look identical and mean
  // opposite things, so they must not render the same way.
  let metricValue: number | null = null;
  let metricFound = false;
  if (kpi.found) {
    for (const [label, value] of Object.entries(kpi.row)) {
      if (nameKey(label) === nameKey(metric)) {
        metricValue = value ?? 0;
        metricFound = true;
        break;
      }
    }
    if (!metricFound) notes.push(`'${metric}' isn't a column in this workbook's KPI table.`);
  }

  const billingRes = readBilling(billingSheet, recruiter, week.getUTCFullYear());
  notes.push(...billingRes.notes);

  // Phone: this week's rows for everyone, and this recruiter's own row.
  const goalSeconds = Math.round(phoneGoalHours * 3600);
  const weekPhone = phoneRows.filter((p) => isoDay(p.weekStart) === isoDay(week));
  let calls: Record<string, unknown> = { available: false };
  let phone: Record<string, unknown> = { available: false, goal_seconds: goalSeconds };
  let team: Record<string, unknown> = { available: false, goal_seconds: goalSeconds };

  if (weekPhone.length > 0) {
    const members = weekPhone
      .map((p) => ({ name: p.recruiter, avg_seconds: p.secondsAvgPerDay }))
      .sort((a, b) => b.avg_seconds - a.avg_seconds);
    const mine = weekPhone.find((p) => nameKey(p.recruiter).split(" ")[0] === nameKey(firstName));
    team = {
      available: true,
      goal_seconds: weekPhone[0].goalSeconds,
      members,
      highlight: mine ? mine.recruiter : null,
      of: members.length,
      ...(mine ? { rank: members.findIndex((m) => m.name === mine.recruiter) + 1 } : {}),
    };
    if (mine) {
      calls = {
        available: true,
        days: DAY_LABELS.map((d, i) => ({ day: d, count: mine.calls[i] })),
        avg_per_day: mine.callsAvgPerDay,
      };
      phone = {
        available: true,
        goal_seconds: mine.goalSeconds,
        days: DAY_LABELS.map((d, i) => ({ day: d, seconds: mine.seconds[i] })),
        avg_seconds: mine.secondsAvgPerDay,
        goal_met: mine.metGoal,
      };
    } else {
      notes.push(`${recruiter} isn't in the phone roster, so the personal phone cards are blank.`);
    }
  }

  return {
    generated_at: now.toISOString(),
    source: masterName,
    recruiter: { display_name: recruiter, first_name: firstName, firm },
    week_of: isoDay(week),
    week_tab: tab,
    current_as_of: isoDay(now),
    highlight_metric: {
      available: metricFound,
      label: metricLabel,
      value: metricValue,
      sublabel: `Week of ${fmtDatePadded(week)}`,
    },
    kpi_row: kpi.row,
    billing: billingRes.billing,
    pipeline,
    calls,
    phone,
    team_phone: team,
    notes,
  };
}
