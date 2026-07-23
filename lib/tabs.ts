/**
 * Tab classification + KPI-grid extraction for recruiter pipeline
 * workbooks.
 *
 * Two responsibilities, both pure (no I/O) so they can be unit-tested
 * against fixtures:
 *
 *   1. Tab classification (planKeep / isCoreTab / parseWeekTab) — decides
 *      which sheets the recap keeps. Unchanged from v2.
 *
 *   2. KPI-grid extraction (buildKpiTable) — NEW. Reads the embedded
 *      running-KPI grid inside a single weekly tab and returns one clean
 *      row per week with all 8 metrics. This is what makes every metric
 *      (not just First/Second Interviews) reliably recoverable after the
 *      M365 connector flattens the file to text: the recap gets a
 *      denormalized KPI_TABLE sheet where each week is ONE self-contained
 *      row, so flattening can't scramble which number belongs to which
 *      week or metric.
 *
 * Why the grid can't just be copied verbatim: the recruiter's KPI area is
 * several side-by-side / stacked mini-tables (Week|First|Second in one
 * block, Multiple|Offers in another, Accepted|Submittals|Agreements|
 * Resumes in a third). Each block is week-indexed with the SAME ordering
 * and the same quarter/YTD subtotal rows, but they don't share a single
 * header row. Reading them through the Graph workbook API preserves exact
 * (row,col) coordinates, so we anchor each metric to ITS OWN header and
 * join every metric by ordinal week position. That survives whether the
 * blocks are laid out horizontally or stacked vertically.
 */

const MONTHS: Record<string, number> = {
  january: 0, february: 1, march: 2, april: 3, may: 4, june: 5,
  july: 6, august: 7, september: 8, october: 9, november: 10, december: 11,
  jan: 0, feb: 1, mar: 2, apr: 3, jun: 5, jul: 6, aug: 7, sep: 8,
  sept: 8, oct: 9, nov: 10, dec: 11,
};

/** Fuzzy month lookup: exact, prefix, or single-typo tolerant. */
function monthIndex(raw: string): number | null {
  const w = raw.toLowerCase();
  if (w in MONTHS) return MONTHS[w];
  for (const [name, idx] of Object.entries(MONTHS)) {
    if (name.length >= 3 && (w.startsWith(name) || name.startsWith(w)) && w.length >= 3) return idx;
  }
  for (const [name, idx] of Object.entries(MONTHS)) {
    if (name.length < 4) continue;
    if (editDistanceLeq1ish(w, name)) return idx;
  }
  return null;
}

/** Cheap check: within edit distance 1, or one char inserted/deleted/substituted. */
function editDistanceLeq1ish(a: string, b: string): boolean {
  if (Math.abs(a.length - b.length) > 1) return false;
  if (a === b) return true;
  const [s, l] = a.length <= b.length ? [a, b] : [b, a];
  let i = 0, j = 0, edits = 0;
  while (i < s.length && j < l.length) {
    if (s[i] === l[j]) { i++; j++; continue; }
    edits++;
    if (edits > 1) return false;
    if (s.length === l.length) { i++; j++; } else { j++; }
  }
  return edits + (l.length - j) + (s.length - i) <= 1;
}

/**
 * Attempt to parse a weekly tab NAME into a Date. Lenient: strips noise
 * words/suffixes, tolerates missing commas/spaces, ordinals, periods,
 * typo'd months, and a missing year (uses defaultYear).
 */
export function parseWeekTab(name: string, defaultYear?: number): Date | null {
  let s = name
    .trim()
    .toLowerCase()
    .replace(/\(.*?\)/g, " ")
    .replace(/\bcurrent\b/g, " ")
    .replace(/(\d)(st|nd|rd|th)\b/g, "$1")
    .replace(/[.,]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  s = s.replace(/([a-z])(\d)/g, "$1 $2");

  const m = s.match(/^([a-z]+)\s+(\d{1,2})(?:\s+(\d{4}))?$/);
  if (!m) return null;
  const month = monthIndex(m[1]);
  if (month === null) return null;
  const day = Number(m[2]);
  if (day < 1 || day > 31) return null;
  const year = m[3] ? Number(m[3]) : defaultYear;
  if (year === undefined) return null;
  const d = new Date(Date.UTC(year, month, day));
  if (d.getUTCMonth() !== month) return null;
  return d;
}

/** Known non-weekly tabs, matched loosely (lowercased, trimmed, prefix). */
const CORE_PREFIXES = [
  "active po",
  "client",
  "candidate",
  "pipeline meeting",
  "sheet",
  "master",
  "kpi_table", // never treat our own derived sheet as a weekly tab
];

export function isCoreTab(name: string): boolean {
  const n = name.trim().toLowerCase();
  return CORE_PREFIXES.some((p) => n.startsWith(p));
}

export interface KeepPlan {
  keep: string[];
  dropped: string[];
}

/** Extract a default year from context, e.g. filename "… Pipeline 2026.xlsx". */
export function inferYear(masterName: string, sheetNames: string[]): number | undefined {
  const fromFile = masterName.match(/\b(20\d{2})\b/);
  if (fromFile) return Number(fromFile[1]);
  for (const n of sheetNames) {
    const m = n.match(/\b(20\d{2})\b/);
    if (m) return Number(m[1]);
  }
  return undefined;
}

/**
 * Decide which tabs the recap keeps.
 * @param sheetNames sheets in workbook position order
 * @param weeksToKeep how many most-recent weekly tabs to retain
 * @param defaultYear year assumed for year-less tab names
 */
export function planKeep(sheetNames: string[], weeksToKeep: number, defaultYear?: number): KeepPlan {
  const weekly: Array<{ name: string; position: number; date: Date | null }> = [];
  for (let i = 0; i < sheetNames.length; i++) {
    const name = sheetNames[i];
    if (isCoreTab(name)) continue;
    weekly.push({ name, position: i, date: parseWeekTab(name, defaultYear) });
  }

  const sorted = [...weekly].sort((a, b) => {
    if (a.date && b.date && a.date.getTime() !== b.date.getTime()) {
      return a.date.getTime() - b.date.getTime();
    }
    return a.position - b.position;
  });

  const keepWeekly = new Set(sorted.slice(-weeksToKeep).map((w) => w.name + "@" + w.position));
  const keep: string[] = [];
  const dropped: string[] = [];
  for (let i = 0; i < sheetNames.length; i++) {
    const name = sheetNames[i];
    if (isCoreTab(name) || keepWeekly.has(name + "@" + i)) keep.push(name);
    else dropped.push(name);
  }
  return { keep, dropped };
}

/**
 * Given the kept sheet names, return the newest weekly tab (the one whose
 * embedded KPI grid is most up to date). Its running KPI table already
 * contains every week of the year, so we only need to parse this one tab
 * to build the full KPI_TABLE.
 */
export function newestWeeklyTab(keep: string[], defaultYear?: number): string | null {
  const weekly = keep
    .map((name, i) => ({ name, position: i, date: parseWeekTab(name, defaultYear) }))
    .filter((_, i) => !isCoreTab(keep[i]));
  if (weekly.length === 0) return null;
  weekly.sort((a, b) => {
    if (a.date && b.date && a.date.getTime() !== b.date.getTime()) {
      return a.date.getTime() - b.date.getTime();
    }
    return a.position - b.position;
  });
  return weekly[weekly.length - 1].name;
}

/* ======================================================================
 * KPI-grid extraction
 * ==================================================================== */

/** Canonical, exact (normalized) header labels for the 8 metrics + Week. */
const KPI_LABELS = {
  week: "week",
  firstInterviews: "first interviews",
  secondInterviews: "second interviews",
  multipleInterviews: "multiple interviews",
  submittals: "submittals",
  offers: "offers",
  acceptedOffers: "accepted offers",
  agreementsExecuted: "agreements executed",
  resumes: "resumes",
} as const;

type MetricKey = Exclude<keyof typeof KPI_LABELS, "week">;

const METRIC_KEYS: MetricKey[] = [
  "firstInterviews",
  "secondInterviews",
  "multipleInterviews",
  "submittals",
  "offers",
  "acceptedOffers",
  "agreementsExecuted",
  "resumes",
];

export interface KpiRow {
  week: string; // ISO yyyy-mm-dd
  firstInterviews: number | null;
  secondInterviews: number | null;
  multipleInterviews: number | null;
  submittals: number | null;
  offers: number | null;
  acceptedOffers: number | null;
  agreementsExecuted: number | null;
  resumes: number | null;
}

export interface KpiDiagnostics {
  weekHeader: { row: number; col: number } | null;
  /** Per metric: the header cell chosen (0-based, grid-local), or null. */
  found: Record<string, { row: number; col: number } | null>;
  /** Per metric: how many candidate header matches were seen. */
  matchCounts: Record<string, number>;
  weekSlots: number; // total ordinal slots read from the Week column
  weekRows: number; // of those, how many were real week dates
  notes: string[];
}

export interface KpiExtraction {
  rows: KpiRow[];
  diagnostics: KpiDiagnostics;
}

function norm(v: unknown): string {
  return String(v ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[:]+$/, "");
}

/** Excel serial date -> ISO yyyy-mm-dd (UTC). 25569 = days 1899-12-30→1970-01-01. */
function serialToISO(n: number): string | null {
  if (!isFinite(n) || n < 1 || n > 60000) return null;
  const ms = Math.round((n - 25569) * 86400 * 1000);
  const d = new Date(ms);
  if (isNaN(d.getTime())) return null;
  return d.toISOString().slice(0, 10);
}

/** Try to read a cell as a date, using its number format as a hint. */
function cellToISO(value: unknown, fmt: string | undefined): string | null {
  const isDateFmt = !!fmt && /[dmy]/i.test(fmt) && fmt !== "General";
  if (typeof value === "number" && isDateFmt) return serialToISO(value);
  if (typeof value === "number" && value > 40000 && value < 50000) return serialToISO(value);
  if (typeof value === "string") {
    const s = value.trim();
    let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
      const d = new Date(Date.UTC(Number(m[3]), Number(m[1]) - 1, Number(m[2])));
      if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
    }
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  }
  return null;
}

/** Coerce a metric cell to a number, or null. Keeps fractional split credits. */
function cellToNumber(value: unknown): number | null {
  if (typeof value === "number" && isFinite(value)) return value;
  if (typeof value === "string") {
    const s = value.trim();
    if (s === "") return null;
    const n = Number(s);
    if (isFinite(n)) return n;
  }
  return null;
}

/**
 * Build the denormalized KPI table from ONE weekly tab's used-range grid.
 *
 * @param values      2D values, grid-local indices ([0][0] = used-range top-left)
 * @param numberFormat 2D number formats, same shape as values
 *
 * Strategy:
 *   1. Find the topmost cell whose text == "week"  → the Week column.
 *   2. Read down the Week column, recording every slot in order (week
 *      dates AND total rows) until the YTD total, inclusive. Totals keep
 *      their ordinal slot so metric columns stay aligned.
 *   3. For each metric, find its header (exact label match, at/below the
 *      Week header row) and read its column from ITS OWN header downward.
 *      Ordinal k in the Week sequence == ordinal k below the metric header.
 *   4. Emit one row per Week slot that is a real date.
 *
 * Any metric whose header can't be found is left null — never fabricated.
 */
export function buildKpiTable(values: unknown[][], numberFormat: string[][] = []): KpiExtraction {
  const notes: string[] = [];
  const found: KpiDiagnostics["found"] = {};
  const matchCounts: KpiDiagnostics["matchCounts"] = {};
  for (const k of METRIC_KEYS) {
    found[k] = null;
    matchCounts[k] = 0;
  }

  const nRows = values.length;
  const nCols = values.reduce((m, r) => Math.max(m, r?.length ?? 0), 0);
  const fmt = (r: number, c: number): string | undefined => numberFormat?.[r]?.[c];

  // 1. Locate the Week header (topmost, leftmost exact "week").
  let weekHeader: { row: number; col: number } | null = null;
  for (let r = 0; r < nRows && !weekHeader; r++) {
    for (let c = 0; c < nCols; c++) {
      if (norm(values[r]?.[c]) === KPI_LABELS.week) {
        weekHeader = { row: r, col: c };
        break;
      }
    }
  }

  if (!weekHeader) {
    notes.push('No "Week" header found — KPI grid could not be located.');
    return {
      rows: [],
      diagnostics: { weekHeader: null, found, matchCounts, weekSlots: 0, weekRows: 0, notes },
    };
  }

  // 2. Read the Week column into ordinal slots.
  type Slot = { kind: "week" | "total" | "other"; iso?: string };
  const slots: Slot[] = [];
  let trailingBlanks = 0;
  for (let r = weekHeader.row + 1; r < nRows; r++) {
    const raw = values[r]?.[weekHeader.col];
    const n = norm(raw);
    if (n.includes("ytd")) {
      slots.push({ kind: "total" });
      break; // YTD is the last row of the grid
    }
    if (n.includes("total")) {
      slots.push({ kind: "total" });
      trailingBlanks = 0;
      continue;
    }
    if (n === "") {
      trailingBlanks++;
      if (trailingBlanks >= 3) break; // grid has ended
      slots.push({ kind: "other" });
      continue;
    }
    trailingBlanks = 0;
    const iso = cellToISO(raw, fmt(r, weekHeader.col));
    slots.push(iso ? { kind: "week", iso } : { kind: "other" });
  }
  while (slots.length && slots[slots.length - 1].kind === "other") slots.pop();

  const L = slots.length;

  // 3. Locate each metric header and read its column by ordinal offset.
  const metricValues: Record<MetricKey, Array<number | null>> = {
    firstInterviews: [], secondInterviews: [], multipleInterviews: [], submittals: [],
    offers: [], acceptedOffers: [], agreementsExecuted: [], resumes: [],
  };

  for (const key of METRIC_KEYS) {
    const label = KPI_LABELS[key];
    let chosen: { row: number; col: number } | null = null;
    for (let r = weekHeader.row; r < nRows; r++) {
      for (let c = 0; c < nCols; c++) {
        if (norm(values[r]?.[c]) === label) {
          matchCounts[key]++;
          if (!chosen) chosen = { row: r, col: c }; // topmost wins
        }
      }
    }
    found[key] = chosen;
    if (!chosen) {
      notes.push(`Metric "${label}" header not found — column left blank.`);
      metricValues[key] = new Array(L).fill(null);
      continue;
    }
    const col = chosen.col;
    const origin = chosen.row;
    const out: Array<number | null> = [];
    for (let k = 0; k < L; k++) {
      const r = origin + 1 + k;
      out.push(r < nRows ? cellToNumber(values[r]?.[col]) : null);
    }
    metricValues[key] = out;
  }

  // 4. Emit one row per real week slot.
  const rows: KpiRow[] = [];
  let weekRows = 0;
  for (let k = 0; k < L; k++) {
    if (slots[k].kind !== "week" || !slots[k].iso) continue;
    weekRows++;
    rows.push({
      week: slots[k].iso as string,
      firstInterviews: metricValues.firstInterviews[k] ?? null,
      secondInterviews: metricValues.secondInterviews[k] ?? null,
      multipleInterviews: metricValues.multipleInterviews[k] ?? null,
      submittals: metricValues.submittals[k] ?? null,
      offers: metricValues.offers[k] ?? null,
      acceptedOffers: metricValues.acceptedOffers[k] ?? null,
      agreementsExecuted: metricValues.agreementsExecuted[k] ?? null,
      resumes: metricValues.resumes[k] ?? null,
    });
  }

  return {
    rows,
    diagnostics: { weekHeader, found, matchCounts, weekSlots: L, weekRows, notes },
  };
}

/** Column order for the emitted KPI_TABLE sheet (header row). */
export const KPI_TABLE_HEADERS = [
  "Week",
  "First Interviews",
  "Second Interviews",
  "Multiple Interviews",
  "Submittals",
  "Offers",
  "Accepted Offers",
  "Agreements Executed",
  "Resumes",
];

/** Turn a KpiRow into a plain array matching KPI_TABLE_HEADERS order. */
export function kpiRowToArray(r: KpiRow): Array<string | number | null> {
  return [
    r.week,
    r.firstInterviews,
    r.secondInterviews,
    r.multipleInterviews,
    r.submittals,
    r.offers,
    r.acceptedOffers,
    r.agreementsExecuted,
    r.resumes,
  ];
}