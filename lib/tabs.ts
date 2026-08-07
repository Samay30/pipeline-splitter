/**
 * Tab classification for recruiter pipeline workbooks — v2, hardened
 * against real-world tab naming observed across all 11 recruiter
 * workbooks in production ("Jan 12" with no year, "Current February,
 * 16, 2026", "March30, 2026", "arch 23, 2026", "Februrary 23,2026",
 * "March 2nd, 2026 (current)", "February 9, 2026 (2)", trailing and
 * leading spaces, lowercase months, "March 30th. 2026", ...).
 *
 * v1 assumed weekly tabs reliably parse as dates and kept everything
 * that didn't. In production that inverted the trim for workbooks with
 * year-less tab names (nothing parsed -> nothing dropped -> recap as
 * big as the master). v2 flips the logic:
 *
 *   1. CORE tabs are recognized by fuzzy name matching against a small
 *      known set (Active Positions, Clients, Candidates, Pipeline
 *      Meeting, Master*, Sheet*). Always kept.
 *   2. EVERYTHING else is a weekly tab. Keep the last N.
 *   3. Ordering of weekly tabs: parsed date when available, otherwise
 *      inferred from sheet position (workbook order is chronological
 *      in practice — recruiters append new weeks). Position is the
 *      tiebreaker and the fallback, so unparseable names still sort
 *      correctly relative to their neighbors.
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
  // prefix match ("janu", "febr")
  for (const [name, idx] of Object.entries(MONTHS)) {
    if (name.length >= 3 && (w.startsWith(name) || name.startsWith(w)) && w.length >= 3) return idx;
  }
  // one-edit tolerance for typos like "arch" (March) or "Februrary"
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
 * Attempt to parse a weekly tab name into a Date. Extremely lenient:
 * strips noise words/suffixes, tolerates missing commas/spaces,
 * ordinals, periods, typo'd months, and a missing year (returns a
 * year-less sentinel handled by the caller via defaultYear).
 */
export function parseWeekTab(name: string, defaultYear?: number): Date | null {
  let s = name
    .trim()
    .toLowerCase()
    .replace(/\(.*?\)/g, " ")            // "(2)", "(current)"
    .replace(/\bcurrent\b/g, " ")         // "Current Feb 9, 2026"
    .replace(/(\d)(st|nd|rd|th)\b/g, "$1") // ordinals
    .replace(/[.,]/g, " ")                // periods and commas -> space
    .replace(/\s+/g, " ")
    .trim();

  // "march30 2026" -> "march 30 2026"
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
  if (d.getUTCMonth() !== month) return null; // impossible date rolled over
  return d;
}

/** Known non-weekly tabs, matched loosely (lowercased, trimmed, prefix). */
const CORE_PREFIXES = [
  "active po",        // "Active Positions", "Active Postitions" (typo)
  "client",
  "candidate",
  "pipeline meeting",
  "sheet",
  "master",           // "Master 2026 Pipeline"
  "phone",            // derived tab this job writes; never a weekly tab
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

  // Order weekly tabs chronologically. Primary key: parsed date.
  // Undated tabs inherit ordering from sheet position (workbooks are
  // chronological in practice); position also breaks date ties.
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