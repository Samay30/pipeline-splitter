/**
 * Tab classification for recruiter pipeline workbooks.
 *
 * Convention observed in real EES workbooks: weekly tabs are named like
 *   "May 25, 2026" / "June 1 2026" / "July 20, 2026"
 * (comma after the day is optional — recruiters are inconsistent).
 * Everything else ("Active Positions", "Clients", "Candidates",
 * "Pipeline Meeting", stray "Sheet1"s) is a non-date tab.
 *
 * Keep rule: ALL non-date tabs + the N most recent weekly tabs.
 * Non-date tabs are few and small, so keeping them unconditionally is
 * safer than maintaining an allow-list that breaks when a recruiter
 * adds a tab we didn't anticipate.
 */

const MONTHS: Record<string, number> = {
  january: 0, february: 1, march: 2, april: 3, may: 4, june: 5,
  july: 6, august: 7, september: 8, october: 9, november: 10, december: 11,
};

const WEEK_TAB_RE = /^([A-Za-z]+)\.?\s+(\d{1,2})(?:st|nd|rd|th)?\s*,?\s+(\d{4})$/;

/** Parse a weekly tab name into a Date, or null if it isn't a weekly tab. */
export function parseWeekTab(name: string): Date | null {
  const m = name.trim().match(WEEK_TAB_RE);
  if (!m) return null;
  const month = MONTHS[m[1].toLowerCase()];
  if (month === undefined) return null;
  const day = Number(m[2]);
  const year = Number(m[3]);
  if (day < 1 || day > 31) return null;
  const d = new Date(Date.UTC(year, month, day));
  // Reject impossible dates like "February 30, 2026"
  if (d.getUTCMonth() !== month || d.getUTCDate() !== day) return null;
  return d;
}

export interface KeepPlan {
  /** Sheet names to copy into the recap workbook, in original workbook order. */
  keep: string[];
  /** Weekly tabs dropped (oldest first), for reporting. */
  dropped: string[];
}

/**
 * Given all sheet names in workbook order, decide what the recap keeps.
 * @param sheetNames sheets in position order
 * @param weeksToKeep how many most-recent weekly tabs to retain
 */
export function planKeep(sheetNames: string[], weeksToKeep: number): KeepPlan {
  const weekly: Array<{ name: string; date: Date }> = [];
  for (const name of sheetNames) {
    const d = parseWeekTab(name);
    if (d) weekly.push({ name, date: d });
  }
  weekly.sort((a, b) => a.date.getTime() - b.date.getTime());
  const keepWeekly = new Set(weekly.slice(-weeksToKeep).map((w) => w.name));
  const droppedWeekly = weekly.slice(0, Math.max(0, weekly.length - weeksToKeep)).map((w) => w.name);

  const keep = sheetNames.filter((n) => parseWeekTab(n) === null || keepWeekly.has(n));
  return { keep, dropped: droppedWeekly };
}
