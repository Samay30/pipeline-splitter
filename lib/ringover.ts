/**
 * Ringover -> weekly phone roll-up for the PHONE tab of each recap workbook.
 *
 * Why this lives here and not in the skill: Claude's sandbox can't reach
 * public-api.ringover.com (egress allowlist), and a skill folder is the wrong
 * place for an API key. This job already runs on a schedule with credentials,
 * so it pulls the numbers once for the whole firm and writes them next to the
 * pipeline data the skill already reads.
 *
 * Everything here is best-effort. If Ringover is down, misconfigured, or
 * simply not set up yet, `collectPhoneWeeks` returns a result with an `error`
 * and no rows — the recap sync carries on and the PHONE tab is omitted. A
 * phone outage must never cost us the recap files.
 *
 * API shape (v2):
 *   GET https://public-api.ringover.com/v2/calls
 *       ?start_date=&end_date=&limit_count=1000&limit_offset=0
 *   Header: Authorization: <raw key>        (NOT "Bearer <key>")
 *   Returns { total_call_count, call_list: [...] }, or 204 for an empty window.
 */

const DEFAULT_BASE_URL = "https://public-api.ringover.com/v2";
const PAGE_SIZE = 1000;
const MAX_PAGES = 200;

export const DAY_LABELS = ["Mon", "Tue", "Wed", "Thu", "Fri"] as const;

export interface PhoneRow {
  /** Monday of the week, as a real Date so Excel keeps it a date. */
  weekStart: Date;
  recruiter: string;
  /** Index 0 = Monday .. 4 = Friday. */
  calls: number[];
  seconds: number[];
  callsTotal: number;
  secondsTotal: number;
  /** Always total / 5. A quiet Monday should pull the average down. */
  callsAvgPerDay: number;
  secondsAvgPerDay: number;
  goalSeconds: number;
  metGoal: boolean;
}

export interface PhoneResult {
  rows: PhoneRow[];
  weeks: string[];
  rosterSize: number;
  callsPulled: number;
  /** Ringover users seen on calls who aren't in the roster. */
  unmappedUsers: string[];
  ms: number;
  error?: string;
  skipped?: string;
}

/* -------------------------------------------------------------------------- */
/* config                                                                     */
/* -------------------------------------------------------------------------- */

/** Parse RINGOVER_ROSTER: {"aaronr@eggersesearch.com":"Aaron", ...} */
export function parseRoster(raw: string | undefined): Map<string, string> {
  const map = new Map<string, string>();
  if (!raw || !raw.trim()) return map;
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error("RINGOVER_ROSTER is not valid JSON");
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new Error('RINGOVER_ROSTER must be an object of {"email":"Display Name"}');
  }
  for (const [email, name] of Object.entries(parsed as Record<string, unknown>)) {
    if (typeof name === "string" && name.trim()) {
      map.set(email.trim().toLowerCase(), name.trim());
    }
  }
  return map;
}

/** Monday (UTC) of the week containing `d`. */
export function mondayOf(d: Date): Date {
  const copy = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  const dow = copy.getUTCDay(); // 0 = Sunday
  const shift = dow === 0 ? -6 : 1 - dow;
  copy.setUTCDate(copy.getUTCDate() + shift);
  return copy;
}

function addDays(d: Date, n: number): Date {
  const copy = new Date(d);
  copy.setUTCDate(copy.getUTCDate() + n);
  return copy;
}

function isoDay(d: Date): string {
  return d.toISOString().slice(0, 10);
}

/* -------------------------------------------------------------------------- */
/* HTTP                                                                       */
/* -------------------------------------------------------------------------- */

interface RingoverPage {
  total_call_count?: number;
  call_list?: unknown[];
  calls?: unknown[];
}

async function fetchCalls(
  apiKey: string,
  baseUrl: string,
  start: Date,
  end: Date
): Promise<unknown[]> {
  const all: unknown[] = [];
  let offset = 0;

  for (let page = 0; page < MAX_PAGES; page++) {
    const url =
      `${baseUrl}/calls?start_date=${encodeURIComponent(start.toISOString())}` +
      `&end_date=${encodeURIComponent(end.toISOString())}` +
      `&limit_count=${PAGE_SIZE}&limit_offset=${offset}`;

    const res = await fetch(url, {
      headers: { Authorization: apiKey, Accept: "application/json" },
    });

    if (res.status === 204) break; // documented "no calls in window"
    if (res.status === 429) {
      const wait = Number(res.headers.get("Retry-After")) || 5;
      await new Promise((r) => setTimeout(r, wait * 1000));
      continue;
    }
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`Ringover ${res.status} on /calls: ${body.slice(0, 300)}`);
    }

    const json = (await res.json()) as RingoverPage | unknown[];
    const list: unknown[] = Array.isArray(json)
      ? json
      : (json.call_list ?? json.calls ?? []);
    all.push(...list);

    const total = Array.isArray(json)
      ? all.length
      : Number(json.total_call_count ?? all.length);
    offset += list.length;
    if (list.length === 0 || list.length < PAGE_SIZE || offset >= total) break;
  }

  return all;
}

/* -------------------------------------------------------------------------- */
/* normalisation                                                              */
/* -------------------------------------------------------------------------- */

/**
 * Ringover's call objects vary a little by plan, so every field is read as
 * "first of these keys that has a value". If a field ever comes back empty,
 * add the real key here — this function is the only place names appear.
 */
interface NormalCall {
  email: string;
  direction: "in" | "out" | null;
  at: Date | null;
  seconds: number;
}

function pickString(obj: Record<string, unknown>, keys: string[]): string {
  for (const k of keys) {
    const v = obj[k];
    if (typeof v === "string" && v.trim()) return v.trim();
  }
  return "";
}

function pickSeconds(obj: Record<string, unknown>, keys: string[]): number {
  for (const k of keys) {
    const v = obj[k];
    if (typeof v === "number" && Number.isFinite(v) && v >= 0) return Math.round(v);
    if (typeof v === "string" && v.trim()) {
      if (v.includes(":")) {
        const parts = v.split(":").map((p) => Number(p));
        if (parts.every((p) => Number.isFinite(p))) {
          return parts.reduce((acc, p) => acc * 60 + p, 0);
        }
      }
      const n = Number(v);
      if (Number.isFinite(n) && n >= 0) return Math.round(n);
    }
  }
  return 0;
}

function normalize(raw: unknown): NormalCall {
  const call = (raw ?? {}) as Record<string, unknown>;
  const user = (call.user ?? call.agent ?? call.user_data ?? {}) as Record<string, unknown>;

  const email = (
    pickString(user, ["email"]) || pickString(call, ["user_email", "agent_email"])
  ).toLowerCase();

  const dirRaw = pickString(call, ["direction", "type", "way"]).toLowerCase();
  const direction = dirRaw.includes("out") ? "out" : dirRaw.includes("in") ? "in" : null;

  const whenRaw = pickString(call, ["start_time", "started_at", "start_date", "date"]);
  const at = whenRaw ? new Date(whenRaw) : null;

  return {
    email,
    direction,
    // incall_duration is talk time rather than ringing time, which is what the
    // KPI sheet is measuring.
    seconds: pickSeconds(call, [
      "incall_duration",
      "talk_duration",
      "total_duration",
      "duration",
    ]),
    at: at && !Number.isNaN(at.valueOf()) ? at : null,
  };
}

/* -------------------------------------------------------------------------- */
/* roll-up                                                                    */
/* -------------------------------------------------------------------------- */

export interface CollectOptions {
  /** How many weeks back, including the current one. */
  weeks: number;
  /** Daily phone-time goal, in hours. */
  goalHours: number;
  /** Directions counted as an "outbound call". */
  countDirections: Set<string>;
  /** Directions whose talk time counts toward phone time. */
  timeDirections: Set<string>;
  /** Treated as "now" — injectable for tests. */
  now?: Date;
}

export function optionsFromEnv(): CollectOptions {
  const dirs = (v: string | undefined, fallback: string) =>
    new Set(
      (v || fallback)
        .split(",")
        .map((s) => s.trim().toLowerCase())
        .filter(Boolean)
    );
  return {
    weeks: Math.max(1, Number(process.env.PHONE_WEEKS_TO_KEEP || "4")),
    goalHours: Number(process.env.PHONE_DAILY_GOAL_HOURS || "3"),
    countDirections: dirs(process.env.PHONE_COUNT_DIRECTIONS, "out"),
    timeDirections: dirs(process.env.PHONE_TIME_DIRECTIONS, "out,in"),
  };
}

/**
 * Pull the last N weeks of calls once for the whole firm and roll them into
 * one row per recruiter per week. Never throws — failures come back on
 * `error` so the caller can carry on without the PHONE tab.
 */
export async function collectPhoneWeeks(opts?: Partial<CollectOptions>): Promise<PhoneResult> {
  const t0 = Date.now();
  const empty: PhoneResult = {
    rows: [],
    weeks: [],
    rosterSize: 0,
    callsPulled: 0,
    unmappedUsers: [],
    ms: 0,
  };

  const apiKey = (process.env.RINGOVER_API_KEY || "").trim();
  if (!apiKey) {
    return { ...empty, skipped: "RINGOVER_API_KEY not set — PHONE tab omitted" };
  }

  let roster: Map<string, string>;
  try {
    roster = parseRoster(process.env.RINGOVER_ROSTER);
  } catch (err) {
    return { ...empty, error: err instanceof Error ? err.message : String(err) };
  }
  if (roster.size === 0) {
    return { ...empty, skipped: "RINGOVER_ROSTER not set — PHONE tab omitted" };
  }

  const cfg = { ...optionsFromEnv(), ...opts };
  const baseUrl = (process.env.RINGOVER_BASE_URL || DEFAULT_BASE_URL).replace(/\/$/, "");
  const now = cfg.now ?? new Date();

  const thisMonday = mondayOf(now);
  const firstMonday = addDays(thisMonday, -7 * (cfg.weeks - 1));
  const lastFriday = addDays(thisMonday, 4);

  let calls: unknown[];
  try {
    calls = await fetchCalls(
      apiKey,
      baseUrl,
      new Date(`${isoDay(firstMonday)}T00:00:00.000Z`),
      new Date(`${isoDay(lastFriday)}T23:59:59.999Z`)
    );
  } catch (err) {
    return {
      ...empty,
      rosterSize: roster.size,
      ms: Date.now() - t0,
      error: err instanceof Error ? err.message : String(err),
    };
  }

  // weekKey -> email -> per-weekday buckets
  const buckets = new Map<string, Map<string, { calls: number[]; seconds: number[] }>>();
  const weekKeys: string[] = [];
  for (let i = 0; i < cfg.weeks; i++) {
    const key = isoDay(addDays(firstMonday, 7 * i));
    weekKeys.push(key);
    const perPerson = new Map<string, { calls: number[]; seconds: number[] }>();
    for (const email of roster.keys()) {
      perPerson.set(email, { calls: [0, 0, 0, 0, 0], seconds: [0, 0, 0, 0, 0] });
    }
    buckets.set(key, perPerson);
  }

  const unmapped = new Set<string>();
  for (const raw of calls) {
    const c = normalize(raw);
    if (!c.at) continue;

    const dow = c.at.getUTCDay();
    if (dow === 0 || dow === 6) continue; // the sheet has five columns
    const dayIndex = dow - 1;

    const weekKey = isoDay(mondayOf(c.at));
    const perPerson = buckets.get(weekKey);
    if (!perPerson) continue; // outside the requested window

    const slot = perPerson.get(c.email);
    if (!slot) {
      if (c.email) unmapped.add(c.email);
      continue; // never attribute a call to the wrong person
    }

    if (c.direction && cfg.countDirections.has(c.direction)) slot.calls[dayIndex] += 1;
    if (c.direction && cfg.timeDirections.has(c.direction)) slot.seconds[dayIndex] += c.seconds;
  }

  const goalSeconds = Math.round(cfg.goalHours * 3600);
  const rows: PhoneRow[] = [];
  for (const weekKey of weekKeys) {
    const perPerson = buckets.get(weekKey)!;
    for (const [email, name] of roster) {
      const slot = perPerson.get(email)!;
      const callsTotal = slot.calls.reduce((a, b) => a + b, 0);
      const secondsTotal = slot.seconds.reduce((a, b) => a + b, 0);
      const secondsAvg = Math.round(secondsTotal / 5);
      rows.push({
        weekStart: new Date(`${weekKey}T00:00:00.000Z`),
        recruiter: name,
        calls: slot.calls,
        seconds: slot.seconds,
        callsTotal,
        secondsTotal,
        callsAvgPerDay: Math.round(callsTotal / 5),
        secondsAvgPerDay: secondsAvg,
        goalSeconds,
        metGoal: secondsAvg >= goalSeconds,
      });
    }
  }

  return {
    rows,
    weeks: weekKeys,
    rosterSize: roster.size,
    callsPulled: calls.length,
    unmappedUsers: [...unmapped].sort(),
    ms: Date.now() - t0,
  };
}
