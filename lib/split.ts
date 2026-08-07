/**
 * Reads a master pipeline workbook through the Graph *workbook API*
 * (per-sheet JSON — no renderer size cap, unlike the text-extraction
 * renderer the Claude connector uses), then writes a slim values-only
 * copy ("recap file") back to the same SharePoint folder.
 *
 * The recap file mirrors the master's most recent weekly tabs plus all
 * non-date tabs, preserving cell values and number formats (so dates
 * stay real dates). Formulas are flattened to their computed values —
 * intentional: omitted sheets would otherwise leave broken references,
 * and the recap is a read-only derived artifact anyway.
 *
 * It also appends a PHONE tab holding the firm's weekly Ringover roll-up
 * (see lib/ringover.ts). Every recap carries the same PHONE tab: the KPI
 * sheet's team-comparison card needs everyone's numbers, and the tab is a few
 * dozen rows, so duplicating it is cheaper than making the skill open a second
 * file. When no phone data is available the tab is simply omitted.
 */

import ExcelJS from "exceljs";
import { graphFetch, graphJson, uploadFile } from "./graph";
import { inferYear, planKeep } from "./tabs";
import { DAY_LABELS, type PhoneRow } from "./ringover";
import type { SheetData, SheetMap } from "./kpi";

const CHUNK_ROWS = 500;

/**
 * Every sheet in the recap gets this marker injected as its first row.
 * The M365 connector's text extraction flattens all sheets into one blob
 * and strips sheet names; this banner puts the tab's identity into the
 * cell data itself, where flattening can't lose it.
 */
const SHEET_BANNER = (name: string) => `=== SHEET: ${name} ===`;

/** Name of the derived phone tab. Kept in sync with the skill's reader. */
export const PHONE_SHEET_NAME = "PHONE";

interface RangeMeta {
  address: string;
  rowCount: number;
  columnCount: number;
}

interface RangeData {
  values: unknown[][];
  numberFormat: string[][];
}

export interface FileReport {
  master: string;
  recap: string;
  keptTabs: string[];
  droppedWeeklyTabs: string[];
  recapBytes: number;
  ms: number;
  phoneRows?: number;
  error?: string;
}

/** splitWorkbook's result plus the sheet data it read, so the KPI payload can
 *  be built without a second pass over Graph. */
export interface SplitResult {
  report: FileReport;
  sheets: SheetMap;
}

function wsUrl(driveId: string, itemId: string, sheetName: string, tail: string): string {
  // Single quotes inside sheet names must be doubled per OData rules.
  const escaped = sheetName.replace(/'/g, "''");
  return `/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(escaped)}')${tail}`;
}

function colLetter(n: number): string {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Parse "A1:P108" (sheet prefix already stripped) into numeric bounds. */
function parseA1Range(a1: string): { r1: number; c1: number; r2: number; c2: number } {
  const cellRe = /([A-Z]+)(\d+)/g;
  const cells: Array<{ col: number; row: number }> = [];
  let m: RegExpExecArray | null;
  while ((m = cellRe.exec(a1)) !== null) {
    let col = 0;
    for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
    cells.push({ col, row: Number(m[2]) });
  }
  if (cells.length === 0) throw new Error(`Unparseable range address: ${a1}`);
  const first = cells[0];
  const last = cells[cells.length - 1];
  return { r1: first.row, c1: first.col, r2: last.row, c2: last.col };
}

/** Seconds -> "h:mm:ss", for the one human-readable column on the PHONE tab. */
function hms(seconds: number): string {
  const s = Math.max(0, Math.round(seconds));
  const mm = String(Math.floor((s % 3600) / 60)).padStart(2, "0");
  const ss = String(s % 60).padStart(2, "0");
  return `${Math.floor(s / 3600)}:${mm}:${ss}`;
}

const PHONE_HEADERS = [
  "Week Starting",
  "Recruiter",
  ...DAY_LABELS.map((d) => `${d} Calls`),
  "Total Calls",
  "Avg Calls/Day",
  ...DAY_LABELS.map((d) => `${d} Seconds`),
  "Total Seconds",
  "Avg Seconds/Day",
  "Avg Time/Day",
  "Daily Goal Seconds",
  "Met Goal",
];

/**
 * Write the phone roll-up as a plain rectangular table. Durations are stored
 * as raw seconds (numbers) so nothing has to round-trip through Excel time
 * formatting; the single "Avg Time/Day" text column exists purely so a human
 * opening the file can read it at a glance.
 */
export function addPhoneSheet(out: ExcelJS.Workbook, rows: PhoneRow[]): void {
  const ws = out.addWorksheet(PHONE_SHEET_NAME);
  ws.addRow([SHEET_BANNER(PHONE_SHEET_NAME)]);
  ws.addRow(PHONE_HEADERS);

  // Newest week first: the skill wants the current week, and so does a human
  // opening the file.
  const ordered = [...rows].sort(
    (a, b) =>
      b.weekStart.getTime() - a.weekStart.getTime() ||
      a.recruiter.localeCompare(b.recruiter)
  );

  for (const r of ordered) {
    const row = ws.addRow([
      r.weekStart,
      r.recruiter,
      ...r.calls,
      r.callsTotal,
      r.callsAvgPerDay,
      ...r.seconds,
      r.secondsTotal,
      r.secondsAvgPerDay,
      hms(r.secondsAvgPerDay),
      r.goalSeconds,
      r.metGoal ? "Yes" : "No",
    ]);
    row.getCell(1).numFmt = "yyyy-mm-dd";
  }

  ws.getColumn(1).width = 14;
  ws.getColumn(2).width = 16;
  for (let c = 3; c <= PHONE_HEADERS.length; c++) {
    ws.getColumn(c).width = Math.max(12, PHONE_HEADERS[c - 1].length + 2);
  }
}

export async function splitWorkbook(opts: {
  driveId: string;
  folderPath: string;
  itemId: string;
  masterName: string;
  recapSuffix: string;
  weeksToKeep: number;
  dryRun: boolean;
  /** Firm-wide phone roll-up. Omit or pass [] to skip the PHONE tab. */
  phoneRows?: PhoneRow[];
}): Promise<SplitResult> {
  const t0 = Date.now();
  const {
    driveId,
    folderPath,
    itemId,
    masterName,
    recapSuffix,
    weeksToKeep,
    dryRun,
    phoneRows = [],
  } = opts;
  const recapName = masterName.replace(/\.xlsx$/i, "") + recapSuffix + ".xlsx";
  const base = `/drives/${driveId}/items/${itemId}/workbook`;

  // Read-only workbook session: better perf + consistent snapshot across calls.
  const session = await graphJson<{ id: string }>(`${base}/createSession`, {
    method: "POST",
    body: JSON.stringify({ persistChanges: false }),
  });
  const sh = { "workbook-session-id": session.id };

  try {
    const sheets = await graphJson<{ value: Array<{ name: string; position: number }> }>(
      `${base}/worksheets?$select=name,position`,
      {},
      sh
    );
    const ordered = sheets.value.sort((a, b) => a.position - b.position).map((s) => s.name);
    const plan = planKeep(ordered, weeksToKeep, inferYear(masterName, ordered));

    const out = new ExcelJS.Workbook();
    out.creator = "eggers-internal pipeline-sync";
    out.created = new Date();
    const collected: SheetMap = new Map();

    for (const sheetName of plan.keep) {
      const ws = out.addWorksheet(sheetName);

      // Discover used range; a truly empty sheet 404s or returns A1 only.
      let meta: RangeMeta;
      try {
        meta = await graphJson<RangeMeta>(
          wsUrl(driveId, itemId, sheetName, `/usedRange?$select=address,rowCount,columnCount`),
          {},
          sh
        );
      } catch {
        continue; // empty sheet — keep it as a blank tab
      }
      const a1 = meta.address.includes("!") ? meta.address.slice(meta.address.lastIndexOf("!") + 1) : meta.address;
      const { r1, c1, r2, c2 } = parseA1Range(a1);
      const endCol = colLetter(c2);
      const colWidths = new Map<number, number>();
      // Row/col indices here are zero-based within the used range, which is
      // what kpi.ts expects — it never relies on absolute sheet coordinates.
      const grid: SheetData = { values: [], numberFormat: [] };

      for (let rowStart = r1; rowStart <= r2; rowStart += CHUNK_ROWS) {
        const rowEnd = Math.min(rowStart + CHUNK_ROWS - 1, r2);
        const addr = `${colLetter(c1)}${rowStart}:${endCol}${rowEnd}`;
        const data = await graphJson<RangeData>(
          wsUrl(
            driveId,
            itemId,
            sheetName,
            `/range(address='${addr}')?$select=values,numberFormat`
          ),
          {},
          sh
        );
        for (let i = 0; i < data.values.length; i++) {
          grid.values.push(data.values[i] ?? []);
          grid.numberFormat.push(data.numberFormat?.[i] ?? []);
          const excelRow = rowStart + i;
          for (let j = 0; j < data.values[i].length; j++) {
            const v = data.values[i][j];
            if (v === "" || v === null || v === undefined) continue;
            const cell = ws.getCell(excelRow, c1 + j);
            cell.value = v as ExcelJS.CellValue;
            const fmt = data.numberFormat?.[i]?.[j];
            if (fmt && fmt !== "General") cell.numFmt = fmt;
            const isDateFmt = !!fmt && /[dmy]/i.test(fmt) && fmt !== "General";
            const rendered = typeof v === "number" && isDateFmt ? 12 : String(v).length;
            const col = c1 + j;
            if ((colWidths.get(col) ?? 0) < rendered) colWidths.set(col, rendered);
          }
        }
      }

      for (const [col, w] of colWidths) {
        ws.getColumn(col).width = Math.min(Math.max(w + 2, 10), 45);
      }
      ws.spliceRows(1, 0, [SHEET_BANNER(sheetName)]);
      collected.set(sheetName, grid);
    }

    // Derived tab, appended last so it never shifts the mirrored tabs.
    if (phoneRows.length > 0) {
      addPhoneSheet(out, phoneRows);
    }

    const buffer = new Uint8Array(await out.xlsx.writeBuffer());
    if (!dryRun) {
      await uploadFile(driveId, folderPath, recapName, buffer);
    }

    const report: FileReport = {
      master: masterName,
      recap: recapName,
      keptTabs: phoneRows.length > 0 ? [...plan.keep, PHONE_SHEET_NAME] : plan.keep,
      droppedWeeklyTabs: plan.dropped,
      recapBytes: buffer.byteLength,
      phoneRows: phoneRows.length,
      ms: Date.now() - t0,
    };
    return { report, sheets: collected };
  } finally {
    // Best-effort session close; never let cleanup mask a real result.
    graphFetch(`${base}/closeSession`, { method: "POST", body: "{}" }, sh).catch(() => {});
  }
}