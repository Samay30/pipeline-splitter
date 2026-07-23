/**
 * Reads a master pipeline workbook through the Graph *workbook API*
 * (per-sheet JSON — no renderer size cap, unlike the text-extraction
 * renderer the Claude connector uses), then writes a slim values-only
 * copy ("recap file") back to the same SharePoint folder.
 *
 * The recap file mirrors the master's most recent weekly tabs plus all
 * non-date tabs, preserving cell values and number formats (so dates
 * stay real dates). Formulas are flattened to their computed values.
 *
 * NEW: the recap also gets a derived KPI_TABLE sheet — one row per week
 * with all 8 metrics denormalized onto that single row. The embedded KPI
 * grid inside a weekly tab is several side-by-side / stacked mini-tables;
 * when the connector flattens the file to text, which number belongs to
 * which week/metric becomes ambiguous for everything except First/Second
 * Interviews. KPI_TABLE removes that ambiguity: because every metric for a
 * week sits on ONE spreadsheet row, flattening keeps them together. Built
 * from the newest kept weekly tab, whose running grid already spans the
 * whole year, so we parse only one tab.
 */

import ExcelJS from "exceljs";
import { graphFetch, graphJson, uploadFile } from "./graph";
import {
  inferYear,
  planKeep,
  newestWeeklyTab,
  buildKpiTable,
  kpiRowToArray,
  KPI_TABLE_HEADERS,
  type KpiDiagnostics,
} from "./tabs";

const CHUNK_ROWS = 500;

/** Injected as each sheet's first row so flattening can't lose tab identity. */
const SHEET_BANNER = (name: string) => `=== SHEET: ${name} ===`;

/** Name of the derived KPI sheet. */
const KPI_SHEET_NAME = "KPI_TABLE";

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
  /** NEW: how many week rows the KPI_TABLE sheet ended up with (0 = not built). */
  kpiRowCount?: number;
  /** NEW: which weekly tab the KPI grid was parsed from. */
  kpiSourceTab?: string;
  /** NEW: full extraction diagnostics — read this on the first dry run. */
  kpiDiagnostics?: KpiDiagnostics;
  error?: string;
}

function wsUrl(driveId: string, itemId: string, sheetName: string, tail: string): string {
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

/**
 * Fetch a sheet's entire used range as dense grid-local 2D arrays
 * (values + numberFormat), indices starting at [0][0] = used-range
 * top-left. Used to feed buildKpiTable. Returns null for an empty sheet.
 */
async function fetchSheetGrid(
  driveId: string,
  itemId: string,
  sheetName: string,
  sh: Record<string, string>
): Promise<{ values: unknown[][]; numberFormat: string[][] } | null> {
  let meta: RangeMeta;
  try {
    meta = await graphJson<RangeMeta>(
      wsUrl(driveId, itemId, sheetName, `/usedRange?$select=address,rowCount,columnCount`),
      {},
      sh
    );
  } catch {
    return null;
  }
  const a1 = meta.address.includes("!") ? meta.address.slice(meta.address.lastIndexOf("!") + 1) : meta.address;
  const { r1, c1, r2, c2 } = parseA1Range(a1);
  const endCol = colLetter(c2);
  const width = c2 - c1 + 1;

  const values: unknown[][] = [];
  const numberFormat: string[][] = [];
  for (let rowStart = r1; rowStart <= r2; rowStart += CHUNK_ROWS) {
    const rowEnd = Math.min(rowStart + CHUNK_ROWS - 1, r2);
    const addr = `${colLetter(c1)}${rowStart}:${endCol}${rowEnd}`;
    const data = await graphJson<RangeData>(
      wsUrl(driveId, itemId, sheetName, `/range(address='${addr}')?$select=values,numberFormat`),
      {},
      sh
    );
    for (let i = 0; i < data.values.length; i++) {
      const vRow = data.values[i] ?? [];
      const fRow = data.numberFormat?.[i] ?? [];
      const outV: unknown[] = new Array(width).fill("");
      const outF: string[] = new Array(width).fill("General");
      for (let j = 0; j < width; j++) {
        outV[j] = vRow[j] ?? "";
        outF[j] = fRow[j] ?? "General";
      }
      values.push(outV);
      numberFormat.push(outF);
    }
  }
  return { values, numberFormat };
}

export async function splitWorkbook(opts: {
  driveId: string;
  folderPath: string;
  itemId: string;
  masterName: string;
  recapSuffix: string;
  weeksToKeep: number;
  dryRun: boolean;
}): Promise<FileReport> {
  const t0 = Date.now();
  const { driveId, folderPath, itemId, masterName, recapSuffix, weeksToKeep, dryRun } = opts;
  const recapName = masterName.replace(/\.xlsx$/i, "") + recapSuffix + ".xlsx";
  const base = `/drives/${driveId}/items/${itemId}/workbook`;

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
    const defaultYear = inferYear(masterName, ordered);
    const plan = planKeep(ordered, weeksToKeep, defaultYear);

    const out = new ExcelJS.Workbook();
    out.creator = "eggers-internal pipeline-sync";
    out.created = new Date();

    for (const sheetName of plan.keep) {
      const ws = out.addWorksheet(sheetName);

      let meta: RangeMeta;
      try {
        meta = await graphJson<RangeMeta>(
          wsUrl(driveId, itemId, sheetName, `/usedRange?$select=address,rowCount,columnCount`),
          {},
          sh
        );
      } catch {
        continue;
      }
      const a1 = meta.address.includes("!") ? meta.address.slice(meta.address.lastIndexOf("!") + 1) : meta.address;
      const { r1, c1, r2, c2 } = parseA1Range(a1);
      const endCol = colLetter(c2);
      const colWidths = new Map<number, number>();

      for (let rowStart = r1; rowStart <= r2; rowStart += CHUNK_ROWS) {
        const rowEnd = Math.min(rowStart + CHUNK_ROWS - 1, r2);
        const addr = `${colLetter(c1)}${rowStart}:${endCol}${rowEnd}`;
        const data = await graphJson<RangeData>(
          wsUrl(driveId, itemId, sheetName, `/range(address='${addr}')?$select=values,numberFormat`),
          {},
          sh
        );
        for (let i = 0; i < data.values.length; i++) {
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
    }

    // ---- Derived KPI_TABLE sheet ---------------------------------------
    let kpiRowCount = 0;
    let kpiSourceTab: string | undefined;
    let kpiDiagnostics: KpiDiagnostics | undefined;
    try {
      const source = newestWeeklyTab(plan.keep, defaultYear);
      if (source) {
        kpiSourceTab = source;
        const grid = await fetchSheetGrid(driveId, itemId, source, sh);
        if (grid) {
          const { rows, diagnostics } = buildKpiTable(grid.values, grid.numberFormat);
          kpiDiagnostics = diagnostics;
          kpiRowCount = rows.length;

          const ws = out.addWorksheet(KPI_SHEET_NAME);
          ws.addRow([SHEET_BANNER(KPI_SHEET_NAME)]);
          ws.addRow(KPI_TABLE_HEADERS);
          for (const r of rows) {
            const arr = kpiRowToArray(r);
            const row = ws.addRow(arr);
            // Render the Week cell as a real date.
            const wk = row.getCell(1);
            if (typeof arr[0] === "string" && /^\d{4}-\d{2}-\d{2}$/.test(arr[0])) {
              wk.value = new Date(arr[0] + "T00:00:00Z");
              wk.numFmt = "yyyy-mm-dd";
            }
          }
          ws.getColumn(1).width = 12;
          for (let c = 2; c <= KPI_TABLE_HEADERS.length; c++) ws.getColumn(c).width = 20;
        }
      }
    } catch (kpiErr) {
      // KPI sheet is additive — never let it fail the whole recap.
      kpiDiagnostics = {
        weekHeader: null,
        found: {},
        matchCounts: {},
        weekSlots: 0,
        weekRows: 0,
        notes: [`KPI build error: ${kpiErr instanceof Error ? kpiErr.message : String(kpiErr)}`],
      };
    }

    const buffer = new Uint8Array(await out.xlsx.writeBuffer());
    if (!dryRun) {
      await uploadFile(driveId, folderPath, recapName, buffer);
    }

    return {
      master: masterName,
      recap: recapName,
      keptTabs: plan.keep,
      droppedWeeklyTabs: plan.dropped,
      recapBytes: buffer.byteLength,
      ms: Date.now() - t0,
      kpiRowCount,
      kpiSourceTab,
      kpiDiagnostics,
    };
  } finally {
    graphFetch(`${base}/closeSession`, { method: "POST", body: "{}" }, sh).catch(() => {});
  }
}