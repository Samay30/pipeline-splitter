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
 */

import ExcelJS from "exceljs";
import { graphFetch, graphJson, uploadFile } from "./graph";
import { planKeep } from "./tabs";

const CHUNK_ROWS = 500;

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
  error?: string;
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
    const plan = planKeep(ordered, weeksToKeep);

    const out = new ExcelJS.Workbook();
    out.creator = "eggers-internal pipeline-sync";
    out.created = new Date();

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
          const excelRow = rowStart + i;
          for (let j = 0; j < data.values[i].length; j++) {
            const v = data.values[i][j];
            if (v === "" || v === null || v === undefined) continue;
            const cell = ws.getCell(excelRow, c1 + j);
            cell.value = v as ExcelJS.CellValue;
            const fmt = data.numberFormat?.[i]?.[j];
            if (fmt && fmt !== "General") cell.numFmt = fmt;
          }
        }
      }
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
    };
  } finally {
    // Best-effort session close; never let cleanup mask a real result.
    graphFetch(`${base}/closeSession`, { method: "POST", body: "{}" }, sh).catch(() => {});
  }
}
