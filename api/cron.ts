/**
 * /api/cron — regenerates slim *_RECAP.xlsx copies of every recruiter's
 * master pipeline workbook so the Claude M365 connector (whose renderer
 * caps how much of a workbook it will extract) can always read them.
 *
 * Invocation:
 *   - Vercel Cron (schedule in vercel.json). Vercel automatically sends
 *     "Authorization: Bearer <CRON_SECRET>" when CRON_SECRET is set.
 *   - Manual:  curl -H "Authorization: Bearer $CRON_SECRET" \
 *                https://<deployment>/api/cron
 *   - Options: ?dryRun=1        build but don't upload (test mode)
 *              ?file=<name>     sync only one master workbook
 *              ?noPhone=1       skip the Ringover pull entirely
 *
 * Each recap also carries a PHONE tab with the firm's weekly Ringover
 * roll-up. That pull happens once per run for everyone, before the loop, and
 * is strictly best-effort: if it fails the recaps are still written without
 * the tab and the failure is reported. A phone outage must never cost us the
 * recap files, which are the reason this job exists.
 */

import type { VercelRequest, VercelResponse } from "@vercel/node";
import { listFolder, optionalEnv, resolveDriveId } from "../lib/graph";
import { splitWorkbook, type FileReport } from "../lib/split";
import { collectPhoneWeeks, type PhoneResult } from "../lib/ringover";
import { buildKpiPayload, readBilling, recruiterFromFilename, type SheetData } from "../lib/kpi";
import { graphJson, uploadFile } from "../lib/graph";

export default async function handler(req: VercelRequest, res: VercelResponse) {
  const secret = process.env.CRON_SECRET;
  if (!secret) {
    return res.status(500).json({ error: "CRON_SECRET is not configured" });
  }
  if (req.headers.authorization !== `Bearer ${secret}`) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  const dryRun = req.query.dryRun === "1" || req.query.dryRun === "true";
  const onlyFile = typeof req.query.file === "string" ? req.query.file : null;
  const noPhone = req.query.noPhone === "1" || req.query.noPhone === "true";

  const folderPath = optionalEnv("PIPELINE_FOLDER_PATH", "eggers-pipeline");
  const masterPattern = new RegExp(optionalEnv("MASTER_PATTERN", "_Pipeline_.*\\.xlsx$"), "i");
  const recapSuffix = optionalEnv("RECAP_SUFFIX", "_RECAP");
  const kpiSuffix = optionalEnv("KPI_SUFFIX", "_KPI");
  const billingsPattern = new RegExp(
    optionalEnv("BILLINGS_PATTERN", "Master.*Billing.*\\.xlsx$"),
    "i"
  );
  const noKpi = req.query.noKpi === "1" || req.query.noKpi === "true";
  const weeksToKeep = Number(optionalEnv("WEEKS_TO_KEEP", "10"));

  const started = new Date().toISOString();
  const reports: FileReport[] = [];
  let phone: PhoneResult | null = null;
  let billingSheet: SheetData | undefined;
  let billingsName: string | null = null;
  const kpiNotes: string[] = [];
  const kpiFiles: Array<{ recruiter: string; file: string; week: string; error?: string }> = [];

  try {
    const driveId = await resolveDriveId();
    const items = await listFolder(driveId, folderPath);

    const masters = items.filter(
      (i) =>
        i.file &&
        masterPattern.test(i.name) &&
        !i.name.toLowerCase().includes(recapSuffix.toLowerCase()) &&
        (!onlyFile || i.name === onlyFile)
    );

    // One pull for the whole firm, before the loop. collectPhoneWeeks never
    // throws — a problem comes back on .error and we carry on without the tab.
    if (!noPhone) {
      phone = await collectPhoneWeeks();
    }

    // One read of the shared billings workbook, also before the loop: one file
    // serves every recruiter in the run.
    if (!noKpi) {
      const billingsFile = items.find((i) => i.file && billingsPattern.test(i.name));
      if (billingsFile) {
        try {
          billingSheet = await readBillingSheet(driveId, billingsFile.id);
          billingsName = billingsFile.name;
        } catch (err) {
          kpiNotes.push(
            `Billings workbook found but not readable: ` +
              `${err instanceof Error ? err.message : String(err)}`
          );
        }
      } else {
        kpiNotes.push("No billings workbook matched BILLINGS_PATTERN — billing card will be blank.");
      }
    }

    if (masters.length === 0) {
      return res.status(200).json({
        started,
        dryRun,
        folderPath,
        message: onlyFile
          ? `No file named "${onlyFile}" matching MASTER_PATTERN in folder`
          : "No master workbooks matched MASTER_PATTERN in folder",
        filesInFolder: items.map((i) => i.name),
        phone: phoneSummary(phone),
      });
    }

    // Sequential on purpose: parallel workbook sessions on one drive
    // invite Graph throttling, and the cron has a generous time budget.
    for (const master of masters) {
      try {
        const { report, sheets } = await splitWorkbook({
          driveId,
          folderPath,
          itemId: master.id,
          masterName: master.name,
          recapSuffix,
          weeksToKeep,
          dryRun,
          phoneRows: phone?.rows ?? [],
        });
        reports.push(report);

        // Derived KPI payload. Never allowed to fail the recap that just
        // succeeded — the recap file is the reason this job exists.
        if (!noKpi) {
          const recruiter = recruiterFromFilename(master.name);
          try {
            const payload = buildKpiPayload({
              masterName: master.name,
              sheets,
              billingSheet,
              phoneRows: phone?.rows ?? [],
            });
            const kpiName = master.name.replace(/\.xlsx$/i, "") + kpiSuffix + ".json";
            if (!dryRun) {
              await uploadFile(
                driveId,
                folderPath,
                kpiName,
                new TextEncoder().encode(JSON.stringify(payload, null, 2)),
                "application/json"
              );
            }
            kpiFiles.push({
              recruiter,
              file: kpiName,
              week: String(payload.week_of ?? "unknown"),
            });
          } catch (err) {
            kpiFiles.push({
              recruiter,
              file: "",
              week: "",
              error: err instanceof Error ? err.message : String(err),
            });
          }
        }
      } catch (err) {
        reports.push({
          master: master.name,
          recap: "",
          keptTabs: [],
          droppedWeeklyTabs: [],
          recapBytes: 0,
          phoneRows: 0,
          ms: 0,
          error: err instanceof Error ? err.message : String(err),
        });
      }
    }

    const failed = reports.filter((r) => r.error);
    return res.status(failed.length === reports.length && reports.length > 0 ? 500 : 200).json({
      started,
      finished: new Date().toISOString(),
      dryRun,
      folderPath,
      synced: reports.length - failed.length,
      failed: failed.length,
      phone: phoneSummary(phone),
      kpi: {
        status: noKpi ? "disabled" : "ok",
        billingsWorkbook: billingsName,
        written: kpiFiles,
        notes: kpiNotes,
      },
      reports,
    });
  } catch (err) {
    return res.status(500).json({
      started,
      error: err instanceof Error ? err.message : String(err),
      phone: phoneSummary(phone),
      reports,
    });
  }
}

/** Read the whole billings workbook's year-goals sheets through the workbook API. */
async function readBillingSheet(driveId: string, itemId: string): Promise<SheetData> {
  const base = `/drives/${driveId}/items/${itemId}/workbook`;
  const sheets = await graphJson<{ value: Array<{ name: string }> }>(
    `${base}/worksheets?$select=name`
  );
  // Prefer the "<year> Goals" tab for the current year, else the last Goals tab.
  const year = String(new Date().getUTCFullYear());
  const goals = sheets.value.map((s) => s.name).filter((n) => /goal/i.test(n));
  const target = goals.find((n) => n.includes(year)) ?? goals[goals.length - 1] ?? sheets.value[0]?.name;
  if (!target) throw new Error("billings workbook has no sheets");

  const escaped = encodeURIComponent(target.replace(/'/g, "''"));
  const data = await graphJson<{ values: unknown[][]; numberFormat: string[][] }>(
    `${base}/worksheets('${escaped}')/usedRange?$select=values,numberFormat`
  );
  return { values: data.values ?? [], numberFormat: data.numberFormat ?? [] };
}

/** Compact phone status for the run report — read this before trusting the tab. */
function phoneSummary(phone: PhoneResult | null) {
  if (!phone) return { status: "disabled" as const };
  if (phone.error) return { status: "failed" as const, error: phone.error, ms: phone.ms };
  if (phone.skipped) return { status: "skipped" as const, reason: phone.skipped };
  return {
    status: "ok" as const,
    weeks: phone.weeks,
    recruiters: phone.rosterSize,
    callsPulled: phone.callsPulled,
    rows: phone.rows.length,
    unmappedUsers: phone.unmappedUsers,
    ms: phone.ms,
  };
}
