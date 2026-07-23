/**
 * /api/cron — regenerates slim *_RECAP.xlsx copies of every recruiter's
 * master pipeline workbook so the Claude M365 connector (whose renderer
 * caps how much of a workbook it will extract) can always read them.
 *
 * Each recap now also carries a derived KPI_TABLE sheet: one row per week
 * with all 8 metrics denormalized onto that row, so every metric — not
 * just First/Second Interviews — is unambiguous after the connector
 * flattens the file to text.
 *
 * Invocation:
 *   - Vercel Cron (schedule in vercel.json). Vercel automatically sends
 *     "Authorization: Bearer <CRON_SECRET>" when CRON_SECRET is set.
 *   - Manual:  curl -H "Authorization: Bearer $CRON_SECRET" \
 *                https://<deployment>/api/cron
 *   - Options: ?dryRun=1        build but don't upload (test mode)
 *              ?file=<name>     sync only one master workbook
 *              ?verbose=1       include full KPI extraction diagnostics
 *                               (which metric columns were located, etc.)
 *
 * First-run tip: run once with ?dryRun=1&verbose=1 and read each report's
 * kpiDiagnostics before trusting the schedule — it shows exactly which
 * metric headers the KPI parser found and how many week rows it produced.
 */

import type { VercelRequest, VercelResponse } from "@vercel/node";
import { listFolder, optionalEnv, resolveDriveId } from "../lib/graph";
import { splitWorkbook, type FileReport } from "../lib/split";

export default async function handler(req: VercelRequest, res: VercelResponse) {
  const secret = process.env.CRON_SECRET;
  if (!secret) {
    return res.status(500).json({ error: "CRON_SECRET is not configured" });
  }
  if (req.headers.authorization !== `Bearer ${secret}`) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  const dryRun = req.query.dryRun === "1" || req.query.dryRun === "true";
  const verbose = req.query.verbose === "1" || req.query.verbose === "true";
  const onlyFile = typeof req.query.file === "string" ? req.query.file : null;

  const folderPath = optionalEnv("PIPELINE_FOLDER_PATH", "eggers-pipeline");
  const masterPattern = new RegExp(optionalEnv("MASTER_PATTERN", "_Pipeline_.*\\.xlsx$"), "i");
  const recapSuffix = optionalEnv("RECAP_SUFFIX", "_RECAP");
  const weeksToKeep = Number(optionalEnv("WEEKS_TO_KEEP", "10"));

  const started = new Date().toISOString();
  const reports: FileReport[] = [];

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

    if (masters.length === 0) {
      return res.status(200).json({
        started,
        dryRun,
        folderPath,
        message: onlyFile
          ? `No file named "${onlyFile}" matching MASTER_PATTERN in folder`
          : "No master workbooks matched MASTER_PATTERN in folder",
        filesInFolder: items.map((i) => i.name),
      });
    }

    for (const master of masters) {
      try {
        reports.push(
          await splitWorkbook({
            driveId,
            folderPath,
            itemId: master.id,
            masterName: master.name,
            recapSuffix,
            weeksToKeep,
            dryRun,
          })
        );
      } catch (err) {
        reports.push({
          master: master.name,
          recap: "",
          keptTabs: [],
          droppedWeeklyTabs: [],
          recapBytes: 0,
          ms: 0,
          error: err instanceof Error ? err.message : String(err),
        });
      }
    }

    // Unless ?verbose=1, keep the response light: drop the (large) per-file
    // KPI diagnostics but always keep the row count + source tab so the
    // schedule's health is visible at a glance.
    const shaped = reports.map((r) => {
      if (verbose) return r;
      const { kpiDiagnostics, ...rest } = r;
      return rest;
    });

    const failed = reports.filter((r) => r.error);
    const kpiEmpty = reports.filter((r) => !r.error && !r.kpiRowCount).map((r) => r.master);

    return res.status(failed.length === reports.length && reports.length > 0 ? 500 : 200).json({
      started,
      finished: new Date().toISOString(),
      dryRun,
      verbose,
      folderPath,
      synced: reports.length - failed.length,
      failed: failed.length,
      /** Masters whose KPI_TABLE came back empty — worth a look with ?verbose=1. */
      kpiEmpty,
      reports: shaped,
    });
  } catch (err) {
    return res.status(500).json({
      started,
      error: err instanceof Error ? err.message : String(err),
      reports,
    });
  }
}