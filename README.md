# pipeline-sync

Scheduled job that keeps a slim, connector-readable copy of every recruiter's
pipeline workbook on SharePoint.

## Why this exists

The Claude Microsoft 365 connector reads Excel files through Microsoft's
text-extraction renderer, which caps how much content it renders. A full-year
pipeline workbook (30+ tabs) exceeds the cap and the connector gets a 406.
The Graph **workbook API** has no such cap — it reads per-sheet as JSON. This
job uses it to read each master, then writes a small values-only copy
(`<master>_RECAP.xlsx`) containing all non-date tabs (Active Positions,
Clients, Candidates, etc.) plus the `WEEKS_TO_KEEP` most recent weekly tabs.
The eggers-pipeline-recap skill reads the `_RECAP` file; the master is never
touched, and recruiters change nothing about how they work.

```
cron (weekday 6am/12pm/4pm CT)
  └─ Graph workbook API reads master (any size)
       └─ writes  Aaron_Rider_Pipeline_2026_RECAP.xlsx  next to master
            └─ recruiter: "recap pipeline this week"
                 └─ connector reads the small _RECAP file → skill → rundown
```

## One-time setup

### 1. Entra ID permissions
The existing invoicing app registration can be reused **if** its Graph
application permissions cover the site that hosts the pipeline folder.

- Simplest: `Sites.ReadWrite.All` (Application) with admin consent (Jason).
- Least-privilege: `Sites.Selected` (Application), then grant `write` on the
  Eggers_GTM_Collission site specifically via a Graph call — ask if you want
  the one-liner for this.

If the invoicing app was consented for `Sites.ReadWrite.All` already, no
Entra change is needed at all.

### 2. Environment variables
Copy `.env.example` → Vercel project env. Notes:

- `PIPELINE_FOLDER_PATH`: for a Teams channel, files live under the channel
  folder — usually `General/eggers-pipeline` or `<ChannelName>/eggers-pipeline`.
  Open the folder in SharePoint and read the path from the URL to confirm.
- `DRIVE_ID` skips two Graph lookups per run if you have it (same format as
  the invoicing SharePoint drive IDs, `b!...`).
- `CRON_SECRET`: `openssl rand -hex 32`. Vercel attaches it to cron calls
  automatically as `Authorization: Bearer <secret>`.

### 3. Deploy (monorepo)
```bash
# from eggers-internal root
cd packages/pipeline-sync
npm install
npx tsc --noEmit          # typecheck
vercel link               # new Vercel project, root dir = packages/pipeline-sync
vercel env pull           # or add env vars in dashboard first
vercel deploy --prod
```

### 4. First run — dry, then real
```bash
# builds every recap in memory, uploads nothing, returns a full report
curl -H "Authorization: Bearer $CRON_SECRET" \
  "https://<deployment>/api/cron?dryRun=1"

# single-file real run against the test workbook
curl -H "Authorization: Bearer $CRON_SECRET" \
  "https://<deployment>/api/cron?file=Aaron_Rider_Pipeline_2026_TEST_9wk.xlsx"

# full real run
curl -H "Authorization: Bearer $CRON_SECRET" "https://<deployment>/api/cron"
```
The JSON report lists, per master: kept tabs, dropped weekly tabs, recap file
size, and duration — read it before trusting the schedule.

### 5. Update the skill (one paragraph)
Add this to `eggers-pipeline-recap/SKILL.md` under "What this skill actually
does", so it targets the derived file:

> **Which file to read:** pipeline workbooks exist in two forms on SharePoint —
> the recruiter's master (e.g. `Aaron_Rider_Pipeline_2026.xlsx`) and an
> auto-generated slim copy ending in `_RECAP.xlsx` that contains the same
> tabs minus old weeks. Always read the `_RECAP` file; the master is too
> large for the file renderer and will fail with a 406. The `_RECAP` copy is
> regenerated several times each weekday, values-only, and is never edited
> by hand.

## Schedule

`vercel.json` runs weekdays at 11:00, 17:00, 21:00 UTC = 6am, 12pm, 4pm
Central. Numbers a recruiter types in are reflected in their next recap after
the following run; adjust the cron expression if that staleness window is too
wide (or hit the endpoint manually for an instant refresh).

## Design decisions

- **Values-only copy.** Formulas are flattened to computed values. Dropped
  sheets would break cross-tab references anyway, and each weekly tab's
  embedded KPI table (weekly rows, quarter totals, YTD) is self-contained,
  so nothing the skill reads is lost. Number formats are preserved so date
  columns remain real dates.
- **Keep rule = all non-date tabs + last N weekly tabs.** No allow-list to
  maintain; a recruiter adding an unexpected tab can't break the job.
- **Sequential processing, read-only workbook sessions, retry with
  Retry-After.** Plays nice with Graph throttling.
- **Masters are opened read-only and never written.** The only write is the
  `_RECAP` upload (create-or-replace).

## Troubleshooting

- `401 Unauthorized` from the endpoint → CRON_SECRET mismatch.
- `403` from Graph → app registration lacks permission on this site
  (see step 1) or consent not granted.
- `Drive "Documents" not found` → set `DRIVE_NAME` to one of the names the
  error lists, or set `DRIVE_ID` directly.
- Empty `filesInFolder` in the report → `PIPELINE_FOLDER_PATH` is wrong;
  remember Teams channels nest under the channel name.
- Recap still 406s in the connector → lower `WEEKS_TO_KEEP`.
