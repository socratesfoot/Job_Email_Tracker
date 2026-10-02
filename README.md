# Job Email Tracker

A Google Apps Script that scans Gmail for job-related messages, logs them to a
"Jobs" sheet, scores them against your targets, ages entries
(NEW → OLD → OUTDATED), and tracks interview/rejection replies so nothing
slips through.

## Usage

1. Open a Google Sheet, then **Extensions > Apps Script**.
2. Paste in `code.gs` (replacing the default `Code.gs`).
3. (Optional) If you want distance filtering for hybrid/onsite jobs, enable
   the Maps service: in the Apps Script editor, **Services (+) > Maps**.
   Without it, distance checks are skipped and a warning is logged each run.
4. Run `setup()` once. This grants permissions, creates the **Jobs** and
   **Settings** sheets (with headers and color-coded status formatting), and
   installs a trigger that scans every 15 minutes.
5. Open the **Settings** tab and set your filters (all off by default):
   require contact email / location / phone, distance cap for hybrid/onsite
   jobs (ZIP + miles), blacklist of phrases.
6. Run `backfill()` once to scan the last 90 days of mail.
7. Reopening the spreadsheet adds a **Job Tracker** menu: open settings, scan
   now, backfill last 90 days.

Day to day it runs itself. New postings appear as NEW rows; rows age to OLD
after 7 days of inactivity and OUTDATED after 90 days. Rows in an
INTERVIEW REQUEST stage are never auto-aged, and any new message on one keeps
it IMPORTANT so requested replies don't get missed.

## How postings are scored (0–100)

Every posting starts at 40 and gains or loses points against your targets:

- Full-time employment: +30. Contract: +15 if 12+ months stated.
- Remote: +20. Hybrid: +5.
- Compensation at/above $200k/yr (hourly annualized at 2,080 hrs): +25;
  $150k+: +10.
- Each specialty hit (key management, information security, AI, TMS,
  project management): +5, up to 5 hits.
- Score 75+ → Priority column marked IMPORTANT.

Filtered out as non-serious: educational/training pitches, part-time roles,
and contracts under 12 months when a duration is stated. Unknown-duration
contracts are kept but score lower. Duplicate postings for the same company +
position are merged into one row.

## Sheet columns

Status, Priority, Stage, Date Received, Company, Position, Type, Work Mode,
Location, Contact Name, Contact Email, Contact Phone, Subject, Thread ID,
Message IDs, Last Activity, Salary, Gmail Link, Score.

- **Salary**: compensation pulled from the email ($70/hr, $120k–$150k, ...).
- **Gmail Link**: direct link to the email thread.
- **Score**: 0–100 fit score (see above).

## Script structure (`code.gs`)

- **Config (top of file)** — sheet name, column headers, aging thresholds
  (7 / 90 days), active stages, and tuning knobs: `TARGET_ANNUAL_SALARY`,
  `MIN_CONTRACT_MONTHS`, `BASE_SCORE`, `IMPORTANT_AT`, `SPECIALTIES`,
  `NOT_SERIOUS` patterns.
- **Settings admin console** — `getSettingsSheet_()` / `getSettings_()` build
  and read the Settings tab (require email/location/phone, distance cap,
  blacklist). Filters apply only to newly found postings, never to rows
  already captured.
- **Entry points** — `setup()` (one-time install), `backfill()` (90-day
  scan), `run()` (recent-mail scan, on the 15-min trigger), `onOpen()`
  (Job Tracker menu).
- **Mailbox scan** — `processMailbox_()` searches Gmail, `handleMessage_()`
  classifies each message, `buildRow_()` extracts fields, `applySignal_()`
  updates stage/priority on replies, `touch_()` refreshes activity.
- **Parsing helpers** — `extractSalary_` / `salaryAnnualMax_`
  (compensation), `extractContractMonths_` (contract length), `isPosting_`
  (posting detection), `parseSender_` / `companyFromDomain_` /
  `cleanSubject_`, `geocodeCached_` / `haversineMiles_` (distance via the
  Maps service).
- **Dedup & scoring** — `findDuplicate_()` merges re-sent postings for the
  same company + position; `scorePosting_()` computes the 0–100 score.
- **Sheet I/O** — `ensureHeaders_()`, `readRows_()`, `writeDirtyRows_()`
  (only changed rows are written back each run), `refreshStatuses_()`
  (aging NEW → OLD → OUTDATED).

## Notes

- Tune the targets at the top of `code.gs` to taste.
- Distance filtering warns in the log when enabled without the Maps service.
