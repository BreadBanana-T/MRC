# Deploying MRC Ops from this repo (no more file-by-file drift)

The whole app lives in this repo. Push it all at once with clasp so the Apps
Script editor is *exactly* `main` — never a mix of old/new files.

## One-time setup
1. Install: `npm install -g @google/clasp`
2. `clasp login`
3. Get your **Script ID**: Apps Script editor → ⚙ Project Settings → "IDs".
4. Copy `.clasp.json.example` to `.clasp.json` and paste your Script ID in.

## Every update
```
git pull            # get latest main
clasp push          # overwrites the Apps Script project with this repo, exactly
```
Then in the editor: **Deploy → Manage deployments → ✏ → New version** (or use the
Test deployment `/dev` URL to verify first).

## MRC Workforce (break planner)
Same deployment, second page: open `<web app URL>?app=workforce` (or the calendar-clock
button in the MRC header). Nothing extra to deploy or enable.
Run `setupDemandTrigger()` once from the editor: it logs outages / weather / Alarm.com next to our
alarm trend every hour, which is how the planner learns which conditions really raise our volume.

## Services to enable once (Editor → Services → +)
- People API   (identifier `People`, v1)   — profile photos
- Google Sheets API (identifier `Sheets`, v4) — fast batch reads

clasp pushes `.gs`, `.html`, and `appsscript.json`; the `.xlsx`/`.md`/`.txt`
files are ignored automatically.
