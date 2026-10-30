# MRC Operations Portal — "What's new" video

Prompt for Claude Motion, plus the screenshots it uses. Everything in this folder
is images and Markdown only (`clasp push` ignores it).

- `side-by-side/` — 1920×1080 boards, Before on the left, After on the right, page title on top.
- `before/` and `after/` — the raw screenshots (1500×1000) if you want to animate them yourself.

"Before" is the portal as of 7 Sep 2026 (`main` at `73047c3`), before this overhaul.
"After" is today's `main`. Both were rendered from the same mock data, so differences are
the product, not the data.

---

## Prompt (copy everything below this line into Claude Motion)

Create a 100–120 second product update video for an internal operations tool called the
**MRC Operations Portal**, used by a workforce/real-time team (WFM schedules, a live agent
floor on Genesys Cloud, power-outage and alarm monitoring, management reporting).
Audience: operations leads and managers. Tone: calm, confident, professional — a release
walkthrough, not an ad. No hype words, no emoji, no stock footage, no people.

**Visual style**
- 16:9, 1920×1080, 30 fps.
- Background: soft off-white `#EEF2FB`; ink text `#0F172A`; secondary text `#64748B`;
  single accent indigo `#4F46E5`. Good = `#0CA30C`, bad = `#D03B3B`, used only on numbers
  that are good or bad.
- Type: Inter (headings Bold, body Medium). Titles 56–64 px, captions 28–32 px.
- Motion: smooth and restrained — eased slides and fades (300–600 ms), gentle push-ins on
  screenshots, wipes between Before and After. Nothing bouncing or spinning; nothing loops.
- Screenshots sit on the background with a thin 1 px border `#CBD5E1` and a soft shadow.
  When showing a Before/After pair, either use the provided side-by-side board or a
  horizontal wipe from Before to After on the same frame.

**Assets** (all in `docs/whats-new/`):
`side-by-side/01_dashboard.png` … `18_notifications.png`, and the same names in
`before/` and `after/`.

**Storyboard** (times are targets; on-screen text in quotes, voice-over in italics)

1. **Title — 0:00–0:06**
   "MRC Operations Portal" / "What's new" — indigo underline draws in under the title.
   *"Here's what changed in the MRC Operations Portal, and why it matters."*

2. **The problem we fixed first — 0:06–0:20**
   Big number card animating from left to right:
   "Dashboard at rest: 106% of a CPU core → 0.2%"
   second line: "Genesys page: 75% → 0.2%"
   small caption: "measured with the page idle; Chrome had been flagging the tab as slowing the browser"
   *"Chrome kept flagging the portal for slowing the computer. The cause was animations that
   never stopped — on PCs without graphics acceleration that redraws the whole screen sixty
   times a second. Every animation now plays once, when something changes, then rests."*

3. **Dashboard — 0:20–0:34** — `side-by-side/01_dashboard.png`, then push into
   `after/01_dashboard.png` top row, then the "Right now" card.
   Callouts (appear one by one, pointing at the After): "Live numbers that matter now",
   "'Right now' in plain words", "One chart per question".
   *"The dashboard now leads with the live numbers — service level, response time, who's on
   the floor, unplanned absences, customers without power — and a 'Right now' card that says
   what's happening in plain words."*

4. **Management view — 0:34–0:50** — `side-by-side/03_management.png`, then push into the
   After's headline row and the "What stands out" card.
   Callouts: "Headline numbers with trend", "Compared like for like: this week to date vs
   last week to date", "Findings, not just figures".
   *"The management view opens on four headline numbers with their trend, compared fairly
   against the same days last period, followed by what stands out — the days under target,
   the biggest staffing gap, where absence lands."*

5. **Trackers — 0:50–1:05** — quick sequence, ~3 s each, Before→After wipe:
   `05_furlough`, `08_safe`, `09_safe_coverage`, `10_absences`, `11_training`.
   Caption: "Every tracker: headline numbers, findings, cleaner charts".
   *"Every tracker follows the same pattern: the numbers first, then the findings, then the
   detail — Furlough, Coaching, Overtime, SAFE, Absences and Training."*

6. **Master Operations Tracker — 1:05–1:20** — `side-by-side/12_master_ops.png`.
   Number cards: "Page size: ~29,000 → 4,100 elements", "Scrolling: 2.8 s → 0.24 s of work
   per second", "Tower filter: 0 agents → everyone with Tower hours".
   *"The Master Operations Tracker was the slowest page. It now builds an agent's detail only
   when you open it, scrolls smoothly, and the Tower filter finally shows the people who
   worked Tower."*

7. **Floor and tools — 1:20–1:38** — 3 s each: `04_genesys`, `13_red_flags`, `14_scripts`,
   `15_weather`, `16_calculator`.
   Captions: "Genesys: coverage gaps first", "Red Flags: the rule and the action, per table",
   "Calculator: result against target".
   *"The Genesys floor shows queue gaps first, Red Flags explains each rule and what to do
   about it, and the calculator presents its result against the target — with the same
   report you copy into Chat."*

8. **Also fixed — 1:38–1:52** — a clean checklist that ticks in line by line:
   - "Unplanned absences counted correctly"
   - "Maintenance banner with countdown and alerts"
   - "Outages keep their last good reading when a feed goes down"
   - "Data coverage: see which days are missing, and import them"
   - "Copy for Chat keeps its formatting"
   - "Genesys: Roster and Queues merged into one page"
   *"Plus the fixes the floor asked for — absences, maintenance alerts, outage feeds, data
   coverage, one Genesys page, and reports that paste cleanly into Chat."*

9. **Close — 1:52–2:00**
   "MRC Operations Portal" / "Faster. Clearer. One design across every page."
   *"Faster, clearer, and one design across every page."*

**Do not** invent features, numbers or names beyond what's above; keep every figure
exactly as written.
