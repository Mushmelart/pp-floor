# PP Floor — install

Four tablets' worth of setup, once. Roughly 30 minutes.

## How the pieces talk

```
  Art room (dashboard, 8002)
        │  artist clicks Agree / Disagree+reason
        ▼
  Flask /port/approve ──── pushes the release ────┐   outbound https
        │                                          │   from the office Mac
        ├── logs/port_approvals.json (local truth) │
        │                                          ▼
        │                            Google Sheet "PP Floor"
        │                            RELEASES · PRESS · DARKROOM
        │                                    ▲        │
        │                          reads     │        │  weekly PDF
        │                          rotations │        ▼
        └──────────────────────►  Tablets (GitHub Pages, https)  ──► email
```

The tablets never talk to the office machine. They can't — an https page is
not allowed to fetch `http://<shop-ip>:8002`, and no server setting changes
that. So the pipeline pushes releases out to the sheet and the tablets read
them back over https.

That's the better arrangement anyway:

- a tablet works on any network, not just shop wifi
- the press still sees its rotation when the office Mac is asleep
- the floor only ever sees rotations the art room actually released

The cost: a job the art room has **not** released has nothing on the tablet.
The operator gets "Nothing released for 754205 yet" instead of a raw
prediction. That is the correct behaviour — an unreleased rotation is not
something anyone should be printing from.

## 1 · Apps Script

1. Open the **PP Floor** sheet → Extensions → Apps Script.
2. Delete the boilerplate. Paste `PP_Floor_AppsScript.gs`.
3. File → New → Script → paste `PP_Floor_Report.gs`. Set `RECIPIENTS` at the
   top of it.
4. Deploy → New deployment → **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Copy the URL. It ends in `/exec`.

"Anyone" is what lets the tablets and the pipeline POST without an OAuth
handshake. `SHARED_TOKEN` is the actual gate, and the sheet itself stays
private.

The tabs (PRESS, DARKROOM, RELEASES) create themselves with their headers on
the first write. Run `formatSheets()` once by hand afterwards to widen the
readable columns.

## 2 · Point the pipeline at it

In `VectorSEP_PP/config.json`, `print_order` section:

```json
"floor_sheet_url": "https://script.google.com/macros/s/…/exec",
"floor_sheet_token": "pp-floor-dzQCyQbIUaPF2uZfOlkezE_z"
```

The token is already filled in and matches the Apps Script. Restart the 8002
app.

Test: approve a rotation in the dashboard. A row appears on the RELEASES tab
within a few seconds. If it doesn't, the Flask log line is
`floor release push failed:` — the push runs on a background thread, so a dead
webhook never makes the artist wait.

## 3 · Host the tablet page

The tablet page is `index.html` in this folder.

1. **`CFG.SHEET_URL` and `CFG.SHEET_TOKEN` both stay empty here.**

   GitHub Pages only publishes from a **public** repo on a free plan, so
   everything committed to this repo is world-readable. Both values are
   entered once per tablet in ⚙ and live in that tablet's browser storage, so
   the published source carries nothing worth having.

   ⚙ takes them as one paste — the `/exec` URL, a `|`, then the token:

   ```
   https://script.google.com/macros/s/AKfyc…/exec | pp-floor-…
   ```

   Pasting just a URL later repoints a tablet at a new deployment and keeps
   its token.

2. Create the GitHub repo — `pp-floor`, **public**.
3. Push this folder to it.
4. Settings → Pages → Deploy from a branch → `main` / root.

Deploying a change afterwards is: edit `index.html`, commit, push. Pages
redeploys in about a minute. Add `?v=2` to the URL to force a tablet past its
cache while testing.

## 4 · The tablets

On each one, in Chrome:

1. Open the Pages URL.
2. ⚙ → paste the setup line (URL `|` token) from step 3.
3. ⚙ → **Where is it**: *At a press* or *Screen room*. This decides which tabs
   the tablet shows — a press operator never scrolls past a screen-room log,
   and the dark room never sees a press rotation.
4. Station name (`Press 3`, `Dark room`) and the operator's name.
5. Menu → **Add to Home Screen**. It opens full-screen from then on.

Built for 11" / 1200×1920 portrait. Nothing tappable is under 64px.

## 5 · The weekly report

In the Apps Script editor, pick `installWeeklyTrigger` from the function
dropdown and press Run. It clears its own duplicates, so running it twice is
safe. Monday 7am.

`sendWeeklyReport` sends one immediately if you want to see it now.

## The numbers, and what they refuse to do

Every threshold is a named constant at the top of the file that uses it — one
knob each, no hunting.

| | |
|---|---|
| Setup ceiling | 25 min/screen |
| Run ceiling | 600 min |
| Wait ceiling | 480 min |
| Setup target | 8 min/screen |

A self-reported timer's failure mode is being left running, so there are three
guards: a check-in prompt that stops and flags an unconfirmed timer, a
pause-and-hold that freezes a parked job's clock, and the hard ceilings above.

When a value runs away it is **excluded from every average and listed by order
number**, never quietly trimmed. Every average that touches a runaway-prone
number applies the same test, so no two figures in the report can contradict
each other.

Setup rate is normalised per screen (setup ÷ colors), which is what makes a
physically impossible value visible regardless of how big the job was.

## What isn't built yet

- No offline cache of releases. A tablet with no connection can still run its
  timers and queues its submissions in the browser, but it can't look up a
  rotation.
- Monthly / quarterly roll-ups. The weekly is the pattern; they're the same
  script with a wider window.
- The HOW-TO tab (the EN+ES technique sheets) is out for now. The Flask route
  `/floor/procedures/<name>` is still there for when it goes back in.
