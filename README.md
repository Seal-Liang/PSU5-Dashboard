# Weekly Project Status Dashboard

A local-first dashboard for analysing weekly project work-hours. Upload the team's weekly Excel/xlsm status reports and a BU classification file; the app parses them, stores everything in a single local `data.json`, and renders resource-distribution and workforce-utilisation charts.

Runs two ways: as a local Python web app, or as a standalone Windows `.exe` you can hand to colleagues — **no installation, no internet required**.

---

## Quick start (developers)

```bash
pip install -r requirements.txt
python server.py
```

`server.py` opens a desktop window at <http://localhost:5000> (via `flaskwebgui`). Your data is read from / written to `data.json` in the project root.

## Build the standalone app (Windows)

```bat
build.bat
```

This installs dependencies + PyInstaller and produces a one-file executable in **`WeeklyDashboard_Release\`**. Zip that folder and send it to colleagues — they just double-click `WeeklyDashboard.exe`. All assets (charts, fonts, scripts) are bundled, so it works fully offline.

> In the packaged app, data is stored in `%APPDATA%\WeeklyDashboard\data.json`, **not** next to the exe. On first launch, a `data.json` placed beside the exe is auto-imported.

---

## How to use

The app has three views (left sidebar):

### 1. Upload Data
- **Weekly Status Reports** — drag & drop one or many `.xlsx` / `.xlsm` files. Re-uploading the same week overwrites it cleanly.
- **BU Classification** (optional) — an `.xlsx` mapping Product Code → BU / BG / Product Name. Uploading it also re-labels all previously imported records retroactively.

### 2. Dashboard
Filter by **Team**, **Time Frame** (All / Year / Quarter / Month), **Interval**, and **Group By** (BU or Product Code). Charts: hours distribution (pie), hours-by-category trend, workforce utilisation %, seasonality radar, and top-10 projects. Every widget has copy-to-clipboard and download-as-image buttons; "Download All" exports a ZIP.

### 3. Settings
- **Holidays** — add manually or bulk-upload (columns `Date`, `Name`). Holidays reduce a week's available working days, which raises the utilisation rate.
- **Headcount** — per team, set an **initial baseline** and a dated **change log** (onboard `+1`, depart / leave / no-tracking `−1`). Utilisation capacity is computed dynamically per week from these.

---

## Input file format (important)

The weekly Excel reports must follow the team's established layout, or rows are silently skipped:

- **Header row is row 5** (the 6th row); data starts below it.
- Required columns: `PROJECT NAME`, `PIC`, `BU`, `Hours`. The **`BU` column actually holds the Product Code**, which is matched against the BU Classification file to resolve the real BU/BG.
- The first cell (row 1, column 1) holds week metadata as `2025WK31 : 2025/07/28`.
- Sheets named `Total` and `AMS` are treated specially / skipped.
- Files whose name contains `MBDC-ID` or `MBDC-UX` are parsed as detailed per-sub-team reports.

Sample files with the exact expected layout are in [`FileSamples/`](FileSamples/).

---

## Tech stack

| Layer | Tech |
|-------|------|
| Backend / API | Flask |
| Excel parsing | pandas + openpyxl |
| Storage | a single `data.json` file (no database) |
| Frontend | vanilla JS + Chart.js (no build step) |
| Desktop packaging | flaskwebgui + PyInstaller |

Architecture and internal invariants for contributors are documented in [CLAUDE.md](CLAUDE.md) and [`architecture/`](architecture/).

---

## Troubleshooting

- **`data.json` won't load / looks corrupted** — the app auto-backs-up a bad file to `data.json.bak_<timestamp>` and starts fresh. To recover the packaged app's data, run `python rescue.py` (repairs the newest backup in `%APPDATA%\WeeklyDashboard`).
- **Check how many records are unmapped** (BU shows as `Unknown`) — run `python check.py`. Usually means the BU Classification file is missing some Product Codes.
- **`pip install` fails building pandas wheels** — you're likely on a too-new Python. Use Python 3.11/3.12, or install without pinned versions.
- **Charts don't update after a code change** — bump the `?v=N` on `app.js` in `static/index.html` to bust the browser cache.

---

## Project layout

```
server.py              Flask routes (upload / data / settings)
tools/
  excel_parser.py      Excel -> normalized records
  data_store.py        read / merge / persist data.json
static/
  index.html           single-page UI
  js/app.js            all UI + utilisation math
  js/, css/, fonts/    vendored libraries + self-hosted font
architecture/          per-layer design SOPs
FileSamples/           sample input files
build.bat              one-click Windows build
```
