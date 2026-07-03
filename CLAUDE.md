# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A local-first **Weekly Project Status Dashboard**. Users upload weekly Excel/xlsm reports (work hours per project) plus a BU classification file; the app parses them into a single `data.json` store and renders resource-distribution and workforce-utilization charts. It ships two ways: run as a local dev server, or bundled by PyInstaller + `flaskwebgui` into a standalone Windows desktop `.exe` for non-technical colleagues.

Tech stack: **Flask** (thin API/router) · **pandas + openpyxl** (Excel parsing) · **vanilla JS + Chart.js + html2canvas + JSZip** (frontend, no build step). No database — state is a JSON file.

## Planning documents

Write plan-mode docs, task breakdowns, design proposals, and handoff notes into **[`plans/`](plans/)** in this repo — **not** the global `~/.claude/plans/`. Keep them versioned with the code. One Markdown file per plan, named `YYYY-MM-DD-<slug>.md`. See [plans/README.md](plans/README.md).

## Commands

```bash
python server.py          # Run locally. flaskwebgui opens a desktop window on http://localhost:5000
build.bat                 # Windows: install deps + PyInstaller one-file build into WeeklyDashboard_Release/
python check.py           # Diagnostic: report how many WeeklyData records have unmapped BU (bu == "Unknown")
python rescue.py          # Repair a corrupted data.json from the newest .bak in %APPDATA%\WeeklyDashboard
pip install -r requirements.txt
```

There is no test suite, linter, or type checker configured. `FileSamples/` holds `*_FAKE*.xlsx` fixtures with the exact real-world layout — use these to validate parser changes manually.

## Architecture (3 layers)

Data flows one direction: **HTTP → parse → persist → serve → render client-side**.

- **Layer 2 — Routing** ([server.py](server.py)): Flask endpoints only. Saves uploads to `.tmp/`, calls the tools, returns JSON. Holds zero analytics logic. Endpoints: `/api/upload_reports`, `/api/upload_bu`, `/api/upload_holidays`, `/api/upload_logs`, `GET /api/data`, `POST /api/settings`.
- **Layer 1 — Tools** ([tools/excel_parser.py](tools/excel_parser.py), [tools/data_store.py](tools/data_store.py)): deterministic Excel→dict parsing and all read/write/merge logic for `data.json`. Keep business logic here, not in `server.py`.
- **State** (`data.json`): the single source of truth. Schema is documented in [gemini.md](gemini.md) (`Settings` / `BUMapping` / `WeeklyData`). It is `.gitignore`d and can be multi-MB.
- **Presentation** ([static/](static/)): `index.html` + [static/js/app.js](static/js/app.js) (~1400 lines, all UI + **all utilization math**). The backend never computes utilization — it only stores raw hours, headcount logs, and holidays; the client derives everything.

The `architecture/*.md` files are the original per-layer SOPs. Treat them and `gemini.md` as design intent, but this file plus the code are authoritative when they conflict.

## Invariants that break things silently if changed

These encode the real Excel format and are easy to violate without an error:

- **Header row.** Weekly-report sheets have their header at **row 5** (`xl.parse(sheet, header=5)`, 0-indexed). Required columns after strip: `PROJECT NAME`, `PIC`, `BU`, `Hours`. A sheet missing any is silently skipped.
- **The Excel `BU` column is actually the Product Code.** It is looked up (uppercased) against `BUMapping.product_code` to resolve real `bu`/`bg`. Unmatched codes are kept with `bu`/`bg` = `"Unknown"`.
- **Week metadata** lives at sheet 0, row 0, col 0 as `"2025WK31 : 2025/07/28"` (split on `" : "`, regex fallback).
- **Skipped sheets:** `Total` and `AMS` are not parsed as departments. `Total` is re-read separately to inject leftover hours (see below).
- **Comma product codes** (`"BZHY, ABCD"`) split one row's hours **evenly** across each code into separate records.

## Non-obvious mechanisms (require reading multiple files)

- **source_type dedup.** `parse_weekly_report` tags every record `GENERAL`, `ID_DETAIL`, or `UX_DETAIL` based on the **filename** (`MBDC-ID` / `MBDC-UX` substring). On re-upload, `save_weekly_data` drops existing records of the same `source_type` for that `week_id` (clean re-upload), and when a detail file is present it **suppresses the `GENERAL` MBDC-ID/UX summary rows** so hours aren't double-counted. Individual-person sheets in detail files are remapped to their parent department (`MBDC-ID`/`MBDC-UX`).
- **Total-sheet leftover injection.** For detail files, the parser diffs each project's hours in `Total` vs the sum across sub-team sheets and injects only the positive difference under the parent department — this captures hours logged directly at parent level. Changing sub-team parsing means re-checking this second pass.
- **Utilization is client-side.** In `app.js`: weekly capacity = `dynamicHeadcount × workingDays × 8`. `workingDays = (week.working_days || 5) − holidaysFallingInThatMon–Fri window`. `dynamicHeadcount = team.initial + Σ(log.delta for logs on/before the week date)`. `MBDC-ID`/`MBDC-UX` roll up their `-1/-2/...` children, but the parent rollup is skipped if any child has its own headcount configured (prevents double-counting capacity). Configure headcount at sub-team **or** parent level, never both.
- **Headcount deltas:** log types containing `Leave` / `Depart` / `NoTracking` / `離職` / `留停` → `delta = -1`; everything else → `+1`. This string matching is duplicated in `server.py` (bulk upload) and `app.js` (manual add) — keep them in sync.

## Data-store hazards

- **`data.json` location differs by mode.** Dev: project root. Frozen `.exe`: `%APPDATA%\WeeklyDashboard\data.json` (first run auto-copies a `data.json` sitting next to the exe). Never assume the root file is live in a packaged build.
- **Encoding robustness is load-bearing.** Excel/CSV reads fall back through utf-8 → big5; `load_data` retries big5 then resaves as UTF-8; a fully corrupt `data.json` is renamed to `.bak_<ts>` and reinitialized empty. Preserve these fallbacks — the source files are Windows/Traditional-Chinese exports.
- Writes go through a `threading.Lock`; keep using `save_data` rather than opening the file directly. Uploads write to `.tmp/` and delete in a `try/except finally` because Windows keeps pandas file handles locked (`WinError 32`) — don't remove the guarded cleanup.

## Conventions

- Product codes and BU lookups are compared **uppercased and stripped** everywhere. Match that when adding lookups.
- Bump the `app.js?v=N` query string in [index.html](static/index.html) when shipping frontend changes so the packaged app busts its cache.
- **All frontend assets are vendored** (offline-first): Chart.js, JSZip, and html2canvas live in `static/js/`; the Inter font (latin subset) is self-hosted in `static/fonts/` via `static/css/inter.css`. Do **not** reintroduce CDN `<script>`/`<link>` tags — the packaged `.exe` must render with no internet. `--add-data "static;static"` bundles all of it automatically.
