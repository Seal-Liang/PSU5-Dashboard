# plans/

In-repo home for **planning documents** — task breakdowns, implementation
plans, design proposals, investigation notes, and handoffs.

## Why this exists

By default Claude Code writes plan-mode documents to the global
`~/.claude/plans/` folder, where they're detached from the project and
invisible to teammates. This project keeps its plans **in the repo** so they
are versioned alongside the code they describe and travel with `git`.

## Convention

- One Markdown file per plan, named `YYYY-MM-DD-<short-slug>.md`
  (e.g. `2026-07-03-offline-vendoring.md`).
- Put status/progress notes for an active effort at the top of its plan file;
  don't spread one initiative across several files.
- These files are committed. Delete or archive a plan once it's fully
  superseded — stale plans are worse than none.

## Existing files

- `task_plan.md`, `progress.md`, `findings.md` — legacy planning notes from the
  original Gemini-era build, kept for history.
