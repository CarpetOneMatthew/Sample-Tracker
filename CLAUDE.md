# Sample Tracker — CLAUDE.md

## What this is
Single-page vanilla JS/HTML/CSS app for tracking flooring samples ordered from
suppliers across commercial projects (RDCO commercial division). No build step,
no package.json, no framework — `index.html` (~1,570 lines) contains structure,
inline `<style>`, and inline `<script>` in one file. Hosted on GitHub Pages,
served directly from this repo.

## Architecture
- **Frontend**: `index.html` only. Everything — markup, CSS, JS — lives inline.
  Static assets (favicons, icons, `background.webp`, `manifest.json`) support
  PWA install/splash and are wired via `<link>`/`<meta>` tags in `<head>`.
- **Backend**: Google Apps Script Web App, NOT in this repo. Lives in the
  Apps Script editor as `Code.gs`, deployed as a `/exec` endpoint. The frontend
  calls it via `fetch` for `getAll` (GET) and writes (POST) — see the
  `CONFIGURE ME` block near line 600 of `index.html` for `API_URL`/`API_TOKEN`.
- **Database**: A Google Sheet. `Code.gs` reads/writes tabs via `upsert_()`.

## Critical gotcha — read before touching Code.gs
`upsert_()` in `Code.gs` writes **positionally**, driven by a hardcoded headers
array (`TABS.LINES.headers`). If a column is added to the sheet without also
updating that headers array in the same change, writes silently misalign and
corrupt data — no error thrown. Any schema change to `SampleLines` requires
updating both the sheet AND `TABS.LINES.headers` together.

## Known accepted risk (parked, not a bug to "fix" unprompted)
`API_URL` and `API_TOKEN` are plaintext in `index.html`, visible via View
Source. This repo is public. This has been reviewed and deliberately parked —
don't flag it as new news or start hardening it without being asked.

## Visual system (current — flat console aesthetic)
- Typefaces: Archivo (UI text) + IBM Plex Mono (numerals/data)
- Hairline separators between rows, not cards/shadows
- Border radii: 3–4px, no drop shadows
- Teal accent reserved strictly for interactive elements
- Status shown once, via a single colored dot — not repeated with badges/text

### The CSS is TWO stacked layers — read both before styling anything
`index.html` has one `<style>` block containing two systems in sequence:

1. **"Vista"** (from the top) — the ORIGINAL card-based sheet. SF Pro/Inter,
   `--radius:18px`, real shadows, status encoded three ways (a coloured left
   bar on `.rowcard::before`, a filled pill, and the status word).
2. **"FLAT LAYER"** (from the `FLAT LAYER — the console system` banner
   comment onward) — overrides Vista **wholesale** and is what actually
   ships. It redefines `:root` (`--radius:4px`, `--shadow:none`, Archivo,
   IBM Plex Mono), flattens `.rowcard` to hairlines, kills
   `.rowcard::before`, and re-encodes status as a single dot via
   `.pill::before`.

Everything in the bullet list above is the FLAT LAYER. Vista is dead weight
kept only because the flat layer overrides rather than replaces it.

**Author all new CSS in the flat layer, at the bottom.** Styling a new
component up in the Vista block looks right in isolation and then reads
wrong in the app: it inherits Vista's radii and shadows, and a more
specific selector (e.g. `.mything .count`) silently beats the flat layer's
`.count{font-family:var(--mono)}`, so figures come out in the wrong face.
Reading only the top of the file will mislead you about the entire design.

## Behavioral conventions already built in
- Auto-refresh: silent 5-min background poll. Defers if a modal is open, a
  save is in-flight, or the user is actively typing — never interrupts an
  in-progress edit.
- Supplier history view: every sample ever ordered from a supplier, across
  all projects, inline-editable, with Overdue/Problems/Excluded/Unmeasured
  filters. Speed/accuracy badges are computed off this data — editing a
  record here corrects the scoring.
- Handed-off archive (project modal only): once a sample is received AND
  handed off it folds behind a `HANDED OFF · n` disclosure row, so a long
  project shows only what still needs doing. Purely derived —
  `isArchived()` is `Status === 'Delivered'`, nothing is stored, no column,
  no backfill. Expanding gives back the real rows, still fully editable.
  Anything handed off while the project is open stays visible until you
  reopen it (`justHandedOff`), so a row never vanishes under your cursor.
  **Archiving is visibility only — archived samples still count toward
  supplier speed/accuracy scoring.** `Exclude` remains the one and only
  mechanism for removing a sample from the SCORING.
  Today and Receive never rendered Delivered lines, and `markDelivered()`
  already flips a project to Complete once every line lands (and the
  Projects list defaults to Active), so the project modal was the only
  surface that needed this.

## Workflow note
Previous commits were made by uploading files through the GitHub web UI, not
git. Treat this as effectively the first real git history for this repo —
don't assume commit messages before now reflect atomic, reviewable changes.
