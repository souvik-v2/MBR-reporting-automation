# MBR Pipeline — how it works, and what changed

This document describes the automated Monthly Business Review (MBR) deck
pipeline: `build_mbr.py` and the modules it calls. If you're looking for
the generic Excel/CSV/PPTX merger tool, see [README.md](README.md) instead
— that's a separate, unrelated tool that happens to live in the same repo.

## Quick start

```bash
pip install pandas openpyxl python-pptx
# .odp team decks also need LibreOffice installed (headless conversion):
#   sudo apt install libreoffice      (Linux)
#   brew install --cask libreoffice   (macOS)
#   or download from libreoffice.org  (Windows)

python3 build_mbr.py \
  --template sample_data/final-data/06-01-2026.pptx \
  --graph-dir sample_data/Graph \
  --ppt-data-dir sample_data/ppt-data \
  --month March \
  --report-date 2026-04-02 \
  --out output/MBR_March_2026.pptx
```

`--report-date` stamps the title slide; omit it to use today's date, or
pass `--no-date-update` to leave the title slide's date untouched.

This prints a per-project summary of every chart updated and every slide
copied, and writes `output/MBR_March_2026_REVIEW_CHECKLIST.md` listing
anything that needs a human's judgment call before the deck goes out.

`--template` should be **last month's finished MBR deck** (it already
contains the correctly-formatted native charts and section structure;
this run just refreshes the data and adds this month's content), not a
blank template.

## What this pipeline automates, end to end

1. **Reads weekly delivery metrics** for every project from whichever
   `LT-SimplifiedMetrics-*.xlsx` file is available (`graph_data_extractor
   .py`). These files are full copies of one master workbook containing
   every project's sheet, so any one of them works as the source.
2. **Updates the native PowerPoint charts** ("Rework Trend", "Throughput
   Distribution", "Carryover Trend") on each "Project Metrics - ..."
   slide with real data for the target month, using python-pptx's
   `chart.replace_data()` (`chart_updater.py`). These are real, still-
   editable PPT charts afterward — not pasted-in images.
3. **Converts any `.odp` team deck to `.pptx`** automatically via headless
   LibreOffice (`odp_converter.py`), so a team submitting in OpenDocument
   format doesn't break the pipeline.
4. **Copies each team's "Goals Achieved" / "In Progress" slides** into the
   consolidated deck, byte-for-byte (text formatting, tables, images, and
   embedded charts all preserved), inserted right after the "Project
   Details" section divider (`slide_copy.py`, `goals_consolidator.py`).
5. **Updates the title slide's date stamp** to the last day of the
   reporting month (e.g. "31-March-2026" for a March report) -- derived
   from the actual Graph data for that month, not from whatever day the
   script happens to be run on. Pass `--report-date YYYY-MM-DD` to set a
   specific date instead, or `--no-date-update` to leave it alone.

6. **Optionally inserts each team's full Goals Achieved / In Progress
   slides** into the deck, if `--insert-goal-slides` is passed (off by
   default). The Key Highlights table in step 4 is refreshed either way
   -- this flag only controls whether the full original slides also get
   added as extra pages.
7. **Flags anything that needs a human look** instead of guessing: a
   project with no team deck configured, a slide title that looks
   relevant but doesn't match the standard pattern, etc.

## What this pipeline deliberately does NOT automate

Two pieces of human-authored content are left untouched, on purpose:

- **The "Key Highlights Across Workstreams" table** (the slides right
  after the title slide).
- **The "Key Achievements and Project Roadmap" narrative text box** on
  each Project Metrics slide.

Both are short, hand-picked summaries of a much longer Goals Achieved
slide. Earlier code in this project tried to auto-generate this kind of
summary with a formatting heuristic ("a bullet is a heading if it's
capitalized, doesn't end in a colon, and is 10-150 characters long").
Tested against this project's real team decks, that heuristic isn't
reliable — some teams bold their heading bullets and some don't, so it
would sometimes misclassify a detail bullet as a heading or vice versa.
Rather than risk silently misrepresenting what a team reported, the
pipeline copies their original slide untouched and leaves the *summarized*
version for a person to write, with the full source material now sitting
right there in the deck to reference. The `_REVIEW_CHECKLIST.md` file
this script generates points you at exactly this.

## "LT Manager" names -- read from the Graph workbook's "Client POC" field

The Key Highlights table's "LT Manager" column (the LendingTree-side
account manager per project) is read from the **"Client POC"** field in
each project's sheet inside the Graph workbook
(`sample_data/Graph/LT-SimplifiedMetrics-*.xlsx`). That field had no
other use in this sheet (every project's actual client-feedback notes go
in the separate "Client Feedback" field), so it's repurposed here to
carry the LT manager's name.

To change an LT manager, edit that cell directly in the Graph workbook --
no code change needed. `ProjectGraphData.lt_manager`
(`graph_data_extractor.py`) reads it dynamically every run.

`ProjectEntry.lt_manager` in `project_registry.py` still exists as a
**fallback only**, used if the Excel field is ever blank for a project
(seeded with the real names found in the original template: Robert
Ploeger / Douglas Brown / Sushanth Kancharla / David Wilson). The Excel
value always wins when present.

## "Key Achievements" slide and text box

Several team decks have a real slide -- titled "Key Achievements" (FLEX)
or "Key Contributions Beyond Ticket Goals" (CanopyNext, Core_SVC/
Consumer Web) -- with genuine, team-authored content distinct from the
Goals Achieved / In Progress slides. `key_achievements_finder.py` looks
for this slide by title and, when found:
  - copies it into the consolidated deck right after the project's
    "In Progress Ticket" slide (or after "Goals Achieved" if that title
    isn't present in this deck) -- this happens regardless of whether
    `--insert-goal-slides` is on, since it's separate, always-useful
    content;
  - replaces the body bullets under the metrics slide's "Key
    Achievements and Project Roadmap" heading with that same real text,
    leaving the heading and the "Risks / Constraints" section untouched.

Projects whose deck has no such slide (SEOD, WPM, CET, Phoenix as of
this writing) keep their existing metrics-slide text exactly as it was
in the template -- nothing is invented when the source material doesn't
exist.

## PDF export

`build_mbr.py` exports a PDF copy of the finished deck automatically
(via headless LibreOffice -- the same dependency already required for
.odp conversion). Pass `--no-pdf` to skip it. If LibreOffice isn't
installed, this is a soft skip with a message, not a build failure.

## Chatbot (chatbot_app.py)

A small Flask app (`chatbot_app.py`) serves a browser chat window at
http://localhost:5000, backed by:
- `knowledge_base.py` -- builds one in-memory snapshot of every
  project's manager names, achievements, in-progress items, and roadmap
  by reusing the exact same extraction functions build_mbr.py uses, so
  the chatbot and the generated reports can never disagree with each
  other.
- `chat_engine.py` -- a rule-based intent classifier + project-name
  matcher (no external AI API; see its module docstring for why).
  Recognizes the example questions from the spec directly: key
  achievements, project roadmap, project manager, LT manager,
  in-progress/completed tickets, and "generate the latest report".

The knowledge base is rebuilt fresh on every question -- there is no
caching, so editing a source file and asking again immediately reflects
the change.

## Slide titles and the Key Highlights table's "Project" column are now dynamic

Both now use the Graph workbook's **"Sub Project"** field, read fresh
every run -- not the static text baked into the template. Renaming
happens on the metrics slide ("Project Metrics - {Sub Project}") and on
the matching row's "Project" cell in the Key Highlights table.

**By explicit request, this is intentionally unvalidated.** Whatever is
typed into a project's "Sub Project" cell in
`sample_data/Graph/LT-SimplifiedMetrics-*.xlsx` is shown exactly as-is --
including if it's wrong, blank, or copy-pasted from a different
project's row. There is no plausibility check against the project's
known identity (sheet name, slide title hint). Keeping this field
accurate is the data owner's responsibility; the pipeline will not
second-guess or silently keep the old title if it looks wrong.

The rename always re-matches against the *original* template's static
title text on every run (`build_mbr()` loads a fresh, unmodified copy of
the template each time), so this never breaks next month's slide
lookup -- there's nothing to "undo," since the rename only ever exists
in that run's output file, never in the template itself.

## Two bugs fixed after real-world testing in PowerPoint

**Blank/overlapping copied slides in real PowerPoint.** Slides copied
from team decks (slide_copy.py) looked fine in LibreOffice but showed
blank or overlapping text in actual PowerPoint. Cause: placeholder
shapes (title, body, slide number) often have no explicit position --
they inherit it from a matching placeholder on their *own* slide's
layout. The destination slide is built on a "Blank" layout with no
placeholders to inherit from, so a copied shape with no explicit
geometry had nothing to fall back to once moved. LibreOffice tolerates
this with some default; PowerPoint doesn't. Fixed by resolving each
shape's effective position on the source (where inheritance still
works) and baking those exact values into the copy explicitly, so it no
longer depends on inheriting from anything.

**Editing one Graph Excel file was silently ignored.** Every
`LT-SimplifiedMetrics-*.xlsx` file *should* be an identical copy of the
same master workbook -- but the moment any one file is edited
independently (the normal, expected way someone would test or update
data), that's no longer true. The pipeline previously always read from
whichever file `glob()` happened to list first alphabetically, so an
edit made to `LT-SimplifiedMetrics-seod.xlsx` could be silently
overridden by stale data still sitting in a *different* file's copy of
the SEOD sheet. Fixed in `graph_workbook_for_entry()` (build_mbr.py):
each project now reads from its own matching file first (matched by
filename slug against the project's sheet name / slide title hint),
falling back to "any file" only if no project-specific file exists.

If you're testing changes by editing sample_data/ and re-running, this
is what makes that actually work now -- both data changes (Excel) and
slide-template changes should be picked up on the next run with no
caching.

## The project registry — the one file you'll need to maintain

`project_registry.py` is the single source of truth mapping a project
across three naming systems that, in practice, don't line up
algorithmically:

| | Example |
|---|---|
| Template slide title | `"Project Metrics - SLE"` |
| Graph workbook sheet name | `"CanopyNext"` |
| Team deck filename | `V2Solutions_LT_CanopyNext_March_Report.pptx` |

Real examples of why this can't be auto-matched reliably:
- The CanopyNext Graph sheet's project is labeled **"SLE"** on its slide.
- "Consumer Web" the slide / "Consumer Web Kanban" the sheet is produced
  by a deck whose filename says **"Core_SVC"** and whose title slide says
  **"LT Consumer Website (LCW)"**.
- "Consumer Experience Team (CET)" has its own Graph sheet but no
  dedicated Graph filename and no team deck at all.

When a project is renamed, added, or its deck filename pattern changes,
add or edit one `ProjectEntry` in `project_registry.py`:

```python
ProjectEntry(
    sheet_name="CanopyNext",       # exact sheet name in the Graph workbook
    slide_title_hint="SLE",         # substring expected in the slide title
    deck_glob="*CanopyNext*.pptx",  # filename pattern in ppt-data/
    display_name="SLE (CanopyNext)",
),
```

Set `deck_glob=None` for a project with no team deck (its Goals/In-Progress
slides simply won't be inserted, and this is reported, not silently
skipped).

## Chart field mapping

`chart_updater.py`'s `CHART_FIELD_MAP` controls which columns from the
weekly metrics feed each chart's series. The source data has a single
combined "Rework" column (no V2/LT split), so the "Rework Trend" chart's
V2 series is currently reported as 0 with all rework attributed to LT —
this matches what's actually in the real data (V2's reopen/rework figures
are 0 in every sampled week) rather than guessing a split that doesn't
exist. If a project ever reports a genuine V2/LT rework split, update
`_v2_rework` / `_lt_rework` in `chart_updater.py` accordingly.

## Slide-matching pattern for Goals/In-Progress extraction

`goals_consolidator.py`'s `STRICT_PATTERN` matches slide titles containing
"Goals Achieved", "In Progress Ticket", or a "Continue…." continuation
slide. This was checked against all 6 real team decks in this project and
correctly identifies the relevant slide(s) in every one. Titles that look
related but don't match (`NEARMISS_PATTERN`: "Carryover", "Reopen",
"Completed Epic", etc.) are reported in the review checklist rather than
silently dropped or silently guessed into the deck.

## Verifying a build visually

Since the output is a real `.pptx`, the easiest sanity check is to convert
it to PDF and look at it:

```bash
soffice --headless --convert-to pdf --outdir output output/MBR_March_2026.pptx
```

## What changed from the previous version of this pipeline

The repository previously contained ~10 overlapping scripts
(`build_complete_mbr.py`, `create_mbr_complete.py`, `create_consolidated_mbr
.py`, `update_template_with_march.py`, and others — now in `archive/old_scripts/`)
and documentation describing a Windows-COM-automation approach
(`pptx_auto_consolidator.py`, requiring `pywin32` and a PowerPoint
installation) that was never finished, alongside contradictory docs that
described manual copy-paste as the actual fallback.

Specific bugs found and fixed while rebuilding this pipeline:

- **`graph_data_extractor.py` read the wrong Excel sheet for every
  project** (`pd.read_excel(file, header=1)` reads pandas' default first
  sheet, never the actual per-project sheet), producing identical
  fabricated-looking numbers for every project. This bug is visible in
  the old `archive/old_docs/FINAL_SUMMARY.md`'s own example output. Fixed by
  detecting the correct sheet via its header signature and matching it
  against the registry.
- **Date parsing had no year disambiguation**, so a chronological Oct→May
  weekly series was assigned the wrong absolute year (and, for one
  project whose sprint-naming convention changes mid-series, would have
  misread a naming-convention transition as a year wraparound). Fixed
  with a two-pass approach that prefers any year embedded directly in a
  week label, falling back to a chronological walk only when absent.
- **`update_template_with_march.py`'s slide reordering was a no-op** — it
  mutated a Python list copy of the slide order, not the underlying XML
  (`<p:sldIdLst>`), so slides it claimed to insert at a specific position
  actually landed at the end of the deck. Fixed in `slide_copy.py`'s
  `move_slide_to()`, which edits the XML element list directly.
- **The actual chart-copying problem (the subject of most of the old
  docs) turned out not to need solving at all** — the real template
  already contains live, editable python-pptx chart objects on every
  Project Metrics slide. `chart_updater.py` updates them in place with
  `chart.replace_data()`. One real wrinkle: this template's charts link
  to an *external* SharePoint URL rather than an embedded workbook, which
  crashes python-pptx's `replace_data()` unless the external link is
  detached first — handled in `_detach_external_data_link()`.
- **`.odp` files (one team's monthly deck) crash python-pptx outright**
  (`KeyError` on the OPC package loader, since ODF isn't OOXML). Fixed by
  converting via headless LibreOffice first (`odp_converter.py`).
- **The documented slide layout didn't match the real template** — the
  docs described a single project-summary table on slide 3; the real
  template has a two-slide "Key Highlights" table (slides 3-4) and a
  separate, contentless "Project Details" divider slide (slide 6), which
  is where Goals/In-Progress content actually belongs.

## A mistake made (and fixed) while building this pipeline

The first pass at cleaning up this repo moved `mbr_report_builder.py` and
`chart_generator.py` to `archive/old_scripts/` without checking that `main.py`
(the separate, working generic-merger tool) actually imports them for its
own Excel-based Sprint MBR report (`output/mbr_report.xlsx`) -- a
completely different, unrelated feature from the broken PowerPoint-COM
pipeline this rebuild replaced. Both files are openpyxl-only (no COM, no
python-pptx) and were working correctly; they've been restored.

While fixing that, a second, pre-existing and unrelated bug surfaced:
`writer.py` and `mbr_report_builder.py` both crashed with
`IllegalCharacterError` on any PPTX-sourced text containing a vertical-tab
character (`\x0b`, which PowerPoint uses internally as a soft line-break
inside text boxes -- e.g. a manager-name cell like "Dayanand\x0bHarpreet").
This wasn't introduced by this rebuild, but since `main.py` was being
touched anyway, both files now sanitize cell values through openpyxl's
own `ILLEGAL_CHARACTERS_RE` before writing, replacing control characters
with a space rather than crashing the whole run.

If you're auditing this rebuild: the lesson is to grep for *all* callers
of a file before moving it, not just the ones already known to be broken.
