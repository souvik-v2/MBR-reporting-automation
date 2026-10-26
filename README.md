# Excel / CSV / PPTX → Master Excel Merger & Sprint MBR Reporting System

Scans a folder (including subfolders) for `.xlsx`, `.xls`, `.csv`/`.tsv`, and
`.pptx` files, reconciles their different column structures, and produces:

1. **Master Excel workbook** — keeping only the **latest** version of each record
   when the same ID shows up in multiple files
2. **MBR Report** — automated sprint metrics, KPI calculations, and consolidated
   reporting with charts and goal achievement tracking

## What it does

- **Excel/CSV**: reads every sheet of every file. Column headers don't need
  to match exactly — `Order ID`, `Ref No`, and `id` are all recognized as the
  same column if you list them as aliases in `config.py` (see below). Columns
  that aren't explicitly aliased are still merged together if they only
  differ by case/spacing (e.g. `Customer` and `customer`).
- **PPTX**:
  - **Tables** on slides are read like a spreadsheet — if the table's header
    row matches your known columns, its rows get merged in with everything
    else.
  - **Text** (titles, bullet points, body placeholders) goes into its own
    `PPTX_Text` sheet (slide number, placeholder type, text).
  - **Charts** (bar/line/pie, etc.) have their title, series names, and
    category/value data points extracted into a `PPTX_Charts` sheet, since
    chart data doesn't fit a generic row structure.
- **Deduplication**: rows sharing the same ID (per `DEDUP_ID_COLUMN`) are
  collapsed down to one — the one with the latest date in
  `DEDUP_DATE_COLUMN`. Every dedup decision is written to a `Run_Log` sheet
  so you can see exactly which row was kept and why.

## Output

The script produces **two files**, both built from the same merged data:

1. **`master_file.xlsx`** — the formatted workbook (open in Excel/Sheets)
2. **`dashboard.html`** — a standalone, scrollable web dashboard (open by
   double-clicking; works fully offline, no server needed)

### The Excel workbook has 5 sheets:

| Sheet         | Contents                                                               |
| ------------- | ---------------------------------------------------------------------- |
| `Master`      | One row per unique ID — the latest version of each record              |
| `All_Records` | Every row found, before deduplication (for audit)                      |
| `PPTX_Text`   | Slide titles / body text extracted from PPTX files                     |
| `PPTX_Charts` | Chart title + category/series/value data from PPTX charts              |
| `Run_Log`     | Plain-English log of every file read, column issue, and dedup decision |

### The HTML dashboard

A single page with:

- **Summary cards** at the top — master row count, total rows found, PPTX
  text/chart counts, and a warning/error count
- **Tabs** for Master / All Records / PPTX Text / PPTX Charts / Run Log —
  same data as the Excel sheets, browsable in one page
- **Search box** — filters every visible column as you type
- **Source-file filter** — narrow the table down to rows from one input file
- **Click any column header to sort**, click again to reverse
- A scrollable table with a sticky header, so you can scroll through
  thousands of rows without losing the column names

Just open `dashboard.html` in any browser (Chrome, Edge, Firefox, Safari) —
double-click it from File Explorer. No internet connection or install needed;
all the data is embedded directly in the file.

## Ask questions instead of reading the deck

```bash
python3 chatbot_app.py
```

Open http://localhost:5000 in a browser. Ask things like "What are the
key achievements this month?", "Who is the LT Manager for SEOD?", or
"What are the in-progress tickets for FLEX?". No external AI API is
used -- every answer is read directly from sample_data/ on the spot, so
editing a source file and asking again immediately reflects the change.

## Run everything in one command

```bash
python3 run_all.py
```

Re-run this any time files in sample_data/ change -- new team deck,
updated Graph workbook, new month's template. Nothing needs editing
first; the month, template, and project list are all auto-detected.
Produces (in output/): master_file.xlsx, dashboard.html, mbr_report.xlsx,
MBR_<Month>_<Year>.pptx, and a REVIEW_CHECKLIST.md.

## Monthly Business Review (MBR) deck automation

This is a **separate pipeline** from the merger tool above. It takes last
month's finished MBR PowerPoint deck, this month's per-project metrics
(Excel) and team status decks (PowerPoint/OpenDocument), and produces a
fully updated deck — no manual chart copying, no Windows, no PowerPoint
installation required.

```bash
python3 build_mbr.py \
  --template sample_data/final-data/06-01-2026.pptx \
  --graph-dir sample_data/Graph \
  --ppt-data-dir sample_data/ppt-data \
  --month March \
  --out output/MBR_March_2026.pptx
```

(Add `--report-date YYYY-MM-DD` to stamp a specific date on the title
slide, or `--no-date-update` to leave it as-is.)

This automatically:

- Reads each project's weekly metrics from the `LT-SimplifiedMetrics-*.xlsx`
  files and updates the native "Rework Trend", "Throughput Distribution",
  and "Carryover Trend" charts on each Project Metrics slide with real data
  for the target month, in place, using python-pptx's chart API.
- Converts any `.odp` team deck to `.pptx` automatically (via headless
  LibreOffice) and copies that team's "Goals Achieved" / "In Progress"
  slides into the deck, preserving original formatting exactly.
- Writes an `_REVIEW_CHECKLIST.md` file alongside the output, listing
  anything that genuinely needs a human's judgment call (a missing team
  deck, an unusually-titled slide that wasn't auto-matched, etc.) — see
  **[MBR_PIPELINE.md](MBR_PIPELINE.md)** for exactly what is and isn't
  automated, and why.

**📖 See [MBR_PIPELINE.md](MBR_PIPELINE.md) for complete documentation,**
**including the project registry you'll need to update if your team list**
**changes.**

The previous version of this pipeline's documentation (in `archive/old_docs/` and `archive/old_scripts/`)
described a Windows-COM-automation approach that was never completed, and
made claims about chart automation that didn't match the actual code. See
MBR_PIPELINE.md's "What changed" section for specifics.

## Setup

```bash
pip install pandas openpyxl python-pptx
# or on Windows:
py -m pip install pandas openpyxl python-pptx
```

## Configure

Open **`config.py`** and adjust to match your data — this is the only file
you should need to edit:

- `SOURCE_FOLDER` / `OUTPUT_FILE` — default paths (can be overridden on the
  command line instead, see below).
- `COLUMN_ALIASES` — group together any column names from different files
  that mean the same thing, e.g.:
  ```python
  COLUMN_ALIASES = {
      "ID": ["id", "record id", "order id", "ref no"],
      "Date": ["date", "order date", "timestamp"],
      "Customer Name": ["customer", "client", "account name"],
  }
  ```
- `DEDUP_ID_COLUMN` / `DEDUP_DATE_COLUMN` — which canonical columns to use
  for "same record" and "which is newest".
- `DATE_DAYFIRST` — dates like `07/01/2026` are ambiguous. Set this once to
  match your data instead of letting the script guess. ISO dates
  (`2026-01-07`) are unaffected — those are never ambiguous.
- `NO_ID_POLICY` / `KEEP_UNPARSEABLE_DATES` — what to do with rows that have
  no ID, or a date that couldn't be parsed.
- `EXTRACT_PPTX_TABLES` / `_TEXT` / `_CHARTS` — turn any of the three PPTX
  extraction types on/off.

## Run

```bash
python main.py
```

Or override paths without touching `config.py`:

```bash
python main.py /path/to/your/folder /path/to/master_output.xlsx /path/to/dashboard.html
```

(Each argument is optional and falls back to the matching `config.py` value if omitted.)

## Always check the Run_Log sheet

After every run, open the `Run_Log` sheet first. It tells you:

- which files were read and how many rows came from each
- any files/sheets that failed to read, and why
- every duplicate ID found, which file/row was kept, and why
- any dates that were ambiguous and how they were interpreted

This is the fastest way to confirm the merge did what you expected, since
your data's structure varies between files.

## Project structure

```
main.py                        merger tool entry point -- run this
build_mbr.py                   MBR PowerPoint pipeline entry point -- run this
run_all.py                     runs both of the above in one command
chatbot_app.py                  web chatbot -- run this, then open localhost:5000
requirements.txt
README.md                      this file
MBR_PIPELINE.md                full documentation for the MBR PowerPoint pipeline

src/
  merger/                      generic Excel/CSV/PPTX merger tool (used by main.py)
    config.py                  <- edit this to match your data
    column_utils.py             column-name matching + date parsing
    readers_tabular.py          Excel/CSV reading
    readers_pptx.py             PPTX table/text/chart extraction
    dedup.py                    "keep latest per ID" logic
    writer.py                   builds the formatted Excel workbook
    dashboard.py                builds the standalone HTML dashboard
    metrics_calculator.py       Sprint KPI calculations
    pptx_slide_extractor.py     extracts goal slides from presentations
    chart_generator.py          Excel-native chart generation (openpyxl)
    mbr_report_builder.py       builds mbr_report.xlsx (KPIs, trends, goals, run log)
    pptx_goals_extractor.py     standalone ad-hoc goals-slide extraction report

  mbr_pipeline/                MBR PowerPoint deck pipeline (used by build_mbr.py)
    project_registry.py        verified slide/sheet/deck name mapping (edit when teams change)
    graph_data_extractor.py    reads weekly metrics from LT-SimplifiedMetrics-*.xlsx
    chart_updater.py           updates native PPT charts via python-pptx's chart API
    slide_copy.py               cross-presentation slide copying (text/images/charts)
    goals_consolidator.py      finds + copies each team's Goals/In-Progress slides
    odp_converter.py           converts legacy .odp team decks to .pptx
    key_highlights_extractor.py  extractive Completed/In-progress summaries
    key_highlights_table.py    refreshes the Key Highlights table in place
    text_fit.py / textbox_fitter.py  shrink-to-fit for text boxes with fixed layouts
    auto_discovery.py          finds month/template/folders automatically from sample_data/
    knowledge_base.py          one shared data snapshot used by both build_mbr.py and chatbot_app.py
    chat_engine.py             rule-based question answering (no external AI API)
    roadmap_extractor.py       finds each team's Project Roadmap slide content
    pdf_export.py              converts the finished deck to PDF via LibreOffice

archive/                       superseded scripts and docs, kept for reference only
  old_docs/                    contradictory/outdated planning docs from earlier sessions
  old_scripts/                 COM-automation attempts, earlier broken graph reader, etc.

sample_data/                   example input files (safe to delete/replace)
  Graph/                       LT-SimplifiedMetrics Excel files
  ppt-data/                    Team presentation files (.pptx and .odp)
  final-data/                  MBR template PowerPoint (prior month's finished deck)

output/                        generated files land here:
  master_file.xlsx                       (merger tool output)
  dashboard.html                         (interactive web dashboard)
  mbr_report.xlsx                        (Sprint MBR Excel report)
  MBR_<Month>_<Year>.pptx                 (this month's finished MBR deck)
  MBR_<Month>_<Year>_REVIEW_CHECKLIST.md  (what still needs a human look)
```

## Notes & limitations

- PPTX **tables** only merge cleanly into the Master/All_Records sheets if
  their header row text matches one of your known column names/aliases.
  Tables with unrecognized headers still get included as rows, just with
  whatever column names were on the slide.
- Very large folders (thousands of files) will work but will take longer;
  there's no parallelism built in.
- `.xls` (old Excel format) requires the `xlrd` package — install with
  `pip install xlrd` if you have legacy `.xls` files.
