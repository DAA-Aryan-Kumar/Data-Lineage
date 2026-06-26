# Data Lineage Builder

Turns Atlan **impact report** exports into client-formatted Excel lineage
workbooks that strictly follow the layout of
`RevOps reports - Data Lineage v2.0.xlsx`.

## What it produces

One workbook containing:

| Sheet | Contents |
|---|---|
| **List of Reports** | Index of every report with stats; report names link to their sheets. |
| **\<Report name\>** | Summary lineage in the client layout: compacted tree, purple/pink headers, spacer columns, merged Query blocks, "List of sources" and "Remarks" columns. Every object is expanded **exactly once, at its lowest depth**; later occurrences are blue hyperlinks that jump to the cell where the object is fully expanded. |
| **\<Report name\> (Detailed)** | Same layout, but every branch fully expanded (no de-duplication). |
| **Source Tables** | One row per distinct source table across all reports — the migration worklist — with the reports that use it. Root queries that have no lineage of their own are kept but separated into a labelled section below the main table. |

## Usage

### GUI (recommended for most users)

Double-click `Run Lineage Builder.bat`, or run:

```
python lineage_ui.py
```

Add one or more impact reports, tweak options/rules if needed, press
**Build Lineage Workbook**. Settings persist to `lineage_settings.json`.

### Command line

```
python data_lineage.py -i report1.xlsx [report2.csv ...] [options]

  -o, --output PATH   output workbook (default: <input>__Lineage.xlsx)
  -s, --separate      write Detailed sheets to a separate _Detailed.xlsx
  --no-detailed       skip the Detailed sheets (much faster)
  --unformatted       raw wide dump without client styling
  --drop-no-lineage   ignore root queries that have no upstream lineage, e.g.
                      tables hardcoded inside PowerBI (ON by default;
                      use --no-drop-no-lineage to keep them)
  --config FILE.json  override rules / formatting / source labels
  --gui               launch the GUI
```

With no arguments it prompts for an input path interactively.

### Standalone executable (optional)

`lineage_app.py` is a single entry point that can be packaged into one
Windows `.exe` (no Python needed on the target machine). The same binary
serves both modes: **no arguments → GUI** (it prints a short loading line and
hides its own console once the window is up), **arguments → CLI** (attaches to
the calling terminal). Build it with PyInstaller:

```
pyinstaller --onefile --console --name "Data Lineage Builder" ^
  --icon app.ico --add-data "app.ico;." lineage_app.py
```

(The repo `.gitignore` excludes the generated `.exe`, `build/`, and `dist/`.)
For a watermarked personal build, add `--add-data "watermark.txt;."` where
`watermark.txt` contains the mark text (see below).

**Tip — build a *small* exe.** PyInstaller bundles whatever is in the build
environment. Building from an Anaconda base produces a ~250 MB exe because its
NumPy links Intel MKL. Building from a clean pip venv (NumPy ships the much
smaller OpenBLAS) yields a ~38 MB exe with identical behaviour:

```
py -m venv buildenv
buildenv\Scripts\activate
pip install "pandas==2.2.3" openpyxl lxml pyinstaller   # pin pandas 2.x
pyinstaller --onefile --console --name "Data Lineage Builder" ^
  --icon app.ico --add-data "app.ico;." lineage_app.py
```

The build machine needs Python; the resulting `.exe` does not. (If that venv
is created from a *conda* Python, also put `…\anaconda3\Library\bin` on `PATH`
during the build so PyInstaller can find `ffi-8.dll` for `_ctypes`.)

## Configuration

Defaults live at the top of `data_lineage.py`; anything can be overridden
per-run with `--config file.json` (the GUI's `lineage_settings.json` uses the
same schema):

* **`TRIVIAL_RULES`** — trim noise. Every list holds **glob patterns** matched
  against an object's `DATABASE.SCHEMA.OBJECT` triple; `*` matches within a
  segment and entries with fewer than three segments auto-pad with `.*`
  (`TEMP_DB` → `TEMP_DB.*.*`, `DB.SCHEMA` → `DB.SCHEMA.*`), so `*.DATAWAREHOUSE.*`,
  `PROD_*.*.BRANCH*` all work; `*.*.*` matches everything (logged). Matching is
  **case-insensitive** unless you **wrap the entry in quotes**
  (`"PROD_DATALAKE.LAWPROD.attrep_changes*"`), which matches case-sensitively —
  for Snowflake quoted identifiers. The four sections are: `exclude_patterns`
  (hidden from output, not expanded), `block_patterns` (shown but not
  expanded), and `table_patterns` / `view_patterns` (shown but not expanded
  once *beyond* `table_level` / `view_level`). `drop_no_lineage_summary` /
  `drop_no_lineage_detailed` drop root queries with no upstream, and
  `apply_to_summary` / `apply_to_detailed` choose which sheet set the pattern
  rules touch — each independently per sheet set.
* **`EXCEL_FORMAT`** — colors, font, borders, padding, query-block merging,
  hyperlinks, optional freeze panes.
* **`SOURCE_LABELS`** — `glob pattern` → label map for the "List of sources"
  column, using the same glob syntax as the rules (e.g.
  `*.CRM_MSCRM.*` → `CRM`; quote an entry for case-sensitive matching); the
  **first matching pattern wins**, so ordering sets priority.
  `SOURCE_LABEL_FALLBACK` picks what to show when nothing matches
  (`schema`, `db` or `blank`).
* **Watermark** — a small muted mark in the GUI's top-right corner, **off by
  default**. Turn it on (without forking the code) via any of, in priority
  order: the `LINEAGE_WATERMARK` environment variable, a `watermark.txt`
  bundled into a packaged `.exe`, or a `"watermark"` key in
  `lineage_settings.json`. This is how the personal "AK" build is produced
  from the same source.

The app icon is `app.ico` (vector master `app.svg`); it is used for both the
window title bar and the packaged `.exe`.

## Requirements

Python 3.10+ with `pandas` and `openpyxl` (`pip install pandas openpyxl`).
The GUI uses only the standard library on top of that.

## How it works

1. The impact report is parsed into an upstream adjacency map keyed by
   `Database.Schema.Name` (PowerBI objects use their bare name).
   Root objects are those at the file's minimum Lineage Depth.
2. A depth-first walk builds the lineage tree. Cycles are cut per-path.
   In Summary mode a BFS first computes each object's minimum depth; the
   walk then expands each object only at that depth, the first time it is
   seen — everything else becomes a hyperlinked duplicate. This guarantees
   every reachable object is expanded exactly once.
3. The tree is laid out into a plain Python grid (parents share the row of
   their first child — the client's compacted layout) and bulk-written via
   openpyxl with prebuilt shared styles, which keeps large Detailed sheets
   fast.
