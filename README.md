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
| **Source Tables** | One row per distinct source table across all reports — the migration worklist — with the reports that use it. |

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
  --drop-no-lineage   ignore root queries that have no upstream lineage
                      (e.g. tables hardcoded inside PowerBI)
  --config FILE.json  override rules / formatting / source labels
  --gui               launch the GUI
```

With no arguments it prompts for an input path interactively.

## Configuration

Defaults live at the top of `data_lineage.py`; anything can be overridden
per-run with `--config file.json` (the GUI's `lineage_settings.json` uses the
same schema):

* **`TRIVIAL_RULES`** — stop expanding "trivial" objects in Summary sheets:
  blocked namespaces (`DB`, `DB.SCHEMA` or `DB.SCHEMA.OBJECT`), and
  level-limits for listed tables/views (`"ALL"` = blanket).
  `apply_rules_to_detailed` extends the rules to Detailed sheets.
* **`EXCEL_FORMAT`** — colors, font, borders, padding, query-block merging,
  hyperlinks, optional freeze panes.
* **`SOURCE_LABELS`** — prefix → label map for the "List of sources" column
  (e.g. `PROD_DATALAKE.CRM_MSCRM` → `CRM`); `SOURCE_LABEL_FALLBACK` picks
  what to show when nothing matches (`schema`, `db` or `blank`).

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
