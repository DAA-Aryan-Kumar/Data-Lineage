"""
data_lineage.py
===============
Builds recursive data-lineage workbooks from Atlan "impact report" exports.

Given one or more impact reports (.csv / .xlsx), the script traces every
root object (lowest Lineage Depth, e.g. PowerBI queries) back to its source
tables and writes an Excel workbook that strictly follows the client's
layout ("RevOps reports - Data Lineage v2.0.xlsx"):

  * one sheet per report, compacted tree layout (parents never repeat),
  * spacer columns between level groups, purple/pink headers, Messina Sans,
  * merged Query blocks, "List of sources" and "Remarks" columns,
  * a "List of Reports" index sheet and a "Source Tables" rollup sheet,
  * native internal hyperlinks from every pruned duplicate to the cell
    where that object is fully expanded.

Usage:
    python data_lineage.py -i report1.xlsx [report2.csv ...] [-o out.xlsx]
                           [-s] [--no-detailed] [--unformatted]
                           [--config rules.json] [--gui]
"""

from __future__ import annotations

import argparse
import gc
import json
import logging
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
from collections import deque
from xml.sax.saxutils import escape as _xml_escape, unescape as _xml_unescape

import pandas as pd

log = logging.getLogger("data_lineage")

# --------------------------------------------------------------------------
# CONFIGURATION: PRUNING RULES
# Every list holds glob patterns matched against an object's
# DATABASE.SCHEMA.OBJECT triple. '*' matches within a segment; entries with
# fewer than three segments auto-pad with '.*' (so 'TEMP_DB' -> 'TEMP_DB.*.*'
# and 'DB.SCHEMA' -> 'DB.SCHEMA.*'). '*.*.*' matches everything (logged).
# Matching is case-INSENSITIVE by default; wrap an entry in quotes to make it
# case-SENSITIVE, e.g.  "PROD_DATALAKE.LAWPROD.attrep_changes*"  (Snowflake
# quoted identifiers are case-sensitive). A JSON file passed via --config (or
# the GUI) overrides these.
# --------------------------------------------------------------------------
TRIVIAL_RULES = {
    # Objects that must NOT appear in the output at all (hidden and not expanded).
    'exclude_patterns': [],

    # Objects shown in the output, but whose upstream is NOT expanded.
    'block_patterns': [],

    # Stop expanding matching TABLES once they sit beyond this level.
    'table_patterns': [],
    'table_level': 8,

    # Stop expanding matching VIEWS once they sit beyond this level.
    'view_patterns': [],
    'view_level': 8,

    # Drop root queries that have no upstream lineage, per sheet set.
    'drop_no_lineage_summary': True,
    'drop_no_lineage_detailed': True,

    # Which sheet sets the pattern rules apply to.
    'apply_to_summary': True,
    'apply_to_detailed': False,
}

# --------------------------------------------------------------------------
# CONFIGURATION: EXCEL FORMATTING (matches the client reference workbook)
# --------------------------------------------------------------------------
EXCEL_FORMAT = {
    'font_name': 'Messina Sans',
    'font_size': 10,
    'header_fill': '7030A0',        # purple group-header band
    'header_font_color': 'FFFFFF',
    'subheader_fill': 'F2CEEF',     # pink db/schema/object/type band
    'subheader_font_color': '000000',
    'index_header_fill': 'E5E8EE',  # grey band on the index sheets
    'apply_borders': True,
    'border_color': '000000',
    'border_style': 'thin',
    'hyperlink_duplicates': True,   # link pruned duplicates to their expansion
    'hyperlink_color': '0563C1',
    'pad_empty_cells': True,        # client pads empty cells with NBSP
    'merge_query_blocks': True,     # merge the Query cell over its block
    'freeze_panes': False,          # client sheets are not frozen
    'column_widths': {
        'edge': 5.57, 'query': 39.0, 'db': 22.0, 'schema': 16.0,
        'object': 36.0, 'type': 7.0, 'spacer': 1.7,
        'sources': 19.0, 'remarks': 45.0,
    },
}

# --------------------------------------------------------------------------
# CONFIGURATION: "List of sources" labels
# 'glob pattern' : 'label', using the same DATABASE.SCHEMA.OBJECT glob syntax
# as the pruning rules (quote an entry for case-sensitive matching). The FIRST
# matching pattern (top-to-bottom) wins, so ordering sets priority. Fallback
# when nothing matches: 'schema', 'db' or 'blank'.
# --------------------------------------------------------------------------
SOURCE_LABELS = {
    'PROD_DATALAKE.CRM_MSCRM.*': 'CRM',
    'PROD_DATALAKE.LAWPROD.*': 'LAWSON',
    'DATALAKE.PUBLIC.*': 'LAWSON',
}
SOURCE_LABEL_FALLBACK = 'schema'

NBSP = ' '

# Node statuses -------------------------------------------------------------
ST_EXPANDED = 'expanded'        # children rendered below/right of this node
ST_SOURCE = 'source'            # true source table (no upstream at all)
ST_DUP = 'duplicate'            # pruned: already expanded elsewhere
ST_RULE = 'rule'                # pruned: matched a trivial rule
ST_CYCLE = 'cycle'              # pruned: all children already on this path
ST_NO_UPSTREAM = 'no_upstream'  # root object with no lineage in the report


# --------------------------------------------------------------------------
# GLOB NAMESPACE MATCHING (shared by pruning rules and source labels)
# --------------------------------------------------------------------------
def _clean_field(value) -> str:
    """A metadata field as a plain string, with NaN/None -> ''."""
    if value is None or (isinstance(value, float) and value != value):
        return ''
    return str(value)


def compile_patterns(patterns):
    """Compile glob patterns into [(original, [db_re, schema_re, object_re]), ...].
    Matching is case-insensitive UNLESS the whole entry is wrapped in quotes
    ("..." or '...'), which strips the quotes and matches case-sensitively
    (for Snowflake quoted identifiers). Each (unquoted) pattern is split on '.'
    into at most three segments, padded with '*' to three, then each segment
    glob-compiled. A pattern reducing to '*.*.*' is flagged in the log."""
    compiled = []
    for raw in patterns:
        raw = raw.strip()
        if not raw:
            continue
        if len(raw) >= 2 and raw[0] == raw[-1] and raw[0] in '"\'':
            pat, flags = raw[1:-1], 0              # quoted -> case-sensitive
        else:
            pat, flags = raw, re.IGNORECASE
        parts = pat.split('.')
        if len(parts) < 3:
            parts = parts + ['*'] * (3 - len(parts))
        elif len(parts) > 3:                       # object name containing dots
            parts = parts[:2] + ['.'.join(parts[2:])]
        if all(p.strip() == '*' for p in parts):
            log.warning("Rule '%s' matches every object - applying it anyway.", raw)
        seg_res = [re.compile(re.escape(seg).replace(r'\*', '.*').replace(r'\?', '.') + r'\Z',
                              flags)
                   for seg in parts]
        compiled.append((raw, seg_res))
    return compiled


def match_patterns(fields, compiled):
    """Return the first pattern in `compiled` whose three segment-regexes all
    match the (db, schema, object) `fields`, or None."""
    for pat, seg_res in compiled:
        if all(rx.match(f) for rx, f in zip(seg_res, fields)):
            return pat
    return None


class Node:
    """One occurrence of an object in the lineage tree."""
    __slots__ = ('name', 'level', 'children', 'status', 'remark', 'coord')

    def __init__(self, name: str, level: int):
        self.name = name
        self.level = level
        self.children: list['Node'] = []
        self.status = ST_SOURCE
        self.remark = ''
        self.coord = None  # Excel coordinate of the object cell, set on render

    def leaves(self):
        if not self.children:
            yield self
        else:
            for child in self.children:
                yield from child.leaves()

    def max_level(self) -> int:
        return max((c.max_level() for c in self.children), default=self.level)


def derive_report_name(path: str) -> str:
    """'New Candidate Funnel Dashboard_Upstream.csv' -> 'New Candidate Funnel Dashboard'."""
    stem = os.path.splitext(os.path.basename(path))[0]
    stem = re.sub(r'(?i)(__lineage|_upstream|\s*-\s*copy|\s*\(\d+\))+$', '', stem)
    return stem.strip(' _-') or stem


_CANON_COLS = ('Name', 'Database', 'Schema', 'Type', 'Connector', 'Lineage Depth',
               'Immediate upstream', 'Immediate downstream')


def _canon_cols(df):
    """Rename input columns to canonical names, matching case/whitespace-
    insensitively (e.g. 'Lineage depth', 'Immediate Downstream ')."""
    norm = lambda s: re.sub(r'\s+', ' ', str(s).strip()).lower()
    lookup = {norm(c): c for c in df.columns}        # first occurrence wins
    renames = {lookup[norm(c)]: c for c in _CANON_COLS if norm(c) in lookup}
    return df.rename(columns=renames) if renames else df


def _downstream_names(raw) -> list:
    """Readable names from an Atlan 'Name (default/.../id), Other (...)' value
    (the text before each parenthesised qualified name)."""
    if not isinstance(raw, str):
        return []
    out = []
    for nm in re.findall(r'([^(]+?)\s*\([^)]*\)', raw):
        nm = nm.strip().strip(',').strip()
        if nm:
            out.append(nm)
    return out


def _data_report_name(df, fallback: str) -> str:
    """Derive the report's name from its own data, falling back to `fallback`
    (the filename-based name) when the data is ambiguous:
      * dashboard export -> the PowerBI root rows' shared downstream container
        (e.g. 'ATS Pilot Dashboard'), so a renamed input file still names right;
      * table-level export -> the single root object's Name (e.g. 'BRANCH_DIM')."""
    if 'Lineage Depth' not in df.columns or 'Name' not in df.columns:
        return fallback
    depth = pd.to_numeric(df['Lineage Depth'], errors='coerce')
    if depth.notna().sum() == 0:
        return fallback
    root = depth.eq(depth.min())
    is_pbi = (df['Connector'].astype(str).str.strip().str.lower().eq('powerbi')
              if 'Connector' in df.columns else pd.Series(False, index=df.index))
    if 'Immediate downstream' in df.columns:
        names = []
        for raw in df.loc[root & is_pbi, 'Immediate downstream']:
            names.extend(_downstream_names(raw))
        if names:
            from collections import Counter
            return Counter(names).most_common(1)[0][0]
    distinct = [n for n in dict.fromkeys(df.loc[root, 'Name'].astype(str).map(str.strip))
                if n and n.lower() != 'nan']
    return distinct[0] if len(distinct) == 1 else fallback


def peek_report_name(path: str) -> str:
    """Cheap data-driven name for `path` without building the lineage graph
    (used to assign merge sheet names before the parallel render)."""
    fallback = derive_report_name(path)
    try:
        if path.lower().endswith('.xlsx'):
            df = pd.read_excel(path)
        elif path.lower().endswith('.csv'):
            df = pd.read_csv(path)
        else:
            return fallback
        return _data_report_name(_canon_cols(df), fallback)
    except Exception:
        return fallback


# ===========================================================================
# LINEAGE ENGINE
# ===========================================================================
class DataLineageBuilder:
    """Loads an Atlan impact report and builds detailed / summary lineage trees."""

    REQUIRED_COLS = ['Name', 'Database', 'Schema', 'Type', 'Connector',
                     'Lineage Depth', 'Immediate upstream']

    def __init__(self, input_file_path: str, rules: dict | None = None):
        if not input_file_path:
            raise ValueError("Input file path must be provided.")
        self.input_file_path = os.path.abspath(input_file_path)
        self.report_name = derive_report_name(self.input_file_path)
        self.rules = TRIVIAL_RULES if rules is None else rules
        self.details: dict[str, dict] = {}     # name -> {Database, Schema, Name, Type}
        self.adjacency: dict[str, tuple] = {}  # name -> upstream names, in order
        self.roots: list[str] = []             # objects at the minimum lineage depth
        self._rc_exclude = self._rc_block = self._rc_table = self._rc_view = []

    # -- loading / preprocessing -------------------------------------------
    def load(self):
        if not os.path.exists(self.input_file_path):
            raise FileNotFoundError(f"Input file not found: {self.input_file_path}")
        log.info("Loading %s", os.path.basename(self.input_file_path))
        if self.input_file_path.lower().endswith('.xlsx'):
            df = pd.read_excel(self.input_file_path)
        elif self.input_file_path.lower().endswith('.csv'):
            df = pd.read_csv(self.input_file_path)
        else:
            raise ValueError("Unsupported file format, expected .csv or .xlsx")

        df = _canon_cols(df)
        missing = [c for c in self.REQUIRED_COLS if c not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns in input: {missing}")
        self._preprocess(df)
        return self

    @staticmethod
    def _parse_upstream(raw) -> tuple:
        """Parse Atlan's 'Name (default/connector/id/DB/SCHEMA/NAME)' lists."""
        if not isinstance(raw, str) or not raw.strip():
            return ()
        names = []
        for inner in re.findall(r'\(([^()]*)\)', raw):
            parts = inner.strip().split('/')
            names.append('.'.join(parts[3:6]) if len(parts) >= 6 else parts[-1])
        return tuple(dict.fromkeys(names))  # de-dupe, keep order

    def _preprocess(self, df: pd.DataFrame):
        # Match the connector case-insensitively: Atlan exports vary between
        # 'powerbi' and 'Powerbi'. PowerBI objects have no Database/Schema, so
        # they use their bare Name; everything else is Database.Schema.Name.
        is_pbi = df['Connector'].astype(str).str.strip().str.lower().eq('powerbi')
        full = (df['Database'].astype(str) + '.' + df['Schema'].astype(str)
                + '.' + df['Name'].astype(str))
        df = df.assign(**{'Consolidated Name': full.where(~is_pbi, df['Name'])})

        first = df.drop_duplicates(subset=['Consolidated Name'])
        self.details = (first.set_index('Consolidated Name')
                        [['Database', 'Schema', 'Name', 'Type']].to_dict('index'))

        self.adjacency = {
            cn: ups for cn, ups in zip(first['Consolidated Name'],
                                       first['Immediate upstream'].map(self._parse_upstream))
            if ups
        }

        min_depth = df['Lineage Depth'].min()
        self.roots = first.loc[first['Lineage Depth'] == min_depth,
                               'Consolidated Name'].tolist()
        # Name the report from its own data (dashboard downstream / root object),
        # keeping the filename-derived name only as a fallback.
        self.report_name = _data_report_name(df, self.report_name)
        log.info("  %d objects, %d roots (depth %s), %d with upstream",
                 len(first), len(self.roots), min_depth, len(self.adjacency))

    def object_meta(self, name: str) -> dict:
        """Metadata for a node; falls back to parsing the consolidated name."""
        meta = self.details.get(name)
        if meta:
            return meta
        parts = name.split('.')
        if len(parts) >= 3:
            return {'Database': parts[0], 'Schema': parts[1],
                    'Name': '.'.join(parts[2:]), 'Type': ''}
        return {'Database': '', 'Schema': '', 'Name': name, 'Type': ''}

    # -- pruning rules -------------------------------------------------------
    def _compile_rules(self):
        """Compile each rule section's glob patterns once per build."""
        r = self.rules
        self._rc_exclude = compile_patterns(r.get('exclude_patterns', []))
        self._rc_block = compile_patterns(r.get('block_patterns', []))
        self._rc_table = compile_patterns(r.get('table_patterns', []))
        self._rc_view = compile_patterns(r.get('view_patterns', []))

    def _rule_action(self, name: str, level: int) -> tuple[str, str | None]:
        """Decide what happens to a node: 'exclude' (hide it entirely),
        'block'/'stop' (show it but do not expand its upstream), or 'expand'.
        Returns (action, remark)."""
        meta = self.object_meta(name)
        fields = (_clean_field(meta.get('Database')),
                  _clean_field(meta.get('Schema')),
                  _clean_field(meta.get('Name')))
        if match_patterns(fields, self._rc_exclude):
            return 'exclude', None
        if match_patterns(fields, self._rc_block):
            return 'block', 'Not expanded: matched a block rule'
        obj_type = str(meta.get('Type') or '').upper()
        if 'TABLE' in obj_type and self._rc_table:
            limit = self.rules.get('table_level', 99)
            if level > limit and match_patterns(fields, self._rc_table):
                return 'stop', f'Not expanded: table beyond level L{limit}'
        if 'VIEW' in obj_type and self._rc_view:
            limit = self.rules.get('view_level', 99)
            if level > limit and match_patterns(fields, self._rc_view):
                return 'stop', f'Not expanded: view beyond level L{limit}'
        return 'expand', None

    # -- tree building -------------------------------------------------------
    def _min_depths(self, apply_rules: bool) -> dict[str, int]:
        """BFS shortest depth per reachable node, honouring pruning rules.
        Excluded objects are unreachable; blocked/stopped objects appear but
        their upstream is not traversed."""
        depth = {}
        queue = deque()
        for r in self.roots:
            if apply_rules and self._rule_action(r, 0)[0] == 'exclude':
                continue
            depth[r] = 0
            queue.append(r)
        while queue:
            name = queue.popleft()
            if apply_rules and self._rule_action(name, depth[name])[0] != 'expand':
                continue
            for child in self.adjacency.get(name, ()):
                clevel = depth[name] + 1
                if apply_rules and self._rule_action(child, clevel)[0] == 'exclude':
                    continue
                if child not in depth:
                    depth[child] = clevel
                    queue.append(child)
        return depth

    def build_tree(self, summarize: bool,
                   drop_no_lineage: bool = True) -> tuple[list[Node], dict[str, Node]]:
        """
        Returns (root nodes, expanded_at). In summary mode every object is
        expanded exactly once, at its minimum depth; later occurrences become
        ST_DUP leaves pointing back at the expansion (via expanded_at).

        If drop_no_lineage is True, root objects with no upstream at all
        (e.g. tables hard-coded inside PowerBI) are omitted entirely rather
        than emitted as a lone Query row.
        """
        sys.setrecursionlimit(max(sys.getrecursionlimit(), 50_000))
        self._compile_rules()
        apply_rules = (self.rules.get('apply_to_summary', True) if summarize
                       else self.rules.get('apply_to_detailed', False))
        min_depth = self._min_depths(apply_rules) if summarize else {}
        expanded_at: dict[str, Node] = {}

        def make(name: str, level: int, path: frozenset) -> Node | None:
            action, reason = (self._rule_action(name, level) if apply_rules
                              else ('expand', None))
            if action == 'exclude':
                return None
            node = Node(name, level)
            children = self.adjacency.get(name, ())
            if not children:
                node.status = ST_NO_UPSTREAM if level == 0 else ST_SOURCE
                if level == 0:
                    node.remark = 'No upstream lineage found in the impact report'
                return node
            if action in ('block', 'stop'):
                node.status, node.remark = ST_RULE, reason
                return node
            if summarize and (name in expanded_at or level != min_depth.get(name, level)):
                node.status = ST_DUP
                return node
            kids = [c for c in children if c not in path]
            if not kids:
                node.status, node.remark = ST_CYCLE, 'Circular reference'
                return node
            node.status = ST_EXPANDED
            if summarize:
                expanded_at[name] = node
            child_path = path | {name}
            node.children = [n for n in (make(c, level + 1, child_path) for c in kids)
                             if n is not None]
            return node

        root_names = self.roots
        if drop_no_lineage:
            root_names = [r for r in self.roots if self.adjacency.get(r)]
            dropped = len(self.roots) - len(root_names)
            if dropped:
                log.info("  Ignoring %d root object(s) with no upstream lineage", dropped)

        roots = [n for n in (make(r, 0, frozenset()) for r in root_names) if n is not None]
        mode = 'summary' if summarize else 'detailed'
        log.info("  Built %s tree: %d rows, max depth L%d", mode,
                 sum(1 for r in roots for _ in r.leaves()),
                 max((r.max_level() for r in roots), default=0))
        return roots, expanded_at

    def source_tables(self) -> list[str]:
        """True source tables — reached objects (depth > 0) with no upstream.
        Excludes root queries that simply have no lineage of their own; those
        are reported separately by rootless_queries()."""
        reachable = self._min_depths(apply_rules=False)
        return [n for n, depth in reachable.items()
                if depth > 0 and not self.adjacency.get(n)]

    def rootless_queries(self) -> list[str]:
        """Root objects with no upstream lineage at all (e.g. tables hardcoded
        inside PowerBI). These are dropped from the lineage sheets by default
        but still listed (separately) in the Source Tables rollup."""
        return [r for r in self.roots if not self.adjacency.get(r)]


# ===========================================================================
# EXCEL RENDERER (strict client layout)
# ===========================================================================
class ClientExcelRenderer:
    DATA_START_ROW = 4   # row 1 blank, row 2 group header, row 3 sub header

    def __init__(self, fmt: dict | None = None, source_labels: dict | None = None,
                 source_fallback: str | None = None):
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

        self.fmt = EXCEL_FORMAT if fmt is None else fmt
        labels = SOURCE_LABELS if source_labels is None else source_labels
        self.source_fallback = (SOURCE_LABEL_FALLBACK if source_fallback is None
                                else source_fallback)
        # Compile labels to (label, seg_res) preserving order (first match wins).
        self._source_labels = []
        for pat, label in labels.items():
            comp = compile_patterns([pat])
            if comp:
                self._source_labels.append((label, comp[0][1]))
        f = self.fmt
        self.font_data = Font(name=f['font_name'], size=f['font_size'])
        self.font_data_bold = Font(name=f['font_name'], size=f['font_size'], bold=True)
        self.font_header = Font(name=f['font_name'], size=f['font_size'], bold=True,
                                color=f['header_font_color'])
        self.font_subheader = Font(name=f['font_name'], size=f['font_size'], bold=True,
                                   color=f['subheader_font_color'])
        self.font_link = Font(name=f['font_name'], size=f['font_size'],
                              color=f['hyperlink_color'], underline='single')
        self.fill_header = PatternFill('solid', start_color=f['header_fill'])
        self.fill_subheader = PatternFill('solid', start_color=f['subheader_fill'])
        self.fill_white = PatternFill('solid', start_color='FFFFFF')
        self.fill_index = PatternFill('solid', start_color=f['index_header_fill'])
        side = Side(border_style=f['border_style'], color=f['border_color'])
        self.border = (Border(left=side, right=side, top=side, bottom=side)
                       if f['apply_borders'] else None)
        self.align_center = Alignment(horizontal='center', vertical='center',
                                      wrap_text=True)

    SEED_SHEET = '__seed__'        # hidden style-seed lineage sheet
    SEED_GRID = '__seed_grid__'    # hidden style-seed index/source grid

    def seed_styles(self, wb):
        """Render a tiny example through the REAL renderer onto hidden sheets so
        the workbook's style table is complete and byte-identical across every
        report and the index/source aux. openpyxl prunes styles that no live
        cell references, and font_link / fill_index are used only conditionally,
        so without this a no-dups report or the aux would get a different
        styles.xml and the raw XML-splice merge would mis-map style indices.
        Driving it through render_lineage_sheet + _write_grid (rather than a
        hand-listed palette) keeps it automatically in sync with the renderer.
        The seed sheets are hidden and never copied into the merged output."""
        class _StubBuilder:
            def object_meta(self, name):
                p = name.split('.')
                if len(p) >= 3:
                    return {'Database': p[0], 'Schema': p[1],
                            'Name': '.'.join(p[2:]), 'Type': 'View'}
                return {'Database': '', 'Schema': '', 'Name': name, 'Type': ''}

        # The seed must exercise the FULL Query-column merge geometry, or the
        # per-report files and the aux end up with different style tables and
        # the XML splice mis-maps indices. openpyxl's MergedCellRange.format()
        # (run at save) strips a vertical merge's interior cells down to a
        # left+right-only border -- but only when the block is >=3 rows tall,
        # so the "covered cell" style is born only from a tall block. A 2-row
        # seed (the old one) never created it, so it drifted to a divergent
        # cellXfs slot and covered Query cells rendered bold + no side border.
        # Root 1: a tall (>=3 row) merged block -> top / interior / bottom
        # border variants. Root 2: a 1-row block -> the unmerged anchor style
        # (bold + centered + full border), for single-row queries / merge off.
        root = Node('Q', 0); root.status = ST_EXPANDED
        view = Node('DB.SCH.VIEW', 1); view.status = ST_EXPANDED
        tbl = Node('DB.SCH.TBL', 2); tbl.status = ST_SOURCE
        view.children = [tbl]                                # block row 1
        extra = Node('DB.SCH.EXTRA', 1); extra.status = ST_SOURCE  # row 2 (interior)
        dup = Node('DB.SCH.VIEW', 1); dup.status = ST_DUP    # row 3; forces font_link
        root.children = [view, extra, dup]                   # 3 rows -> tall merge
        solo = Node('Q2', 0); solo.status = ST_EXPANDED
        solo_src = Node('DB.SCH.SOLO', 1); solo_src.status = ST_SOURCE
        solo.children = [solo_src]                           # 1 row -> unmerged
        self.render_lineage_sheet(wb, self.SEED_SHEET, _StubBuilder(),
                                  [root, solo], {'DB.SCH.VIEW': view}, link_dups=True)
        grid = wb.create_sheet(self.SEED_GRID)
        self._write_grid(grid, ['a', 'b'], [['x', 'y']], [10, 10])
        # The Source Tables "rootless" section label is bold with NO border
        # (see _write_source_sheet); seed that exact combination -- with a
        # border it seeds the wrong style and the aux's styles.xml diverges.
        label = grid.cell(row=10, column=2, value='r')
        label.font = self.font_data_bold
        for name in (self.SEED_SHEET, self.SEED_GRID):
            wb[name].sheet_state = 'hidden'

    # -- helpers -------------------------------------------------------------
    def source_label(self, name: str, meta: dict) -> str:
        fields = (_clean_field(meta.get('Database')),
                  _clean_field(meta.get('Schema')),
                  _clean_field(meta.get('Name')))
        for label, seg_res in self._source_labels:        # first match wins
            if all(rx.match(f) for rx, f in zip(seg_res, fields)):
                return label
        if self.source_fallback == 'schema':
            return _clean_field(meta.get('Schema'))
        if self.source_fallback == 'db':
            return _clean_field(meta.get('Database'))
        return ''

    # Low-information words dropped first when abbreviating a long sheet name.
    _SHEET_FILLER = ('Dashboard', 'Report', 'Analysis', 'Analytics', 'Activity',
                     'Program', 'Center', 'Data')

    @staticmethod
    def safe_sheet_name(name: str, taken: set, suffix: str = '') -> str:
        """Fit a report name into Excel's 31-char sheet-name limit by smart
        abbreviation: strip illegal chars, drop filler words, then shorten
        words, before any hard truncation. Ensures uniqueness within `taken`."""
        clean = re.sub(r"[\[\]:*?/\\']", ' ', str(name))
        clean = re.sub(r'\s+', ' ', clean).strip()
        budget = 31 - len(suffix)

        candidate = clean
        if len(candidate) > budget:                       # 1. drop filler words
            words = candidate.split(' ')
            for fw in ClientExcelRenderer._SHEET_FILLER:
                if len(' '.join(words)) <= budget:
                    break
                trimmed = [w for w in words if w.lower() != fw.lower()]
                words = trimmed or words
            candidate = ' '.join(words)
        if len(candidate) > budget:                       # 2. shorten multi-word names
            words = candidate.split(' ')
            if len(words) > 1:                            # a single token (e.g. a snowflake
                candidate = ' '.join(w[:4] for w in words)  # table VW_...) keeps its prefix
        if len(candidate) > budget:                       # 3. last-resort truncate
            candidate = candidate[:budget].rstrip(' _-')

        candidate = (candidate + suffix).strip()
        base, n = candidate, 2
        while candidate.lower() in taken:                 # 4. ensure uniqueness
            tail = f' ({n})'
            candidate = base[:31 - len(tail)].rstrip(' _-') + tail
            n += 1
        taken.add(candidate.lower())
        return candidate

    # -- lineage sheet ---------------------------------------------------------
    def render_lineage_sheet(self, wb, title: str, builder: DataLineageBuilder,
                             roots: list[Node], expanded_at: dict[str, Node],
                             link_dups: bool) -> dict:
        from openpyxl.utils import get_column_letter
        from openpyxl.worksheet.hyperlink import Hyperlink

        ws = wb.create_sheet(title)
        max_depth = max([1] + [r.max_level() for r in roots])

        col_query = 2
        group_col = lambda lvl: 3 + 5 * (lvl - 1)          # db col of level group
        col_sources = 3 + 5 * max_depth
        col_remarks = col_sources + 1
        last_col = col_remarks

        # ---- lay the tree out into a plain Python grid first (fast), then
        # bulk-append to the sheet; cell-by-cell writes are far too slow ----
        pad = NBSP if self.fmt.get('pad_empty_cells', True) else None

        def new_list() -> list:
            # list index i maps to sheet column i+1; column A stays empty
            values = [pad] * last_col
            values[0] = None
            return values

        header2 = new_list()
        header3 = new_list()
        header2[col_query - 1] = 'Query'
        for lvl in range(1, max_depth + 1):
            c = group_col(lvl) - 1
            header2[c] = 'Sources (L1)' if lvl == 1 else f'Underlying sources (L{lvl})'
            header3[c:c + 4] = ('db', 'schema', 'object', 'type')
        header2[col_sources - 1] = 'List of sources'
        header2[col_remarks - 1] = 'Remarks'

        grid: list[list] = []      # data rows only (sheet rows 4..)
        dup_nodes: list[Node] = []
        query_blocks: list[tuple[int, int]] = []

        def new_row() -> list:
            grid.append(new_list())
            return grid[-1]

        def write_node(node: Node, row: list) -> list:
            sheet_row = len(grid) + self.DATA_START_ROW - 1
            if node.level == 0:
                row[col_query - 1] = node.name
                node.coord = f'{get_column_letter(col_query)}{sheet_row}'
            else:
                meta = builder.object_meta(node.name)
                c = group_col(node.level) - 1
                row[c:c + 4] = (meta['Database'], meta['Schema'],
                                meta['Name'], meta['Type'])
                node.coord = f'{get_column_letter(c + 3)}{sheet_row}'
            if node.children:
                for i, child in enumerate(node.children):
                    row = write_node(child, row if i == 0 else new_row())
                return row
            # leaf row: list-of-sources + remarks
            if node.level > 0:
                row[col_sources - 1] = self.source_label(
                    node.name, builder.object_meta(node.name))
            if node.remark:
                row[col_remarks - 1] = node.remark
            if node.status == ST_DUP:
                dup_nodes.append(node)
            return row

        for root in roots:
            start = len(grid) + self.DATA_START_ROW
            write_node(root, new_row())
            query_blocks.append((start, len(grid) + self.DATA_START_ROW - 1))
        last_row = len(grid) + self.DATA_START_ROW - 1

        ws.append([])              # row 1 stays blank (client leaves an edge)
        ws.append(header2)
        ws.append(header3)
        for row in grid:
            ws.append(row)

        # ---- style pass: assign a prebuilt style array (font + border) to
        # every cell. Copying the array is ~10x faster than setting
        # .font/.border individually on two million cells. Each cell needs
        # its OWN copy: openpyxl style setters mutate the array in place,
        # so sharing one instance would bleed later changes everywhere. ----
        from copy import copy as _copy
        probe = ws.cell(row=1, column=1)
        original = probe._style
        probe.font = self.font_data
        if self.border:
            probe.border = self.border
        base_style = probe._style
        probe._style = original
        for cells in ws.iter_rows(min_row=2, max_row=last_row,
                                  min_col=2, max_col=last_col):
            for cell in cells:
                cell._style = _copy(base_style)

        # ---- header styling & merges ----
        spacer_cols = {group_col(l) + 4 for l in range(1, max_depth + 1)}
        for r in (2, 3):
            for c in range(2, last_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.alignment = self.align_center
                if c in spacer_cols:
                    cell.fill = self.fill_white
                    cell.font = self.font_header if r == 2 else self.font_subheader
                elif r == 2:
                    cell.fill, cell.font = self.fill_header, self.font_header
                else:
                    cell.fill, cell.font = self.fill_subheader, self.font_subheader
        ws.merge_cells(start_row=2, start_column=col_query, end_row=3, end_column=col_query)
        for lvl in range(1, max_depth + 1):
            c = group_col(lvl)
            ws.merge_cells(start_row=2, start_column=c, end_row=2, end_column=c + 3)
        for c in (col_sources, col_remarks):
            ws.merge_cells(start_row=2, start_column=c, end_row=3, end_column=c)

        # ---- query blocks: bold + centered, optionally merged over the block ----
        merge = self.fmt.get('merge_query_blocks', True)
        for start, end in query_blocks:
            if merge and end > start:
                ws.merge_cells(start_row=start, start_column=col_query,
                               end_row=end, end_column=col_query)
            top = ws.cell(row=start, column=col_query)
            top.alignment = self.align_center
            top.font = self.font_data_bold

        # ---- hyperlinks: pruned duplicates -> expansion location ----
        n_links = 0
        if link_dups and self.fmt.get('hyperlink_duplicates', True):
            for node in dup_nodes:
                target = expanded_at.get(node.name)
                if target is None or target.coord is None:
                    continue
                cell = ws[node.coord]
                cell.hyperlink = Hyperlink(ref=node.coord,
                                           location=f"'{title}'!{target.coord}",
                                           tooltip=f'Expanded at {target.coord}')
                cell.font = self.font_link
                n_links += 1

        # ---- column widths ----
        w = self.fmt['column_widths']
        ws.column_dimensions['A'].width = w['edge']
        # Autofit the Query column to its (bold) contents. openpyxl has no real
        # autofit, so approximate: widest query name + a little padding, with
        # a 1.12 factor for the bold face, clamped to a sane range.
        header_chars = len('Query')
        content_chars = max((len(str(r.name)) for r in roots), default=header_chars)
        query_width = min(max(content_chars * 1.12 + 2, w['query'] * 0.4), 90)
        ws.column_dimensions[get_column_letter(col_query)].width = query_width
        for lvl in range(1, max_depth + 1):
            c = group_col(lvl)
            for off, key in enumerate(('db', 'schema', 'object', 'type')):
                ws.column_dimensions[get_column_letter(c + off)].width = w[key]
            ws.column_dimensions[get_column_letter(c + 4)].width = w['spacer']
        ws.column_dimensions[get_column_letter(col_sources)].width = w['sources']
        ws.column_dimensions[get_column_letter(col_remarks)].width = w['remarks']

        if self.fmt.get('freeze_panes'):
            ws.freeze_panes = ws.cell(row=self.DATA_START_ROW, column=3)

        return {'sheet': title, 'rows': last_row - self.DATA_START_ROW + 1,
                'max_depth': max_depth, 'queries': len(roots), 'links': n_links}

    # -- simple styled grid (index / rollup sheets) ----------------------------
    def _write_grid(self, ws, header: list[str], rows: list[list],
                    widths: list[float], start_row: int = 2) -> int:
        """Write a header + data block starting at `start_row`. Returns the last
        row written (so callers can stack a second section below)."""
        from openpyxl.utils import get_column_letter
        ws.column_dimensions['A'].width = self.fmt['column_widths']['edge']
        for j, (text, width) in enumerate(zip(header, widths)):
            cell = ws.cell(row=start_row, column=2 + j, value=text)
            cell.fill = self.fill_index
            cell.font = self.font_data
            if self.border:
                cell.border = self.border
            ws.column_dimensions[get_column_letter(2 + j)].width = width
        for i, row in enumerate(rows):
            for j, value in enumerate(row):
                cell = ws.cell(row=start_row + 1 + i, column=2 + j, value=value)
                cell.font = self.font_data
                if self.border:
                    cell.border = self.border
        return start_row + len(rows)  # last row written (header if no rows)


# ===========================================================================
# PARALLEL RENDERING (one process per report)
# ===========================================================================
class BuildCancelled(Exception):
    """Raised when a caller's cancel signal (a threading.Event) is set mid-build,
    so the run unwinds cleanly (temp dirs removed, pool workers terminated)."""


def _resolve_workers(max_workers, n_tasks: int) -> int:
    """Resolve the worker count, clamped to [1, n_tasks]. The default ("auto",
    when max_workers is unset) is one worker per report, capped at the CPU count
    -- more processes than cores can't speed up the CPU-bound rendering and only
    multiplies peak memory. An explicit max_workers acts as an upper cap."""
    if not max_workers or max_workers < 1:
        max_workers = os.cpu_count() or 2
    return max(1, min(int(max_workers), max(1, n_tasks)))


def _source_rows(builder: 'DataLineageBuilder', renderer: 'ClientExcelRenderer', names) -> list[list]:
    """[db, schema, object, type, label, consolidated-name] per source object."""
    rows = []
    for n in names:
        m = builder.object_meta(n)
        rows.append([m['Database'], m['Schema'], m['Name'], m['Type'],
                     renderer.source_label(n, m), n])
    return rows


class _ListLogHandler(logging.Handler):
    """Collects (level, message) records into a list so a child process can ship
    its per-report log lines back to the parent (child logs are otherwise lost
    under the Windows 'spawn' start method)."""

    def __init__(self, sink: list):
        super().__init__()
        self.sink = sink

    def emit(self, record):
        self.sink.append((record.levelno, self.format(record)))


def _render_report_task(task: dict) -> dict:
    """Worker (runs in a child process): build one report's lineage trees and
    render them to a standalone .xlsx, returning picklable metadata + the source
    rows the parent needs, plus the report's own log lines (captured here since
    a spawned child shares no handler with the parent)."""
    import openpyxl
    captured: list = []
    handler = _ListLogHandler(captured)
    handler.setFormatter(logging.Formatter('%(message)s'))
    log.addHandler(handler)
    log.setLevel(logging.INFO)
    try:
        builder = DataLineageBuilder(task['path'], rules=task['rules']).load()
        drop_s, drop_d = task['drop_summary'], task['drop_detailed']
        s_roots, s_exp = builder.build_tree(summarize=True, drop_no_lineage=drop_s)
        d_roots = d_exp = None
        if task['include_detailed']:
            d_roots, d_exp = builder.build_tree(summarize=False, drop_no_lineage=drop_d)

        renderer = ClientExcelRenderer(task['fmt'], task['source_labels'], task['source_fallback'])
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        if task.get('seed'):                       # for merge: identical styles.xml across files
            renderer.seed_styles(wb)
        s_info = renderer.render_lineage_sheet(wb, task['summary_sheet'], builder,
                                               s_roots, s_exp, link_dups=True)
        d_info = None
        if d_roots is not None:
            d_info = renderer.render_lineage_sheet(wb, task['detailed_sheet'], builder,
                                                   d_roots, d_exp, link_dups=False)
        if task['include_source']:
            _add_source_rollup(wb, renderer, [{'builder': builder}])
        out = _safe_save(wb, task['out_path'])
        result = {
            'report_name': builder.report_name, 'out_path': out,
            'input_basename': os.path.basename(builder.input_file_path),
            'summary_sheet': task['summary_sheet'], 'detailed_sheet': task['detailed_sheet'],
            'queries': s_info['queries'], 'max_depth': s_info['max_depth'],
            'summary_rows': s_info['rows'], 'detailed_rows': (d_info['rows'] if d_info else '-'),
            'source_count': len(builder.source_tables()),
            'source_rows': _source_rows(builder, renderer, builder.source_tables()),
            'rootless_rows': _source_rows(builder, renderer, builder.rootless_queries()),
        }
        # release the report's heavy objects before the worker takes the next
        # task (matters when a worker is reused: more reports than workers)
        del wb, builder, s_roots, s_exp, d_roots, d_exp
        gc.collect()
    finally:
        log.removeHandler(handler)
    result['logs'] = captured
    return result


def _run_report_tasks(tasks: list[dict], workers: int, report_progress,
                      errors: list | None = None, stop_on_error: bool = False,
                      cancel=None) -> list[dict]:
    """Render all report tasks, sequentially (workers<=1) or across a process
    pool. Per-report progress is logged as each one finishes. A task that fails
    (e.g. a malformed input file) is logged and skipped (its slot stays None)
    so the other reports still build; the failure is recorded in `errors`.
    With `stop_on_error`, the first failure aborts the whole run instead. If
    `cancel` (a threading.Event) is set, the run stops promptly (pool workers
    are terminated) and raises BuildCancelled."""
    n = len(tasks)
    results: list = [None] * n

    def _cancelled():
        return cancel is not None and cancel.is_set()

    def _failed(task, exc, done):
        if stop_on_error:
            raise RuntimeError(f"{task['name']}: {exc}") from exc
        # one consistently-indented warning line (was a 0-indent error + a
        # duplicate 2-indent line); WARNING so the GUI colours it
        log.warning("  [%d/%d] skipped %s: %s", done, n, task['name'], exc)
        if errors is not None:
            errors.append((task['name'], str(exc)))

    def _record(fut, i, task, done):
        try:
            results[i] = fut.result()
            report_progress(f"  [{done}/{n}] rendered {results[i]['report_name']}")
            # replay the child's captured per-report logs (they didn't stream
            # live from the worker process); levels are preserved so the GUI
            # colours them and the CLI prefixes them.
            for lvl, msg in results[i].get('logs', ()):
                log.log(lvl, "      %s", msg)
        except Exception as exc:
            _failed(task, exc, done)

    if workers <= 1:
        for i, task in enumerate(tasks):
            if _cancelled():
                raise BuildCancelled()
            try:
                results[i] = _render_report_task(task)
                report_progress(f"  [{i + 1}/{n}] rendered {results[i]['report_name']}")
                for lvl, msg in results[i].get('logs', ()):
                    log.log(lvl, "      %s", msg)
            except BuildCancelled:
                raise
            except Exception as exc:
                _failed(task, exc, i + 1)
        return results

    from concurrent.futures import ProcessPoolExecutor, wait, FIRST_COMPLETED
    # Recycle each worker after one report (max_tasks_per_child=1, Python 3.11+)
    # so its peak memory is returned to the OS immediately rather than the worker
    # sitting at its high-water mark until the whole batch finishes. The one-file
    # bundle re-extraction this implies is fast, so it's a good trade.
    pool_kw = {'max_workers': workers}
    if sys.version_info >= (3, 11):
        pool_kw['max_tasks_per_child'] = 1
    pool = ProcessPoolExecutor(**pool_kw)
    futs = {}
    try:
        futs = {pool.submit(_render_report_task, t): (i, t) for i, t in enumerate(tasks)}
        pending, done = set(futs), 0
        while pending:
            if _cancelled():
                raise BuildCancelled()
            # poll so a cancel is noticed within ~0.3s even mid-render
            finished, pending = wait(pending, timeout=0.3, return_when=FIRST_COMPLETED)
            for fut in finished:
                i, task = futs[fut]
                done += 1
                _record(fut, i, task, done)
    except BaseException:
        # cancel queued work and terminate any still-running workers so a Stop
        # (or a stop_on_error abort) doesn't leave child processes running
        for f in futs:
            f.cancel()
        for proc in list(getattr(pool, '_processes', {}).values()):
            try:
                proc.terminate()
            except Exception:
                pass
        raise
    finally:
        pool.shutdown(wait=False)
    return results


def _make_report_tasks(input_files, eff_rules, fmt, source_labels, source_fallback,
                       include_detailed, include_source, out_paths, seed=False,
                       defer_naming=False) -> list[dict]:
    """Build per-report task dicts.

    defer_naming=True (the merge path): do NOT read files here -- the workers do
    all the loading in parallel. Each report renders under fixed 'Summary' /
    'Detailed' sheet names and the parent assigns the final, unique, data-driven
    sheet names at splice time. defer_naming=False (separate-files): peek each
    file so its output is named after its report and sheet names are unique."""
    drop_s = eff_rules.get('drop_no_lineage_summary', True)
    drop_d = eff_rules.get('drop_no_lineage_detailed', True)
    taken = {'list of reports', 'source tables'}
    tasks = []
    for path, out_path in zip(input_files, out_paths):
        if defer_naming:
            name = derive_report_name(path)      # label for errors only (no file read)
            s_sheet = 'Summary'
            d_sheet = 'Detailed' if include_detailed else None
        else:
            name = peek_report_name(path)        # data-driven, matches the child's name
            s_sheet = ClientExcelRenderer.safe_sheet_name(name, taken)
            d_sheet = (ClientExcelRenderer.safe_sheet_name(name, taken, suffix=' (Detailed)')
                       if include_detailed else None)
        tasks.append({'path': path, 'name': name, 'rules': eff_rules, 'fmt': fmt,
                      'source_labels': source_labels, 'source_fallback': source_fallback,
                      'include_detailed': include_detailed, 'include_source': include_source,
                      'seed': seed, 'drop_summary': drop_s, 'drop_detailed': drop_d,
                      'summary_sheet': s_sheet, 'detailed_sheet': d_sheet, 'out_path': out_path})
    return tasks


def _build_separate_files(input_files, output_path, eff_rules, fmt, source_labels,
                          source_fallback, include_detailed, max_workers, report_progress,
                          errors=None, stop_on_error=False, cancel=None, avoid_overwrite=False):
    """One standalone workbook per report (Summary + Detailed + Source Tables),
    rendered in parallel. No index sheet (it's meaningless for single files)."""
    out_paths = _separate_out_paths(input_files, output_path)
    if avoid_overwrite:
        out_paths = [_avoid_existing(p) for p in out_paths]
    tasks = _make_report_tasks(input_files, eff_rules, fmt, source_labels, source_fallback,
                               include_detailed, include_source=True, out_paths=out_paths)
    workers = _resolve_workers(max_workers, len(tasks))
    report_progress(f"Rendering {len(tasks)} report(s) to separate files "
                    f"with {workers} worker process(es)...")
    results = _run_report_tasks(tasks, workers, report_progress, errors, stop_on_error, cancel)
    return [r['out_path'] for r in results if r]


# ===========================================================================
# XML-LEVEL MERGE (cheap: splice pre-rendered worksheets into one workbook)
# ===========================================================================
_NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
_WS_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'


class _StyleMergeMismatch(Exception):
    """A per-report file's styles.xml diverged from the aux template, so its
    verbatim `s=` style indices can't be safely spliced. Signals build_workbooks
    to fall back to the (correct, slower) serial combine instead of emitting a
    workbook with mis-mapped cell styles."""


def _sheet_xml_by_name(zf: zipfile.ZipFile, name: str) -> bytes:
    """Return the worksheet XML for sheet `name` from an .xlsx zip, resolved via
    workbook.xml -> rels (robust to sheet ordering / hidden seed sheets)."""
    wbxml = zf.read('xl/workbook.xml').decode('utf-8')
    rels = zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid = None
    for tag in re.findall(r'<sheet\b[^>]*?/?>', wbxml):
        nm = re.search(r'\bname="([^"]*)"', tag)
        ri = re.search(r'\br:id="([^"]*)"', tag)
        if nm and ri and _xml_unescape(nm.group(1)) == name:
            rid = ri.group(1)
            break
    if rid is None:
        raise KeyError(f"sheet {name!r} not found")
    target = None
    for tag in re.findall(r'<Relationship\b[^>]*?/?>', rels):
        if re.search(r'\bId="' + re.escape(rid) + r'"', tag):
            target = re.search(r'\bTarget="([^"]*)"', tag).group(1)
            break
    target = target.lstrip('/')
    if not target.startswith('xl/'):
        target = 'xl/' + target
    return zf.read(target)


def _rename_summary_refs(xml: bytes, new_name: str) -> bytes:
    """A merge-path worker renders its summary sheet under the fixed name
    'Summary', so its intra-sheet dup->expansion hyperlinks point at
    location="'Summary'!...". Rewrite those to the final sheet name so the links
    stay valid once the sheet is renamed in the merged workbook. (The detailed
    sheet has link_dups=False, so it has no such refs.)"""
    new = ("location=\"'%s'!" % _xml_escape(new_name)).encode('utf-8')
    return xml.replace(b"location=\"'Summary'!", new)


def _safe_replace(tmp_path: str, target: str) -> str:
    """Move tmp_path onto target; if target is locked (open in Excel), fall back
    to a numbered name."""
    try:
        os.replace(tmp_path, target)
        return target
    except PermissionError:
        base, ext = os.path.splitext(target)
        for n in range(1, 100):
            alt = f'{base} ({n}){ext}'
            try:
                os.replace(tmp_path, alt)
                log.warning("'%s' is open elsewhere - saved as '%s' instead.",
                            os.path.basename(target), os.path.basename(alt))
                return alt
            except PermissionError:
                continue
        raise


def _build_aux(results, fmt, source_labels, source_fallback, include_detailed) -> str:
    """Build a (seeded) workbook holding just the combined index + merged Source
    Tables, save it to a temp file, and return the path. Its styles.xml matches
    the per-report files (seed_styles), so its sheet XML can be spliced in."""
    import openpyxl
    renderer = ClientExcelRenderer(fmt, source_labels, source_fallback)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    renderer.seed_styles(wb)

    index_ws = wb.create_sheet('List of Reports')
    index_rows = [[r['report_name'], r['queries'], r['source_count'], f"L{r['max_depth']}",
                   r['summary_rows'], r['detailed_rows'], r['input_basename']] for r in results]
    renderer._write_grid(index_ws,
                         ['Report', 'Queries', 'Source tables', 'Max depth',
                          'Summary rows', 'Detailed rows', 'Input file'],
                         index_rows, [45, 10, 13, 10, 13, 13, 45])
    from openpyxl.worksheet.hyperlink import Hyperlink
    for i, res in enumerate(results):
        cell = index_ws.cell(row=3 + i, column=2)
        cell.hyperlink = Hyperlink(ref=cell.coordinate,
                                   location="'%s'!B2" % res['summary_sheet'])
        cell.font = renderer.font_link

    sources, rootless = _aggregate_source_rows(results, renderer)
    _write_source_sheet(wb, renderer, sources, rootless)

    fd, aux_path = tempfile.mkstemp(suffix='.xlsx', dir=_temp_root())
    os.close(fd)
    wb.save(aux_path)
    return aux_path


def _merge_xlsx(ordered_sheets: list[tuple], template_path: str, output_path: str) -> str:
    """Assemble one workbook from pre-rendered worksheet XML by raw zip writing
    (no openpyxl re-render). Every source must share template_path's styles.xml
    (guaranteed by seed_styles). ordered_sheets is [(sheet_name, sheet_xml_bytes)]."""
    n = len(ordered_sheets)
    with zipfile.ZipFile(template_path) as tz:
        styles = tz.read('xl/styles.xml')
        theme = tz.read('xl/theme/theme1.xml')
        dotrels = tz.read('_rels/.rels')
        core = tz.read('docProps/core.xml')

    sheets = ''.join(f'<sheet name="{_xml_escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
                     for i, (name, _) in enumerate(ordered_sheets, 1))
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                f'xmlns:r="{_NS_R}"><sheets>{sheets}</sheets></workbook>')

    rels = [f'<Relationship Id="rId{i}" Type="{_NS_R}/worksheet" '
            f'Target="worksheets/sheet{i}.xml"/>' for i in range(1, n + 1)]
    rels.append(f'<Relationship Id="rId{n + 1}" Type="{_NS_R}/styles" Target="styles.xml"/>')
    rels.append(f'<Relationship Id="rId{n + 2}" Type="{_NS_R}/theme" Target="theme/theme1.xml"/>')
    wbrels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              + ''.join(rels) + '</Relationships>')

    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
          '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
          '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>']
    ct += [f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="{_WS_CT}"/>'
           for i in range(1, n + 1)]
    ct.append('</Types>')
    content_types = ''.join(ct)

    titles = ''.join(f'<vt:lpstr>{_xml_escape(name)}</vt:lpstr>' for name, _ in ordered_sheets)
    app = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
           'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
           '<Application>Microsoft Excel</Application>'
           f'<TitlesOfParts><vt:vector size="{n}" baseType="lpstr">{titles}</vt:vector></TitlesOfParts>'
           '</Properties>')

    tmp_out = output_path + '.tmp'
    with zipfile.ZipFile(tmp_out, 'w', zipfile.ZIP_DEFLATED) as out:
        out.writestr('[Content_Types].xml', content_types)
        out.writestr('_rels/.rels', dotrels)
        out.writestr('docProps/app.xml', app)
        out.writestr('docProps/core.xml', core)
        out.writestr('xl/workbook.xml', workbook)
        out.writestr('xl/_rels/workbook.xml.rels', wbrels)
        out.writestr('xl/styles.xml', styles)
        out.writestr('xl/theme/theme1.xml', theme)
        for i, (_, xml) in enumerate(ordered_sheets, 1):
            out.writestr(f'xl/worksheets/sheet{i}.xml', xml)
    return _safe_replace(tmp_out, output_path)


def _build_combined_merged(input_files, output_path, eff_rules, fmt, source_labels,
                           source_fallback, include_detailed, max_workers, report_progress,
                           errors=None, stop_on_error=False, cancel=None):
    """Render every report to a standalone file in parallel, then splice their
    Summary/Detailed sheets + a combined index + merged Source Tables into one
    workbook by raw XML assembly."""
    tmpdir = tempfile.mkdtemp(prefix='merge_', dir=_temp_root())
    aux_path = None
    try:
        out_paths = [os.path.join(tmpdir, f'r{i}.xlsx') for i in range(len(input_files))]
        tasks = _make_report_tasks(input_files, eff_rules, fmt, source_labels, source_fallback,
                                   include_detailed, include_source=False, out_paths=out_paths,
                                   seed=True, defer_naming=True)
        workers = _resolve_workers(max_workers, len(tasks))
        report_progress(f"Rendering {len(tasks)} report(s) with {workers} worker process(es)...")
        results = [r for r in _run_report_tasks(tasks, workers, report_progress,
                                                errors, stop_on_error, cancel) if r]
        if not results:
            return []          # every report failed; build_workbooks reports it

        report_progress("Merging into one workbook...")
        # Workers rendered under fixed 'Summary'/'Detailed' names; assign the
        # final, unique, data-driven sheet names now (data-driven names were
        # computed in the workers and returned as report_name).
        taken = {'list of reports', 'source tables'}
        for res in results:
            res['summary_sheet'] = ClientExcelRenderer.safe_sheet_name(res['report_name'], taken)
            res['detailed_sheet'] = (ClientExcelRenderer.safe_sheet_name(
                res['report_name'], taken, suffix=' (Detailed)') if include_detailed else None)
        aux_path = _build_aux(results, fmt, source_labels, source_fallback, include_detailed)
        ordered = []
        with zipfile.ZipFile(aux_path) as az:
            aux_styles = az.read('xl/styles.xml')
            ordered.append(('List of Reports', _sheet_xml_by_name(az, 'List of Reports')))
            for res in results:
                with zipfile.ZipFile(res['out_path']) as rz:
                    # The splice keeps each sheet's `s=` indices verbatim, so
                    # every file MUST share the aux's styles.xml (seed_styles
                    # guarantees this). Refuse to emit a mis-styled workbook.
                    if rz.read('xl/styles.xml') != aux_styles:
                        raise _StyleMergeMismatch(res['report_name'])
                    summary = _rename_summary_refs(_sheet_xml_by_name(rz, 'Summary'),
                                                   res['summary_sheet'])
                    ordered.append((res['summary_sheet'], summary))
                    if res['detailed_sheet']:
                        ordered.append((res['detailed_sheet'],
                                        _sheet_xml_by_name(rz, 'Detailed')))
            ordered.append(('Source Tables', _sheet_xml_by_name(az, 'Source Tables')))
        final = _merge_xlsx(ordered, aux_path, output_path)
        report_progress(f"Saved {final}")
        return [final]
    finally:
        if aux_path and os.path.exists(aux_path):
            os.remove(aux_path)
        shutil.rmtree(tmpdir, ignore_errors=True)


# ===========================================================================
# WORKBOOK ASSEMBLY
# ===========================================================================
def _temp_root() -> str:
    """A single app-owned scratch folder under the system temp dir, so all of
    our intermediate files live in one place (easy to find / sweep). The system
    temp dir is always writable and cross-platform, unlike the app's own folder
    which may be read-only (Program Files, a network share, etc.)."""
    root = os.path.join(tempfile.gettempdir(), 'DataLineageBuilder')
    os.makedirs(root, exist_ok=True)
    return root


def _sweep_stale_temp(max_age_hours: float = 1.0):
    """Best-effort cleanup of leftover scratch files from earlier runs that were
    killed before their normal cleanup (e.g. ended via Task Manager). Only items
    older than max_age_hours are removed, so a concurrent/active run's temp is
    never touched. Failures (locked files) are ignored."""
    root = os.path.join(tempfile.gettempdir(), 'DataLineageBuilder')
    if not os.path.isdir(root):
        return
    cutoff = time.time() - max_age_hours * 3600
    for name in os.listdir(root):
        p = os.path.join(root, name)
        try:
            if os.path.getmtime(p) >= cutoff:
                continue
            if os.path.isdir(p):
                shutil.rmtree(p, ignore_errors=True)
            else:
                os.remove(p)
        except OSError:
            pass


def _avoid_existing(path: str) -> str:
    """Return `path`, or the first 'name (N).ext' (N starting at 1) that doesn't
    exist yet, so an existing output is kept instead of overwritten (opt-in)."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    for n in range(1, 1000):
        alt = f'{base} ({n}){ext}'
        if not os.path.exists(alt):
            return alt
    return path


def _dedup_inputs(input_files: list[str]) -> list[str]:
    """Drop duplicate input paths (normalised), keeping first-occurrence order."""
    seen, out = set(), []
    for f in input_files:
        key = os.path.normcase(os.path.abspath(f))
        if key not in seen:
            seen.add(key)
            out.append(f)
    return out


def _default_output_path(input_files: list[str]) -> str:
    """Auto output path when none is given: a single report after its
    dashboard/table ('<name> Lineage.xlsx'), else 'Data Lineage Workbook.xlsx',
    next to the first input."""
    first = os.path.abspath(input_files[0])
    if len(input_files) == 1:
        safe = re.sub(r'[\\/:*?"<>|]', ' ', peek_report_name(first)).strip()
        name = f'{safe} Lineage.xlsx' if safe else 'Data Lineage Workbook.xlsx'
    else:
        name = 'Data Lineage Workbook.xlsx'
    return os.path.join(os.path.dirname(first), name)


def _separate_out_paths(input_files: list[str], output_path: str) -> list[str]:
    """Per-report output paths for separate-files mode ('<name> Lineage.xlsx');
    a single report honours output_path exactly."""
    if len(input_files) == 1:
        return [output_path]
    out_dir = os.path.dirname(output_path)
    taken, paths = set(), []
    for path in input_files:
        base = re.sub(r"[\\/:*?\"<>|]", ' ', peek_report_name(path)).strip() or 'report'
        cand, n = base, 2
        while cand.lower() in taken:
            cand = f'{base} ({n})'
            n += 1
        taken.add(cand.lower())
        paths.append(os.path.join(out_dir, f'{cand} Lineage.xlsx'))
    return paths


def _planned_outputs(input_files, output_path, combine, unformatted,
                     separate_detailed) -> list[str]:
    """The path(s) build_workbooks would write — for a pre-build overwrite check."""
    if not input_files:
        return []
    input_files = _dedup_inputs(input_files)
    output_path = os.path.abspath(output_path or _default_output_path(input_files))
    if not combine and not unformatted:
        return _separate_out_paths(input_files, output_path)
    paths = [output_path]
    if combine and not unformatted and separate_detailed:
        paths.append(re.sub(r'\.xlsx$', '', output_path, flags=re.I) + ' (Detailed).xlsx')
    return paths


def build_workbooks(input_files: list[str], output_path: str | None = None,
                    rules: dict | None = None, fmt: dict | None = None,
                    source_labels: dict | None = None, source_fallback: str | None = None,
                    include_detailed: bool = True, separate_detailed: bool = False,
                    unformatted: bool = False, combine: bool = True,
                    max_workers: int | None = None, write_error_log: bool = False,
                    stop_on_error: bool = False, avoid_overwrite: bool = False,
                    cancel=None, progress=None) -> list[str]:
    """Process every input file and write the output workbook(s).
    Returns the list of files written. `progress` is an optional callback(str).

    A file that can't be processed (e.g. not a valid impact report) is logged
    and skipped so the rest still build (unless `stop_on_error`, which aborts
    the whole run on the first failure); if `write_error_log` is set, the
    skipped files are also written to 'Lineage Errors.txt' in the output folder.

    Whether root queries with no upstream lineage are dropped is read per sheet
    set from rules['drop_no_lineage_summary'/'_detailed']."""
    import openpyxl

    eff_rules = TRIVIAL_RULES if rules is None else rules
    drop_summary = eff_rules.get('drop_no_lineage_summary', True)
    drop_detailed = eff_rules.get('drop_no_lineage_detailed', True)
    errors: list = []           # (report-name, message) for files that were skipped

    def report_progress(msg):
        log.info(msg)
        if progress:
            progress(msg)

    def _finish(written):
        """Common tail: report/write skipped-file errors, then return the
        written paths (or raise if nothing could be built at all)."""
        if errors:
            report_progress(f"  {len(errors)} of {len(input_files)} report(s) "
                            f"could not be processed and were skipped.")
            if write_error_log:
                _write_error_report(errors, output_path, report_progress)
        if not written:
            raise ValueError(f"None of the {len(input_files)} input file(s) could be "
                             f"processed. See the log for details.")
        return written

    if not input_files:
        raise ValueError("No input files provided.")
    _sweep_stale_temp()           # clear scratch left by any previously-killed run
    # Drop duplicate inputs (same file listed twice) so the output never gets
    # duplicate sheets; keep first-occurrence order.
    deduped = _dedup_inputs(input_files)
    if len(deduped) != len(input_files):
        report_progress(f"  Ignored {len(input_files) - len(deduped)} duplicate input(s).")
    input_files = deduped
    if output_path is None:
        output_path = _default_output_path(input_files)
    output_path = os.path.abspath(output_path)
    if avoid_overwrite:           # keep an existing file; write a numbered copy
        output_path = _avoid_existing(output_path)

    # ---- one standalone file per report (parallel, no merge) ----
    if not combine and not unformatted:
        written = _build_separate_files(input_files, output_path, eff_rules, fmt,
                                        source_labels, source_fallback, include_detailed,
                                        max_workers, report_progress, errors, stop_on_error,
                                        cancel, avoid_overwrite)
        for path in written:
            report_progress(f"Saved {path}")
        return _finish(written)

    # ---- combined workbook, parallel render + cheap XML merge ----
    # (the serial path below still handles separate_detailed and the 1-worker
    #  fallback, where no cross-file splice is needed)
    if (combine and not unformatted and not separate_detailed
            and _resolve_workers(max_workers, len(input_files)) > 1):
        try:
            return _finish(_build_combined_merged(
                input_files, output_path, eff_rules, fmt, source_labels, source_fallback,
                include_detailed, max_workers, report_progress, errors, stop_on_error, cancel))
        except _StyleMergeMismatch as exc:
            report_progress(f"  Style tables diverged for '{exc}' - falling back to a "
                            f"serial combine to keep formatting correct.")
            errors.clear()      # the serial path re-processes everything below
            # fall through to the serial combined path below

    # ---- build all lineage trees (serial: combined / unformatted) ----
    reports = []
    for path in input_files:
        if cancel is not None and cancel.is_set():
            raise BuildCancelled()
        try:
            builder = DataLineageBuilder(path, rules=rules).load()
            report_progress(f"Tracing lineage: {builder.report_name}")
            summary_roots, expanded_at = builder.build_tree(
                summarize=True, drop_no_lineage=drop_summary)
            detailed = (builder.build_tree(summarize=False, drop_no_lineage=drop_detailed)
                        if include_detailed else (None, None))
            reports.append({'builder': builder, 'summary': (summary_roots, expanded_at),
                            'detailed': detailed})
        except Exception as exc:
            if stop_on_error:
                raise RuntimeError(f"{derive_report_name(path)}: {exc}") from exc
            label = derive_report_name(path)
            log.warning("  skipped %s: %s", label, exc)
            errors.append((label, str(exc)))

    if unformatted:
        return _finish([_export_unformatted(reports, output_path)] if reports else [])

    if not reports:
        return _finish([])      # every input failed -> nothing to write

    renderer = ClientExcelRenderer(fmt, source_labels, source_fallback)
    written = []

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    multi = len(reports) > 1          # an index sheet is pointless for one report
    index_ws = wb.create_sheet('List of Reports') if multi else None

    detail_wb = None
    detail_path = None
    if include_detailed and separate_detailed:
        detail_wb = openpyxl.Workbook()
        detail_wb.remove(detail_wb.active)
        detail_path = re.sub(r'\.xlsx$', '', output_path, flags=re.I) + ' (Detailed).xlsx'
        if avoid_overwrite:
            detail_path = _avoid_existing(detail_path)

    taken_main, taken_detail = {'list of reports', 'source tables'}, set()
    index_rows = []
    for rep in reports:
        if cancel is not None and cancel.is_set():
            raise BuildCancelled()
        builder = rep['builder']
        report_progress(f"Writing sheets: {builder.report_name}")
        summary_roots, expanded_at = rep['summary']
        s_name = renderer.safe_sheet_name(builder.report_name, taken_main)
        s_info = renderer.render_lineage_sheet(wb, s_name, builder, summary_roots,
                                               expanded_at, link_dups=True)
        d_info = None
        if include_detailed:
            d_roots, d_exp = rep['detailed']
            target_wb = detail_wb if separate_detailed else wb
            taken = taken_detail if separate_detailed else taken_main
            d_name = renderer.safe_sheet_name(builder.report_name, taken,
                                              suffix='' if separate_detailed else ' (Detailed)')
            d_info = renderer.render_lineage_sheet(target_wb, d_name, builder, d_roots,
                                                   d_exp, link_dups=False)
        rep['s_info'], rep['d_info'] = s_info, d_info
        index_rows.append([builder.report_name, s_info['queries'],
                           len(builder.source_tables()), f"L{s_info['max_depth']}",
                           s_info['rows'], d_info['rows'] if d_info else '-',
                           os.path.basename(builder.input_file_path)])

    # ---- index sheet (only when there's more than one report) ----
    if multi:
        renderer._write_grid(index_ws,
                             ['Report', 'Queries', 'Source tables', 'Max depth',
                              'Summary rows', 'Detailed rows', 'Input file'],
                             index_rows,
                             [45, 10, 13, 10, 13, 13, 45])
        from openpyxl.worksheet.hyperlink import Hyperlink
        for i, rep in enumerate(reports):
            cell = index_ws.cell(row=3 + i, column=2)
            cell.hyperlink = Hyperlink(ref=cell.coordinate,
                                       location=f"'{rep['s_info']['sheet']}'!B2")
            cell.font = renderer.font_link

    # ---- source tables rollup ----
    _add_source_rollup(wb, renderer, reports)

    written.append(_safe_save(wb, output_path))
    if detail_wb is not None:
        written.append(_safe_save(detail_wb, detail_path))
    for path in written:
        report_progress(f"Saved {path}")
    return _finish(written)


_SOURCE_HEADER = ['db', 'schema', 'object', 'type', 'List of sources',
                  'Used by # reports', 'Reports']
_SOURCE_WIDTHS = [24, 18, 40, 8, 19, 16, 60]


def _write_source_sheet(wb, renderer: ClientExcelRenderer, sources: list[list],
                        rootless: list[list]):
    """Write the 'Source Tables' sheet: real source tables (migration worklist)
    on top; no-lineage root queries in a labelled section below."""
    ws = wb.create_sheet('Source Tables')
    # Skip the real-sources header entirely if there are none (e.g. an
    # all-rootless report), so the rootless section isn't preceded by an
    # empty orphan header.
    last = renderer._write_grid(ws, _SOURCE_HEADER, sources, _SOURCE_WIDTHS) if sources else 1
    if rootless:
        title_row = last + 2  # one blank spacer row, then a section label
        label = ws.cell(row=title_row, column=2,
                        value='Queries with no upstream lineage (not migrated)')
        label.font = renderer.font_data_bold
        renderer._write_grid(ws, _SOURCE_HEADER, rootless, _SOURCE_WIDTHS,
                             start_row=title_row + 1)


def _rollup_rows(usage: dict, renderer, exclude=frozenset()) -> list[list]:
    rows = []
    for name in sorted(usage):
        if name in exclude:
            continue
        meta, reps = usage[name]['meta'], usage[name]['reports']
        rows.append([meta['Database'], meta['Schema'], meta['Name'], meta['Type'],
                     renderer.source_label(name, meta), len(reps),
                     ', '.join(sorted(set(reps)))])
    return rows


def _add_source_rollup(wb, renderer: ClientExcelRenderer, reports: list[dict]):
    """Source-table rollup across all reports (serial path; aggregates from
    the in-memory builders)."""
    def collect(getter):
        usage: dict[str, dict] = {}
        for rep in reports:
            builder = rep['builder']
            for name in getter(builder):
                entry = usage.setdefault(name, {'meta': builder.object_meta(name), 'reports': []})
                entry['reports'].append(builder.report_name)
        return usage
    src = collect(lambda b: b.source_tables())
    root = collect(lambda b: b.rootless_queries())
    sources = _rollup_rows(src, renderer)
    rootless = _rollup_rows(root, renderer, exclude=set(src))
    _write_source_sheet(wb, renderer, sources, rootless)


def _aggregate_source_rows(results: list[dict], renderer) -> tuple[list[list], list[list]]:
    """Source-table rollup for the parallel/merge path: aggregate the picklable
    source rows the child processes returned (each row is
    [db, schema, object, type, label, consolidated-name])."""
    def collect(key):
        usage: dict[str, dict] = {}
        for res in results:
            for row in res[key]:
                name = row[5]
                meta = {'Database': row[0], 'Schema': row[1], 'Name': row[2],
                        'Type': row[3], '_label': row[4]}
                entry = usage.setdefault(name, {'meta': meta, 'reports': []})
                entry['reports'].append(res['report_name'])
        return usage

    def rows(usage, exclude=frozenset()):
        out = []
        for name in sorted(usage):
            if name in exclude:
                continue
            m, reps = usage[name]['meta'], usage[name]['reports']
            out.append([m['Database'], m['Schema'], m['Name'], m['Type'], m['_label'],
                        len(set(reps)), ', '.join(sorted(set(reps)))])
        return out
    src = collect('source_rows')
    root = collect('rootless_rows')
    return rows(src), rows(root, exclude=set(src))


def _export_unformatted(reports: list[dict], output_path: str) -> str:
    """Raw wide dump: one row per leaf path, plain single-row header."""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for rep in reports:
            builder = rep['builder']
            for label, data in (('Summary', rep['summary']), ('Detailed', rep['detailed'])):
                if data is None or data[0] is None:
                    continue
                roots = data[0]
                records = []
                for root in roots:
                    for leaf in root.leaves():
                        path = _path_to(root, leaf)
                        rec = {'L0': path[0].name}
                        for node in path[1:]:
                            meta = builder.object_meta(node.name)
                            i = node.level
                            rec.update({f'Database_L{i}': meta['Database'],
                                        f'Schema_L{i}': meta['Schema'],
                                        f'L{i}': meta['Name'],
                                        f'Type_L{i}': meta['Type']})
                        rec['Status'] = leaf.status
                        rec['Remark'] = leaf.remark
                        records.append(rec)
                sheet = ClientExcelRenderer.safe_sheet_name(
                    builder.report_name, {s.lower() for s in writer.sheets},
                    suffix=f' {label}')
                pd.DataFrame(records).to_excel(writer, sheet_name=sheet, index=False)
    return output_path


def _path_to(root: Node, leaf: Node) -> list[Node]:
    """Root-to-leaf node path (DFS)."""
    path = []

    def walk(node):
        path.append(node)
        if node is leaf:
            return True
        for child in node.children:
            if walk(child):
                return True
        path.pop()
        return False

    walk(root)
    return path


def _write_error_report(errors: list, output_path: str, report_progress) -> str | None:
    """Write the skipped-file errors to 'Lineage Errors.txt' in the output
    folder, one section per report/object. Off by default (opt-in)."""
    path = os.path.join(os.path.dirname(output_path) or '.', 'Lineage Errors.txt')
    try:
        with open(path, 'w', encoding='utf-8') as fh:
            fh.write("Reports that could not be processed\n")
            fh.write("=" * 38 + "\n\n")
            for name, msg in errors:
                fh.write(f"## {name}\n{msg}\n\n")
        report_progress(f"Wrote error report: {path}")
        return path
    except OSError as exc:
        report_progress(f"  Could not write error report: {exc}")
        return None


def _safe_save(wb, path: str) -> str:
    """Save the workbook; fall back to a numbered name if the file is locked by Excel."""
    try:
        wb.save(path)
        return path
    except PermissionError:
        base, ext = os.path.splitext(path)
        for n in range(1, 100):
            alt = f'{base} ({n}){ext}'
            try:
                wb.save(alt)
                log.warning("'%s' is open in Excel — saved as '%s' instead.",
                            os.path.basename(path), os.path.basename(alt))
                return alt
            except PermissionError:
                continue
        raise


# ===========================================================================
# CONFIG / CLI
# ===========================================================================
def load_config(path: str):
    """Merge a JSON config file over the module-level defaults."""
    with open(path, encoding='utf-8') as fh:
        cfg = json.load(fh)
    TRIVIAL_RULES.update(cfg.get('trivial_rules', {}))
    EXCEL_FORMAT.update(cfg.get('excel_format', {}))
    if 'source_labels' in cfg:
        SOURCE_LABELS.clear()
        SOURCE_LABELS.update(cfg['source_labels'])
    global SOURCE_LABEL_FALLBACK
    SOURCE_LABEL_FALLBACK = cfg.get('source_label_fallback', SOURCE_LABEL_FALLBACK)
    log.info("Loaded config overrides from %s", path)


def main(argv=None):
    import multiprocessing
    multiprocessing.freeze_support()   # safe re-entry for parallel worker processes
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(
        description="Recursively expand data lineage from Atlan impact reports "
                    "into a client-formatted Excel workbook.")
    parser.add_argument('-i', '--input', nargs='+',
                        help="One or more Atlan impact reports (.csv or .xlsx)")
    parser.add_argument('-o', '--output', help="Output .xlsx path (optional)")
    parser.add_argument('-s', '--separate', action='store_true',
                        help="Write Detailed sheets to a separate _Detailed.xlsx file")
    parser.add_argument('--no-detailed', action='store_true',
                        help="Skip the Detailed sheets entirely")
    parser.add_argument('--unformatted', action='store_true',
                        help="Raw wide dump without client styling")
    parser.add_argument('--no-combine', action='store_true',
                        help="Write one workbook per report instead of merging into one")
    parser.add_argument('--workers', type=int, default=None,
                        help="Max worker processes for rendering (default: one per "
                             "report, capped at the CPU count)")
    parser.add_argument('--error-log', action='store_true',
                        help="Write skipped files to 'Lineage Errors.txt' in the output folder")
    parser.add_argument('--stop-on-error', action='store_true',
                        help="Abort the whole run on the first file that fails "
                             "(default: skip it and continue)")
    parser.add_argument('-y', '--yes', action='store_true',
                        help="Overwrite existing output files without prompting")
    parser.add_argument('--drop-no-lineage', action=argparse.BooleanOptionalAction,
                        default=True,
                        help="Ignore root queries that have no upstream lineage "
                             "(e.g. tables hardcoded inside PowerBI). On by default; "
                             "use --no-drop-no-lineage to keep them.")
    parser.add_argument('--config', help="JSON file overriding rules/format/source labels")
    parser.add_argument('--gui', action='store_true', help="Launch the graphical interface")
    args = parser.parse_args(argv)

    if args.config:
        load_config(args.config)

    if args.gui:
        import lineage_ui
        lineage_ui.run()
        return

    inputs = args.input
    if not inputs:
        try:
            raw = input("Path to the Atlan impact report (.csv or .xlsx): ").strip().strip('"')
        except EOFError:
            raw = ''
        if not raw:
            print("No input provided. Exiting.")
            sys.exit(1)
        inputs = [raw]

    cli_rules = dict(TRIVIAL_RULES)
    cli_rules['drop_no_lineage_summary'] = args.drop_no_lineage
    cli_rules['drop_no_lineage_detailed'] = args.drop_no_lineage

    # Mirror the GUI's option interactions: rather than fail or produce bogus
    # output, proceed with the sensible behaviour and tell the user what was
    # ignored.
    if args.separate and args.no_detailed:
        log.warning("--separate ignored: there are no Detailed sheets (--no-detailed).")
    if args.separate and args.no_combine:
        log.warning("--separate ignored: each per-report file (--no-combine) already "
                    "contains its own Detailed sheet.")
    if args.unformatted and args.separate:
        log.warning("--separate ignored: --unformatted writes a single raw workbook.")

    # Confirm before overwriting existing output (skip with --yes). The prompt
    # offers a third choice -- save a numbered copy instead of overwriting.
    cli_avoid_overwrite = False
    if not args.yes:
        existing = [p for p in _planned_outputs(inputs, args.output, not args.no_combine,
                                                args.unformatted, args.separate)
                    if os.path.exists(p)]
        if existing:
            print("These output file(s) already exist:")
            for p in existing:
                print(f"  {p}")
            print("  [O]verwrite, save a [N]umbered copy (keeps the originals), or [C]ancel?")
            try:
                resp = input("Choice [o/n/C]: ").strip().lower()
            except EOFError:
                resp = ''
            if resp in ('o', 'overwrite'):
                pass
            elif resp in ('n', 'number', 'numbered', 'copy'):
                cli_avoid_overwrite = True
            else:
                print("Cancelled.")
                sys.exit(0)

    try:
        written = build_workbooks(inputs, args.output, rules=cli_rules,
                                  include_detailed=not args.no_detailed,
                                  separate_detailed=args.separate,
                                  unformatted=args.unformatted,
                                  combine=not args.no_combine,
                                  max_workers=args.workers,
                                  write_error_log=args.error_log,
                                  stop_on_error=args.stop_on_error,
                                  avoid_overwrite=cli_avoid_overwrite)
    except Exception as exc:
        log.error("Failed to process lineage: %s", exc)
        sys.exit(1)
    print("Done. Output file(s):")
    for path in written:
        print(f"  {path}")


if __name__ == '__main__':
    main()
