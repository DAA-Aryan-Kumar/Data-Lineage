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
import json
import logging
import os
import re
import sys
from collections import deque

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

        df = self._canonicalize_columns(df)
        missing = [c for c in self.REQUIRED_COLS if c not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns in input: {missing}")
        self._preprocess(df)
        return self

    def _canonicalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Rename input columns to the canonical names, matching case- and
        whitespace-insensitively (e.g. 'Lineage depth' or 'Immediate Upstream ')."""
        norm = lambda s: re.sub(r'\s+', ' ', str(s).strip()).lower()
        lookup = {norm(c): c for c in df.columns}  # first occurrence wins
        renames = {lookup[norm(canon)]: canon
                   for canon in self.REQUIRED_COLS if norm(canon) in lookup}
        return df.rename(columns=renames) if renames else df

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
        if len(candidate) > budget:                       # 2. shorten each word
            words = candidate.split(' ')
            candidate = ' '.join(w[:4] for w in words)
        if len(candidate) > budget:                       # 3. last-resort truncate
            candidate = candidate[:budget].rstrip()

        candidate = (candidate + suffix).strip()
        base, n = candidate, 2
        while candidate.lower() in taken:                 # 4. ensure uniqueness
            tail = f' ({n})'
            candidate = base[:31 - len(tail)].rstrip() + tail
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
def _resolve_workers(max_workers, n_tasks: int) -> int:
    """Clamp the worker count to [1, n_tasks], defaulting to cpu_count-1."""
    if not max_workers or max_workers < 1:
        max_workers = max(1, (os.cpu_count() or 2) - 1)
    return max(1, min(int(max_workers), max(1, n_tasks)))


def _source_rows(builder: 'DataLineageBuilder', renderer: 'ClientExcelRenderer', names) -> list[list]:
    """[db, schema, object, type, label, consolidated-name] per source object."""
    rows = []
    for n in names:
        m = builder.object_meta(n)
        rows.append([m['Database'], m['Schema'], m['Name'], m['Type'],
                     renderer.source_label(n, m), n])
    return rows


def _render_report_task(task: dict) -> dict:
    """Worker (runs in a child process): build one report's lineage trees and
    render them to a standalone .xlsx, returning picklable metadata + the source
    rows the parent needs for the combined index / Source Tables sheets."""
    import openpyxl
    builder = DataLineageBuilder(task['path'], rules=task['rules']).load()
    drop_s, drop_d = task['drop_summary'], task['drop_detailed']
    s_roots, s_exp = builder.build_tree(summarize=True, drop_no_lineage=drop_s)
    d_roots = d_exp = None
    if task['include_detailed']:
        d_roots, d_exp = builder.build_tree(summarize=False, drop_no_lineage=drop_d)

    renderer = ClientExcelRenderer(task['fmt'], task['source_labels'], task['source_fallback'])
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    s_info = renderer.render_lineage_sheet(wb, task['summary_sheet'], builder,
                                           s_roots, s_exp, link_dups=True)
    d_info = None
    if d_roots is not None:
        d_info = renderer.render_lineage_sheet(wb, task['detailed_sheet'], builder,
                                               d_roots, d_exp, link_dups=False)
    if task['include_source']:
        _add_source_rollup(wb, renderer, [{'builder': builder}])
    out = _safe_save(wb, task['out_path'])
    return {
        'report_name': builder.report_name, 'out_path': out,
        'input_basename': os.path.basename(builder.input_file_path),
        'summary_sheet': task['summary_sheet'], 'detailed_sheet': task['detailed_sheet'],
        'queries': s_info['queries'], 'max_depth': s_info['max_depth'],
        'summary_rows': s_info['rows'], 'detailed_rows': (d_info['rows'] if d_info else '-'),
        'source_count': len(builder.source_tables()),
        'source_rows': _source_rows(builder, renderer, builder.source_tables()),
        'rootless_rows': _source_rows(builder, renderer, builder.rootless_queries()),
    }


def _run_report_tasks(tasks: list[dict], workers: int, report_progress) -> list[dict]:
    """Render all report tasks, sequentially (workers<=1) or across a process
    pool. Per-report progress is logged as each one finishes."""
    n = len(tasks)
    results: list = [None] * n
    if workers <= 1:
        for i, task in enumerate(tasks):
            results[i] = _render_report_task(task)
            report_progress(f"  [{i + 1}/{n}] rendered {results[i]['report_name']}")
        return results
    from concurrent.futures import ProcessPoolExecutor, as_completed
    with ProcessPoolExecutor(max_workers=workers) as pool:
        futs = {pool.submit(_render_report_task, t): i for i, t in enumerate(tasks)}
        done = 0
        for fut in as_completed(futs):
            i = futs[fut]
            results[i] = fut.result()
            done += 1
            report_progress(f"  [{done}/{n}] rendered {results[i]['report_name']}")
    return results


def _make_report_tasks(input_files, eff_rules, fmt, source_labels, source_fallback,
                       include_detailed, include_source, out_paths) -> list[dict]:
    """Build per-report task dicts with globally-unique sheet names so the
    rendered sheets never collide when merged and internal hyperlinks stay valid."""
    drop_s = eff_rules.get('drop_no_lineage_summary', True)
    drop_d = eff_rules.get('drop_no_lineage_detailed', True)
    taken = {'list of reports', 'source tables'}
    tasks = []
    for path, out_path in zip(input_files, out_paths):
        name = derive_report_name(path)
        s_sheet = ClientExcelRenderer.safe_sheet_name(name, taken)
        d_sheet = (ClientExcelRenderer.safe_sheet_name(name, taken, suffix=' (Detailed)')
                   if include_detailed else None)
        tasks.append({'path': path, 'rules': eff_rules, 'fmt': fmt,
                      'source_labels': source_labels, 'source_fallback': source_fallback,
                      'include_detailed': include_detailed, 'include_source': include_source,
                      'drop_summary': drop_s, 'drop_detailed': drop_d,
                      'summary_sheet': s_sheet, 'detailed_sheet': d_sheet, 'out_path': out_path})
    return tasks


def _build_separate_files(input_files, output_path, eff_rules, fmt, source_labels,
                          source_fallback, include_detailed, max_workers, report_progress):
    """One standalone workbook per report (Summary + Detailed + Source Tables),
    rendered in parallel. No index sheet (it's meaningless for single files)."""
    out_dir = os.path.dirname(output_path)
    taken = set()
    out_paths = []
    for path in input_files:
        base = re.sub(r"[\\/:*?\"<>|]", ' ', derive_report_name(path)).strip() or 'report'
        cand, n = base, 2
        while cand.lower() in taken:
            cand = f'{base} ({n})'
            n += 1
        taken.add(cand.lower())
        out_paths.append(os.path.join(out_dir, cand + '.xlsx'))
    tasks = _make_report_tasks(input_files, eff_rules, fmt, source_labels, source_fallback,
                               include_detailed, include_source=True, out_paths=out_paths)
    workers = _resolve_workers(max_workers, len(tasks))
    report_progress(f"Rendering {len(tasks)} report(s) to separate files "
                    f"with {workers} worker process(es)...")
    results = _run_report_tasks(tasks, workers, report_progress)
    return [r['out_path'] for r in results if r]


# ===========================================================================
# WORKBOOK ASSEMBLY
# ===========================================================================
def build_workbooks(input_files: list[str], output_path: str | None = None,
                    rules: dict | None = None, fmt: dict | None = None,
                    source_labels: dict | None = None, source_fallback: str | None = None,
                    include_detailed: bool = True, separate_detailed: bool = False,
                    unformatted: bool = False, combine: bool = True,
                    max_workers: int | None = None, progress=None) -> list[str]:
    """Process every input file and write the output workbook(s).
    Returns the list of files written. `progress` is an optional callback(str).

    Whether root queries with no upstream lineage are dropped is read per sheet
    set from rules['drop_no_lineage_summary'/'_detailed']."""
    import openpyxl

    eff_rules = TRIVIAL_RULES if rules is None else rules
    drop_summary = eff_rules.get('drop_no_lineage_summary', True)
    drop_detailed = eff_rules.get('drop_no_lineage_detailed', True)

    def report_progress(msg):
        log.info(msg)
        if progress:
            progress(msg)

    if not input_files:
        raise ValueError("No input files provided.")
    if output_path is None:
        first = os.path.abspath(input_files[0])
        output_path = os.path.join(os.path.dirname(first), 'Data Lineage Workbook.xlsx')
    output_path = os.path.abspath(output_path)

    # ---- one standalone file per report (parallel, no merge) ----
    if not combine and not unformatted:
        written = _build_separate_files(input_files, output_path, eff_rules, fmt,
                                        source_labels, source_fallback, include_detailed,
                                        max_workers, report_progress)
        for path in written:
            report_progress(f"Saved {path}")
        return written

    # ---- build all lineage trees (serial: combined / unformatted) ----
    reports = []
    for path in input_files:
        builder = DataLineageBuilder(path, rules=rules).load()
        report_progress(f"Tracing lineage: {builder.report_name}")
        summary_roots, expanded_at = builder.build_tree(
            summarize=True, drop_no_lineage=drop_summary)
        detailed = (builder.build_tree(summarize=False, drop_no_lineage=drop_detailed)
                    if include_detailed else (None, None))
        reports.append({'builder': builder, 'summary': (summary_roots, expanded_at),
                        'detailed': detailed})

    if unformatted:
        return [_export_unformatted(reports, output_path)]

    renderer = ClientExcelRenderer(fmt, source_labels, source_fallback)
    written = []

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    index_ws = wb.create_sheet('List of Reports')

    detail_wb = None
    detail_path = None
    if include_detailed and separate_detailed:
        detail_wb = openpyxl.Workbook()
        detail_wb.remove(detail_wb.active)
        detail_path = re.sub(r'\.xlsx$', '', output_path, flags=re.I) + '_Detailed.xlsx'

    taken_main, taken_detail = {'list of reports', 'source tables'}, set()
    index_rows = []
    for rep in reports:
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

    # ---- index sheet ----
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
    return written


def _add_source_rollup(wb, renderer: ClientExcelRenderer, reports: list[dict]):
    """Source-table rollup across all reports. The real source tables (the
    migration worklist) come first; root queries that have no lineage of their
    own are kept but separated into a labelled section below."""
    def collect(getter, exclude=frozenset()) -> tuple[list[list], set]:
        usage: dict[str, dict] = {}
        for rep in reports:
            builder = rep['builder']
            for name in getter(builder):
                if name in exclude:
                    continue
                entry = usage.setdefault(name, {'meta': builder.object_meta(name), 'reports': []})
                entry['reports'].append(builder.report_name)
        rows = []
        for name in sorted(usage):
            meta, reps = usage[name]['meta'], usage[name]['reports']
            rows.append([meta['Database'], meta['Schema'], meta['Name'], meta['Type'],
                         renderer.source_label(name, meta), len(reps),
                         ', '.join(sorted(set(reps)))])
        return rows, set(usage)

    sources, source_names = collect(lambda b: b.source_tables())
    # A name that is a real source in any report belongs in the worklist above,
    # never the rootless section — even if another report has it as a root.
    rootless, _ = collect(lambda b: b.rootless_queries(), exclude=source_names)
    header = ['db', 'schema', 'object', 'type', 'List of sources',
              'Used by # reports', 'Reports']
    widths = [24, 18, 40, 8, 19, 16, 60]

    ws = wb.create_sheet('Source Tables')
    # Skip the real-sources header entirely if there are none (e.g. an
    # all-rootless report), so the rootless section isn't preceded by an
    # empty orphan header.
    last = renderer._write_grid(ws, header, sources, widths) if sources else 1
    if rootless:
        title_row = last + 2  # one blank spacer row, then a section label
        label = ws.cell(row=title_row, column=2,
                        value='Queries with no upstream lineage (not migrated)')
        label.font = renderer.font_data_bold
        renderer._write_grid(ws, header, rootless, widths, start_row=title_row + 1)


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
    try:
        written = build_workbooks(inputs, args.output, rules=cli_rules,
                                  include_detailed=not args.no_detailed,
                                  separate_detailed=args.separate,
                                  unformatted=args.unformatted)
    except Exception as exc:
        log.error("Failed to process lineage: %s", exc)
        sys.exit(1)
    print("Done. Output file(s):")
    for path in written:
        print(f"  {path}")


if __name__ == '__main__':
    main()
