"""
lineage_ui.py
=============
Graphical front-end for data_lineage.py — lets non-technical users build
client-formatted lineage workbooks from Atlan impact reports.

Launch with:  python lineage_ui.py        (or: python data_lineage.py --gui)

No dependencies beyond the engine's (pandas + openpyxl); the UI itself is
pure standard-library Tkinter. Settings persist to lineage_settings.json
next to this file, and the same file works as --config for the CLI.
"""

from __future__ import annotations

import json
import logging
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import data_lineage as engine
except ImportError as exc:
    # Launched (e.g. by pythonw) on an interpreter missing pandas/openpyxl —
    # there's no console to show the traceback, so surface it as a dialog.
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showerror(
        "Missing Python library",
        f"The Data Lineage Builder could not start because a required "
        f"library is missing:\n\n    {exc}\n\n"
        f"Install the dependencies with:\n\n    pip install pandas openpyxl")
    raise SystemExit(1)

# When frozen by PyInstaller, __file__ lives in a temp extraction dir, so keep
# the settings next to the executable instead (a stable, writable location).
_APP_DIR = (os.path.dirname(sys.executable) if getattr(sys, 'frozen', False)
            else os.path.dirname(os.path.abspath(__file__)))
SETTINGS_FILE = os.path.join(_APP_DIR, 'lineage_settings.json')

ACCENT = '#7030A0'        # client purple
ACCENT_LIGHT = '#F2CEEF'  # client pink


def _safe_config(widget, **opts):
    """Apply widget options one at a time, skipping any the widget rejects
    (e.g. a Listbox has no 'insertbackground')."""
    for key, val in opts.items():
        try:
            widget.configure(**{key: val})
        except tk.TclError:
            pass


class QueueLogHandler(logging.Handler):
    """Routes engine log records into the UI thread via a queue."""

    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(('LOG', record.levelno, self.format(record)))


class LineageApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Data Lineage Builder")
        # Fit the window within the screen so nothing is clipped on shorter
        # displays (dark mode's clam widgets are a touch taller); tab content
        # scrolls, and the action bar stays pinned at the bottom.
        sh = root.winfo_screenheight()
        root.geometry(f"980x{min(720, max(520, sh - 90))}")
        root.minsize(820, 460)
        self._set_window_icon()

        self.log_queue: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.cancel_event: threading.Event | None = None   # set to stop a running build
        self.written_files: list[str] = []
        self._files: list[str] = []   # input report paths, in processing order
        self.dark = tk.BooleanVar(value=False)
        self._themed = []   # (tk widget, light_bg, light_fg) recoloured on theme switch
        self._canvases = []           # scrollable-tab canvases (recoloured on theme switch)
        self._tab_canvases = {}       # notebook tab path -> its canvas (for mouse wheel)
        self._watermark_lbl = None

        self._apply_theme()
        self._build_layout()
        self._load_settings()
        self._apply_theme()   # re-apply now that widgets exist (honours saved dark)
        self._poll_log_queue()

    def _register_themed(self, widget, light_bg='#ffffff', light_fg='#000000'):
        """Track a classic-tk widget so dark mode can recolour it."""
        self._themed.append((widget, light_bg, light_fg))
        return widget

    def _set_window_icon(self):
        """Give the title bar and taskbar a crisp icon at any DPI. Prefer
        iconphoto with per-size PNGs (Tk picks the best size, sharper than
        scaling one .ico frame); fall back to app.ico. Also set an explicit
        AppUserModelID so Windows uses our icon on the taskbar rather than
        grouping under python/pythonw."""
        base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "Accordion.DataLineageBuilder")
        except Exception:
            pass
        # Base icon from the .ico (reliable on the existing root); then override
        # with crisp per-size PNGs via iconphoto when they're available.
        try:
            ico = os.path.join(base, 'app.ico')
            if os.path.exists(ico):
                self.root.iconbitmap(ico)
        except Exception:
            pass
        try:
            imgs = []
            for size in (256, 128, 64, 48, 32, 24, 16):
                png = os.path.join(base, f'app_{size}.png')
                if os.path.exists(png):
                    imgs.append(tk.PhotoImage(file=png))
            if imgs:
                self._icon_imgs = imgs   # keep refs alive
                self.root.iconphoto(True, *imgs)
        except Exception:
            pass

    # ------------------------------------------------------------------ UI --
    def _apply_theme(self, *_):
        """Apply light (native 'vista') or dark ('clam' + palette) theme.
        Light is the original look, untouched; dark recolours ttk styles, the
        root, and the registered classic-tk widgets."""
        style = ttk.Style(self.root)
        dark = self.dark.get()
        if not dark:
            try:
                style.theme_use('vista')
            except tk.TclError:
                pass
            style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'), foreground=ACCENT)
            style.configure('Sub.TLabel', font=('Segoe UI', 9), foreground='#666666')
            style.configure('Run.TButton', font=('Segoe UI', 11, 'bold'))
            style.configure('TNotebook.Tab', font=('Segoe UI', 10), padding=(14, 6))
            try:
                self.root.configure(background='SystemButtonFace')
            except tk.TclError:
                pass
            for w, lbg, lfg in self._themed:
                _safe_config(w, bg=lbg, fg=lfg, insertbackground=lfg,
                             selectbackground='#cce8ff', selectforeground='#000000',
                             highlightbackground='SystemButtonFace')
            for c in self._canvases:
                _safe_config(c, bg='SystemButtonFace')
            if self._watermark_lbl:
                self._watermark_lbl.configure(background='SystemButtonFace', foreground='#C2C2C2')
            self._apply_log_tags()
            return

        BG, FG, FIELD, SUB, SEL = '#2b2b2b', '#e6e6e6', '#3c3f41', '#9aa0a6', '#365880'
        style.theme_use('clam')
        style.configure('.', background=BG, foreground=FG, fieldbackground=FIELD,
                        bordercolor='#555555', lightcolor=BG, darkcolor=BG,
                        troughcolor=FIELD, arrowcolor=FG, insertcolor=FG)
        for cls in ('TFrame', 'TLabel', 'TLabelframe', 'TLabelframe.Label', 'TCheckbutton'):
            style.configure(cls, background=BG, foreground=FG)
        style.map('TCheckbutton', background=[('active', BG)])
        # clam's default checkbox mark renders as an ambiguous "X" on a dark
        # palette, so swap in a drawn box + green tick image (dark only; light
        # mode keeps the native vista checkbox).
        self._install_dark_check(style, BG, FIELD)
        style.configure('TButton', background='#3c3f41', foreground=FG)
        style.map('TButton', background=[('active', '#4a4d4f'), ('pressed', '#4a4d4f')])
        style.configure('TEntry', fieldbackground=FIELD, foreground=FG)
        style.configure('TSpinbox', fieldbackground=FIELD, foreground=FG, arrowcolor=FG)
        style.configure('TCombobox', fieldbackground=FIELD, foreground=FG, arrowcolor=FG)
        style.map('TCombobox', fieldbackground=[('readonly', FIELD)], foreground=[('readonly', FG)])
        style.configure('TNotebook', background=BG, bordercolor='#555555')
        style.configure('TNotebook.Tab', background='#3c3f41', foreground=FG,
                        font=('Segoe UI', 10), padding=(14, 6))
        style.map('TNotebook.Tab', background=[('selected', BG)], foreground=[('selected', FG)])
        style.configure('Horizontal.TProgressbar', background=ACCENT, troughcolor=FIELD)
        style.configure('Title.TLabel', background=BG, foreground='#C9A6E8',
                        font=('Segoe UI', 16, 'bold'))
        style.configure('Sub.TLabel', background=BG, foreground=SUB, font=('Segoe UI', 9))
        style.configure('Run.TButton', font=('Segoe UI', 11, 'bold'))
        try:
            self.root.configure(background=BG)
        except tk.TclError:
            pass
        for w, lbg, lfg in self._themed:
            _safe_config(w, bg=FIELD, fg=FG, insertbackground=FG,
                         selectbackground=SEL, selectforeground='#ffffff',
                         highlightbackground=BG)
        for c in self._canvases:
            _safe_config(c, bg=BG)
        if self._watermark_lbl:
            self._watermark_lbl.configure(background=BG, foreground='#5e5e5e')
        self._apply_log_tags()

    def _apply_log_tags(self):
        """Colour the log tags (errors / warnings / completions) for the active
        theme so they stand out. Safe to call before the log widget exists."""
        if not hasattr(self, 'log_text'):
            return
        if self.dark.get():
            cols = {'err': '#ff7b72', 'warn': '#e3b341', 'ok': '#7ee787', 'info': '#e6e6e6'}
        else:
            cols = {'err': '#c0392b', 'warn': '#b9770e', 'ok': '#1e7e34', 'info': '#222222'}
        for tag, col in cols.items():
            bold = tag in ('err', 'warn')
            self.log_text.tag_configure(tag, foreground=col, spacing1=1,
                                        font=('Consolas', 9, 'bold') if bold else ('Consolas', 9))

    def _make_check_img(self, checked: bool, bg: str, field: str,
                        disabled: bool = False) -> tk.PhotoImage:
        """A 16px checkbox indicator: a bordered box, plus a green tick if
        checked. `disabled` mutes the colours so a disabled box reads as greyed
        out (clam doesn't dim a custom image indicator on its own)."""
        n = 16
        border = '#4a4a4a' if disabled else '#8a8a8a'
        interior = '#333333' if disabled else field
        tick = '#5e7d5e' if disabled else '#86e086'
        img = tk.PhotoImage(master=self.root, width=n, height=n)
        img.put(bg, to=(0, 0, n, n))                  # blend into the dark background
        img.put(border, to=(2, 2, n - 1, n - 1))      # box border
        img.put(interior, to=(3, 3, n - 2, n - 2))    # box interior
        if checked:                                    # draw a thick check stroke
            stroke = [(3, 8), (4, 9), (5, 10), (6, 11),
                      (7, 10), (8, 9), (9, 8), (10, 7), (11, 6), (12, 5), (13, 4)]
            for x, y in stroke:
                img.put(tick, to=(x, y, x + 2, y + 2))
        return img

    def _install_dark_check(self, style, bg: str, field: str):
        """Replace clam's checkbox indicator (its checked mark looks like an X)
        with a drawn box + green tick, including muted disabled variants.
        Created once; the layout change applies only to clam, so light mode
        keeps its native vista checkbox."""
        if not getattr(self, '_dark_check', None):
            self._dark_check = {
                'off': self._make_check_img(False, bg, field),
                'on': self._make_check_img(True, bg, field),
                'off_dis': self._make_check_img(False, bg, field, disabled=True),
                'on_dis': self._make_check_img(True, bg, field, disabled=True),
            }
            try:
                c = self._dark_check
                # state specs are matched in order, most-specific first
                style.element_create('Dark.Checkbutton.indicator', 'image', c['off'],
                                     ('disabled', 'selected', c['on_dis']),
                                     ('disabled', c['off_dis']),
                                     ('selected', c['on']),
                                     padding=(0, 0, 6, 0), sticky='')
            except tk.TclError:
                pass   # already registered (theme toggled before)
        try:
            style.layout('TCheckbutton', [
                ('Checkbutton.padding', {'sticky': 'nswe', 'children': [
                    ('Dark.Checkbutton.indicator', {'side': 'left', 'sticky': ''}),
                    ('Checkbutton.focus', {'side': 'left', 'sticky': 'w', 'children': [
                        ('Checkbutton.label', {'sticky': 'nswe'})]})]})])
        except tk.TclError:
            pass

    def _scrollable(self, outer):
        """Return a padded inner frame inside a vertically-scrollable canvas that
        fills `outer` (a notebook tab). Content taller than the viewport scrolls;
        shorter content stretches to fill (so e.g. the log box still expands)."""
        canvas = tk.Canvas(outer, highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        inner = ttk.Frame(canvas, padding=12)
        win = canvas.create_window((0, 0), window=inner, anchor='nw')

        def _sync(_=None):
            cw, ch = canvas.winfo_width(), canvas.winfo_height()
            canvas.itemconfigure(win, width=cw, height=max(inner.winfo_reqheight(), ch))
            canvas.configure(scrollregion=(0, 0, cw, inner.winfo_reqheight()))
        inner.bind('<Configure>', _sync)
        canvas.bind('<Configure>', _sync)
        self._canvases.append(canvas)
        self._tab_canvases[str(outer)] = canvas
        return inner

    def _on_mousewheel(self, event):
        """A scrollable list/text under the pointer (the file list, the log, the
        rule boxes) scrolls itself first when it has hidden content; otherwise
        the wheel scrolls the whole tab."""
        delta = -1 if event.delta > 0 else 1
        w = self.root.winfo_containing(event.x_root, event.y_root)
        if isinstance(w, (tk.Listbox, tk.Text)):
            try:
                first, last = w.yview()
                if first > 0.0 or last < 1.0:        # has rows hidden above/below
                    w.yview_scroll(delta, 'units')
                    return
            except tk.TclError:
                pass
        cvs = self._tab_canvases.get(self.notebook.select())
        if cvs is not None:
            cvs.yview_scroll(delta, 'units')

    def _build_layout(self):
        header = ttk.Frame(self.root, padding=(16, 12, 16, 4))
        header.pack(fill='x')
        ttk.Label(header, text="Data Lineage Builder", style='Title.TLabel').pack(anchor='w')
        ttk.Label(header,
                  text="Turn Atlan impact reports into client-formatted lineage workbooks",
                  style='Sub.TLabel').pack(anchor='w')

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=12, pady=8)
        self._build_run_tab()
        self._build_rules_tab()
        self._build_format_tab()
        self.root.bind_all('<MouseWheel>', self._on_mousewheel)

        # status bar
        bar = ttk.Frame(self.root, padding=(16, 4, 16, 8))
        bar.pack(fill='x')
        self.progress = ttk.Progressbar(bar, mode='indeterminate', length=180)
        self.progress.pack(side='left')
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(bar, textvariable=self.status_var).pack(side='left', padx=10)
        # Packed right-to-left so the row reads [Stop | Open output folder |
        # Open workbook]. Stop is active only during a build and the Open
        # buttons only after, so they never compete and Stop isn't hit by
        # mistake.
        self.open_file_btn = ttk.Button(bar, text="Open workbook",
                                        command=self._open_file, state='disabled')
        self.open_file_btn.pack(side='right', padx=(6, 0))
        self.open_folder_btn = ttk.Button(bar, text="Open output folder",
                                          command=self._open_folder, state='disabled')
        self.open_folder_btn.pack(side='right', padx=6)
        self.stop_btn = ttk.Button(bar, text="■ Stop", command=self._stop, state='disabled')
        self.stop_btn.pack(side='right', padx=6)

        self._add_watermark()

    def _watermark_text(self) -> str:
        """Optional corner watermark, OFF by default. Enabled (e.g. "AK") via,
        in order of precedence: the LINEAGE_WATERMARK env var, a 'watermark.txt'
        bundled into a packaged build, or a 'watermark' key in the local
        settings file. Lets one codebase produce a plain or watermarked build."""
        txt = os.environ.get('LINEAGE_WATERMARK', '').strip()
        if txt:
            return txt
        if getattr(sys, 'frozen', False):
            bundled = os.path.join(getattr(sys, '_MEIPASS', ''), 'watermark.txt')
            try:
                if os.path.exists(bundled):
                    with open(bundled, encoding='utf-8') as fh:
                        return fh.read().strip()
            except OSError:
                pass
        try:
            if os.path.exists(SETTINGS_FILE):
                with open(SETTINGS_FILE, encoding='utf-8') as fh:
                    return str(json.load(fh).get('watermark', '')).strip()
        except (OSError, json.JSONDecodeError):
            pass
        return ''

    def _add_watermark(self):
        """Draw a small, muted maker's mark in the top-right corner if one is
        configured (see _watermark_text). No-op otherwise."""
        text = self._watermark_text()
        if not text:
            return
        mark = tk.Label(self.root, text=text, font=('Segoe UI', 8),
                        foreground='#C2C2C2', background=self.root.cget('background'))
        mark.place(relx=1.0, rely=0.0, x=-10, y=6, anchor='ne')
        mark.lift()
        self._watermark_lbl = mark
        self._apply_theme()   # match its colours to the current theme

    # ---- tab 1: run ----------------------------------------------------------
    def _build_run_tab(self):
        outer = ttk.Frame(self.notebook)
        self.notebook.add(outer, text='  Run  ')
        tab = self._scrollable(outer)

        topbar = ttk.Frame(tab)
        topbar.pack(fill='x')
        ttk.Checkbutton(topbar, text="🌙 Dark mode", variable=self.dark,
                        command=self._apply_theme).pack(side='right')

        files_frame = ttk.LabelFrame(tab, text="Atlan impact reports (.csv / .xlsx)",
                                     padding=8)
        files_frame.pack(fill='x', pady=(4, 0))
        list_row = ttk.Frame(files_frame)
        list_row.pack(fill='x')
        self.files_list = tk.Listbox(list_row, height=5, selectmode='extended',
                                     activestyle='dotbox')
        self._register_themed(self.files_list)
        self.files_list.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(list_row, command=self.files_list.yview)
        scroll.pack(side='left', fill='y')
        self.files_list.config(yscrollcommand=scroll.set)
        btns = ttk.Frame(list_row)
        btns.pack(side='left', fill='y', padx=(8, 0))
        ttk.Button(btns, text="Add files…", command=self._add_files).pack(fill='x')
        ttk.Button(btns, text="Remove", command=self._remove_files).pack(fill='x', pady=4)
        ttk.Button(btns, text="Clear", command=self._clear_files).pack(fill='x')

        out_frame = ttk.Frame(tab)
        out_frame.pack(fill='x', pady=(10, 0))
        ttk.Label(out_frame, text="Output workbook:").pack(side='left')
        self.output_var = tk.StringVar()
        ttk.Entry(out_frame, textvariable=self.output_var).pack(
            side='left', fill='x', expand=True, padx=8)
        ttk.Button(out_frame, text="Browse…", command=self._pick_output).pack(side='left')
        naming_hint = ttk.Label(
            tab, text="Leave blank to auto-name the output next to the first input file — "
                      "a single report after its dashboard/table, a combined workbook as "
                      "\"Data Lineage Workbook.xlsx\".",
            style='Sub.TLabel', justify='left')
        naming_hint.pack(anchor='w', fill='x')
        # wrap rather than clip at narrow widths
        tab.bind('<Configure>',
                 lambda e: naming_hint.configure(wraplength=max(280, e.width - 24)), add='+')

        self.opt_detailed = tk.BooleanVar(value=True)
        self.opt_combine = tk.BooleanVar(value=True)
        self.opt_separate = tk.BooleanVar(value=False)
        self.worker_count = tk.IntVar(value=os.cpu_count() or 4)

        opts = ttk.Frame(tab)
        opts.pack(fill='x', pady=(10, 0))
        det = ttk.LabelFrame(opts, text="Detailed sheets", padding=8)
        det.pack(side='left', fill='both', expand=True)
        ttk.Checkbutton(det, text="Include Detailed sheets", variable=self.opt_detailed
                        ).pack(anchor='w', padx=6, pady=2)
        self.sep_cb = ttk.Checkbutton(det, text="Detailed sheets in a separate file",
                                      variable=self.opt_separate)
        self.sep_cb.pack(anchor='w', padx=6, pady=2)

        comb = ttk.LabelFrame(opts, text="Combining", padding=8)
        comb.pack(side='left', fill='both', expand=True, padx=(10, 0))
        self.combine_cb = ttk.Checkbutton(comb, text="Combine into one workbook",
                                          variable=self.opt_combine)
        self.combine_cb.pack(anchor='w', padx=6, pady=2)
        self.opt_auto_workers = tk.BooleanVar(value=True)
        ttk.Checkbutton(comb, text="Auto worker count (one per report, capped at CPUs)",
                        variable=self.opt_auto_workers, command=self._toggle_workers
                        ).pack(anchor='w', padx=6, pady=2)
        wrow = ttk.Frame(comb)
        wrow.pack(anchor='w', padx=6, pady=2)
        ttk.Label(wrow, text="Worker process cap:").pack(side='left')
        self.worker_spin = ttk.Spinbox(wrow, from_=1, to=max(1, os.cpu_count() or 8),
                                       width=4, textvariable=self.worker_count)
        self.worker_spin.pack(side='left', padx=4)
        self._toggle_workers()
        self.opt_detailed.trace_add('write', self._toggle_detailed)
        self.opt_combine.trace_add('write', self._toggle_detailed)
        self._toggle_detailed()

        actions = ttk.Frame(tab)
        actions.pack(fill='x', pady=12)
        self.run_btn = ttk.Button(actions, text="▶  Build Lineage Workbook",
                                  style='Run.TButton', command=self._run)
        self.run_btn.pack(side='left')
        ttk.Button(actions, text="Save settings",
                   command=self._save_settings).pack(side='right')

        log_frame = ttk.LabelFrame(tab, text="Log", padding=4)
        log_frame.pack(fill='both', expand=True)
        self.opt_error_log = tk.BooleanVar(value=False)
        self.opt_stop_on_error = tk.BooleanVar(value=False)
        ttk.Checkbutton(log_frame, variable=self.opt_error_log,
                        text="Write an error report (.txt) listing any files that were skipped"
                        ).pack(anchor='w', padx=2, pady=(0, 2))
        ttk.Checkbutton(log_frame, variable=self.opt_stop_on_error,
                        text="Stop the whole build if any file fails (instead of skipping it)"
                        ).pack(anchor='w', padx=2, pady=(0, 2))
        log_body = ttk.Frame(log_frame)
        log_body.pack(fill='both', expand=True)
        self.log_text = tk.Text(log_body, height=10, state='disabled',
                                font=('Consolas', 9), background='#FAF7FC')
        self._register_themed(self.log_text, light_bg='#FAF7FC')
        self.log_text.pack(side='left', fill='both', expand=True)
        log_scroll = ttk.Scrollbar(log_body, command=self.log_text.yview)
        log_scroll.pack(side='left', fill='y')
        self.log_text.config(yscrollcommand=log_scroll.set)

    def _toggle_detailed(self, *_):
        """'Detailed in a separate file' only applies to a combined workbook
        that includes detailed sheets."""
        ok = self.opt_detailed.get() and self.opt_combine.get()
        self.sep_cb.configure(state='normal' if ok else 'disabled')

    def _toggle_workers(self, *_):
        """The worker-cap spinbox is only editable when Auto is off."""
        self.worker_spin.configure(
            state='disabled' if self.opt_auto_workers.get() else 'normal')

    # ---- tab 2: pruning rules -------------------------------------------------
    def _build_rules_tab(self):
        outer = ttk.Frame(self.notebook)
        self.notebook.add(outer, text='  Pruning Rules  ')
        tab = self._scrollable(outer)

        # --- per-sheet toggles at the top ---
        self.drop_summary = tk.BooleanVar(value=True)
        self.drop_detailed = tk.BooleanVar(value=True)
        self.apply_summary = tk.BooleanVar(value=True)
        self.apply_detailed = tk.BooleanVar(value=False)
        toggles = ttk.Frame(tab)
        toggles.pack(fill='x', pady=(0, 8))
        for r, (label, sv, dv) in enumerate([
                ("Ignore queries with no upstream lineage:", self.drop_summary, self.drop_detailed),
                ("Apply the pruning rules below:", self.apply_summary, self.apply_detailed)]):
            ttk.Label(toggles, text=label).grid(row=r, column=0, sticky='w', pady=2)
            ttk.Checkbutton(toggles, text="Summary", variable=sv).grid(
                row=r, column=1, sticky='w', padx=(12, 8))
            ttk.Checkbutton(toggles, text="Detailed", variable=dv).grid(
                row=r, column=2, sticky='w')

        help_lbl = ttk.Label(
            tab,
            text="Each box takes one glob pattern per line, matched against "
                 "DATABASE.SCHEMA.OBJECT (e.g.  *.DATAWAREHOUSE.*,  PROD_*.*.BRANCH*,  "
                 "*.CRM*.* ). Fewer than three segments auto-fill with * "
                 "(TEMP_DB = TEMP_DB.*.*). Matching ignores case unless you quote "
                 'the entry, e.g.  "PROD_DATALAKE.LAWPROD.attrep_changes*"  '
                 "(for case-sensitive Snowflake quoted identifiers).",
            style='Sub.TLabel', justify='left')
        help_lbl.pack(anchor='w', fill='x', pady=(0, 8))
        # Wrap to the actual available width so the text never clips at min width.
        tab.bind('<Configure>',
                 lambda e: help_lbl.configure(wraplength=max(280, e.width - 24)))

        grid = ttk.Frame(tab)
        grid.pack(fill='both', expand=True)
        for i in (0, 1):
            grid.columnconfigure(i, weight=1, uniform='rules')
            grid.rowconfigure(i, weight=1)

        def make_section(r, c, title, info, level_var=None):
            f = ttk.LabelFrame(grid, text=title, padding=8)
            f.grid(row=r, column=c, sticky='nsew',
                   padx=(0, 5) if c == 0 else (5, 0),
                   pady=(0, 5) if r == 0 else (5, 0))
            if level_var is None:
                ttk.Label(f, text=info, style='Sub.TLabel',
                          wraplength=380, justify='left').pack(anchor='w')
            else:
                row = ttk.Frame(f)
                row.pack(anchor='w', fill='x')
                ttk.Label(row, text=info, style='Sub.TLabel').pack(side='left')
                ttk.Spinbox(row, from_=0, to=99, width=4,
                            textvariable=level_var).pack(side='left', padx=4)
            txt = tk.Text(f, height=6, font=('Consolas', 9))
            self._register_themed(txt)
            txt.pack(fill='both', expand=True, pady=(4, 0))
            return txt

        self.rule_tbl_limit = tk.IntVar(value=8)
        self.rule_vw_limit = tk.IntVar(value=8)
        self.ex_text = make_section(
            0, 0, "Exclude entirely",
            "Hidden from the output and not expanded.")
        self.blk_text = make_section(
            0, 1, "Block expansion",
            "Shown in the output, but upstream is not expanded.")
        self.tbl_text = make_section(
            1, 0, "Stop expanding tables",
            "Matching TABLES shown but not expanded beyond level",
            self.rule_tbl_limit)
        self.vw_text = make_section(
            1, 1, "Stop expanding views",
            "Matching VIEWS shown but not expanded beyond level",
            self.rule_vw_limit)

    # ---- tab 3: formatting & sources -----------------------------------------
    def _build_format_tab(self):
        outer = ttk.Frame(self.notebook)
        self.notebook.add(outer, text='  Formatting & Sources  ')
        tab = self._scrollable(outer)
        self._fmt_lockable = []   # widgets disabled while "Unformatted" is on

        self.opt_unformatted = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            tab, text="Unformatted raw export  (a plain wide dump — disables every "
                      "styling option below except the source labels)",
            variable=self.opt_unformatted).pack(anchor='w')

        fmt = ttk.LabelFrame(tab, text="Workbook appearance", padding=8)
        fmt.pack(fill='x', pady=(8, 0))
        self.fmt_links = tk.BooleanVar(value=True)
        self.fmt_merge = tk.BooleanVar(value=True)
        self.fmt_pad = tk.BooleanVar(value=True)
        self.fmt_freeze = tk.BooleanVar(value=False)
        for col, (text, var) in enumerate([
                ("Hyperlink duplicates to their expansion", self.fmt_links),
                ("Merge Query cells over their block", self.fmt_merge),
                ("Pad empty cells like the client file", self.fmt_pad),
                ("Freeze header rows & Query column", self.fmt_freeze)]):
            cb = ttk.Checkbutton(fmt, text=text, variable=var)
            cb.grid(row=col // 2, column=col % 2, sticky='w', padx=6, pady=2)
            self._fmt_lockable.append(cb)
        colors = ttk.Frame(fmt)
        colors.grid(row=2, column=0, columnspan=2, sticky='w', pady=(6, 0))
        ttk.Label(colors, text="Header fill:").pack(side='left')
        self.fmt_header_fill = tk.StringVar(value=engine.EXCEL_FORMAT['header_fill'])
        self.fmt_sub_fill = tk.StringVar(value=engine.EXCEL_FORMAT['subheader_fill'])
        self.fmt_font = tk.StringVar(value=engine.EXCEL_FORMAT['font_name'])
        e1 = ttk.Entry(colors, textvariable=self.fmt_header_fill, width=8)
        e1.pack(side='left', padx=4)
        ttk.Label(colors, text="Sub-header fill:").pack(side='left', padx=(10, 0))
        e2 = ttk.Entry(colors, textvariable=self.fmt_sub_fill, width=8)
        e2.pack(side='left', padx=4)
        ttk.Label(colors, text="Font:").pack(side='left', padx=(10, 0))
        e3 = ttk.Entry(colors, textvariable=self.fmt_font, width=16)
        e3.pack(side='left', padx=4)
        self._fmt_lockable += [e1, e2, e3]

        src = ttk.LabelFrame(
            tab, text='"List of sources" labels — glob = label, one per line, first '
                      'match wins (e.g.  *.CRM_MSCRM.* = CRM ; quote for case-sensitive)',
            padding=8)
        src.pack(fill='both', expand=True, pady=(10, 0))
        self.src_text = tk.Text(src, height=8, font=('Consolas', 9))
        self._register_themed(self.src_text)
        self.src_text.pack(fill='both', expand=True)
        fb_row = ttk.Frame(src)
        fb_row.pack(anchor='w', pady=(6, 0))
        ttk.Label(fb_row, text="When no label matches, use:").pack(side='left')
        self.src_fallback = tk.StringVar(value=engine.SOURCE_LABEL_FALLBACK)
        ttk.Combobox(fb_row, textvariable=self.src_fallback, width=10,
                     values=('schema', 'db', 'blank'), state='readonly'
                     ).pack(side='left', padx=6)

        self.opt_unformatted.trace_add('write', self._toggle_unformatted)
        self._toggle_unformatted()

    def _toggle_unformatted(self, *_):
        """Grey out the styling controls when Unformatted export is selected
        (source labels stay editable — they apply to the raw dump too)."""
        state = 'disabled' if self.opt_unformatted.get() else 'normal'
        for w in self._fmt_lockable:
            try:
                w.configure(state=state)
            except tk.TclError:
                pass

    # ------------------------------------------------------------- actions --
    def _refresh_files(self):
        """Redraw the listbox with 1-based numbers showing processing order."""
        self.files_list.delete(0, 'end')
        for i, path in enumerate(self._files, 1):
            self.files_list.insert('end', f'{i:>2}.  {path}')
        self._update_combine_state()

    def _update_combine_state(self):
        """Combining is only meaningful with two or more reports; grey it out
        otherwise (the engine names a single report after its dashboard with no
        index sheet regardless of this toggle)."""
        if hasattr(self, 'combine_cb'):
            self.combine_cb.configure(state='normal' if len(self._files) > 1 else 'disabled')

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select Atlan impact reports",
            filetypes=[("Excel / CSV", "*.xlsx *.csv"), ("All files", "*.*")])
        if not paths:
            return
        norm = lambda p: os.path.normcase(os.path.abspath(p))
        # de-dupe the incoming selection (preserving its order) ...
        seen, incoming = set(), []
        for p in paths:
            k = norm(p)
            if k not in seen:
                seen.add(k)
                incoming.append(p)
        # ... then drop any existing entries for re-added files and append the
        # new selection, so the latest order wins and there are no duplicates.
        self._files = [f for f in self._files if norm(f) not in seen] + incoming
        self._refresh_files()

    def _remove_files(self):
        for idx in sorted(self.files_list.curselection(), reverse=True):
            del self._files[idx]
        self._refresh_files()

    def _clear_files(self):
        self._files.clear()
        self._refresh_files()

    def _pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Output workbook", defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.output_var.set(path)

    def _collect_rules(self) -> dict:
        def lines(widget):
            return [ln.strip() for ln in widget.get('1.0', 'end').splitlines() if ln.strip()]
        return {
            'exclude_patterns': lines(self.ex_text),
            'block_patterns': lines(self.blk_text),
            'table_patterns': lines(self.tbl_text),
            'table_level': self.rule_tbl_limit.get(),
            'view_patterns': lines(self.vw_text),
            'view_level': self.rule_vw_limit.get(),
            'drop_no_lineage_summary': self.drop_summary.get(),
            'drop_no_lineage_detailed': self.drop_detailed.get(),
            'apply_to_summary': self.apply_summary.get(),
            'apply_to_detailed': self.apply_detailed.get(),
        }

    def _collect_format(self) -> dict:
        fmt = dict(engine.EXCEL_FORMAT)
        fmt.update({
            'hyperlink_duplicates': self.fmt_links.get(),
            'merge_query_blocks': self.fmt_merge.get(),
            'pad_empty_cells': self.fmt_pad.get(),
            'freeze_panes': self.fmt_freeze.get(),
            'header_fill': self.fmt_header_fill.get().strip() or '7030A0',
            'subheader_fill': self.fmt_sub_fill.get().strip() or 'F2CEEF',
            'font_name': self.fmt_font.get().strip() or 'Messina Sans',
        })
        return fmt

    def _collect_source_labels(self) -> dict:
        labels = {}
        for line in self.src_text.get('1.0', 'end').splitlines():
            if '=' in line:
                prefix, _, label = line.partition('=')
                if prefix.strip():
                    labels[prefix.strip()] = label.strip()
        return labels

    def _ask_overwrite(self, existing: list) -> str:
        """Modal 3-way dialog when output files already exist.
        Returns 'overwrite', 'suffix' (save a numbered copy), or 'cancel'."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Output file already exists")
        dlg.transient(self.root)
        dlg.resizable(False, False)
        dlg.configure(background=self.root.cget('background'))
        shown = "\n".join("     • " + os.path.basename(p) for p in existing[:12])
        more = "" if len(existing) <= 12 else f"\n     …and {len(existing) - 12} more"
        msg = ("These output file(s) already exist:\n\n" + shown + more +
               "\n\n• Overwrite — replace them.\n"
               "• Save a copy — keep the originals; write \"name (1).xlsx\", etc.\n"
               "• Cancel — don't build.")
        ttk.Label(dlg, text=msg, justify='left', wraplength=440).pack(
            padx=16, pady=(16, 10), anchor='w')
        choice = {'v': 'cancel'}

        def pick(v):
            choice['v'] = v
            dlg.destroy()

        row = ttk.Frame(dlg)
        row.pack(padx=16, pady=(0, 14), fill='x')
        # packed right-to-left: [Overwrite] is leftmost, Cancel rightmost
        ttk.Button(row, text="Cancel", command=lambda: pick('cancel')).pack(side='right')
        ttk.Button(row, text="Save a copy", command=lambda: pick('suffix')).pack(
            side='right', padx=6)
        ttk.Button(row, text="Overwrite", command=lambda: pick('overwrite')).pack(side='right')
        dlg.bind('<Escape>', lambda e: pick('cancel'))
        dlg.protocol("WM_DELETE_WINDOW", lambda: pick('cancel'))
        dlg.grab_set()
        self.root.wait_window(dlg)
        return choice['v']

    def _run(self):
        files = list(self._files)
        if not files:
            messagebox.showwarning("No input", "Add at least one Atlan impact report first.")
            return
        missing = [f for f in files if not os.path.exists(f)]
        if missing:
            messagebox.showerror("File not found", "\n".join(missing))
            return
        # confirm before overwriting any existing output (on the UI thread,
        # before the build starts, so there's no cross-thread dialog)
        avoid_overwrite = False
        try:
            planned = engine._planned_outputs(
                files, self.output_var.get().strip() or None,
                self.opt_combine.get(), self.opt_unformatted.get(), self.opt_separate.get())
        except Exception:
            planned = []
        existing = [p for p in planned if os.path.exists(p)]
        if existing:
            choice = self._ask_overwrite(existing)
            if choice == 'cancel':
                self.status_var.set("Cancelled")
                return
            avoid_overwrite = (choice == 'suffix')
        self._save_settings(silent=True)
        self.cancel_event = threading.Event()
        self.run_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.open_file_btn.config(state='disabled')
        self.open_folder_btn.config(state='disabled')
        self.progress.start(12)
        self.status_var.set("Building…")
        self._log_clear()

        kwargs = dict(
            input_files=files,
            output_path=self.output_var.get().strip() or None,
            rules=self._collect_rules(),
            fmt=self._collect_format(),
            source_labels=self._collect_source_labels(),
            source_fallback=self.src_fallback.get(),
            include_detailed=self.opt_detailed.get(),
            separate_detailed=self.opt_separate.get(),
            unformatted=self.opt_unformatted.get(),
            combine=self.opt_combine.get(),
            max_workers=None if self.opt_auto_workers.get() else self.worker_count.get(),
            write_error_log=self.opt_error_log.get(),
            stop_on_error=self.opt_stop_on_error.get(),
            avoid_overwrite=avoid_overwrite,
            cancel=self.cancel_event,
        )
        self.worker = threading.Thread(target=self._worker_main, args=(kwargs,), daemon=True)
        self.worker.start()

    def _stop(self):
        """Signal the running build to stop; it unwinds and terminates its
        worker processes, then the poller resets the buttons."""
        if self.cancel_event is not None:
            self.cancel_event.set()
        self.stop_btn.config(state='disabled')
        self.status_var.set("Stopping…")
        self._log_line("\nStopping — terminating workers…", 'warn')

    def _worker_main(self, kwargs):
        handler = QueueLogHandler(self.log_queue)
        handler.setFormatter(logging.Formatter('%(message)s'))
        engine.log.addHandler(handler)
        engine.log.setLevel(logging.INFO)
        try:
            written = engine.build_workbooks(**kwargs)
            self.written_files = written
            self.log_queue.put(('DONE', written))
        except engine.BuildCancelled:
            self.log_queue.put(('CANCELLED', None))
        except Exception as exc:
            self.log_queue.put(('ERROR', str(exc)))
        finally:
            engine.log.removeHandler(handler)

    # ------------------------------------------------------------- logging --
    def _poll_log_queue(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                if isinstance(item, tuple) and item and item[0] == 'LOG':
                    _, levelno, msg = item
                    self._log_line(msg, self._tag_for(levelno, msg))
                elif isinstance(item, tuple):
                    kind, payload = item
                    self.progress.stop()
                    self.run_btn.config(state='normal')
                    self.stop_btn.config(state='disabled')
                    if kind == 'DONE':
                        self.status_var.set("Done — " + os.path.basename(payload[0]))
                        self._log_line("\nFinished. Output:", 'ok')
                        for path in payload:
                            self._log_line(f"  {path}", 'ok')
                        # one file -> "Open workbook" makes sense; multiple files
                        # (separate-files mode) -> only "Open output folder".
                        self.open_file_btn.config(
                            state='normal' if len(payload) == 1 else 'disabled')
                        self.open_folder_btn.config(state='normal')
                    elif kind == 'CANCELLED':
                        self.status_var.set("Stopped")
                        self._log_line("\nStopped. No output written.", 'warn')
                    else:
                        self.status_var.set("Failed")
                        self._log_line(f"\nERROR: {payload}", 'err')
                        messagebox.showerror("Build failed", payload)
                else:
                    self._log_line(item)
        except queue.Empty:
            pass
        self.root.after(120, self._poll_log_queue)

    @staticmethod
    def _tag_for(levelno: int, msg: str) -> str:
        if levelno >= logging.ERROR:
            return 'err'
        if levelno >= logging.WARNING:
            return 'warn'
        s = msg.lstrip()
        if s.startswith(('Saved', 'Finished', 'Done')) or 'rendered ' in s:
            return 'ok'
        return 'info'

    def _log_line(self, text: str, tag: str = 'info'):
        self.log_text.config(state='normal')
        self.log_text.insert('end', text + '\n', tag)
        self.log_text.see('end')
        self.log_text.config(state='disabled')

    def _log_clear(self):
        self.log_text.config(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.config(state='disabled')

    def _open_file(self):
        if self.written_files:
            os.startfile(self.written_files[0])

    def _open_folder(self):
        if self.written_files:
            subprocess.Popen(['explorer', '/select,', self.written_files[0]])

    # ------------------------------------------------------------ settings --
    def _save_settings(self, silent: bool = False):
        cfg = {
            'trivial_rules': self._collect_rules(),
            'excel_format': self._collect_format(),
            'source_labels': self._collect_source_labels(),
            'source_label_fallback': self.src_fallback.get(),
            'ui': {
                'files': list(self._files),
                'output': self.output_var.get(),
                'include_detailed': self.opt_detailed.get(),
                'separate_detailed': self.opt_separate.get(),
                'unformatted': self.opt_unformatted.get(),
                'combine': self.opt_combine.get(),
                'auto_workers': self.opt_auto_workers.get(),
                'workers': self.worker_count.get(),
                'error_log': self.opt_error_log.get(),
                'stop_on_error': self.opt_stop_on_error.get(),
                'dark': self.dark.get(),
            },
        }
        try:
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as fh:
                json.dump(cfg, fh, indent=2)
            if not silent:
                self.status_var.set(f"Settings saved to {os.path.basename(SETTINGS_FILE)}")
        except OSError as exc:
            if not silent:
                messagebox.showerror("Could not save settings", str(exc))

    def _load_settings(self):
        rules = dict(engine.TRIVIAL_RULES)
        labels = dict(engine.SOURCE_LABELS)
        ui = {}
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, encoding='utf-8') as fh:
                    cfg = json.load(fh)
                rules.update(cfg.get('trivial_rules', {}))
                labels = cfg.get('source_labels', labels)
                fmt = cfg.get('excel_format', {})
                self.fmt_links.set(fmt.get('hyperlink_duplicates', True))
                self.fmt_merge.set(fmt.get('merge_query_blocks', True))
                self.fmt_pad.set(fmt.get('pad_empty_cells', True))
                self.fmt_freeze.set(fmt.get('freeze_panes', False))
                self.fmt_header_fill.set(fmt.get('header_fill', '7030A0'))
                self.fmt_sub_fill.set(fmt.get('subheader_fill', 'F2CEEF'))
                self.fmt_font.set(fmt.get('font_name', 'Messina Sans'))
                self.src_fallback.set(cfg.get('source_label_fallback', 'schema'))
                ui = cfg.get('ui', {})
            except (OSError, json.JSONDecodeError):
                pass

        self.rule_tbl_limit.set(rules.get('table_level', 8))
        self.rule_vw_limit.set(rules.get('view_level', 8))
        self.ex_text.insert('1.0', '\n'.join(rules.get('exclude_patterns', [])))
        self.blk_text.insert('1.0', '\n'.join(rules.get('block_patterns', [])))
        self.tbl_text.insert('1.0', '\n'.join(rules.get('table_patterns', [])))
        self.vw_text.insert('1.0', '\n'.join(rules.get('view_patterns', [])))
        self.src_text.insert('1.0', '\n'.join(f'{k} = {v}' for k, v in labels.items()))
        self.drop_summary.set(rules.get('drop_no_lineage_summary', True))
        self.drop_detailed.set(rules.get('drop_no_lineage_detailed', True))
        self.apply_summary.set(rules.get('apply_to_summary', True))
        self.apply_detailed.set(rules.get('apply_to_detailed', False))

        self._files = [p for p in ui.get('files', []) if os.path.exists(p)]
        self._refresh_files()
        self.output_var.set(ui.get('output', ''))
        self.opt_detailed.set(ui.get('include_detailed', True))
        self.opt_separate.set(ui.get('separate_detailed', False))
        self.opt_unformatted.set(ui.get('unformatted', False))
        self.opt_combine.set(ui.get('combine', True))
        self.opt_auto_workers.set(ui.get('auto_workers', True))
        self.worker_count.set(ui.get('workers', os.cpu_count() or 4))
        self._toggle_workers()
        self.opt_error_log.set(ui.get('error_log', False))
        self.opt_stop_on_error.set(ui.get('stop_on_error', False))
        self.dark.set(ui.get('dark', False))


def run():
    root = tk.Tk()
    LineageApp(root)
    root.mainloop()


if __name__ == '__main__':
    run()
