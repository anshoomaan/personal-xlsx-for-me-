"""
mesh_table.py  –  Task Mesh Personal Database
===============================================
A Tkinter-based personal task manager with a multi-column spreadsheet UI,
per-cell styling, section titles, and plain-text export.

Run:
    python mesh_table.py

Keyboard shortcuts
------------------
  Tab            – begin chord (release within 350 ms to trigger an action):
      Tab  alone         → insert a section title ribbon ABOVE current row
      Tab → C            → insert a blank column AFTER current column
      Tab → R            → insert a blank data row AFTER current row
  Shift+Tab      – also inserts a blank data row AFTER current row
  Return         – move focus down to the next data row (same column)

Column tricks
-------------
  Shift+Click on any column header  – insert a new column AFTER that column
  Right-Click  on any column header – same as Shift+Click
  Click  [+]  button in header      – append a column at the far right

File format
-----------
  Files are saved / loaded as JSON (extension .mesh).
  Save is atomic (temp-file + rename) with an automatic .bak backup.
  The loader validates and repairs every field so older files always open.
  The plain-text export ("Export .txt") remains available separately.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import tempfile
import tkinter as tk
import tkinter.font as tkfont
from dataclasses import asdict, dataclass, field
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk
from typing import Any

# ---------------------------------------------------------------------------
# Constants & Configuration
# ---------------------------------------------------------------------------

SCRIPT_DIR: str = os.path.dirname(os.path.abspath(__file__))

DEFAULT_COLUMNS: int = 10
DEFAULT_DATA_ROWS: int = 28
EXPORT_LINE_WIDTH: int = 72
EXPORT_COL_PAD: int = 15

WINDOW_TITLE: str = "Task Mesh – Personal Database"
WINDOW_GEOMETRY: str = "1350x760"
WINDOW_MIN_SIZE: tuple[int, int] = (900, 500)

# Column widths (pixels)
COL0_W: int = 80
DATA_W: int = 130
ADDCOL_W: int = 34

# Palette
TOOLBAR_BG: str = "#1a1a1a"
HEADER_BG: str = "#d0d0d0"
HEADER_FG: str = "#222222"
TITLE_BG: str = "#0d0d0d"
TITLE_FG: str = "#ff3333"
ROW_BG_A: str = "#ffffff"
ROW_BG_B: str = "#f7f7f7"
SERIAL_FG: str = "#aaaaaa"
SEPARATOR_BG: str = "#dedede"

DEFAULT_FONT_FAMILY: str = "Consolas"
DEFAULT_FONT_SIZE: str = "10"
DEFAULT_TEXT_COLOR: str = "#111111"
DEFAULT_FILL_COLOR: str = "#ffffff"
DEFAULT_ALIGNMENT: str = "left"

FONT_SIZES: tuple[str, ...] = (
    "8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "32"
)

ALIGN_SYMBOLS: tuple[tuple[str, str], ...] = (
    ("≡L", "left"),
    ("≡C", "center"),
    ("≡R", "right"),
)

# File format version — increment only on breaking schema changes
SAVE_VERSION: int = 2


# ---------------------------------------------------------------------------
# Data Model
# ---------------------------------------------------------------------------

@dataclass
class CellStyle:
    """Stores per-cell visual style attributes."""
    family: str = DEFAULT_FONT_FAMILY
    size: int = int(DEFAULT_FONT_SIZE)
    bold: bool = False
    italic: bool = False
    fg: str = DEFAULT_TEXT_COLOR
    bg: str = DEFAULT_FILL_COLOR
    justify: str = DEFAULT_ALIGNMENT


@dataclass
class RowItem:
    """Base representation of a single row in the table model."""
    row_type: str  # "title" | "header" | "data"


@dataclass
class TitleRow(RowItem):
    text: str = "title"
    has_header: bool = False
    row_type: str = field(default="title", init=False)


@dataclass
class HeaderRow(RowItem):
    values: list[str] = field(default_factory=list)
    row_type: str = field(default="header", init=False)


@dataclass
class DataRow(RowItem):
    values: list[str] = field(default_factory=list)
    row_type: str = field(default="data", init=False)


def _make_default_items(num_cols: int) -> list[RowItem]:
    """Return the initial table state: one header row + data rows."""
    items: list[RowItem] = [
        HeaderRow(values=[f"Column {i + 1}" for i in range(num_cols)])
    ]
    items += [DataRow(values=[""] * num_cols) for _ in range(DEFAULT_DATA_ROWS)]
    return items


def _compute_serials(items: list[RowItem]) -> dict[int, str | None]:
    """
    Assign sequential 3-digit serial numbers to data rows.
    Counters reset after every title/header row.
    """
    serials: dict[int, str | None] = {}
    counter = 1
    for idx, item in enumerate(items):
        if item.row_type in ("header", "title"):
            counter = 1
            serials[idx] = None
        else:
            serials[idx] = f"{counter:03d}"
            counter += 1
    return serials


# ---------------------------------------------------------------------------
# Serialisation helpers
# ---------------------------------------------------------------------------

# Known CellStyle field names and their defaults – used for schema repair
_CELL_STYLE_DEFAULTS: dict[str, Any] = {
    "family": DEFAULT_FONT_FAMILY,
    "size": int(DEFAULT_FONT_SIZE),
    "bold": False,
    "italic": False,
    "fg": DEFAULT_TEXT_COLOR,
    "bg": DEFAULT_FILL_COLOR,
    "justify": DEFAULT_ALIGNMENT,
}

_VALID_HEX_COLOR = re.compile(r"^#[0-9a-fA-F]{3}(?:[0-9a-fA-F]{3})?$")
_VALID_JUSTIFY = {"left", "center", "right"}


def _repair_cell_style(raw: dict) -> CellStyle:
    """
    Build a CellStyle from *raw*, filling in defaults for missing or
    invalid fields so a corrupt or old-format style never crashes load.
    """
    d = dict(_CELL_STYLE_DEFAULTS)  # start from safe defaults
    if isinstance(raw, dict):
        if isinstance(raw.get("family"), str) and raw["family"].strip():
            d["family"] = raw["family"]
        try:
            sz = int(raw.get("size", d["size"]))
            d["size"] = max(6, min(sz, 144))
        except (TypeError, ValueError):
            pass
        d["bold"] = bool(raw.get("bold", False))
        d["italic"] = bool(raw.get("italic", False))
        for field_name in ("fg", "bg"):
            val = raw.get(field_name, d[field_name])
            if isinstance(val, str) and _VALID_HEX_COLOR.match(val):
                d[field_name] = val
        justify = raw.get("justify", d["justify"])
        if justify in _VALID_JUSTIFY:
            d["justify"] = justify
    return CellStyle(**d)


def _content_hash(payload: dict) -> str:
    """SHA-256 of the JSON body (without the 'checksum' key itself)."""
    body = {k: v for k, v in payload.items() if k != "checksum"}
    raw = json.dumps(body, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(raw.encode()).hexdigest()


def _serialize(
    items: list[RowItem],
    cell_styles: dict[tuple[int, int], CellStyle],
    num_cols: int,
) -> dict:
    """Convert the full table state to a plain-Python dict (JSON-safe)."""
    rows_out: list[dict] = []
    for item in items:
        if isinstance(item, TitleRow):
            rows_out.append(
                {"type": "title", "text": item.text, "has_header": item.has_header}
            )
        elif isinstance(item, HeaderRow):
            rows_out.append({"type": "header", "values": list(item.values)})
        elif isinstance(item, DataRow):
            rows_out.append({"type": "data", "values": list(item.values)})

    styles_out: dict[str, dict] = {}
    for (r, c), style in cell_styles.items():
        styles_out[f"{r},{c}"] = asdict(style)

    payload = {
        "version": SAVE_VERSION,
        "num_cols": num_cols,
        "rows": rows_out,
        "cell_styles": styles_out,
    }
    payload["checksum"] = _content_hash(payload)
    return payload


def _deserialize(data: dict) -> tuple[list[RowItem], dict[tuple[int, int], CellStyle], int]:
    """
    Reconstruct items, cell_styles, and num_cols from a serialised dict.

    Safety guarantees
    -----------------
    - Missing / wrong-type fields are replaced with safe defaults.
    - Checksum mismatch raises a ValueError with a descriptive message
      (caller decides whether to warn or abort).
    - Each CellStyle is run through _repair_cell_style().
    - Every row's values list is normalised to exactly num_cols entries.
    - Unknown row types are silently skipped.
    """
    # ── Checksum ──────────────────────────────────────────────────────
    stored = data.get("checksum")
    if stored is not None:
        expected = _content_hash(data)
        if stored != expected:
            raise ValueError(
                "Checksum mismatch – the file may be corrupted or hand-edited.\n"
                "Load anyway? (data will be repaired where possible)"
            )

    # ── num_cols ──────────────────────────────────────────────────────
    try:
        num_cols = max(1, int(data.get("num_cols", DEFAULT_COLUMNS)))
    except (TypeError, ValueError):
        num_cols = DEFAULT_COLUMNS

    # ── rows ──────────────────────────────────────────────────────────
    items: list[RowItem] = []
    for row in data.get("rows", []):
        if not isinstance(row, dict):
            continue
        rtype = row.get("type")
        if rtype == "title":
            text = str(row.get("text", "")) if row.get("text") is not None else ""
            has_hdr = bool(row.get("has_header", False))
            items.append(TitleRow(text=text, has_header=has_hdr))
        elif rtype == "header":
            raw_vals = row.get("values", [])
            vals = [str(v) for v in raw_vals] if isinstance(raw_vals, list) else []
            # Pad / trim to num_cols
            while len(vals) < num_cols:
                vals.append(f"Column {len(vals) + 1}")
            vals = vals[:num_cols]
            items.append(HeaderRow(values=vals))
        elif rtype == "data":
            raw_vals = row.get("values", [])
            vals = [str(v) for v in raw_vals] if isinstance(raw_vals, list) else []
            while len(vals) < num_cols:
                vals.append("")
            vals = vals[:num_cols]
            items.append(DataRow(values=vals))
        # else: unknown row type → skip silently

    # ── cell_styles ───────────────────────────────────────────────────
    cell_styles: dict[tuple[int, int], CellStyle] = {}
    for key, sd in data.get("cell_styles", {}).items():
        try:
            r_str, c_str = str(key).split(",")
            r, c = int(r_str), int(c_str)
            cell_styles[(r, c)] = _repair_cell_style(sd)
        except (ValueError, AttributeError):
            continue  # malformed key → skip

    return items, cell_styles, num_cols


def _atomic_write(path: str, text: str) -> None:
    """
    Write *text* to *path* atomically:
      1. Write to a sibling temp file.
      2. If an existing file is present, copy it to <path>.bak first.
      3. os.replace() the temp file into place (atomic on POSIX & Windows).
    """
    dir_name = os.path.dirname(path) or "."
    # Backup existing file
    if os.path.exists(path):
        shutil.copy2(path, path + ".bak")
    # Write to temp then rename
    fd, tmp = tempfile.mkstemp(dir=dir_name, suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(text)
        os.replace(tmp, path)
    except Exception:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise



# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------

class MeshTable:
    """
    Main controller and view for the Task Mesh personal database.

    Responsibilities:
      - Build and manage the Tkinter UI (toolbar + scrollable grid).
      - Maintain the list-of-RowItem data model.
      - Apply per-cell styling (font, colour, alignment).
      - Handle keyboard shortcuts and user mutations (add/remove rows & cols).
      - Save/load the full table state as JSON (.mesh files).
      - Export the table to a formatted plain-text file.
    """

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self._configure_window()

        self.num_cols: int = DEFAULT_COLUMNS
        self.items: list[RowItem] = _make_default_items(self.num_cols)

        # {(row_index, col_index): CellStyle}
        self.cell_styles: dict[tuple[int, int], CellStyle] = {}

        # Toolbar state (tk variables bound to widgets)
        self.fnt_family = tk.StringVar(value=DEFAULT_FONT_FAMILY)
        self.fnt_size = tk.StringVar(value=DEFAULT_FONT_SIZE)
        self.bold_var = tk.BooleanVar(value=False)
        self.italic_var = tk.BooleanVar(value=False)
        self.txt_color: str = DEFAULT_TEXT_COLOR
        self.fill_color: str = DEFAULT_FILL_COLOR
        self._align = tk.StringVar(value=DEFAULT_ALIGNMENT)
        self._align_btns: dict[str, tk.Button] = {}

        # Focus tracking
        self.focused_item: int | None = None
        self.focused_col: int = 0
        self.cell_widgets: dict[tuple[int, int], tk.Entry] = {}

        # Tab-chord state: Tab alone = title, Tab→C = col after, Tab→R = row after
        self._chord_pending: bool = False
        self._chord_after_id: str | None = None   # tk after() handle

        # Preview labels updated when colour pickers run
        self._pen_prev: tk.Label
        self._fill_prev: tk.Label

        self._build_toolbar()
        self._build_table_area()
        self._render()

    # ------------------------------------------------------------------
    # Window setup
    # ------------------------------------------------------------------

    def _configure_window(self) -> None:
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_GEOMETRY)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=TOOLBAR_BG)

    # ------------------------------------------------------------------
    # Data helpers
    # ------------------------------------------------------------------

    def _flush_all(self) -> None:
        """
        Persist every live Entry widget's current value back to self.items.
        Must be called before any structural mutation to avoid data loss.
        """
        for (row_idx, col), widget in list(self.cell_widgets.items()):
            if not isinstance(widget, tk.Entry):
                continue
            try:
                value = widget.get()
            except tk.TclError:
                continue

            item = self.items[row_idx]
            if isinstance(item, TitleRow):
                item.text = value
            elif isinstance(item, HeaderRow):
                self._ensure_col(item.values, col, f"Column {col + 1}")
                item.values[col] = value
            elif isinstance(item, DataRow):
                self._ensure_col(item.values, col, "")
                item.values[col] = value

    @staticmethod
    def _ensure_col(values: list[str], col: int, fill: str) -> None:
        """Extend *values* in-place so index *col* is valid."""
        while len(values) <= col:
            values.append(fill)

    def _get_row_bg(self, row_idx: int) -> str:
        return ROW_BG_A if row_idx % 2 == 0 else ROW_BG_B

    # ------------------------------------------------------------------
    # Cell style helpers
    # ------------------------------------------------------------------

    def _resolve_cell_style(self, row_idx: int, col: int) -> CellStyle:
        """
        Return the CellStyle for (row_idx, col), falling back to the
        current toolbar state when no explicit style has been saved.
        """
        saved = self.cell_styles.get((row_idx, col))
        if saved:
            return saved
        return CellStyle(
            family=self.fnt_family.get(),
            size=int(self.fnt_size.get()),
            bold=self.bold_var.get(),
            italic=self.italic_var.get(),
            fg=self.txt_color,
            bg=self.fill_color,
            justify=self._align.get(),
        )

    def _cell_font(self, row_idx: int, col: int) -> tuple[str, int, str]:
        s = self._resolve_cell_style(row_idx, col)
        weight = "bold" if s.bold else ""
        slant = "italic" if s.italic else ""
        style = " ".join(filter(None, [weight, slant])) or "normal"
        return (s.family, s.size, style)

    def _cell_colors(self, row_idx: int, col: int, default_bg: str) -> tuple[str, str]:
        s = self._resolve_cell_style(row_idx, col)
        bg = s.bg if s.bg not in (ROW_BG_A, ROW_BG_B) else default_bg
        return s.fg, bg

    def _cell_align(self, row_idx: int, col: int) -> str:
        return self._resolve_cell_style(row_idx, col).justify

    def _save_cell_style(self, row_idx: int, col: int) -> None:
        """Snapshot the current toolbar state as the style for (row_idx, col)."""
        self.cell_styles[(row_idx, col)] = CellStyle(
            family=self.fnt_family.get(),
            size=int(self.fnt_size.get()),
            bold=self.bold_var.get(),
            italic=self.italic_var.get(),
            fg=self.txt_color,
            bg=self.fill_color,
            justify=self._align.get(),
        )

    def _load_cell_style_into_toolbar(self, row_idx: int, col: int) -> None:
        """Reflect a cell's saved style back into the toolbar controls."""
        s = self.cell_styles.get((row_idx, col))
        if not s:
            return
        if s.family:
            self.fnt_family.set(s.family)
        if s.size:
            self.fnt_size.set(str(s.size))
        self.bold_var.set(s.bold)
        self.italic_var.set(s.italic)
        if s.fg:
            self.txt_color = s.fg
            self._pen_prev.config(bg=s.fg)
        if s.bg not in (None, ROW_BG_A, ROW_BG_B):
            self.fill_color = s.bg
            self._fill_prev.config(bg=s.bg)
        if s.justify:
            self._align.set(s.justify)
            self._update_align_btns()

    # ------------------------------------------------------------------
    # Toolbar
    # ------------------------------------------------------------------

    def _build_toolbar(self) -> None:
        label_kw: dict[str, Any] = dict(
            bg=TOOLBAR_BG, fg="#aaaaaa", font=("Segoe UI", 9)
        )
        toolbar = tk.Frame(self.root, bg=TOOLBAR_BG)
        toolbar.pack(fill=tk.X)

        self._build_toolbar_row1(toolbar, label_kw)
        self._build_toolbar_row2(toolbar, label_kw)

    def _build_toolbar_row1(
        self, toolbar: tk.Frame, label_kw: dict[str, Any]
    ) -> None:
        row = tk.Frame(toolbar, bg=TOOLBAR_BG)
        row.pack(fill=tk.X, padx=12, pady=(8, 3))

        tk.Label(
            row, text="TASK MANAGER", bg=TOOLBAR_BG, fg="#f0f0f0",
            font=("Segoe UI", 13, "bold"),
        ).pack(side=tk.LEFT)

        for label, command, bg in (
            ("EXPORT .TXT", self._export_txt, "#222"),
            ("SAVE  (.mesh)", self._save, "#1a472a"),
            ("OPEN FILE", self._open, "#3a3a3a"),
        ):
            tk.Button(
                row, text=label, bg=bg, fg="white", relief=tk.FLAT,
                padx=12, pady=3, font=("Segoe UI", 9, "bold"),
                activebackground="#555", activeforeground="white",
                cursor="hand2", command=command,
            ).pack(side=tk.RIGHT, padx=3)

    def _build_toolbar_row2(
        self, toolbar: tk.Frame, label_kw: dict[str, Any]
    ) -> None:
        row = tk.Frame(toolbar, bg=TOOLBAR_BG)
        row.pack(fill=tk.X, padx=12, pady=(3, 9))

        # Font family
        tk.Label(row, text="Font:", **label_kw).pack(side=tk.LEFT, padx=(0, 4))
        family_cb = ttk.Combobox(
            row, textvariable=self.fnt_family,
            values=sorted(tkfont.families()), width=17, state="readonly",
        )
        family_cb.pack(side=tk.LEFT)
        family_cb.bind("<<ComboboxSelected>>", self._on_toolbar_changed)

        # Font size
        tk.Label(row, text="  Size:", **label_kw).pack(side=tk.LEFT, padx=(6, 4))
        size_cb = ttk.Combobox(
            row, textvariable=self.fnt_size,
            values=list(FONT_SIZES), width=4, state="readonly",
        )
        size_cb.pack(side=tk.LEFT)
        size_cb.bind("<<ComboboxSelected>>", self._on_toolbar_changed)

        tk.Label(row, text="  ", bg=TOOLBAR_BG).pack(side=tk.LEFT)

        # Bold / Italic toggles
        bold_ital_kw: dict[str, Any] = dict(
            bg=TOOLBAR_BG, fg="white", selectcolor="#3a6ea5",
            activebackground=TOOLBAR_BG, cursor="hand2",
            command=self._on_toolbar_changed,
        )
        tk.Checkbutton(
            row, text=" B ", variable=self.bold_var,
            font=("Segoe UI", 10, "bold"), **bold_ital_kw,
        ).pack(side=tk.LEFT, padx=2)
        tk.Checkbutton(
            row, text=" I ", variable=self.italic_var,
            font=("Segoe UI", 10, "italic"), **bold_ital_kw,
        ).pack(side=tk.LEFT, padx=2)

        self._toolbar_separator(row)

        # Text colour
        tk.Button(
            row, text="A  Text", bg="#2e2e2e", fg="#ddd", relief=tk.FLAT,
            padx=7, font=("Segoe UI", 9), cursor="hand2",
            command=self._pick_txt_color,
        ).pack(side=tk.LEFT)
        self._pen_prev = tk.Label(
            row, bg=self.txt_color, width=3, height=1, relief=tk.RIDGE, bd=2,
        )
        self._pen_prev.pack(side=tk.LEFT, padx=(2, 10))

        # Fill colour
        tk.Button(
            row, text="▬  Fill", bg="#2e2e2e", fg="#ddd", relief=tk.FLAT,
            padx=7, font=("Segoe UI", 9), cursor="hand2",
            command=self._pick_fill_color,
        ).pack(side=tk.LEFT)
        self._fill_prev = tk.Label(
            row, bg=self.fill_color, width=3, height=1, relief=tk.RIDGE, bd=2,
        )
        self._fill_prev.pack(side=tk.LEFT, padx=(2, 6))

        self._toolbar_separator(row)

        # Alignment buttons
        tk.Label(row, text="Align:", **label_kw).pack(side=tk.LEFT, padx=(0, 4))
        for symbol, value in ALIGN_SYMBOLS:
            btn = tk.Button(
                row, text=symbol, bg="#2e2e2e", fg="#ccc", relief=tk.FLAT,
                padx=6, font=("Courier New", 10), cursor="hand2",
                command=lambda v=value: self._set_align(v),
            )
            btn.pack(side=tk.LEFT, padx=1)
            self._align_btns[value] = btn
        self._update_align_btns()

        tk.Label(
            row,
            text="  Tab→C = col after  |  Tab→R = row after  |  Tab alone = section title  |  Shift+Tab = row after",
            bg=TOOLBAR_BG, fg="#555", font=("Segoe UI", 8, "italic"),
        ).pack(side=tk.RIGHT, padx=6)

    @staticmethod
    def _toolbar_separator(parent: tk.Frame) -> None:
        tk.Frame(parent, bg="#444", width=1, height=20).pack(
            side=tk.LEFT, padx=10, fill=tk.Y
        )

    # -- Toolbar event handlers ----------------------------------------

    def _on_toolbar_changed(self, *_: Any) -> None:
        """Re-render after any toolbar control changes."""
        self._flush_all()
        if self.focused_item is not None:
            self._save_cell_style(self.focused_item, self.focused_col)
        self._render()

    def _pick_txt_color(self) -> None:
        color = colorchooser.askcolor(color=self.txt_color, title="Text Color")[1]
        if color:
            self.txt_color = color
            self._pen_prev.config(bg=color)
            self._on_toolbar_changed()

    def _pick_fill_color(self) -> None:
        color = colorchooser.askcolor(color=self.fill_color, title="Cell Fill")[1]
        if color:
            self.fill_color = color
            self._fill_prev.config(bg=color)
            self._on_toolbar_changed()

    def _set_align(self, value: str) -> None:
        self._align.set(value)
        self._update_align_btns()
        self._on_toolbar_changed()

    def _update_align_btns(self) -> None:
        current = self._align.get()
        for value, btn in self._align_btns.items():
            btn.config(bg="#3a6ea5" if value == current else "#2e2e2e")

    # ------------------------------------------------------------------
    # Canvas / scroll area
    # ------------------------------------------------------------------

    def _build_table_area(self) -> None:
        outer = tk.Frame(self.root, bg="white")
        outer.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(outer, bg="white", highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=self.canvas.yview)
        hsb = ttk.Scrollbar(outer, orient=tk.HORIZONTAL, command=self.canvas.xview)

        self.canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.inner = tk.Frame(self.canvas, bg="white")
        self._canvas_window = self.canvas.create_window(
            (0, 0), window=self.inner, anchor="nw"
        )

        self.inner.bind(
            "<Configure>",
            lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfig(self._canvas_window, width=e.width),
        )
        self.root.bind_all(
            "<MouseWheel>",
            lambda e: self.canvas.yview_scroll(int(-1 * e.delta / 120), "units"),
        )

    # ------------------------------------------------------------------
    # Render
    # ------------------------------------------------------------------

    def _render(self) -> None:
        """Tear down and rebuild the entire grid from self.items."""
        yview = self.canvas.yview()[0]

        for widget in self.inner.winfo_children():
            widget.destroy()
        self.cell_widgets.clear()

        self._configure_grid_columns()

        serial_map = _compute_serials(self.items)
        for idx, item in enumerate(self.items):
            if isinstance(item, TitleRow):
                self._render_title_row(idx, item)
            elif isinstance(item, HeaderRow):
                self._render_header_row(idx, item)
            else:
                self._render_data_row(idx, item, serial_map[idx])

        tk.Button(
            self.inner, text="  ＋  Add Row  ",
            bg="#f0f0f0", fg="#888", relief=tk.FLAT,
            font=("Segoe UI", 9), cursor="hand2",
            activebackground="#e0e0e0",
            command=self._add_row,
        ).grid(
            row=len(self.items), column=0,
            columnspan=self.num_cols + 2,
            sticky="w", padx=10, pady=8,
        )

        self.inner.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.root.after(20, lambda: self.canvas.yview_moveto(yview))

        if self.focused_item is not None:
            widget = self.cell_widgets.get((self.focused_item, self.focused_col))
            if widget:
                self.root.after(35, widget.focus_set)

    def _configure_grid_columns(self) -> None:
        self.inner.columnconfigure(0, minsize=COL0_W)
        for col in range(1, self.num_cols + 1):
            self.inner.columnconfigure(col, minsize=DATA_W)
        self.inner.columnconfigure(self.num_cols + 1, minsize=ADDCOL_W)

    # ------------------------------------------------------------------
    # Row builders
    # ------------------------------------------------------------------

    def _render_title_row(self, row_idx: int, item: TitleRow) -> None:
        """Dark ribbon: [checkbox]  [§]  [editable red title text]."""
        host = tk.Frame(self.inner, bg=TITLE_BG)
        host.grid(row=row_idx, column=0, columnspan=self.num_cols + 2, sticky="ew")

        # "Add Row Here" separator (between sections only)
        if row_idx > 0:
            self._add_row_here_separator(host, row_idx)

        # Header toggle checkbox
        hvar = tk.BooleanVar(value=item.has_header)
        tk.Checkbutton(
            host, variable=hvar, bg=TITLE_BG,
            activebackground=TITLE_BG, selectcolor="#333",
            cursor="hand2",
            command=lambda: self._toggle_section_header(row_idx, hvar),
        ).pack(side=tk.LEFT, padx=(8, 0))

        tk.Label(
            host, text="§", bg=TITLE_BG, fg="#555",
            font=("Segoe UI", 10),
        ).pack(side=tk.LEFT, padx=2)

        entry = tk.Entry(
            host, bg=TITLE_BG, fg=TITLE_FG,
            font=("Segoe UI", 12, "bold"),
            relief=tk.FLAT, justify="center",
            insertbackground=TITLE_FG, bd=0,
        )
        entry.insert(0, item.text)
        entry.pack(fill=tk.BOTH, expand=True, ipady=7, padx=20)

        entry.bind("<FocusIn>", lambda _e, w=entry: w.selection_range(0, tk.END))
        entry.bind(
            "<FocusOut>",
            lambda _e, i=row_idx, w=entry: setattr(self.items[i], "text", w.get()),
        )
        entry.bind(
            "<Return>",
            lambda _e, i=row_idx, w=entry: setattr(self.items[i], "text", w.get()),
        )
        entry.bind(
            "<Tab>",
            lambda _e, i=row_idx: self._remove_title(i) or "break",
        )
        self.cell_widgets[(row_idx, -1)] = entry

    def _add_row_here_separator(self, host: tk.Frame, row_idx: int) -> None:
        sep = tk.Frame(host, bg=SEPARATOR_BG, height=20)
        sep.pack(fill=tk.X, side=tk.TOP)
        sep.pack_propagate(False)
        tk.Button(
            sep, text="＋  Add Row Here",
            bg=SEPARATOR_BG, fg="#999999", relief=tk.FLAT,
            font=("Segoe UI", 8), cursor="hand2",
            activebackground="#c8c8c8", activeforeground="#444",
            command=lambda i=row_idx: self._add_row_at(i),
        ).place(relx=0.5, rely=0.5, anchor="center")

    def _toggle_section_header(self, row_idx: int, hvar: tk.BooleanVar) -> None:
        """Insert or remove the header row that follows a title ribbon."""
        self._flush_all()
        enabled = hvar.get()
        self.items[row_idx].has_header = enabled
        next_idx = row_idx + 1

        if enabled:
            if next_idx >= len(self.items) or not isinstance(
                self.items[next_idx], HeaderRow
            ):
                self.items.insert(
                    next_idx,
                    HeaderRow(values=[f"Column {c + 1}" for c in range(self.num_cols)]),
                )
        else:
            if next_idx < len(self.items) and isinstance(
                self.items[next_idx], HeaderRow
            ):
                self.items.pop(next_idx)

        self._render()

    def _render_header_row(self, row_idx: int, item: HeaderRow) -> None:
        """
        Column heading row with editable names and [+] add-column button.

        Shift+Click on any header cell inserts a new column AFTER that column.
        """
        tk.Label(
            self.inner, text="  #  ",
            bg=HEADER_BG, fg=HEADER_FG,
            font=("Segoe UI", 9, "bold"), relief=tk.GROOVE,
        ).grid(row=row_idx, column=0, sticky="nsew", ipady=5)

        for col in range(self.num_cols):
            val = item.values[col] if col < len(item.values) else f"Column {col + 1}"
            entry = tk.Entry(
                self.inner, bg=HEADER_BG, fg=HEADER_FG,
                font=("Segoe UI", 9, "bold"),
                relief=tk.GROOVE, justify="center", bd=1,
            )
            entry.insert(0, val)
            entry.grid(row=row_idx, column=col + 1, sticky="nsew", ipady=5)
            entry.bind(
                "<FocusOut>",
                lambda _e, i=row_idx, c=col, w=entry: self._save_header_cell(i, c, w),
            )
            entry.bind(
                "<Return>",
                lambda _e, i=row_idx, c=col, w=entry: self._save_header_cell(i, c, w),
            )
            # Shift+Click → insert a column AFTER this one
            entry.bind(
                "<Shift-Button-1>",
                lambda _e, c=col: self._add_column_after(c),
            )
            entry.bind(
                "<Button-3>",   # right-click also works as an alternative
                lambda _e, c=col: self._add_column_after(c),
            )
            self.cell_widgets[(row_idx, col)] = entry

        tk.Button(
            self.inner, text="+",
            bg=HEADER_BG, fg="#444",
            relief=tk.FLAT, font=("Segoe UI", 14, "bold"),
            cursor="hand2", command=self._add_column,
        ).grid(row=row_idx, column=self.num_cols + 1, sticky="nsew")

    def _save_header_cell(self, row_idx: int, col: int, widget: tk.Entry) -> None:
        item = self.items[row_idx]
        if isinstance(item, HeaderRow):
            self._ensure_col(item.values, col, f"Column {col + 1}")
            item.values[col] = widget.get()

    def _render_data_row(
        self, row_idx: int, item: DataRow, serial: str | None
    ) -> None:
        """Data row: serial label in col0, editable Entry cells for data."""
        row_bg = self._get_row_bg(row_idx)

        # Serial number cell (col 0)
        frame = tk.Frame(self.inner, bg=row_bg, relief=tk.GROOVE, bd=1)
        frame.grid(row=row_idx, column=0, sticky="nsew")
        tk.Label(
            frame, text=serial or "   ",
            bg=row_bg, fg=SERIAL_FG,
            font=("Courier New", 9),
        ).pack(expand=True)

        # Data cells
        for col in range(self.num_cols):
            val = item.values[col] if col < len(item.values) else ""
            fg, bg = self._cell_colors(row_idx, col, row_bg)

            entry = tk.Entry(
                self.inner, bg=bg, fg=fg,
                font=self._cell_font(row_idx, col),
                relief=tk.GROOVE, bd=1,
                justify=self._cell_align(row_idx, col),
            )
            entry.insert(0, val)
            entry.grid(row=row_idx, column=col + 1, sticky="nsew", ipady=3)

            entry.bind(
                "<FocusIn>",
                lambda _e, i=row_idx, c=col: self._on_cell_focus(i, c),
            )
            entry.bind(
                "<FocusOut>",
                lambda _e, i=row_idx, c=col, w=entry: self._save_data_cell(i, c, w),
            )
            entry.bind("<Tab>", self._on_tab)
            entry.bind("<Shift-Tab>", self._on_shift_tab)
            # Cross-platform: Linux sometimes uses ISO_Left_Tab for Shift+Tab
            entry.bind("<ISO_Left_Tab>", self._on_shift_tab)
            entry.bind("<Return>", self._on_return)
            self.cell_widgets[(row_idx, col)] = entry

    def _on_cell_focus(self, row_idx: int, col: int) -> None:
        """Update focus tracking and reflect the cell's style in the toolbar."""
        self.focused_item = row_idx
        self.focused_col = col
        self._load_cell_style_into_toolbar(row_idx, col)

    def _save_data_cell(self, row_idx: int, col: int, widget: tk.Entry) -> None:
        item = self.items[row_idx]
        if isinstance(item, DataRow):
            self._ensure_col(item.values, col, "")
            item.values[col] = widget.get()

    # ------------------------------------------------------------------
    # Keyboard event handlers
    # ------------------------------------------------------------------

    # ── Tab chord dispatcher ──────────────────────────────────────────
    # Press Tab → chord window opens (350 ms).
    #   • Next key = C  → insert column AFTER focused column
    #   • Next key = R  → insert row AFTER focused row
    #   • Timeout / any other key → insert title ribbon ABOVE current row

    def _on_tab(self, event: tk.Event) -> str:
        """
        Tab keypress: start the chord window.
        If Tab arrives while a chord is already pending we treat it as a
        'plain Tab' so rapid double-Tab still inserts a title.
        """
        if self._chord_pending:
            self._chord_cancel()
            self._do_insert_title()
            return "break"

        if self.focused_item is None:
            return "break"

        # Snapshot focused cell now (focus may move before timer fires)
        row_idx = self.focused_item
        col = self.focused_col
        widget = self.cell_widgets.get((row_idx, col))
        if widget:
            self._save_data_cell(row_idx, col, widget)

        self._chord_pending = True
        self._chord_after_id = self.root.after(350, self._chord_timeout)

        # Bind C and R globally for the duration of the chord window
        self.root.bind_all("<KeyPress-c>", self._on_chord_c, add=False)
        self.root.bind_all("<KeyPress-C>", self._on_chord_c, add=False)
        self.root.bind_all("<KeyPress-r>", self._on_chord_r, add=False)
        self.root.bind_all("<KeyPress-R>", self._on_chord_r, add=False)
        return "break"

    def _chord_cancel(self) -> None:
        """Cancel a pending chord window (timer + bindings)."""
        if self._chord_after_id is not None:
            self.root.after_cancel(self._chord_after_id)
            self._chord_after_id = None
        self._chord_pending = False
        for seq in ("<KeyPress-c>", "<KeyPress-C>", "<KeyPress-r>", "<KeyPress-R>"):
            try:
                self.root.unbind_all(seq)
            except tk.TclError:
                pass

    def _chord_timeout(self) -> None:
        """350 ms elapsed with no chord key → run the default Tab action."""
        self._chord_after_id = None
        if self._chord_pending:
            self._chord_pending = False
            for seq in ("<KeyPress-c>", "<KeyPress-C>", "<KeyPress-r>", "<KeyPress-R>"):
                try:
                    self.root.unbind_all(seq)
                except tk.TclError:
                    pass
            self._do_insert_title()

    def _on_chord_c(self, _event: tk.Event) -> str:
        """Tab → C: insert a blank column AFTER the currently focused column."""
        if not self._chord_pending:
            return ""
        self._chord_cancel()
        if self.focused_item is not None:
            self._add_column_after(self.focused_col)
        return "break"

    def _on_chord_r(self, _event: tk.Event) -> str:
        """Tab → R: insert a blank data row AFTER the currently focused row."""
        if not self._chord_pending:
            return ""
        self._chord_cancel()
        if self.focused_item is not None:
            self._insert_row_after(self.focused_item)
        return "break"

    def _do_insert_title(self) -> None:
        """Insert a title ribbon ABOVE the currently focused row."""
        if self.focused_item is None:
            return
        row_idx = self.focused_item
        self.items.insert(row_idx, TitleRow())
        self.focused_item = row_idx + 1
        self._render()
        title_entry = self.cell_widgets.get((row_idx, -1))
        if title_entry:
            self.root.after(
                50, lambda w=title_entry: (w.focus_set(), w.selection_range(0, tk.END))
            )

    def _on_shift_tab(self, _event: tk.Event) -> str:
        """Shift+Tab: insert a blank data row immediately AFTER the current row."""
        if self.focused_item is None:
            return "break"
        row_idx = self.focused_item
        col = self.focused_col
        widget = self.cell_widgets.get((row_idx, col))
        if widget:
            self._save_data_cell(row_idx, col, widget)
        self._insert_row_after(row_idx)
        return "break"

    def _on_return(self, _event: tk.Event) -> str:
        """Return: move focus to the same column in the next data row."""
        if self.focused_item is None:
            return "break"

        next_idx = self.focused_item + 1
        col = self.focused_col
        while next_idx < len(self.items):
            if isinstance(self.items[next_idx], DataRow):
                widget = self.cell_widgets.get((next_idx, col))
                if widget:
                    widget.focus_set()
                break
            next_idx += 1
        return "break"

    # ------------------------------------------------------------------
    # Structural mutations
    # ------------------------------------------------------------------

    def _add_row(self) -> None:
        self._flush_all()
        self.items.append(DataRow(values=[""] * self.num_cols))
        self._render()

    def _insert_row_after(self, row_idx: int) -> None:
        """
        Core helper: insert a blank DataRow at row_idx+1, re-key cell_styles,
        update focused_item, and re-render.  Called by Shift+Tab, Tab→R, and
        _add_row_at (which uses the 'before' variant).
        """
        self._flush_all()
        insert_at = row_idx + 1
        # Shift styles: rows strictly after row_idx move down one
        new_styles: dict[tuple[int, int], CellStyle] = {}
        for (r, c), style in self.cell_styles.items():
            new_styles[(r + 1, c) if r > row_idx else (r, c)] = style
        self.cell_styles = new_styles
        self.items.insert(insert_at, DataRow(values=[""] * self.num_cols))
        self.focused_item = insert_at
        self.focused_col = 0
        self._render()

    def _add_row_at(self, index: int) -> None:
        """Insert a blank data row immediately BEFORE *index* (used by separator button)."""
        self._flush_all()
        # Rows at or after the insertion point shift down
        new_styles: dict[tuple[int, int], CellStyle] = {}
        for (r, c), style in self.cell_styles.items():
            new_styles[(r + 1, c) if r >= index else (r, c)] = style
        self.cell_styles = new_styles
        self.items.insert(index, DataRow(values=[""] * self.num_cols))
        self._render()


    def _add_column(self) -> None:
        """Append a new column at the far right (the [+] header button)."""
        self._add_column_after(self.num_cols - 1)

    def _add_column_after(self, after_col: int) -> None:
        """
        Insert a new blank column immediately AFTER *after_col*.

        Steps:
          1. Flush all live widgets to the model.
          2. In every HeaderRow / DataRow, splice a blank string at
             position after_col + 1 and trim to the new num_cols.
          3. Re-key cell_styles so indices beyond the insertion point
             shift right by 1 (column axis).
          4. Re-render.
        """
        self._flush_all()
        insert_at = after_col + 1
        self.num_cols += 1

        for item in self.items:
            if isinstance(item, HeaderRow):
                # Ensure list is long enough before splicing
                while len(item.values) < self.num_cols - 1:
                    item.values.append(f"Column {len(item.values) + 1}")
                item.values.insert(insert_at, f"Column {insert_at + 1}")
                # Trim to exact num_cols length
                item.values = item.values[: self.num_cols]
                # Re-number column labels after the insertion point
                for i in range(insert_at + 1, self.num_cols):
                    # Only rename columns that still carry the default name pattern
                    pass  # user-customised names are preserved as-is

            elif isinstance(item, DataRow):
                while len(item.values) < self.num_cols - 1:
                    item.values.append("")
                item.values.insert(insert_at, "")
                item.values = item.values[: self.num_cols]

        # Re-key cell styles (column axis: indices >= insert_at shift right)
        new_styles: dict[tuple[int, int], CellStyle] = {}
        for (r, c), style in self.cell_styles.items():
            if c >= insert_at:
                new_styles[(r, c + 1)] = style
            else:
                new_styles[(r, c)] = style
        self.cell_styles = new_styles

        self._render()

    def _remove_title(self, row_idx: int) -> None:
        """
        Tab on a title entry removes the ribbon (and its linked HeaderRow
        if the header-toggle was active).
        """
        self._flush_all()
        self.items.pop(row_idx)
        if row_idx < len(self.items) and isinstance(self.items[row_idx], HeaderRow):
            self.items.pop(row_idx)
        self.focused_item = None
        self._render()

    # ------------------------------------------------------------------
    # File I/O  –  Save / Load (.mesh JSON)
    # ------------------------------------------------------------------

    def _save(self) -> None:
        """
        Save the full table state as a JSON .mesh file.

        Safety:
         - Flushes every live widget before serialising.
         - Embeds a SHA-256 checksum of the payload.
         - Atomic write: temp-file → os.replace (no half-written files).
         - Backs up the previous version to <name>.mesh.bak automatically.
        """
        raw_name = simpledialog.askstring(
            "Save File", "Enter filename (without extension):", parent=self.root
        )
        if not raw_name or not raw_name.strip():
            return

        filename = raw_name.strip()
        if not filename.lower().endswith(".mesh"):
            filename += ".mesh"
        path = os.path.join(SCRIPT_DIR, filename)

        self._flush_all()
        payload = _serialize(self.items, self.cell_styles, self.num_cols)
        text = json.dumps(payload, ensure_ascii=False, indent=2)

        try:
            _atomic_write(path, text)
        except Exception as exc:
            messagebox.showerror("Save Failed", f"Could not write file:\n{exc}")
            return

        bak = path + ".bak"
        bak_note = f"\n(Previous version backed up to {os.path.basename(bak)})" if os.path.exists(bak) else ""
        messagebox.showinfo("Saved ✓", f"Saved to:\n{path}{bak_note}")

    def _open(self) -> None:
        """
        Open a previously saved .mesh file and restore the full table state.

        Safety:
         - Validates JSON structure before touching live state.
         - Verifies SHA-256 checksum; on mismatch asks the user whether to
           proceed with best-effort repair rather than crashing.
         - Every row value list is padded/trimmed to num_cols.
         - Every CellStyle is repaired field-by-field with known good defaults.
         - If the file is completely unparseable the current session is kept.
        """
        path = filedialog.askopenfilename(
            initialdir=SCRIPT_DIR,
            filetypes=[
                ("Mesh files", "*.mesh"),
                ("All files", "*.*"),
            ],
            title="Open .mesh File",
        )
        if not path:
            return

        # ── Parse JSON ────────────────────────────────────────────────
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except json.JSONDecodeError as exc:
            messagebox.showerror(
                "Open Failed",
                f"File is not valid JSON:\n{exc}\n\nYour current session is unchanged.",
            )
            return
        except Exception as exc:
            messagebox.showerror("Open Failed", f"Could not read file:\n{exc}")
            return

        if not isinstance(data, dict):
            messagebox.showerror("Open Failed", "File does not contain a valid table object.")
            return

        # ── Deserialise (checksum-aware) ───────────────────────────────
        try:
            items, cell_styles, num_cols = _deserialize(data)
        except ValueError as exc:
            # Checksum mismatch – ask user
            proceed = messagebox.askyesno(
                "Checksum Warning",
                str(exc),
                icon="warning",
            )
            if not proceed:
                return
            # Retry without checksum enforcement
            patched = dict(data)
            patched.pop("checksum", None)
            try:
                items, cell_styles, num_cols = _deserialize(patched)
            except Exception as exc2:
                messagebox.showerror("Repair Failed", f"Could not repair file:\n{exc2}")
                return
        except Exception as exc:
            messagebox.showerror("Parse Error", f"Unexpected error parsing file:\n{exc}")
            return

        # ── Apply to live state ────────────────────────────────────────
        self.items = items
        self.cell_styles = cell_styles
        self.num_cols = num_cols
        self.focused_item = None
        self.focused_col = 0
        self._chord_cancel()

        self._render()
        messagebox.showinfo(
            "Opened ✓",
            f"Loaded:\n{path}\n\n"
            f"{len(items)} rows  ·  {num_cols} columns",
        )


    # ------------------------------------------------------------------
    # Plain-text export  (unchanged, kept as separate action)
    # ------------------------------------------------------------------

    def _export_txt(self) -> None:
        """Export the table to a formatted plain-text (.txt) file."""
        raw_name = simpledialog.askstring(
            "Export .txt", "Enter filename (without extension):", parent=self.root
        )
        if not raw_name or not raw_name.strip():
            return

        filename = raw_name.strip()
        if not filename.lower().endswith(".txt"):
            filename += ".txt"
        path = os.path.join(SCRIPT_DIR, filename)

        self._flush_all()
        lines = self._build_export_lines()

        with open(path, "w", encoding="utf-8") as fh:
            fh.write("\n".join(lines))

        messagebox.showinfo("Exported ✓", f"Exported to:\n{path}")

    def _build_export_lines(self) -> list[str]:
        """Return the plain-text export lines for all rows."""
        lines: list[str] = []
        serial_map = _compute_serials(self.items)
        divider = "-" * EXPORT_LINE_WIDTH
        header_divider = "=" * EXPORT_LINE_WIDTH

        for idx, item in enumerate(self.items):
            if isinstance(item, TitleRow):
                text = item.text.upper().center(EXPORT_LINE_WIDTH - 4)
                lines += ["", header_divider, f"  {text}", header_divider, ""]

            elif isinstance(item, HeaderRow):
                padded = "\t".join(v.ljust(EXPORT_COL_PAD) for v in item.values)
                lines += [padded, divider]

            elif isinstance(item, DataRow):
                serial = serial_map.get(idx) or "   "
                cells = "\t".join(
                    (
                        item.values[c].ljust(EXPORT_COL_PAD)
                        if c < len(item.values)
                        else " " * EXPORT_COL_PAD
                    )
                    for c in range(self.num_cols)
                )
                lines.append(f"{serial}  {cells}")

        return lines


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def _apply_ttk_theme(root: tk.Tk) -> None:
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(
        "TCombobox",
        fieldbackground="#2e2e2e",
        background="#2e2e2e",
        foreground="#f0f0f0",
        selectbackground="#444",
        selectforeground="#fff",
        arrowcolor="#ccc",
    )
    style.configure("TScrollbar", background="#ccc", troughcolor="#eee")


def main() -> None:
    root = tk.Tk()
    _apply_ttk_theme(root)
    MeshTable(root)
    root.mainloop()


if __name__ == "__main__":
    main()
