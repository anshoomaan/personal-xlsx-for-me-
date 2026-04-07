"""
mesh_table.py  –  Task Mesh Personal Database
Run:  python mesh_table.py
"""

import os
import tkinter as tk
from tkinter import ttk, colorchooser, filedialog, messagebox, simpledialog
import tkinter.font as tkfont

# Script directory (used for save / open dialogs)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


class MeshTable:
    COL0_W   = 80
    DATA_W   = 130
    ADDCOL_W = 34

    TOOLBAR_BG = "#1a1a1a"
    HEADER_BG  = "#d0d0d0"
    HEADER_FG  = "#222222"
    TITLE_BG   = "#0d0d0d"
    TITLE_FG   = "#ff3333"
    ROW_BG_A   = "#ffffff"
    ROW_BG_B   = "#f7f7f7"
    SERIAL_FG  = "#aaaaaa"

    def __init__(self, root):
        self.root = root
        self.root.title("Task Mesh – Personal Database")
        self.root.geometry("1350x760")
        self.root.minsize(900, 500)
        self.root.configure(bg=self.TOOLBAR_BG)

        self.num_cols = 10
        self.items = []
        self._init_items()

        # Per-cell style: {(ii, col): {family, size, bold, italic, fg, bg, justify}}
        self.cell_styles = {}

        # Toolbar state
        self.fnt_family = tk.StringVar(value="Consolas")
        self.fnt_size   = tk.StringVar(value="10")
        self.bold_var   = tk.BooleanVar(value=False)
        self.italic_var = tk.BooleanVar(value=False)
        self.txt_color  = "#111111"
        self.fill_color = "#ffffff"
        self._align     = tk.StringVar(value="left")
        self._align_btns = {}

        # Focus tracking
        self.focused_item = None
        self.focused_col  = 0
        self.cell_widgets = {}

        self._build_toolbar()
        self._build_table_area()
        self._render()

    # ── Data model ─────────────────────────────────────────────────────────────

    def _init_items(self):
        self.items = [
            {"type": "header", "values": [f"Column {i+1}" for i in range(self.num_cols)]}
        ] + [
            {"type": "data", "values": [""] * self.num_cols}
            for _ in range(28)
        ]

    def _compute_serials(self):
        out, n = {}, 1
        for i, item in enumerate(self.items):
            if item["type"] in ("header", "title"):
                n = 1
                out[i] = None
            else:
                out[i] = f"{n:03d}"
                n += 1
        return out

    def _flush_all(self):
        """Persist all Entry widgets back to self.items."""
        for (ii, col), w in list(self.cell_widgets.items()):
            if not isinstance(w, tk.Entry):
                continue
            try:
                val = w.get()
            except tk.TclError:
                continue
            item = self.items[ii]
            if item["type"] == "title":
                item["text"] = val
            elif item["type"] == "header":
                while len(item["values"]) <= col:
                    item["values"].append(f"Column {col+1}")
                item["values"][col] = val
            elif item["type"] == "data":
                while len(item["values"]) <= col:
                    item["values"].append("")
                item["values"][col] = val

    # ── Toolbar ────────────────────────────────────────────────────────────────

    def _build_toolbar(self):
        C = self.TOOLBAR_BG
        lkw = dict(bg=C, fg="#aaaaaa", font=("Segoe UI", 9))
        tb = tk.Frame(self.root, bg=C)
        tb.pack(fill=tk.X)

        # Row 1
        r1 = tk.Frame(tb, bg=C)
        r1.pack(fill=tk.X, padx=12, pady=(8, 3))
        tk.Label(r1, text="TASK MANAGER", bg=C, fg="#f0f0f0",
                 font=("Segoe UI", 13, "bold")).pack(side=tk.LEFT)
        for label, cmd, bg in [("SAVE AS .TXT", self._save, "#222"),
                                 ("OPEN FILE",    self._open, "#3a3a3a")]:
            tk.Button(r1, text=label, bg=bg, fg="white", relief=tk.FLAT,
                      padx=12, pady=3, font=("Segoe UI", 9, "bold"),
                      activebackground="#555", activeforeground="white",
                      cursor="hand2", command=cmd).pack(side=tk.RIGHT, padx=3)

        # Row 2
        r2 = tk.Frame(tb, bg=C)
        r2.pack(fill=tk.X, padx=12, pady=(3, 9))

        tk.Label(r2, text="Font:", **lkw).pack(side=tk.LEFT, padx=(0, 4))
        fcb = ttk.Combobox(r2, textvariable=self.fnt_family,
                           values=sorted(tkfont.families()), width=17, state="readonly")
        fcb.pack(side=tk.LEFT)
        fcb.bind("<<ComboboxSelected>>", self._toolbar_changed)

        tk.Label(r2, text="  Size:", **lkw).pack(side=tk.LEFT, padx=(6, 4))
        scb = ttk.Combobox(r2, textvariable=self.fnt_size,
                           values=["8","9","10","11","12","14","16","18","20","24","28","32"],
                           width=4, state="readonly")
        scb.pack(side=tk.LEFT)
        scb.bind("<<ComboboxSelected>>", self._toolbar_changed)

        tk.Label(r2, text="  ", bg=C).pack(side=tk.LEFT)

        # Bold
        self.bold_btn = tk.Checkbutton(r2, text=" B ", variable=self.bold_var,
                                        bg=C, fg="white", selectcolor="#3a6ea5",
                                        activebackground=C, font=("Segoe UI", 10, "bold"),
                                        cursor="hand2", command=self._toolbar_changed)
        self.bold_btn.pack(side=tk.LEFT, padx=2)

        # Italic
        self.ital_btn = tk.Checkbutton(r2, text=" I ", variable=self.italic_var,
                                        bg=C, fg="white", selectcolor="#3a6ea5",
                                        activebackground=C, font=("Segoe UI", 10, "italic"),
                                        cursor="hand2", command=self._toolbar_changed)
        self.ital_btn.pack(side=tk.LEFT, padx=2)

        # Separator
        tk.Frame(r2, bg="#444", width=1, height=20).pack(side=tk.LEFT, padx=10, fill=tk.Y)

        # Text color
        tk.Button(r2, text="A  Text", bg="#2e2e2e", fg="#ddd", relief=tk.FLAT,
                  padx=7, font=("Segoe UI", 9), cursor="hand2",
                  command=self._pick_txt_color).pack(side=tk.LEFT)
        self._pen_prev = tk.Label(r2, bg=self.txt_color, width=3, height=1,
                                   relief=tk.RIDGE, bd=2)
        self._pen_prev.pack(side=tk.LEFT, padx=(2, 10))

        # Fill color
        tk.Button(r2, text="▬  Fill", bg="#2e2e2e", fg="#ddd", relief=tk.FLAT,
                  padx=7, font=("Segoe UI", 9), cursor="hand2",
                  command=self._pick_fill_color).pack(side=tk.LEFT)
        self._fill_prev = tk.Label(r2, bg=self.fill_color, width=3, height=1,
                                    relief=tk.RIDGE, bd=2)
        self._fill_prev.pack(side=tk.LEFT, padx=(2, 6))

        tk.Frame(r2, bg="#444", width=1, height=20).pack(side=tk.LEFT, padx=10, fill=tk.Y)

        # Alignment
        tk.Label(r2, text="Align:", **lkw).pack(side=tk.LEFT, padx=(0, 4))
        for sym, val in [("≡L","left"),("≡C","center"),("≡R","right")]:
            b = tk.Button(r2, text=sym, bg="#2e2e2e", fg="#ccc", relief=tk.FLAT,
                          padx=6, font=("Courier New", 10), cursor="hand2",
                          command=lambda v=val: self._set_align(v))
            b.pack(side=tk.LEFT, padx=1)
            self._align_btns[val] = b
        self._update_align_btns()

        tk.Label(r2, text="  Tab = insert section title",
                 bg=C, fg="#555", font=("Segoe UI", 8, "italic")).pack(side=tk.RIGHT, padx=6)

    def _toolbar_changed(self, *_):
        self._flush_all()
        if self.focused_item is not None:
            self._write_cell_style(self.focused_item, self.focused_col)
        self._render()

    def _write_cell_style(self, ii, col):
        self.cell_styles[(ii, col)] = {
            "family":  self.fnt_family.get(),
            "size":    int(self.fnt_size.get()),
            "bold":    self.bold_var.get(),
            "italic":  self.italic_var.get(),
            "fg":      self.txt_color,
            "bg":      self.fill_color,
            "justify": self._align.get(),
        }

    def _cell_font(self, ii, col):
        s    = self.cell_styles.get((ii, col), {})
        fam  = s.get("family", self.fnt_family.get())
        sz   = s.get("size",   int(self.fnt_size.get()))
        bold = s.get("bold",   self.bold_var.get())
        ital = s.get("italic", self.italic_var.get())
        st   = " ".join(filter(None, ["bold" if bold else "", "italic" if ital else ""])) or "normal"
        return (fam, sz, st)

    def _cell_colors(self, ii, col, default_bg):
        s  = self.cell_styles.get((ii, col), {})
        fg = s.get("fg", self.txt_color)
        bg = s.get("bg", default_bg)
        return fg, bg

    def _cell_align(self, ii, col):
        return self.cell_styles.get((ii, col), {}).get("justify", self._align.get())

    def _pick_txt_color(self):
        c = colorchooser.askcolor(color=self.txt_color, title="Text Color")[1]
        if c:
            self.txt_color = c
            self._pen_prev.config(bg=c)
            self._toolbar_changed()

    def _pick_fill_color(self):
        c = colorchooser.askcolor(color=self.fill_color, title="Cell Fill")[1]
        if c:
            self.fill_color = c
            self._fill_prev.config(bg=c)
            self._toolbar_changed()

    def _set_align(self, val):
        self._align.set(val)
        self._update_align_btns()
        self._toolbar_changed()

    def _update_align_btns(self):
        cur = self._align.get()
        for val, b in self._align_btns.items():
            b.config(bg="#3a6ea5" if val == cur else "#2e2e2e")

    # ── Canvas / scroll area ───────────────────────────────────────────────────

    def _build_table_area(self):
        outer = tk.Frame(self.root, bg="white")
        outer.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(outer, bg="white", highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient=tk.VERTICAL,   command=self.canvas.yview)
        hsb = ttk.Scrollbar(outer, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        vsb.pack(side=tk.RIGHT,  fill=tk.Y)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.inner = tk.Frame(self.canvas, bg="white")
        self._cwin = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>",
                        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>",
                         lambda e: self.canvas.itemconfig(self._cwin, width=e.width))
        self.root.bind_all("<MouseWheel>",
                           lambda e: self.canvas.yview_scroll(int(-1*e.delta/120), "units"))

    # ── Render ─────────────────────────────────────────────────────────────────

    def _render(self):
        yview = self.canvas.yview()[0]
        for w in self.inner.winfo_children():
            w.destroy()
        self.cell_widgets.clear()

        self.inner.columnconfigure(0, minsize=self.COL0_W)
        for c in range(1, self.num_cols + 1):
            self.inner.columnconfigure(c, minsize=self.DATA_W)
        self.inner.columnconfigure(self.num_cols + 1, minsize=self.ADDCOL_W)

        smap = self._compute_serials()
        for ii, item in enumerate(self.items):
            t = item["type"]
            if   t == "title":  self._row_title(ii, item)
            elif t == "header": self._row_header(ii, item)
            else:               self._row_data(ii, item, smap[ii])

        tk.Button(self.inner, text="  ＋  Add Row  ",
                  bg="#f0f0f0", fg="#888", relief=tk.FLAT,
                  font=("Segoe UI", 9), cursor="hand2",
                  activebackground="#e0e0e0",
                  command=self._add_row
                  ).grid(row=len(self.items), column=0,
                         columnspan=self.num_cols + 2,
                         sticky="w", padx=10, pady=8)

        self.inner.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.root.after(20, lambda: self.canvas.yview_moveto(yview))

        if self.focused_item is not None:
            w = self.cell_widgets.get((self.focused_item, self.focused_col))
            if w:
                self.root.after(35, w.focus_set)

    # ── Row builders ───────────────────────────────────────────────────────────

    def _row_title(self, ii, item):
        """Dark ribbon: [checkbox]  [§]  [editable red title text]"""
        ncols = self.num_cols + 2
        host  = tk.Frame(self.inner, bg=self.TITLE_BG)
        host.grid(row=ii, column=0, columnspan=ncols, sticky="ew")

        hvar = tk.BooleanVar(value=item.get("has_header", False))

        def _toggle(i=ii, v=hvar):
            val = v.get()
            self._flush_all()
            self.items[i]["has_header"] = val
            nxt = i + 1
            if val:
                if nxt >= len(self.items) or self.items[nxt]["type"] != "header":
                    self.items.insert(nxt, {
                        "type":   "header",
                        "values": [f"Column {c+1}" for c in range(self.num_cols)]
                    })
            else:
                if nxt < len(self.items) and self.items[nxt]["type"] == "header":
                    self.items.pop(nxt)
            self._render()

        tk.Checkbutton(host, variable=hvar, bg=self.TITLE_BG,
                       activebackground=self.TITLE_BG, selectcolor="#333",
                       cursor="hand2", command=_toggle
                       ).pack(side=tk.LEFT, padx=(8, 0))

        tk.Label(host, text="§", bg=self.TITLE_BG, fg="#555",
                 font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=2)

        e = tk.Entry(host, bg=self.TITLE_BG, fg=self.TITLE_FG,
                     font=("Segoe UI", 12, "bold"),
                     relief=tk.FLAT, justify="center",
                     insertbackground=self.TITLE_FG, bd=0)
        e.insert(0, item.get("text", "title"))
        e.pack(fill=tk.BOTH, expand=True, ipady=7, padx=20)
        e.bind("<FocusIn>",  lambda ev, w=e: w.selection_range(0, tk.END))
        e.bind("<FocusOut>", lambda ev, i=ii, w=e: self.items[i].__setitem__("text", w.get()))
        e.bind("<Return>",   lambda ev, i=ii, w=e: self.items[i].__setitem__("text", w.get()))
        self.cell_widgets[(ii, -1)] = e

    def _row_header(self, ii, item):
        """Column heading row with editable names + [+] add-column button."""
        tk.Label(self.inner, text="  #  ",
                 bg=self.HEADER_BG, fg=self.HEADER_FG,
                 font=("Segoe UI", 9, "bold"), relief=tk.GROOVE
                 ).grid(row=ii, column=0, sticky="nsew", ipady=5)

        for col in range(self.num_cols):
            val = item["values"][col] if col < len(item["values"]) else f"Column {col+1}"
            e = tk.Entry(self.inner, bg=self.HEADER_BG, fg=self.HEADER_FG,
                         font=("Segoe UI", 9, "bold"),
                         relief=tk.GROOVE, justify="center", bd=1)
            e.insert(0, val)
            e.grid(row=ii, column=col + 1, sticky="nsew", ipady=5)

            def _save_h(ev=None, i=ii, c=col, w=e):
                while len(self.items[i]["values"]) <= c:
                    self.items[i]["values"].append(f"Column {c+1}")
                self.items[i]["values"][c] = w.get()

            e.bind("<FocusOut>", _save_h)
            e.bind("<Return>",   _save_h)
            self.cell_widgets[(ii, col)] = e

        tk.Button(self.inner, text="+",
                  bg=self.HEADER_BG, fg="#444",
                  relief=tk.FLAT, font=("Segoe UI", 14, "bold"),
                  cursor="hand2", command=self._add_column
                  ).grid(row=ii, column=self.num_cols + 1, sticky="nsew")

    def _row_data(self, ii, item, serial):
        """Data row: serial number only in col0, Entry cells for data."""
        row_bg = self.ROW_BG_A if ii % 2 == 0 else self.ROW_BG_B

        # Col-0: serial number only (no checkbox)
        f0 = tk.Frame(self.inner, bg=row_bg, relief=tk.GROOVE, bd=1)
        f0.grid(row=ii, column=0, sticky="nsew")
        tk.Label(f0, text=serial or "   ",
                 bg=row_bg, fg=self.SERIAL_FG,
                 font=("Courier New", 9)).pack(expand=True)

        # Data cells
        for col in range(self.num_cols):
            val = item["values"][col] if col < len(item["values"]) else ""
            fg, bg = self._cell_colors(ii, col, row_bg)
            e = tk.Entry(self.inner, bg=bg, fg=fg,
                         font=self._cell_font(ii, col),
                         relief=tk.GROOVE, bd=1,
                         justify=self._cell_align(ii, col))
            e.insert(0, val)
            e.grid(row=ii, column=col + 1, sticky="nsew", ipady=3)

            def _on_focus(ev, i=ii, c=col):
                self.focused_item = i
                self.focused_col  = c
                # Reflect this cell's style in the toolbar
                s = self.cell_styles.get((i, c), {})
                if s.get("family"):
                    self.fnt_family.set(s["family"])
                if s.get("size"):
                    self.fnt_size.set(str(s["size"]))
                if "bold"   in s: self.bold_var.set(s["bold"])
                if "italic" in s: self.italic_var.set(s["italic"])
                if s.get("fg"):
                    self.txt_color = s["fg"]
                    self._pen_prev.config(bg=s["fg"])
                if s.get("bg") not in (None, self.ROW_BG_A, self.ROW_BG_B):
                    self.fill_color = s["bg"]
                    self._fill_prev.config(bg=s["bg"])
                if s.get("justify"):
                    self._align.set(s["justify"])
                    self._update_align_btns()

            def _save_cell(ev=None, i=ii, c=col, w=e):
                while len(self.items[i]["values"]) <= c:
                    self.items[i]["values"].append("")
                self.items[i]["values"][c] = w.get()

            e.bind("<FocusIn>",  _on_focus)
            e.bind("<FocusOut>", _save_cell)
            e.bind("<Tab>",      self._on_tab)
            e.bind("<Return>",   self._on_return)
            self.cell_widgets[(ii, col)] = e

    # ── Key events ─────────────────────────────────────────────────────────────

    def _on_tab(self, ev):
        if self.focused_item is None:
            return "break"
        ii, col = self.focused_item, self.focused_col
        w = self.cell_widgets.get((ii, col))
        if w:
            while len(self.items[ii]["values"]) <= col:
                self.items[ii]["values"].append("")
            self.items[ii]["values"][col] = w.get()
        self.items.insert(ii, {"type": "title", "text": "title", "has_header": False})
        self.focused_item = ii + 1
        self._render()
        tw = self.cell_widgets.get((ii, -1))
        if tw:
            self.root.after(50, lambda: (tw.focus_set(), tw.selection_range(0, tk.END)))
        return "break"

    def _on_return(self, ev):
        if self.focused_item is None:
            return "break"
        nxt, col = self.focused_item + 1, self.focused_col
        while nxt < len(self.items):
            if self.items[nxt]["type"] == "data":
                w = self.cell_widgets.get((nxt, col))
                if w:
                    w.focus_set()
                break
            nxt += 1
        return "break"

    # ── Mutations ──────────────────────────────────────────────────────────────

    def _add_row(self):
        self._flush_all()
        self.items.append({"type": "data", "values": [""] * self.num_cols})
        self._render()

    def _add_column(self):
        self._flush_all()
        self.num_cols += 1
        for item in self.items:
            if item["type"] in ("header", "data"):
                while len(item["values"]) < self.num_cols:
                    item["values"].append("")
        for item in self.items:
            if item["type"] == "header":
                item["values"][self.num_cols - 1] = f"Column {self.num_cols}"
        self._render()

    # ── File I/O ───────────────────────────────────────────────────────────────

    def _save(self):
        name = simpledialog.askstring(
            "Save File", "Enter filename (without extension):", parent=self.root)
        if not name or not name.strip():
            return
        name = name.strip()
        if not name.lower().endswith(".txt"):
            name += ".txt"
        path = os.path.join(SCRIPT_DIR, name)

        self._flush_all()
        lines = []
        smap  = self._compute_serials()
        for i, item in enumerate(self.items):
            if item["type"] == "title":
                lines += ["", "=" * 72,
                          f"  {item.get('text','title').upper().center(68)}",
                          "=" * 72, ""]
            elif item["type"] == "header":
                lines.append("\t".join(v.ljust(15) for v in item["values"]))
                lines.append("-" * 72)
            elif item["type"] == "data":
                serial = smap.get(i, "   ")
                row = f"{serial}  " + "\t".join(
                    (item["values"][c].ljust(15) if c < len(item["values"]) else " "*15)
                    for c in range(self.num_cols))
                lines.append(row)

        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        messagebox.showinfo("Saved ✓", f"Saved to:\n{path}")

    def _open(self):
        path = filedialog.askopenfilename(
            initialdir=SCRIPT_DIR,
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title="Open File")
        if not path:
            return
        messagebox.showinfo("Open File",
                            f"Selected:\n{path}\n\n(Import parsing coming soon)")


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("TCombobox",
                    fieldbackground="#2e2e2e", background="#2e2e2e",
                    foreground="#f0f0f0", selectbackground="#444",
                    selectforeground="#fff", arrowcolor="#ccc")
    style.configure("TScrollbar", background="#ccc", troughcolor="#eee")
    MeshTable(root)
    root.mainloop()
