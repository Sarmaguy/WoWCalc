#!/usr/bin/env python3
"""
WoW Crafting Profit Calculator
===============================
Parses Auction House data, manages crafting formulas, and calculates
profit/loss for each craftable item — accounting for AH cut and modifiers.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import sys
import csv
import io
import math
from datetime import datetime

# ---------------------------------------------------------------------------w
# Paths — config files live next to the script (or next to the .exe)
# ---------------------------------------------------------------------------
if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
FORMULAS_FILE = os.path.join(APP_DIR, "formulas.json")
SETTINGS_FILE = os.path.join(APP_DIR, "settings.json")
HISTORY_FILE = os.path.join(APP_DIR, "history.json")

# Oldest snapshots are dropped once the log grows past this.
MAX_SNAPSHOTS = 1000

DEFAULT_SETTINGS = {
    "ah_cut_percent": 5.0,
    "change_highlight_percent": 5.0,
    "modifiers": {
        "multicraft_modifier": 1.0,
    },
}

# ---------------------------------------------------------------------------
# Dark palette — every colour in the UI comes from here
# ---------------------------------------------------------------------------
BG_BASE = "#1F2124"        # window and panel background
BG_SURFACE = "#2A2D31"     # raised surface — buttons, tabs, table headings
BG_SURFACE_HI = "#363A3F"  # the same, hovered
BG_FIELD = "#17191C"       # sunken surface — entries, text box, tables
BORDER = "#3A3F45"
FG_TEXT = "#E4E6EB"
FG_MUTED = "#9BA1A8"       # secondary labels, axis text
FG_DIM = "#6E757C"         # disabled controls, hairlines
ACCENT = "#3E6FA8"         # selection
ACCENT_FG = "#FFFFFF"

PROFIT_FG = "#5FD37A"
LOSS_FG = "#FF7B72"

# Movement indicators — a dark row wash for big movers, a bare arrow for the
# rest, so a table where everything shifted a little still reads calmly.
TINT_UP = "#1B3326"
TINT_DOWN = "#3A2126"
FLAT_EPSILON = 0.05     # |%| below this counts as unchanged
DASH = "—"          # shown when there is nothing to compare against

# Chart scale modes.
MODE_LOG = "Log (gold)"
MODE_LINEAR = "Linear (gold)"
MODE_PCT = "% change"
SCALE_MODES = (MODE_LOG, MODE_LINEAR, MODE_PCT)

# Line colours for the history chart. Light enough to double as the label
# colour in the item picker, against the dark panel behind it.
SERIES_COLORS = (
    "#6CB6FF", "#FF7B72", "#5FD37A", "#FFA657", "#C297FF",
    "#D2A06E", "#FF9BCE", "#AAB4BE", "#E3DC5C", "#5FD8E0",
    "#7C8CE0", "#D07BC0", "#9DBF6A", "#C9A046", "#D97A6A",
    "#8FA8D8", "#B98FD9", "#7FC9A8", "#C4B37E", "#E08A8A",
)

DEFAULT_FORMULAS = [
    {
        "output_item": "Titanium Bar",
        "output_quantity": 1,
        "modifier": None,
        "ingredients": [
            {"item": "Saronite Bar", "quantity": 8},
        ],
    },
    {
        "output_item": "Flask of the Blood Knights Tier 1",
        "output_quantity": 2,
        "modifier": "multicraft_modifier",
        "ingredients": [
            {"item": "Nocturnal Lotus", "quantity": 1},
            {"item": "Sanguithorn Tier 1", "quantity": 6},
            {"item": "Argentleaf Tier 1", "quantity": 8},
            {"item": "Mote of Wild Magic", "quantity": 2},
            {"item": "Sunglass Vial Tier 1", "quantity": 2},
        ],
    },
]

EXAMPLE_DATA = (
    '"Price","Name","Item Level","Owned?","Available"\n'
    '217800,"Mana Lily Tier 1",23,"",88426\n'
    '240000,"Saronite Bar",11,"",4915\n'
    '1860000,"Titanium Bar",12,"Yes",2576\n'
    '15489400,"Nocturnal Lotus",21,"",8773\n'
    '5529700,"Flask of the Blood Knights Tier 1",278,"",9683\n'
    '40000,"Sanguithorn Tier 1",23,"",422310\n'
    '404000,"Argentleaf Tier 1",23,"",74428\n'
    '29800,"Sunglass Vial Tier 1",21,"",116238\n'
    '37800,"Mote of Wild Magic",23,"",564414\n'
)


# ---------------------------------------------------------------------------
# Dark theme
# ---------------------------------------------------------------------------
def dark_titlebar(window):
    """Ask Windows to paint the title bar dark too, so it matches the app.

    Call this once the window has been built: the title bar belongs to a frame
    window Tk only creates when the window is mapped, and DWM re-reads the
    attribute on the next non-client paint — hence the update and the nudge.
    """
    if sys.platform != "win32":
        return
    try:
        import ctypes
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id()) or window.winfo_id()
        on = ctypes.c_int(1)
        for attribute in (20, 19):      # DWMWA_USE_IMMERSIVE_DARK_MODE, pre-20H1
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, attribute, ctypes.byref(on), ctypes.sizeof(on)
            ) == 0:
                break
        NUDGE = 0x0001 | 0x0002 | 0x0004 | 0x0020   # NOSIZE NOMOVE NOZORDER FRAMECHANGED
        ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, NUDGE)
    except Exception:
        pass                            # cosmetic only — never worth crashing


def apply_dark_theme(root):
    """Repaint the whole app dark. Call before any widget is built.

    ttk on Windows defaults to the 'vista' theme, whose widgets are drawn by
    the OS and silently ignore colour options — so switch to 'clam', which Tk
    draws itself and which honours everything set below. Classic tk widgets
    (Canvas, Text, the Combobox drop-down list) are coloured at their own call
    sites or through the option database.
    """
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    root.configure(bg=BG_BASE)

    style.configure(
        ".", background=BG_BASE, foreground=FG_TEXT, fieldbackground=BG_FIELD,
        troughcolor=BG_FIELD, bordercolor=BORDER, lightcolor=BG_SURFACE,
        darkcolor=BG_SURFACE, insertcolor=FG_TEXT, focuscolor=ACCENT,
        selectbackground=ACCENT, selectforeground=ACCENT_FG,
    )
    style.map(".", foreground=[("disabled", FG_DIM)])

    style.configure("TFrame", background=BG_BASE)
    style.configure("TLabel", background=BG_BASE, foreground=FG_TEXT)
    style.configure("Muted.TLabel", background=BG_BASE, foreground=FG_MUTED)
    style.configure("Status.TLabel", background=BG_SURFACE, foreground=FG_MUTED)
    style.configure("TLabelframe", background=BG_BASE, bordercolor=BORDER,
                    lightcolor=BORDER, darkcolor=BORDER)
    style.configure("TLabelframe.Label", background=BG_BASE, foreground=FG_MUTED)
    style.configure("TSeparator", background=BORDER)
    style.configure("TPanedwindow", background=BG_BASE)
    style.configure("Sash", background=BORDER, sashthickness=6, gripcount=0)

    style.configure("TButton", background=BG_SURFACE, foreground=FG_TEXT,
                    bordercolor=BORDER, lightcolor=BG_SURFACE,
                    darkcolor=BG_SURFACE, focusthickness=0, padding=(8, 4))
    style.map(
        "TButton",
        background=[("pressed", ACCENT), ("active", BG_SURFACE_HI),
                    ("disabled", BG_BASE)],
        foreground=[("pressed", ACCENT_FG), ("disabled", FG_DIM)],
        bordercolor=[("active", ACCENT)],
    )

    for widget in ("TEntry", "TCombobox", "TSpinbox"):
        style.configure(widget, fieldbackground=BG_FIELD, foreground=FG_TEXT,
                        background=BG_SURFACE, insertcolor=FG_TEXT,
                        arrowcolor=FG_MUTED, bordercolor=BORDER,
                        lightcolor=BORDER, darkcolor=BORDER, padding=3)
        style.map(
            widget,
            fieldbackground=[("disabled", BG_BASE)],
            foreground=[("disabled", FG_DIM)],
            background=[("active", BG_SURFACE_HI)],
            bordercolor=[("focus", ACCENT)],
            arrowcolor=[("active", FG_TEXT)],
            # A read-only Combobox draws its own text as a selection; without
            # this the closed widget sits there with a blue bar across it.
            selectbackground=[("readonly", BG_FIELD)],
            selectforeground=[("readonly", FG_TEXT)],
        )

    # The Combobox drop-down is a classic Listbox built by Tk itself, so the
    # option database is the only way in.
    for option, value in (
        ("background", BG_FIELD), ("foreground", FG_TEXT),
        ("selectBackground", ACCENT), ("selectForeground", ACCENT_FG),
    ):
        root.option_add("*TCombobox*Listbox." + option, value)

    style.configure("TNotebook", background=BG_BASE, bordercolor=BORDER)
    style.configure("TNotebook.Tab", background=BG_SURFACE, foreground=FG_MUTED,
                    bordercolor=BORDER, lightcolor=BG_SURFACE,
                    darkcolor=BG_SURFACE, padding=(10, 5))
    style.map(
        "TNotebook.Tab",
        background=[("selected", BG_BASE), ("active", BG_SURFACE_HI)],
        foreground=[("selected", FG_TEXT)],
        lightcolor=[("selected", BG_BASE)],
    )

    style.configure("Treeview", background=BG_FIELD, fieldbackground=BG_FIELD,
                    foreground=FG_TEXT, bordercolor=BORDER, rowheight=21)
    style.map("Treeview", background=[("selected", ACCENT)],
              foreground=[("selected", ACCENT_FG)])
    style.configure("Treeview.Heading", background=BG_SURFACE, foreground=FG_TEXT,
                    bordercolor=BORDER, lightcolor=BG_SURFACE,
                    darkcolor=BG_SURFACE, relief="flat", padding=(4, 4))
    style.map("Treeview.Heading", background=[("active", BG_SURFACE_HI)])

    style.configure("TScrollbar", background=BG_SURFACE, troughcolor=BG_BASE,
                    bordercolor=BG_BASE, arrowcolor=FG_MUTED,
                    lightcolor=BG_SURFACE, darkcolor=BG_SURFACE)
    style.map("TScrollbar",
              background=[("pressed", ACCENT), ("active", BG_SURFACE_HI)],
              arrowcolor=[("active", FG_TEXT)])
    return style


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def format_gold(copper):
    """Convert a copper value to a readable  Xg XXs XXc  string."""
    if copper is None:
        return "N/A"
    neg = copper < 0
    c = abs(int(round(copper)))
    g = c // 10000
    s = (c % 10000) // 100
    cp = c % 100
    text = f"{g:,}g {s:02d}s {cp:02d}c"
    return f"-{text}" if neg else text


def format_gold_short(copper):
    """Compact gold label — for chart axes and delta columns."""
    if copper is None:
        return "N/A"
    neg = copper < 0
    g = abs(copper) / 10000.0
    if g >= 100:
        text = f"{g:,.0f}g"
    elif g >= 10:
        text = f"{g:,.1f}g"
    elif g >= 1:
        text = f"{g:,.2f}g"
    else:
        text = f"{abs(copper):,.0f}c"
    return f"-{text}" if neg else text


def _axis_gold(copper):
    """Axis label for a round tick value — no trailing zeros."""
    if copper >= 10000:
        return f"{copper / 10000.0:,.2f}".rstrip("0").rstrip(".") + "g"
    if copper >= 100:
        return f"{copper / 10000.0:,.4f}".rstrip("0").rstrip(".") + "g"
    return f"{copper:,.0f}c"


def pct_change(new, old):
    """Percent move of *new* against *old*, or None when it is undefined."""
    if new is None or not old:
        return None
    return (new - old) / abs(old) * 100.0


def price_change_cell(new, old):
    """(cell text, percent) for a price-movement column."""
    if old is None:
        return ("new" if new is not None else DASH), None
    p = pct_change(new, old)
    if p is None:
        return DASH, None
    if abs(p) < FLAT_EPSILON:
        return DASH, 0.0
    return f"{'▲' if p > 0 else '▼'} {p:+.1f}%", p


def profit_change_cell(new, old):
    """(cell text, percent) for the profit-per-craft movement column.

    Shows the gold swing first — profit can cross zero, which makes a bare
    percentage useless — with the percentage in brackets when it is meaningful.
    """
    if new is None or old is None:
        return DASH, None
    delta = new - old
    if abs(delta) < 100:            # under a silver — noise
        return DASH, 0.0
    p = pct_change(new, old)
    text = f"{'▲' if delta > 0 else '▼'} {format_gold_short(abs(delta))}"
    if p is not None and abs(p) < 1000:
        text += f" ({p:+.1f}%)"
    return text, (p if p is not None else (999.0 if delta > 0 else -999.0))


def change_tag(pct, threshold):
    """Row tag for the mover tint, or "" when the move is unremarkable."""
    if pct is None or abs(pct) < max(threshold, FLAT_EPSILON):
        return ""
    return "up" if pct > 0 else "down"


def _nice_ticks(lo, hi, target=6):
    """Round tick values covering [lo, hi] on a linear axis."""
    span = hi - lo
    if span <= 0:
        return [lo]
    raw = span / target
    mag = 10.0 ** math.floor(math.log10(raw))
    step = 10 * mag
    for m in (1, 2, 2.5, 5):
        if raw <= m * mag:
            step = m * mag
            break
    ticks, v = [], math.ceil(lo / step) * step
    while v <= hi + step * 1e-9 and len(ticks) < 40:
        ticks.append(v)
        v += step
    return ticks


def _y_ticks(mode, lo, hi):
    """[(plotted value, label)] for the chart's vertical axis."""
    if mode == MODE_LOG:
        # 1/2/5 per decade while the span is narrow, decades only beyond that.
        mantissas = (1, 2, 5) if (hi - lo) <= 4 else (1,)
        out = []
        for d in range(math.floor(lo), math.ceil(hi) + 1):
            for m in mantissas:
                raw = m * 10.0 ** d
                pos = math.log10(raw)
                if lo <= pos <= hi:
                    out.append((pos, _axis_gold(raw)))
        return out
    if mode == MODE_PCT:
        return [(v, f"{v:+.0f}%") for v in _nice_ticks(lo, hi)]
    return [(v, _axis_gold(v)) for v in _nice_ticks(lo, hi)]


def _fmt_ts(iso, one_line=False):
    try:
        dt = datetime.fromisoformat(iso)
    except (TypeError, ValueError):
        return str(iso or "?")
    return dt.strftime("%Y-%m-%d %H:%M") if one_line else dt.strftime("%m-%d" + chr(10) + "%H:%M")


def _ellipsis(text, n):
    return text if len(text) <= n else text[: n - 1] + "…"


# ╔═════════════════════════════════════════════════════════════════════════╗
# ║  FORMULA DIALOG                                                        ║
# ╚═════════════════════════════════════════════════════════════════════════╝
class FormulaDialog(tk.Toplevel):
    """Modal dialog for creating / editing a crafting formula."""

    def __init__(self, parent, item_names=None, formula=None, modifier_names=None):
        super().__init__(parent)
        self.title("Edit Formula" if formula else "Add Formula")
        self.geometry("600x500")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.configure(bg=BG_BASE)

        self.result = None
        self.item_names = sorted(item_names or [])
        self.modifier_names = modifier_names or []
        self.ingredient_widgets = []  # list of (frame, combobox, spinbox)

        self._build(formula)

        # Centre on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(0,x)}+{max(0,y)}")
        dark_titlebar(self)

        self.wait_window()

    # ---- build ----
    def _build(self, formula):
        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)

        # --- Output ---
        out = ttk.LabelFrame(main, text="Output (what you craft)", padding=8)
        out.pack(fill="x", pady=(0, 6))

        ttk.Label(out, text="Item:").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.out_item = ttk.Combobox(out, values=self.item_names, width=38)
        self.out_item.grid(row=0, column=1, columnspan=2, sticky="ew", padx=4, pady=2)

        ttk.Label(out, text="Quantity produced per craft:").grid(row=1, column=0, sticky="w", padx=4, pady=2)
        self.out_qty = ttk.Spinbox(out, from_=1, to=100, width=6)
        self.out_qty.grid(row=1, column=1, sticky="w", padx=4, pady=2)
        self.out_qty.set(1)

        ttk.Label(out, text="Modifier:").grid(row=2, column=0, sticky="w", padx=4, pady=2)
        self.mod_var = tk.StringVar(value="None")
        self.mod_combo = ttk.Combobox(
            out, textvariable=self.mod_var,
            values=["None"] + self.modifier_names, width=28,
        )
        self.mod_combo.grid(row=2, column=1, sticky="w", padx=4, pady=2)
        out.columnconfigure(1, weight=1)

        # --- Ingredients ---
        ing_lf = ttk.LabelFrame(main, text="Ingredients (materials consumed)", padding=8)
        ing_lf.pack(fill="both", expand=True, pady=(0, 6))

        canvas_frame = ttk.Frame(ing_lf)
        canvas_frame.pack(fill="both", expand=True)

        self._canvas = tk.Canvas(canvas_frame, highlightthickness=0, height=130,
                                 bg=BG_BASE)
        vsb = ttk.Scrollbar(canvas_frame, orient="vertical", command=self._canvas.yview)
        self._ing_inner = ttk.Frame(self._canvas)

        self._ing_inner.bind(
            "<Configure>",
            lambda _: self._canvas.configure(scrollregion=self._canvas.bbox("all")),
        )
        self._canvas.create_window((0, 0), window=self._ing_inner, anchor="nw")
        self._canvas.configure(yscrollcommand=vsb.set)
        self._canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        ttk.Button(ing_lf, text="+ Add Ingredient", command=lambda: self._add_row()).pack(pady=4)

        # --- Buttons ---
        bf = ttk.Frame(main)
        bf.pack(fill="x")
        ttk.Button(bf, text="Save", command=self._save).pack(side="right", padx=5)
        ttk.Button(bf, text="Cancel", command=self.destroy).pack(side="right")

        # Populate when editing
        if formula:
            self.out_item.set(formula["output_item"])
            self.out_qty.set(formula["output_quantity"])
            self.mod_var.set(formula.get("modifier") or "None")
            for ing in formula["ingredients"]:
                self._add_row(ing["item"], ing["quantity"])
        else:
            self._add_row()

    def _add_row(self, item="", qty=1):
        row = ttk.Frame(self._ing_inner)
        row.pack(fill="x", pady=1)

        cb = ttk.Combobox(row, values=self.item_names, width=32)
        cb.pack(side="left", padx=2)
        cb.set(item)

        ttk.Label(row, text="×").pack(side="left")

        sp = ttk.Spinbox(row, from_=1, to=9999, width=6)
        sp.pack(side="left", padx=2)
        sp.set(qty)

        def _remove(r=row, c=cb, s=sp):
            self.ingredient_widgets = [
                (rr, cc, ss) for rr, cc, ss in self.ingredient_widgets if rr is not r
            ]
            r.destroy()

        ttk.Button(row, text="✕", width=3, command=_remove).pack(side="left", padx=2)
        self.ingredient_widgets.append((row, cb, sp))

    def _save(self):
        name = self.out_item.get().strip()
        if not name:
            messagebox.showwarning("Missing", "Output item name is required.", parent=self)
            return

        try:
            qty = int(self.out_qty.get())
            assert qty >= 1
        except (ValueError, AssertionError):
            messagebox.showwarning("Invalid", "Output quantity must be a positive integer.", parent=self)
            return

        mod = self.mod_var.get()
        mod = None if mod == "None" else mod

        ingredients = []
        for _, cb, sp in self.ingredient_widgets:
            iname = cb.get().strip()
            if not iname:
                continue
            try:
                iqty = int(sp.get())
                assert iqty >= 1
            except (ValueError, AssertionError):
                messagebox.showwarning(
                    "Invalid", f"Quantity for '{iname}' must be a positive integer.", parent=self,
                )
                return
            ingredients.append({"item": iname, "quantity": iqty})

        if not ingredients:
            messagebox.showwarning("Missing", "Add at least one ingredient.", parent=self)
            return

        self.result = {
            "output_item": name,
            "output_quantity": qty,
            "modifier": mod,
            "ingredients": ingredients,
        }
        self.destroy()


# ╔═════════════════════════════════════════════════════════════════════════╗
# ║  SETTINGS DIALOG                                                       ║
# ╚═════════════════════════════════════════════════════════════════════════╝
class SettingsDialog(tk.Toplevel):
    """Modal dialog for AH cut and modifier values."""

    def __init__(self, parent, settings):
        super().__init__(parent)
        self.title("Settings")
        self.geometry("470x470")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.configure(bg=BG_BASE)

        self.result = None
        self.mod_entries = {}

        self._build(settings)

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(0,x)}+{max(0,y)}")
        dark_titlebar(self)

        self.wait_window()

    def _build(self, settings):
        main = ttk.Frame(self, padding=15)
        main.pack(fill="both", expand=True)

        # AH cut
        ahf = ttk.LabelFrame(main, text="Auction House", padding=10)
        ahf.pack(fill="x", pady=(0, 10))
        ttk.Label(ahf, text="AH Cut (%):").grid(row=0, column=0, sticky="w")
        self.ah_entry = ttk.Entry(ahf, width=10)
        self.ah_entry.grid(row=0, column=1, padx=8)
        self.ah_entry.insert(0, str(settings.get("ah_cut_percent", 5.0)))
        ttk.Label(ahf, text="(deducted from revenue when selling)").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(2, 0),
        )

        # Display
        df = ttk.LabelFrame(main, text="Display", padding=10)
        df.pack(fill="x", pady=(0, 10))
        ttk.Label(df, text="Highlight price change ≥ (%):").grid(row=0, column=0, sticky="w")
        self.hl_entry = ttk.Entry(df, width=10)
        self.hl_entry.grid(row=0, column=1, padx=8)
        self.hl_entry.insert(0, str(settings.get("change_highlight_percent", 5.0)))
        ttk.Label(
            df, text="(bigger moves get a tinted row — smaller ones just show ▲ / ▼)",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))

        # Modifiers
        mf = ttk.LabelFrame(main, text="Modifiers  (multiplied with output quantity)", padding=10)
        mf.pack(fill="both", expand=True, pady=(0, 10))

        ttk.Label(mf, text="1.0 = no bonus · 1.15 = 15 % more items on average").pack(anchor="w")

        self._mod_frame = ttk.Frame(mf)
        self._mod_frame.pack(fill="both", expand=True, pady=5)

        for name, val in settings.get("modifiers", {}).items():
            self._add_mod_row(name, val)

        addf = ttk.Frame(mf)
        addf.pack(fill="x", pady=4)
        ttk.Label(addf, text="New modifier name:").pack(side="left")
        self.new_name = ttk.Entry(addf, width=22)
        self.new_name.pack(side="left", padx=5)
        ttk.Button(addf, text="Add", command=self._add_mod).pack(side="left")

        # Buttons
        bf = ttk.Frame(main)
        bf.pack(fill="x")
        ttk.Button(bf, text="Save", command=self._save).pack(side="right", padx=5)
        ttk.Button(bf, text="Cancel", command=self.destroy).pack(side="right")

    def _add_mod_row(self, name, val):
        row = ttk.Frame(self._mod_frame)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text=name, width=28, anchor="w").pack(side="left")
        e = ttk.Entry(row, width=10)
        e.pack(side="left", padx=5)
        e.insert(0, str(val))

        def _remove(r=row, n=name):
            self.mod_entries.pop(n, None)
            r.destroy()

        ttk.Button(row, text="✕", width=3, command=_remove).pack(side="left")
        self.mod_entries[name] = e

    def _add_mod(self):
        name = self.new_name.get().strip()
        if not name:
            return
        if name in self.mod_entries:
            messagebox.showwarning("Exists", f"'{name}' already exists.", parent=self)
            return
        self._add_mod_row(name, 1.0)
        self.new_name.delete(0, "end")

    def _save(self):
        try:
            ah_cut = float(self.ah_entry.get())
        except ValueError:
            messagebox.showwarning("Invalid", "AH cut must be a number.", parent=self)
            return

        try:
            highlight = abs(float(self.hl_entry.get()))
        except ValueError:
            messagebox.showwarning(
                "Invalid", "Highlight threshold must be a number.", parent=self,
            )
            return

        mods = {}
        for name, entry in self.mod_entries.items():
            try:
                mods[name] = float(entry.get())
            except ValueError:
                messagebox.showwarning("Invalid", f"'{name}' value must be a number.", parent=self)
                return

        self.result = {
            "ah_cut_percent": ah_cut,
            "change_highlight_percent": highlight,
            "modifiers": mods,
        }
        self.destroy()


# ╔═════════════════════════════════════════════════════════════════════════╗
# ║  PRICE HISTORY CHART                                                   ║
# ╚═════════════════════════════════════════════════════════════════════════╝
class PriceChart(ttk.Frame):
    """Multi-series price-over-time chart drawn straight onto a tk.Canvas.

    Hand-rolled rather than embedding matplotlib: this app ships as a one-file
    PyInstaller build, and matplotlib would take it from ~10 MB to ~60 MB for a
    handful of line plots.
    """

    BG = BG_FIELD           # plot area — a shade below the panel around it
    GRID = "#2B2F34"
    AXIS = "#4E555C"
    TEXT = FG_MUTED
    HAIR = FG_DIM
    TIP_BG = BG_SURFACE     # hover tooltip
    PAD_L, PAD_R, PAD_T, PAD_B = 92, 20, 18, 50
    HOVER_RADIUS = 40

    def __init__(self, parent, app):
        super().__init__(parent, padding=6)
        self.app = app
        self.snapshots = []
        self.item_vars = {}     # {item name: BooleanVar}
        self.colors = {}        # {item name: hex}
        self._points = []       # [(x, y, name, index, copper)] for hit-testing
        self._hover_ids = []
        self._plot_box = (0, 0, 0, 0)
        self._plot_snaps = []
        self._resize_id = None
        self._build()

    # ---------------------------------------------------------------- build
    def _build(self):
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(0, 6))

        ttk.Label(bar, text="Scale:").pack(side="left")
        self.scale_var = tk.StringVar(value=MODE_LOG)
        scale = ttk.Combobox(
            bar, textvariable=self.scale_var, values=list(SCALE_MODES),
            state="readonly", width=13,
        )
        scale.pack(side="left", padx=(4, 14))
        scale.bind("<<ComboboxSelected>>", lambda _: self._redraw())

        ttk.Label(bar, text="Range:").pack(side="left")
        self.range_var = tk.StringVar(value="All")
        rng = ttk.Combobox(
            bar, textvariable=self.range_var, state="readonly", width=9,
            values=["All", "Last 10", "Last 25", "Last 50", "Last 100"],
        )
        rng.pack(side="left", padx=(4, 14))
        rng.bind("<<ComboboxSelected>>", lambda _: self._redraw())

        ttk.Button(bar, text="Clear History", command=self._clear_history).pack(side="right")
        self.info_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.info_var, style="Muted.TLabel").pack(
            side="right", padx=12,
        )

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(
            body, bg=self.BG, highlightthickness=1, highlightbackground=BORDER,
        )
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<Configure>", self._on_resize)
        self.canvas.bind("<Motion>", self._on_hover)
        self.canvas.bind("<Leave>", lambda _: self._clear_hover())

        # ---- series picker, doubles as the legend ----
        side = ttk.LabelFrame(body, text=" Items ", padding=4)
        side.pack(side="right", fill="y", padx=(6, 0))

        btns = ttk.Frame(side)
        btns.pack(fill="x")
        ttk.Button(btns, text="All", width=5,
                   command=lambda: self._select("all")).pack(side="left")
        ttk.Button(btns, text="None", width=6,
                   command=lambda: self._select("none")).pack(side="left", padx=2)
        ttk.Button(btns, text="Recipes", width=9,
                   command=lambda: self._select("recipes")).pack(side="left")

        wrap = ttk.Frame(side)
        wrap.pack(fill="both", expand=True, pady=(4, 0))
        self._list_canvas = tk.Canvas(
            wrap, width=205, highlightthickness=0, bg=BG_BASE,
        )
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self._list_canvas.yview)
        self._list_inner = ttk.Frame(self._list_canvas)
        self._list_inner.bind(
            "<Configure>",
            lambda _: self._list_canvas.configure(
                scrollregion=self._list_canvas.bbox("all")),
        )
        self._list_canvas.create_window((0, 0), window=self._list_inner, anchor="nw")
        self._list_canvas.configure(yscrollcommand=vsb.set)
        self._list_canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._bind_wheel(self._list_canvas)
        self._bind_wheel(self._list_inner)

    def _bind_wheel(self, widget):
        widget.bind(
            "<MouseWheel>",
            lambda e: self._list_canvas.yview_scroll(int(-e.delta / 120), "units"),
        )

    # ----------------------------------------------------------------- data
    def set_history(self, snapshots):
        self.snapshots = snapshots
        self._rebuild_series_list()
        self._redraw()

    def _recipe_items(self):
        names = set()
        for f in self.app.formulas:
            names.add(f["output_item"])
            names.update(i["item"] for i in f["ingredients"])
        return names

    def _rebuild_series_list(self):
        items = sorted({n for s in self.snapshots for n in s["prices"]})
        self.colors = {
            n: SERIES_COLORS[i % len(SERIES_COLORS)] for i, n in enumerate(items)
        }
        if set(items) == set(self.item_vars):
            return      # same items — keep whatever the user has ticked

        remembered = {n: v.get() for n, v in self.item_vars.items()}
        for child in self._list_inner.winfo_children():
            child.destroy()
        self.item_vars = {}

        # Default to the items that actually appear in a recipe; everything
        # else is usually noise on a shared chart. If the log happens to hold
        # none of them, fall back to showing everything rather than nothing.
        default_on = self._recipe_items() & set(items) or set(items)
        for name in items:
            var = tk.BooleanVar(value=remembered.get(name, name in default_on))
            cb = tk.Checkbutton(
                self._list_inner, text=_ellipsis(name, 26), variable=var,
                anchor="w", fg=self.colors[name], bg=BG_BASE,
                activebackground=BG_BASE, activeforeground=self.colors[name],
                selectcolor=BG_FIELD, highlightthickness=0, bd=0, padx=2,
                font=("Segoe UI", 8), command=self._redraw,
            )
            cb.pack(fill="x")
            self._bind_wheel(cb)
            self.item_vars[name] = var

    def _select(self, which):
        recipes = self._recipe_items()
        for name, var in self.item_vars.items():
            if which == "all":
                var.set(True)
            elif which == "none":
                var.set(False)
            else:
                var.set(name in recipes)
        self._redraw()

    def _clear_history(self):
        if not self.app.history:
            return
        if messagebox.askyesno(
            "Clear History",
            "Delete all %d recorded snapshots?" % len(self.app.history)
            + chr(10) + "This cannot be undone.",
            parent=self,
        ):
            self.app.clear_history()

    def _visible_snapshots(self):
        choice = self.range_var.get()
        if choice.startswith("Last"):
            return self.snapshots[-int(choice.split()[1]):]
        return self.snapshots

    # ----------------------------------------------------------------- draw
    def _on_resize(self, _event):
        if self._resize_id:
            self.after_cancel(self._resize_id)
        self._resize_id = self.after(80, self._redraw)

    def _center_text(self, w, h, text):
        self.canvas.create_text(
            w / 2, h / 2, text=text, fill=FG_MUTED, justify="center",
            font=("Segoe UI", 10), width=max(200, w - 80),
        )

    def _redraw(self):
        c = self.canvas
        c.delete("all")
        self._points = []
        self._hover_ids = []

        n_total = len(self.snapshots)
        self.info_var.set(
            "no snapshots yet" if not n_total
            else "%d snapshot%s" % (n_total, "" if n_total == 1 else "s")
        )

        w, h = c.winfo_width(), c.winfo_height()
        if w < 80 or h < 80:
            return

        snaps = self._visible_snapshots()
        self._plot_snaps = snaps
        if not snaps:
            self._center_text(
                w, h,
                "No price history yet." + chr(10)
                + "Paste Auction House data on the Calculator tab "
                  "and the first snapshot is recorded automatically.",
            )
            return

        selected = [n for n, v in self.item_vars.items() if v.get()]
        if not selected:
            self._center_text(w, h, "No items selected — tick some on the right.")
            return

        mode = self.scale_var.get()
        series = self._build_series(selected, snaps, mode)
        if not series:
            self._center_text(w, h, "The selected items have no recorded prices.")
            return

        values = [p[1] for pts in series.values() for p in pts]
        lo, hi = min(values), max(values)
        pad = (hi - lo) * 0.08 if hi - lo > 1e-9 else (abs(hi) * 0.1 or 1.0)
        lo, hi = lo - pad, hi + pad
        if mode == MODE_PCT:
            lo, hi = min(lo, 0.0), max(hi, 0.0)
        elif mode == MODE_LINEAR:
            lo = min(lo, 0.0)
        if hi - lo < 1e-9:
            hi = lo + 1.0

        x0, y0 = self.PAD_L, self.PAD_T
        x1, y1 = w - self.PAD_R, h - self.PAD_B
        if x1 - x0 < 60 or y1 - y0 < 60:
            return
        self._plot_box = (x0, y0, x1, y1)

        n = len(snaps)
        span_x = x1 - x0
        if n == 1:
            def px(i):
                return x0 + span_x / 2
        else:
            def px(i):
                return x0 + span_x * i / (n - 1)

        def py(v):
            return y1 - (v - lo) / (hi - lo) * (y1 - y0)

        # horizontal gridlines + value labels
        for value, label in _y_ticks(mode, lo, hi):
            y = py(value)
            c.create_line(x0, y, x1, y, fill=self.GRID)
            c.create_text(x0 - 8, y, text=label, anchor="e",
                          fill=self.TEXT, font=("Segoe UI", 8))

        if mode == MODE_PCT and lo <= 0 <= hi:
            c.create_line(x0, py(0), x1, py(0), fill=FG_DIM, dash=(3, 3))

        # vertical gridlines + timestamps
        for i in self._label_indices(n, span_x):
            x = px(i)
            c.create_line(x, y0, x, y1, fill=self.GRID)
            c.create_text(x, y1 + 6, text=_fmt_ts(snaps[i]["ts"]), anchor="n",
                          fill=self.TEXT, font=("Segoe UI", 8), justify="center")

        c.create_line(x0, y0, x0, y1, fill=self.AXIS)
        c.create_line(x0, y1, x1, y1, fill=self.AXIS)

        show_dots = n <= 40
        for name, pts in series.items():
            color = self.colors.get(name, FG_MUTED)
            segment, prev_i = [], None
            for i, value, copper in pts:
                if prev_i is not None and i != prev_i + 1:
                    self._stroke(segment, color, show_dots)   # gap in coverage
                    segment = []
                x, y = px(i), py(value)
                segment.append((x, y))
                self._points.append((x, y, name, i, copper))
                prev_i = i
            self._stroke(segment, color, show_dots)

    def _build_series(self, selected, snaps, mode):
        """{item: [(snapshot index, plotted value, copper), ...]}"""
        series = {}
        for name in selected:
            pts, first = [], None
            for i, snap in enumerate(snaps):
                copper = snap["prices"].get(name)
                if copper is None or copper <= 0:
                    continue        # a log axis has no room for zero / missing
                if mode == MODE_LOG:
                    value = math.log10(copper)
                elif mode == MODE_PCT:
                    first = copper if first is None else first
                    value = (copper / first - 1.0) * 100.0
                else:
                    value = float(copper)
                pts.append((i, value, copper))
            if pts:
                series[name] = pts
        return series

    @staticmethod
    def _label_indices(n, span_x):
        """Snapshot indices to label, keeping ~70 px between timestamps."""
        if n == 1:
            return [0]
        step = max(1, math.ceil(n / max(2, int(span_x // 70))))
        idxs = list(range(0, n, step))
        if idxs[-1] != n - 1:
            # Always label the newest snapshot: extend when there is room,
            # otherwise move the last label rather than crowding two together.
            if n - 1 - idxs[-1] >= step:
                idxs.append(n - 1)
            else:
                idxs[-1] = n - 1
        return idxs

    def _stroke(self, segment, color, show_dots):
        if len(segment) >= 2:
            self.canvas.create_line(
                [v for point in segment for v in point],
                fill=color, width=2, capstyle="round", joinstyle="round",
            )
        if show_dots or len(segment) == 1:
            r = 2.5
            for x, y in segment:
                self.canvas.create_oval(
                    x - r, y - r, x + r, y + r, fill=color, outline=color,
                )

    # ---------------------------------------------------------------- hover
    def _clear_hover(self):
        for item in self._hover_ids:
            self.canvas.delete(item)
        self._hover_ids = []

    def _on_hover(self, event):
        if not self._points:
            return
        best, best_d = None, self.HOVER_RADIUS ** 2
        for point in self._points:
            d = (point[0] - event.x) ** 2 + (point[1] - event.y) ** 2
            if d <= best_d:
                best, best_d = point, d
        if best is None:
            self._clear_hover()
            return
        self._draw_hover(*best)

    def _draw_hover(self, x, y, name, index, copper):
        self._clear_hover()
        c = self.canvas
        x0, y0, x1, y1 = self._plot_box
        snaps = self._plot_snaps
        color = self.colors.get(name, FG_MUTED)

        ids = self._hover_ids
        ids.append(c.create_line(x, y0, x, y1, fill=self.HAIR, dash=(2, 3)))
        ids.append(c.create_oval(x - 4.5, y - 4.5, x + 4.5, y + 4.5,
                                 outline=color, width=2, fill=self.BG))

        previous = snaps[index - 1]["prices"].get(name) if index > 0 else None
        lines = [name, _fmt_ts(snaps[index]["ts"], one_line=True), format_gold(copper)]
        pct = pct_change(copper, previous)
        if pct is not None:
            arrow = "▲" if pct > 0 else ("▼" if pct < 0 else "•")
            lines.append("%s %+.2f%% vs previous" % (arrow, pct))

        text_id = c.create_text(
            0, 0, text=chr(10).join(lines), anchor="nw", justify="left",
            font=("Segoe UI", 8), fill=FG_TEXT,
        )
        bx0, by0, bx1, by1 = c.bbox(text_id)
        tw, th = bx1 - bx0, by1 - by0

        tx, ty = x + 14, y - th - 16
        if tx + tw + 12 > x1:
            tx = x - tw - 26
        tx = max(tx, x0 + 2)
        if ty < y0:
            ty = y + 16
        ty = min(ty, y1 - th - 12)

        c.coords(text_id, tx + 6, ty + 5)
        box_id = c.create_rectangle(
            tx, ty, tx + tw + 12, ty + th + 10, fill=self.TIP_BG, outline=color,
        )
        c.tag_lower(box_id, text_id)
        ids.extend([box_id, text_id])


# ╔═════════════════════════════════════════════════════════════════════════╗
# ║  MAIN APPLICATION                                                      ║
# ╚═════════════════════════════════════════════════════════════════════════╝
class WoWCraftCalc:
    """Top-level application controller."""

    def __init__(self, root):
        self.root = root
        self.root.title("WoW Crafting Profit Calculator")
        self.root.geometry("1250x820")
        self.root.minsize(950, 620)

        self.item_prices = {}       # {item_name: price_copper}  — current parse
        self.item_available = {}    # {item_name: "Available" column, as text}
        self.baseline_prices = {}   # prices from the snapshot before this one
        self.history = []           # [{"ts": iso, "prices": {...}}, ...]
        self.formulas = []
        self.settings = {}
        self._parse_note = ""       # appended to the next status message

        self._load_settings()
        self._load_formulas()
        self._load_history()
        self._build_ui()
        self._refresh_formulas_tree()
        self.chart.set_history(self.history)

    # ------------------------------------------------------------------ IO
    def _load_settings(self):
        try:
            with open(SETTINGS_FILE, encoding="utf-8") as fh:
                self.settings = json.load(fh)
        except (FileNotFoundError, json.JSONDecodeError):
            self.settings = json.loads(json.dumps(DEFAULT_SETTINGS))

    def _save_settings(self):
        with open(SETTINGS_FILE, "w", encoding="utf-8") as fh:
            json.dump(self.settings, fh, indent=2)

    def _load_formulas(self):
        try:
            with open(FORMULAS_FILE, encoding="utf-8") as fh:
                self.formulas = json.load(fh)
        except (FileNotFoundError, json.JSONDecodeError):
            self.formulas = json.loads(json.dumps(DEFAULT_FORMULAS))
            self._save_formulas()

    def _save_formulas(self):
        with open(FORMULAS_FILE, "w", encoding="utf-8") as fh:
            json.dump(self.formulas, fh, indent=2)

    def _load_history(self):
        try:
            with open(HISTORY_FILE, encoding="utf-8") as fh:
                data = json.load(fh)
            snaps = data.get("snapshots", []) if isinstance(data, dict) else data
            self.history = [
                s for s in snaps
                if isinstance(s, dict) and isinstance(s.get("prices"), dict) and s["prices"]
            ]
        except (OSError, json.JSONDecodeError, AttributeError, TypeError):
            self.history = []

    def _save_history(self):
        payload = {"version": 1, "snapshots": self.history}
        try:
            with open(HISTORY_FILE, "w", encoding="utf-8") as fh:
                json.dump(payload, fh)
        except OSError as exc:
            self.status_var.set(f"Could not save price history: {exc}")

    def _record_snapshot(self):
        """Log the current parse, unless it repeats the most recent snapshot.

        Auto-parse re-runs on every edit to the paste box, so identical
        re-parses are dropped — otherwise typing would bury the chart under
        duplicate points. Either way the delta baseline is refreshed to the
        snapshot *before* the current one.
        """
        added = False
        if self.item_prices:
            last = self.history[-1]["prices"] if self.history else None
            if last != self.item_prices:
                self.history.append({
                    "ts": datetime.now().isoformat(timespec="seconds"),
                    "prices": dict(self.item_prices),
                })
                del self.history[:-MAX_SNAPSHOTS]
                self._save_history()
                added = True

        self.baseline_prices = (
            self.history[-2]["prices"] if len(self.history) >= 2 else {}
        )
        return added

    def clear_history(self):
        self.history = []
        self.baseline_prices = {}
        self._save_history()
        self.chart.set_history(self.history)
        self._refresh_items_tree()
        self._auto_calculate()
        self.status_var.set("Price history cleared.")

    def _highlight_threshold(self):
        try:
            return abs(float(self.settings.get("change_highlight_percent", 5.0)))
        except (TypeError, ValueError):
            return 5.0

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=(6, 0))

        calc_tab = ttk.Frame(self.notebook)
        self.notebook.add(calc_tab, text="  Calculator  ")

        hist_tab = ttk.Frame(self.notebook)
        self.notebook.add(hist_tab, text="  Price History  ")

        self._build_calculator_tab(calc_tab)

        self.chart = PriceChart(hist_tab, self)
        self.chart.pack(fill="both", expand=True)

        # Status bar — shared by both tabs
        self.status_var = tk.StringVar(
            value="Ready — paste Auction House data and click  Parse Data"
        )
        ttk.Label(
            self.root, textvariable=self.status_var, style="Status.TLabel",
            anchor="w", padding=(6, 3),
        ).pack(fill="x", side="bottom")

    def _build_calculator_tab(self, parent):
        # Main vertical split
        vpane = ttk.PanedWindow(parent, orient="vertical")
        vpane.pack(fill="both", expand=True, padx=6, pady=6)

        # ==================== TOP — AH data input ====================
        top = ttk.LabelFrame(vpane, text="  Auction House Data  ", padding=5)

        toolbar = ttk.Frame(top)
        toolbar.pack(fill="x", pady=(0, 4))
        ttk.Button(toolbar, text="📋 Parse Data", command=self._parse_data).pack(side="left", padx=3)
        ttk.Button(toolbar, text="🗑 Clear All", command=self._clear_all).pack(side="left", padx=3)
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(toolbar, text="⚙ Settings", command=self._open_settings).pack(side="left", padx=3)
        ttk.Button(
            toolbar, text="📊 Calculate Profits", command=self._calculate,
        ).pack(side="right", padx=3)

        # A plain Text plus a ttk scrollbar rather than ScrolledText: the
        # classic tk.Scrollbar that builds is drawn by Windows itself and stays
        # white whatever colours it is handed.
        text_wrap = ttk.Frame(top)
        text_wrap.pack(fill="both", expand=True)
        self.data_text = tk.Text(
            text_wrap, height=9, font=("Consolas", 9), wrap="none",
            bg=BG_FIELD, fg=FG_TEXT, insertbackground=FG_TEXT,
            selectbackground=ACCENT, selectforeground=ACCENT_FG,
            relief="flat", borderwidth=0, highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT,
        )
        text_vsb = ttk.Scrollbar(text_wrap, orient="vertical",
                                 command=self.data_text.yview)
        self.data_text.configure(yscrollcommand=text_vsb.set)
        self.data_text.pack(side="left", fill="both", expand=True)
        text_vsb.pack(side="right", fill="y")
        self.data_text.insert("1.0", EXAMPLE_DATA)
        # <<Modified>> only fires on a False->True flip, and the insert above
        # already set the flag — without this reset the first paste is silent.
        self.data_text.edit_modified(False)

        # Auto-parse when text changes (paste, type, delete)
        self._text_debounce_id = None
        self.data_text.bind("<<Modified>>", self._on_text_modified)

        vpane.add(top, weight=1)

        # ==================== MIDDLE — Items + Formulas ====================
        hpane = ttk.PanedWindow(vpane, orient="horizontal")

        # ---------- Parsed items ----------
        items_lf = ttk.LabelFrame(hpane, text="  Parsed Items  ", padding=5)
        item_cols = ("name", "price_copper", "price_gold", "available", "change")
        self.items_tree = ttk.Treeview(items_lf, columns=item_cols, show="headings", height=7)
        for cid, txt, w in [
            ("name", "Item Name", 210),
            ("price_copper", "Price (copper)", 105),
            ("price_gold", "Price (gold)", 130),
            ("available", "Available", 80),
            ("change", "vs Last", 95),
        ]:
            self.items_tree.heading(cid, text=txt)
            self.items_tree.column(cid, width=w, minwidth=50)
        self.items_tree.column("change", anchor="e")
        self.items_tree.tag_configure("up", background=TINT_UP)
        self.items_tree.tag_configure("down", background=TINT_DOWN)
        _add_scrollbar(items_lf, self.items_tree)
        self.items_tree.bind("<Double-1>", self._on_item_double_click)
        hpane.add(items_lf, weight=1)

        # ---------- Formulas ----------
        form_lf = ttk.LabelFrame(hpane, text="  Crafting Formulas  ", padding=5)
        ftb = ttk.Frame(form_lf)
        ftb.pack(fill="x", pady=(0, 4))
        ttk.Button(ftb, text="Add", command=self._add_formula).pack(side="left", padx=2)
        ttk.Button(ftb, text="Edit", command=self._edit_formula).pack(side="left", padx=2)
        ttk.Button(ftb, text="Delete", command=self._delete_formula).pack(side="left", padx=2)

        self.formulas_tree = ttk.Treeview(form_lf, columns=("desc",), show="headings", height=7)
        self.formulas_tree.heading("desc", text="Formula")
        self.formulas_tree.column("desc", width=450, minwidth=150)
        _add_scrollbar(form_lf, self.formulas_tree)
        self.formulas_tree.bind("<Double-1>", lambda _: self._edit_formula())
        hpane.add(form_lf, weight=1)

        vpane.add(hpane, weight=1)

        # ==================== BOTTOM — Results ====================
        res_lf = ttk.LabelFrame(vpane, text="  Profit Analysis  ", padding=5)

        res_cols = (
            "item", "market", "market_chg", "craft_cost", "eff_qty",
            "cost_unit", "revenue_unit", "profit_unit", "profit_craft",
            "profit_chg", "margin",
        )
        self.results_tree = ttk.Treeview(res_lf, columns=res_cols, show="headings", height=10)
        for cid, txt, w in [
            # Trimmed from the original widths so the two movement columns
            # fit on screen at the default window size instead of hiding
            # behind the horizontal scrollbar.
            ("item", "Item", 205),
            ("market", "Market Price", 104),
            ("market_chg", "Market vs Last", 95),
            ("craft_cost", "Total Craft Cost", 112),
            ("eff_qty", "Eff. Output", 70),
            ("cost_unit", "Cost / Unit", 104),
            ("revenue_unit", "Revenue / Unit", 106),
            ("profit_unit", "Profit / Unit", 104),
            ("profit_craft", "Profit / Craft", 108),
            ("profit_chg", "Profit vs Last", 124),
            ("margin", "Margin", 64),
        ]:
            self.results_tree.heading(cid, text=txt)
            self.results_tree.column(cid, width=w, minwidth=50)
        for cid in ("market_chg", "profit_chg"):
            self.results_tree.column(cid, anchor="e")

        # Horizontal scroll for the wide table
        xsb = ttk.Scrollbar(res_lf, orient="horizontal", command=self.results_tree.xview)
        ysb = ttk.Scrollbar(res_lf, orient="vertical", command=self.results_tree.yview)
        self.results_tree.configure(xscrollcommand=xsb.set, yscrollcommand=ysb.set)
        self.results_tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        res_lf.rowconfigure(0, weight=1)
        res_lf.columnconfigure(0, weight=1)

        # Foreground still carries profit/loss; the tinted background carries the
        # move since the last snapshot. Treeview resolves one tag per row, so
        # every combination gets its own.
        for state, fg in (("profit", PROFIT_FG), ("loss", LOSS_FG)):
            self.results_tree.tag_configure(state, foreground=fg)
            self.results_tree.tag_configure(f"{state}_up", foreground=fg, background=TINT_UP)
            self.results_tree.tag_configure(f"{state}_down", foreground=fg, background=TINT_DOWN)

        vpane.add(res_lf, weight=2)

    # ------------------------------------------------------ Auto-parse
    def _on_text_modified(self, event=None):
        """Debounce text changes and auto-parse after 500ms."""
        if not self.data_text.edit_modified():
            return
        self.data_text.edit_modified(False)
        if self._text_debounce_id:
            self.root.after_cancel(self._text_debounce_id)
        self._text_debounce_id = self.root.after(500, self._parse_data)

    # -------------------------------------------------------------- Parse
    def _parse_data(self, event=None):
        raw = self.data_text.get("1.0", "end").strip()
        if not raw:
            self.status_var.set("Paste Auction House data into the text area first.")
            return

        prices, available = {}, {}
        try:
            for row in csv.DictReader(io.StringIO(raw)):
                name = (row.get("Name") or "").strip()
                if not name:
                    continue
                prices[name] = int(float(row.get("Price", 0)))
                available[name] = row.get("Available", "")
        except Exception as exc:
            self.status_var.set(f"Parse Error: {exc}")
            return

        self.item_prices = prices
        self.item_available = available
        added = self._record_snapshot()

        if added:
            self.chart.set_history(self.history)
            self._parse_note = f"   New snapshot #{len(self.history)} recorded."
        else:
            self._parse_note = "   (no price change since last snapshot)"

        self._refresh_items_tree()
        if self.formulas:
            self._calculate()       # its status line carries the note instead
        else:
            self.status_var.set(f"Parsed {len(prices)} items.{self._take_note()}")

    def _take_note(self):
        """Pop the pending snapshot note so it is shown exactly once."""
        note, self._parse_note = self._parse_note, ""
        return note

    def _clear_all(self):
        self.data_text.delete("1.0", "end")
        self.item_prices.clear()
        self.item_available.clear()
        self._clear_tree(self.items_tree)
        self._clear_tree(self.results_tree)
        self.status_var.set("Cleared.  (Recorded price history is kept.)")

    # -------------------------------------------------------- Parsed items
    def _item_row_values(self, name):
        """(column values, percent move) for one row of the Parsed Items table."""
        price = self.item_prices.get(name)
        if self.baseline_prices:
            text, pct = price_change_cell(price, self.baseline_prices.get(name))
        else:
            text, pct = DASH, None      # nothing recorded before this parse
        values = (
            name,
            f"{price:,}",
            format_gold(price),
            self.item_available.get(name, ""),
            text,
        )
        return values, pct

    def _refresh_items_tree(self):
        self._clear_tree(self.items_tree)
        threshold = self._highlight_threshold()
        for name in self.item_prices:
            values, pct = self._item_row_values(name)
            tag = change_tag(pct, threshold)
            self.items_tree.insert("", "end", values=values, tags=(tag,) if tag else ())

    # ----------------------------------------------------------- Formulas
    @staticmethod
    def _formula_label(f):
        mod = f" × {f['modifier']}" if f.get("modifier") else ""
        ings = " + ".join(f"{i['quantity']}×{i['item']}" for i in f["ingredients"])
        return f"{f['output_quantity']}×{f['output_item']}{mod}  ←  {ings}"

    def _refresh_formulas_tree(self):
        self._clear_tree(self.formulas_tree)
        for f in self.formulas:
            self.formulas_tree.insert("", "end", values=(self._formula_label(f),))

    def _modifier_names(self):
        return list(self.settings.get("modifiers", {}).keys())

    def _add_formula(self):
        dlg = FormulaDialog(
            self.root,
            item_names=list(self.item_prices.keys()),
            modifier_names=self._modifier_names(),
        )
        if dlg.result:
            self.formulas.append(dlg.result)
            self._save_formulas()
            self._refresh_formulas_tree()
            self._auto_calculate()

    def _edit_formula(self):
        sel = self.formulas_tree.selection()
        if not sel:
            self.status_var.set("Select a formula to edit first.")
            return
        idx = self.formulas_tree.index(sel[0])
        dlg = FormulaDialog(
            self.root,
            item_names=list(self.item_prices.keys()),
            formula=self.formulas[idx],
            modifier_names=self._modifier_names(),
        )
        if dlg.result:
            self.formulas[idx] = dlg.result
            self._save_formulas()
            self._refresh_formulas_tree()
            self._auto_calculate()

    def _delete_formula(self):
        sel = self.formulas_tree.selection()
        if not sel:
            self.status_var.set("Select a formula to delete first.")
            return
        idx = self.formulas_tree.index(sel[0])
        name = self.formulas[idx]["output_item"]
        if messagebox.askyesno("Confirm Delete", f"Delete formula for  \"{name}\"?"):
            del self.formulas[idx]
            self._save_formulas()
            self._refresh_formulas_tree()
            self._auto_calculate()

    # ----------------------------------------------------------- Settings
    def _open_settings(self):
        dlg = SettingsDialog(self.root, self.settings)
        if dlg.result:
            self.settings = dlg.result
            self._save_settings()
            self.status_var.set(
                f"Settings saved.  AH cut = {self.settings['ah_cut_percent']}%  |  "
                + "  |  ".join(
                    f"{k} = {v}" for k, v in self.settings.get("modifiers", {}).items()
                )
            )
            self._auto_calculate()

    # ------------------------------------------------------- Inline Editing
    def _on_item_double_click(self, event):
        """Allow editing price by double-clicking a row in Parsed Items."""
        item_id = self.items_tree.identify_row(event.y)
        column = self.items_tree.identify_column(event.x)
        if not item_id or column not in ("#2", "#3"):
            return  # only allow editing price columns

        # Get bounding box of the cell
        bbox = self.items_tree.bbox(item_id, column)
        if not bbox:
            return

        values = self.items_tree.item(item_id, "values")
        current_copper = int(values[1].replace(",", ""))

        # Create entry widget over the cell
        entry = ttk.Entry(self.items_tree, width=15)
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        entry.insert(0, str(current_copper))
        entry.select_range(0, "end")
        entry.focus_set()

        def _commit(e=None):
            try:
                new_price = int(float(entry.get()))
            except ValueError:
                entry.destroy()
                return
            entry.destroy()
            name = values[0]
            self.item_prices[name] = new_price
            new_values, pct = self._item_row_values(name)
            tag = change_tag(pct, self._highlight_threshold())
            self.items_tree.item(
                item_id, values=new_values, tags=(tag,) if tag else (),
            )
            self._auto_calculate()

        def _cancel(e=None):
            entry.destroy()

        entry.bind("<Return>", _commit)
        entry.bind("<FocusOut>", _commit)
        entry.bind("<Escape>", _cancel)

    # --------------------------------------------------------- Auto Calculate
    def _auto_calculate(self):
        """Recalculate profits automatically if data is available."""
        if self.item_prices and self.formulas:
            self._calculate()

    # --------------------------------------------------------- Calculation
    def _evaluate(self, prices):
        """Run every formula against *prices*.

        Returns ({formula index: result}, missing item names). Keyed by index
        rather than item name so two recipes for the same output stay distinct.
        """
        ah_cut = self.settings.get("ah_cut_percent", 5.0) / 100.0
        modifiers = self.settings.get("modifiers", {})

        results, missing = {}, set()
        for idx, formula in enumerate(self.formulas):
            out_item = formula["output_item"]
            out_qty = formula["output_quantity"]
            mod_name = formula.get("modifier")

            mod_val = modifiers.get(mod_name, 1.0) if mod_name else 1.0

            if out_item not in prices:
                missing.add(out_item)
                continue

            market = prices[out_item]

            # Ingredient cost
            ing_cost = 0
            ok = True
            for ing in formula["ingredients"]:
                if ing["item"] not in prices:
                    missing.add(ing["item"])
                    ok = False
                    break
                ing_cost += prices[ing["item"]] * ing["quantity"]

            if not ok:
                continue

            eff_output = out_qty * mod_val
            cost_per_unit = ing_cost / eff_output if eff_output else 0
            revenue_per_unit = market * (1 - ah_cut)
            profit_per_unit = revenue_per_unit - cost_per_unit
            profit_per_craft = profit_per_unit * eff_output
            margin = (profit_per_unit / revenue_per_unit * 100) if revenue_per_unit else 0

            results[idx] = {
                "item": out_item,
                "market": market,
                "craft_cost": ing_cost,
                "eff_qty": eff_output,
                "cost_unit": cost_per_unit,
                "revenue_unit": revenue_per_unit,
                "profit_unit": profit_per_unit,
                "profit_craft": profit_per_craft,
                "margin": margin,
            }
        return results, missing

    def _calculate(self):
        if not self.item_prices:
            self.status_var.set("No AH data — parse Auction House data first.")
            return
        if not self.formulas:
            self.status_var.set("No formulas — add at least one crafting formula.")
            return

        results, missing = self._evaluate(self.item_prices)

        # Same recipes re-priced at the previous snapshot, so the profit delta
        # also picks up ingredient moves — not just the output's market price.
        baseline = (
            self._evaluate(self.baseline_prices)[0] if self.baseline_prices else {}
        )

        ordered = sorted(
            results.items(), key=lambda kv: kv[1]["profit_craft"], reverse=True,
        )
        threshold = self._highlight_threshold()

        # Render
        self._clear_tree(self.results_tree)
        for idx, r in ordered:
            prev = baseline.get(idx)
            eff = r["eff_qty"]
            eff_str = f"{eff:.2f}" if eff != int(eff) else str(int(eff))

            market_txt = DASH
            if self.baseline_prices:
                market_txt = price_change_cell(
                    r["market"], self.baseline_prices.get(r["item"]),
                )[0]
            profit_txt, profit_pct = profit_change_cell(
                r["profit_craft"], prev["profit_craft"] if prev else None,
            )

            state = "profit" if r["profit_craft"] >= 0 else "loss"
            move = change_tag(profit_pct, threshold)
            tag = f"{state}_{move}" if move else state

            self.results_tree.insert("", "end", values=(
                r["item"],
                format_gold(r["market"]),
                market_txt,
                format_gold(r["craft_cost"]),
                eff_str,
                format_gold(r["cost_unit"]),
                format_gold(r["revenue_unit"]),
                format_gold(r["profit_unit"]),
                format_gold(r["profit_craft"]),
                profit_txt,
                f"{r['margin']:+.1f}%",
            ), tags=(tag,))

        n_profit = sum(1 for _, r in ordered if r["profit_craft"] >= 0)
        missing_note = f"  [Skipped: {', '.join(sorted(missing))}]" if missing else ""
        self.status_var.set(
            f"Done — {len(ordered)} recipes evaluated:  "
            f"{n_profit} profitable,  {len(ordered) - n_profit} unprofitable.   "
            f"(AH cut {self.settings['ah_cut_percent']}%)   "
            f"[{len(self.history)} snapshots]{self._take_note()}{missing_note}"
        )

    # ------------------------------------------------------------- Helpers
    @staticmethod
    def _clear_tree(tree):
        for item in tree.get_children():
            tree.delete(item)


def _add_scrollbar(parent, tree):
    """Pack a Treeview with a vertical scrollbar inside *parent*."""
    vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    tree.pack(side="left", fill="both", expand=True)
    vsb.pack(side="right", fill="y")


# ╔═════════════════════════════════════════════════════════════════════════╗
# ║  ENTRY POINT                                                           ║
# ╚═════════════════════════════════════════════════════════════════════════╝
def main():
    root = tk.Tk()
    apply_dark_theme(root)
    WoWCraftCalc(root)
    dark_titlebar(root)
    root.mainloop()


if __name__ == "__main__":
    main()
