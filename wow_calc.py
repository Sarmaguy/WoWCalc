#!/usr/bin/env python3
"""
WoW Crafting Profit Calculator
===============================
Parses Auction House data, manages crafting formulas, and calculates
profit/loss for each craftable item — accounting for AH cut and modifiers.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import json
import os
import sys
import csv
import io

# ---------------------------------------------------------------------------w
# Paths — config files live next to the script (or next to the .exe)
# ---------------------------------------------------------------------------
if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
FORMULAS_FILE = os.path.join(APP_DIR, "formulas.json")
SETTINGS_FILE = os.path.join(APP_DIR, "settings.json")

DEFAULT_SETTINGS = {
    "ah_cut_percent": 5.0,
    "modifiers": {
        "multicraft_modifier": 1.0,
    },
}

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

        self._canvas = tk.Canvas(canvas_frame, highlightthickness=0, height=130)
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
        self.geometry("450x400")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.result = None
        self.mod_entries = {}

        self._build(settings)

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(0,x)}+{max(0,y)}")

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

        mods = {}
        for name, entry in self.mod_entries.items():
            try:
                mods[name] = float(entry.get())
            except ValueError:
                messagebox.showwarning("Invalid", f"'{name}' value must be a number.", parent=self)
                return

        self.result = {"ah_cut_percent": ah_cut, "modifiers": mods}
        self.destroy()


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

        self.item_prices = {}   # {item_name: price_copper}
        self.formulas = []
        self.settings = {}

        self._load_settings()
        self._load_formulas()
        self._build_ui()
        self._refresh_formulas_tree()

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

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        # Main vertical split
        vpane = ttk.PanedWindow(self.root, orient="vertical")
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

        self.data_text = scrolledtext.ScrolledText(top, height=9, font=("Consolas", 9), wrap="none")
        self.data_text.pack(fill="both", expand=True)
        self.data_text.insert("1.0", EXAMPLE_DATA)

        # Auto-parse when text changes (paste, type, delete)
        self._text_debounce_id = None
        self.data_text.bind("<<Modified>>", self._on_text_modified)

        vpane.add(top, weight=1)

        # ==================== MIDDLE — Items + Formulas ====================
        hpane = ttk.PanedWindow(vpane, orient="horizontal")

        # ---------- Parsed items ----------
        items_lf = ttk.LabelFrame(hpane, text="  Parsed Items  ", padding=5)
        item_cols = ("name", "price_copper", "price_gold", "available")
        self.items_tree = ttk.Treeview(items_lf, columns=item_cols, show="headings", height=7)
        for cid, txt, w in [
            ("name", "Item Name", 210),
            ("price_copper", "Price (copper)", 105),
            ("price_gold", "Price (gold)", 130),
            ("available", "Available", 80),
        ]:
            self.items_tree.heading(cid, text=txt)
            self.items_tree.column(cid, width=w, minwidth=50)
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
            "item", "market", "craft_cost", "eff_qty",
            "cost_unit", "revenue_unit", "profit_unit", "profit_craft", "margin",
        )
        self.results_tree = ttk.Treeview(res_lf, columns=res_cols, show="headings", height=10)
        for cid, txt, w in [
            ("item", "Item", 210),
            ("market", "Market Price", 120),
            ("craft_cost", "Total Craft Cost", 120),
            ("eff_qty", "Eff. Output", 80),
            ("cost_unit", "Cost / Unit", 120),
            ("revenue_unit", "Revenue / Unit", 120),
            ("profit_unit", "Profit / Unit", 120),
            ("profit_craft", "Profit / Craft", 120),
            ("margin", "Margin", 70),
        ]:
            self.results_tree.heading(cid, text=txt)
            self.results_tree.column(cid, width=w, minwidth=50)

        # Horizontal scroll for the wide table
        xsb = ttk.Scrollbar(res_lf, orient="horizontal", command=self.results_tree.xview)
        ysb = ttk.Scrollbar(res_lf, orient="vertical", command=self.results_tree.yview)
        self.results_tree.configure(xscrollcommand=xsb.set, yscrollcommand=ysb.set)
        self.results_tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        res_lf.rowconfigure(0, weight=1)
        res_lf.columnconfigure(0, weight=1)

        self.results_tree.tag_configure("profit", foreground="#1B8C1B")
        self.results_tree.tag_configure("loss", foreground="#CC2222")

        vpane.add(res_lf, weight=2)

        # Status bar
        self.status_var = tk.StringVar(value="Ready — paste Auction House data and click  Parse Data")
        ttk.Label(
            self.root, textvariable=self.status_var,
            relief="sunken", anchor="w", padding=(6, 3),
        ).pack(fill="x", side="bottom")

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

        self.item_prices.clear()
        self._clear_tree(self.items_tree)

        try:
            reader = csv.DictReader(io.StringIO(raw))
            count = 0
            for row in reader:
                name = row.get("Name", "").strip()
                price = int(float(row.get("Price", 0)))
                avail = row.get("Available", "")
                if name:
                    self.item_prices[name] = price
                    self.items_tree.insert("", "end", values=(
                        name, f"{price:,}", format_gold(price), avail,
                    ))
                    count += 1
            self.status_var.set(f"Parsed {count} items.")
            self._auto_calculate()
        except Exception as exc:
            self.status_var.set(f"Parse Error: {exc}")

    def _clear_all(self):
        self.data_text.delete("1.0", "end")
        self.item_prices.clear()
        self._clear_tree(self.items_tree)
        self._clear_tree(self.results_tree)
        self.status_var.set("Cleared.")

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
            self.items_tree.item(item_id, values=(
                name, f"{new_price:,}", format_gold(new_price), values[3],
            ))
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
    def _calculate(self):
        if not self.item_prices:
            self.status_var.set("No AH data — parse Auction House data first.")
            return
        if not self.formulas:
            self.status_var.set("No formulas — add at least one crafting formula.")
            return

        ah_cut = self.settings.get("ah_cut_percent", 5.0) / 100.0
        modifiers = self.settings.get("modifiers", {})

        results = []
        missing = set()

        for formula in self.formulas:
            out_item = formula["output_item"]
            out_qty = formula["output_quantity"]
            mod_name = formula.get("modifier")

            mod_val = modifiers.get(mod_name, 1.0) if mod_name else 1.0

            if out_item not in self.item_prices:
                missing.add(out_item)
                continue

            market = self.item_prices[out_item]

            # Ingredient cost
            ing_cost = 0
            ok = True
            for ing in formula["ingredients"]:
                if ing["item"] not in self.item_prices:
                    missing.add(ing["item"])
                    ok = False
                    break
                ing_cost += self.item_prices[ing["item"]] * ing["quantity"]

            if not ok:
                continue

            eff_output = out_qty * mod_val
            cost_per_unit = ing_cost / eff_output if eff_output else 0
            revenue_per_unit = market * (1 - ah_cut)
            profit_per_unit = revenue_per_unit - cost_per_unit
            profit_per_craft = profit_per_unit * eff_output
            margin = (profit_per_unit / revenue_per_unit * 100) if revenue_per_unit else 0

            results.append({
                "item": out_item,
                "market": market,
                "craft_cost": ing_cost,
                "eff_qty": eff_output,
                "cost_unit": cost_per_unit,
                "revenue_unit": revenue_per_unit,
                "profit_unit": profit_per_unit,
                "profit_craft": profit_per_craft,
                "margin": margin,
            })

        # Sort descending by profit per craft
        results.sort(key=lambda r: r["profit_craft"], reverse=True)

        # Render
        self._clear_tree(self.results_tree)
        for r in results:
            tag = "profit" if r["profit_craft"] >= 0 else "loss"
            eff = r["eff_qty"]
            eff_str = f"{eff:.2f}" if eff != int(eff) else str(int(eff))

            self.results_tree.insert("", "end", values=(
                r["item"],
                format_gold(r["market"]),
                format_gold(r["craft_cost"]),
                eff_str,
                format_gold(r["cost_unit"]),
                format_gold(r["revenue_unit"]),
                format_gold(r["profit_unit"]),
                format_gold(r["profit_craft"]),
                f"{r['margin']:+.1f}%",
            ), tags=(tag,))

        n_profit = sum(1 for r in results if r["profit_craft"] >= 0)
        missing_note = f"  [Skipped: {', '.join(sorted(missing))}]" if missing else ""
        self.status_var.set(
            f"Done — {len(results)} recipes evaluated:  "
            f"{n_profit} profitable,  {len(results) - n_profit} unprofitable.   "
            f"(AH cut {self.settings['ah_cut_percent']}%){missing_note}"
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
    WoWCraftCalc(root)
    root.mainloop()


if __name__ == "__main__":
    main()
