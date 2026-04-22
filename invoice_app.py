#!/usr/bin/env python3
"""
Invoice Generator — Portable App
SQLite-backed contacts & line items database
Auto-installs: Pillow, reportlab
Run: python invoice_app.py
"""

import sys, os, subprocess, importlib, importlib.util

def ensure_deps():
    deps = {'Pillow': 'PIL', 'reportlab': 'reportlab'}
    missing = [pkg for pkg, mod in deps.items() if importlib.util.find_spec(mod) is None]
    if missing:
        print(f"[Setup] Installing: {', '.join(missing)} ...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install'] + missing + ['--quiet'])

ensure_deps()

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import datetime, json, sqlite3, pathlib

from PIL import Image, ImageTk
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, Image as RLImage, HRFlowable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER, TA_LEFT


# ─────────────────────────────────────────────────────────────────
#  DATABASE
# ─────────────────────────────────────────────────────────────────
class InvoiceDB:
    FIELDS = ['Name / Company', 'Address Line 1', 'Address Line 2',
              'City, State, Postcode', 'Phone', 'Email']

    def __init__(self):
        db_path = pathlib.Path(__file__).parent / 'invoice_data.db'
        self.conn = sqlite3.connect(str(db_path))
        self.conn.row_factory = sqlite3.Row
        self._init_schema()

    def _init_schema(self):
        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS contacts (
                id       INTEGER PRIMARY KEY AUTOINCREMENT,
                label    TEXT NOT NULL UNIQUE,
                name     TEXT, address1 TEXT, address2 TEXT,
                city     TEXT, phone    TEXT, email    TEXT,
                updated  TEXT DEFAULT (datetime('now'))
            );
            CREATE TABLE IF NOT EXISTS line_items (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                description TEXT NOT NULL UNIQUE,
                rate        REAL DEFAULT 0,
                used_count  INTEGER DEFAULT 1,
                updated     TEXT DEFAULT (datetime('now'))
            );
        """)
        self.conn.commit()

    # contacts
    def save_contact(self, fields: dict):
        label = (fields.get('Name / Company') or '').strip()
        if not label:
            return
        row = self.conn.execute("SELECT id FROM contacts WHERE label=?", (label,)).fetchone()
        vals = (fields.get('Name / Company',''), fields.get('Address Line 1',''),
                fields.get('Address Line 2',''), fields.get('City, State, Postcode',''),
                fields.get('Phone',''), fields.get('Email',''))
        if row:
            self.conn.execute(
                "UPDATE contacts SET name=?,address1=?,address2=?,city=?,phone=?,email=?,updated=datetime('now') WHERE label=?",
                (*vals, label))
        else:
            self.conn.execute(
                "INSERT INTO contacts (label,name,address1,address2,city,phone,email) VALUES (?,?,?,?,?,?,?)",
                (label, *vals))
        self.conn.commit()

    def get_contact_names(self):
        return [r['label'] for r in self.conn.execute(
            "SELECT label FROM contacts ORDER BY updated DESC").fetchall()]

    def get_contact(self, label: str):
        row = self.conn.execute("SELECT * FROM contacts WHERE label=?", (label,)).fetchone()
        if not row:
            return None
        return {'Name / Company': row['name'] or '',
                'Address Line 1': row['address1'] or '',
                'Address Line 2': row['address2'] or '',
                'City, State, Postcode': row['city'] or '',
                'Phone': row['phone'] or '',
                'Email': row['email'] or ''}

    def delete_contact(self, label: str):
        self.conn.execute("DELETE FROM contacts WHERE label=?", (label,))
        self.conn.commit()

    def all_contacts(self):
        return self.conn.execute("SELECT * FROM contacts ORDER BY updated DESC").fetchall()

    # line items
    def save_line_item(self, description: str, rate: float):
        desc = description.strip()
        if not desc:
            return
        if self.conn.execute("SELECT id FROM line_items WHERE description=?", (desc,)).fetchone():
            self.conn.execute(
                "UPDATE line_items SET rate=?,used_count=used_count+1,updated=datetime('now') WHERE description=?",
                (rate, desc))
        else:
            self.conn.execute("INSERT INTO line_items (description,rate) VALUES (?,?)", (desc, rate))
        self.conn.commit()

    def search_line_items(self, query: str):
        return [{'description': r['description'], 'rate': r['rate']}
                for r in self.conn.execute(
                    "SELECT description,rate FROM line_items WHERE description LIKE ? ORDER BY used_count DESC LIMIT 8",
                    (f'%{query}%',)).fetchall()]

    def all_line_items(self):
        return self.conn.execute(
            "SELECT * FROM line_items ORDER BY used_count DESC").fetchall()

    def delete_line_item(self, desc: str):
        self.conn.execute("DELETE FROM line_items WHERE description=?", (desc,))
        self.conn.commit()

    def close(self):
        self.conn.close()


# ─────────────────────────────────────────────────────────────────
#  AUTOCOMPLETE ENTRY
# ─────────────────────────────────────────────────────────────────
class AutocompleteEntry(tk.Entry):
    def __init__(self, master, suggestion_fn, on_select=None, placeholder='', **kwargs):
        self._var = tk.StringVar()
        super().__init__(master, textvariable=self._var, **kwargs)
        self._suggestion_fn = suggestion_fn
        self._on_select     = on_select
        self._placeholder   = placeholder
        self._popup         = None
        self._listbox       = None
        self._skip_trace    = False

        if placeholder:
            self._put_placeholder()
            self.bind('<FocusIn>',  self._focus_in)
            self.bind('<FocusOut>', self._focus_out)

        self._var.trace_add('write', self._on_change)
        self.bind('<Down>',   self._to_list)
        self.bind('<Escape>', self._close)
        self.bind('<Return>', self._close)

    def _put_placeholder(self):
        self._skip_trace = True
        self._var.set(self._placeholder)
        self.config(fg='#9CA3AF')
        self._skip_trace = False

    def _focus_in(self, _):
        if self._var.get() == self._placeholder:
            self._skip_trace = True
            self._var.set('')
            self.config(fg='#1F2937')
            self._skip_trace = False

    def _focus_out(self, _):
        if not self._var.get().strip():
            self._put_placeholder()
        self._close()

    def _on_change(self, *_):
        if self._skip_trace:
            return
        q = self._var.get().strip()
        if q and q != self._placeholder:
            results = self._suggestion_fn(q)
            if results:
                self._show(results); return
        self._close()

    def _show(self, items):
        self._close()
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 1
        w = max(self.winfo_width(), 240)
        h = min(len(items) * 26 + 4, 200)

        self._popup = tk.Toplevel(self)
        self._popup.wm_overrideredirect(True)
        self._popup.wm_geometry(f"{w}x{h}+{x}+{y}")
        self._popup.lift()

        frm = tk.Frame(self._popup, bg='white',
                       highlightbackground='#D1D5DB', highlightthickness=1)
        frm.pack(fill='both', expand=True)
        sb = tk.Scrollbar(frm, orient='vertical')
        self._listbox = tk.Listbox(frm, yscrollcommand=sb.set,
                                    font=('Segoe UI', 9), relief='flat',
                                    bg='white', fg='#1F2937', bd=0,
                                    activestyle='none',
                                    selectbackground='#EFF6FF',
                                    selectforeground='#1D4ED8')
        sb.config(command=self._listbox.yview)
        sb.pack(side='right', fill='y')
        self._listbox.pack(fill='both', expand=True)
        for item in items:
            self._listbox.insert('end', f"  {item}")
        self._listbox.bind('<ButtonRelease-1>', self._pick)
        self._listbox.bind('<Return>', self._pick)

    def _to_list(self, _):
        if self._listbox:
            self._listbox.focus_set()
            self._listbox.selection_set(0)

    def _pick(self, _):
        if not self._listbox:
            return
        sel = self._listbox.curselection()
        if sel:
            val = self._listbox.get(sel[0]).strip()
            self._skip_trace = True
            self._var.set(val)
            self.config(fg='#1F2937')
            self._skip_trace = False
            self._close()
            if self._on_select:
                self._on_select(val)

    def _close(self, *_):
        if self._popup:
            self._popup.destroy()
            self._popup = None
            self._listbox = None

    def get_value(self):
        v = self._var.get().strip()
        return '' if v == self._placeholder else v

    def set_value(self, v: str):
        self._skip_trace = True
        if v:
            self._var.set(v)
            self.config(fg='#1F2937')
        else:
            self._put_placeholder()
        self._skip_trace = False


# ─────────────────────────────────────────────────────────────────
#  MAIN APP
# ─────────────────────────────────────────────────────────────────
class InvoiceApp:
    BG        = "#F0F4F8"
    CARD      = "#FFFFFF"
    ACCENT    = "#1D4ED8"
    TEXT      = "#1F2937"
    MUTED     = "#6B7280"
    BORDER    = "#D1D5DB"
    DANGER    = "#DC2626"
    DANGER_BG = "#FEE2E2"
    HEADER    = "#0F172A"
    STRIPE    = "#F8FAFC"

    ADDR_FIELDS = ['Name / Company', 'Address Line 1', 'Address Line 2',
                   'City, State, Postcode', 'Phone', 'Email']

    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Generator")
        self.root.geometry("1020x860")
        self.root.minsize(820, 680)
        self.root.configure(bg=self.BG)
        try:
            self.root.iconbitmap(default='')
        except Exception:
            pass

        self.db = InvoiceDB()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.logo_path     = None
        self.logo_photoimg = None
        self.item_rows     = []
        self.tax_enabled   = tk.BooleanVar(value=False)
        self.tax_rate      = tk.StringVar(value="6")
        self.currency      = tk.StringVar(value="RM")

        self._setup_styles()
        self._build_ui()

    def _on_close(self):
        self.db.close()
        self.root.destroy()

    def _setup_styles(self):
        s = ttk.Style()
        s.theme_use('clam')
        s.configure('TFrame',       background=self.BG)
        s.configure('TLabel',       background=self.BG, foreground=self.TEXT, font=('Segoe UI', 10))
        s.configure('TCheckbutton', background=self.CARD, foreground=self.TEXT, font=('Segoe UI', 10))
        s.map('TCheckbutton',       background=[('active', self.CARD)])
        s.configure('Horizontal.TSeparator', background=self.BORDER)

    # ── UI scaffold ───────────────────────────────────────────────
    def _build_ui(self):
        self._build_topbar()
        outer  = tk.Frame(self.root, bg=self.BG)
        outer.pack(fill='both', expand=True)
        canvas = tk.Canvas(outer, bg=self.BG, highlightthickness=0)
        vsb    = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        self._sf   = tk.Frame(canvas, bg=self.BG)
        self._cwin = canvas.create_window((0,0), window=self._sf, anchor='nw')
        self._sf.bind('<Configure>',    lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>',      lambda e: canvas.itemconfig(self._cwin, width=e.width))
        canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))
        self._build_body(self._sf)
        self._statusvar = tk.StringVar(value="Ready — fill in details and Export PDF")
        tk.Label(self.root, textvariable=self._statusvar, bg=self.HEADER, fg='#94A3B8',
                 font=('Segoe UI', 8), anchor='w', padx=14).pack(fill='x', side='bottom')

    def _build_topbar(self):
        bar = tk.Frame(self.root, bg=self.HEADER, height=52)
        bar.pack(fill='x')
        bar.pack_propagate(False)
        tk.Label(bar, text="📄  Invoice Generator", bg=self.HEADER, fg='white',
                 font=('Segoe UI', 14, 'bold')).pack(side='left', padx=18, pady=10)
        tk.Label(bar, text="Currency:", bg=self.HEADER, fg='#94A3B8',
                 font=('Segoe UI', 9)).pack(side='right', padx=(0,4))
        ttk.Combobox(bar, textvariable=self.currency, width=6,
                     values=["RM","USD","SGD","EUR","GBP","JPY"],
                     font=('Segoe UI', 9), state='readonly'
                     ).pack(side='right', padx=(0,14))
        for label, cmd, bg, fg in [
            ("💾 Export PDF",  self._export_pdf,      "#1D4ED8", "white"),
            ("📋 Save JSON",   self._save_json,       "#065F46", "white"),
            ("📂 Load JSON",   self._load_json,       "#92400E", "white"),
            ("🗄 Manage DB",   self._open_db_manager, "#374151", "#E5E7EB"),
            ("🗑 Clear",       self._clear_all,       "#3F3F46", "#FCA5A5"),
        ]:
            tk.Button(bar, text=label, command=cmd, bg=bg, fg=fg,
                      font=('Segoe UI', 9, 'bold'), relief='flat', cursor='hand2',
                      padx=10, pady=6, activebackground=bg, activeforeground=fg
                      ).pack(side='right', padx=(0,5), pady=8)

    def _build_body(self, parent):
        p = {'padx': 18, 'pady': 6}
        row1 = tk.Frame(parent, bg=self.BG)
        row1.pack(fill='x', **p)
        # meta MUST be packed first (side='right') before logo (side='left')
        self._build_meta_card(row1)
        self._build_logo_card(row1)

        row2 = tk.Frame(parent, bg=self.BG)
        row2.pack(fill='x', **p)
        self._build_address_card(row2, "FROM  (Your Details)", 'from').pack(
            side='left', fill='both', expand=True, padx=(0,8))
        self._build_address_card(row2, "BILL TO  (Client Details)", 'to').pack(
            side='left', fill='both', expand=True)

        self._build_items_card(parent, p)
        self._build_totals_card(parent, p)

    # ── Logo ──────────────────────────────────────────────────────
    def _build_logo_card(self, parent):
        card = self._card(parent)
        card.pack(side='left', fill='y', padx=(0,8), ipadx=10, ipady=8)
        tk.Label(card, text="COMPANY LOGO", bg=self.CARD, fg=self.MUTED,
                 font=('Segoe UI', 8)).pack(pady=(8,4), padx=10)
        self._logo_frame = tk.Frame(card, bg='#F1F5F9', width=170, height=110,
                                     highlightbackground=self.BORDER, highlightthickness=1)
        self._logo_frame.pack(padx=10)
        self._logo_frame.pack_propagate(False)
        self._logo_lbl = tk.Label(self._logo_frame, text="Click to upload\nlogo image",
                                   bg='#F1F5F9', fg=self.MUTED, font=('Segoe UI', 9), cursor='hand2')
        self._logo_lbl.pack(fill='both', expand=True)
        self._logo_lbl.bind('<Button-1>', lambda e: self._upload_logo())
        bf = tk.Frame(card, bg=self.CARD)
        bf.pack(fill='x', padx=10, pady=(6,4))
        self._btn(bf, "⬆  Upload Logo", self._upload_logo, self.ACCENT, 'white').pack(fill='x', ipady=3)
        self._btn(bf, "✕  Remove", self._remove_logo, '#F1F5F9', self.DANGER).pack(fill='x', pady=(4,0), ipady=2)

    # ── Meta ──────────────────────────────────────────────────────
    def _build_meta_card(self, parent):
        outer = tk.Frame(parent, bg=self.CARD,
                         highlightbackground=self.BORDER, highlightthickness=1)
        outer.pack(side='right', fill='y', padx=(8, 0))

        gf = tk.Frame(outer, bg=self.CARD)
        gf.pack(padx=20, pady=14)

        self._meta = {}
        today = datetime.date.today()
        for i, (lbl, key, default) in enumerate([
            ("Invoice #:", 'inv_num', "INV-001"),
            ("Date:",       'date',   today.strftime('%d/%m/%Y')),
            ("Due Date:",   'due',    (today + datetime.timedelta(days=30)).strftime('%d/%m/%Y')),
            ("PO Number:",  'po',     ""),
        ]):
            tk.Label(gf, text=lbl, bg=self.CARD, fg=self.MUTED,
                     font=('Segoe UI', 9, 'bold')).grid(row=i, column=0, sticky='e', padx=(0, 8), pady=4)
            v = tk.StringVar(value=default)
            self._meta[key] = v
            tk.Entry(gf, textvariable=v, width=18, font=('Segoe UI', 10),
                     relief='flat', bg='#F1F5F9', fg=self.TEXT, bd=0,
                     highlightbackground=self.BORDER, highlightthickness=1
                     ).grid(row=i, column=1, sticky='w', pady=4, ipady=5, padx=2)


    # ── Address card ──────────────────────────────────────────────
    def _build_address_card(self, parent, title, prefix):
        card = self._card(parent)
        hdr  = tk.Frame(card, bg=self.CARD)
        hdr.pack(fill='x', padx=12, pady=(10,4))
        tk.Label(hdr, text=title, bg=self.CARD, fg=self.ACCENT,
                 font=('Segoe UI', 9, 'bold')).pack(side='left')

        # Save & Delete buttons
        self._btn(hdr, "💾 Save Contact",
                  lambda p=prefix: self._save_contact(p),
                  '#EFF6FF', self.ACCENT).pack(side='right', padx=(4,0), ipady=2, ipadx=4)
        self._btn(hdr, "🗑 Delete",
                  lambda p=prefix: self._delete_contact_from_card(p),
                  self.DANGER_BG, self.DANGER).pack(side='right', ipady=2, ipadx=4)

        addr_entries = {}
        addr_vars    = {}

        for i, field in enumerate(self.ADDR_FIELDS):
            if i == 0:
                # Name row: autocomplete entry + ▼ Select button
                name_row = tk.Frame(card, bg=self.CARD)
                name_row.pack(fill='x', padx=12, pady=2)
                name_row.columnconfigure(0, weight=1)

                e = AutocompleteEntry(
                    name_row,
                    suggestion_fn=lambda q: [n for n in self.db.get_contact_names()
                                             if q.lower() in n.lower()],
                    on_select=lambda val, p=prefix: self._load_contact_by_name(val, p),
                    placeholder=field,
                    font=('Segoe UI', 10), relief='flat',
                    bg='#F8FAFC', bd=0,
                    highlightbackground=self.BORDER, highlightthickness=1
                )
                e.grid(row=0, column=0, sticky='ew', ipady=5)

                tk.Button(
                    name_row,
                    text="▼ Select",
                    command=lambda p=prefix: self._show_contact_picker(p),
                    bg=self.ACCENT, fg='white',
                    font=('Segoe UI', 8, 'bold'),
                    relief='flat', cursor='hand2', padx=8,
                    activebackground=self.ACCENT, activeforeground='white'
                ).grid(row=0, column=1, padx=(4,0), ipady=5, sticky='ns')

                addr_entries[field] = e
                addr_vars[field]    = e._var
            else:
                v  = tk.StringVar(value=field)
                e2 = tk.Entry(card, textvariable=v, font=('Segoe UI', 10),
                              relief='flat', bg='#F8FAFC', fg=self.MUTED, bd=0,
                              highlightbackground=self.BORDER, highlightthickness=1)
                e2.pack(fill='x', padx=12, pady=2, ipady=5)

                def on_in(ev, ew=e2, vw=v, ph=field):
                    if vw.get() == ph:
                        vw.set(''); ew.config(fg=self.TEXT)
                def on_out(ev, ew=e2, vw=v, ph=field):
                    if not vw.get().strip():
                        vw.set(ph); ew.config(fg=self.MUTED)

                e2.bind('<FocusIn>',  on_in)
                e2.bind('<FocusOut>', on_out)
                addr_entries[field] = e2
                addr_vars[field]    = v

        setattr(self, f'_{prefix}_entries', addr_entries)
        setattr(self, f'_{prefix}_vars',    addr_vars)
        tk.Frame(card, bg=self.CARD, height=6).pack()
        return card

    def _show_contact_picker(self, prefix):
        """Open a floating panel listing all saved contacts for selection."""
        names = self.db.get_contact_names()
        if not names:
            messagebox.showinfo("No Contacts",
                "No contacts saved yet.\n\nFill in the fields and click '💾 Save Contact' first.")
            return

        # Anchor position: root centre
        rx = self.root.winfo_rootx() + self.root.winfo_width()  // 2
        ry = self.root.winfo_rooty() + self.root.winfo_height() // 2

        pop = tk.Toplevel(self.root)
        pop.title("Select Contact")
        pop.geometry(f"360x420+{rx-180}+{ry-210}")
        pop.configure(bg=self.BG)
        pop.resizable(False, True)
        pop.grab_set()
        pop.lift()

        # Search bar
        search_frame = tk.Frame(pop, bg=self.HEADER)
        search_frame.pack(fill='x')
        tk.Label(search_frame, text="👥  Select a Contact",
                 bg=self.HEADER, fg='white',
                 font=('Segoe UI', 10, 'bold')).pack(side='left', padx=12, pady=10)

        sv = tk.StringVar()
        search_e = tk.Entry(search_frame, textvariable=sv,
                            font=('Segoe UI', 9), relief='flat',
                            bg='#1E293B', fg='white',
                            insertbackground='white',
                            highlightbackground='#334155', highlightthickness=1)
        search_e.pack(side='right', padx=10, pady=8, ipady=4, fill='x', expand=True)
        tk.Label(search_frame, text="🔍", bg=self.HEADER, fg='#94A3B8',
                 font=('Segoe UI', 10)).pack(side='right')

        # List
        list_frame = tk.Frame(pop, bg=self.CARD,
                              highlightbackground=self.BORDER, highlightthickness=1)
        list_frame.pack(fill='both', expand=True, padx=10, pady=8)

        sb = tk.Scrollbar(list_frame, orient='vertical')
        lb = tk.Listbox(list_frame, yscrollcommand=sb.set,
                        font=('Segoe UI', 10), relief='flat',
                        bg=self.CARD, fg=self.TEXT, bd=0,
                        activestyle='none',
                        selectbackground=self.ACCENT,
                        selectforeground='white',
                        highlightthickness=0)
        sb.config(command=lb.yview)
        sb.pack(side='right', fill='y')
        lb.pack(fill='both', expand=True, padx=4, pady=4)

        def populate(query=''):
            lb.delete(0, 'end')
            for name in names:
                if query.lower() in name.lower():
                    lb.insert('end', f"  {name}")

        populate()
        sv.trace_add('write', lambda *_: populate(sv.get()))

        def do_select():
            sel = lb.curselection()
            if sel:
                chosen = lb.get(sel[0]).strip()
                self._load_contact_by_name(chosen, prefix)
                pop.destroy()

        lb.bind('<Double-1>',     lambda e: do_select())
        lb.bind('<Return>',       lambda e: do_select())
        search_e.bind('<Return>', lambda e: do_select())
        search_e.focus_set()

        # Buttons
        btn_row = tk.Frame(pop, bg=self.BG)
        btn_row.pack(fill='x', padx=10, pady=(0,10))
        self._btn(btn_row, "✓  Load Selected", do_select,
                  self.ACCENT, 'white').pack(side='left', ipady=6, ipadx=10)
        self._btn(btn_row, "✕  Cancel", pop.destroy,
                  '#F1F5F9', self.MUTED).pack(side='right', ipady=6, ipadx=10)

    def _get_addr_data(self, prefix):
        entries = getattr(self, f'_{prefix}_entries')
        vars_d  = getattr(self, f'_{prefix}_vars')
        result  = {}
        for field in self.ADDR_FIELDS:
            if isinstance(entries[field], AutocompleteEntry):
                result[field] = entries[field].get_value()
            else:
                v = vars_d[field].get()
                result[field] = '' if v == field else v
        return result

    def _set_addr_data(self, prefix, data):
        entries = getattr(self, f'_{prefix}_entries')
        vars_d  = getattr(self, f'_{prefix}_vars')
        for field in self.ADDR_FIELDS:
            val = data.get(field, '')
            if isinstance(entries[field], AutocompleteEntry):
                entries[field].set_value(val)
            else:
                if val:
                    vars_d[field].set(val)
                    entries[field].config(fg=self.TEXT)
                else:
                    vars_d[field].set(field)
                    entries[field].config(fg=self.MUTED)

    def _save_contact(self, prefix):
        data = self._get_addr_data(prefix)
        name = data.get('Name / Company', '').strip()
        if not name:
            messagebox.showwarning("Save Contact", "Name / Company is required.")
            return
        self.db.save_contact(data)
        self._set_status(f"✅ Contact saved: {name}")

    def _delete_contact_from_card(self, prefix):
        data = self._get_addr_data(prefix)
        name = data.get('Name / Company', '').strip()
        if not name:
            messagebox.showwarning("Delete", "Enter a name first.")
            return
        if not self.db.get_contact(name):
            messagebox.showinfo("Delete", f'"{name}" is not in the database.')
            return
        if messagebox.askyesno("Delete Contact", f'Delete "{name}" from database?'):
            self.db.delete_contact(name)
            self._set_status(f"🗑 Deleted: {name}")

    def _load_contact_by_name(self, name, prefix):
        contact = self.db.get_contact(name)
        if contact:
            self._set_addr_data(prefix, contact)
            self._set_status(f"📋 Loaded: {name}")

    # ── Items card ────────────────────────────────────────────────
    def _build_items_card(self, parent, p):
        card = self._card(parent)
        card.pack(fill='x', **p)
        hdr  = tk.Frame(card, bg=self.HEADER)
        hdr.pack(fill='x')
        for col, (txt, w, anc) in enumerate([
            ("#", 4, 'center'), ("Description", 0, 'w'),
            ("Qty", 8, 'center'), ("Unit Price", 14, 'e'), ("Amount", 14, 'e'), ("", 4, 'center')
        ]):
            kw = {'width': w} if w else {}
            tk.Label(hdr, text=txt, bg=self.HEADER, fg='white',
                     font=('Segoe UI', 9, 'bold'), anchor=anc, **kw
                     ).grid(row=0, column=col, padx=6, pady=7, sticky='ew' if not w else '')
        hdr.columnconfigure(1, weight=1)
        self._items_cont = tk.Frame(card, bg=self.CARD)
        self._items_cont.pack(fill='x')
        self._items_cont.columnconfigure(1, weight=1)
        for _ in range(3):
            self._add_row()
        self._btn(card, "＋  Add Line Item", self._add_row, '#EFF6FF', self.ACCENT
                  ).pack(fill='x', ipady=7)

    def _add_row(self):
        idx  = len(self.item_rows)
        bg   = self.CARD if idx % 2 == 0 else self.STRIPE
        rf   = tk.Frame(self._items_cont, bg=bg)
        rf.grid(row=idx, column=0, columnspan=6, sticky='ew')
        rf.columnconfigure(1, weight=1)

        num = tk.Label(rf, text=str(idx+1), bg=bg, fg=self.MUTED,
                       font=('Segoe UI', 9), width=4, anchor='center')
        num.grid(row=0, column=0, padx=6, pady=4)

        # Description with autocomplete
        desc_e = AutocompleteEntry(
            rf,
            suggestion_fn=lambda q: [r['description'] for r in self.db.search_line_items(q)],
            on_select=None,
            font=('Segoe UI', 10), relief='flat',
            bg=bg, bd=0,
            highlightbackground=self.BORDER, highlightthickness=1
        )
        desc_e.grid(row=0, column=1, padx=4, pady=4, ipady=5, sticky='ew')

        def mk_entry(col, w, just='right'):
            e = tk.Entry(rf, font=('Segoe UI', 10), relief='flat',
                         bg=bg, fg=self.TEXT, bd=0, justify=just, width=w,
                         highlightbackground=self.BORDER, highlightthickness=1)
            e.grid(row=0, column=col, padx=4, pady=4, ipady=5)
            return e

        qty_e  = mk_entry(2, 8,  'center')
        rate_e = mk_entry(3, 14, 'right')
        qty_e.insert(0, "1")
        rate_e.insert(0, "0.00")

        amt_lbl = tk.Label(rf, text="0.00", width=14, anchor='e',
                           font=('Segoe UI', 10, 'bold'), bg=bg, fg=self.TEXT)
        amt_lbl.grid(row=0, column=4, padx=6, pady=4)

        row_data = {'frame': rf, 'num_lbl': num,
                    'desc': desc_e, 'qty': qty_e,
                    'rate': rate_e, 'amt_lbl': amt_lbl}
        self.item_rows.append(row_data)

        # When description picked from dropdown → auto-fill its saved rate
        def on_desc_pick(val, rd=row_data):
            for m in self.db.search_line_items(val):
                if m['description'] == val:
                    rd['rate'].delete(0, 'end')
                    rd['rate'].insert(0, f"{m['rate']:.2f}")
                    self._recalc()
                    break

        desc_e._on_select = on_desc_pick

        def del_row(rd=row_data):
            rd['frame'].destroy()
            self.item_rows.remove(rd)
            self._renumber()
            self._recalc()

        tk.Button(rf, text="✕", command=del_row,
                  bg=self.DANGER_BG, fg=self.DANGER,
                  font=('Segoe UI', 9, 'bold'), relief='flat', cursor='hand2', width=3
                  ).grid(row=0, column=5, padx=6, pady=4)

        for w in (qty_e, rate_e):
            w.bind('<KeyRelease>', lambda e: self._recalc())
        return row_data

    def _renumber(self):
        for i, rd in enumerate(self.item_rows):
            rd['num_lbl'].config(text=str(i+1))
            rd['frame'].grid(row=i)

    # ── Totals ────────────────────────────────────────────────────
    def _build_totals_card(self, parent, p):
        card = self._card(parent)
        card.pack(fill='x', **p, ipadx=12, ipady=12)
        outer = tk.Frame(card, bg=self.CARD)
        outer.pack(fill='x', padx=12, pady=8)

        nf = tk.Frame(outer, bg=self.CARD)
        nf.pack(side='left', fill='both', expand=True, padx=(0,24))
        tk.Label(nf, text="NOTES / PAYMENT INFO", bg=self.CARD, fg=self.MUTED,
                 font=('Segoe UI', 8, 'bold')).pack(anchor='w', pady=(0,4))
        self._notes = tk.Text(nf, height=7, font=('Segoe UI', 9),
                               bg='#F8FAFC', fg=self.TEXT, relief='flat', wrap='word',
                               highlightbackground=self.BORDER, highlightthickness=1, bd=0)
        self._notes.pack(fill='both', expand=True)
        self._notes.insert('1.0', "Bank: Maybank\nAccount No: 1234-5678-9012\nRef: Invoice #")

        tf = tk.Frame(outer, bg=self.CARD, width=300)
        tf.pack(side='right', fill='y')
        tf.pack_propagate(False)

        def tot_row(lbl_text, attr, bold=False):
            r = tk.Frame(tf, bg=self.CARD)
            r.pack(fill='x', pady=2)
            fw = 'bold' if bold else 'normal'
            tk.Label(r, text=lbl_text, bg=self.CARD, fg=self.TEXT,
                     font=('Segoe UI', 10, fw)).pack(side='left')
            l = tk.Label(r, text=f"{self.currency.get()} 0.00",
                          bg=self.CARD, fg=self.TEXT, font=('Segoe UI', 10, fw))
            l.pack(side='right')
            setattr(self, attr, l)

        tot_row("Subtotal:", '_lbl_sub')

        tax_row = tk.Frame(tf, bg=self.CARD)
        tax_row.pack(fill='x', pady=4)
        tk.Checkbutton(tax_row, text="Add Tax", variable=self.tax_enabled,
                        bg=self.CARD, fg=self.TEXT, font=('Segoe UI', 10),
                        activebackground=self.CARD, cursor='hand2',
                        command=self._recalc).pack(side='left')
        tk.Entry(tax_row, textvariable=self.tax_rate, width=5, justify='center',
                 font=('Segoe UI', 10), relief='flat', bg='#F1F5F9', fg=self.TEXT, bd=0,
                 highlightbackground=self.BORDER, highlightthickness=1
                 ).pack(side='left', padx=4, ipady=3)
        tk.Label(tax_row, text="%", bg=self.CARD, fg=self.MUTED,
                 font=('Segoe UI', 10)).pack(side='left')
        self._lbl_tax = tk.Label(tax_row, text=f"{self.currency.get()} 0.00",
                                  bg=self.CARD, fg=self.TEXT, font=('Segoe UI', 10))
        self._lbl_tax.pack(side='right')
        self.tax_rate.trace_add('write', lambda *_: self._recalc())

        ttk.Separator(tf, orient='horizontal').pack(fill='x', pady=6)

        grand = tk.Frame(tf, bg=self.ACCENT)
        grand.pack(fill='x', ipady=10)
        tk.Label(grand, text="TOTAL", bg=self.ACCENT, fg='white',
                 font=('Segoe UI', 14, 'bold')).pack(side='left', padx=14)
        self._lbl_total = tk.Label(grand, text=f"{self.currency.get()} 0.00",
                                    bg=self.ACCENT, fg='white',
                                    font=('Segoe UI', 15, 'bold'))
        self._lbl_total.pack(side='right', padx=14)

    def _recalc(self):
        cur = self.currency.get()
        sub = 0.0
        for rd in self.item_rows:
            try:
                amt = float(rd['qty'].get() or 0) * float(rd['rate'].get() or 0)
            except ValueError:
                amt = 0.0
            rd['amt_lbl'].config(text=f"{amt:,.2f}")
            sub += amt
        tax = 0.0
        if self.tax_enabled.get():
            try:
                tax = sub * float(self.tax_rate.get() or 0) / 100
            except ValueError:
                pass
        self._lbl_sub.config(text=f"{cur} {sub:,.2f}")
        self._lbl_tax.config(text=f"{cur} {tax:,.2f}")
        self._lbl_total.config(text=f"{cur} {sub+tax:,.2f}")

    # ── Auto-save contacts + items to DB ──────────────────────────
    def _autosave_to_db(self):
        for prefix in ('from', 'to'):
            data = self._get_addr_data(prefix)
            if data.get('Name / Company', '').strip():
                self.db.save_contact(data)
        for rd in self.item_rows:
            desc = rd['desc'].get_value() if isinstance(rd['desc'], AutocompleteEntry) \
                   else rd['desc'].get().strip()
            if desc:
                try:
                    rate = float(rd['rate'].get() or 0)
                except ValueError:
                    rate = 0.0
                self.db.save_line_item(desc, rate)

    # ── Logo ──────────────────────────────────────────────────────
    def _upload_logo(self):
        path = filedialog.askopenfilename(
            filetypes=[("Image files","*.png *.jpg *.jpeg *.gif *.bmp *.webp"),("All","*.*")])
        if not path:
            return
        try:
            img = Image.open(path).convert("RGBA")
            img.thumbnail((170, 110), Image.LANCZOS)
            self.logo_photoimg = ImageTk.PhotoImage(img)
            self._logo_lbl.config(image=self.logo_photoimg, text='', bg=self.CARD)
            self.logo_path = path
            self._set_status(f"Logo: {os.path.basename(path)}")
        except Exception as ex:
            messagebox.showerror("Logo Error", str(ex))

    def _remove_logo(self):
        self.logo_path = None; self.logo_photoimg = None
        self._logo_lbl.config(image='', text="Click to upload\nlogo image", bg='#F1F5F9')

    # ── PDF ───────────────────────────────────────────────────────
    def _export_pdf(self):
        self._autosave_to_db()
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF","*.pdf")],
            initialfile=f"Invoice_{self._meta['inv_num'].get()}.pdf")
        if not path:
            return
        try:
            self._generate_pdf(path)
            self._set_status(f"PDF saved → {path}")
            if messagebox.askyesno("Exported", f"Saved!\n{path}\n\nOpen now?"):
                if sys.platform == 'win32':
                    os.startfile(path)
                else:
                    os.system(f'xdg-open "{path}"')
        except Exception as ex:
            messagebox.showerror("Export Error", str(ex))

    def _generate_pdf(self, path):
        doc  = SimpleDocTemplate(path, pagesize=A4,
                                  rightMargin=18*mm, leftMargin=18*mm,
                                  topMargin=12*mm, bottomMargin=12*mm)
        DARK = colors.HexColor('#0F172A')
        BLUE = colors.HexColor('#1D4ED8')
        MUT  = colors.HexColor('#6B7280')
        LIT  = colors.HexColor('#F1F5F9')
        STR  = colors.HexColor('#F8FAFC')
        W    = A4[0] - 36*mm
        cur  = self.currency.get()

        def ps(n='', **kw):
            return ParagraphStyle(n or 'x', parent=getSampleStyleSheet()['Normal'], **kw)

        story = []

        # Header
        logo_cell = ''
        if self.logo_path:
            try:
                li = RLImage(self.logo_path); li._restrictSize(55*mm, 32*mm)
                logo_cell = li
            except Exception:
                pass

        mp = [Paragraph(f'<font color="#6B7280">Invoice #</font>  <b>{self._meta["inv_num"].get()}</b>',
                         ps(fontSize=9, alignment=TA_RIGHT, spaceAfter=2)),
              Paragraph(f'<font color="#6B7280">Date</font>  <b>{self._meta["date"].get()}</b>',
                         ps(fontSize=9, alignment=TA_RIGHT, spaceAfter=2)),
              Paragraph(f'<font color="#6B7280">Due</font>  <b>{self._meta["due"].get()}</b>',
                         ps(fontSize=9, alignment=TA_RIGHT, spaceAfter=2))]
        if self._meta['po'].get():
            mp.append(Paragraph(f'<font color="#6B7280">PO #</font>  <b>{self._meta["po"].get()}</b>',
                                  ps(fontSize=9, alignment=TA_RIGHT)))

        hdr = Table([[logo_cell or Spacer(1,1), mp]], colWidths=[W*.5, W*.5])
        hdr.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                                  ('BOTTOMPADDING',(0,0),(-1,-1),8)]))
        story += [hdr, HRFlowable(width='100%', thickness=1.5, color=DARK, spaceAfter=8)]

        def addr_block(prefix, title):
            data  = self._get_addr_data(prefix)
            paras = [Paragraph(f'<font color="{BLUE.hexval()}"><b>{title}</b></font>',
                                 ps(fontSize=8, fontName='Helvetica-Bold', spaceAfter=4))]
            for f in self.ADDR_FIELDS:
                v = data.get(f, '').strip()
                if v:
                    paras.append(Paragraph(v, ps(fontSize=10, spaceAfter=1)))
            return paras

        addr = Table([[addr_block('from','FROM'), addr_block('to','BILL TO')]],
                      colWidths=[W*.5, W*.5])
        addr.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),
                                   ('BOTTOMPADDING',(0,0),(-1,-1),8),
                                   ('TOPPADDING',(0,0),(-1,-1),4)]))
        story += [addr, HRFlowable(width='100%', thickness=0.5, color=LIT, spaceAfter=10)]

        rows = [[Paragraph(f'<font color="white"><b>{h}</b></font>', ps(fontSize=9, alignment=a))
                 for h, a in zip(['#','Description','Qty','Unit Price',f'Amount ({cur})'],
                                 [TA_CENTER,TA_LEFT,TA_CENTER,TA_RIGHT,TA_RIGHT])]]
        subtotal = 0.0
        for i, rd in enumerate(self.item_rows):
            desc = rd['desc'].get_value() if isinstance(rd['desc'], AutocompleteEntry) \
                   else rd['desc'].get().strip()
            if not desc:
                continue
            try:
                qty = float(rd['qty'].get() or 0)
                rate = float(rd['rate'].get() or 0)
                amt  = qty * rate; subtotal += amt
            except ValueError:
                qty, rate, amt = 0, 0, 0
            rows.append([
                Paragraph(str(i+1), ps(fontSize=9, alignment=TA_CENTER)),
                Paragraph(desc, ps(fontSize=10)),
                Paragraph(f"{qty:g}", ps(fontSize=10, alignment=TA_CENTER)),
                Paragraph(f"{rate:,.2f}", ps(fontSize=10, alignment=TA_RIGHT)),
                Paragraph(f"{amt:,.2f}", ps(fontSize=10, alignment=TA_RIGHT, fontName='Helvetica-Bold')),
            ])

        it = Table(rows, colWidths=[W*.06, W*.42, W*.10, W*.20, W*.22], repeatRows=1)
        it.setStyle(TableStyle([
            ('BACKGROUND',    (0,0),(-1,0), DARK),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white, STR]),
            ('GRID',          (0,0),(-1,-1),0.3,colors.HexColor('#E5E7EB')),
            ('TOPPADDING',    (0,0),(-1,-1),6),
            ('BOTTOMPADDING', (0,0),(-1,-1),6),
            ('LEFTPADDING',   (1,0),(1,-1), 8),
        ]))
        story += [it, Spacer(1, 6*mm)]

        tax_pct = 0.0; tax_amt = 0.0
        if self.tax_enabled.get():
            try:
                tax_pct = float(self.tax_rate.get() or 0)
                tax_amt = subtotal * tax_pct / 100
            except ValueError:
                pass

        trows = [[Paragraph("Subtotal:", ps(fontSize=9, alignment=TA_RIGHT, textColor=MUT)),
                  Paragraph(f"{cur} {subtotal:,.2f}", ps(fontSize=10, alignment=TA_RIGHT))]]
        if tax_pct:
            trows.append([
                Paragraph(f"Tax ({tax_pct:.0f}%):", ps(fontSize=9, alignment=TA_RIGHT, textColor=MUT)),
                Paragraph(f"{cur} {tax_amt:,.2f}", ps(fontSize=10, alignment=TA_RIGHT))
            ])

        tt = Table(trows, colWidths=[W*.22, W*.20])
        tt.setStyle(TableStyle([('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2)]))
        rb = Table([[Spacer(1,1), tt]], colWidths=[W*.58, W*.42])
        rb.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP')]))
        story += [rb, Spacer(1, 3*mm)]

        gt = Table([[Paragraph('TOTAL', ps(fontSize=14, fontName='Helvetica-Bold', textColor=colors.white)),
                     Paragraph(f'{cur} {subtotal+tax_amt:,.2f}',
                                ps(fontSize=15, fontName='Helvetica-Bold',
                                   textColor=colors.white, alignment=TA_RIGHT))]],
                    colWidths=[W*.78, W*.22])
        gt.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),BLUE),
            ('TOPPADDING',(0,0),(-1,-1),10),
            ('BOTTOMPADDING',(0,0),(-1,-1),10),
            ('LEFTPADDING',(0,0),(0,-1),14),
            ('RIGHTPADDING',(1,0),(1,-1),14),
        ]))
        story += [gt, Spacer(1, 8*mm)]

        notes = self._notes.get('1.0','end').strip()
        if notes:
            story.append(HRFlowable(width='100%', thickness=0.5, color=LIT, spaceAfter=4))
            story.append(Paragraph('<b>Notes / Payment Information</b>',
                                    ps(fontSize=9, fontName='Helvetica-Bold', textColor=MUT, spaceAfter=4)))
            for line in notes.split('\n'):
                if line.strip():
                    story.append(Paragraph(line, ps(fontSize=9, textColor=MUT, spaceAfter=2)))

        story += [Spacer(1,12*mm),
                  HRFlowable(width='100%', thickness=0.5, color=LIT, spaceAfter=4),
                  Paragraph('Thank you for your business.',
                              ps(fontSize=9, textColor=MUT, alignment=TA_CENTER))]
        doc.build(story)

    # ── DB Manager ────────────────────────────────────────────────
    def _open_db_manager(self):
        win = tk.Toplevel(self.root)
        win.title("Database Manager")
        win.geometry("760x560")
        win.configure(bg=self.BG)
        win.grab_set()

        nb = ttk.Notebook(win)
        nb.pack(fill='both', expand=True, padx=12, pady=12)

        # Contacts tab
        cf = tk.Frame(nb, bg=self.BG)
        nb.add(cf, text="  👥 Contacts  ")

        ctb = tk.Frame(cf, bg=self.BG)
        ctb.pack(fill='x', pady=(0,6))
        tk.Label(ctb, text="Saved Contacts", bg=self.BG, fg=self.TEXT,
                 font=('Segoe UI', 11, 'bold')).pack(side='left')

        ccols = ('Name / Company','Address','Phone','Email','Updated')
        c_tree = ttk.Treeview(cf, columns=ccols, show='headings', height=14)
        for col, w in zip(ccols, [180,190,110,160,100]):
            c_tree.heading(col, text=col)
            c_tree.column(col, width=w, anchor='w')
        csb = ttk.Scrollbar(cf, orient='vertical', command=c_tree.yview)
        c_tree.configure(yscrollcommand=csb.set)
        csb.pack(side='right', fill='y')
        c_tree.pack(fill='both', expand=True)

        def refresh_contacts():
            for r in c_tree.get_children():
                c_tree.delete(r)
            for r in self.db.all_contacts():
                c_tree.insert('','end', values=(
                    r['name'] or '',
                    ((r['address1'] or '') + ' ' + (r['address2'] or '')).strip(),
                    r['phone'] or '', r['email'] or '',
                    (r['updated'] or '')[:10]))

        def del_contact():
            sel = c_tree.selection()
            if not sel:
                return
            label = c_tree.item(sel[0])['values'][0]
            if messagebox.askyesno("Delete", f'Delete "{label}"?', parent=win):
                self.db.delete_contact(label)
                refresh_contacts()

        def load_to_from(ev):
            sel = c_tree.selection()
            if sel:
                label = c_tree.item(sel[0])['values'][0]
                self._load_contact_by_name(label, 'from')
                win.destroy()

        c_tree.bind('<Double-1>', load_to_from)

        cb = tk.Frame(cf, bg=self.BG)
        cb.pack(fill='x', pady=4)
        self._btn(cb, "🗑 Delete Selected", del_contact, self.DANGER_BG, self.DANGER).pack(side='left')
        tk.Label(cb, text="Double-click to load into FROM card",
                 bg=self.BG, fg=self.MUTED, font=('Segoe UI', 8)).pack(side='left', padx=8)

        refresh_contacts()

        # Line Items tab
        lf = tk.Frame(nb, bg=self.BG)
        nb.add(lf, text="  📦 Line Items  ")

        ltb = tk.Frame(lf, bg=self.BG)
        ltb.pack(fill='x', pady=(0,6))
        tk.Label(ltb, text="Saved Line Items", bg=self.BG, fg=self.TEXT,
                 font=('Segoe UI', 11, 'bold')).pack(side='left')

        lcols = ('Description','Default Rate','Times Used','Last Used')
        l_tree = ttk.Treeview(lf, columns=lcols, show='headings', height=14)
        for col, w in zip(lcols, [300,120,100,110]):
            l_tree.heading(col, text=col)
            l_tree.column(col, width=w, anchor='w')
        lsb = ttk.Scrollbar(lf, orient='vertical', command=l_tree.yview)
        l_tree.configure(yscrollcommand=lsb.set)
        lsb.pack(side='right', fill='y')
        l_tree.pack(fill='both', expand=True)

        def refresh_items():
            for r in l_tree.get_children():
                l_tree.delete(r)
            for r in self.db.all_line_items():
                l_tree.insert('','end', values=(
                    r['description'],
                    f"{self.currency.get()} {r['rate']:.2f}",
                    r['used_count'],
                    (r['updated'] or '')[:10]))

        def del_item():
            sel = l_tree.selection()
            if not sel:
                return
            desc = l_tree.item(sel[0])['values'][0]
            if messagebox.askyesno("Delete", f'Delete "{desc}"?', parent=win):
                self.db.delete_line_item(desc)
                refresh_items()

        lb = tk.Frame(lf, bg=self.BG)
        lb.pack(fill='x', pady=4)
        self._btn(lb, "🗑 Delete Selected", del_item, self.DANGER_BG, self.DANGER).pack(side='left')
        refresh_items()

        self._btn(win, "✓  Close", win.destroy, self.ACCENT, 'white'
                  ).pack(pady=8, ipadx=20, ipady=4)

    # ── Save/Load JSON ────────────────────────────────────────────
    def _save_json(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".json", filetypes=[("JSON","*.json")],
            initialfile=f"Invoice_{self._meta['inv_num'].get()}.json")
        if not path:
            return
        items = []
        for rd in self.item_rows:
            desc = rd['desc'].get_value() if isinstance(rd['desc'], AutocompleteEntry) \
                   else rd['desc'].get()
            items.append({'desc': desc, 'qty': rd['qty'].get(), 'rate': rd['rate'].get()})
        data = {
            'meta':        {k: v.get() for k,v in self._meta.items()},
            'currency':    self.currency.get(),
            'from':        self._get_addr_data('from'),
            'to':          self._get_addr_data('to'),
            'items':       items,
            'tax_enabled': self.tax_enabled.get(),
            'tax_rate':    self.tax_rate.get(),
            'notes':       self._notes.get('1.0','end').strip(),
            'logo_path':   self.logo_path or '',
        }
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        self._set_status(f"Saved → {path}")

    def _load_json(self):
        path = filedialog.askopenfilename(filetypes=[("JSON","*.json"),("All","*.*")])
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception as ex:
            messagebox.showerror("Load Error", str(ex)); return

        for k, v in data.get('meta', {}).items():
            if k in self._meta:
                self._meta[k].set(v)
        self.currency.set(data.get('currency', 'RM'))
        self._set_addr_data('from', data.get('from', {}))
        self._set_addr_data('to',   data.get('to',   {}))

        for rd in self.item_rows:
            rd['frame'].destroy()
        self.item_rows.clear()

        for item in data.get('items', []):
            rd = self._add_row()
            rd['desc'].set_value(item.get('desc',''))
            rd['qty'].delete(0,'end'); rd['qty'].insert(0, item.get('qty','1'))
            rd['rate'].delete(0,'end'); rd['rate'].insert(0, item.get('rate','0.00'))
        if not self.item_rows:
            self._add_row()

        self.tax_enabled.set(data.get('tax_enabled', False))
        self.tax_rate.set(data.get('tax_rate','6'))
        self._notes.delete('1.0','end')
        self._notes.insert('1.0', data.get('notes',''))

        logo = data.get('logo_path','')
        if logo and os.path.exists(logo):
            self.logo_path = logo
            try:
                img = Image.open(logo).convert('RGBA')
                img.thumbnail((170,110), Image.LANCZOS)
                self.logo_photoimg = ImageTk.PhotoImage(img)
                self._logo_lbl.config(image=self.logo_photoimg, text='', bg=self.CARD)
            except Exception:
                pass

        self._recalc()
        self._set_status(f"Loaded → {path}")

    # ── Clear ─────────────────────────────────────────────────────
    def _clear_all(self):
        if not messagebox.askyesno("Clear All", "Reset all invoice data?"):
            return
        today = datetime.date.today()
        self._meta['inv_num'].set('INV-001')
        self._meta['date'].set(today.strftime('%d/%m/%Y'))
        self._meta['due'].set((today + datetime.timedelta(days=30)).strftime('%d/%m/%Y'))
        self._meta['po'].set('')
        blank = {f: '' for f in self.ADDR_FIELDS}
        self._set_addr_data('from', blank)
        self._set_addr_data('to',   blank)
        for rd in self.item_rows:
            rd['frame'].destroy()
        self.item_rows.clear()
        for _ in range(3):
            self._add_row()
        self.tax_enabled.set(False)
        self.tax_rate.set('6')
        self._notes.delete('1.0','end')
        self._notes.insert('1.0',"Bank: Maybank\nAccount No: 1234-5678-9012\nRef: Invoice #")
        self._remove_logo()
        self._recalc()
        self._set_status("Cleared.")

    # ── Helpers ───────────────────────────────────────────────────
    def _card(self, parent, **kw):
        return tk.Frame(parent, bg=self.CARD,
                        highlightbackground=self.BORDER, highlightthickness=1, **kw)

    def _btn(self, parent, text, cmd, bg, fg):
        return tk.Button(parent, text=text, command=cmd,
                         bg=bg, fg=fg, font=('Segoe UI', 9, 'bold'),
                         relief='flat', cursor='hand2',
                         activebackground=bg, activeforeground=fg)

    def _lbl_tag(self, parent, text, color, size, bold=False):
        fw = 'bold' if bold else 'normal'
        return tk.Label(parent, text=text, bg=self.CARD, fg=color,
                        font=('Segoe UI', size, fw))

    def _set_status(self, msg):
        self._statusvar.set(msg)


# ── Entry point ───────────────────────────────────────────────────
if __name__ == '__main__':
    root = tk.Tk()
    app  = InvoiceApp(root)
    root.mainloop()
