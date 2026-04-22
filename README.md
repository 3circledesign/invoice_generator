# 🧾 Invoice Generator

A portable, offline-first desktop invoice generator built with **Python + Tkinter**. No installation needed — just run and go.

---

## ✨ Features

- **Company logo** — upload PNG/JPG, appears on the exported PDF
- **From / Bill To** — full address cards for your details and your client
- **Contact database** — saved contacts auto-complete as you type, or browse with **▼ Select**
- **Line items** — unlimited rows with description, qty, unit price, and auto-calculated amount
- **Smart description autocomplete** — saved items suggest themselves and auto-fill the last-used rate
- **Live totals** — subtotal, optional tax (any %), and grand total update in real time
- **Multi-currency** — RM, USD, SGD, EUR, GBP, JPY
- **Export to PDF** — clean A4 invoice, opens automatically after export
- **Save / Load JSON** — save your invoice state and reload it later
- **Database manager** — browse, search, and delete saved contacts and line items
- **Portable** — single `.py` file, runs anywhere Python 3.9+ is installed

---

## 🚀 Quick Start

### Windows
```
Double-click run.bat
```
Dependencies are installed automatically on first launch.

### macOS / Linux
```bash
pip install Pillow reportlab
python invoice_app.py
```

---

## 📋 Requirements

- Python 3.9+
- Pillow
- ReportLab

> Both are auto-installed by `run.bat` on first run.

---

## 📁 File Structure

```
invoice_app/
├── invoice_app.py       # Main application — single file
├── invoice_data.db      # SQLite database (auto-created on first run)
├── run.bat              # Windows launcher
├── requirements.txt     # pip dependencies
└── README.md
```

---

## 🗄 Database

The app stores data locally in `invoice_data.db` (SQLite) — no server, no internet required.

| Table | Stores |
|---|---|
| `contacts` | Name, address, phone, email — for FROM and BILL TO cards |
| `line_items` | Description + last-used rate, sorted by frequency |

**Contacts** are saved automatically every time you export a PDF, or manually via the **💾 Save Contact** button on each card.

**Line items** are saved automatically on PDF export and suggested via autocomplete when typing in the description field.

---

## 🖨 PDF Export

The exported PDF includes:

- Your logo (top left)
- Invoice #, Date, Due Date (top right)
- FROM and BILL TO address blocks
- Line items table with qty, unit price, and amount
- Subtotal, tax (optional, any %), and grand total
- Notes / payment info section

---

## 💾 Save & Load

Invoices can be saved as `.json` files and reloaded later — preserving all fields, line items, tax settings, notes, and logo path.

---

## 📸 Screenshots

> _Add your screenshots here_

---

## 📄 License

MIT — free to use, modify, and distribute.
