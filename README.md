# Invoice Generator — Portable Desktop App

A clean, portable invoice generator built with Python + Tkinter.  
No installer needed. Just run and go.

## Quick Start

**Windows:**
```
Double-click run.bat
```

**Any OS (manual):**
```bash
pip install Pillow reportlab
python invoice_app.py
```

> Dependencies (Pillow + reportlab) are also auto-installed on first launch.

---

## Features

| Feature | Details |
|---|---|
| 📋 Invoice Header | Invoice #, Date, Due Date, PO Number |
| 🏢 From / To | Your details + client billing address |
| 🖼 Logo | Upload PNG/JPG — appears on the PDF |
| 📝 Line Items | Unlimited rows — Description, Qty, Unit Price, auto Amount |
| 🧮 Auto Totals | Live subtotal, tax, and grand total |
| 💰 Tax Toggle | Enable/disable tax with custom % |
| 💱 Currency | RM, USD, SGD, EUR, GBP, JPY |
| 📄 Export PDF | Professional A4 invoice PDF |
| 💾 Save/Load | Save as JSON, reload later |
| 🗑 Clear | Reset for a fresh invoice |

---

## File Structure

```
invoice_app/
├── invoice_app.py   ← Main application (single file)
├── run.bat          ← Windows launcher (auto-installs deps)
├── requirements.txt ← pip dependencies
└── README.md        ← This file
```

---

## Tips

- **Placeholder text** in address fields clears when you click them.
- **Tax** — tick the checkbox and enter any %, e.g. `6` for SST 6%.
- **Logo** — PNG with transparent background looks best on PDF.
- **JSON save** — stores everything including logo path for reuse.
- **PDF** — opens automatically after export (Windows).

---

Made with Python 3 · tkinter · Pillow · ReportLab
