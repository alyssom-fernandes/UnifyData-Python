# UnifyData

> **Intelligent payroll reconciliation tool powered by Python and Streamlit.**

UnifyData automates the monthly cross-referencing of employee benefit consumption (extracted from Linx AutoSystem) against HR payroll records. Unlike a simple lookup, it uses **fuzzy name matching** to detect and surface typos, accent differences, and name inconsistencies between systems — making it ideal for environments where records don't share a common enrollment code.

---

## ✨ Features

- **Fuzzy Name Matching** — uses `thefuzz` (`token_sort_ratio`) to identify similar names across systems, catching typos and formatting inconsistencies automatically
- **Dual-Value Capture** — extracts both the *overdue* and *total debt* amounts from each Linx report entry, displaying them as separate columns
- **Interactive Divergence Review** — ambiguous or unmatched names surface in a dedicated review panel with options to approve, ignore, or manually link by CPF
- **Multi-Company Support** — handles multiple HR spreadsheets at once, generating one Excel sheet per company unit
- **Local Web Interface** — runs as a Streamlit app in your browser with a custom dark-mode design
- **Excel Export** — generates a `.xlsx` report with color-coded headers, currency formatting, gray rows for zero-consumption employees, and auto-width columns

---

## 🗂️ Project Structure

```
unifydata-python/
├── unifydata.py             # Main Streamlit application
├── requirements.txt         # Python dependencies
├── Iniciar_UnifyData.bat    # One-click Windows launcher
└── README.md
```

---

## 🚀 How to Run

### Prerequisites
- Python 3.8+ installed on your machine

### Option A — Windows (one click)
Double-click **`Iniciar_UnifyData.bat`**.

It will automatically install dependencies on first run and open the app in your browser.

### Option B — Manual (any OS)

```bash
# Install dependencies
pip install -r requirements.txt

# Launch the app
streamlit run unifydata.py
```

The app will open at `http://localhost:8501`.

---

## 📖 How to Use

1. **Upload HR Files** — drag your company HR spreadsheets (`.xlsx`, `.xls`, or `.csv`). Each file name (e.g. `Empregados Posto Rosário.xls`) is used as the company unit name.
2. **Upload Linx Reports** — drag the monthly consumption CSV files exported from Linx AutoSystem.
3. **Click "Process and Unify Data"** — the app cross-references both sources and detects divergences.
4. **Review Divergences** — for each flagged entry, choose to:
   - ✅ **Approve** a suggested name match
   - 🔗 **Link by CPF** for manual correction
   - 🚫 **Ignore** to exclude the entry from the report
5. **Download** the consolidated `.xlsx` report.

---

## 📊 Output Format

Each company unit gets its own sheet. Columns per sheet:

| Funcionário | CPF | Overdue [Origin] | Total [Origin] | ... | Total Discount |
|---|---|---|---|---|---|
| JOHN DOE | 000.000.000-00 | R$ 0,00 | R$ 150,00 | ... | R$ 150,00 |

- Header row: dark burgundy background, white bold text
- Employees with zero consumption: displayed in gray
- Footer row: grand total in bold

---

## 🛠️ Tech Stack

| Technology | Role |
|---|---|
| [Python 3](https://python.org) | Core language |
| [Streamlit](https://streamlit.io) | Web interface framework |
| [pandas](https://pandas.pydata.org) | Data processing and pivoting |
| [thefuzz](https://github.com/seatgeek/thefuzz) | Fuzzy name matching |
| [openpyxl](https://openpyxl.readthedocs.io) | Excel report generation |

---

## 📋 Dependencies

```
streamlit
pandas
openpyxl
thefuzz[speedup]
python-Levenshtein
```

Install with:
```bash
pip install -r requirements.txt
```

---

## 🔄 Difference from the Web Version

| | UnifyData Web | UnifyData Python |
|---|---|---|
| **Installation** | None (open HTML) | Python required |
| **Matching method** | Enrollment code (exact) | Fuzzy name matching |
| **Persistence** | `localStorage` (permanent) | Session only (re-upload each month) |
| **Best for** | Fast monthly closing | Systems without common enrollment codes |

---

## 🤝 Contributing

Feel free to open issues or pull requests. Contributions that improve CSV format compatibility or add new matching strategies are especially welcome.

---

## 📄 License

MIT License — free to use, modify, and distribute.

---

*Developed by AFN Systems · 2026*
