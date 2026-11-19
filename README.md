# X2T - Excel to Tally Converter (v2.0)

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://python.org)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D6.svg)](https://microsoft.com/windows)
[![UI Framework](https://img.shields.io/badge/UI-CustomTkinter-blueviolet.svg)](https://github.com/TomSchimansky/CustomTkinter)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

A high-performance, standalone Windows Desktop Application built in Python with a modern CustomTkinter interface. **X2T** is designed as a drop-in, enterprise-grade replacement for legacy Excel VBA Macro tools (X2T Beta). It parses complex multi-line Excel/CSV spreadsheets and generates standard, schema-compliant Tally XML or pushes data directly into **Tally Prime** and **Tally ERP 9** via local HTTP XML synchronization (port 9000).

---

## 🚀 Key Features & Modules

### 1. 🛒 Sales & Debit Note Module
- **Multi-Line Item Support**: Groups multiple stock item lines under the same invoice by Voucher Number (`INV-001`, `INV-002`, etc.).
- **Inventory Mappings**: Stock Item Name, Quantity, Unit of Measure (UOM), Rate, Godown / Location.
- **GST Bifurcations**: Automatic or column-mapped CGST, SGST, and IGST ledgers (e.g. `CGST 9%`, `SGST 9%`, `Output CGST`, `Output IGST`).
- **Level 1 Discount**: Post-discount unit rate or item-level discount percentage (`Discount %`).
- **Level 2 Discount**: Invoice-level cash discount or trade discount ledger line item (e.g. `Cash Discount`).
- **Auto Round-off**: Automatically calculates difference between gross sum and rounded invoice total, inserting a balancing `Round Off` ledger entry.
- **Pure Accounting Mode**: Also supports service and non-inventory invoices.

### 2. 📦 Purchase & Credit Note Module
- Maps Supplier Invoice Number (`REFERENCE`) and Supplier Invoice Date (`REFERENCEDATE`).
- Handles Supplier Party Ledger (`Sundry Creditors`), Purchase nominal ledgers, and Input GST (`Input CGST`, `Input SGST`, `Input IGST`).
- Credit Note toggle for purchase returns.

### 3. 🏦 Bank Statement Import Module
- **Dual Format Support**:
  - **2-Column Format**: Separate `Withdrawal (Dr)` and `Deposit (Cr)` columns.
  - **1-Column Format**: Single `Amount` column with `Dr/Cr` or `Type` indicator.
- **Auto-Categorization Rules Engine**:
  - Automatically matches narration keywords to target ledgers and voucher types:
    - `"ATM"` / `"CASH WDL"` $\rightarrow$ **Cash** (`Contra`)
    - `"CHG"` / `"CHARGES"` / `"SMS CHG"` $\rightarrow$ **Bank Charges** (`Payment`)
    - `"INT"` / `"INTEREST"` $\rightarrow$ **Bank Interest Received** (`Receipt`)
    - `"SALARY"` $\rightarrow$ **Salary Expenses** (`Payment`)
    - `"RENT"` $\rightarrow$ **Rent Expenses** (`Payment`)
    - `"BESCOM"` / `"POWER"` $\rightarrow$ **Electricity Expenses** (`Payment`)
    - Unknown narrations $\rightarrow$ **Suspense A/c** (flagged for review).
- **Interactive Reassignment**: Double-click any row to edit or override the assigned ledger before export.
- **Rules Manager**: Add, edit, delete, or reset categorization rules saved persistently to `templates/rules.json`.

### 4. 📑 Vertical / Generic Journal Import Format
- Supports multi-row debit and credit entries (`Payment`, `Receipt`, `Contra`, `Journal`).
- Supports Cost Center allocations (`Primary Cost Category`).
- **Strict Dr == Cr Validation**: Validates that every voucher satisfies $\sum \text{Debit} == \sum \text{Credit}$ before export, flagging imbalanced vouchers with exact difference amounts.

### 5. 🗂️ Master Generators (Pre-requisites for Tally)
- **"Fetch Ledgers"**: Scans all transaction sheets across tabs or external files, extracts unique ledgers, auto-classifies their parent groups (`Sundry Debtors`, `Sundry Creditors`, `Sales Accounts`, `Purchase Accounts`, `Duties & Taxes`, `Bank Accounts`), and generates Tally Ledger Master XML.
- **"Fetch Stock Items"**: Extracts unique stock items, parent groups, and base units (`PCS`, `NOS`, `SET`), generating Unit Masters and Stock Item Masters XML.
- Push masters directly to Tally Prime before importing vouchers.

### 6. ⚡ Output Actions & Direct Sync
- **Generate XML**: Exports Tally standard envelope XML (`<ENVELOPE><HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER><BODY><IMPORTDATA>...`).
- **Interactive XML Preview**: View syntax, copy to clipboard, or save to disk.
- **Push to Tally (Direct Sync)**: Sends XML directly to Tally's HTTP port (`http://localhost:9000`), parsing response tags (`<CREATED>`, `<ALTERED>`, `<ERRORS>`, `<LINEERROR>`) with detailed diagnostic cards.
- **Sample Template Exporter**: One-click download of sample Excel workbooks pre-formatted for Sales, Purchase, Bank Statements, and Vertical Journals.

---

## 📐 Tally XML Accounting Sign Conventions

Tally internal XML requires strict adherence to sign and direction conventions:
- **Debit Entry (Dr)**: `<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>` and `<AMOUNT>-amount</AMOUNT>` (Negative value).
- **Credit Entry (Cr)**: `<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>` and `<AMOUNT>+amount</AMOUNT>` (Positive value).
- **Mathematical Identity**: In every single voucher, the sum of all XML `<AMOUNT>` tags must equal `0.00`:
  $$\sum \text{Debit Amounts } (-X) + \sum \text{Credit Amounts } (+X) = 0$$

Dates are strictly formatted as `YYYYMMDD` (e.g., `20240401`).

---

## 📂 Project Directory Structure

```
EXCEL TO TALLY EXPORT/
├── dist/
│   └── X2T_Converter.exe        # Standalone Windows executable
├── x2t_desktop/
│   ├── app.py                   # Main CustomTkinter GUI application
│   ├── build.spec               # PyInstaller build specification
│   ├── requirements.txt         # Dependencies
│   ├── core/
│   │   ├── excel_parser.py      # Pandas & OpenPyXL parser & validator
│   │   ├── tally_xml_builder.py # Tally XML generator (Vouchers & Masters)
│   │   ├── tally_connector.py   # Direct HTTP connector to Tally port 9000
│   │   ├── rules_engine.py      # Bank narration keyword rules engine
│   │   └── resource_path.py     # Bundled asset path resolver
│   ├── ui/
│   │   ├── common.py            # ModernTable, XMLPreviewModal, SyncResponseModal
│   │   ├── tabs_sales.py        # Sales & Debit Note tab
│   │   ├── tabs_purchase.py     # Purchase & Credit Note tab
│   │   ├── tabs_bank.py         # Bank Statement tab
│   │   ├── tabs_vertical.py     # Vertical / Journal tab
│   │   ├── tabs_masters.py      # Ledger & Stock Master generator tab
│   │   └── tabs_settings.py     # Settings & Bank Rules Manager tab
│   └── templates/
│       ├── rules.json           # Categorization keyword rules
│       └── sample_generator.py  # Sample Excel generator script
├── sample_data/                 # Pre-generated sample Excel workbooks
│   ├── sample_sales.xlsx
│   ├── sample_purchase.xlsx
│   ├── sample_bank_statement.xlsx
│   └── sample_vertical_journal.xlsx
├── run.py                       # Root launcher script
├── build_exe.bat                # 1-Click PyInstaller compilation script
└── README.md                    # Documentation
```

---

## 🛠️ Installation & Running from Source

### Prerequisites
- Python 3.10+ (tested on Python 3.10, 3.11, 3.12, 3.13, 3.14)
- Windows 10 or Windows 11

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Run Application
```bash
python run.py
```
Or run directly:
```bash
python x2t_desktop/app.py
```

---

## 📦 Building Single Standalone .EXE

To compile the entire application into a single standalone `.exe` without console popup:

### Option A: Using the Batch Script
Double-click `build_exe.bat` or run:
```cmd
build_exe.bat
```

### Option B: Exact PyInstaller Command
```cmd
python -m PyInstaller --noconfirm --onefile --windowed ^
    --add-data "x2t_desktop\templates;templates" ^
    --add-data "sample_data;sample_data" ^
    --collect-all customtkinter ^
    --name "X2T_Converter" ^
    run.py
```

The resulting standalone executable will be generated at:
`dist\X2T_Converter.exe`

---

## 🔌 Configuring Tally Prime / ERP 9 for Direct Sync

To enable direct XML push to Tally:
1. Open **Tally Prime** (or Tally ERP 9).
2. Go to **F1: Help** > **Settings** > **Connectivity**.
3. Under **Client/Server configuration**:
   - Set **TallyPrime acts as**: `Both` (or `Server`).
   - Set **Enable ODBC**: `Yes`.
   - Set **Port**: `9000`.
4. Restart Tally Prime.
5. In **X2T**, click **🔌 Test Link** in the top-right header or in Settings tab to verify connection and display the active company name!
