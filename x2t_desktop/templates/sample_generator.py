"""
Sample Excel Templates Generator
Creates pre-filled, compliant Excel files for testing Sales, Purchase, Bank, and Vertical sheets.
"""

import os
import pandas as pd


def generate_all_samples(output_dir: str):
    os.makedirs(output_dir, exist_ok=True)

    # 1. Sales Sample
    sales_data = [
        # Invoice 1 (Multi-line: 2 items)
        {
            "Voucher No": "INV-2024-001",
            "Date": "2024-04-01",
            "Party Name": "Alpha Infotech Ltd",
            "Stock Item": "Laptop Dell Core i7",
            "Qty": 2,
            "Rate": 60000.0,
            "Unit": "NOS",
            "Godown": "Main Location",
            "Sales Ledger": "Sales GST 18%",
            "Discount %": 0.0,
            "Taxable Amount": 120000.0,
            "GST Rate": 18,
            "CGST": 10800.0,
            "SGST": 10800.0,
            "IGST": 0.0,
            "Discount Amount": 0.0,
            "Discount Ledger": "",
            "Round Off": 0.0,
            "Total Amount": 141600.0,
            "Narration": "Being IT hardware sold under PO-4451"
        },
        {
            "Voucher No": "INV-2024-001",
            "Date": "2024-04-01",
            "Party Name": "Alpha Infotech Ltd",
            "Stock Item": "Wireless Keyboard Mouse Combo",
            "Qty": 5,
            "Rate": 1500.0,
            "Unit": "SET",
            "Godown": "Main Location",
            "Sales Ledger": "Sales GST 18%",
            "Discount %": 5.0,  # 5% item discount
            "Taxable Amount": 7125.0,
            "GST Rate": 18,
            "CGST": 641.25,
            "SGST": 641.25,
            "IGST": 0.0,
            "Discount Amount": 0.0,
            "Discount Ledger": "",
            "Round Off": 0.0,
            "Total Amount": 8407.50,
            "Narration": "Included accessories combo"
        },
        # Invoice 2 (Single item + Level 2 discount + Round off)
        {
            "Voucher No": "INV-2024-002",
            "Date": "2024-04-02",
            "Party Name": "Beta Enterprises",
            "Stock Item": "27-inch 4K IPS Monitor",
            "Qty": 1,
            "Rate": 27500.0,
            "Unit": "NOS",
            "Godown": "Main Location",
            "Sales Ledger": "Sales GST 18%",
            "Discount %": 0.0,
            "Taxable Amount": 27500.0,
            "GST Rate": 18,
            "CGST": 2475.0,
            "SGST": 2475.0,
            "IGST": 0.0,
            "Discount Amount": 500.0,  # Level 2 cash discount
            "Discount Ledger": "Cash Discount",
            "Round Off": 0.0,
            "Total Amount": 31950.0,
            "Narration": "Sold with special corporate discount"
        },
        # Invoice 3 (Inter-state sale with IGST)
        {
            "Voucher No": "INV-2024-003",
            "Date": "2024-04-03",
            "Party Name": "Global Logistics Corp",
            "Stock Item": "Network Router Dual Band",
            "Qty": 3,
            "Rate": 3200.33,
            "Unit": "NOS",
            "Godown": "Main Location",
            "Sales Ledger": "Interstate Sales 18%",
            "Discount %": 0.0,
            "Taxable Amount": 9600.99,
            "GST Rate": 18,
            "CGST": 0.0,
            "SGST": 0.0,
            "IGST": 1728.18,
            "Discount Amount": 0.0,
            "Discount Ledger": "",
            "Round Off": -0.17,  # Auto round-off
            "Total Amount": 11329.0,
            "Narration": "Inter-state supply to branch office"
        }
    ]
    df_sales = pd.DataFrame(sales_data)
    sales_file = os.path.join(output_dir, "sample_sales.xlsx")
    with pd.ExcelWriter(sales_file, engine="openpyxl") as writer:
        df_sales.to_excel(writer, sheet_name="Sales Invoices", index=False)

    # 2. Purchase Sample
    purch_data = [
        {
            "Voucher No": "PUR-001",
            "Supplier Inv No": "SUP-8890",
            "Date": "2024-04-05",
            "Supplier Inv Date": "2024-04-04",
            "Supplier Name": "Zenith Distributors",
            "Stock Item": "Laptop Dell Core i7",
            "Qty": 5,
            "Rate": 48000.0,
            "Unit": "NOS",
            "Godown": "Main Location",
            "Purchase Ledger": "Purchase Local 18%",
            "Discount %": 0.0,
            "Taxable Amount": 240000.0,
            "GST Rate": 18,
            "Input CGST": 21600.0,
            "Input SGST": 21600.0,
            "Input IGST": 0.0,
            "Discount Amount": 0.0,
            "Round Off": 0.0,
            "Total Amount": 283200.0,
            "Narration": "Inward supply of laptops"
        },
        {
            "Voucher No": "PUR-001",
            "Supplier Inv No": "SUP-8890",
            "Date": "2024-04-05",
            "Supplier Inv Date": "2024-04-04",
            "Supplier Name": "Zenith Distributors",
            "Stock Item": "Wireless Keyboard Mouse Combo",
            "Qty": 20,
            "Rate": 950.0,
            "Unit": "SET",
            "Godown": "Main Location",
            "Purchase Ledger": "Purchase Local 18%",
            "Discount %": 2.0,
            "Taxable Amount": 18620.0,
            "GST Rate": 18,
            "Input CGST": 1675.80,
            "Input SGST": 1675.80,
            "Input IGST": 0.0,
            "Discount Amount": 0.0,
            "Round Off": 0.0,
            "Total Amount": 21971.60,
            "Narration": "Bulk peripheral procurement"
        }
    ]
    df_purch = pd.DataFrame(purch_data)
    purch_file = os.path.join(output_dir, "sample_purchase.xlsx")
    with pd.ExcelWriter(purch_file, engine="openpyxl") as writer:
        df_purch.to_excel(writer, sheet_name="Purchase Bills", index=False)

    # 3. Bank Statement Sample (both 2-col and 1-col sheets)
    bank_2col = [
        {"Txn Date": "2024-04-01", "Particulars": "ATM CASH WDL 5542 INDIRANAGAR", "Chq/Ref No": "WDL001", "Withdrawal": 10000.0, "Deposit": 0.0, "Balance": 245000.0},
        {"Txn Date": "2024-04-02", "Particulars": "SMS CHG Q4 2024 HDFC", "Chq/Ref No": "", "Withdrawal": 17.70, "Deposit": 0.0, "Balance": 244982.30},
        {"Txn Date": "2024-04-03", "Particulars": "CREDIT INT REC UPTO MAR 24", "Chq/Ref No": "", "Withdrawal": 0.0, "Deposit": 1845.0, "Balance": 246827.30},
        {"Txn Date": "2024-04-04", "Particulars": "SALARY FOR STAFF APRIL", "Chq/Ref No": "NEFT123", "Withdrawal": 65000.0, "Deposit": 0.0, "Balance": 181827.30},
        {"Txn Date": "2024-04-05", "Particulars": "OFFICE RENT BESIDE METRO", "Chq/Ref No": "RTGS445", "Withdrawal": 25000.0, "Deposit": 0.0, "Balance": 156827.30},
        {"Txn Date": "2024-04-06", "Particulars": "ELECTRICITY BILL BESCOM AC 998", "Chq/Ref No": "UPI998", "Withdrawal": 4200.0, "Deposit": 0.0, "Balance": 152627.30},
        {"Txn Date": "2024-04-07", "Particulars": "CDM CASH DEPOSIT BR 102", "Chq/Ref No": "CDM102", "Withdrawal": 0.0, "Deposit": 30000.0, "Balance": 182627.30},
        {"Txn Date": "2024-04-08", "Particulars": "UPI/9988112233/CLIENT RECEIPT", "Chq/Ref No": "UPI551", "Withdrawal": 0.0, "Deposit": 12500.0, "Balance": 195127.30}
    ]
    bank_1col = [
        {"Date": "2024-04-01", "Narration": "ATM CASH WDL INDIRANAGAR", "Ref No": "WDL001", "Amount": 10000.0, "Dr/Cr": "DR"},
        {"Date": "2024-04-02", "Narration": "SMS CHG Q4 2024", "Ref No": "", "Amount": 17.70, "Dr/Cr": "DR"},
        {"Date": "2024-04-03", "Narration": "INTEREST REC Q4", "Ref No": "", "Amount": 1845.0, "Dr/Cr": "CR"},
        {"Date": "2024-04-04", "Narration": "SALARY APRIL", "Ref No": "NEFT123", "Amount": 65000.0, "Dr/Cr": "DR"},
        {"Date": "2024-04-05", "Narration": "UPI CLIENT PAYMENT", "Ref No": "UPI551", "Amount": 12500.0, "Dr/Cr": "CR"}
    ]
    bank_file = os.path.join(output_dir, "sample_bank_statement.xlsx")
    with pd.ExcelWriter(bank_file, engine="openpyxl") as writer:
        pd.DataFrame(bank_2col).to_excel(writer, sheet_name="Bank 2-Column", index=False)
        pd.DataFrame(bank_1col).to_excel(writer, sheet_name="Bank 1-Column", index=False)

    # 4. Vertical Journal Sample
    vertical_data = [
        # Voucher 1: Rent Payment (Bank Payment)
        {"Voucher No": "VCH-001", "Voucher Type": "Payment", "Date": "2024-04-01", "Ledger Name": "Rent Expenses", "Dr / Cr": "Dr", "Amount": 30000.0, "Narration": "April office rent", "Cost Center": "Head Office"},
        {"Voucher No": "VCH-001", "Voucher Type": "Payment", "Date": "2024-04-01", "Ledger Name": "HDFC Bank Account", "Dr / Cr": "Cr", "Amount": 30000.0, "Narration": "April office rent", "Cost Center": ""},
        
        # Voucher 2: Depreciation Journal Entry
        {"Voucher No": "VCH-002", "Voucher Type": "Journal", "Date": "2024-04-02", "Ledger Name": "Depreciation on Furniture", "Dr / Cr": "Dr", "Amount": 4500.0, "Narration": "Monthly furniture depreciation", "Cost Center": ""},
        {"Voucher No": "VCH-002", "Voucher Type": "Journal", "Date": "2024-04-02", "Ledger Name": "Accumulated Depreciation", "Dr / Cr": "Cr", "Amount": 4500.0, "Narration": "Monthly furniture depreciation", "Cost Center": ""},

        # Voucher 3: Multi-line Salary Distribution
        {"Voucher No": "VCH-003", "Voucher Type": "Payment", "Date": "2024-04-03", "Ledger Name": "Salary - Engineering", "Dr / Cr": "Dr", "Amount": 75000.0, "Narration": "April Staff Payroll", "Cost Center": "Tech Dept"},
        {"Voucher No": "VCH-003", "Voucher Type": "Payment", "Date": "2024-04-03", "Ledger Name": "Salary - Sales", "Dr / Cr": "Dr", "Amount": 45000.0, "Narration": "April Staff Payroll", "Cost Center": "Sales Dept"},
        {"Voucher No": "VCH-003", "Voucher Type": "Payment", "Date": "2024-04-03", "Ledger Name": "Professional Tax Payable", "Dr / Cr": "Cr", "Amount": 400.0, "Narration": "PT deduction", "Cost Center": ""},
        {"Voucher No": "VCH-003", "Voucher Type": "Payment", "Date": "2024-04-03", "Ledger Name": "HDFC Bank Account", "Dr / Cr": "Cr", "Amount": 119600.0, "Narration": "Net salary paid via bank", "Cost Center": ""},

        # Voucher 4: Contra (Cash withdrawal)
        {"Voucher No": "VCH-004", "Voucher Type": "Contra", "Date": "2024-04-04", "Ledger Name": "Cash", "Dr / Cr": "Dr", "Amount": 15000.0, "Narration": "Petty cash withdrawal", "Cost Center": ""},
        {"Voucher No": "VCH-004", "Voucher Type": "Contra", "Date": "2024-04-04", "Ledger Name": "HDFC Bank Account", "Dr / Cr": "Cr", "Amount": 15000.0, "Narration": "Petty cash withdrawal", "Cost Center": ""}
    ]
    df_vertical = pd.DataFrame(vertical_data)
    vertical_file = os.path.join(output_dir, "sample_vertical_journal.xlsx")
    with pd.ExcelWriter(vertical_file, engine="openpyxl") as writer:
        df_vertical.to_excel(writer, sheet_name="Journal Entries", index=False)

    return {
        "sales": sales_file,
        "purchase": purch_file,
        "bank": bank_file,
        "vertical": vertical_file
    }


if __name__ == "__main__":
    out = os.path.join(os.path.dirname(os.path.dirname(__file__)), "..", "sample_data")
    created = generate_all_samples(out)
    for k, v in created.items():
        print(f"Generated {k}: {v}")
