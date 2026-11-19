"""
Excel & CSV Parser Engine
Parses, cleans, validates, and groups transaction and master data from spreadsheets.
"""

import os
import re
import pandas as pd
import numpy as np
from typing import List, Dict, Any, Tuple, Optional
from datetime import datetime


def find_column(df_columns: List[str], candidates: List[str]) -> Optional[str]:
    """Finds a column in df_columns matching any candidate (case-insensitive & stripped)."""
    clean_cols = {re.sub(r"[_\s\-\.]+", "", str(col)).lower(): col for col in df_columns}
    for cand in candidates:
        cand_clean = re.sub(r"[_\s\-\.]+", "", cand).lower()
        if cand_clean in clean_cols:
            return clean_cols[cand_clean]
    # Secondary check: substring match
    for cand in candidates:
        cand_clean = re.sub(r"[_\s\-\.]+", "", cand).lower()
        for c_clean, orig_col in clean_cols.items():
            if cand_clean in c_clean or c_clean in cand_clean:
                return orig_col
    return None


def clean_num(val: Any, default: float = 0.0) -> float:
    """Cleans a numerical value from strings, floats, ints, currency symbols."""
    if pd.isna(val) or val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val) if not np.isnan(val) else default
    s = str(val).replace(",", "").replace("₹", "").replace("$", "").replace(" ", "").strip()
    try:
        return float(s)
    except ValueError:
        return default


def clean_str(val: Any, default: str = "") -> str:
    """Cleans a string value."""
    if pd.isna(val) or val is None:
        return default
    s = str(val).strip()
    return "" if s.lower() == "nan" else s


class ExcelParser:
    """Parser for Sales, Purchase, Bank Statements, and Vertical Journal sheets."""

    @staticmethod
    def get_sheet_names(file_path: str) -> List[str]:
        """Returns sheet names of an Excel workbook."""
        if not os.path.exists(file_path):
            return []
        try:
            if file_path.lower().endswith(".csv"):
                return ["CSV Data"]
            xl = pd.ExcelFile(file_path, engine="openpyxl")
            return xl.sheet_names
        except Exception:
            return []

    @staticmethod
    def load_sheet_preview(file_path: str, sheet_name: Optional[str] = None, nrows: int = 100) -> pd.DataFrame:
        """Loads first N rows of a sheet as DataFrame for preview."""
        if not os.path.exists(file_path):
            return pd.DataFrame()
        try:
            if file_path.lower().endswith(".csv"):
                df = pd.read_csv(file_path, nrows=nrows)
            else:
                df = pd.read_excel(file_path, sheet_name=sheet_name or 0, nrows=nrows, engine="openpyxl")
            return df
        except Exception:
            return pd.DataFrame()

    @staticmethod
    def load_full_sheet(file_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
        """Loads entire sheet."""
        if file_path.lower().endswith(".csv"):
            return pd.read_csv(file_path)
        return pd.read_excel(file_path, sheet_name=sheet_name or 0, engine="openpyxl")

    # =========================================================================
    # SALES / DEBIT NOTE PARSER
    # =========================================================================
    @classmethod
    def parse_sales_sheet(cls, df: pd.DataFrame, col_map: Optional[Dict[str, str]] = None) -> Tuple[List[Dict[str, Any]], List[str]]:
        """
        Parses Sales or Debit Note transactions from DataFrame.
        Groups multi-line stock items under the same invoice by Voucher Number.
        Returns: (invoices_list, validation_errors_list)
        """
        cols = list(df.columns)
        errors = []

        # Auto-detect or use user mappings
        c_vch_no = col_map.get("vch_no") if col_map else find_column(cols, ["Voucher No", "Vch No", "Invoice No", "Inv No", "Bill No"])
        c_date = col_map.get("date") if col_map else find_column(cols, ["Date", "Voucher Date", "Inv Date", "Invoice Date"])
        c_party = col_map.get("party") if col_map else find_column(cols, ["Party Name", "Party Ledger", "Customer Name", "Buyer Name", "Party"])
        c_item = col_map.get("item") if col_map else find_column(cols, ["Stock Item", "Item Name", "Item Description", "Product", "Item"])
        c_qty = col_map.get("qty") if col_map else find_column(cols, ["Qty", "Quantity", "Billed Qty", "Actual Qty"])
        c_rate = col_map.get("rate") if col_map else find_column(cols, ["Rate", "Price", "Unit Rate", "Unit Price"])
        c_unit = col_map.get("unit") if col_map else find_column(cols, ["Unit", "UOM", "Unit of Measure"])
        c_godown = col_map.get("godown") if col_map else find_column(cols, ["Godown", "Location", "Warehouse"])
        c_sales_ledger = col_map.get("sales_ledger") if col_map else find_column(cols, ["Sales Ledger", "Income Ledger", "Sales Account"])
        c_taxable = col_map.get("taxable") if col_map else find_column(cols, ["Taxable Amount", "Taxable Value", "Taxable", "Basic Amount"])
        c_gst_pct = col_map.get("gst_pct") if col_map else find_column(cols, ["GST %", "GST Rate", "Tax Rate", "Rate %"])
        c_cgst = col_map.get("cgst") if col_map else find_column(cols, ["CGST", "CGST Amount", "Central Tax"])
        c_sgst = col_map.get("sgst") if col_map else find_column(cols, ["SGST", "SGST Amount", "State Tax"])
        c_igst = col_map.get("igst") if col_map else find_column(cols, ["IGST", "IGST Amount", "Integrated Tax"])
        c_disc1 = col_map.get("disc1") if col_map else find_column(cols, ["Discount %", "Item Disc %", "Level 1 Discount", "Disc %"])
        c_disc2 = col_map.get("disc2") if col_map else find_column(cols, ["Discount Amount", "Cash Discount", "Level 2 Discount", "Trade Discount"])
        c_disc2_ledger = col_map.get("disc2_ledger") if col_map else find_column(cols, ["Discount Ledger", "Disc Ledger"])
        c_round_off = col_map.get("round_off") if col_map else find_column(cols, ["Round Off", "Roundoff"])
        c_total = col_map.get("total") if col_map else find_column(cols, ["Total Amount", "Invoice Amount", "Net Amount", "Grand Total", "Invoice Value"])
        c_narration = col_map.get("narration") if col_map else find_column(cols, ["Narration", "Remarks", "Description"])
        c_reference = col_map.get("reference") if col_map else find_column(cols, ["Reference", "PO No", "Ref No", "Order No"])

        if not c_vch_no and not c_date:
            errors.append("Critical Error: Could not detect 'Voucher Number' or 'Date' columns.")
            return [], errors

        invoices_dict: Dict[str, Dict[str, Any]] = {}
        row_count = 0

        for idx, row in df.iterrows():
            row_idx = idx + 2  # 1-based index including header
            vch_no = clean_str(row.get(c_vch_no)) if c_vch_no else ""
            date_val = row.get(c_date) if c_date else None
            
            # Skip empty rows
            if not vch_no and pd.isna(date_val):
                continue

            # If vch_no is empty, fallback to Date + Row
            if not vch_no:
                vch_no = f"INV-{row_idx}"

            party = clean_str(row.get(c_party)) if c_party else "Cash"
            if not party:
                party = "Cash"

            # Create or update invoice record
            if vch_no not in invoices_dict:
                invoices_dict[vch_no] = {
                    "vch_no": vch_no,
                    "date": date_val,
                    "party_name": party,
                    "reference": clean_str(row.get(c_reference)) if c_reference else vch_no,
                    "narration": clean_str(row.get(c_narration)) if c_narration else "",
                    "items": [],
                    "taxes_dict": {},
                    "discount_amount": 0.0,
                    "discount_ledger": clean_str(row.get(c_disc2_ledger)) if c_disc2_ledger else "Discount Allowed",
                    "manual_round_off": None,
                    "expected_total": None,
                    "rows_indexed": []
                }

            inv = invoices_dict[vch_no]
            inv["rows_indexed"].append(row_idx)

            # Item details
            item_name = clean_str(row.get(c_item)) if c_item else ""
            qty = clean_num(row.get(c_qty), 0.0) if c_qty else 0.0
            rate = clean_num(row.get(c_rate), 0.0) if c_rate else 0.0
            unit = clean_str(row.get(c_unit), "PCS") if c_unit else "PCS"
            godown = clean_str(row.get(c_godown), "Main Location") if c_godown else "Main Location"
            sales_ledger = clean_str(row.get(c_sales_ledger), "Sales Account") if c_sales_ledger else "Sales Account"
            disc1_pct = clean_num(row.get(c_disc1), 0.0) if c_disc1 else 0.0

            # Level 1 Discount calculation
            gross_amount = qty * rate if (qty and rate) else clean_num(row.get(c_taxable), 0.0)
            item_taxable = gross_amount
            if disc1_pct > 0:
                item_taxable = gross_amount * (1.0 - (disc1_pct / 100.0))
            elif c_taxable and clean_num(row.get(c_taxable), 0.0) > 0:
                item_taxable = clean_num(row.get(c_taxable), 0.0)

            if item_name:
                inv["items"].append({
                    "stock_item": item_name,
                    "qty": qty if qty else 1.0,
                    "unit": unit or "PCS",
                    "rate": rate if rate else item_taxable,
                    "discount_pct": disc1_pct,
                    "amount": round(item_taxable, 2),
                    "godown": godown or "Main Location",
                    "sales_ledger": sales_ledger or "Sales Account"
                })

            # Taxes
            cgst_val = clean_num(row.get(c_cgst), 0.0) if c_cgst else 0.0
            sgst_val = clean_num(row.get(c_sgst), 0.0) if c_sgst else 0.0
            igst_val = clean_num(row.get(c_igst), 0.0) if c_igst else 0.0

            # If GST % given but separate tax amounts missing
            gst_pct = clean_num(row.get(c_gst_pct), 0.0) if c_gst_pct else 0.0
            if (cgst_val == 0 and sgst_val == 0 and igst_val == 0) and gst_pct > 0:
                # Default to intra-state (CGST + SGST)
                half_rate = gst_pct / 2.0
                cgst_val = round(item_taxable * (half_rate / 100.0), 2)
                sgst_val = round(item_taxable * (half_rate / 100.0), 2)
                cgst_key = f"CGST {half_rate:g}%"
                sgst_key = f"SGST {half_rate:g}%"
            else:
                cgst_key = "CGST"
                sgst_key = "SGST"

            if cgst_val > 0:
                inv["taxes_dict"][cgst_key] = inv["taxes_dict"].get(cgst_key, 0.0) + cgst_val
            if sgst_val > 0:
                inv["taxes_dict"][sgst_key] = inv["taxes_dict"].get(sgst_key, 0.0) + sgst_val
            if igst_val > 0:
                inv["taxes_dict"]["IGST"] = inv["taxes_dict"].get("IGST", 0.0) + igst_val

            # Level 2 Discount
            l2_disc = clean_num(row.get(c_disc2), 0.0) if c_disc2 else 0.0
            if l2_disc > 0:
                inv["discount_amount"] += l2_disc

            # Round Off & Total
            if c_round_off and not pd.isna(row.get(c_round_off)):
                inv["manual_round_off"] = clean_num(row.get(c_round_off), 0.0)
            if c_total and not pd.isna(row.get(c_total)):
                inv["expected_total"] = clean_num(row.get(c_total), 0.0)

            row_count += 1

        # Post-process invoices: consolidate totals, round-offs, taxes
        result_invoices = []
        for vch_no, inv in invoices_dict.items():
            # If no items were added (e.g. pure accounting sales), check if taxable exists
            items_total = sum(itm["amount"] for itm in inv["items"])
            tax_total = sum(inv["taxes_dict"].values())
            subtotal = items_total + tax_total - inv["discount_amount"]

            # Round off determination
            if inv["manual_round_off"] is not None:
                round_off = inv["manual_round_off"]
                grand_total = subtotal + round_off
            elif inv["expected_total"] is not None and inv["expected_total"] > 0:
                grand_total = inv["expected_total"]
                round_off = round(grand_total - subtotal, 2)
            else:
                grand_total = round(subtotal)
                round_off = round(grand_total - subtotal, 2)

            taxes_list = [{"ledger": k, "amount": round(v, 2)} for k, v in inv["taxes_dict"].items()]

            inv_obj = {
                "vch_no": inv["vch_no"],
                "date": inv["date"],
                "party_name": inv["party_name"],
                "reference": inv["reference"],
                "narration": inv["narration"],
                "items": inv["items"],
                "taxes": taxes_list,
                "taxable_amount": items_total,
                "discount_amount": round(inv["discount_amount"], 2),
                "discount_ledger": inv["discount_ledger"],
                "round_off": round_off,
                "round_off_ledger": "Round Off",
                "party_amount": round(grand_total, 2)
            }
            result_invoices.append(inv_obj)

        return result_invoices, errors

    # =========================================================================
    # PURCHASE / CREDIT NOTE PARSER
    # =========================================================================
    @classmethod
    def parse_purchase_sheet(cls, df: pd.DataFrame, col_map: Optional[Dict[str, str]] = None) -> Tuple[List[Dict[str, Any]], List[str]]:
        """
        Parses Purchase or Credit Note transactions from DataFrame.
        """
        cols = list(df.columns)
        errors = []

        c_vch_no = col_map.get("vch_no") if col_map else find_column(cols, ["Voucher No", "Vch No", "Purchase No", "Entry No"])
        c_supp_inv = col_map.get("supp_inv") if col_map else find_column(cols, ["Supplier Inv No", "Supplier Invoice No", "Bill No", "Ref No", "Invoice No"])
        c_date = col_map.get("date") if col_map else find_column(cols, ["Date", "Voucher Date", "Purchase Date"])
        c_supp_date = col_map.get("supp_date") if col_map else find_column(cols, ["Supplier Inv Date", "Bill Date", "Invoice Date"])
        c_party = col_map.get("party") if col_map else find_column(cols, ["Supplier Name", "Party Name", "Vendor Name", "Party Ledger", "Party"])
        c_item = col_map.get("item") if col_map else find_column(cols, ["Stock Item", "Item Name", "Item Description", "Product", "Item"])
        c_qty = col_map.get("qty") if col_map else find_column(cols, ["Qty", "Quantity", "Billed Qty"])
        c_rate = col_map.get("rate") if col_map else find_column(cols, ["Rate", "Price", "Unit Rate"])
        c_unit = col_map.get("unit") if col_map else find_column(cols, ["Unit", "UOM"])
        c_godown = col_map.get("godown") if col_map else find_column(cols, ["Godown", "Location"])
        c_purch_ledger = col_map.get("purch_ledger") if col_map else find_column(cols, ["Purchase Ledger", "Expense Ledger", "Purchase Account"])
        c_taxable = col_map.get("taxable") if col_map else find_column(cols, ["Taxable Amount", "Taxable Value", "Basic Amount"])
        c_gst_pct = col_map.get("gst_pct") if col_map else find_column(cols, ["GST %", "GST Rate", "Tax Rate"])
        c_cgst = col_map.get("cgst") if col_map else find_column(cols, ["Input CGST", "CGST", "CGST Amount"])
        c_sgst = col_map.get("sgst") if col_map else find_column(cols, ["Input SGST", "SGST", "SGST Amount"])
        c_igst = col_map.get("igst") if col_map else find_column(cols, ["Input IGST", "IGST", "IGST Amount"])
        c_disc1 = col_map.get("disc1") if col_map else find_column(cols, ["Discount %", "Item Disc %"])
        c_disc2 = col_map.get("disc2") if col_map else find_column(cols, ["Discount Amount", "Level 2 Discount", "Discount Received"])
        c_disc2_ledger = col_map.get("disc2_ledger") if col_map else find_column(cols, ["Discount Ledger"])
        c_round_off = col_map.get("round_off") if col_map else find_column(cols, ["Round Off", "Roundoff"])
        c_total = col_map.get("total") if col_map else find_column(cols, ["Total Amount", "Invoice Amount", "Net Amount", "Grand Total"])
        c_narration = col_map.get("narration") if col_map else find_column(cols, ["Narration", "Remarks"])

        if not c_vch_no and not c_supp_inv and not c_date:
            errors.append("Critical Error: Could not detect Voucher Number, Supplier Invoice No, or Date columns.")
            return [], errors

        invoices_dict: Dict[str, Dict[str, Any]] = {}

        for idx, row in df.iterrows():
            row_idx = idx + 2
            vch_no = clean_str(row.get(c_vch_no)) if c_vch_no else ""
            supp_inv = clean_str(row.get(c_supp_inv)) if c_supp_inv else ""
            date_val = row.get(c_date) if c_date else None

            if not vch_no and not supp_inv and pd.isna(date_val):
                continue

            group_key = vch_no or supp_inv or f"PUR-{row_idx}"
            party = clean_str(row.get(c_party)) if c_party else "Cash"
            if not party:
                party = "Cash"

            if group_key not in invoices_dict:
                invoices_dict[group_key] = {
                    "vch_no": group_key,
                    "supplier_inv_no": supp_inv or group_key,
                    "date": date_val,
                    "supplier_inv_date": row.get(c_supp_date) if c_supp_date else date_val,
                    "party_name": party,
                    "narration": clean_str(row.get(c_narration)) if c_narration else "",
                    "items": [],
                    "taxes_dict": {},
                    "discount_amount": 0.0,
                    "discount_ledger": clean_str(row.get(c_disc2_ledger)) if c_disc2_ledger else "Discount Received",
                    "manual_round_off": None,
                    "expected_total": None
                }

            inv = invoices_dict[group_key]

            item_name = clean_str(row.get(c_item)) if c_item else ""
            qty = clean_num(row.get(c_qty), 0.0) if c_qty else 0.0
            rate = clean_num(row.get(c_rate), 0.0) if c_rate else 0.0
            unit = clean_str(row.get(c_unit), "PCS") if c_unit else "PCS"
            godown = clean_str(row.get(c_godown), "Main Location") if c_godown else "Main Location"
            purch_ledger = clean_str(row.get(c_purch_ledger), "Purchase Account") if c_purch_ledger else "Purchase Account"
            disc1_pct = clean_num(row.get(c_disc1), 0.0) if c_disc1 else 0.0

            gross_amount = qty * rate if (qty and rate) else clean_num(row.get(c_taxable), 0.0)
            item_taxable = gross_amount
            if disc1_pct > 0:
                item_taxable = gross_amount * (1.0 - (disc1_pct / 100.0))
            elif c_taxable and clean_num(row.get(c_taxable), 0.0) > 0:
                item_taxable = clean_num(row.get(c_taxable), 0.0)

            if item_name:
                inv["items"].append({
                    "stock_item": item_name,
                    "qty": qty if qty else 1.0,
                    "unit": unit or "PCS",
                    "rate": rate if rate else item_taxable,
                    "discount_pct": disc1_pct,
                    "amount": round(item_taxable, 2),
                    "godown": godown or "Main Location",
                    "purchase_ledger": purch_ledger or "Purchase Account"
                })

            cgst_val = clean_num(row.get(c_cgst), 0.0) if c_cgst else 0.0
            sgst_val = clean_num(row.get(c_sgst), 0.0) if c_sgst else 0.0
            igst_val = clean_num(row.get(c_igst), 0.0) if c_igst else 0.0
            gst_pct = clean_num(row.get(c_gst_pct), 0.0) if c_gst_pct else 0.0

            if (cgst_val == 0 and sgst_val == 0 and igst_val == 0) and gst_pct > 0:
                half_rate = gst_pct / 2.0
                cgst_val = round(item_taxable * (half_rate / 100.0), 2)
                sgst_val = round(item_taxable * (half_rate / 100.0), 2)
                cgst_key = f"Input CGST {half_rate:g}%"
                sgst_key = f"Input SGST {half_rate:g}%"
            else:
                cgst_key = "Input CGST"
                sgst_key = "Input SGST"

            if cgst_val > 0:
                inv["taxes_dict"][cgst_key] = inv["taxes_dict"].get(cgst_key, 0.0) + cgst_val
            if sgst_val > 0:
                inv["taxes_dict"][sgst_key] = inv["taxes_dict"].get(sgst_key, 0.0) + sgst_val
            if igst_val > 0:
                inv["taxes_dict"]["Input IGST"] = inv["taxes_dict"].get("Input IGST", 0.0) + igst_val

            l2_disc = clean_num(row.get(c_disc2), 0.0) if c_disc2 else 0.0
            if l2_disc > 0:
                inv["discount_amount"] += l2_disc

            if c_round_off and not pd.isna(row.get(c_round_off)):
                inv["manual_round_off"] = clean_num(row.get(c_round_off), 0.0)
            if c_total and not pd.isna(row.get(c_total)):
                inv["expected_total"] = clean_num(row.get(c_total), 0.0)

        result_invoices = []
        for key, inv in invoices_dict.items():
            items_total = sum(itm["amount"] for itm in inv["items"])
            tax_total = sum(inv["taxes_dict"].values())
            subtotal = items_total + tax_total - inv["discount_amount"]

            if inv["manual_round_off"] is not None:
                round_off = inv["manual_round_off"]
                grand_total = subtotal + round_off
            elif inv["expected_total"] is not None and inv["expected_total"] > 0:
                grand_total = inv["expected_total"]
                round_off = round(grand_total - subtotal, 2)
            else:
                grand_total = round(subtotal)
                round_off = round(grand_total - subtotal, 2)

            taxes_list = [{"ledger": k, "amount": round(v, 2)} for k, v in inv["taxes_dict"].items()]

            inv_obj = {
                "vch_no": inv["vch_no"],
                "supplier_inv_no": inv["supplier_inv_no"],
                "date": inv["date"],
                "supplier_inv_date": inv["supplier_inv_date"],
                "party_name": inv["party_name"],
                "narration": inv["narration"],
                "items": inv["items"],
                "taxes": taxes_list,
                "taxable_amount": items_total,
                "discount_amount": round(inv["discount_amount"], 2),
                "discount_ledger": inv["discount_ledger"],
                "round_off": round_off,
                "round_off_ledger": "Round Off",
                "party_amount": round(grand_total, 2)
            }
            result_invoices.append(inv_obj)

        return result_invoices, errors

    # =========================================================================
    # BANK STATEMENT PARSER (2-Column & 1-Column)
    # =========================================================================
    @classmethod
    def parse_bank_sheet(cls, df: pd.DataFrame, col_map: Optional[Dict[str, str]] = None) -> Tuple[List[Dict[str, Any]], List[str]]:
        """
        Parses Bank Statement transactions.
        Handles both 2-column format (Withdrawal/Debit & Deposit/Credit)
        and 1-column format (Single Amount + Dr/Cr).
        Returns: (transactions_list, errors_list)
        """
        cols = list(df.columns)
        errors = []

        c_date = col_map.get("date") if col_map else find_column(cols, ["Txn Date", "Transaction Date", "Date", "Value Date"])
        c_narration = col_map.get("narration") if col_map else find_column(cols, ["Narration", "Description", "Particulars", "Remarks"])
        c_ref = col_map.get("ref") if col_map else find_column(cols, ["Chq/Ref No", "Cheque No", "Ref No", "Transaction ID", "Reference"])
        
        # 2-column checks
        c_withdrawal = col_map.get("withdrawal") if col_map else find_column(cols, ["Withdrawal", "Debit", "Dr", "Withdrawal (Dr)", "Debit Amount"])
        c_deposit = col_map.get("deposit") if col_map else find_column(cols, ["Deposit", "Credit", "Cr", "Deposit (Cr)", "Credit Amount"])
        
        # 1-column checks
        c_amount = col_map.get("amount") if col_map else find_column(cols, ["Amount", "Txn Amount", "Transaction Amount"])
        c_drcr = col_map.get("drcr") if col_map else find_column(cols, ["Dr / Cr", "Dr/Cr", "Type", "Transaction Type", "CR/DR"])

        is_two_column = (c_withdrawal is not None or c_deposit is not None)

        if not c_date:
            errors.append("Critical Error: Could not detect 'Date' column in bank statement.")
            return [], errors

        if not is_two_column and not c_amount:
            errors.append("Critical Error: Could not detect Amount columns (Withdrawal/Deposit or Amount+Dr/Cr).")
            return [], errors

        transactions = []
        for idx, row in df.iterrows():
            row_idx = idx + 2
            date_val = row.get(c_date)
            if pd.isna(date_val):
                continue

            narration = clean_str(row.get(c_narration)) if c_narration else ""
            # Ignore summary / footer rows
            if any(skip_kw in narration.upper() for skip_kw in ["OPENING BALANCE", "CLOSING BALANCE", "TOTAL", "PAGE"]):
                continue

            ref_no = clean_str(row.get(c_ref)) if c_ref else ""

            if is_two_column:
                w_val = clean_num(row.get(c_withdrawal), 0.0) if c_withdrawal else 0.0
                d_val = clean_num(row.get(c_deposit), 0.0) if c_deposit else 0.0
                
                if w_val > 0.0001:
                    transactions.append({
                        "row_id": row_idx,
                        "date": date_val,
                        "vch_no": f"TXN{row_idx:05d}",
                        "ref_no": ref_no,
                        "narration": narration,
                        "direction": "WITHDRAWAL",
                        "amount": round(w_val, 2),
                        "assigned_ledger": "",  # To be filled by rules engine
                        "vch_type": "Payment"
                    })
                elif d_val > 0.0001:
                    transactions.append({
                        "row_id": row_idx,
                        "date": date_val,
                        "vch_no": f"TXN{row_idx:05d}",
                        "ref_no": ref_no,
                        "narration": narration,
                        "direction": "DEPOSIT",
                        "amount": round(d_val, 2),
                        "assigned_ledger": "",  # To be filled by rules engine
                        "vch_type": "Receipt"
                    })
            else:
                # 1-column format
                amt_val = clean_num(row.get(c_amount), 0.0)
                drcr_val = clean_str(row.get(c_drcr)).upper() if c_drcr else ""

                if abs(amt_val) < 0.0001:
                    continue

                # Determine direction from Dr/Cr or sign of amount
                if "DR" in drcr_val or "DEBIT" in drcr_val or "WITHDRAWAL" in drcr_val or amt_val < 0:
                    direction = "WITHDRAWAL"
                    vch_type = "Payment"
                else:
                    direction = "DEPOSIT"
                    vch_type = "Receipt"

                transactions.append({
                    "row_id": row_idx,
                    "date": date_val,
                    "vch_no": f"TXN{row_idx:05d}",
                    "ref_no": ref_no,
                    "narration": narration,
                    "direction": direction,
                    "amount": round(abs(amt_val), 2),
                    "assigned_ledger": "",
                    "vch_type": vch_type
                })

        return transactions, errors

    # =========================================================================
    # VERTICAL / JOURNAL PARSER
    # =========================================================================
    @classmethod
    def parse_vertical_sheet(cls, df: pd.DataFrame, col_map: Optional[Dict[str, str]] = None) -> Tuple[List[Dict[str, Any]], List[str]]:
        """
        Parses Vertical / Journal vouchers from DataFrame.
        Multi-row Dr/Cr entries grouped by Voucher Number.
        Validates that for each voucher, Total Debit == Total Credit.
        Returns: (vouchers_list, validation_warnings_and_errors)
        """
        cols = list(df.columns)
        errors = []

        c_vch_no = col_map.get("vch_no") if col_map else find_column(cols, ["Voucher No", "Vch No", "Voucher Number", "Entry No"])
        c_vch_type = col_map.get("vch_type") if col_map else find_column(cols, ["Voucher Type", "Vch Type", "Type"])
        c_date = col_map.get("date") if col_map else find_column(cols, ["Date", "Voucher Date"])
        c_ledger = col_map.get("ledger") if col_map else find_column(cols, ["Ledger Name", "Ledger", "Account Name", "Particulars"])
        c_drcr = col_map.get("drcr") if col_map else find_column(cols, ["Dr / Cr", "Dr/Cr", "DrCr", "Debit/Credit"])
        c_amount = col_map.get("amount") if col_map else find_column(cols, ["Amount", "Ledger Amount", "Value"])
        c_narration = col_map.get("narration") if col_map else find_column(cols, ["Narration", "Remarks", "Line Narration"])
        c_cost_center = col_map.get("cost_center") if col_map else find_column(cols, ["Cost Center", "Cost Centre", "Department"])

        if not c_vch_no and not c_date:
            errors.append("Critical Error: Missing Voucher Number or Date columns in Vertical sheet.")
            return [], errors

        if not c_ledger or not c_amount:
            errors.append("Critical Error: Missing Ledger Name or Amount column in Vertical sheet.")
            return [], errors

        vouchers_dict: Dict[str, Dict[str, Any]] = {}

        for idx, row in df.iterrows():
            row_idx = idx + 2
            vch_no = clean_str(row.get(c_vch_no)) if c_vch_no else ""
            date_val = row.get(c_date) if c_date else None
            ledger = clean_str(row.get(c_ledger)) if c_ledger else ""
            amt = clean_num(row.get(c_amount), 0.0) if c_amount else 0.0

            if not vch_no and pd.isna(date_val) and not ledger:
                continue

            if not vch_no:
                vch_no = f"JV-{row_idx}"

            vch_type = clean_str(row.get(c_vch_type), "Journal") if c_vch_type else "Journal"
            if not vch_type:
                vch_type = "Journal"

            drcr_str = clean_str(row.get(c_drcr), "Dr") if c_drcr else "Dr"
            is_dr = drcr_str.lower() in ["dr", "debit", "d"]
            drcr_normalized = "Dr" if is_dr else "Cr"

            narration = clean_str(row.get(c_narration)) if c_narration else ""
            cost_center = clean_str(row.get(c_cost_center)) if c_cost_center else ""

            if vch_no not in vouchers_dict:
                vouchers_dict[vch_no] = {
                    "vch_no": vch_no,
                    "vch_type": vch_type,
                    "date": date_val,
                    "narration": narration,
                    "entries": [],
                    "total_dr": 0.0,
                    "total_cr": 0.0,
                    "balanced": True,
                    "diff": 0.0
                }

            vch = vouchers_dict[vch_no]
            # Accumulate narration if voucher header narration was empty
            if not vch["narration"] and narration:
                vch["narration"] = narration

            vch["entries"].append({
                "ledger": ledger,
                "type": drcr_normalized,
                "amount": round(abs(amt), 2),
                "cost_center": cost_center
            })

            if is_dr:
                vch["total_dr"] += abs(amt)
            else:
                vch["total_cr"] += abs(amt)

        # Validate balancing
        result_vouchers = []
        for vch_no, vch in vouchers_dict.items():
            total_dr = round(vch["total_dr"], 2)
            total_cr = round(vch["total_cr"], 2)
            diff = round(abs(total_dr - total_cr), 2)
            vch["total_dr"] = total_dr
            vch["total_cr"] = total_cr
            vch["diff"] = diff
            vch["balanced"] = (diff <= 0.01)

            if not vch["balanced"]:
                errors.append(f"Voucher {vch_no} is Imbalanced: Total Dr={total_dr:.2f}, Total Cr={total_cr:.2f} (Diff: {diff:.2f})")

            result_vouchers.append(vch)

        return result_vouchers, errors

    # =========================================================================
    # MASTER EXTRACTORS
    # =========================================================================
    @classmethod
    def extract_masters_from_transactions(
        cls,
        sales_invoices: Optional[List[Dict[str, Any]]] = None,
        purch_invoices: Optional[List[Dict[str, Any]]] = None,
        bank_txs: Optional[List[Dict[str, Any]]] = None,
        vertical_vchs: Optional[List[Dict[str, Any]]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """
        Extracts unique Ledgers and Stock Items across all parsed transaction sets.
        Returns: (ledgers_list, stock_items_list)
        """
        ledgers_map: Dict[str, Dict[str, Any]] = {}
        stock_items_map: Dict[str, Dict[str, Any]] = {}

        # 1. From Sales
        if sales_invoices:
            for inv in sales_invoices:
                party = inv.get("party_name", "").strip()
                if party and party.lower() != "cash" and party not in ledgers_map:
                    ledgers_map[party] = {"name": party, "parent": "Sundry Debtors", "type": "Customer"}

                for itm in inv.get("items", []):
                    itm_name = itm.get("stock_item", "").strip()
                    if itm_name and itm_name not in stock_items_map:
                        stock_items_map[itm_name] = {
                            "name": itm_name,
                            "parent": "Primary",
                            "unit": itm.get("unit", "PCS"),
                            "godown": itm.get("godown", "Main Location")
                        }
                    sl = itm.get("sales_ledger", "").strip()
                    if sl and sl not in ledgers_map:
                        ledgers_map[sl] = {"name": sl, "parent": "Sales Accounts", "type": "Income"}

                for tax in inv.get("taxes", []):
                    tl = tax.get("ledger", "").strip()
                    if tl and tl not in ledgers_map:
                        ledgers_map[tl] = {"name": tl, "parent": "Duties & Taxes", "type": "Tax"}

                dl = inv.get("discount_ledger", "").strip()
                if dl and dl not in ledgers_map:
                    ledgers_map[dl] = {"name": dl, "parent": "Indirect Expenses", "type": "Discount"}

                ro = inv.get("round_off_ledger", "Round Off").strip()
                if ro and ro not in ledgers_map:
                    ledgers_map[ro] = {"name": ro, "parent": "Indirect Expenses", "type": "Round Off"}

        # 2. From Purchase
        if purch_invoices:
            for inv in purch_invoices:
                party = inv.get("party_name", "").strip()
                if party and party.lower() != "cash" and party not in ledgers_map:
                    ledgers_map[party] = {"name": party, "parent": "Sundry Creditors", "type": "Vendor"}

                for itm in inv.get("items", []):
                    itm_name = itm.get("stock_item", "").strip()
                    if itm_name and itm_name not in stock_items_map:
                        stock_items_map[itm_name] = {
                            "name": itm_name,
                            "parent": "Primary",
                            "unit": itm.get("unit", "PCS"),
                            "godown": itm.get("godown", "Main Location")
                        }
                    pl = itm.get("purchase_ledger", "").strip()
                    if pl and pl not in ledgers_map:
                        ledgers_map[pl] = {"name": pl, "parent": "Purchase Accounts", "type": "Expense"}

                for tax in inv.get("taxes", []):
                    tl = tax.get("ledger", "").strip()
                    if tl and tl not in ledgers_map:
                        ledgers_map[tl] = {"name": tl, "parent": "Duties & Taxes", "type": "Tax"}

        # 3. From Bank
        if bank_txs:
            for tx in bank_txs:
                al = tx.get("assigned_ledger", "").strip()
                if al and al not in ledgers_map:
                    # Guess parent
                    parent = "Indirect Expenses"
                    if "SUSPENSE" in al.upper():
                        parent = "Suspense A/c"
                    elif "CASH" in al.upper():
                        parent = "Cash-in-Hand"
                    elif "INTEREST" in al.upper():
                        parent = "Indirect Incomes"
                    elif "TAX" in al.upper() or "GST" in al.upper() or "TDS" in al.upper():
                        parent = "Duties & Taxes"
                    ledgers_map[al] = {"name": al, "parent": parent, "type": "Bank Contra/Expense"}

        # 4. From Vertical
        if vertical_vchs:
            for vch in vertical_vchs:
                for entry in vch.get("entries", []):
                    l_name = entry.get("ledger", "").strip()
                    if l_name and l_name not in ledgers_map:
                        ledgers_map[l_name] = {"name": l_name, "parent": "Current Assets", "type": "General"}

        return list(ledgers_map.values()), list(stock_items_map.values())
