"""
Tally XML Builder Engine
Generates standard, schema-compliant XML for Tally ERP 9 and Tally Prime.
Handles Sales, Debit Note, Purchase, Credit Note, Bank Statements, Vertical Journals,
and Ledger / Stock Item Masters.
"""

import html
import re
from typing import List, Dict, Any, Optional


def escape_xml(val: Any) -> str:
    """Escapes XML special characters for safe inclusion in Tally XML."""
    if val is None:
        return ""
    s = str(val).strip()
    return html.escape(s, quote=True)


def format_tally_date(date_str: Any) -> str:
    """
    Converts various date formats to Tally standard YYYYMMDD.
    Accepts: '2024-04-01', '01-04-2024', '01/04/2024', datetime objects, etc.
    """
    if date_str is None:
        return "20240401"
    
    # If already datetime or pandas Timestamp
    if hasattr(date_str, "strftime"):
        return date_str.strftime("%Y%m%d")
    
    s = str(date_str).strip()
    # Remove time part if any (e.g. '2024-04-01 00:00:00')
    s = s.split(" ")[0].split("T")[0]
    
    # Matches YYYYMMDD directly
    if re.match(r"^\d{8}$", s):
        return s
    
    # Matches YYYY-MM-DD or YYYY/MM/DD
    m_iso = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$", s)
    if m_iso:
        y, m, d = m_iso.groups()
        return f"{int(y):04d}{int(m):02d}{int(d):02d}"
    
    # Matches DD-MM-YYYY or DD/MM/YYYY
    m_in = re.match(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$", s)
    if m_in:
        d, m, y = m_in.groups()
        return f"{int(y):04d}{int(m):02d}{int(d):02d}"
    
    # Fallback: strip non-digits if length 8
    clean = re.sub(r"\D", "", s)
    if len(clean) == 8:
        return clean
        
    return "20240401"


def format_amount(amt: Any) -> float:
    """Converts amount to float, handling strings with commas or currency symbols."""
    if amt is None:
        return 0.0
    if isinstance(amt, (int, float)):
        return float(amt)
    s = str(amt).replace(",", "").replace("₹", "").replace("$", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def wrap_envelope(messages_xml: str, report_name: str = "Vouchers", company_name: str = "") -> str:
    """Wraps TallyMessage XML blocks into standard Tally Import Envelope."""
    company_tag = f"<SVCURRENTCOMPANY>{escape_xml(company_name)}</SVCURRENTCOMPANY>" if company_name else ""
    return f"""<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Import Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>{report_name}</REPORTNAME>
        <STATICVARIABLES>
          {company_tag}
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
{messages_xml}
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>"""


class TallyXMLBuilder:
    """
    Builder for all Tally ERP 9 / Tally Prime XML import formats.
    """

    @staticmethod
    def build_sales_or_dn_voucher(inv: Dict[str, Any], vch_type: str = "Sales") -> str:
        """
        Builds XML for a single Sales or Debit Note invoice.
        inv structure:
        {
            "vch_no": "INV-001",
            "date": "2024-04-01",
            "party_name": "Customer Name",
            "reference": "PO-123",
            "ref_date": "2024-04-01",
            "narration": "Sales for April",
            "items": [
                {
                    "stock_item": "Item A",
                    "qty": 10.0,
                    "unit": "PCS",
                    "rate": 1000.0,
                    "discount_pct": 0.0,  # Level 1 discount %
                    "amount": 10000.0,    # Taxable amount
                    "godown": "Main Location",
                    "sales_ledger": "Sales GST 18%"
                }
            ],
            "taxes": [
                {"ledger": "CGST 9%", "amount": 900.0},
                {"ledger": "SGST 9%", "amount": 900.0},
                {"ledger": "IGST 18%", "amount": 0.0}
            ],
            "discount_ledger": "Cash Discount",  # Level 2 discount
            "discount_amount": 0.0,              # Level 2 discount amount
            "round_off": 0.0,                    # Auto or manual round off
            "party_amount": 11800.0              # Final grand total
        }
        """
        vch_no = escape_xml(inv.get("vch_no", ""))
        vch_date = format_tally_date(inv.get("date", ""))
        party_name = escape_xml(inv.get("party_name", "Cash"))
        reference = escape_xml(inv.get("reference", vch_no))
        narration = escape_xml(inv.get("narration", ""))
        
        party_total = format_amount(inv.get("party_amount", 0.0))
        round_off = format_amount(inv.get("round_off", 0.0))
        discount_l2 = format_amount(inv.get("discount_amount", 0.0))
        discount_l2_ledger = escape_xml(inv.get("discount_ledger", "Discount Allowed"))

        is_sales = vch_type.lower() in ["sales", "sale"]
        # In Sales: Party is Debited (ISDEEMEDPOSITIVE=Yes, Amount=-abs(total))
        # In Debit Note (Sales Return by customer/buyer): Party is Credited (ISDEEMEDPOSITIVE=No, Amount=+abs(total))
        # Or Debit Note to Vendor: Party is Debited.
        # Standard convention: Sales debits Customer; Debit Note (to Vendor or Customer debit) debits Party.
        is_party_debit = True  # Sales debits party, Credit Note credits party
        if vch_type.lower() in ["credit note", "cr note"]:
            is_party_debit = False

        party_deemed_positive = "Yes" if is_party_debit else "No"
        party_xml_amount = f"{-abs(party_total):.2f}" if is_party_debit else f"{abs(party_total):.2f}"
        
        has_inventory = len(inv.get("items", [])) > 0
        obj_view = "Invoice Voucher View" if has_inventory else "Accounting Voucher View"
        is_invoice = "Yes"

        xml_lines = []
        xml_lines.append(f'        <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        xml_lines.append(f'          <VOUCHER VCHTYPE="{escape_xml(vch_type)}" ACTION="Create" OBJVIEW="{obj_view}">')
        xml_lines.append(f'            <DATE>{vch_date}</DATE>')
        xml_lines.append(f'            <VOUCHERTYPENAME>{escape_xml(vch_type)}</VOUCHERTYPENAME>')
        xml_lines.append(f'            <VOUCHERNUMBER>{vch_no}</VOUCHERNUMBER>')
        if reference:
            xml_lines.append(f'            <REFERENCE>{reference}</REFERENCE>')
        xml_lines.append(f'            <PARTYLEDGERNAME>{party_name}</PARTYLEDGERNAME>')
        xml_lines.append(f'            <PARTYNAME>{party_name}</PARTYNAME>')
        xml_lines.append(f'            <PERSISTEDVIEW>{obj_view}</PERSISTEDVIEW>')
        xml_lines.append(f'            <ISINVOICE>{is_invoice}</ISINVOICE>')
        if narration:
            xml_lines.append(f'            <NARRATION>{narration}</NARRATION>')

        # 1. Party Ledger Entry
        xml_lines.append(f'            <LEDGERENTRIES.LIST>')
        xml_lines.append(f'              <LEDGERNAME>{party_name}</LEDGERNAME>')
        xml_lines.append(f'              <ISDEEMEDPOSITIVE>{party_deemed_positive}</ISDEEMEDPOSITIVE>')
        xml_lines.append(f'              <ISPARTYLEDGER>Yes</ISPARTYLEDGER>')
        xml_lines.append(f'              <AMOUNT>{party_xml_amount}</AMOUNT>')
        if vch_no:
            xml_lines.append(f'              <BILLALLOCATIONS.LIST>')
            xml_lines.append(f'                <NAME>{vch_no}</NAME>')
            xml_lines.append(f'                <BILLTYPE>New Ref</BILLTYPE>')
            xml_lines.append(f'                <AMOUNT>{party_xml_amount}</AMOUNT>')
            xml_lines.append(f'              </BILLALLOCATIONS.LIST>')
        xml_lines.append(f'            </LEDGERENTRIES.LIST>')

        # 2. Inventory Items
        for item in inv.get("items", []):
            item_name = escape_xml(item.get("stock_item", ""))
            qty = format_amount(item.get("qty", 0.0))
            unit = escape_xml(item.get("unit", "PCS"))
            rate = format_amount(item.get("rate", 0.0))
            disc_pct = format_amount(item.get("discount_pct", 0.0))
            amt = format_amount(item.get("amount", qty * rate))
            godown = escape_xml(item.get("godown", "Main Location"))
            sales_ledger = escape_xml(item.get("sales_ledger", "Sales Account"))

            # In Sales, inventory is Credited (ISDEEMEDPOSITIVE=No, Amount=+amt)
            item_deemed = "No" if is_sales else "Yes"
            item_xml_amt = f"{abs(amt):.2f}" if is_sales else f"{-abs(amt):.2f}"

            qty_str = f"{qty:.2f} {unit}".strip()
            rate_str = f"{rate:.2f}/{unit}".strip()

            xml_lines.append(f'            <ALLINVENTORYENTRIES.LIST>')
            xml_lines.append(f'              <STOCKITEMNAME>{item_name}</STOCKITEMNAME>')
            xml_lines.append(f'              <ISDEEMEDPOSITIVE>{item_deemed}</ISDEEMEDPOSITIVE>')
            xml_lines.append(f'              <RATE>{rate_str}</RATE>')
            if disc_pct > 0:
                xml_lines.append(f'              <DISCOUNT>{disc_pct:.2f}</DISCOUNT>')
            xml_lines.append(f'              <AMOUNT>{item_xml_amt}</AMOUNT>')
            xml_lines.append(f'              <ACTUALQTY>{qty_str}</ACTUALQTY>')
            xml_lines.append(f'              <BILLEDQTY>{qty_str}</BILLEDQTY>')

            # Godown / Batch allocation
            xml_lines.append(f'              <BATCHALLOCATIONS.LIST>')
            xml_lines.append(f'                <GODOWNNAME>{godown}</GODOWNNAME>')
            xml_lines.append(f'                <BATCHNAME>Primary Batch</BATCHNAME>')
            xml_lines.append(f'                <AMOUNT>{item_xml_amt}</AMOUNT>')
            xml_lines.append(f'                <ACTUALQTY>{qty_str}</ACTUALQTY>')
            xml_lines.append(f'                <BILLEDQTY>{qty_str}</BILLEDQTY>')
            xml_lines.append(f'              </BATCHALLOCATIONS.LIST>')

            # Nominal / Revenue Allocation
            xml_lines.append(f'              <ACCOUNTINGALLOCATIONS.LIST>')
            xml_lines.append(f'                <LEDGERNAME>{sales_ledger}</LEDGERNAME>')
            xml_lines.append(f'                <ISDEEMEDPOSITIVE>{item_deemed}</ISDEEMEDPOSITIVE>')
            xml_lines.append(f'                <AMOUNT>{item_xml_amt}</AMOUNT>')
            xml_lines.append(f'              </ACCOUNTINGALLOCATIONS.LIST>')
            xml_lines.append(f'            </ALLINVENTORYENTRIES.LIST>')

        # If pure accounting invoice (no items), add the sales ledger directly
        if not has_inventory and inv.get("taxable_amount"):
            taxable = format_amount(inv.get("taxable_amount", 0.0))
            sales_ledger = escape_xml(inv.get("sales_ledger", "Sales Account"))
            sales_deemed = "No" if is_sales else "Yes"
            sales_amt = f"{abs(taxable):.2f}" if is_sales else f"{-abs(taxable):.2f}"
            xml_lines.append(f'            <LEDGERENTRIES.LIST>')
            xml_lines.append(f'              <LEDGERNAME>{sales_ledger}</LEDGERNAME>')
            xml_lines.append(f'              <ISDEEMEDPOSITIVE>{sales_deemed}</ISDEEMEDPOSITIVE>')
            xml_lines.append(f'              <AMOUNT>{sales_amt}</AMOUNT>')
            xml_lines.append(f'            </LEDGERENTRIES.LIST>')

        # 3. Level 2 Discount Ledger Line Item
        if discount_l2 > 0:
            # Discount given reduces sales, so it is Debited in sales (ISDEEMEDPOSITIVE=Yes, Amount=-disc)
            disc_amt_str = f"{-abs(discount_l2):.2f}"
            xml_lines.append(f'            <LEDGERENTRIES.LIST>')
            xml_lines.append(f'              <LEDGERNAME>{discount_l2_ledger}</LEDGERNAME>')
            xml_lines.append(f'              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>')
            xml_lines.append(f'              <AMOUNT>{disc_amt_str}</AMOUNT>')
            xml_lines.append(f'            </LEDGERENTRIES.LIST>')

        # 4. Tax Ledger Entries (CGST, SGST, IGST, etc.)
        for tax in inv.get("taxes", []):
            tax_ledger = escape_xml(tax.get("ledger", ""))
            tax_amt = format_amount(tax.get("amount", 0.0))
            if tax_ledger and abs(tax_amt) > 0.0001:
                # In Sales, Output Tax is Credited (ISDEEMEDPOSITIVE=No, Amount=+tax)
                tax_deemed = "No" if is_sales else "Yes"
                tax_str = f"{abs(tax_amt):.2f}" if is_sales else f"{-abs(tax_amt):.2f}"
                xml_lines.append(f'            <LEDGERENTRIES.LIST>')
                xml_lines.append(f'              <LEDGERNAME>{tax_ledger}</LEDGERNAME>')
                xml_lines.append(f'              <ISDEEMEDPOSITIVE>{tax_deemed}</ISDEEMEDPOSITIVE>')
                xml_lines.append(f'              <AMOUNT>{tax_str}</AMOUNT>')
                xml_lines.append(f'            </LEDGERENTRIES.LIST>')

        # 5. Round-off Ledger Entry
        if abs(round_off) > 0.0001:
            ro_ledger = escape_xml(inv.get("round_off_ledger", "Round Off"))
            # If round_off is positive, it increases total invoice (Credit: ISDEEMEDPOSITIVE=No, Amount=+ro)
            # If round_off is negative, it reduces total invoice (Debit: ISDEEMEDPOSITIVE=Yes, Amount=-abs(ro))
            if round_off > 0:
                ro_deemed = "No"
                ro_str = f"{abs(round_off):.2f}"
            else:
                ro_deemed = "Yes"
                ro_str = f"{-abs(round_off):.2f}"
            xml_lines.append(f'            <LEDGERENTRIES.LIST>')
            xml_lines.append(f'              <LEDGERNAME>{ro_ledger}</LEDGERNAME>')
            xml_lines.append(f'              <ISDEEMEDPOSITIVE>{ro_deemed}</ISDEEMEDPOSITIVE>')
            xml_lines.append(f'              <AMOUNT>{ro_str}</AMOUNT>')
            xml_lines.append(f'            </LEDGERENTRIES.LIST>')

        xml_lines.append(f'          </VOUCHER>')
        xml_lines.append(f'        </TALLYMESSAGE>')
        return "\n".join(xml_lines)

    @staticmethod
    def build_purchase_or_cn_voucher(inv: Dict[str, Any], vch_type: str = "Purchase") -> str:
        """
        Builds XML for a single Purchase or Credit Note invoice.
        In Purchase:
        - Party (Supplier) is Credited (ISDEEMEDPOSITIVE=No, Amount=+party_total)
        - Inventory / Purchase Account is Debited (ISDEEMEDPOSITIVE=Yes, Amount=-amt)
        - Input GST is Debited (ISDEEMEDPOSITIVE=Yes, Amount=-tax)
        - Round-off balances
        """
        vch_no = escape_xml(inv.get("vch_no", ""))
        vch_date = format_tally_date(inv.get("date", ""))
        party_name = escape_xml(inv.get("party_name", "Cash"))
        supplier_inv_no = escape_xml(inv.get("supplier_inv_no", inv.get("reference", vch_no)))
        supplier_inv_date = format_tally_date(inv.get("supplier_inv_date", inv.get("date", "")))
        narration = escape_xml(inv.get("narration", ""))

        party_total = format_amount(inv.get("party_amount", 0.0))
        round_off = format_amount(inv.get("round_off", 0.0))
        discount_l2 = format_amount(inv.get("discount_amount", 0.0))
        discount_l2_ledger = escape_xml(inv.get("discount_ledger", "Discount Received"))

        is_purchase = vch_type.lower() in ["purchase", "purchases"]
        # In Purchase, Party is Credited (+party_total, ISDEEMEDPOSITIVE=No)
        # In Credit Note (Purchase Return to vendor), Party is Debited (-party_total, ISDEEMEDPOSITIVE=Yes)
        is_party_credit = is_purchase
        party_deemed = "No" if is_party_credit else "Yes"
        party_xml_amt = f"{abs(party_total):.2f}" if is_party_credit else f"{-abs(party_total):.2f}"

        has_inventory = len(inv.get("items", [])) > 0
        obj_view = "Invoice Voucher View" if has_inventory else "Accounting Voucher View"

        xml_lines = []
        xml_lines.append(f'        <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        xml_lines.append(f'          <VOUCHER VCHTYPE="{escape_xml(vch_type)}" ACTION="Create" OBJVIEW="{obj_view}">')
        xml_lines.append(f'            <DATE>{vch_date}</DATE>')
        xml_lines.append(f'            <VOUCHERTYPENAME>{escape_xml(vch_type)}</VOUCHERTYPENAME>')
        xml_lines.append(f'            <VOUCHERNUMBER>{vch_no}</VOUCHERNUMBER>')
        if supplier_inv_no:
            xml_lines.append(f'            <REFERENCE>{supplier_inv_no}</REFERENCE>')
            xml_lines.append(f'            <REFERENCEDATE>{supplier_inv_date}</REFERENCEDATE>')
        xml_lines.append(f'            <PARTYLEDGERNAME>{party_name}</PARTYLEDGERNAME>')
        xml_lines.append(f'            <PARTYNAME>{party_name}</PARTYNAME>')
        xml_lines.append(f'            <PERSISTEDVIEW>{obj_view}</PERSISTEDVIEW>')
        xml_lines.append(f'            <ISINVOICE>Yes</ISINVOICE>')
        if narration:
            xml_lines.append(f'            <NARRATION>{narration}</NARRATION>')

        # 1. Party Ledger Entry (Credit for Purchase)
        xml_lines.append(f'            <LEDGERENTRIES.LIST>')
        xml_lines.append(f'              <LEDGERNAME>{party_name}</LEDGERNAME>')
        xml_lines.append(f'              <ISDEEMEDPOSITIVE>{party_deemed}</ISDEEMEDPOSITIVE>')
        xml_lines.append(f'              <ISPARTYLEDGER>Yes</ISPARTYLEDGER>')
        xml_lines.append(f'              <AMOUNT>{party_xml_amt}</AMOUNT>')
        if supplier_inv_no or vch_no:
            bill_ref = supplier_inv_no or vch_no
            xml_lines.append(f'              <BILLALLOCATIONS.LIST>')
            xml_lines.append(f'                <NAME>{bill_ref}</NAME>')
            xml_lines.append(f'                <BILLTYPE>New Ref</BILLTYPE>')
            xml_lines.append(f'                <AMOUNT>{party_xml_amt}</AMOUNT>')
            xml_lines.append(f'              </BILLALLOCATIONS.LIST>')
        xml_lines.append(f'            </LEDGERENTRIES.LIST>')

        # 2. Inventory Items
        for item in inv.get("items", []):
            item_name = escape_xml(item.get("stock_item", ""))
            qty = format_amount(item.get("qty", 0.0))
            unit = escape_xml(item.get("unit", "PCS"))
            rate = format_amount(item.get("rate", 0.0))
            disc_pct = format_amount(item.get("discount_pct", 0.0))
            amt = format_amount(item.get("amount", qty * rate))
            godown = escape_xml(item.get("godown", "Main Location"))
            purch_ledger = escape_xml(item.get("purchase_ledger", "Purchase Account"))

            # In Purchase, inventory is Debited (ISDEEMEDPOSITIVE=Yes, Amount=-amt)
            item_deemed = "Yes" if is_purchase else "No"
            item_xml_amt = f"{-abs(amt):.2f}" if is_purchase else f"{abs(amt):.2f}"

            qty_str = f"{qty:.2f} {unit}".strip()
            rate_str = f"{rate:.2f}/{unit}".strip()

            xml_lines.append(f'            <ALLINVENTORYENTRIES.LIST>')
            xml_lines.append(f'              <STOCKITEMNAME>{item_name}</STOCKITEMNAME>')
            xml_lines.append(f'              <ISDEEMEDPOSITIVE>{item_deemed}</ISDEEMEDPOSITIVE>')
            xml_lines.append(f'              <RATE>{rate_str}</RATE>')
            if disc_pct > 0:
                xml_lines.append(f'              <DISCOUNT>{disc_pct:.2f}</DISCOUNT>')
            xml_lines.append(f'              <AMOUNT>{item_xml_amt}</AMOUNT>')
            xml_lines.append(f'              <ACTUALQTY>{qty_str}</ACTUALQTY>')
            xml_lines.append(f'              <BILLEDQTY>{qty_str}</BILLEDQTY>')

            # Godown / Batch allocation
            xml_lines.append(f'              <BATCHALLOCATIONS.LIST>')
            xml_lines.append(f'                <GODOWNNAME>{godown}</GODOWNNAME>')
            xml_lines.append(f'                <BATCHNAME>Primary Batch</BATCHNAME>')
            xml_lines.append(f'                <AMOUNT>{item_xml_amt}</AMOUNT>')
            xml_lines.append(f'                <ACTUALQTY>{qty_str}</ACTUALQTY>')
            xml_lines.append(f'                <BILLEDQTY>{qty_str}</BILLEDQTY>')
            xml_lines.append(f'              </BATCHALLOCATIONS.LIST>')

            # Purchase nominal ledger allocation
            xml_lines.append(f'              <ACCOUNTINGALLOCATIONS.LIST>')
            xml_lines.append(f'                <LEDGERNAME>{purch_ledger}</LEDGERNAME>')
            xml_lines.append(f'                <ISDEEMEDPOSITIVE>{item_deemed}</ISDEEMEDPOSITIVE>')
            xml_lines.append(f'                <AMOUNT>{item_xml_amt}</AMOUNT>')
            xml_lines.append(f'              </ACCOUNTINGALLOCATIONS.LIST>')
            xml_lines.append(f'            </ALLINVENTORYENTRIES.LIST>')

        # If pure accounting invoice (no items), add purchase ledger entry directly
        if not has_inventory and inv.get("taxable_amount"):
            taxable = format_amount(inv.get("taxable_amount", 0.0))
            purch_ledger = escape_xml(inv.get("purchase_ledger", "Purchase Account"))
            p_deemed = "Yes" if is_purchase else "No"
            p_amt = f"{-abs(taxable):.2f}" if is_purchase else f"{abs(taxable):.2f}"
            xml_lines.append(f'            <LEDGERENTRIES.LIST>')
            xml_lines.append(f'              <LEDGERNAME>{purch_ledger}</LEDGERNAME>')
            xml_lines.append(f'              <ISDEEMEDPOSITIVE>{p_deemed}</ISDEEMEDPOSITIVE>')
            xml_lines.append(f'              <AMOUNT>{p_amt}</AMOUNT>')
            xml_lines.append(f'            </LEDGERENTRIES.LIST>')

        # 3. Level 2 Discount Ledger Line Item
        if discount_l2 > 0:
            # Discount received reduces purchase expenditure, so it is Credited (ISDEEMEDPOSITIVE=No, Amount=+disc)
            disc_amt_str = f"{abs(discount_l2):.2f}"
            xml_lines.append(f'            <LEDGERENTRIES.LIST>')
            xml_lines.append(f'              <LEDGERNAME>{discount_l2_ledger}</LEDGERNAME>')
            xml_lines.append(f'              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>')
            xml_lines.append(f'              <AMOUNT>{disc_amt_str}</AMOUNT>')
            xml_lines.append(f'            </LEDGERENTRIES.LIST>')

        # 4. Input Tax Ledger Entries (Debit in Purchase)
        for tax in inv.get("taxes", []):
            tax_ledger = escape_xml(tax.get("ledger", ""))
            tax_amt = format_amount(tax.get("amount", 0.0))
            if tax_ledger and abs(tax_amt) > 0.0001:
                tax_deemed = "Yes" if is_purchase else "No"
                tax_str = f"{-abs(tax_amt):.2f}" if is_purchase else f"{abs(tax_amt):.2f}"
                xml_lines.append(f'            <LEDGERENTRIES.LIST>')
                xml_lines.append(f'              <LEDGERNAME>{tax_ledger}</LEDGERNAME>')
                xml_lines.append(f'              <ISDEEMEDPOSITIVE>{tax_deemed}</ISDEEMEDPOSITIVE>')
                xml_lines.append(f'              <AMOUNT>{tax_str}</AMOUNT>')
                xml_lines.append(f'            </LEDGERENTRIES.LIST>')

        # 5. Round-off Ledger Entry
        if abs(round_off) > 0.0001:
            ro_ledger = escape_xml(inv.get("round_off_ledger", "Round Off"))
            if round_off > 0:
                ro_deemed = "Yes"
                ro_str = f"{-abs(round_off):.2f}"
            else:
                ro_deemed = "No"
                ro_str = f"{abs(round_off):.2f}"
            xml_lines.append(f'            <LEDGERENTRIES.LIST>')
            xml_lines.append(f'              <LEDGERNAME>{ro_ledger}</LEDGERNAME>')
            xml_lines.append(f'              <ISDEEMEDPOSITIVE>{ro_deemed}</ISDEEMEDPOSITIVE>')
            xml_lines.append(f'              <AMOUNT>{ro_str}</AMOUNT>')
            xml_lines.append(f'            </LEDGERENTRIES.LIST>')

        xml_lines.append(f'          </VOUCHER>')
        xml_lines.append(f'        </TALLYMESSAGE>')
        return "\n".join(xml_lines)

    @staticmethod
    def build_bank_voucher(tx: Dict[str, Any], bank_ledger: str) -> str:
        """
        Builds XML for a single Bank Statement transaction.
        tx structure:
        {
            "date": "2024-04-01",
            "vch_no": "TX1001",
            "ref_no": "CHQ123456",
            "narration": "UPI/ATM/Cheque description",
            "assigned_ledger": "Bank Charges",
            "vch_type": "Payment",  # "Payment", "Receipt", "Contra"
            "direction": "WITHDRAWAL",  # or "DEPOSIT"
            "amount": 250.0
        }
        """
        vch_date = format_tally_date(tx.get("date", ""))
        vch_no = escape_xml(tx.get("vch_no", ""))
        ref_no = escape_xml(tx.get("ref_no", ""))
        narration = escape_xml(tx.get("narration", ""))
        vch_type = escape_xml(tx.get("vch_type", "Payment"))
        assigned_ledger = escape_xml(tx.get("assigned_ledger", "Suspense A/c"))
        bank_ledger = escape_xml(bank_ledger or "Bank Account")
        amount = format_amount(tx.get("amount", 0.0))
        direction = tx.get("direction", "WITHDRAWAL").upper()

        # If withdrawal:
        # Bank is Credited (ISDEEMEDPOSITIVE=No, Amount=+amount)
        # Assigned Ledger is Debited (ISDEEMEDPOSITIVE=Yes, Amount=-amount)
        # If deposit:
        # Bank is Debited (ISDEEMEDPOSITIVE=Yes, Amount=-amount)
        # Assigned Ledger is Credited (ISDEEMEDPOSITIVE=No, Amount=+amount)

        is_withdrawal = (direction == "WITHDRAWAL")

        bank_deemed = "No" if is_withdrawal else "Yes"
        bank_xml_amt = f"{amount:.2f}" if is_withdrawal else f"{-amount:.2f}"

        party_deemed = "Yes" if is_withdrawal else "No"
        party_xml_amt = f"{-amount:.2f}" if is_withdrawal else f"{amount:.2f}"

        xml_lines = []
        xml_lines.append(f'        <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        xml_lines.append(f'          <VOUCHER VCHTYPE="{vch_type}" ACTION="Create">')
        xml_lines.append(f'            <DATE>{vch_date}</DATE>')
        xml_lines.append(f'            <VOUCHERTYPENAME>{vch_type}</VOUCHERTYPENAME>')
        if vch_no:
            xml_lines.append(f'            <VOUCHERNUMBER>{vch_no}</VOUCHERNUMBER>')
        if ref_no:
            xml_lines.append(f'            <REFERENCE>{ref_no}</REFERENCE>')
        if narration:
            xml_lines.append(f'            <NARRATION>{narration}</NARRATION>')

        # Assigned Ledger Entry (Expense / Customer / Contra)
        xml_lines.append(f'            <ALLLEDGERENTRIES.LIST>')
        xml_lines.append(f'              <LEDGERNAME>{assigned_ledger}</LEDGERNAME>')
        xml_lines.append(f'              <ISDEEMEDPOSITIVE>{party_deemed}</ISDEEMEDPOSITIVE>')
        xml_lines.append(f'              <AMOUNT>{party_xml_amt}</AMOUNT>')
        xml_lines.append(f'            </ALLLEDGERENTRIES.LIST>')

        # Bank Ledger Entry
        xml_lines.append(f'            <ALLLEDGERENTRIES.LIST>')
        xml_lines.append(f'              <LEDGERNAME>{bank_ledger}</LEDGERNAME>')
        xml_lines.append(f'              <ISDEEMEDPOSITIVE>{bank_deemed}</ISDEEMEDPOSITIVE>')
        xml_lines.append(f'              <AMOUNT>{bank_xml_amt}</AMOUNT>')
        if ref_no or narration:
            xml_lines.append(f'              <BANKALLOCATIONS.LIST>')
            xml_lines.append(f'                <DATE>{vch_date}</DATE>')
            if ref_no:
                xml_lines.append(f'                <INSTRUMENTNUMBER>{ref_no}</INSTRUMENTNUMBER>')
            xml_lines.append(f'                <TRANSACTIONTYPE>Cheque/eFund</TRANSACTIONTYPE>')
            xml_lines.append(f'                <AMOUNT>{bank_xml_amt}</AMOUNT>')
            xml_lines.append(f'              </BANKALLOCATIONS.LIST>')
        xml_lines.append(f'            </ALLLEDGERENTRIES.LIST>')

        xml_lines.append(f'          </VOUCHER>')
        xml_lines.append(f'        </TALLYMESSAGE>')
        return "\n".join(xml_lines)

    @staticmethod
    def build_vertical_voucher(vch: Dict[str, Any]) -> str:
        """
        Builds XML for a Vertical / Generic Journal voucher.
        vch structure:
        {
            "vch_no": "JV-01",
            "vch_type": "Journal",  # Payment, Receipt, Journal, Contra
            "date": "2024-04-01",
            "narration": "Voucher narration",
            "entries": [
                {
                    "ledger": "Salary Expenses",
                    "type": "Dr",  # Dr or Cr
                    "amount": 50000.0,
                    "cost_center": ""
                },
                {
                    "ledger": "Bank Account",
                    "type": "Cr",
                    "amount": 50000.0,
                    "cost_center": ""
                }
            ]
        }
        """
        vch_no = escape_xml(vch.get("vch_no", ""))
        vch_type = escape_xml(vch.get("vch_type", "Journal"))
        vch_date = format_tally_date(vch.get("date", ""))
        narration = escape_xml(vch.get("narration", ""))

        xml_lines = []
        xml_lines.append(f'        <TALLYMESSAGE xmlns:UDF="TallyUDF">')
        xml_lines.append(f'          <VOUCHER VCHTYPE="{vch_type}" ACTION="Create">')
        xml_lines.append(f'            <DATE>{vch_date}</DATE>')
        xml_lines.append(f'            <VOUCHERTYPENAME>{vch_type}</VOUCHERTYPENAME>')
        if vch_no:
            xml_lines.append(f'            <VOUCHERNUMBER>{vch_no}</VOUCHERNUMBER>')
        if narration:
            xml_lines.append(f'            <NARRATION>{narration}</NARRATION>')

        for entry in vch.get("entries", []):
            ledger = escape_xml(entry.get("ledger", ""))
            dr_cr = str(entry.get("type", "Dr")).strip().lower()
            amt = format_amount(entry.get("amount", 0.0))
            cost_center = escape_xml(entry.get("cost_center", ""))

            is_dr = dr_cr in ["dr", "debit", "d"]
            deemed = "Yes" if is_dr else "No"
            amt_str = f"{-abs(amt):.2f}" if is_dr else f"{abs(amt):.2f}"

            xml_lines.append(f'            <ALLLEDGERENTRIES.LIST>')
            xml_lines.append(f'              <LEDGERNAME>{ledger}</LEDGERNAME>')
            xml_lines.append(f'              <ISDEEMEDPOSITIVE>{deemed}</ISDEEMEDPOSITIVE>')
            xml_lines.append(f'              <AMOUNT>{amt_str}</AMOUNT>')
            if cost_center:
                xml_lines.append(f'              <CATEGORYALLOCATIONS.LIST>')
                xml_lines.append(f'                <CATEGORY>Primary Cost Category</CATEGORY>')
                xml_lines.append(f'                <COSTCENTREALLOCATIONS.LIST>')
                xml_lines.append(f'                  <NAME>{cost_center}</NAME>')
                xml_lines.append(f'                  <AMOUNT>{amt_str}</AMOUNT>')
                xml_lines.append(f'                </COSTCENTREALLOCATIONS.LIST>')
                xml_lines.append(f'              </CATEGORYALLOCATIONS.LIST>')
            xml_lines.append(f'            </ALLLEDGERENTRIES.LIST>')

        xml_lines.append(f'          </VOUCHER>')
        xml_lines.append(f'        </TALLYMESSAGE>')
        return "\n".join(xml_lines)

    @staticmethod
    def build_ledger_master_xml(ledgers: List[Dict[str, Any]], company_name: str = "") -> str:
        """
        Builds XML for Ledger Masters.
        ledgers item structure:
        {
            "name": "Customer ABC",
            "parent": "Sundry Debtors",
            "opening_balance": 0.0,
            "is_billwise": "Yes"
        }
        """
        messages = []
        for l in ledgers:
            name = escape_xml(l.get("name", ""))
            parent = escape_xml(l.get("parent", "Sundry Debtors"))
            op_bal = format_amount(l.get("opening_balance", 0.0))
            is_billwise = escape_xml(l.get("is_billwise", "Yes"))
            
            if not name:
                continue

            op_xml = ""
            if abs(op_bal) > 0.0001:
                op_deemed = "Yes" if op_bal >= 0 else "No"
                op_amt = f"{-abs(op_bal):.2f}" if op_bal >= 0 else f"{abs(op_bal):.2f}"
                op_xml = f"""
            <OPENINGBALANCE>{op_amt}</OPENINGBALANCE>
            <ISDEEMEDPOSITIVE>{op_deemed}</ISDEEMEDPOSITIVE>"""

            msg = f"""        <TALLYMESSAGE xmlns:UDF="TallyUDF">
          <LEDGER NAME="{name}" ACTION="Create">
            <NAME>{name}</NAME>
            <PARENT>{parent}</PARENT>
            <ISBILLWISEON>{is_billwise}</ISBILLWISEON>
            <AFFECTSSTOCK>No</AFFECTSSTOCK>{op_xml}
          </LEDGER>
        </TALLYMESSAGE>"""
            messages.append(msg)

        all_msgs = "\n".join(messages)
        return wrap_envelope(all_msgs, report_name="All Masters", company_name=company_name)

    @staticmethod
    def build_stock_item_master_xml(stock_items: List[Dict[str, Any]], company_name: str = "") -> str:
        """
        Builds XML for Stock Item Masters and their Base Units.
        stock_items item structure:
        {
            "name": "Widget A",
            "parent": "Primary",
            "unit": "PCS",
            "hsn": "8471",
            "tax_rate": 18.0
        }
        """
        messages = []
        # First gather all unique units
        units_set = set()
        for itm in stock_items:
            u = itm.get("unit", "").strip()
            if u:
                units_set.add(u)

        # Generate Unit Masters first
        for u in sorted(units_set):
            u_esc = escape_xml(u)
            u_msg = f"""        <TALLYMESSAGE xmlns:UDF="TallyUDF">
          <UNIT NAME="{u_esc}" ACTION="Create">
            <NAME>{u_esc}</NAME>
            <ISSIMPLEUNIT>Yes</ISSIMPLEUNIT>
            <ORIGINALNAME>{u_esc}</ORIGINALNAME>
          </UNIT>
        </TALLYMESSAGE>"""
            messages.append(u_msg)

        # Generate Stock Item Masters
        for itm in stock_items:
            name = escape_xml(itm.get("name", ""))
            parent = escape_xml(itm.get("parent", "Primary"))
            unit = escape_xml(itm.get("unit", "PCS"))
            hsn = escape_xml(itm.get("hsn", ""))
            if not name:
                continue

            hsn_xml = f"\n            <HSNCODE>{hsn}</HSNCODE>" if hsn else ""

            msg = f"""        <TALLYMESSAGE xmlns:UDF="TallyUDF">
          <STOCKITEM NAME="{name}" ACTION="Create">
            <NAME>{name}</NAME>
            <PARENT>{parent}</PARENT>
            <BASEUNITS>{unit}</BASEUNITS>
            <ISBATCHWISEON>No</ISBATCHWISEON>{hsn_xml}
            <OPENINGBALANCE>0</OPENINGBALANCE>
          </STOCKITEM>
        </TALLYMESSAGE>"""
            messages.append(msg)

        all_msgs = "\n".join(messages)
        return wrap_envelope(all_msgs, report_name="All Masters", company_name=company_name)
