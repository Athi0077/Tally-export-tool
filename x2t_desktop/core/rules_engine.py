"""
Bank Statement Categorization Rules Engine
Auto-categorizes transactions based on narration keywords and transaction direction.
"""

import os
import json
import re
from typing import List, Dict, Tuple, Optional
from core.resource_path import get_resource_path


DEFAULT_RULES: List[Dict[str, str]] = [
    # Cash / ATM
    {"keyword": "ATM", "ledger": "Cash", "vch_type": "Contra", "match_type": "contains"},
    {"keyword": "CASH WDL", "ledger": "Cash", "vch_type": "Contra", "match_type": "contains"},
    {"keyword": "CWDR", "ledger": "Cash", "vch_type": "Contra", "match_type": "contains"},
    {"keyword": "CDM", "ledger": "Cash", "vch_type": "Contra", "match_type": "contains"},
    {"keyword": "CASH DEP", "ledger": "Cash", "vch_type": "Contra", "match_type": "contains"},
    {"keyword": "CDEP", "ledger": "Cash", "vch_type": "Contra", "match_type": "contains"},
    {"keyword": "SELF", "ledger": "Cash", "vch_type": "Contra", "match_type": "contains"},
    
    # Bank Charges
    {"keyword": "CHG", "ledger": "Bank Charges", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "CHARGES", "ledger": "Bank Charges", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "CHQ RTN", "ledger": "Bank Charges", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "SMS CHG", "ledger": "Bank Charges", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "ANNUAL FEE", "ledger": "Bank Charges", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "CONSOLIDATED CHG", "ledger": "Bank Charges", "vch_type": "Payment", "match_type": "contains"},
    
    # Interest
    {"keyword": "INT PD", "ledger": "Bank Interest Received", "vch_type": "Receipt", "match_type": "contains"},
    {"keyword": "INTEREST", "ledger": "Bank Interest Received", "vch_type": "Receipt", "match_type": "contains"},
    {"keyword": "INT.REC", "ledger": "Bank Interest Received", "vch_type": "Receipt", "match_type": "contains"},
    {"keyword": "CREDIT INT", "ledger": "Bank Interest Received", "vch_type": "Receipt", "match_type": "contains"},

    # Recurring utilities / expenses
    {"keyword": "SALARY", "ledger": "Salary Expenses", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "PAYROLL", "ledger": "Salary Expenses", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "ELECTRICITY", "ledger": "Electricity Expenses", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "POWER", "ledger": "Electricity Expenses", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "BESCOM", "ledger": "Electricity Expenses", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "RENT", "ledger": "Rent Expenses", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "OFFICE RENT", "ledger": "Rent Expenses", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "INTERNET", "ledger": "Internet & Telephone Expenses", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "AIRTEL", "ledger": "Internet & Telephone Expenses", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "JIO", "ledger": "Internet & Telephone Expenses", "vch_type": "Payment", "match_type": "contains"},

    # Taxes & Government Dues
    {"keyword": "GST", "ledger": "GST Payable", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "TDS", "ledger": "TDS Payable", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "CHALLAN", "ledger": "Direct Taxes", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "INCOME TAX", "ledger": "Income Tax", "vch_type": "Payment", "match_type": "contains"},

    # Insurance
    {"keyword": "LIC", "ledger": "Life Insurance", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "INSURANCE", "ledger": "Insurance Expenses", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "PREMIUM", "ledger": "Insurance Expenses", "vch_type": "Payment", "match_type": "contains"},

    # Fuel
    {"keyword": "PETROL", "ledger": "Fuel & Conveyance", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "DIESEL", "ledger": "Fuel & Conveyance", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "HPCL", "ledger": "Fuel & Conveyance", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "BPCL", "ledger": "Fuel & Conveyance", "vch_type": "Payment", "match_type": "contains"},
    {"keyword": "IOCL", "ledger": "Fuel & Conveyance", "vch_type": "Payment", "match_type": "contains"}
]


class RulesEngine:
    def __init__(self, rules_file: Optional[str] = None):
        self.rules_file = rules_file or get_resource_path("templates/rules.json")
        self.rules: List[Dict[str, str]] = []
        self.default_suspense_ledger = "Suspense A/c"
        self.load_rules()

    def load_rules(self) -> None:
        if os.path.exists(self.rules_file):
            try:
                with open(self.rules_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        self.rules = data
                        return
            except Exception:
                pass
        # Fallback to default rules
        self.rules = [dict(r) for r in DEFAULT_RULES]
        self.save_rules()

    def save_rules(self) -> bool:
        try:
            os.makedirs(os.path.dirname(self.rules_file), exist_ok=True)
            with open(self.rules_file, "w", encoding="utf-8") as f:
                json.dump(self.rules, f, indent=2)
            return True
        except Exception:
            return False

    def reset_to_defaults(self) -> None:
        self.rules = [dict(r) for r in DEFAULT_RULES]
        self.save_rules()

    def add_rule(self, keyword: str, ledger: str, vch_type: str = "Payment", match_type: str = "contains") -> bool:
        keyword = keyword.strip()
        ledger = ledger.strip()
        if not keyword or not ledger:
            return False
        # Avoid duplicate keywords
        for rule in self.rules:
            if rule["keyword"].lower() == keyword.lower():
                rule["ledger"] = ledger
                rule["vch_type"] = vch_type
                rule["match_type"] = match_type
                self.save_rules()
                return True
        self.rules.append({
            "keyword": keyword,
            "ledger": ledger,
            "vch_type": vch_type,
            "match_type": match_type
        })
        self.save_rules()
        return True

    def delete_rule(self, index: int) -> bool:
        if 0 <= index < len(self.rules):
            self.rules.pop(index)
            self.save_rules()
            return True
        return False

    def match(self, narration: str) -> Optional[Dict[str, str]]:
        if not narration or not isinstance(narration, str):
            return None
        text = narration.strip().upper()
        for rule in self.rules:
            kw = rule.get("keyword", "").upper()
            mtype = rule.get("match_type", "contains").lower()
            if not kw:
                continue
            if mtype == "contains" and kw in text:
                return rule
            elif mtype == "startswith" and text.startswith(kw):
                return rule
            elif mtype == "endswith" and text.endswith(kw):
                return rule
            elif mtype == "regex":
                try:
                    if re.search(kw, text, re.IGNORECASE):
                        return rule
                except re.error:
                    continue
        return None

    def categorize(self, narration: str, tx_direction: str = "WITHDRAWAL") -> Tuple[str, str]:
        """
        Categorizes transaction.
        tx_direction: 'WITHDRAWAL' (Dr to Bank / Outflow) or 'DEPOSIT' (Cr to Bank / Inflow)
        Returns: (ledger_name, vch_type)
        """
        rule = self.match(narration)
        if rule:
            ledger = rule["ledger"]
            # If rule specifies a voucher type (like Contra), honor it, otherwise default based on flow
            vch_type = rule.get("vch_type", "")
            if not vch_type:
                vch_type = "Payment" if tx_direction.upper() == "WITHDRAWAL" else "Receipt"
            return ledger, vch_type
        
        # Default fallback
        ledger = self.default_suspense_ledger
        vch_type = "Payment" if tx_direction.upper() == "WITHDRAWAL" else "Receipt"
        return ledger, vch_type
