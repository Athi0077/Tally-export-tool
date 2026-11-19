"""
Bank Statement Import Tab
Handles 2-column (Withdrawal/Deposit) & 1-column (Amount + Dr/Cr) bank statements.
Auto-categorization engine with narration keyword mapping.
"""

import os
import shutil
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd

from core.excel_parser import ExcelParser
from core.tally_xml_builder import TallyXMLBuilder, wrap_envelope
from core.rules_engine import RulesEngine
from core.resource_path import ensure_sample_file
from ui.common import ModernTable, XMLPreviewModal, TallySyncResponseModal


class BankTab(ctk.CTkFrame):
    def __init__(self, master, get_tally_connector_func, rules_engine: RulesEngine, **kwargs):
        super().__init__(master, **kwargs)
        self.get_tally_connector = get_tally_connector_func
        self.rules_engine = rules_engine

        self.current_file_path = ""
        self.current_sheet_name = ""
        self.df_raw = pd.DataFrame()
        self.transactions = []
        self.validation_errors = []

        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_top_controls()
        self._build_stats_bar()
        self._build_table_preview()
        self._build_action_bar()

    def _build_top_controls(self):
        ctrl_frame = ctk.CTkFrame(self)
        ctrl_frame.grid(row=0, column=0, padx=15, pady=(15, 5), sticky="ew")
        ctrl_frame.grid_columnconfigure(1, weight=1)

        # File Selection
        lbl_file = ctk.CTkLabel(ctrl_frame, text="Bank Statement:", font=("Segoe UI", 12, "bold"))
        lbl_file.grid(row=0, column=0, padx=(15, 5), pady=10, sticky="w")

        self.entry_file = ctk.CTkEntry(ctrl_frame, placeholder_text="Select Bank Statement Excel or CSV file...")
        self.entry_file.grid(row=0, column=1, padx=5, pady=10, sticky="ew")

        btn_browse = ctk.CTkButton(ctrl_frame, text="📁 Browse...", width=100, command=self.browse_file)
        btn_browse.grid(row=0, column=2, padx=5, pady=10)

        # Sheet Selection
        lbl_sheet = ctk.CTkLabel(ctrl_frame, text="Sheet:", font=("Segoe UI", 12, "bold"))
        lbl_sheet.grid(row=0, column=3, padx=(10, 5), pady=10)

        self.opt_sheet = ctk.CTkOptionMenu(ctrl_frame, values=["(No sheets)"], command=self.on_sheet_changed, width=150)
        self.opt_sheet.grid(row=0, column=4, padx=5, pady=10)

        # Bank Ledger Name in Tally
        lbl_bank = ctk.CTkLabel(ctrl_frame, text="Bank Ledger in Tally:", font=("Segoe UI", 12, "bold"))
        lbl_bank.grid(row=0, column=5, padx=(10, 5), pady=10)

        self.entry_bank_ledger = ctk.CTkEntry(ctrl_frame, width=180)
        self.entry_bank_ledger.insert(0, "HDFC Bank Account")
        self.entry_bank_ledger.grid(row=0, column=6, padx=5, pady=10)

        # Sample Template Button
        btn_sample = ctk.CTkButton(
            ctrl_frame,
            text="📥 Sample Template",
            width=130,
            fg_color="#334155",
            command=self.download_sample_template
        )
        btn_sample.grid(row=0, column=7, padx=(5, 15), pady=10)

    def _build_stats_bar(self):
        self.stats_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.stats_frame.grid(row=1, column=0, padx=15, pady=5, sticky="ew")

        self.lbl_stats = ctk.CTkLabel(
            self.stats_frame,
            text="No bank statement loaded.",
            font=("Segoe UI", 11)
        )
        self.lbl_stats.pack(side="left")

    def _build_table_preview(self):
        columns = [
            "#", "Date", "Chq/Ref No", "Narration / Description",
            "Flow", "Amount", "Assigned Ledger", "Voucher Type"
        ]
        self.table = ModernTable(self, columns=columns)
        self.table.grid(row=2, column=0, padx=15, pady=5, sticky="nsew")
        self.table.on_edit_callback = self._on_table_cell_edited

    def _build_action_bar(self):
        act_frame = ctk.CTkFrame(self, fg_color="transparent")
        act_frame.grid(row=3, column=0, padx=15, pady=(5, 15), sticky="ew")

        btn_rules = ctk.CTkButton(
            act_frame,
            text="⚡ Run Auto-Categorization",
            command=self.apply_auto_categorization,
            width=180,
            fg_color="#0284c7"
        )
        btn_rules.pack(side="left", padx=(0, 10))

        self.btn_xml = ctk.CTkButton(
            act_frame,
            text="📄 Generate Tally XML",
            command=self.generate_xml,
            width=160,
            fg_color="#059669"
        )
        self.btn_xml.pack(side="left", padx=10)

        self.btn_sync = ctk.CTkButton(
            act_frame,
            text="🚀 Push to Tally (Direct Sync)",
            command=self.push_to_tally,
            width=210,
            fg_color="#7c3aed"
        )
        self.btn_sync.pack(side="left", padx=10)

        lbl_hint = ctk.CTkLabel(act_frame, text="(Double-click any row to edit Assigned Ledger)", font=("Segoe UI", 10, "italic"), text_color="gray")
        lbl_hint.pack(side="right")

    def browse_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Bank Statements (*.xlsx, *.csv)", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if not path:
            return
        self.current_file_path = path
        self.entry_file.delete(0, "end")
        self.entry_file.insert(0, path)

        sheets = ExcelParser.get_sheet_names(path)
        if sheets:
            self.opt_sheet.configure(values=sheets)
            self.opt_sheet.set(sheets[0])
            self.load_data(sheets[0])
        else:
            messagebox.showerror("Error", "Could not read sheet data from statement.")

    def on_sheet_changed(self, sheet_name: str):
        self.load_data(sheet_name)

    def load_data(self, sheet_name: str):
        if not self.current_file_path:
            return
        self.current_sheet_name = sheet_name
        try:
            self.df_raw = ExcelParser.load_full_sheet(self.current_file_path, sheet_name)
            self.transactions, self.validation_errors = ExcelParser.parse_bank_sheet(self.df_raw)
            self.apply_auto_categorization(show_dialog=False)
            self._populate_table()
            self._update_stats()
        except Exception as e:
            messagebox.showerror("Loading Error", f"Failed to parse bank statement: {str(e)}")

    def apply_auto_categorization(self, show_dialog: bool = True):
        if not self.transactions:
            if show_dialog:
                messagebox.showwarning("Warning", "No transactions to categorize.")
            return

        matched_count = 0
        suspense_count = 0

        for tx in self.transactions:
            narration = tx.get("narration", "")
            direction = tx.get("direction", "WITHDRAWAL")
            ledger, vch_type = self.rules_engine.categorize(narration, direction)
            tx["assigned_ledger"] = ledger
            tx["vch_type"] = vch_type

            if "SUSPENSE" in ledger.upper():
                suspense_count += 1
            else:
                matched_count += 1

        self._populate_table()
        self._update_stats()

        if show_dialog:
            messagebox.showinfo(
                "Categorization Complete",
                f"Auto-categorized {len(self.transactions)} bank transactions:\n\n"
                f"• Automatically matched: {matched_count}\n"
                f"• Suspense A/c (Unmatched): {suspense_count}\n\n"
                "You can double-click on any row in the table to reassign its ledger before export."
            )

    def _populate_table(self):
        rows = []
        for i, tx in enumerate(self.transactions, start=1):
            date_str = str(tx.get("date", "")).split(" ")[0]
            flow_icon = "🔻 Withdrawal" if tx.get("direction") == "WITHDRAWAL" else "🔺 Deposit"
            rows.append([
                str(i),
                date_str,
                tx.get("ref_no", ""),
                tx.get("narration", ""),
                flow_icon,
                f"{tx.get('amount', 0.0):.2f}",
                tx.get("assigned_ledger", "Suspense A/c"),
                tx.get("vch_type", "Payment")
            ])
        self.table.set_data(rows)

    def _on_table_cell_edited(self, row_idx: int, col_name: str, new_val: str):
        if 0 <= row_idx < len(self.transactions):
            tx = self.transactions[row_idx]
            if "Ledger" in col_name:
                tx["assigned_ledger"] = new_val.strip()
            elif "Voucher" in col_name:
                tx["vch_type"] = new_val.strip()
            self._update_stats()

    def _update_stats(self):
        total_tx = len(self.transactions)
        wdl_sum = sum(tx["amount"] for tx in self.transactions if tx["direction"] == "WITHDRAWAL")
        dep_sum = sum(tx["amount"] for tx in self.transactions if tx["direction"] == "DEPOSIT")
        suspense_count = sum(1 for tx in self.transactions if "SUSPENSE" in tx.get("assigned_ledger", "").upper())

        suspense_badge = f" | ⚠️ {suspense_count} in Suspense" if suspense_count > 0 else " | ✅ 100% Categorized"

        self.lbl_stats.configure(
            text=(
                f"Total Entries: {total_tx} | Total Withdrawals: ₹{wdl_sum:,.2f} | "
                f"Total Deposits: ₹{dep_sum:,.2f}{suspense_badge}"
            )
        )

    def get_generated_xml(self) -> str:
        if not self.transactions:
            return ""
        bank_ledger = self.entry_bank_ledger.get().strip() or "Bank Account"
        connector = self.get_tally_connector()
        company = ""
        conn_info = connector.test_connection(timeout=1)
        if conn_info.get("connected"):
            company = conn_info.get("active_company", "")

        msgs = [TallyXMLBuilder.build_bank_voucher(tx, bank_ledger) for tx in self.transactions]
        return wrap_envelope("\n".join(msgs), report_name="Vouchers", company_name=company)

    def generate_xml(self):
        if not self.transactions:
            messagebox.showwarning("Warning", "Please load a bank statement first.")
            return
        xml = self.get_generated_xml()
        XMLPreviewModal(self, xml, title="Bank Vouchers Tally XML Preview", file_name_hint="bank_vouchers_tally.xml")

    def push_to_tally(self):
        if not self.transactions:
            messagebox.showwarning("Warning", "Please load a bank statement first.")
            return
        xml = self.get_generated_xml()
        connector = self.get_tally_connector()

        self.btn_sync.configure(text="⏳ Syncing...", state="disabled")
        self.update_idletasks()

        res = connector.send_xml(xml)
        self.btn_sync.configure(text="🚀 Push to Tally (Direct Sync)", state="normal")
        TallySyncResponseModal(self, res)

    def download_sample_template(self):
        src = ensure_sample_file("bank")
        dst = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="sample_bank_statement.xlsx"
        )
        if dst:
            try:
                shutil.copyfile(src, dst)
                messagebox.showinfo("Saved", f"Sample Bank Statement saved to:\n{dst}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save template: {str(e)}")
