"""
Sales & Debit Note Tab
Multi-line stock items, GST bifurcations, Level 1 & 2 discounts, Auto round-off.
"""

import os
import shutil
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd

from core.excel_parser import ExcelParser
from core.tally_xml_builder import TallyXMLBuilder, wrap_envelope
from core.resource_path import ensure_sample_file
from ui.common import ModernTable, XMLPreviewModal, TallySyncResponseModal


class SalesTab(ctk.CTkFrame):
    def __init__(self, master, get_tally_connector_func, **kwargs):
        super().__init__(master, **kwargs)
        self.get_tally_connector = get_tally_connector_func
        
        self.current_file_path = ""
        self.current_sheet_name = ""
        self.df_raw = pd.DataFrame()
        self.parsed_invoices = []
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
        lbl_file = ctk.CTkLabel(ctrl_frame, text="Excel File:", font=("Segoe UI", 12, "bold"))
        lbl_file.grid(row=0, column=0, padx=(15, 5), pady=10, sticky="w")

        self.entry_file = ctk.CTkEntry(ctrl_frame, placeholder_text="Select Sales / Debit Note Excel or CSV workbook...")
        self.entry_file.grid(row=0, column=1, padx=5, pady=10, sticky="ew")

        btn_browse = ctk.CTkButton(ctrl_frame, text="📁 Browse...", width=100, command=self.browse_file)
        btn_browse.grid(row=0, column=2, padx=5, pady=10)

        # Sheet Selection
        lbl_sheet = ctk.CTkLabel(ctrl_frame, text="Sheet:", font=("Segoe UI", 12, "bold"))
        lbl_sheet.grid(row=0, column=3, padx=(10, 5), pady=10)

        self.opt_sheet = ctk.CTkOptionMenu(ctrl_frame, values=["(No sheets)"], command=self.on_sheet_changed, width=150)
        self.opt_sheet.grid(row=0, column=4, padx=5, pady=10)

        # Voucher Type Toggle
        lbl_type = ctk.CTkLabel(ctrl_frame, text="Type:", font=("Segoe UI", 12, "bold"))
        lbl_type.grid(row=0, column=5, padx=(10, 5), pady=10)

        self.opt_vch_type = ctk.CTkOptionMenu(ctrl_frame, values=["Sales", "Debit Note"], width=120)
        self.opt_vch_type.grid(row=0, column=6, padx=5, pady=10)

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
            text="No data loaded. Select an Excel file above to begin.",
            font=("Segoe UI", 11)
        )
        self.lbl_stats.pack(side="left")

    def _build_table_preview(self):
        columns = [
            "Vch No", "Date", "Party Name", "Stock Item", "Qty", "Unit",
            "Rate", "Disc %", "Taxable", "CGST", "SGST", "IGST", "L2 Disc", "Round Off", "Total"
        ]
        self.table = ModernTable(self, columns=columns)
        self.table.grid(row=2, column=0, padx=15, pady=5, sticky="nsew")

    def _build_action_bar(self):
        act_frame = ctk.CTkFrame(self, fg_color="transparent")
        act_frame.grid(row=3, column=0, padx=15, pady=(5, 15), sticky="ew")

        btn_validate = ctk.CTkButton(
            act_frame,
            text="🔍 Validate Entries",
            command=self.validate_data,
            width=140,
            fg_color="#0284c7"
        )
        btn_validate.pack(side="left", padx=(0, 10))

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

    def browse_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel & CSV Files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
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
            messagebox.showerror("Error", "Could not read sheets from chosen file.")

    def on_sheet_changed(self, sheet_name: str):
        self.load_data(sheet_name)

    def load_data(self, sheet_name: str):
        if not self.current_file_path:
            return
        self.current_sheet_name = sheet_name
        try:
            self.df_raw = ExcelParser.load_full_sheet(self.current_file_path, sheet_name)
            self.parsed_invoices, self.validation_errors = ExcelParser.parse_sales_sheet(self.df_raw)
            self._populate_table()
            self._update_stats()
        except Exception as e:
            messagebox.showerror("Loading Error", f"Failed to parse sheet: {str(e)}")

    def _populate_table(self):
        rows = []
        for inv in self.parsed_invoices:
            vch_no = inv.get("vch_no", "")
            date_str = str(inv.get("date", "")).split(" ")[0]
            party = inv.get("party_name", "")
            items = inv.get("items", [])
            taxes = inv.get("taxes", [])
            l2_disc = inv.get("discount_amount", 0.0)
            round_off = inv.get("round_off", 0.0)
            total = inv.get("party_amount", 0.0)

            # Extract CGST, SGST, IGST totals
            cgst_sum = sum(t["amount"] for t in taxes if "CGST" in t["ledger"].upper())
            sgst_sum = sum(t["amount"] for t in taxes if "SGST" in t["ledger"].upper())
            igst_sum = sum(t["amount"] for t in taxes if "IGST" in t["ledger"].upper())

            if items:
                for i, itm in enumerate(items):
                    rows.append([
                        vch_no if i == 0 else "",
                        date_str if i == 0 else "",
                        party if i == 0 else "",
                        itm.get("stock_item", ""),
                        f"{itm.get('qty', 0):.2f}",
                        itm.get("unit", ""),
                        f"{itm.get('rate', 0):.2f}",
                        f"{itm.get('discount_pct', 0):.2f}%" if itm.get('discount_pct', 0) > 0 else "-",
                        f"{itm.get('amount', 0):.2f}",
                        f"{cgst_sum:.2f}" if i == 0 and cgst_sum > 0 else "",
                        f"{sgst_sum:.2f}" if i == 0 and sgst_sum > 0 else "",
                        f"{igst_sum:.2f}" if i == 0 and igst_sum > 0 else "",
                        f"{l2_disc:.2f}" if i == 0 and l2_disc > 0 else "",
                        f"{round_off:.2f}" if i == 0 and abs(round_off) > 0 else "",
                        f"{total:.2f}" if i == 0 else ""
                    ])
            else:
                # Accounting invoice without items
                rows.append([
                    vch_no, date_str, party, "(Accounting Invoice)", "-", "-", "-", "-",
                    f"{inv.get('taxable_amount', 0):.2f}",
                    f"{cgst_sum:.2f}" if cgst_sum > 0 else "",
                    f"{sgst_sum:.2f}" if sgst_sum > 0 else "",
                    f"{igst_sum:.2f}" if igst_sum > 0 else "",
                    f"{l2_disc:.2f}" if l2_disc > 0 else "",
                    f"{round_off:.2f}" if abs(round_off) > 0 else "",
                    f"{total:.2f}"
                ])

        self.table.set_data(rows)

    def _update_stats(self):
        inv_count = len(self.parsed_invoices)
        total_items = sum(len(inv.get("items", [])) for inv in self.parsed_invoices)
        total_taxable = sum(inv.get("taxable_amount", 0.0) for inv in self.parsed_invoices)
        total_tax = sum(sum(t["amount"] for t in inv.get("taxes", [])) for inv in self.parsed_invoices)
        grand_total = sum(inv.get("party_amount", 0.0) for inv in self.parsed_invoices)

        err_badge = f" | ⚠️ {len(self.validation_errors)} Issues Detected" if self.validation_errors else " | ✅ Validated"

        self.lbl_stats.configure(
            text=(
                f"Invoices: {inv_count} | Items: {total_items} | "
                f"Taxable: ₹{total_taxable:,.2f} | Tax: ₹{total_tax:,.2f} | "
                f"Net Value: ₹{grand_total:,.2f}{err_badge}"
            )
        )

    def validate_data(self):
        if not self.parsed_invoices:
            messagebox.showwarning("Validation", "No invoices loaded to validate.")
            return

        report = []
        report.append(f"Validation Report for {len(self.parsed_invoices)} Invoices:")
        report.append("-" * 50)

        if not self.validation_errors:
            report.append("✅ All invoices passed schema checks!")
            report.append("• Voucher numbers and dates are properly formatted.")
            report.append("• Multi-line item quantities, rates, and totals balance.")
            report.append("• Tax calculations (CGST/SGST/IGST) verified.")
        else:
            report.append(f"Found {len(self.validation_errors)} issues:")
            for err in self.validation_errors:
                report.append(f"• {err}")

        messagebox.showinfo("Validation Result", "\n".join(report))

    def get_generated_xml(self) -> str:
        if not self.parsed_invoices:
            return ""
        vch_type = self.opt_vch_type.get()
        connector = self.get_tally_connector()
        # Ping active company
        company = ""
        conn_info = connector.test_connection(timeout=1)
        if conn_info.get("connected"):
            company = conn_info.get("active_company", "")

        msgs = [TallyXMLBuilder.build_sales_or_dn_voucher(inv, vch_type) for inv in self.parsed_invoices]
        return wrap_envelope("\n".join(msgs), report_name="Vouchers", company_name=company)

    def generate_xml(self):
        if not self.parsed_invoices:
            messagebox.showwarning("Warning", "Please load an Excel sheet with invoices first.")
            return
        xml = self.get_generated_xml()
        vch_type = self.opt_vch_type.get().replace(" ", "_")
        XMLPreviewModal(self, xml, title=f"Tally {self.opt_vch_type.get()} XML Preview", file_name_hint=f"{vch_type}_tally_import.xml")

    def push_to_tally(self):
        if not self.parsed_invoices:
            messagebox.showwarning("Warning", "Please load an Excel sheet with invoices first.")
            return
        xml = self.get_generated_xml()
        connector = self.get_tally_connector()

        self.btn_sync.configure(text="⏳ Syncing...", state="disabled")
        self.update_idletasks()

        res = connector.send_xml(xml)
        self.btn_sync.configure(text="🚀 Push to Tally (Direct Sync)", state="normal")
        TallySyncResponseModal(self, res)

    def download_sample_template(self):
        src = ensure_sample_file("sales")
        dst = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="sample_sales_template.xlsx"
        )
        if dst:
            try:
                shutil.copyfile(src, dst)
                messagebox.showinfo("Saved", f"Sample Sales template saved to:\n{dst}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save template: {str(e)}")
