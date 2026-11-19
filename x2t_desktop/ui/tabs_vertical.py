"""
Vertical Import Format (Journal / Generic Vouchers) Tab
Multi-row Dr/Cr vouchers (Payment, Receipt, Contra, Journal).
Strict Dr == Cr balance validation before export.
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


class VerticalTab(ctk.CTkFrame):
    def __init__(self, master, get_tally_connector_func, **kwargs):
        super().__init__(master, **kwargs)
        self.get_tally_connector = get_tally_connector_func

        self.current_file_path = ""
        self.current_sheet_name = ""
        self.df_raw = pd.DataFrame()
        self.parsed_vouchers = []
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

        self.entry_file = ctk.CTkEntry(ctrl_frame, placeholder_text="Select Vertical Journal Excel workbook...")
        self.entry_file.grid(row=0, column=1, padx=5, pady=10, sticky="ew")

        btn_browse = ctk.CTkButton(ctrl_frame, text="📁 Browse...", width=100, command=self.browse_file)
        btn_browse.grid(row=0, column=2, padx=5, pady=10)

        # Sheet Selection
        lbl_sheet = ctk.CTkLabel(ctrl_frame, text="Sheet:", font=("Segoe UI", 12, "bold"))
        lbl_sheet.grid(row=0, column=3, padx=(10, 5), pady=10)

        self.opt_sheet = ctk.CTkOptionMenu(ctrl_frame, values=["(No sheets)"], command=self.on_sheet_changed, width=160)
        self.opt_sheet.grid(row=0, column=4, padx=5, pady=10)

        # Sample Template Button
        btn_sample = ctk.CTkButton(
            ctrl_frame,
            text="📥 Sample Template",
            width=130,
            fg_color="#334155",
            command=self.download_sample_template
        )
        btn_sample.grid(row=0, column=5, padx=(5, 15), pady=10)

    def _build_stats_bar(self):
        self.stats_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.stats_frame.grid(row=1, column=0, padx=15, pady=5, sticky="ew")

        self.lbl_stats = ctk.CTkLabel(
            self.stats_frame,
            text="No data loaded. Select a Vertical Format Excel file above to begin.",
            font=("Segoe UI", 11)
        )
        self.lbl_stats.pack(side="left")

    def _build_table_preview(self):
        columns = [
            "Vch No", "Type", "Date", "Ledger Name",
            "Dr / Cr", "Amount", "Cost Center", "Narration", "Balance Status"
        ]
        self.table = ModernTable(self, columns=columns)
        self.table.grid(row=2, column=0, padx=15, pady=5, sticky="nsew")

    def _build_action_bar(self):
        act_frame = ctk.CTkFrame(self, fg_color="transparent")
        act_frame.grid(row=3, column=0, padx=15, pady=(5, 15), sticky="ew")

        btn_validate = ctk.CTkButton(
            act_frame,
            text="⚖️ Validate Dr == Cr",
            command=self.validate_data,
            width=160,
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
            self.parsed_vouchers, self.validation_errors = ExcelParser.parse_vertical_sheet(self.df_raw)
            self._populate_table()
            self._update_stats()
        except Exception as e:
            messagebox.showerror("Loading Error", f"Failed to parse vertical sheet: {str(e)}")

    def _populate_table(self):
        rows = []
        for vch in self.parsed_vouchers:
            vch_no = vch.get("vch_no", "")
            vch_type = vch.get("vch_type", "Journal")
            date_str = str(vch.get("date", "")).split(" ")[0]
            narration = vch.get("narration", "")
            is_balanced = vch.get("balanced", True)
            diff = vch.get("diff", 0.0)

            status_text = "✅ Balanced" if is_balanced else f"❌ Diff: ₹{diff:.2f}"

            for i, entry in enumerate(vch.get("entries", [])):
                rows.append([
                    vch_no if i == 0 else "",
                    vch_type if i == 0 else "",
                    date_str if i == 0 else "",
                    entry.get("ledger", ""),
                    entry.get("type", "Dr"),
                    f"{entry.get('amount', 0.0):.2f}",
                    entry.get("cost_center", ""),
                    narration if i == 0 else "",
                    status_text if i == 0 else ""
                ])

        self.table.set_data(rows)

    def _update_stats(self):
        vch_count = len(self.parsed_vouchers)
        total_rows = sum(len(v.get("entries", [])) for v in self.parsed_vouchers)
        total_dr = sum(v.get("total_dr", 0.0) for v in self.parsed_vouchers)
        total_cr = sum(v.get("total_cr", 0.0) for v in self.parsed_vouchers)
        imbalanced = sum(1 for v in self.parsed_vouchers if not v.get("balanced", True))

        balance_badge = f" | ⚠️ {imbalanced} Imbalanced Vouchers" if imbalanced > 0 else " | ✅ All Balanced"

        self.lbl_stats.configure(
            text=(
                f"Vouchers: {vch_count} | Lines: {total_rows} | "
                f"Total Debit: ₹{total_dr:,.2f} | Total Credit: ₹{total_cr:,.2f}{balance_badge}"
            )
        )

    def validate_data(self):
        if not self.parsed_vouchers:
            messagebox.showwarning("Validation", "No vouchers loaded to validate.")
            return

        report = [
            f"Balancing Validation Report ({len(self.parsed_vouchers)} Vouchers):",
            "-" * 50
        ]
        imbalanced = [v for v in self.parsed_vouchers if not v.get("balanced", True)]
        if not imbalanced:
            report.append("✅ 100% Balanced! Every voucher satisfies: Total Debit == Total Credit.")
        else:
            report.append(f"❌ Found {len(imbalanced)} imbalanced voucher(s):")
            for v in imbalanced:
                report.append(f"• Voucher {v['vch_no']}: Dr={v['total_dr']:.2f}, Cr={v['total_cr']:.2f} (Diff: ₹{v['diff']:.2f})")
            report.append("\nNote: Tally will reject vouchers that do not balance.")

        messagebox.showinfo("Balancing Report", "\n".join(report))

    def get_generated_xml(self) -> str:
        if not self.parsed_vouchers:
            return ""
        connector = self.get_tally_connector()
        company = ""
        conn_info = connector.test_connection(timeout=1)
        if conn_info.get("connected"):
            company = conn_info.get("active_company", "")

        msgs = [TallyXMLBuilder.build_vertical_voucher(vch) for vch in self.parsed_vouchers]
        return wrap_envelope("\n".join(msgs), report_name="Vouchers", company_name=company)

    def generate_xml(self):
        if not self.parsed_vouchers:
            messagebox.showwarning("Warning", "Please load vertical vouchers first.")
            return
        imbalanced = [v for v in self.parsed_vouchers if not v.get("balanced", True)]
        if imbalanced:
            if not messagebox.askyesno(
                "Imbalanced Warning",
                f"Warning: {len(imbalanced)} voucher(s) have unequal Debit & Credit amounts.\n"
                "Tally requires vouchers to balance to 0. Do you still wish to generate XML?"
            ):
                return

        xml = self.get_generated_xml()
        XMLPreviewModal(self, xml, title="Vertical Journal Tally XML Preview", file_name_hint="journal_vouchers_tally.xml")

    def push_to_tally(self):
        if not self.parsed_vouchers:
            messagebox.showwarning("Warning", "Please load vertical vouchers first.")
            return
        imbalanced = [v for v in self.parsed_vouchers if not v.get("balanced", True)]
        if imbalanced:
            if not messagebox.askyesno(
                "Imbalanced Warning",
                f"Warning: {len(imbalanced)} voucher(s) are imbalanced.\nTally will reject them. Proceed anyway?"
            ):
                return

        xml = self.get_generated_xml()
        connector = self.get_tally_connector()

        self.btn_sync.configure(text="⏳ Syncing...", state="disabled")
        self.update_idletasks()

        res = connector.send_xml(xml)
        self.btn_sync.configure(text="🚀 Push to Tally (Direct Sync)", state="normal")
        TallySyncResponseModal(self, res)

    def download_sample_template(self):
        src = ensure_sample_file("vertical")
        dst = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="sample_vertical_template.xlsx"
        )
        if dst:
            try:
                shutil.copyfile(src, dst)
                messagebox.showinfo("Saved", f"Sample Vertical Journal template saved to:\n{dst}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save template: {str(e)}")
