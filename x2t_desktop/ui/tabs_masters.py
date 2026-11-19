"""
Master Generators Tab (Pre-requisite for Tally Import)
Extracts unique Ledgers and Stock Items from any loaded sheet or transaction files.
Generates compliant Ledger Masters and Stock Item Masters XML, with Direct Tally Sync.
"""

from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd
from typing import List, Dict, Any, Callable

from core.excel_parser import ExcelParser
from core.tally_xml_builder import TallyXMLBuilder
from ui.common import ModernTable, XMLPreviewModal, TallySyncResponseModal


class MastersTab(ctk.CTkFrame):
    def __init__(self, master, get_tally_connector_func, get_all_transactions_func: Callable[[], Dict[str, Any]], **kwargs):
        super().__init__(master, **kwargs)
        self.get_tally_connector = get_tally_connector_func
        self.get_all_transactions = get_all_transactions_func

        self.ledgers: List[Dict[str, Any]] = []
        self.stock_items: List[Dict[str, Any]] = []

        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_top_controls()
        self._build_subtabs()
        self._build_action_bar()

    def _build_top_controls(self):
        ctrl_frame = ctk.CTkFrame(self)
        ctrl_frame.grid(row=0, column=0, padx=15, pady=(15, 5), sticky="ew")

        lbl_desc = ctk.CTkLabel(
            ctrl_frame,
            text="Master Generator: Extract prerequisite Ledgers and Stock Items before importing vouchers into Tally.",
            font=("Segoe UI", 12)
        )
        lbl_desc.pack(side="left", padx=15, pady=10)

        btn_scan_all = ctk.CTkButton(
            ctrl_frame,
            text="🔄 Scan Current Open Tabs",
            command=self.extract_from_active_tabs,
            width=190,
            fg_color="#0284c7"
        )
        btn_scan_all.pack(side="right", padx=(10, 15), pady=10)

        btn_browse_file = ctk.CTkButton(
            ctrl_frame,
            text="📁 Scan Any Excel File...",
            command=self.browse_and_extract_file,
            width=180,
            fg_color="#334155"
        )
        btn_browse_file.pack(side="right", padx=5, pady=10)

    def _build_subtabs(self):
        self.subtabs = ctk.CTkTabview(self)
        self.subtabs.grid(row=2, column=0, padx=15, pady=5, sticky="nsew")

        self.tab_ledgers = self.subtabs.add("Ledger Masters (0)")
        self.tab_items = self.subtabs.add("Stock Item Masters (0)")

        # Setup Ledgers Table
        self.tab_ledgers.grid_rowconfigure(0, weight=1)
        self.tab_ledgers.grid_columnconfigure(0, weight=1)
        self.table_ledgers = ModernTable(
            self.tab_ledgers,
            columns=["Ledger Name", "Parent Group", "Type", "Opening Bal", "Bill-Wise"]
        )
        self.table_ledgers.grid(row=0, column=0, sticky="nsew")
        self.table_ledgers.on_edit_callback = self._on_ledger_edited

        # Setup Stock Items Table
        self.tab_items.grid_rowconfigure(0, weight=1)
        self.tab_items.grid_columnconfigure(0, weight=1)
        self.table_items = ModernTable(
            self.tab_items,
            columns=["Stock Item Name", "Parent Group", "Base Unit", "Default Godown"]
        )
        self.table_items.grid(row=0, column=0, sticky="nsew")
        self.table_items.on_edit_callback = self._on_item_edited

    def _build_action_bar(self):
        act_frame = ctk.CTkFrame(self, fg_color="transparent")
        act_frame.grid(row=3, column=0, padx=15, pady=(5, 15), sticky="ew")

        # Action for Ledgers
        btn_xml_ledgers = ctk.CTkButton(
            act_frame,
            text="📄 Generate Ledgers XML",
            command=self.generate_ledgers_xml,
            width=180,
            fg_color="#059669"
        )
        btn_xml_ledgers.pack(side="left", padx=(0, 10))

        # Action for Stock Items
        btn_xml_items = ctk.CTkButton(
            act_frame,
            text="📄 Generate Stock XML",
            command=self.generate_items_xml,
            width=170,
            fg_color="#0d9488"
        )
        btn_xml_items.pack(side="left", padx=10)

        # Push to Tally
        btn_sync = ctk.CTkButton(
            act_frame,
            text="🚀 Push Selected Masters to Tally",
            command=self.push_masters_to_tally,
            width=230,
            fg_color="#7c3aed"
        )
        btn_sync.pack(side="left", padx=10)

        lbl_edit_hint = ctk.CTkLabel(
            act_frame,
            text="(Double-click any cell to adjust Parent Group or Unit)",
            font=("Segoe UI", 10, "italic"),
            text_color="gray"
        )
        lbl_edit_hint.pack(side="right")

    def extract_from_active_tabs(self, show_dialog: bool = True):
        data = self.get_all_transactions()
        sales_invs = data.get("sales", [])
        purch_invs = data.get("purch", [])
        bank_txs = data.get("bank", [])
        vert_vchs = data.get("vert", [])

        if not (sales_invs or purch_invs or bank_txs or vert_vchs):
            if show_dialog:
                messagebox.showinfo("Masters", "No transactions currently loaded in Sales, Purchase, Bank, or Vertical tabs.")
            return

        ledgers, items = ExcelParser.extract_masters_from_transactions(
            sales_invs, purch_invs, bank_txs, vert_vchs
        )
        self.ledgers = ledgers
        self.stock_items = items
        self._refresh_tables()
        if show_dialog:
            messagebox.showinfo(
                "Scan Complete",
                f"Extracted from active tabs:\n• {len(ledgers)} Unique Ledgers\n• {len(items)} Unique Stock Items"
            )

    def browse_and_extract_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            sheets = ExcelParser.get_sheet_names(path)
            all_sales = []
            all_purch = []
            all_bank = []
            all_vert = []

            for s in sheets:
                df = ExcelParser.load_full_sheet(path, s)
                # Try parsing sales
                sinv, _ = ExcelParser.parse_sales_sheet(df)
                if sinv:
                    all_sales.extend(sinv)
                # Try parsing purchase
                pinv, _ = ExcelParser.parse_purchase_sheet(df)
                if pinv:
                    all_purch.extend(pinv)
                # Try parsing bank
                btx, _ = ExcelParser.parse_bank_sheet(df)
                if btx:
                    all_bank.extend(btx)
                # Try parsing vertical
                vvch, _ = ExcelParser.parse_vertical_sheet(df)
                if vvch:
                    all_vert.extend(vvch)

            ledgers, items = ExcelParser.extract_masters_from_transactions(
                all_sales, all_purch, all_bank, all_vert
            )
            self.ledgers = ledgers
            self.stock_items = items
            self._refresh_tables()
            messagebox.showinfo(
                "Scan Complete",
                f"Scanned {path}:\n• {len(ledgers)} Unique Ledgers\n• {len(items)} Unique Stock Items"
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract masters: {str(e)}")

    def _refresh_tables(self):
        # Update tab titles
        self.subtabs._segmented_button._buttons_dict["Ledger Masters (0)"].configure(text=f"Ledger Masters ({len(self.ledgers)})") if "Ledger Masters (0)" in self.subtabs._segmented_button._buttons_dict else None
        
        # Populate Ledgers
        l_rows = [
            [l.get("name", ""), l.get("parent", ""), l.get("type", "General"), "0.00", "Yes"]
            for l in self.ledgers
        ]
        self.table_ledgers.set_data(l_rows)

        # Populate Stock Items
        i_rows = [
            [i.get("name", ""), i.get("parent", "Primary"), i.get("unit", "PCS"), i.get("godown", "Main Location")]
            for i in self.stock_items
        ]
        self.table_items.set_data(i_rows)

    def _on_ledger_edited(self, row_idx: int, col_name: str, new_val: str):
        if 0 <= row_idx < len(self.ledgers):
            if "Parent" in col_name:
                self.ledgers[row_idx]["parent"] = new_val.strip()
            elif "Name" in col_name:
                self.ledgers[row_idx]["name"] = new_val.strip()

    def _on_item_edited(self, row_idx: int, col_name: str, new_val: str):
        if 0 <= row_idx < len(self.stock_items):
            if "Unit" in col_name:
                self.stock_items[row_idx]["unit"] = new_val.strip()
            elif "Parent" in col_name:
                self.stock_items[row_idx]["parent"] = new_val.strip()

    def generate_ledgers_xml(self):
        if not self.ledgers:
            messagebox.showwarning("Warning", "No ledgers to generate.")
            return
        connector = self.get_tally_connector()
        conn_info = connector.test_connection(timeout=1)
        company = conn_info.get("active_company", "") if conn_info.get("connected") else ""
        xml = TallyXMLBuilder.build_ledger_master_xml(self.ledgers, company_name=company)
        XMLPreviewModal(self, xml, title="Tally Ledger Masters XML", file_name_hint="tally_ledgers_master.xml")

    def generate_items_xml(self):
        if not self.stock_items:
            messagebox.showwarning("Warning", "No stock items to generate.")
            return
        connector = self.get_tally_connector()
        conn_info = connector.test_connection(timeout=1)
        company = conn_info.get("active_company", "") if conn_info.get("connected") else ""
        xml = TallyXMLBuilder.build_stock_item_master_xml(self.stock_items, company_name=company)
        XMLPreviewModal(self, xml, title="Tally Stock Item Masters XML", file_name_hint="tally_stock_items_master.xml")

    def push_masters_to_tally(self):
        current_subtab = self.subtabs.get()
        connector = self.get_tally_connector()
        conn_info = connector.test_connection()
        company = conn_info.get("active_company", "") if conn_info.get("connected") else ""

        if "Stock" in current_subtab:
            if not self.stock_items:
                messagebox.showwarning("Warning", "No stock items available to push.")
                return
            xml = TallyXMLBuilder.build_stock_item_master_xml(self.stock_items, company_name=company)
        else:
            if not self.ledgers:
                messagebox.showwarning("Warning", "No ledgers available to push.")
                return
            xml = TallyXMLBuilder.build_ledger_master_xml(self.ledgers, company_name=company)

        res = connector.send_xml(xml)
        TallySyncResponseModal(self, res)
