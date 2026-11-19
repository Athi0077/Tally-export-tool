"""
X2T - Excel to Tally Converter
Modern Desktop Application (CustomTkinter)
Standalone, production-ready replacement for Excel VBA Macro tool (X2T Beta).
"""

import sys
import os
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

# Ensure current package can be imported
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from core.tally_connector import TallyConnector
from core.rules_engine import RulesEngine
from ui.tabs_sales import SalesTab
from ui.tabs_purchase import PurchaseTab
from ui.tabs_bank import BankTab
from ui.tabs_vertical import VerticalTab
from ui.tabs_masters import MastersTab
from ui.tabs_settings import SettingsTab


class X2TApplication(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Appearance & Theme Setup
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        self.title("X2T - Excel to Tally Converter v2.0")
        self.geometry("1280x820")
        self.minsize(1024, 680)

        # Core Engines
        self.tally_host = "localhost"
        self.tally_port = 9000
        self.tally_timeout = 15
        self.tally_connector = TallyConnector(host=self.tally_host, port=self.tally_port, timeout=self.tally_timeout)
        self.rules_engine = RulesEngine()

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_header()
        self._build_tabview()
        self._build_footer()

    def get_tally_connector(self) -> TallyConnector:
        return self.tally_connector

    def update_tally_config(self, host: str, port: int, timeout: int):
        self.tally_host = host
        self.tally_port = port
        self.tally_timeout = timeout
        self.tally_connector = TallyConnector(host=host, port=port, timeout=timeout)
        self.check_tally_status_silent()

    def get_all_transactions(self):
        """Aggregates parsed data across all tabs for Master Generation."""
        return {
            "sales": getattr(self.tab_sales, "parsed_invoices", []),
            "purch": getattr(self.tab_purch, "parsed_invoices", []),
            "bank": getattr(self.tab_bank, "transactions", []),
            "vert": getattr(self.tab_vertical, "parsed_vouchers", [])
        }

    def _build_header(self):
        header_frame = ctk.CTkFrame(self, height=65, corner_radius=0, fg_color=("#e2e8f0", "#18181b"))
        header_frame.grid(row=0, column=0, sticky="ew")
        header_frame.grid_columnconfigure(1, weight=1)

        # App Brand & Title
        title_box = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_box.grid(row=0, column=0, padx=20, pady=10, sticky="w")

        lbl_logo = ctk.CTkLabel(
            title_box,
            text="⚡ X2T",
            font=("Segoe UI", 20, "bold"),
            text_color=("#2563eb", "#38bdf8")
        )
        lbl_logo.pack(side="left", padx=(0, 10))

        lbl_title = ctk.CTkLabel(
            title_box,
            text="Excel to Tally Converter",
            font=("Segoe UI", 16, "bold")
        )
        lbl_title.pack(side="left")

        lbl_sub = ctk.CTkLabel(
            title_box,
            text="| v2.0 Enterprise Release",
            font=("Segoe UI", 11),
            text_color="gray"
        )
        lbl_sub.pack(side="left", padx=8)

        # Right Header Widgets
        right_box = ctk.CTkFrame(header_frame, fg_color="transparent")
        right_box.grid(row=0, column=2, padx=20, pady=10, sticky="e")

        # Tally Status Pill Button
        self.btn_tally_pill = ctk.CTkButton(
            right_box,
            text="🔌 Tally: Test Link",
            font=("Segoe UI", 11, "bold"),
            height=30,
            width=170,
            fg_color="#334155",
            hover_color="#475569",
            command=self.check_tally_status_prompt
        )
        self.btn_tally_pill.pack(side="left", padx=(0, 15))

        # Appearance Mode Menu
        lbl_theme = ctk.CTkLabel(right_box, text="Theme:", font=("Segoe UI", 11))
        lbl_theme.pack(side="left", padx=(0, 5))

        self.theme_menu = ctk.CTkOptionMenu(
            right_box,
            values=["Dark", "Light", "System"],
            command=self.change_theme,
            width=100,
            height=30
        )
        self.theme_menu.set("Dark")
        self.theme_menu.pack(side="left")

    def _build_tabview(self):
        self.tabview = ctk.CTkTabview(self, corner_radius=10)
        self.tabview.grid(row=1, column=0, padx=15, pady=(5, 10), sticky="nsew")

        # Tabs
        tab_names = [
            "🛒 Sales / Debit Note",
            "📦 Purchase / Credit Note",
            "🏦 Bank Statement",
            "📑 Vertical / Journal",
            "🗂️ Master Generators",
            "⚙️ Settings & Rules"
        ]

        t_sales = self.tabview.add(tab_names[0])
        t_purch = self.tabview.add(tab_names[1])
        t_bank = self.tabview.add(tab_names[2])
        t_vert = self.tabview.add(tab_names[3])
        t_mast = self.tabview.add(tab_names[4])
        t_sett = self.tabview.add(tab_names[5])

        # 1. Sales Tab
        t_sales.grid_rowconfigure(0, weight=1)
        t_sales.grid_columnconfigure(0, weight=1)
        self.tab_sales = SalesTab(t_sales, self.get_tally_connector)
        self.tab_sales.grid(row=0, column=0, sticky="nsew")

        # 2. Purchase Tab
        t_purch.grid_rowconfigure(0, weight=1)
        t_purch.grid_columnconfigure(0, weight=1)
        self.tab_purch = PurchaseTab(t_purch, self.get_tally_connector)
        self.tab_purch.grid(row=0, column=0, sticky="nsew")

        # 3. Bank Tab
        t_bank.grid_rowconfigure(0, weight=1)
        t_bank.grid_columnconfigure(0, weight=1)
        self.tab_bank = BankTab(t_bank, self.get_tally_connector, self.rules_engine)
        self.tab_bank.grid(row=0, column=0, sticky="nsew")

        # 4. Vertical Tab
        t_vert.grid_rowconfigure(0, weight=1)
        t_vert.grid_columnconfigure(0, weight=1)
        self.tab_vertical = VerticalTab(t_vert, self.get_tally_connector)
        self.tab_vertical.grid(row=0, column=0, sticky="nsew")

        # 5. Masters Tab
        t_mast.grid_rowconfigure(0, weight=1)
        t_mast.grid_columnconfigure(0, weight=1)
        self.tab_masters = MastersTab(t_mast, self.get_tally_connector, self.get_all_transactions)
        self.tab_masters.grid(row=0, column=0, sticky="nsew")

        # 6. Settings Tab
        t_sett.grid_rowconfigure(0, weight=1)
        t_sett.grid_columnconfigure(0, weight=1)
        self.tab_settings = SettingsTab(
            t_sett,
            self.get_tally_connector,
            self.update_tally_config,
            self.rules_engine
        )
        self.tab_settings.grid(row=0, column=0, sticky="nsew")

    def _build_footer(self):
        footer = ctk.CTkFrame(self, height=30, corner_radius=0, fg_color="transparent")
        footer.grid(row=2, column=0, padx=20, pady=(0, 5), sticky="ew")

        lbl_foot = ctk.CTkLabel(
            footer,
            text="Tally ERP 9 / Tally Prime Standard XML Integration Engine | XML Request: Import Data",
            font=("Segoe UI", 10),
            text_color="gray"
        )
        lbl_foot.pack(side="left")

        lbl_port_status = ctk.CTkLabel(
            footer,
            text=f"Target Endpoint: http://{self.tally_host}:{self.tally_port}",
            font=("Segoe UI", 10),
            text_color="gray"
        )
        lbl_port_status.pack(side="right")

    def change_theme(self, new_theme: str):
        ctk.set_appearance_mode(new_theme)
        # Refresh tables to adapt styling
        for tab in [self.tab_sales, self.tab_purch, self.tab_bank, self.tab_vertical]:
            if hasattr(tab, "table") and hasattr(tab.table, "_apply_style"):
                tab.table._apply_style()
                tab.table._update_tag_colors()
        if hasattr(self.tab_masters, "table_ledgers"):
            self.tab_masters.table_ledgers._apply_style()
            self.tab_masters.table_ledgers._update_tag_colors()
            self.tab_masters.table_items._apply_style()
            self.tab_masters.table_items._update_tag_colors()

    def check_tally_status_prompt(self):
        self.btn_tally_pill.configure(text="⏳ Pinging...", fg_color="#334155")
        self.update_idletasks()
        res = self.tally_connector.test_connection()
        if res.get("connected"):
            comp = res.get("active_company", "Connected")
            short_comp = (comp[:18] + "..") if len(comp) > 20 else comp
            self.btn_tally_pill.configure(
                text=f"🟢 {short_comp}",
                fg_color="#15803d",
                hover_color="#166534"
            )
            messagebox.showinfo("Tally Connected", f"Successfully connected to Tally at {self.tally_connector.base_url}!\nActive Company: {comp}")
        else:
            self.btn_tally_pill.configure(
                text="🔴 Tally: Offline",
                fg_color="#b91c1c",
                hover_color="#991b1b"
            )
            messagebox.showwarning("Tally Connection", res.get("message", "Could not reach Tally."))

    def check_tally_status_silent(self):
        try:
            res = self.tally_connector.test_connection()
            if res.get("connected"):
                comp = res.get("active_company", "Connected")
                short_comp = (comp[:18] + "..") if len(comp) > 20 else comp
                self.btn_tally_pill.configure(
                    text=f"🟢 {short_comp}",
                    fg_color="#15803d",
                    hover_color="#166534"
                )
            else:
                self.btn_tally_pill.configure(
                    text="🔴 Tally: Offline",
                    fg_color="#b91c1c",
                    hover_color="#991b1b"
                )
        except Exception:
            pass


def main():
    try:
        app = X2TApplication()
        app.mainloop()
    except Exception as e:
        import traceback
        traceback.print_exc()
        messagebox.showerror("Fatal Error", f"Application encountered an unexpected error:\n{str(e)}")


if __name__ == "__main__":
    main()
