"""
Settings & Rules Manager Tab
Configures Tally HTTP connection parameters and manages bank auto-categorization rules.
"""

from tkinter import messagebox
import customtkinter as ctk

from core.tally_connector import TallyConnector
from core.rules_engine import RulesEngine
from ui.common import ModernTable


class SettingsTab(ctk.CTkFrame):
    def __init__(self, master, get_tally_connector_func, update_tally_config_func, rules_engine: RulesEngine, **kwargs):
        super().__init__(master, **kwargs)
        self.get_tally_connector = get_tally_connector_func
        self.update_tally_config = update_tally_config_func
        self.rules_engine = rules_engine

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_tally_settings()
        self._build_rules_manager()

    def _build_tally_settings(self):
        tally_box = ctk.CTkFrame(self)
        tally_box.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="ew")

        lbl_tally = ctk.CTkLabel(tally_box, text="Tally Prime / ERP 9 Connectivity Settings", font=("Segoe UI", 13, "bold"))
        lbl_tally.grid(row=0, column=0, columnspan=7, padx=15, pady=(10, 5), sticky="w")

        # Host
        lbl_host = ctk.CTkLabel(tally_box, text="Host / IP:")
        lbl_host.grid(row=1, column=0, padx=(15, 5), pady=10)

        self.entry_host = ctk.CTkEntry(tally_box, width=120)
        self.entry_host.insert(0, "localhost")
        self.entry_host.grid(row=1, column=1, padx=5, pady=10)

        # Port
        lbl_port = ctk.CTkLabel(tally_box, text="Port:")
        lbl_port.grid(row=1, column=2, padx=(10, 5), pady=10)

        self.entry_port = ctk.CTkEntry(tally_box, width=70)
        self.entry_port.insert(0, "9000")
        self.entry_port.grid(row=1, column=3, padx=5, pady=10)

        # Timeout
        lbl_timeout = ctk.CTkLabel(tally_box, text="Timeout (s):")
        lbl_timeout.grid(row=1, column=4, padx=(10, 5), pady=10)

        self.entry_timeout = ctk.CTkEntry(tally_box, width=60)
        self.entry_timeout.insert(0, "15")
        self.entry_timeout.grid(row=1, column=5, padx=5, pady=10)

        # Apply / Save & Test
        btn_apply = ctk.CTkButton(tally_box, text="Save Settings", width=110, command=self.apply_tally_settings)
        btn_apply.grid(row=1, column=6, padx=(10, 5), pady=10)

        btn_test = ctk.CTkButton(
            tally_box,
            text="🔌 Test Connection",
            width=130,
            fg_color="#0284c7",
            command=self.test_tally_connection
        )
        btn_test.grid(row=1, column=7, padx=(5, 15), pady=10)

        self.lbl_tally_status = ctk.CTkLabel(
            tally_box,
            text="Status: Untested. Click 'Test Connection' to check local Tally server.",
            font=("Segoe UI", 11)
        )
        self.lbl_tally_status.grid(row=2, column=0, columnspan=8, padx=15, pady=(0, 10), sticky="w")

    def _build_rules_manager(self):
        rules_box = ctk.CTkFrame(self)
        rules_box.grid(row=1, column=0, padx=15, pady=(5, 15), sticky="nsew")
        rules_box.grid_rowconfigure(2, weight=1)
        rules_box.grid_columnconfigure(0, weight=1)

        header_frame = ctk.CTkFrame(rules_box, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=15, pady=(10, 5), sticky="ew")

        lbl_rules = ctk.CTkLabel(
            header_frame,
            text="Bank Statement Narration Auto-Categorization Rules",
            font=("Segoe UI", 13, "bold")
        )
        lbl_rules.pack(side="left")

        btn_reset = ctk.CTkButton(
            header_frame,
            text="↺ Reset to Defaults",
            width=130,
            fg_color="#334155",
            command=self.reset_rules
        )
        btn_reset.pack(side="right")

        # Form to add rule
        add_frame = ctk.CTkFrame(rules_box)
        add_frame.grid(row=1, column=0, padx=15, pady=5, sticky="ew")

        ctk.CTkLabel(add_frame, text="Keyword:").pack(side="left", padx=(10, 5), pady=8)
        self.entry_kw = ctk.CTkEntry(add_frame, placeholder_text="e.g. SWIGGY", width=120)
        self.entry_kw.pack(side="left", padx=5, pady=8)

        ctk.CTkLabel(add_frame, text="Target Ledger:").pack(side="left", padx=(10, 5), pady=8)
        self.entry_rule_ledger = ctk.CTkEntry(add_frame, placeholder_text="e.g. Staff Welfare", width=160)
        self.entry_rule_ledger.pack(side="left", padx=5, pady=8)

        ctk.CTkLabel(add_frame, text="Voucher Type:").pack(side="left", padx=(10, 5), pady=8)
        self.opt_rule_vtype = ctk.CTkOptionMenu(add_frame, values=["Payment", "Receipt", "Contra"], width=110)
        self.opt_rule_vtype.pack(side="left", padx=5, pady=8)

        btn_add = ctk.CTkButton(
            add_frame,
            text="➕ Add / Update Rule",
            width=140,
            fg_color="#059669",
            command=self.add_rule
        )
        btn_add.pack(side="left", padx=(15, 10), pady=8)

        btn_del = ctk.CTkButton(
            add_frame,
            text="🗑️ Delete Selected",
            width=120,
            fg_color="#dc2626",
            command=self.delete_selected_rule
        )
        btn_del.pack(side="right", padx=10, pady=8)

        # Rules Table
        self.table_rules = ModernTable(
            rules_box,
            columns=["#", "Keyword", "Target Ledger", "Voucher Type", "Match Type"]
        )
        self.table_rules.grid(row=2, column=0, padx=15, pady=10, sticky="nsew")

        self._refresh_rules_table()

    def apply_tally_settings(self):
        host = self.entry_host.get().strip() or "localhost"
        try:
            port = int(self.entry_port.get().strip() or "9000")
            timeout = int(self.entry_timeout.get().strip() or "15")
        except ValueError:
            messagebox.showerror("Error", "Port and Timeout must be integer numbers.")
            return

        self.update_tally_config(host, port, timeout)
        messagebox.showinfo("Saved", f"Tally connector updated to http://{host}:{port}")

    def test_tally_connection(self):
        self.apply_tally_settings()
        connector = self.get_tally_connector()
        self.lbl_tally_status.configure(text="Pinging Tally...", text_color="gray")
        self.update_idletasks()

        res = connector.test_connection()
        if res.get("connected"):
            comp = res.get("active_company", "Connected")
            self.lbl_tally_status.configure(
                text=f"🟢 Connected to Tally Prime | Active Company: {comp}",
                text_color="#16a34a"
            )
            messagebox.showinfo("Connection Successful", f"Connected to Tally!\nActive Company: {comp}")
        else:
            msg = res.get("message", "Connection failed")
            self.lbl_tally_status.configure(
                text=f"🔴 {msg.splitlines()[0]}",
                text_color="#dc2626"
            )
            messagebox.showerror("Connection Failed", msg)

    def _refresh_rules_table(self):
        rows = []
        for i, r in enumerate(self.rules_engine.rules, start=1):
            rows.append([
                str(i),
                r.get("keyword", ""),
                r.get("ledger", ""),
                r.get("vch_type", "Payment"),
                r.get("match_type", "contains")
            ])
        self.table_rules.set_data(rows)

    def add_rule(self):
        kw = self.entry_kw.get().strip()
        ledger = self.entry_rule_ledger.get().strip()
        vtype = self.opt_rule_vtype.get()
        if not kw or not ledger:
            messagebox.showwarning("Warning", "Please enter both Keyword and Target Ledger.")
            return

        self.rules_engine.add_rule(kw, ledger, vch_type=vtype)
        self.entry_kw.delete(0, "end")
        self.entry_rule_ledger.delete(0, "end")
        self._refresh_rules_table()
        messagebox.showinfo("Rule Added", f"Rule for '{kw}' -> '{ledger}' saved.")

    def delete_selected_rule(self):
        selected = self.table_rules.tree.selection()
        if not selected:
            messagebox.showwarning("Select Rule", "Please select a rule from the table to delete.")
            return

        idx = self.table_rules.tree.index(selected[0])
        self.rules_engine.delete_rule(idx)
        self._refresh_rules_table()

    def reset_rules(self):
        if messagebox.askyesno("Confirm Reset", "Reset all auto-categorization rules back to system defaults?"):
            self.rules_engine.reset_to_defaults()
            self._refresh_rules_table()
            messagebox.showinfo("Reset Complete", "Rules restored to defaults.")
