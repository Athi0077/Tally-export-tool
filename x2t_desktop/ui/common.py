"""
Common UI components, table widgets, dialogs, and styling helpers for CustomTkinter.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
from typing import List, Dict, Any, Optional, Callable


class ModernTable(ctk.CTkFrame):
    """
    A high-performance modern data table wrapping ttk.Treeview
    styled to match CustomTkinter dark/light aesthetics.
    """
    def __init__(self, master, columns: List[str], height: int = 300, **kwargs):
        super().__init__(master, **kwargs)
        self.columns = columns
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Style Treeview
        self.style = ttk.Style()
        self._apply_style()

        # Treeview
        self.tree = ttk.Treeview(
            self,
            columns=columns,
            show="headings",
            selectmode="extended",
            height=height
        )
        self.tree.grid(row=0, column=0, sticky="nsew")

        # Scrollbars
        self.vsb = ctk.CTkScrollbar(self, orientation="vertical", command=self.tree.yview)
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb = ctk.CTkScrollbar(self, orientation="horizontal", command=self.tree.xview)
        self.hsb.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        # Headings & initial column setup
        for col in columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by(c, False))
            self.tree.column(col, width=130, minwidth=80, anchor="w")

        self.tree.bind("<Double-1>", self._on_double_click)
        self.on_edit_callback: Optional[Callable[[int, str, str], None]] = None

    def _apply_style(self):
        mode = ctk.get_appearance_mode().lower()
        if mode == "dark":
            bg = "#1e1e24"
            fg = "#e6e6e6"
            header_bg = "#2b2b36"
            header_fg = "#ffffff"
            selected_bg = "#1f538d"
            row_alt = "#252530"
        else:
            bg = "#ffffff"
            fg = "#1a1a1a"
            header_bg = "#e0e0e0"
            header_fg = "#000000"
            selected_bg = "#3a7ebf"
            row_alt = "#f4f4f6"

        self.style.theme_use("clam")
        self.style.configure(
            "Treeview",
            background=bg,
            foreground=fg,
            fieldbackground=bg,
            rowheight=28,
            font=("Segoe UI", 10)
        )
        self.style.configure(
            "Treeview.Heading",
            background=header_bg,
            foreground=header_fg,
            relief="flat",
            font=("Segoe UI", 10, "bold")
        )
        self.style.map(
            "Treeview",
            background=[("selected", selected_bg)],
            foreground=[("selected", "#ffffff")]
        )
        self.style.map(
            "Treeview.Heading",
            background=[("active", header_bg)]
        )

    def set_data(self, rows: List[List[Any]]):
        """Clears and re-populates rows in table."""
        self.clear()
        for i, row in enumerate(rows):
            tag = "even" if i % 2 == 0 else "odd"
            clean_row = [str(val) if val is not None else "" for val in row]
            self.tree.insert("", "end", values=clean_row, tags=(tag,))
        self._update_tag_colors()

    def update_columns(self, new_columns: List[str]):
        """Dynamically reconfigures columns."""
        self.columns = new_columns
        self.tree["columns"] = new_columns
        for col in new_columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by(c, False))
            self.tree.column(col, width=130, minwidth=80, anchor="w")

    def clear(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _update_tag_colors(self):
        mode = ctk.get_appearance_mode().lower()
        if mode == "dark":
            self.tree.tag_configure("even", background="#1e1e24")
            self.tree.tag_configure("odd", background="#262632")
        else:
            self.tree.tag_configure("even", background="#ffffff")
            self.tree.tag_configure("odd", background="#f5f6f8")

    def sort_by(self, col: str, reverse: bool):
        """Sorts table contents when column header is clicked."""
        data = [(self.tree.set(child, col), child) for child in self.tree.get_children("")]
        try:
            # Try numeric sort
            data.sort(key=lambda t: float(str(t[0]).replace(",", "").strip()), reverse=reverse)
        except Exception:
            # Fallback to string sort
            data.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)

        for index, (_, child) in enumerate(data):
            self.tree.move(child, "", index)
        self.tree.heading(col, command=lambda: self.sort_by(col, not reverse))

    def _on_double_click(self, event):
        """Allows double click inline cell editing if callback set."""
        if not self.on_edit_callback:
            return
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        if not item or not column:
            return
        col_idx = int(column.replace("#", "")) - 1
        col_name = self.columns[col_idx]
        current_val = self.tree.item(item, "values")[col_idx]
        row_idx = self.tree.index(item)
        
        # Simple input dialog
        dialog = ctk.CTkInputDialog(text=f"Edit {col_name}:", title="Edit Value")
        new_val = dialog.get_input()
        if new_val is not None:
            vals = list(self.tree.item(item, "values"))
            vals[col_idx] = new_val
            self.tree.item(item, values=vals)
            self.on_edit_callback(row_idx, col_name, new_val)


class XMLPreviewModal(ctk.CTkToplevel):
    """
    Modal window displaying generated Tally XML with syntax highlighting,
    character statistics, and copy to clipboard / save options.
    """
    def __init__(self, master, xml_content: str, title: str = "Tally XML Preview", file_name_hint: str = "tally_export.xml"):
        super().__init__(master)
        self.title(title)
        self.geometry("850x650")
        self.xml_content = xml_content
        self.file_name_hint = file_name_hint

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Top Bar
        top_bar = ctk.CTkFrame(self, fg_color="transparent")
        top_bar.grid(row=0, column=0, padx=20, pady=(15, 5), sticky="ew")

        lines_count = len(xml_content.splitlines())
        size_kb = len(xml_content.encode("utf-8")) / 1024.0
        info_lbl = ctk.CTkLabel(
            top_bar,
            text=f"Tally Standard XML Payload | {lines_count} lines | {size_kb:.1f} KB",
            font=("Segoe UI", 12, "bold")
        )
        info_lbl.pack(side="left")

        # Action Buttons
        save_btn = ctk.CTkButton(top_bar, text="💾 Save XML File...", command=self.save_file, width=140)
        save_btn.pack(side="right", padx=(10, 0))

        copy_btn = ctk.CTkButton(top_bar, text="📋 Copy to Clipboard", command=self.copy_clipboard, width=140, fg_color="#334155")
        copy_btn.pack(side="right")

        # Text Display
        self.textbox = ctk.CTkTextbox(self, font=("Consolas", 11), wrap="none")
        self.textbox.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        self.textbox.insert("1.0", xml_content)

        # Bottom Bar
        btm_bar = ctk.CTkFrame(self, fg_color="transparent")
        btm_bar.grid(row=2, column=0, padx=20, pady=(5, 15), sticky="ew")

        close_btn = ctk.CTkButton(btm_bar, text="Close", command=self.destroy, width=100)
        close_btn.pack(side="right")

    def copy_clipboard(self):
        self.clipboard_clear()
        self.clipboard_append(self.xml_content)
        messagebox.showinfo("Copied", "XML copied to system clipboard successfully!", parent=self)

    def save_file(self):
        path = filedialog.asksaveasfilename(
            parent=self,
            defaultextension=".xml",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
            initialfile=self.file_name_hint
        )
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(self.xml_content)
                messagebox.showinfo("Saved", f"XML successfully saved to:\n{path}", parent=self)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {str(e)}", parent=self)


class TallySyncResponseModal(ctk.CTkToplevel):
    """
    Modal displaying results of Direct HTTP Sync to Tally Prime / ERP 9.
    """
    def __init__(self, master, sync_result: Dict[str, Any]):
        super().__init__(master)
        self.title("Tally Prime Direct Sync Report")
        self.geometry("620x520")

        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        is_success = sync_result.get("success", False)
        created = sync_result.get("created", 0)
        altered = sync_result.get("altered", 0)
        errors = sync_result.get("errors", 0)
        ignored = sync_result.get("ignored", 0)
        messages = sync_result.get("messages", [])

        # Header Status Banner
        status_color = "#15803d" if (is_success and errors == 0) else "#b91c1c"
        status_text = "SYNC COMPLETED SUCCESSFULLY" if (is_success and errors == 0) else "SYNC COMPLETED WITH ISSUES"

        header_frame = ctk.CTkFrame(self, fg_color=status_color, corner_radius=8)
        header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")

        h_lbl = ctk.CTkLabel(
            header_frame,
            text=status_text,
            font=("Segoe UI", 16, "bold"),
            text_color="#ffffff"
        )
        h_lbl.pack(pady=12)

        # Metrics Card Frame
        metrics_frame = ctk.CTkFrame(self)
        metrics_frame.grid(row=1, column=0, padx=20, pady=5, sticky="ew")
        metrics_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self._create_stat_card(metrics_frame, 0, "Created", str(created), "#16a34a")
        self._create_stat_card(metrics_frame, 1, "Altered", str(altered), "#2563eb")
        self._create_stat_card(metrics_frame, 2, "Ignored", str(ignored), "#ca8a04")
        self._create_stat_card(metrics_frame, 3, "Errors", str(errors), "#dc2626")

        # Details Textbox
        details_lbl = ctk.CTkLabel(self, text="Tally Response Details & Diagnostic Messages:", font=("Segoe UI", 11, "bold"))
        details_lbl.grid(row=2, column=0, padx=25, pady=(10, 2), sticky="w")

        tb = ctk.CTkTextbox(self, font=("Consolas", 11), wrap="word")
        tb.grid(row=3, column=0, padx=20, pady=(0, 10), sticky="nsew")

        if messages:
            tb.insert("1.0", "\n".join(messages) + "\n\n")
        else:
            tb.insert("1.0", "All vouchers / masters imported into Tally with zero exceptions.\n\n")

        raw_xml = sync_result.get("raw", "")
        if raw_xml:
            tb.insert("end", "--- RAW TALLY XML RESPONSE ---\n" + raw_xml)

        close_btn = ctk.CTkButton(self, text="Close", command=self.destroy, width=120)
        close_btn.grid(row=4, column=0, padx=20, pady=(5, 15), sticky="e")

    def _create_stat_card(self, parent, col, title, value, color):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.grid(row=0, column=col, padx=10, pady=10)
        v = ctk.CTkLabel(f, text=value, font=("Segoe UI", 22, "bold"), text_color=color)
        v.pack()
        t = ctk.CTkLabel(f, text=title, font=("Segoe UI", 11))
        t.pack()
