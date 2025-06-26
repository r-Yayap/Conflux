# Conflux V1
import os
import tkinter as tk
import sys
import customtkinter as ctk
import pandas as pd
import traceback
from tkinter import filedialog, messagebox
from datetime import datetime
from tkinterdnd2 import TkinterDnD, DND_ALL

from core.merger import MergerFacade
from core.validators import CheckConfig

def resource_path(relative_path):
    """Get the absolute path to a resource (used for PyInstaller compatibility)"""
    try:
        base_path = sys._MEIPASS  # PyInstaller's temp folder
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

theme_path = resource_path("style/conflux-dark-red.json")
ctk.set_default_color_theme(theme_path)
ctk.set_appearance_mode("dark")

# Helper to auto-select a header based on keywords.
def auto_select_header(headers, keywords):
    for header in headers:
        lower_header = header.lower()
        for kw in keywords:
            if kw in lower_header:
                return header
    return headers[0] if headers else ""

class CTkDnD(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class MergerGUI:
    """Handles the GUI for the Excel merger with drag-and-drop file selection and theme toggle."""

    def __init__(self, master=None):

        #set_custom_theme("dark")  # or "light" if you prefer the light theme

        ctk.set_appearance_mode("dark")  # Set dark mode at startup


        # Use our custom CTkDnD main window for drag-and-drop support.
        self.mergerApp = CTkDnD() if master is None else ctk.CTkToplevel(master)
        self.mergerApp.title("Conflux")
        self.mergerApp.iconbitmap(resource_path("style/xtractor-logo.ico"))

        # File paths for three Excel files and the output file.
        self.excel1_path = tk.StringVar()
        self.excel2_path = tk.StringVar()
        self.excel3_path = tk.StringVar()
        self.output_path = tk.StringVar()

        # Header selections for each file (via drop-down lists).
        self.ref_column1 = tk.StringVar()
        self.title_column1 = tk.StringVar()
        self.ref_column2 = tk.StringVar()
        self.title_column2 = tk.StringVar()
        self.ref_column3 = tk.StringVar()
        self.title_column3 = tk.StringVar()

        # Boolean variables for report options and comparing Excel3 title.
        self.compare_excel2_title = tk.BooleanVar(value=False)
        self.generate_word_report = tk.BooleanVar(value=False)
        self.compare_excel3_title = tk.BooleanVar(value=False)

        # Boolean variable for theme mode; True = dark mode.
        self.theme_mode = tk.BooleanVar(value=True)

        self.excel1_headers = []
        self.excel2_headers = []
        self.excel3_headers = []

        self._build_gui()

    def _build_gui(self):
        self.mergerApp.grid_rowconfigure(0, weight=1)
        self.mergerApp.grid_columnconfigure(0, weight=1)
        self.mergerApp.grid_columnconfigure(1, weight=1)
        self.mergerApp.grid_columnconfigure(2, weight=1)

        # Create three frames for Excel 1, 2, and 3.
        self.excel1_frame = ctk.CTkFrame(self.mergerApp)
        self.excel1_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.excel2_frame = ctk.CTkFrame(self.mergerApp)
        self.excel2_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self.excel3_frame = ctk.CTkFrame(self.mergerApp)
        self.excel3_frame.grid(row=0, column=2, padx=10, pady=10, sticky="nsew")

        for frame in (self.excel1_frame, self.excel2_frame, self.excel3_frame):
            frame.grid_rowconfigure(0, weight=0)
            frame.grid_columnconfigure(0, weight=0)
            frame.grid_columnconfigure(1, weight=1)

        self._build_excel1_section(self.excel1_frame)
        self._build_excel2_section(self.excel2_frame)
        self._build_excel3_section(self.excel3_frame)


        # Create the comparison check frame
        self.comparison_frame = ctk.CTkFrame(self.mergerApp)
        self.comparison_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=(10, 10), sticky="nsew")
        self._build_comparison_checks(self.comparison_frame)

        # Build Filename Checker frame to the right
        self.filename_frame = ctk.CTkFrame(self.mergerApp)
        self.filename_frame.grid(row=1, column=2, padx=10, pady=(10, 10), sticky="nsew")
        self._build_filename_checker(self.filename_frame)

        # Create controls frame (move down)
        self.controls_frame = ctk.CTkFrame(self.mergerApp)
        self.controls_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        self._build_controls(self.controls_frame)

    def _build_excel1_section(self, parent_frame):
        font_name = "Helvetica"
        font_size = 12
        self.excel1_button = ctk.CTkButton(parent_frame,
            text="\n➕\n\nSelect Extracted Excel or\nDrag & Drop Here",
            command=self._browse_excel1,
            border_width=3,
            fg_color="transparent",
            hover_color=("#D6D6D6", "#505050"),  # Light and dark hover color
            text_color=("#333333", "#FFFFFF"),
            corner_radius=10,
            width=200,
            height=150)
        self.excel1_button.grid(row=0, column=0, columnspan=2, padx=33, pady=33, sticky="ew")
        # Enable drag and drop on Excel1 button.
        self.excel1_button.drop_target_register(DND_ALL)
        self.excel1_button.dnd_bind('<<Drop>>', self.drop_excel1)
        ctk.CTkLabel(parent_frame, text="Reference Column:", font=(font_name, font_size)).grid(
            row=1, column=0, padx=5, pady=2, sticky="e")
        self.ref_option_menu1 = ctk.CTkOptionMenu(parent_frame, variable=self.ref_column1, values=[])
        self.ref_option_menu1.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        ctk.CTkLabel(parent_frame, text="", font=(font_name, font_size)).grid(
            row=2, column=0, padx=5, pady=2, sticky="e")
        ctk.CTkLabel(parent_frame, text="Drawing Title:", font=(font_name, font_size)).grid(
            row=3, column=0, padx=5, pady=2, sticky="e")
        self.title_option_menu1 = ctk.CTkOptionMenu(parent_frame, variable=self.title_column1, values=[], state="disabled")
        self.title_option_menu1.grid(row=3, column=1, padx=5, pady=2, sticky="ew")

    def _build_excel2_section(self, parent_frame):
        font_name = "Helvetica"
        font_size = 12
        self.excel2_button = ctk.CTkButton(parent_frame,
            text="\n➕\n\nSelect DC_LOD Excel or\nDrag & Drop Here",
            command=self._browse_excel2,
            border_width=3,
            fg_color="transparent",
            hover_color=("#D6D6D6", "#505050"),  # Light and dark hover color
            text_color=("#333333", "#FFFFFF"),
            corner_radius=10,
            width=200,
            height=150)
        self.excel2_button.grid(row=0, column=0, columnspan=2, padx=33, pady=33, sticky="ew")
        self.excel2_button.drop_target_register(DND_ALL)
        self.excel2_button.dnd_bind('<<Drop>>', self.drop_excel2)
        ctk.CTkLabel(parent_frame, text="Reference Column:", font=(font_name, font_size)).grid(
            row=1, column=0, padx=5, pady=2, sticky="e")
        self.ref_option_menu2 = ctk.CTkOptionMenu(parent_frame, variable=self.ref_column2, values=[])
        self.ref_option_menu2.grid(row=1, column=1, padx=5, pady=2, sticky="ew")

        ctk.CTkCheckBox(parent_frame, text="Compare Title 2",
                        variable=self.compare_excel2_title, command=self._toggle_title_entries).grid(
            row=2, column=0, columnspan=2, padx=5, pady=2, sticky="w")

        ctk.CTkLabel(parent_frame, text="Drawing Title:", font=(font_name, font_size)).grid(
            row=3, column=0, padx=5, pady=2, sticky="e")
        self.title_option_menu2 = ctk.CTkOptionMenu(parent_frame, variable=self.title_column2, values=[], state="disabled")
        self.title_option_menu2.grid(row=3, column=1, padx=5, pady=2, sticky="ew")

    def _build_excel3_section(self, parent_frame):
        font_name = "Helvetica"
        font_size = 12
        self.excel3_button = ctk.CTkButton(parent_frame,
            text="\n➕\n\nSelect DD_LOD Excel or\nDrag & Drop Here",
            command=self._browse_excel3,
            border_width=3,
            fg_color="transparent",
            hover_color=("#D6D6D6", "#505050"),  # Light and dark hover color
            text_color=("#333333", "#FFFFFF"),
            corner_radius=10,
            width=200,
            height=150)
        self.excel3_button.grid(row=0, column=0, columnspan=2, padx=33, pady=33, sticky="ew")
        self.excel3_button.drop_target_register(DND_ALL)
        self.excel3_button.dnd_bind('<<Drop>>', self.drop_excel3)
        ctk.CTkLabel(parent_frame, text="Reference Column:", font=(font_name, font_size)).grid(
            row=1, column=0, padx=5, pady=2, sticky="e")
        self.ref_option_menu3 = ctk.CTkOptionMenu(parent_frame, variable=self.ref_column3, values=[])
        self.ref_option_menu3.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        # Place the compare title checkbox above the title dropdown.
        ctk.CTkCheckBox(parent_frame, text="Compare Title 3",
                        variable=self.compare_excel3_title, command=self._toggle_title_entries).grid(
            row=2, column=0, columnspan=2, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(parent_frame, text="Drawing Title:", font=(font_name, font_size)).grid(
            row=3, column=0, padx=5, pady=2, sticky="e")
        self.title_option_menu3 = ctk.CTkOptionMenu(parent_frame, variable=self.title_column3, values=[], state="disabled")
        self.title_option_menu3.grid(row=3, column=1, padx=5, pady=2, sticky="ew")

    def _build_comparison_checks(self, parent_frame):
        """Builds the checkboxes, dropdowns, and textboxes for additional validation."""
        # Add column labels
        ctk.CTkLabel(parent_frame, text="Enable", font=("Helvetica", 12, "bold")).grid(row=0, column=0, padx=5, pady=2,
                                                                                       sticky="w")
        ctk.CTkLabel(parent_frame, text="Column Name", font=("Helvetica", 12, "bold")).grid(row=0, column=1, padx=5,
                                                                                            pady=2, sticky="ew")
        ctk.CTkLabel(parent_frame, text="Expected Value", font=("Helvetica", 12, "bold")).grid(row=0, column=2, padx=5,
                                                                                               pady=2, sticky="ew")

        # Status Check (Moved to row=1)
        self.status_enabled = tk.BooleanVar(value=False)
        self.status_column = tk.StringVar()
        self.status_column.trace_add("write",
                                     lambda *a: self._update_preview_combo(self.status_column, self.status_combo))

        self.status_value = tk.StringVar()

        self.status_check = ctk.CTkCheckBox(
            parent_frame, text="Check 1",
            variable=self.status_enabled, command=self._toggle_status
        )
        self.status_check.grid(row=1, column=0, padx=5, pady=2, sticky="w")

        self.status_dropdown = ctk.CTkOptionMenu(parent_frame, variable=self.status_column, values=[], state="disabled")
        self.status_dropdown.grid(row=1, column=1, padx=5, pady=2, sticky="ew")

        self.status_combo = ctk.CTkComboBox(
            parent_frame,
            variable=self.status_value,
            values=[],
            state="disabled",
            justify="left"
        )
        self.status_combo.grid(row=1, column=2, padx=5, pady=2, sticky="ew")

        # Project Name Check (Moved to row=2)
        self.project_enabled = tk.BooleanVar(value=False)
        self.project_column = tk.StringVar()
        self.project_column.trace_add("write",
                                      lambda *a: self._update_preview_combo(self.project_column, self.project_combo))

        self.project_value = tk.StringVar()

        self.project_check = ctk.CTkCheckBox(
            parent_frame, text="Check 2",
            variable=self.project_enabled, command=self._toggle_project
        )
        self.project_check.grid(row=2, column=0, padx=5, pady=2, sticky="w")

        self.project_dropdown = ctk.CTkOptionMenu(parent_frame, variable=self.project_column, values=[],
                                                  state="disabled")
        self.project_dropdown.grid(row=2, column=1, padx=5, pady=2, sticky="ew")

        self.project_combo = ctk.CTkComboBox(
            parent_frame,
            variable=self.project_value,
            values=[],
            state="disabled",
            justify="left"
        )
        self.project_combo.grid(row=2, column=2, padx=5, pady=2, sticky="ew")

        # Add Custom Checks Button (Moved to row=3)
        self.custom_checks = []
        self.add_check_button = ctk.CTkButton(parent_frame, text="+ Add Check", command=self._add_custom_check)
        self.add_check_button.grid(row=3, column=0, columnspan=3, padx=5, pady=2, sticky="ew")

    def _build_filename_checker(self, parent_frame):
        """Builds the UI for filename comparison against reference column."""

        self.filename_enabled = tk.BooleanVar(value=False)
        self.filename_column = tk.StringVar()

        ctk.CTkLabel(parent_frame, text="Filename vs Reference Number", font=("Helvetica", 12, "bold")).grid(
            row=0, column=0, columnspan=2, padx=5, pady=(5, 2), sticky="w")

        self.filename_check = ctk.CTkCheckBox(
            parent_frame, text="Enable Filename Check",
            variable=self.filename_enabled,
            command=self._toggle_filename_check
        )
        self.filename_check.grid(row=1, column=0, columnspan=2, padx=5, pady=2, sticky="w")

        ctk.CTkLabel(parent_frame, text="Filename Column:", font=("Helvetica", 11)).grid(
            row=2, column=0, padx=5, pady=2, sticky="e")

        self.filename_dropdown = ctk.CTkOptionMenu(parent_frame, variable=self.filename_column, values=[],
                                                   state="disabled")
        self.filename_dropdown.grid(row=2, column=1, padx=5, pady=2, sticky="ew")

    def _build_controls(self, parent_frame):
        font_name = "Helvetica"
        font_size = 12
        ctk.CTkLabel(parent_frame, text="Output Path:", font=(font_name, font_size)).grid(
            row=0, column=0, padx=5, pady=2, sticky="e")
        ctk.CTkEntry(parent_frame, textvariable=self.output_path, width=300).grid(
            row=0, column=1, padx=5, pady=2, sticky="ew")
        ctk.CTkButton(parent_frame, text="Use Excel 1 Path", command=self._use_excel1_path).grid(
            row=0, column=2, padx=5, pady=2)
        # Theme toggle switch (no label) placed at the bottom-right of the controls frame.
        self.theme_switch = ctk.CTkSwitch(parent_frame, text="", variable=self.theme_mode,
                                          command=self.toggle_theme, switch_width=20, switch_height=10)
        # Use place to anchor it at the bottom-right corner of the parent frame.
        self.theme_switch.place(relx=1.0, rely=1.0, anchor="se")
        ctk.CTkButton(parent_frame, text="Start Merge", command=self._start_merge).grid(
            row=2, column=0, columnspan=3, pady=10)
        parent_frame.grid_columnconfigure(1, weight=1)

    def _update_preview_combo(self, column_var, combo_widget):
        col_name = column_var.get()
        values = self.preview_values_by_column.get(col_name, [])
        combo_widget.configure(values=values)

    def toggle_theme(self):
        if self.theme_mode.get():
            ctk.set_appearance_mode("dark")
            print("Theme set to dark mode")
        else:
            ctk.set_appearance_mode("light")
            print("Theme set to light mode")

    def _toggle_status(self):
        """Enable or disable status comparison fields and update custom checks."""
        state = "normal" if self.status_enabled.get() else "disabled"
        self.status_dropdown.configure(state=state)
        self.status_combo.configure(state=state)
        if state == "normal" and self.status_column.get():
            col = self.status_column.get()
            self.status_combo.configure(values=self.preview_values_by_column.get(col, []))

        # ✅ Update all custom checks
        self._toggle_custom_checks()

    def _toggle_project(self):
        """Enable or disable project name comparison fields and update custom checks."""
        state = "normal" if self.project_enabled.get() else "disabled"
        self.project_dropdown.configure(state=state)
        self.project_combo.configure(state=state)
        if state == "normal" and self.project_column.get():
            col = self.project_column.get()
            self.project_combo.configure(values=self.preview_values_by_column.get(col, []))

        # ✅ Update all custom checks
        self._toggle_custom_checks()

    def _toggle_custom_checks(self):
        """Enable or disable custom check dropdowns and entry fields dynamically."""
        for enabled_var, column_var, value_var, dropdown_widget, entry_widget in self.custom_checks:
            state = "normal" if enabled_var.get() else "disabled"
            dropdown_widget.configure(state=state)
            entry_widget.configure(state=state)

    def _toggle_title_entries(self):
        # Enable title_option_menu1 if either checkbox is checked
        state_1 = "normal" if (self.compare_excel2_title.get() or self.compare_excel3_title.get()) else "disabled"
        self.title_option_menu1.configure(state=state_1)

        # Enable title_option_menu2 only if compare_excel2 is checked
        state_2 = "normal" if self.compare_excel2_title.get() else "disabled"
        self.title_option_menu2.configure(state=state_2)

        # Enable title_option_menu3 only if compare_excel3_title is checked
        state_3 = "normal" if self.compare_excel3_title.get() else "disabled"
        self.title_option_menu3.configure(state=state_3)

    def _toggle_filename_check(self):
        state = "normal" if self.filename_enabled.get() else "disabled"
        self.filename_dropdown.configure(state=state)

    def _add_custom_check(self):
        """Adds a new custom check row dynamically and ensures dropdown values are populated."""
        row_idx = len(self.custom_checks) + 3  # Start after Status and Project Name

        enabled_var = tk.BooleanVar(value=False)
        column_var = tk.StringVar()
        value_var = tk.StringVar()

        def update_this_combo(*_):
            col_name = column_var.get()
            values = self.preview_values_by_column.get(col_name, [])
            combo.configure(values=values)

        column_var.trace_add("write", update_this_combo)

        check = ctk.CTkCheckBox(
            self.comparison_frame, text=f"Check {row_idx}",
            variable=enabled_var
        )
        check.grid(row=row_idx, column=0, padx=5, pady=2, sticky="w")

        dropdown = ctk.CTkOptionMenu(self.comparison_frame, variable=column_var, values=[], state="disabled")
        dropdown.grid(row=row_idx, column=1, padx=5, pady=2, sticky="ew")

        combo = ctk.CTkComboBox(
            self.comparison_frame,
            variable=value_var,
            values=[],
            state="disabled",
            justify="left"
        )
        combo.grid(row=row_idx, column=2, padx=5, pady=2, sticky="ew")

        # ✅ Store all required elements (including dropdown widget) to update later
        self.custom_checks.append((enabled_var, column_var, value_var, dropdown, combo))

        # ✅ If Excel is already loaded, populate dropdown values immediately
        if self.excel1_headers:
            dropdown.configure(values=self.excel1_headers)
            column_var.set(auto_select_header(self.excel1_headers, ["status", "project"]))

        # ✅ Ensure new checks enable/disable properly when toggled
        enabled_var.trace_add("write", lambda *args: self._toggle_custom_checks())

        # ✅ Move the + button down
        self.add_check_button.grid(row=row_idx + 1, column=0, columnspan=3, padx=5, pady=2, sticky="ew")

    # --- Drag and Drop Handlers ---
    def drop_excel1(self, event):
        file_path = event.data.strip().replace("{", "").replace("}", "")
        if file_path.lower().endswith((".xlsx", ".xlsm")):
            self.excel1_path.set(file_path)
            self._load_excel1_headers(file_path)
            filename = os.path.basename(file_path)
            self.excel1_button.configure(text=filename, fg_color="#990d10")
        else:
            messagebox.showerror("Error", "Please drag and drop a valid Excel file (.xlsx or .xlsm).")

    def drop_excel2(self, event):
        file_path = event.data.strip().replace("{", "").replace("}", "")
        if file_path.lower().endswith((".xlsx", ".xlsm")):
            self.excel2_path.set(file_path)
            self._load_excel2_headers(file_path)
            filename = os.path.basename(file_path)
            self.excel2_button.configure(text=filename, fg_color="#990d10")
        else:
            messagebox.showerror("Error", "Please drag and drop a valid Excel file (.xlsx or .xlsm).")

    def drop_excel3(self, event):
        file_path = event.data.strip().replace("{", "").replace("}", "")
        if file_path.lower().endswith((".xlsx", ".xlsm")):
            self.excel3_path.set(file_path)
            self._load_excel3_headers(file_path)
            filename = os.path.basename(file_path)
            self.excel3_button.configure(text=filename, fg_color="#990d10")
        else:
            messagebox.showerror("Error", "Please drag and drop a valid Excel file (.xlsx or .xlsm).")

    def _browse_excel1(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")],
                                               title="Select Excel File 1")
        if file_path:
            self.excel1_path.set(file_path)
            self._load_excel1_headers(file_path)
            import os
            filename = os.path.basename(file_path)
            self.excel1_button.configure(text=filename, fg_color="#217346")

    def _browse_excel2(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")],
                                               title="Select Excel File 2")
        if file_path:
            self.excel2_path.set(file_path)
            self._load_excel2_headers(file_path)
            import os
            filename = os.path.basename(file_path)
            self.excel2_button.configure(text=filename, fg_color="#217346")

    def _browse_excel3(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")],
                                               title="Select Excel File 3")
        if file_path:
            self.excel3_path.set(file_path)
            self._load_excel3_headers(file_path)
            import os
            filename = os.path.basename(file_path)
            self.excel3_button.configure(text=filename, fg_color="#217346")

    def _load_excel1_headers(self, file_path):
        try:
            df_preview = pd.read_excel(file_path, engine='openpyxl', nrows=6)
            df = df_preview.head(0)  # only headers

            headers = list(df.columns)

            # Populate existing dropdowns
            self.excel1_headers = headers
            self.preview_values_by_column = {
                header: list(dict.fromkeys(df_preview[header].dropna().astype(str).tolist()))
                for header in headers
            }

            self.ref_option_menu1.configure(values=headers)
            self.title_option_menu1.configure(values=headers)
            self.ref_column1.set(auto_select_header(headers, ["drawing", "sheet", "ref", "number"]))
            self.title_column1.set(auto_select_header(headers, ["title"]))

            # Populate Status and Project dropdowns
            self.status_dropdown.configure(values=headers)
            self.status_column.set(auto_select_header(headers, ["status"]))

            self.project_dropdown.configure(values=headers)
            self.project_column.set(auto_select_header(headers, ["project"]))

            # Populate filename dropdown and auto-select
            self.filename_dropdown.configure(values=headers)
            self.filename_column.set(auto_select_header(headers, ["filename", "file name"]))

            # ✅ Ensure correct widgets are updated for custom checks
            for check in self.custom_checks:
                enabled_var, column_var, value_var, dropdown_widget, _ = check
                dropdown_widget.configure(values=headers)  # Populate dropdown options
                column_var.set(auto_select_header(headers, ["status", "project"]))  # Auto-select if applicable

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load headers from Excel File 1: {e}")

    def _load_excel2_headers(self, file_path):
        try:
            df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
            headers = list(df.columns)
            self.excel2_headers = headers
            self.ref_option_menu2.configure(values=headers)
            self.title_option_menu2.configure(values=headers)
            self.ref_column2.set(auto_select_header(headers, ["drawing", "sheet", "ref", "number"]))
            self.title_column2.set(auto_select_header(headers, ["title"]))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load headers from Excel File 2: {e}")

    def _load_excel3_headers(self, file_path):
        try:
            df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
            headers = list(df.columns)
            self.excel3_headers = headers
            self.ref_option_menu3.configure(values=headers)
            self.title_option_menu3.configure(values=headers)
            self.ref_column3.set(auto_select_header(headers, ["drawing", "sheet", "ref", "number"]))
            self.title_column3.set(auto_select_header(headers, ["title"]))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load headers from Excel File 3: {e}")

    def _use_excel1_path(self):
        excel1 = self.excel1_path.get()
        if excel1:
            directory, file_name = os.path.split(excel1)
            name, ext = os.path.splitext(file_name)
            self.output_path.set(os.path.join(directory, f"{name}_merged{ext}"))

    def _start_merge(self):
        # 1. Collect file paths and reference columns
        paths = [self.excel1_path.get(), self.excel2_path.get()]
        refs = [self.ref_column1.get(), self.ref_column2.get()]

        excel3 = self.excel3_path.get().strip()
        if excel3:
            paths.append(excel3)
            refs.append(self.ref_column3.get())

        # 2. Resolve output path (with timestamp if exists)
        output = self.output_path.get().strip()
        if not output:
            messagebox.showerror("Error", "Please provide a valid output path.")
            return
        if os.path.exists(output):
            directory, file_name = os.path.split(output)
            base, ext = os.path.splitext(file_name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output = os.path.join(directory, f"{base}_{timestamp}{ext}")

        # 3. Build title_columns list
        title_cols = [
            (self.title_column1.get()
             if (self.compare_excel2_title.get() or self.compare_excel3_title.get())
             else None),
            (self.title_column2.get()
             if self.compare_excel2_title.get()
             else None),
        ]
        if excel3:
            title_cols.append(
                self.title_column3.get()
                if self.compare_excel3_title.get()
                else None
            )

        # 4. Assemble validation config
        cfg = CheckConfig(
            status_column=self.status_column.get() if self.status_enabled.get() else None,
            status_value=self.status_value.get() if self.status_enabled.get() else None,
            project_column=self.project_column.get() if self.project_enabled.get() else None,
            project_value=self.project_value.get() if self.project_enabled.get() else None,
            custom_checks=[
                (col.get(), val.get())
                for en, col, val, _, _ in self.custom_checks
                if en.get() and col.get()
            ],
            filename_column=self.filename_column.get() if self.filename_enabled.get() else None
        )

        # 5. Delegate to the facade
        try:
            merged_df = MergerFacade.run_merge(
                paths=paths,
                ref_columns=refs,
                output_path=output,
                title_columns=title_cols,
                check_config=cfg
            )

            # open file
            if messagebox.askyesno(
                    "Success",
                    f"Merged file saved at\n\n{output}\n\nWould you like to open it now?"
            ):
                try:
                    os.startfile(output)
                except AttributeError:
                    # not Windows? try the generic way
                    import subprocess, sys
                    if sys.platform == "darwin":
                        subprocess.Popen(["open", output])
                    else:
                        subprocess.Popen(["xdg-open", output])

        except Exception as e:
            # messagebox.showerror("Error", str(e))
            traceback.print_exc()

    def run(self):
        if isinstance(self.mergerApp, ctk.CTk):
            self.mergerApp.mainloop()


if __name__ == "__main__":
    app = MergerGUI()
    app.run()



