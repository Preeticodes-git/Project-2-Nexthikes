import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import math
import csv
import statistics

# Optional imports (taaki agar pandas/matplotlib install na bhi ho toh app crash na ho)
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False


THEMES = {
    "dark": {
        "bg": "#1e1e2e",
        "panel": "#181825",
        "surface": "#313244",
        "num_btn": "#45475a",
        "func_btn": "#313244",
        "text": "#cdd6f4",
        "muted": "#a6adc8",
        "accent": "#89b4fa",
        "green": "#a6e3a1",
        "red": "#f38ba8",
        "orange": "#fab387",
        "yellow": "#f9e2af",
        "purple": "#cba6f7",
        "btn_dark_fg": "#11111b",
        "tab_selected_fg": "#11111b",
        "spine": "#6c7086",
    },
    "light": {
        "bg": "#f4f6fb",
        "panel": "#e6eaf2",
        "surface": "#ffffff",
        "num_btn": "#e2e8f0",
        "func_btn": "#dbeafe",
        "text": "#1e293b",
        "muted": "#64748b",
        "accent": "#2563eb",
        "green": "#16a34a",
        "red": "#e11d48",
        "orange": "#ea580c",
        "yellow": "#facc15",
        "purple": "#7c3aed",
        "btn_dark_fg": "#ffffff",
        "tab_selected_fg": "#ffffff",
        "spine": "#94a3b8",
    },
}


class AdvancedSuperCalculator(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Super Advanced Data Scientist Calculator Suite")
        self.geometry("940x680")

        self.df = None
        self.csv_rows = []
        self.csv_headers = []
        self.manual_override = False

        # Detect initial theme from current system time (06:00 - 18:59 -> Light, else Dark)
        self.current_theme_name = self.get_time_based_theme()
        self.theme = THEMES[self.current_theme_name]

        self.configure(bg=self.theme["bg"])
        self.style = ttk.Style()
        self.style.theme_use("clam")

        # Top Header Bar with Live Clock & Theme Switcher
        self.header_bar = tk.Frame(self, bg=self.theme["panel"], pady=6, padx=12)
        self.header_bar.pack(fill="x", padx=10, pady=(10, 0))

        self.clock_lbl = tk.Label(
            self.header_bar,
            text="",
            font=("Segoe UI", 9, "bold"),
            bg=self.theme["panel"],
            fg=self.theme["accent"],
        )
        self.clock_lbl.pack(side="left")

        self.auto_mode_btn = tk.Button(
            self.header_bar,
            text="Auto (Time)",
            font=("Segoe UI", 8, "bold"),
            bg=self.theme["surface"],
            fg=self.theme["text"],
            relief="flat",
            cursor="hand2",
            command=self.enable_auto_theme,
        )
        self.auto_mode_btn.pack(side="right", padx=(6, 0), ipadx=6, ipady=2)

        self.theme_toggle_btn = tk.Button(
            self.header_bar,
            text="",
            font=("Segoe UI", 9, "bold"),
            bg=self.theme["accent"],
            fg=self.theme["btn_dark_fg"],
            relief="flat",
            cursor="hand2",
            command=self.toggle_theme_manually,
        )
        self.theme_toggle_btn.pack(side="right", ipadx=8, ipady=2)

        # Notebook tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=8)

        self.calc_frame = tk.Frame(self.notebook, bg=self.theme["bg"])
        self.weight_frame = tk.Frame(self.notebook, bg=self.theme["bg"])
        self.temp_frame = tk.Frame(self.notebook, bg=self.theme["bg"])
        self.stats_frame = tk.Frame(self.notebook, bg=self.theme["bg"])
        self.excel_frame = tk.Frame(self.notebook, bg=self.theme["bg"])
        self.gst_frame = tk.Frame(self.notebook, bg=self.theme["bg"])
        self.finance_frame = tk.Frame(self.notebook, bg=self.theme["bg"])
        self.converter_frame = tk.Frame(self.notebook, bg=self.theme["bg"])

        self.notebook.add(self.calc_frame, text=" Calculator & History ")
        self.notebook.add(self.weight_frame, text=" Weight (Kg/Lbs) ")
        self.notebook.add(self.temp_frame, text=" Temperature ")
        self.notebook.add(self.stats_frame, text=" Stats & Graphs ")
        self.notebook.add(self.excel_frame, text=" Excel Tools ")
        self.notebook.add(self.gst_frame, text=" GST ")
        self.notebook.add(self.finance_frame, text=" Finance SI/CI ")
        self.notebook.add(self.converter_frame, text=" Data Units ")

        # Build sections
        self.setup_calculator()
        self.setup_weight_calculator()
        self.setup_temperature_calculator()
        self.setup_stats_graphs()
        self.setup_excel_tools()
        self.setup_gst_calculator()
        self.setup_finance_calculator()
        self.setup_converter()

        # Apply theme styling across all widgets & start clock monitor
        self.apply_theme()
        self.update_clock_and_theme()

        # Enter key shortcut
        self.bind("<Return>", lambda e: self.equal_calc() if self.notebook.index(self.notebook.select()) == 0 else None)
        self.bind("<KP_Enter>", lambda e: self.equal_calc() if self.notebook.index(self.notebook.select()) == 0 else None)

    # ==================== TIME-BASED THEME ENGINE ====================
    def get_time_based_theme(self):
        hour = datetime.now().hour
        # 6:00 AM to 6:59 PM -> Light Mode | 7:00 PM to 5:59 AM -> Dark Mode
        return "light" if 6 <= hour < 19 else "dark"

    def toggle_theme_manually(self):
        self.manual_override = True
        self.current_theme_name = "light" if self.current_theme_name == "dark" else "dark"
        self.theme = THEMES[self.current_theme_name]
        self.apply_theme()

    def enable_auto_theme(self):
        self.manual_override = False
        target = self.get_time_based_theme()
        self.current_theme_name = target
        self.theme = THEMES[target]
        self.apply_theme()

    def update_clock_and_theme(self):
        now = datetime.now()
        time_str = now.strftime("%I:%M %p")
        mode_tag = "Manual" if self.manual_override else "Auto-Time"
        self.clock_lbl.config(
            text=f"🕒 {time_str}  |  Theme: {self.current_theme_name.capitalize()} Mode ({mode_tag})"
        )

        if not self.manual_override:
            target = self.get_time_based_theme()
            if target != self.current_theme_name:
                self.current_theme_name = target
                self.theme = THEMES[target]
                self.apply_theme()

        # Re-check every 15 seconds
        self.after(15000, self.update_clock_and_theme)

    def apply_theme(self):
        t = self.theme
        self.configure(bg=t["bg"])

        # Update ttk Notebook & Combobox styles
        self.style.configure("TNotebook", background=t["bg"], borderwidth=0)
        self.style.configure(
            "TNotebook.Tab",
            background=t["panel"],
            foreground=t["text"],
            padding=[10, 8],
            font=("Segoe UI", 9, "bold"),
        )
        self.style.map(
            "TNotebook.Tab",
            background=[("selected", t["accent"])],
            foreground=[("selected", t["tab_selected_fg"])],
        )

        # Update Header Bar
        self.header_bar.configure(bg=t["panel"])
        self.clock_lbl.configure(bg=t["panel"], fg=t["accent"])
        self.auto_mode_btn.configure(bg=t["surface"], fg=t["text"])
        btn_label = "🌙 Switch to Dark" if self.current_theme_name == "light" else "☀ Switch to Light"
        self.theme_toggle_btn.configure(text=btn_label, bg=t["accent"], fg=t["btn_dark_fg"])

        # Update Tab Frames
        for frame in (
            self.calc_frame, self.weight_frame, self.temp_frame, self.stats_frame,
            self.excel_frame, self.gst_frame, self.finance_frame, self.converter_frame
        ):
            frame.configure(bg=t["bg"])

        # Calculator Tab Widgets
        self.left_panel.configure(bg=t["bg"])
        self.right_panel.configure(bg=t["panel"])
        self.mode_bar.configure(bg=t["bg"])
        self.btn_frame.configure(bg=t["bg"])
        self.calc_entry.configure(
            bg=t["surface"],
            fg=t["green"] if self.current_theme_name == "dark" else t["text"],
            insertbackground=t["text"],
        )
        self.trig_lbl.configure(bg=t["bg"], fg=t["text"])
        for rb in (self.deg_rb, self.rad_rb):
            rb.configure(bg=t["bg"], fg=t["accent"], selectcolor=t["surface"], activebackground=t["bg"])

        for text, btn in self.calc_buttons.items():
            if text == "=":
                btn.configure(bg=t["green"], fg=t["btn_dark_fg"])
            elif text == "C":
                btn.configure(bg=t["red"], fg=t["btn_dark_fg"])
            elif text == "⌫":
                btn.configure(bg=t["orange"], fg=t["btn_dark_fg"])
            elif text in "0123456789.":
                btn.configure(bg=t["num_btn"], fg=t["text"])
            else:
                btn.configure(bg=t["func_btn"], fg=t["accent"])

        self.hist_title_lbl.configure(bg=t["panel"], fg=t["accent"])
        self.hist_sub_lbl.configure(bg=t["panel"], fg=t["muted"])
        self.history_listbox.configure(
            bg=t["surface"], fg=t["text"], selectbackground=t["accent"], selectforeground=t["tab_selected_fg"]
        )
        self.clear_hist_btn.configure(bg=t["red"], fg=t["btn_dark_fg"])

        # Weight Tab
        self.weight_title.configure(bg=t["bg"], fg=t["green"])
        self.weight_card.configure(bg=t["panel"])
        self.weight_lbl.configure(bg=t["panel"], fg=t["text"])
        self.weight_entry.configure(bg=t["surface"], fg=t["text"], insertbackground=t["text"])
        self.weight_btn_row.configure(bg=t["panel"])
        self.kg_to_lbs_btn.configure(bg=t["accent"], fg=t["btn_dark_fg"])
        self.lbs_to_kg_btn.configure(bg=t["yellow"], fg="#11111b")
        self.weight_result_lbl.configure(bg=t["bg"], fg=t["orange"])

        # Temperature Tab
        self.temp_title.configure(bg=t["bg"], fg=t["green"])
        self.temp_card.configure(bg=t["panel"])
        self.temp_input_row.configure(bg=t["panel"])
        self.temp_entry.configure(bg=t["surface"], fg=t["text"], insertbackground=t["text"])
        self.temp_convert_btn.configure(bg=t["accent"], fg=t["btn_dark_fg"])
        self.temp_result_lbl.configure(bg=t["bg"], fg=t["green"])

        # Stats & Graphs Tab
        self.stats_ctrl_bar.configure(bg=t["panel"])
        self.stats_data_lbl.configure(bg=t["panel"], fg=t["text"])
        self.stats_data_entry.configure(bg=t["surface"], fg=t["text"], insertbackground=t["text"])
        self.stats_plot_btn.configure(bg=t["green"], fg=t["btn_dark_fg"])
        self.stats_summary_lbl.configure(bg=t["surface"], fg=t["accent"])
        self.chart_container.configure(bg=t["bg"])

        # Excel Tools Tab
        self.excel_toolbar.configure(bg=t["panel"])
        self.excel_load_btn.configure(bg=t["accent"], fg=t["btn_dark_fg"])
        self.excel_desc_btn.configure(bg=t["green"], fg=t["btn_dark_fg"])
        self.excel_miss_btn.configure(bg=t["yellow"], fg="#11111b")
        self.excel_clean_btn.configure(bg=t["orange"], fg=t["btn_dark_fg"])
        self.excel_text_frame.configure(bg=t["bg"])
        self.excel_console.configure(bg=t["surface"], fg=t["text"])

        # GST Tab
        self.gst_title.configure(bg=t["bg"], fg=t["green"])
        self.gst_form.configure(bg=t["bg"])
        self.gst_amt_lbl.configure(bg=t["bg"], fg=t["text"])
        self.gst_amt_entry.configure(bg=t["surface"], fg=t["text"], insertbackground=t["text"])
        self.gst_rate_lbl.configure(bg=t["bg"], fg=t["text"])
        self.gst_rate_entry.configure(bg=t["surface"], fg=t["text"], insertbackground=t["text"])
        self.gst_btn_box.configure(bg=t["bg"])
        self.gst_add_btn.configure(bg=t["accent"], fg=t["btn_dark_fg"])
        self.gst_rem_btn.configure(bg=t["yellow"], fg="#11111b")
        self.gst_result_lbl.configure(bg=t["bg"], fg=t["orange"])

        # Finance Tab
        self.fin_title.configure(bg=t["bg"], fg=t["green"])
        self.fin_form.configure(bg=t["bg"])
        for lbl in (self.fin_p_lbl, self.fin_r_lbl, self.fin_t_lbl):
            lbl.configure(bg=t["bg"], fg=t["text"])
        for ent in (self.fin_p, self.fin_r, self.fin_t):
            ent.configure(bg=t["surface"], fg=t["text"], insertbackground=t["text"])
        self.fin_calc_btn.configure(bg=t["accent"], fg=t["btn_dark_fg"])
        self.fin_result_lbl.configure(bg=t["bg"], fg=t["orange"])

        # Data Units Tab
        self.conv_title.configure(bg=t["bg"], fg=t["green"])
        self.conv_form.configure(bg=t["bg"])
        self.byte_lbl.configure(bg=t["bg"], fg=t["text"])
        self.byte_entry.configure(bg=t["surface"], fg=t["text"], insertbackground=t["text"])
        self.conv_btn.configure(bg=t["accent"], fg=t["btn_dark_fg"])
        self.conv_result_lbl.configure(bg=t["bg"], fg=t["green"])

    # ==================== 1. SCIENTIFIC CALCULATOR & HISTORY ====================
    def setup_calculator(self):
        self.left_panel = tk.Frame(self.calc_frame)
        self.left_panel.pack(side="left", fill="both", expand=True, padx=(10, 5), pady=10)

        self.right_panel = tk.Frame(self.calc_frame, width=260)
        self.right_panel.pack(side="right", fill="y", padx=(5, 10), pady=10)
        self.right_panel.pack_propagate(False)

        self.calc_entry = tk.Entry(
            self.left_panel,
            font=("Consolas", 20, "bold"),
            bd=0,
            justify="right",
            insertwidth=3,
        )
        self.calc_entry.pack(fill="x", pady=(0, 10), ipady=12)
        self.calc_entry.focus_set()

        self.mode_bar = tk.Frame(self.left_panel)
        self.mode_bar.pack(fill="x", pady=(0, 6))
        self.angle_mode = tk.StringVar(value="DEG")
        self.trig_lbl = tk.Label(self.mode_bar, text="Trig Mode:", font=("Segoe UI", 9, "bold"))
        self.trig_lbl.pack(side="left", padx=(0, 6))
        self.deg_rb = tk.Radiobutton(
            self.mode_bar, text="Degrees (θ°)", variable=self.angle_mode, value="DEG", font=("Segoe UI", 9, "bold")
        )
        self.deg_rb.pack(side="left", padx=4)
        self.rad_rb = tk.Radiobutton(
            self.mode_bar, text="Radians (rad)", variable=self.angle_mode, value="RAD", font=("Segoe UI", 9, "bold")
        )
        self.rad_rb.pack(side="left", padx=4)

        self.btn_frame = tk.Frame(self.left_panel)
        self.btn_frame.pack(fill="both", expand=True)

        buttons = [
            ('sin', 0, 0), ('cos', 0, 1), ('tan', 0, 2), ('log', 0, 3), ('ln', 0, 4), ('C', 0, 5),
            ('asin', 1, 0), ('acos', 1, 1), ('atan', 1, 2), ('eˣ', 1, 3), ('10ˣ', 1, 4), ('⌫', 1, 5),
            ('√', 2, 0), ('x²', 2, 1), ('xʸ', 2, 2), ('(', 2, 3), (')', 2, 4), ('/', 2, 5),
            ('7', 3, 0), ('8', 3, 1), ('9', 3, 2), ('!', 3, 3), ('|x|', 3, 4), ('*', 3, 5),
            ('4', 4, 0), ('5', 4, 1), ('6', 4, 2), ('π', 4, 3), ('e', 4, 4), ('-', 4, 5),
            ('1', 5, 0), ('2', 5, 1), ('3', 5, 2), ('%', 5, 3), ('1/x', 5, 4), ('+', 5, 5),
            ('0', 6, 0), ('.', 6, 1), ('floor', 6, 2), ('ceil', 6, 3), ('=', 6, 4),
        ]

        self.calc_buttons = {}
        for text, r, c in buttons:
            colspan = 2 if text == '=' else 1
            if text == '=':
                cmd = self.equal_calc
            elif text == 'C':
                cmd = self.clear_calc
            elif text == '⌫':
                cmd = self.backspace_calc
            else:
                cmd = lambda x=text: self.press_calc(x)

            btn = tk.Button(
                self.btn_frame,
                text=text,
                font=("Segoe UI", 10, "bold"),
                command=cmd,
                relief="flat",
                cursor="hand2",
            )
            btn.grid(row=r, column=c, columnspan=colspan, sticky="nsew", padx=2, pady=2)
            self.calc_buttons[text] = btn

        for i in range(7):
            self.btn_frame.rowconfigure(i, weight=1)
        for j in range(6):
            self.btn_frame.columnconfigure(j, weight=1)

        self.hist_title_lbl = tk.Label(self.right_panel, text="Calculation History", font=("Segoe UI", 11, "bold"))
        self.hist_title_lbl.pack(pady=(10, 5))

        self.hist_sub_lbl = tk.Label(self.right_panel, text="(Click any item to reuse)", font=("Segoe UI", 8))
        self.hist_sub_lbl.pack(pady=(0, 6))

        self.history_listbox = tk.Listbox(
            self.right_panel, font=("Consolas", 10), bd=0, highlightthickness=0
        )
        self.history_listbox.pack(fill="both", expand=True, padx=10, pady=5)
        self.history_listbox.bind("<<ListboxSelect>>", self.reuse_history)

        self.clear_hist_btn = tk.Button(
            self.right_panel,
            text="Clear History",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            cursor="hand2",
            command=lambda: self.history_listbox.delete(0, tk.END),
        )
        self.clear_hist_btn.pack(fill="x", padx=10, pady=10, ipady=4)

    def press_calc(self, key):
        func_with_parens = {
            'sin': 'sin()', 'cos': 'cos()', 'tan': 'tan()',
            'asin': 'asin()', 'acos': 'acos()', 'atan': 'atan()',
            'log': 'log10()', 'ln': 'log()', 'eˣ': 'exp()',
            '√': 'sqrt()', '!': 'factorial()', '|x|': 'abs()',
            '1/x': '1/()', 'floor': 'floor()', 'ceil': 'ceil()',
        }
        simple_mapping = {
            'x²': '**2', 'xʸ': '**', '10ˣ': '10**', 'π': 'pi', 'e': 'e', '%': '/100',
        }

        self.calc_entry.focus_set()
        if key in func_with_parens:
            token = func_with_parens[key]
            pos = self.calc_entry.index(tk.INSERT)
            self.calc_entry.insert(pos, token)
            self.calc_entry.icursor(pos + len(token) - 1)
        else:
            token = simple_mapping.get(key, str(key))
            pos = self.calc_entry.index(tk.INSERT)
            self.calc_entry.insert(pos, token)

    def _trig_env(self):
        is_deg = self.angle_mode.get() == "DEG"
        return {
            "__builtins__": None,
            "sqrt": math.sqrt, "factorial": math.factorial,
            "log10": math.log10, "log": math.log, "exp": math.exp,
            "abs": abs, "floor": math.floor, "ceil": math.ceil,
            "pi": math.pi, "e": math.e,
            "sin": lambda x: math.sin(math.radians(x)) if is_deg else math.sin(x),
            "cos": lambda x: math.cos(math.radians(x)) if is_deg else math.cos(x),
            "tan": lambda x: math.tan(math.radians(x)) if is_deg else math.tan(x),
            "asin": lambda x: math.degrees(math.asin(x)) if is_deg else math.asin(x),
            "acos": lambda x: math.degrees(math.acos(x)) if is_deg else math.acos(x),
            "atan": lambda x: math.degrees(math.atan(x)) if is_deg else math.atan(x),
        }

    def equal_calc(self):
        expr = self.calc_entry.get().strip()
        if not expr:
            return
        try:
            open_count = expr.count("(") - expr.count(")")
            clean_expr = expr + (")" * max(0, open_count))
            raw_result = eval(clean_expr, self._trig_env())
            result = str(round(raw_result, 10)) if isinstance(raw_result, float) else str(raw_result)
            self.history_listbox.insert(0, f"{clean_expr} = {result}")
            self.calc_entry.delete(0, tk.END)
            self.calc_entry.insert(0, result)
        except Exception:
            self.calc_entry.delete(0, tk.END)
            self.calc_entry.insert(0, "Error")

    def reuse_history(self, event):
        selection = self.history_listbox.curselection()
        if selection:
            entry_text = self.history_listbox.get(selection[0])
            result_part = entry_text.split(" = ")[-1]
            self.calc_entry.delete(0, tk.END)
            self.calc_entry.insert(0, result_part)
            self.calc_entry.focus_set()

    def clear_calc(self):
        self.calc_entry.delete(0, tk.END)
        self.calc_entry.focus_set()

    def backspace_calc(self):
        pos = self.calc_entry.index(tk.INSERT)
        if pos > 0:
            self.calc_entry.delete(pos - 1)
        self.calc_entry.focus_set()

    # ==================== 2. WEIGHT CALCULATOR (KG & LBS) ====================
    def setup_weight_calculator(self):
        self.weight_title = tk.Label(self.weight_frame, text="Weight Converter (Kilograms ↔ Pounds)", font=("Segoe UI", 14, "bold"))
        self.weight_title.pack(pady=20)

        self.weight_card = tk.Frame(self.weight_frame, padx=30, pady=25)
        self.weight_card.pack(pady=10)

        self.weight_lbl = tk.Label(self.weight_card, text="Enter Weight Value:", font=("Segoe UI", 11))
        self.weight_lbl.pack(pady=(0, 8))

        self.weight_entry = tk.Entry(self.weight_card, font=("Consolas", 14), bd=0, justify="center")
        self.weight_entry.pack(ipady=6, ipadx=15, pady=(0, 15))

        self.weight_btn_row = tk.Frame(self.weight_card)
        self.weight_btn_row.pack(pady=5)

        self.kg_to_lbs_btn = tk.Button(
            self.weight_btn_row, text="Kilograms (kg) → Pounds (lbs)", font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2", command=lambda: self.convert_weight("kg_to_lbs")
        )
        self.kg_to_lbs_btn.pack(side="left", padx=8, ipadx=8, ipady=6)

        self.lbs_to_kg_btn = tk.Button(
            self.weight_btn_row, text="Pounds (lbs) → Kilograms (kg)", font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2", command=lambda: self.convert_weight("lbs_to_kg")
        )
        self.lbs_to_kg_btn.pack(side="left", padx=8, ipadx=8, ipady=6)

        self.weight_result_lbl = tk.Label(self.weight_frame, text="", font=("Segoe UI", 13, "bold"))
        self.weight_result_lbl.pack(pady=20)

    def convert_weight(self, mode):
        try:
            val = float(self.weight_entry.get())
            if mode == "kg_to_lbs":
                lbs = val * 2.20462262
                grams = val * 1000
                self.weight_result_lbl.config(text=f"{val} kg  =  {lbs:.4f} lbs  ({grams:,.2f} grams)")
            else:
                kg = val / 2.20462262
                oz = val * 16
                self.weight_result_lbl.config(text=f"{val} lbs  =  {kg:.4f} kg  ({oz:,.2f} ounces)")
        except ValueError:
            messagebox.showerror("Input Error", "Please enter a valid numerical weight.")

    # ==================== 3. TEMPERATURE CALCULATOR ====================
    def setup_temperature_calculator(self):
        self.temp_title = tk.Label(self.temp_frame, text="Temperature Converter (°C / °F / K)", font=("Segoe UI", 14, "bold"))
        self.temp_title.pack(pady=20)

        self.temp_card = tk.Frame(self.temp_frame, padx=30, pady=25)
        self.temp_card.pack(pady=10)

        self.temp_input_row = tk.Frame(self.temp_card)
        self.temp_input_row.pack(pady=10)

        self.temp_entry = tk.Entry(self.temp_input_row, font=("Consolas", 14), width=14, bd=0, justify="center")
        self.temp_entry.pack(side="left", ipady=6, padx=6)

        self.temp_unit_combo = ttk.Combobox(
            self.temp_input_row, values=["Celsius (°C)", "Fahrenheit (°F)", "Kelvin (K)"],
            state="readonly", font=("Segoe UI", 11), width=16
        )
        self.temp_unit_combo.current(0)
        self.temp_unit_combo.pack(side="left", ipady=5, padx=6)

        self.temp_convert_btn = tk.Button(
            self.temp_card, text="Convert Temperature", font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2", command=self.convert_temperature
        )
        self.temp_convert_btn.pack(pady=15, ipadx=15, ipady=6)

        self.temp_result_lbl = tk.Label(self.temp_frame, text="", font=("Segoe UI", 12, "bold"), justify="center")
        self.temp_result_lbl.pack(pady=15)

    def convert_temperature(self):
        try:
            val = float(self.temp_entry.get())
            unit = self.temp_unit_combo.get()
            if "Celsius" in unit:
                c, f, k = val, (val * 9 / 5) + 32, val + 273.15
            elif "Fahrenheit" in unit:
                f = val
                c = (f - 32) * 5 / 9
                k = c + 273.15
            else:
                k = val
                c = k - 273.15
                f = (c * 9 / 5) + 32
            self.temp_result_lbl.config(text=f"Celsius: {c:.2f} °C\nFahrenheit: {f:.2f} °F\nKelvin: {k:.2f} K")
        except ValueError:
            messagebox.showerror("Input Error", "Please enter a valid temperature value.")

    # ==================== 4. STATISTICS & GRAPHS MODULE ====================
    def setup_stats_graphs(self):
        self.stats_ctrl_bar = tk.Frame(self.stats_frame, pady=10, padx=10)
        self.stats_ctrl_bar.pack(fill="x", padx=10, pady=(10, 5))

        self.stats_data_lbl = tk.Label(self.stats_ctrl_bar, text="Data (comma-separated):", font=("Segoe UI", 9, "bold"))
        self.stats_data_lbl.pack(side="left", padx=5)

        self.stats_data_entry = tk.Entry(self.stats_ctrl_bar, font=("Consolas", 10), width=34, bd=0)
        self.stats_data_entry.insert(0, "14, 18, 21, 21, 25, 29, 34, 38, 45, 52")
        self.stats_data_entry.pack(side="left", ipady=5, padx=5)

        self.graph_type_combo = ttk.Combobox(
            self.stats_ctrl_bar, values=["Line Plot", "Bar Chart", "Histogram", "Box Plot", "Scatter Plot"],
            state="readonly", width=13, font=("Segoe UI", 9)
        )
        self.graph_type_combo.current(0)
        self.graph_type_combo.pack(side="left", padx=5)

        self.stats_plot_btn = tk.Button(
            self.stats_ctrl_bar, text="Compute & Plot", font=("Segoe UI", 9, "bold"),
            relief="flat", cursor="hand2", command=self.generate_stats_and_graph
        )
        self.stats_plot_btn.pack(side="left", padx=5, ipadx=8, ipady=4)

        self.stats_summary_lbl = tk.Label(
            self.stats_frame,
            text="Mean: --  |  Median: --  |  Mode: --  |  Std Dev: --  |  Variance: --",
            font=("Consolas", 10, "bold"), pady=8
        )
        self.stats_summary_lbl.pack(fill="x", padx=10, pady=5)

        self.chart_container = tk.Frame(self.stats_frame)
        self.chart_container.pack(fill="both", expand=True, padx=10, pady=5)

    def generate_stats_and_graph(self):
        try:
            t = self.theme
            raw_items = [x.strip() for x in self.stats_data_entry.get().split(",") if x.strip()]
            data = [float(x) for x in raw_items]
            if len(data) < 2:
                messagebox.showwarning("Data Warning", "Please enter at least 2 numerical values.")
                return

            mean_v = statistics.mean(data)
            med_v = statistics.median(data)
            modes = statistics.multimode(data)
            mode_str = ", ".join(f"{m:g}" for m in modes[:2])
            std_v = statistics.stdev(data)
            var_v = statistics.variance(data)

            self.stats_summary_lbl.config(
                text=f"Mean: {mean_v:.2f}  |  Median: {med_v:.2f}  |  Mode: {mode_str}  |  Std Dev: {std_v:.2f}  |  Var: {var_v:.2f}"
            )

            for widget in self.chart_container.winfo_children():
                widget.destroy()

            if not HAS_MATPLOTLIB:
                canvas = tk.Canvas(self.chart_container, bg=t["surface"], highlightthickness=0)
                canvas.pack(fill="both", expand=True)
                canvas.update()
                w, h = max(canvas.winfo_width(), 600), max(canvas.winfo_height(), 320)
                pad = 40
                max_val, min_val = max(data), min(0, min(data))
                span = max(max_val - min_val, 1)
                bar_w = (w - 2 * pad) / len(data)
                for idx, val in enumerate(data):
                    x0 = pad + idx * bar_w + 4
                    x1 = pad + (idx + 1) * bar_w - 4
                    y1 = h - pad
                    y0 = y1 - ((val - min_val) / span) * (h - 2 * pad)
                    canvas.create_rectangle(x0, y0, x1, y1, fill=t["accent"], outline="")
                    canvas.create_text((x0 + x1) / 2, y0 - 10, text=f"{val:g}", fill=t["text"], font=("Consolas", 8))
                return

            fig, ax = plt.subplots(figsize=(7, 4), dpi=100)
            fig.patch.set_facecolor(t["bg"])
            ax.set_facecolor(t["surface"])
            ax.tick_params(colors=t["text"])
            for spine in ax.spines.values():
                spine.set_color(t["spine"])

            chart = self.graph_type_combo.get()
            if chart == "Line Plot":
                ax.plot(range(1, len(data) + 1), data, marker="o", color=t["accent"], linewidth=2)
                ax.set_xlabel("Index", color=t["text"])
            elif chart == "Bar Chart":
                ax.bar(range(1, len(data) + 1), data, color=t["green"])
                ax.set_xlabel("Index", color=t["text"])
            elif chart == "Histogram":
                ax.hist(data, bins="auto", color=t["orange"], edgecolor=t["bg"])
                ax.set_xlabel("Value Bins", color=t["text"])
            elif chart == "Box Plot":
                ax.boxplot(
                    data, patch_artist=True,
                    boxprops=dict(facecolor=t["purple"], color=t["text"]),
                    medianprops=dict(color=t["bg"], linewidth=2),
                )
            elif chart == "Scatter Plot":
                ax.scatter(range(1, len(data) + 1), data, color=t["red"], s=60)
                ax.set_xlabel("Index", color=t["text"])

            ax.set_ylabel("Values", color=t["text"])
            ax.set_title(f"Data Visualization - {chart}", color=t["green"], fontsize=12, fontweight="bold")
            ax.grid(True, linestyle="--", alpha=0.25, color=t["text"])

            canvas = FigureCanvasTkAgg(fig, master=self.chart_container)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
            plt.close(fig)

        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid comma-separated numbers.")

    # ==================== 5. EXCEL TOOLS MODULE ====================
    def setup_excel_tools(self):
        self.excel_toolbar = tk.Frame(self.excel_frame, pady=10, padx=10)
        self.excel_toolbar.pack(fill="x", padx=10, pady=(10, 5))

        self.excel_load_btn = tk.Button(
            self.excel_toolbar, text="Load Excel / CSV", font=("Segoe UI", 9, "bold"),
            relief="flat", cursor="hand2", command=self.excel_load_file
        )
        self.excel_load_btn.pack(side="left", padx=4, ipadx=8, ipady=5)

        self.excel_desc_btn = tk.Button(
            self.excel_toolbar, text="Summary Stats (.describe)", font=("Segoe UI", 9, "bold"),
            relief="flat", cursor="hand2", command=self.excel_describe
        )
        self.excel_desc_btn.pack(side="left", padx=4, ipadx=8, ipady=5)

        self.excel_miss_btn = tk.Button(
            self.excel_toolbar, text="Check Missing & Info", font=("Segoe UI", 9, "bold"),
            relief="flat", cursor="hand2", command=self.excel_missing_info
        )
        self.excel_miss_btn.pack(side="left", padx=4, ipadx=8, ipady=5)

        self.excel_clean_btn = tk.Button(
            self.excel_toolbar, text="Clean Nulls & Export", font=("Segoe UI", 9, "bold"),
            relief="flat", cursor="hand2", command=self.excel_clean_export
        )
        self.excel_clean_btn.pack(side="left", padx=4, ipadx=8, ipady=5)

        self.excel_text_frame = tk.Frame(self.excel_frame)
        self.excel_text_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.excel_console = tk.Text(self.excel_text_frame, font=("Consolas", 10), bd=0, padx=12, pady=12, wrap="none")
        y_scroll = tk.Scrollbar(self.excel_text_frame, orient="vertical", command=self.excel_console.yview)
        x_scroll = tk.Scrollbar(self.excel_text_frame, orient="horizontal", command=self.excel_console.xview)
        self.excel_console.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        y_scroll.pack(side="right", fill="y")
        x_scroll.pack(side="bottom", fill="x")
        self.excel_console.pack(side="left", fill="both", expand=True)

        self.excel_console.insert(
            tk.END, "Welcome to Data Scientist Excel Tools.\nClick 'Load Excel / CSV' to inspect, analyze, and clean spreadsheets."
        )

    def excel_load_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel & CSV Files", "*.xlsx *.xls *.csv")])
        if not filepath:
            return
        try:
            if HAS_PANDAS:
                self.df = pd.read_csv(filepath) if filepath.lower().endswith(".csv") else pd.read_excel(filepath)
                self.excel_console.delete("1.0", tk.END)
                self.excel_console.insert(
                    tk.END,
                    f"File Loaded: {filepath}\nDimensions: {self.df.shape[0]} Rows x {self.df.shape[1]} Columns\n"
                    f"{'=' * 70}\nDATA PREVIEW (Top 20 Rows):\n\n{self.df.head(20).to_string()}",
                )
            else:
                if not filepath.lower().endswith(".csv"):
                    messagebox.showinfo("Library Required", "Install 'pandas' and 'openpyxl' for .xlsx files, or load a .csv file.")
                    return
                with open(filepath, "r", newline="", encoding="utf-8") as f:
                    reader = list(csv.reader(f))
                self.csv_headers = reader[0] if reader else []
                self.csv_rows = reader[1:] if len(reader) > 1 else []
                self.excel_console.delete("1.0", tk.END)
                preview = "\n".join([", ".join(row) for row in reader[:21]])
                self.excel_console.insert(
                    tk.END, f"CSV Loaded: {filepath}\nRows: {len(self.csv_rows)} | Columns: {len(self.csv_headers)}\n{'=' * 70}\n{preview}"
                )
        except Exception as e:
            messagebox.showerror("File Load Error", f"Could not read file:\n{e}")

    def excel_describe(self):
        if HAS_PANDAS and self.df is not None:
            self.excel_console.delete("1.0", tk.END)
            self.excel_console.insert(tk.END, f"DESCRIPTIVE STATISTICS SUMMARY:\n{'=' * 70}\n\n{self.df.describe(include='all').to_string()}")
        elif not HAS_PANDAS and self.csv_rows:
            self.excel_console.delete("1.0", tk.END)
            self.excel_console.insert(tk.END, f"Total Rows: {len(self.csv_rows)}\nColumns: {', '.join(self.csv_headers)}")
        else:
            messagebox.showwarning("No Dataset", "Please load an Excel or CSV file first.")

    def excel_missing_info(self):
        if HAS_PANDAS and self.df is not None:
            nulls = self.df.isnull().sum()
            dtypes = self.df.dtypes
            report = pd.DataFrame({"Data Type": dtypes, "Missing Values": nulls})
            self.excel_console.delete("1.0", tk.END)
            self.excel_console.insert(
                tk.END, f"COLUMN TYPES & MISSING VALUE AUDIT:\n{'=' * 70}\n\n{report.to_string()}\n\nTotal Missing Cells: {int(nulls.sum())}"
            )
        elif not HAS_PANDAS and self.csv_rows:
            empty_cells = sum(1 for row in self.csv_rows for cell in row if not cell.strip())
            self.excel_console.delete("1.0", tk.END)
            self.excel_console.insert(tk.END, f"Total Empty Cells Found: {empty_cells}")
        else:
            messagebox.showwarning("No Dataset", "Please load an Excel or CSV file first.")

    def excel_clean_export(self):
        if HAS_PANDAS and self.df is not None:
            cleaned_df = self.df.dropna()
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Workbook", "*.xlsx")])
            if save_path:
                try:
                    cleaned_df.to_excel(save_path, index=False)
                    messagebox.showinfo("Export Complete", f"Cleaned file saved!\nRows kept: {len(cleaned_df)}")
                except Exception as e:
                    messagebox.showerror("Export Error", str(e))
        elif not HAS_PANDAS and self.csv_rows:
            cleaned = [r for r in self.csv_rows if all(c.strip() for c in r)]
            save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV File", "*.csv")])
            if save_path:
                with open(save_path, "w", newline="", encoding="utf-8") as f:
                    writer = csv.writer(f)
                    writer.writerow(self.csv_headers)
                    writer.writerows(cleaned)
                messagebox.showinfo("Export Complete", f"Cleaned CSV saved ({len(cleaned)} rows kept).")
        else:
            messagebox.showwarning("No Dataset", "Please load an Excel or CSV file first.")

    # ==================== 6. GST CALCULATOR MODULE ====================
    def setup_gst_calculator(self):
        self.gst_title = tk.Label(self.gst_frame, text="Advanced GST & Tax Calculator", font=("Segoe UI", 12, "bold"))
        self.gst_title.pack(pady=15)

        self.gst_form = tk.Frame(self.gst_frame)
        self.gst_form.pack(pady=5)

        self.gst_amt_lbl = tk.Label(self.gst_form, text="Initial Amount ($/₹):", font=("Segoe UI", 10))
        self.gst_amt_lbl.grid(row=0, column=0, sticky="w", pady=8, padx=5)
        self.gst_amt_entry = tk.Entry(self.gst_form, font=("Segoe UI", 10), bd=0)
        self.gst_amt_entry.grid(row=0, column=1, pady=8, padx=5, ipady=4)

        self.gst_rate_lbl = tk.Label(self.gst_form, text="GST Rate (%):", font=("Segoe UI", 10))
        self.gst_rate_lbl.grid(row=1, column=0, sticky="w", pady=8, padx=5)
        self.gst_rate_entry = tk.Entry(self.gst_form, font=("Segoe UI", 10), bd=0)
        self.gst_rate_entry.grid(row=1, column=1, pady=8, padx=5, ipady=4)
        self.gst_rate_entry.insert(0, "18")

        self.gst_btn_box = tk.Frame(self.gst_frame)
        self.gst_btn_box.pack(pady=10)

        self.gst_add_btn = tk.Button(
            self.gst_btn_box, text="Add GST (Exclusive)", font=("Segoe UI", 9, "bold"),
            relief="flat", command=lambda: self.calc_gst("add"), cursor="hand2"
        )
        self.gst_add_btn.pack(side="left", padx=5, ipadx=5, ipady=5)

        self.gst_rem_btn = tk.Button(
            self.gst_btn_box, text="Remove GST (Inclusive)", font=("Segoe UI", 9, "bold"),
            relief="flat", command=lambda: self.calc_gst("remove"), cursor="hand2"
        )
        self.gst_rem_btn.pack(side="left", padx=5, ipadx=5, ipady=5)

        self.gst_result_lbl = tk.Label(self.gst_frame, text="", font=("Segoe UI", 10), justify="left")
        self.gst_result_lbl.pack(pady=15)

    def calc_gst(self, mode):
        try:
            amt = float(self.gst_amt_entry.get())
            rate = float(self.gst_rate_entry.get())
            if mode == "add":
                gst_amount = amt * (rate / 100)
                total_amount = amt + gst_amount
                self.gst_result_lbl.config(text=f"Original Amount: {amt:.2f}\nGST Amount ({rate}%): {gst_amount:.2f}\nTotal Gross Price: {total_amount:.2f}")
            else:
                original = amt / (1 + (rate / 100))
                gst_amount = amt - original
                self.gst_result_lbl.config(text=f"Net Original Amount: {original:.2f}\nGST Amount ({rate}%): {gst_amount:.2f}\nTotal Gross Price: {amt:.2f}")
        except ValueError:
            messagebox.showerror("Input Error", "Please provide valid numbers for amount and GST rate.")

    # ==================== 7. FINANCE MODULE (SI & CI) ====================
    def setup_finance_calculator(self):
        self.fin_title = tk.Label(self.finance_frame, text="Simple & Compound Interest Tool", font=("Segoe UI", 12, "bold"))
        self.fin_title.pack(pady=15)

        self.fin_form = tk.Frame(self.finance_frame)
        self.fin_form.pack(pady=5)

        self.fin_p_lbl = tk.Label(self.fin_form, text="Principal ($/₹):", font=("Segoe UI", 10))
        self.fin_p_lbl.grid(row=0, column=0, sticky="w", pady=6, padx=5)
        self.fin_p = tk.Entry(self.fin_form, font=("Segoe UI", 10), bd=0)
        self.fin_p.grid(row=0, column=1, pady=6, padx=5, ipady=4)

        self.fin_r_lbl = tk.Label(self.fin_form, text="Rate of Interest (% per yr):", font=("Segoe UI", 10))
        self.fin_r_lbl.grid(row=1, column=0, sticky="w", pady=6, padx=5)
        self.fin_r = tk.Entry(self.fin_form, font=("Segoe UI", 10), bd=0)
        self.fin_r.grid(row=1, column=1, pady=6, padx=5, ipady=4)

        self.fin_t_lbl = tk.Label(self.fin_form, text="Time Period (Years):", font=("Segoe UI", 10))
        self.fin_t_lbl.grid(row=2, column=0, sticky="w", pady=6, padx=5)
        self.fin_t = tk.Entry(self.fin_form, font=("Segoe UI", 10), bd=0)
        self.fin_t.grid(row=2, column=1, pady=6, padx=5, ipady=4)

        self.fin_calc_btn = tk.Button(
            self.finance_frame, text="Calculate Interest", font=("Segoe UI", 10, "bold"),
            relief="flat", command=self.calculate_finance, cursor="hand2"
        )
        self.fin_calc_btn.pack(pady=12, ipadx=10, ipady=5)

        self.fin_result_lbl = tk.Label(self.finance_frame, text="", font=("Segoe UI", 10, "bold"), justify="left")
        self.fin_result_lbl.pack(pady=10)

    def calculate_finance(self):
        try:
            p = float(self.fin_p.get())
            r = float(self.fin_r.get())
            t = float(self.fin_t.get())
            si = (p * r * t) / 100
            ci = p * (math.pow((1 + r / 100), t)) - p
            self.fin_result_lbl.config(text=f"Simple Interest (SI): {si:.2f} | Total: {p + si:.2f}\nCompound Interest (CI): {ci:.2f} | Total: {p + ci:.2f}")
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numerical values for principal, rate, and time.")

    # ==================== 8. CONVERTER MODULE ====================
    def setup_converter(self):
        self.conv_title = tk.Label(self.converter_frame, text="Data Storage & Unit Converter", font=("Segoe UI", 12, "bold"))
        self.conv_title.pack(pady=15)

        self.conv_form = tk.Frame(self.converter_frame)
        self.conv_form.pack(pady=5)

        self.byte_lbl = tk.Label(self.conv_form, text="Bytes Value:", font=("Segoe UI", 10))
        self.byte_lbl.grid(row=0, column=0, sticky="w", pady=10, padx=5)
        self.byte_entry = tk.Entry(self.conv_form, font=("Segoe UI", 10), bd=0)
        self.byte_entry.grid(row=0, column=1, pady=10, padx=5, ipady=4)

        self.conv_btn = tk.Button(
            self.conv_form, text="Convert Data Units", font=("Segoe UI", 9, "bold"),
            relief="flat", command=self.convert_bytes, cursor="hand2"
        )
        self.conv_btn.grid(row=0, column=2, padx=10)

        self.conv_result_lbl = tk.Label(self.converter_frame, text="", font=("Segoe UI", 10), justify="left")
        self.conv_result_lbl.pack(pady=20)

    def convert_bytes(self):
        try:
            b = float(self.byte_entry.get())
            kb = b / 1024
            mb = kb / 1024
            gb = mb / 1024
            self.conv_result_lbl.config(text=f"Kilobytes (KB): {kb:.4f}\nMegabytes (MB): {mb:.4f}\nGigabytes (GB): {gb:.4f}")
        except ValueError:
            messagebox.showerror("Input Error", "Please provide a valid number of bytes.")


if __name__ == "__main__":
    app = AdvancedSuperCalculator()
    app.mainloop()
    