import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import os
import threading
from datetime import datetime
import mysql.connector
import fitz  # PyMuPDF

CONFIG_FILE = "shimadzu_machine_config.txt"
DB_CONFIG_FILE = "shimadzu_database_config.txt"

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")


class ShimadzuPDFApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Shimadzu LC-2050")
        self.geometry("1200x750+180+30")
        self.test_code_locked = False

        # extraction mode variable ("single" or "multiple")
        self.mode_var = ctk.StringVar(value="single")

        self.create_main_interface()
        self.load_config()

    def create_main_interface(self):
        frame = ctk.CTkFrame(self)
        frame.pack(padx=20, pady=20, fill="both", expand=True)

        # Mode radio buttons
        mode_frame = ctk.CTkFrame(frame)
        mode_frame.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(mode_frame, text="Extraction Mode:").pack(side="left", padx=(0, 8))
        ctk.CTkRadioButton(mode_frame, text="Single Compound", variable=self.mode_var, value="single",
                           command=self._on_mode_change).pack(side="left", padx=4)
        ctk.CTkRadioButton(mode_frame, text="Multiple Compound", variable=self.mode_var, value="multiple",
                           command=self._on_mode_change).pack(side="left", padx=4)

        ctk.CTkLabel(frame, text="Machine ID:").pack(anchor="w")
        self.machine_id_entry = ctk.CTkEntry(frame, state="disabled")
        self.machine_id_entry.pack(fill="x", pady=5)

        ctk.CTkLabel(frame, text="Sample ID (u_id):").pack(anchor="w")
        self.sample_id_entry = ctk.CTkEntry(frame)
        self.sample_id_entry.pack(fill="x", pady=5)

        ctk.CTkLabel(frame, text="User ID:").pack(anchor="w")
        self.user_id_entry = ctk.CTkEntry(frame)
        self.user_id_entry.pack(fill="x", pady=5)

        ctk.CTkLabel(frame, text="Test Code:").pack(anchor="w")
        self.test_code_entry = ctk.CTkEntry(frame)
        self.test_code_entry.pack(fill="x", pady=5)

        btn_frame = ctk.CTkFrame(frame)
        btn_frame.pack(fill="x", pady=(8, 0))
        ctk.CTkButton(btn_frame, text="Save Config", command=self.save_config).pack(side="left", padx=6)
        ctk.CTkButton(btn_frame, text="Database Settings", command=self.open_db_config).pack(side="left", padx=6)
        ctk.CTkButton(btn_frame, text="Select PDF(s) and Extract", command=self.select_pdfs).pack(side="left", padx=6)
        ctk.CTkButton(btn_frame, text="License", command=self.show_license).pack(side="left", padx=6)

        ctk.CTkLabel(frame, text="Status Log:").pack(anchor="w", pady=(10, 0))
        self.status_box = ctk.CTkTextbox(frame, height=100)
        self.status_box.pack(fill="x", expand=False)

        ctk.CTkLabel(frame, text="Database Inserted Rows:").pack(anchor="w", pady=(10, 0))
        self.tree_container = ctk.CTkFrame(frame)
        self.tree_container.pack(fill="both", expand=True)

        # Initialize treeview according to default mode
        self._build_treeview()

    def _on_mode_change(self):
        """Called when radio button changes — rebuild treeview to match selected mode."""
        self.log_status(f"Mode changed to: {self.mode_var.get()}")
        self._build_treeview()

    def _build_treeview(self):
        # clear previous widgets in container
        for w in self.tree_container.winfo_children():
            w.destroy()

        if self.mode_var.get() == "single":
            columns = [
                "machine_id", "u_id", "user_id", "test_code", "acquired_by", "sample_name_header", "sample_id",
                "tray", "vial", "injection_volume", "data_file", "method_file",
                "batch_file", "report_format_file", "date_acquired", "date_processed",
                "title", "sample_name", "sample_id_ind", "ret_time", "area",
                "tailing_factor", "theoretical_plate"
            ]
        else:  # multiple
            columns = [
                "machine_id", "u_id", "user_id", "test_code", "acquired_by", "sample_name_header", "sample_id",
                "tray", "vial", "injection_volume", "data_file", "method_file",
                "batch_file", "report_format_file", "date_acquired", "date_processed",
                "compound_name",  # new column for multiple
                "title", "sample_name", "sample_id_ind", "ret_time", "area",
                "tailing_factor", "theoretical_plate"
            ]

        self.current_columns = columns
        self.tree = ttk.Treeview(self.tree_container, columns=columns, show='headings')
        for col in columns:
            self.tree.heading(col, text=col)
            # set a reasonable width for compound_name to be visible
            width = 150 if col != "compound_name" else 180
            self.tree.column(col, width=width, anchor="center")

        y_scroll = ttk.Scrollbar(self.tree_container, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(self.tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.tree_container.rowconfigure(0, weight=1)
        self.tree_container.columnconfigure(0, weight=1)

        style = ttk.Style()
        style.configure("Treeview", background="black", foreground="light blue",
                        fieldbackground="black", font=("Helvetica", 10, "bold"))
        style.configure("Treeview.Heading", background="blue", foreground="red",
                        font=("Helvetica", 10, "bold"))
        style.map("Treeview", background=[("selected", "#2a2d2e")])

    def select_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        self.log_status(f"File dialog returned: {files}")  # debug
        if files:
            if not self.test_code_locked:
                test_code = self.test_code_entry.get().strip()
                if test_code:
                    self.save_test_code(test_code)
                    self.test_code_entry.configure(state="disabled")
                    self.test_code_locked = True
                else:
                    messagebox.showwarning("Input Required", "Enter test code before proceeding.")
                    return

            sample_id = self.sample_id_entry.get().strip()
            user_id = self.user_id_entry.get().strip()
            if not sample_id or not user_id:
                messagebox.showwarning("Input Required", "Enter both sample ID (u_id) and User ID.")
                return

            import fitz
            self.log_status(f"PyMuPDF version: {fitz.__doc__}")

            # Process files in current thread — could be moved to a background thread if desired
            for filepath in files:
                filename = os.path.basename(filepath)
                self.log_status(f"About to open: {filepath}")  # debug
                try:
                    with open(filepath, 'rb') as f:
                        data = f.read()
                        self.log_status(f"Read {len(data)} bytes from {filename}")  # debug
                        doc = fitz.open(stream=data, filetype='pdf')
                        lines = [line.strip() for page in doc for line in page.get_text().splitlines() if line.strip()]

                        self.log_status(f"Extracted {len(lines)} lines from {filename}")

                        if len(lines) > 0:
                            if self.mode_var.get() == "single":
                                self.extract_single(lines, self.machine_id_entry.get().strip(), sample_id, user_id)
                            else:
                                self.extract_multiple(lines, self.machine_id_entry.get().strip(), sample_id, user_id)
                        else:
                            self.log_status(f"No text extracted from {filename} (maybe image-based PDF?)")
                except Exception as e:
                    self.log_status(f"Error opening {filename}: {e}")

    # --------------------- shared helpers ---------------------
    def save_test_code(self, code):
        lines = []
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                lines = f.read().splitlines()
        if len(lines) < 3:
            lines += [""] * (3 - len(lines))
        lines[2] = code
        with open(CONFIG_FILE, "w") as f:
            f.write("\n".join(lines) + "\n")
        self.log_status("Test code saved and locked.")

    def save_config(self):
        machine_id = self.machine_id_entry.get().strip()
        test_code = self.test_code_entry.get().strip()
        with open(CONFIG_FILE, "w") as f:
            f.write(machine_id + "\n")
            f.write("dummy_path\n")
            f.write(test_code + "\n")
        self.log_status("Configuration saved.")

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                lines = f.read().splitlines()
                if len(lines) >= 1:
                    self.machine_id_entry.configure(state="normal")
                    self.machine_id_entry.insert(0, lines[0])
                    self.machine_id_entry.configure(state="disabled")
                if len(lines) >= 3:
                    self.test_code_entry.insert(0, lines[2])
                    self.test_code_entry.configure(state="disabled")
                    self.test_code_locked = True

    def open_db_config(self):
        win = ctk.CTkToplevel(self)
        win.title("Database Configuration")
        win.geometry("400x400")
        entries = {}
        labels = ["Host", "Port", "User", "Password", "Database"]

        for label in labels:
            ctk.CTkLabel(win, text=label).pack()
            ent = ctk.CTkEntry(win, show="*" if label == "Password" else None)
            ent.pack(fill="x", padx=10, pady=5)
            entries[label] = ent

        if os.path.exists(DB_CONFIG_FILE):
            with open(DB_CONFIG_FILE, "r") as f:
                for label, line in zip(labels, f.read().splitlines()):
                    entries[label].insert(0, line)

        def save():
            with open(DB_CONFIG_FILE, "w") as f:
                for label in labels:
                    f.write(entries[label].get().strip() + "\n")
            messagebox.showinfo("Saved", "Database configuration saved.")

        ctk.CTkButton(win, text="Save", command=save).pack(pady=10)

    # --------------------- single-compound extraction (your original) ---------------------
    def extract_single(self, lines, machine_id, sample_id, user_id):
        def get_val(label):
            try:
                idx = lines.index(label)
                return lines[idx + 1].lstrip(": ").strip()
            except:
                return ""

        header = {
            "machine_id": machine_id,
            "u_id": sample_id,
            "user_id": user_id,
            "test_code": self.test_code_entry.get().strip(),
            "acquired_by": get_val("Acquired by"),
            "sample_name_header": get_val("Sample Name"),
            "sample_id": get_val("Sample ID"),
            "tray": int(get_val("Tray#") or 0),
            "vial": int(get_val("Vial#") or 0),
            "injection_volume": float(get_val("Injection Volume") or 0),
            "data_file": get_val("Data File"),
            "method_file": get_val("Method File"),
            "batch_file": get_val("Batch File"),
            "report_format_file": get_val("Report Format File"),
            "date_acquired": get_val("Date Acquired"),
            "date_processed": get_val("Date Processed")
        }

        def get_table_section(label):
            try:
                idx = [i for i, val in enumerate(lines) if val.strip().lower() == label.lower()][-1]
                result = []
                for i in range(idx + 1, len(lines)):
                    if lines[i] in ["Title", "Sample Name", "Sample ID", "Ret. Time", "Area", "Tailing Factor", "Theoretical Plate"]:
                        break
                    if lines[i].strip() and not lines[i].startswith(":"):
                        result.append(lines[i].lstrip(": ").strip())
                return result
            except:
                return []

        titles = get_table_section("Title")
        sample_names = get_table_section("Sample Name")
        sample_ids = get_table_section("Sample ID")
        ret_times = get_table_section("Ret. Time")
        areas = get_table_section("Area")
        tailing_factors = get_table_section("Tailing Factor")
        plates = get_table_section("Theoretical Plate")
        if not plates:
            plates = get_table_section("Number of Theoretical Plate(USP)")

        rows = []
        row_count = max(len(titles), len(ret_times))
        for i in range(row_count):
            try:
                row = {
                    **header,
                    "title": titles[i] if i < len(titles) else None,
                    "sample_name": sample_names[i] if i < len(sample_names) else None,
                    "sample_id_ind": sample_ids[i] if i < len(sample_ids) else None,
                    "ret_time": float(ret_times[i]) if i < len(ret_times) else None,
                    "area": float(areas[i]) if i < len(areas) else None,
                    "tailing_factor": float(tailing_factors[i]) if i < len(tailing_factors) else None,
                    "theoretical_plate": float(plates[i]) if i < len(plates) else None
                }
                rows.append(row)
            except Exception as e:
                self.log_status(f"Row {i + 1} skipped: {e}")

        if rows:
            self.insert_single_db(rows)
        else:
            self.log_status("No rows extracted for single-compound file.")

    def insert_single_db(self, rows):
        try:
            with open(DB_CONFIG_FILE, "r") as f:
                host, port, user, pwd, db = f.read().splitlines()
            conn = mysql.connector.connect(host=host, port=port, user=user, password=pwd, database=db)
            cursor = conn.cursor()

            columns_to_insert = [
                "machine_id", "u_id", "user_id", "test_code", "acquired_by", "sample_name_header", "sample_id",
                "tray", "vial", "injection_volume", "data_file", "method_file",
                "batch_file", "report_format_file", "date_acquired", "date_processed",
                "title", "sample_name", "sample_id_ind", "ret_time", "area",
                "tailing_factor", "theoretical_plate"
            ]

            for row in rows:
                values = tuple(row.get(col, None) for col in columns_to_insert)
                cursor.execute(f"""
                    INSERT INTO shimadzu_lc2050_results (
                        {', '.join(columns_to_insert)}
                    ) VALUES ({', '.join(['%s'] * len(columns_to_insert))})
                """, values)

                # Insert into treeview (ensure string safe)
                disp_vals = tuple("" if v is None else v for v in values)
                self.tree.insert("", "end", values=disp_vals)

                now = datetime.now()
                log_file = f"log_{now.strftime('%Y-%m-%d')}.txt"
                with open(log_file, "a") as f:
                    f.write(f"[{now.strftime('%Y-%m-%d %H:%M:%S')}] Inserted Row (single):\n")
                    for col, val in zip(columns_to_insert, values):
                        f.write(f"    {col}: {val}\n")
                    f.write("\n")

            conn.commit()
            conn.close()
            self.log_status(f"Inserted {len(rows)} rows into shimadzu_lc2050_results.")
        except Exception as e:
            self.log_status(f"DB error (single): {e}")

    # --------------------- multiple-compound extraction (your enhanced version) ---------------------
    def _find_compound_starts(self, lines):
        starts = []
        for i, l in enumerate(lines):
            if "compound name" in l.lower():
                parts = l.split(":", 1)
                name = parts[1].strip() if len(parts) > 1 else ""
                starts.append((i, name))
        return starts

    def _sub_get_table_section(self, sublines, label):
        try:
            lc = [s.strip().lower() for s in sublines]
            label_l = label.lower()
            idx = len(lc) - 1 - lc[::-1].index(label_l)  # last occurrence index
            result = []
            for i in range(idx + 1, len(sublines)):
                val = sublines[i]
                if val.strip() in ["Title", "Sample Name", "Sample ID", "Ret. Time", "Area", "Tailing Factor",
                                   "Theoretical Plate", "Number of Theoretical Plate(USP)"]:
                    break
                if val.strip() and not val.startswith(":"):
                    result.append(val.lstrip(": ").strip())
            return result
        except Exception:
            return []

    def extract_multiple(self, lines, machine_id, sample_id, user_id):
        def parse_float(val):
            try:
                return float(val)
            except:
                return None

        def get_val(label):
            try:
                idx = lines.index(label)
                return lines[idx + 1].lstrip(": ").strip()
            except:
                return ""

        header_common = {
            "machine_id": machine_id,
            "u_id": sample_id,
            "user_id": user_id,
            "test_code": self.test_code_entry.get().strip(),
            "acquired_by": get_val("Acquired by"),
            "sample_name_header": get_val("Sample Name"),
            "sample_id": get_val("Sample ID"),
            "tray": int(get_val("Tray#") or 0),
            "vial": int(get_val("Vial#") or 0),
            "injection_volume": float(get_val("Injection Volume") or 0),
            "data_file": get_val("Data File"),
            "method_file": get_val("Method File"),
            "batch_file": get_val("Batch File"),
            "report_format_file": get_val("Report Format File"),
            "date_acquired": get_val("Date Acquired"),
            "date_processed": get_val("Date Processed")
        }

        compound_starts = self._find_compound_starts(lines)
        if not compound_starts:
            # fallback to single-table parse (reuse some logic)
            self.log_status("No compound headers found — falling back to single-table parse.")
            titles = self._sub_get_table_section(lines, "Title")
            sample_names = self._sub_get_table_section(lines, "Sample Name")
            sample_ids = self._sub_get_table_section(lines, "Sample ID")
            ret_times = self._sub_get_table_section(lines, "Ret. Time")
            areas = self._sub_get_table_section(lines, "Area")
            tailing_factors = self._sub_get_table_section(lines, "Tailing Factor")
            plates = self._sub_get_table_section(lines, "Theoretical Plate") or self._sub_get_table_section(lines,
                                                                                                            "Number of Theoretical Plate(USP)")

            rows = []
            row_count = max(len(titles), len(ret_times))
            for i in range(row_count):
                row = {
                    **header_common,
                    "compound_name": header_common.get("sample_name_header", ""),
                    "title": titles[i] if i < len(titles) else None,
                    "sample_name": sample_names[i] if i < len(sample_names) else None,
                    "sample_id_ind": sample_ids[i] if i < len(sample_ids) else None,
                    "ret_time": parse_float(ret_times[i]) if i < len(ret_times) else None,
                    "area": parse_float(areas[i]) if i < len(areas) else None,
                    "tailing_factor": parse_float(tailing_factors[i]) if i < len(tailing_factors) else None,
                    "theoretical_plate": parse_float(plates[i]) if i < len(plates) else None
                }
                rows.append(row)

            if rows:
                self.insert_multi_db(rows)
            else:
                self.log_status("No rows extracted in fallback parse.")
            return

        rows_all = []
        for idx, (start_index, compound_name) in enumerate(compound_starts):
            end_index = len(lines)
            if idx + 1 < len(compound_starts):
                end_index = compound_starts[idx + 1][0]

            # locate 'Title' header inside this block
            title_idx_in_block = None
            for i in range(start_index, end_index):
                if lines[i].strip() == "Title":
                    title_idx_in_block = i
                    break
            if title_idx_in_block is None:
                for i in range(end_index - 1, start_index - 1, -1):
                    if lines[i].strip() == "Title":
                        title_idx_in_block = i
                        break
            if title_idx_in_block is None:
                self.log_status(f"Compound '{compound_name}' at line {start_index} has no 'Title' table header — skipping.")
                continue

            sublines = lines[title_idx_in_block:end_index]
            titles = self._sub_get_table_section(sublines, "Title")
            sample_names = self._sub_get_table_section(sublines, "Sample Name")
            sample_ids = self._sub_get_table_section(sublines, "Sample ID")
            ret_times = self._sub_get_table_section(sublines, "Ret. Time")
            areas = self._sub_get_table_section(sublines, "Area")
            tailing_factors = self._sub_get_table_section(sublines, "Tailing Factor")
            plates = self._sub_get_table_section(sublines, "Theoretical Plate") or self._sub_get_table_section(sublines,
                                                                                                               "Number of Theoretical Plate(USP)")

            row_count = max(len(titles), len(ret_times), len(areas), len(sample_ids), len(sample_names))

            for i in range(row_count):
                r = {
                    **header_common,
                    "compound_name": compound_name or "",
                    "title": titles[i] if i < len(titles) else None,
                    "sample_name": sample_names[i] if i < len(sample_names) else None,
                    "sample_id_ind": sample_ids[i] if i < len(sample_ids) else None,
                    "ret_time": parse_float(ret_times[i]) if i < len(ret_times) else None,
                    "area": parse_float(areas[i]) if i < len(areas) else None,
                    "tailing_factor": parse_float(tailing_factors[i]) if i < len(tailing_factors) else None,
                    "theoretical_plate": parse_float(plates[i]) if i < len(plates) else None
                }
                rows_all.append(r)

        if rows_all:
            self.insert_multi_db(rows_all)
        else:
            self.log_status("No rows extracted for any compound blocks.")

    def insert_multi_db(self, rows):
        try:
            with open(DB_CONFIG_FILE, "r") as f:
                host, port, user, pwd, db = f.read().splitlines()
            conn = mysql.connector.connect(host=host, port=port, user=user, password=pwd, database=db)
            cursor = conn.cursor()

            columns_to_insert = [
                "machine_id", "u_id", "user_id", "test_code", "acquired_by", "sample_name_header", "sample_id",
                "tray", "vial", "injection_volume", "data_file", "method_file",
                "batch_file", "report_format_file", "date_acquired", "date_processed",
                "compound_name",  # new
                "title", "sample_name", "sample_id_ind", "ret_time", "area",
                "tailing_factor", "theoretical_plate"
            ]

            for row in rows:
                values = tuple(row.get(col, None) for col in columns_to_insert)
                cursor.execute(f"""
                    INSERT INTO shimadzu_lc2050_multicom_raw (
                        {', '.join(columns_to_insert)}
                    ) VALUES ({', '.join(['%s'] * len(columns_to_insert))})
                """, values)

                disp_vals = tuple("" if v is None else v for v in values)
                self.tree.insert("", "end", values=disp_vals)

                now = datetime.now()
                log_file = f"log_{now.strftime('%Y-%m-%d')}.txt"
                with open(log_file, "a") as f:
                    f.write(f"[{now.strftime('%Y-%m-%d %H:%M:%S')}] Inserted Row (multi):\n")
                    for col, val in zip(columns_to_insert, values):
                        f.write(f"    {col}: {val}\n")
                    f.write("\n")

            conn.commit()
            conn.close()
            self.log_status(f"Inserted {len(rows)} rows into shimadzu_lc2050_multicomponent.")
        except Exception as e:
            self.log_status(f"DB error (multi): {e}")

    # --------------------- logging & license ---------------------
    def log_status(self, msg):
        now = datetime.now()
        timestamp = now.strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {msg}"
        try:
            self.status_box.insert("end", log_message + "\n")
            self.status_box.see("end")
        except Exception:
            pass

        log_file = f"log_{now.strftime('%Y-%m-%d')}.txt"
        try:
            with open(log_file, "a") as f:
                f.write(log_message + "\n")
        except Exception:
            pass

    def show_license(self):
        license_win = ctk.CTkToplevel(self)
        license_win.title("License Information")
        license_win.geometry("700x500+200+50")
        license_win.resizable(False, False)
        license_win.grab_set()
        license_win.attributes("-topmost", True)

        scroll_frame = ctk.CTkScrollableFrame(license_win, width=680, height=440, corner_radius=10, fg_color="#1a1a2e")
        scroll_frame.pack(padx=5, pady=5, fill="both", expand=True)

        license_lines = [
            ("🌟 ©️ Copyright Metro Solution Limited", "#FFD700"),
            ("🔒 All Rights Reserved", "#00CED1"),
            ("💻 Developed and owned by Metro Solution Limited", "#ADFF2F"),
            ("🌐 www.metrosolutionltd.com", "#1E90FF"),
            ("📞 Contact: +8801819173762", "#32CD32"),
            ("⚙️ Scientific data capturing and automation", "white"),
            ("🚫 Unauthorized use is strictly prohibited", "#FF6347"),
            ("🕓 License Duration: 🟢 Lifetime License", "#00FA9A"),
            ("This license grants the authorized user continuous access to the software without expiration, provided all terms and conditions of Metro Solution Limited are followed.", "white"),
            ("🔑 License Key: Metro Solution Limited – 001199228833774466", "#FFD700", "bold"),
            ("Each license key is uniquely created and verified by Metro Solution Limited.", "white"),
            ("Keep your license key confidential — unauthorized sharing or duplication is strictly prohibited.", "white"),
            ("✅ By using this software, you agree to all licensing terms and conditions set forth by Metro Solution Limited.", "#32CD32"),
            ("Developed & Licensed by:", "white"),
            ("Metro Solution Limited", "#00CED1"),
            ("✨ Empowering Innovation Through Automation", "#ADFF2F")
        ]

        for item in license_lines:
            text = item[0]
            color = item[1]
            font_style = ("Helvetica", 11)
            if len(item) == 3 and item[2] == "bold":
                font_style = ("Helvetica", 11, "bold")
            lbl = ctk.CTkLabel(scroll_frame, text=text, text_color=color,
                               font=font_style, anchor="w", wraplength=660)
            lbl.pack(anchor="w", pady=0)

        close_btn = ctk.CTkButton(license_win, text="Close", command=license_win.destroy,
                                  fg_color="#ff5555", hover_color="#ff7777", height=30)
        close_btn.pack(pady=5)


if __name__ == "__main__":
    app = ShimadzuPDFApp()
    app.mainloop()