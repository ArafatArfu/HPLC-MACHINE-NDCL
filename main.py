import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import os
import threading
from datetime import datetime
import mysql.connector
import fitz  # PyMuPDF

CONFIG_FILE = "shimadzu_machine_config.txt"
DB_CONFIG_FILE = "shimadzu_database_config.txt"

# --- UI SETTINGS ---
ctk.set_appearance_mode("Light")  
ctk.set_default_color_theme("blue")

class ShimadzuPDFApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Shimadzu LC-2050 Data Manager")
        self.geometry("1280x800+50+30")
        self.test_code_locked = False

        # --- Tab Setup ---
        self.tabview = ctk.CTkTabview(self, width=1200, height=750, command=self.tabview_callback)
        self.tabview.pack(padx=10, pady=10, fill="both", expand=True)

        self.tab_general = self.tabview.add("Assay")
        self.tab_disso = self.tabview.add("Dissolution")

        # --- Tab 1 Variables ---
        self.mode_var = ctk.StringVar(value="single")

        # --- Initialize Tabs ---
        self.create_main_interface()      # Tab 1 UI (Assay)
        self.setup_dissolution_tab()      # Tab 2 UI (Dissolution)

        self.load_config() 
        self.init_db_tables() 

        # --- Tab Animation Tracker ---
        self.current_tab_name_tracker = "Assay"

    # =========================================================================
    # TAB 1: ASSAY (General Extraction)
    # =========================================================================
    def create_main_interface(self):
        main_frame = ctk.CTkFrame(self.tab_general, fg_color="#C2E2FA") 
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # --- 1. Settings Area ---
        settings_frame = ctk.CTkFrame(main_frame, fg_color="transparent") 
        settings_frame.pack(fill="x", pady=(0, 5))
        
        ctk.CTkLabel(settings_frame, text="Extraction Mode:", font=("Arial", 12, "bold")).pack(side="left", padx=10, pady=5)
        ctk.CTkRadioButton(settings_frame, text="Single Compound", variable=self.mode_var, value="single", 
                           command=self._on_mode_change).pack(side="left", padx=10)
        ctk.CTkRadioButton(settings_frame, text="Multiple Compound", variable=self.mode_var, value="multiple",
                           command=self._on_mode_change).pack(side="left", padx=10)

        # --- 2. Input Fields Area ---
        input_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        input_frame.pack(fill="x", pady=5)
        
        for i in range(4): input_frame.columnconfigure(i, weight=1)

        # Labels
        ctk.CTkLabel(input_frame, text="Machine ID", font=("Arial", 11)).grid(row=0, column=0, sticky="w", padx=10)
        ctk.CTkLabel(input_frame, text="Test Code (Select)", font=("Arial", 11)).grid(row=0, column=1, sticky="w", padx=10)
        ctk.CTkLabel(input_frame, text="Sample ID (u_id)", font=("Arial", 11)).grid(row=0, column=2, sticky="w", padx=10)
        ctk.CTkLabel(input_frame, text="User ID", font=("Arial", 11)).grid(row=0, column=3, sticky="w", padx=10)

        # Entries
        self.machine_id_entry = ctk.CTkEntry(input_frame, state="disabled", height=30)
        self.machine_id_entry.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))

        # Test Code Dropdown for Assay (10003 - 10026)
        assay_test_codes = [str(i) for i in range(10003, 10027)] 
        self.test_code_entry = ctk.CTkOptionMenu(input_frame, height=30, values=assay_test_codes)
        self.test_code_entry.set("10013") 
        self.test_code_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=(0, 10))

        self.sample_id_entry = ctk.CTkEntry(input_frame, height=30, placeholder_text="Scan Sample ID")
        self.sample_id_entry.grid(row=1, column=2, sticky="ew", padx=10, pady=(0, 10))

        self.user_id_entry = ctk.CTkEntry(input_frame, height=30, placeholder_text="User ID")
        self.user_id_entry.grid(row=1, column=3, sticky="ew", padx=10, pady=(0, 10))

        # --- 3. Buttons Area ---
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=10)

        ctk.CTkButton(btn_frame, text="Save Config", command=self.save_config, fg_color="#3B8ED0", width=100).pack(side="left", padx=(0, 5))
        ctk.CTkButton(btn_frame, text="DB Settings", command=self.open_db_config, fg_color="#607D8B", width=100).pack(side="left", padx=5)
        
        ctk.CTkButton(btn_frame, text="Select & Process", command=self.select_pdfs, 
                      fg_color="#2CC985", hover_color="#229A65", font=("Arial", 13, "bold"), height=35).pack(side="left", padx=20, fill="x", expand=True)
        
        ctk.CTkButton(btn_frame, text="License Info", command=self.show_license, fg_color="#E57373", hover_color="#D32F2F", width=100).pack(side="right")

        # --- 4. Logs and Treeview ---
        log_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        log_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(log_frame, text="Status Log:", font=("Arial", 11, "bold")).pack(anchor="w", padx=5)
        self.status_box = ctk.CTkTextbox(log_frame, height=60, font=("Consolas", 11))
        self.status_box.pack(fill="x", padx=5, pady=(0,5))

        tree_label = ctk.CTkLabel(main_frame, text="Extracted Data Preview:", font=("Arial", 12, "bold"))
        tree_label.pack(anchor="w", pady=(5, 0))
        
        self.tree_container = ctk.CTkFrame(main_frame)
        self.tree_container.pack(fill="both", expand=True, pady=5)
        self._build_treeview()

    def _on_mode_change(self):
        self.log_status(f"Mode changed to: {self.mode_var.get()}")
        self._build_treeview()

    def _build_treeview(self):
        for w in self.tree_container.winfo_children(): w.destroy()
        if self.mode_var.get() == "single":
            cols = ["machine_id", "u_id", "user_id", "test_code", "acquired_by", "sample_name_header", "sample_id",
                    "tray", "vial", "injection_volume", "data_file", "method_file", "batch_file", "report_format_file",
                    "date_acquired", "date_processed", "title", "sample_name", "sample_id_ind", "ret_time", "area",
                    "tailing_factor", "theoretical_plate"]
        else:
            cols = ["machine_id", "u_id", "user_id", "test_code", "acquired_by", "sample_name_header", "sample_id",
                    "tray", "vial", "injection_volume", "data_file", "method_file", "batch_file", "report_format_file",
                    "date_acquired", "date_processed", "compound_name", "title", "sample_name", "sample_id_ind", 
                    "ret_time", "area", "tailing_factor", "theoretical_plate"]
        
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="white", foreground="black", fieldbackground="white", rowheight=25)
        style.configure("Treeview.Heading", background="#E0E0E0", foreground="black", font=("Arial", 9, "bold"))
        
        self.tree = ttk.Treeview(self.tree_container, columns=cols, show='headings')
        for col in cols:
            self.tree.heading(col, text=col)
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

    # =========================================================================
    # TAB 2: DISSOLUTION (UPDATED UI & LOGIC)
    # =========================================================================
    def setup_dissolution_tab(self):
        # Main frame with the specified theme color #E1BEE7
        main_frame = ctk.CTkFrame(self.tab_disso, fg_color="#E1BEE7") 
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # --- 1. Info Frame (Inline - 4 Columns) ---
        info_frame = ctk.CTkFrame(main_frame, fg_color="transparent") 
        info_frame.pack(fill="x", pady=5)
        
        for i in range(4): info_frame.columnconfigure(i, weight=1)

        # Headers
        ctk.CTkLabel(info_frame, text="Machine ID", font=("Arial", 11)).grid(row=0, column=0, sticky="w", padx=5)
        ctk.CTkLabel(info_frame, text="Test Code", font=("Arial", 11)).grid(row=0, column=1, sticky="w", padx=5)
        ctk.CTkLabel(info_frame, text="Sample ID (u_id)", font=("Arial", 11)).grid(row=0, column=2, sticky="w", padx=5)
        ctk.CTkLabel(info_frame, text="User ID", font=("Arial", 11)).grid(row=0, column=3, sticky="w", padx=5)

        # Entries
        self.disso_machine_id_entry = ctk.CTkEntry(info_frame, state="disabled", height=30)
        self.disso_machine_id_entry.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 10))

        self.disso_test_code = ctk.CTkOptionMenu(info_frame, height=30, values=["10010", "10011"])
        self.disso_test_code.grid(row=1, column=1, sticky="ew", padx=5, pady=(0, 10))

        self.disso_sample_id = ctk.CTkEntry(info_frame, height=30, placeholder_text="Scan Sample ID")
        self.disso_sample_id.grid(row=1, column=2, sticky="ew", padx=5, pady=(0, 10))

        self.disso_user_id = ctk.CTkEntry(info_frame, height=30, placeholder_text="User ID")
        self.disso_user_id.grid(row=1, column=3, sticky="ew", padx=5, pady=(0, 10))

        # --- 2. Sample Type Selection (Standard/Non-Standard) ---
        type_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        type_frame.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(type_frame, text="Sample Type:", font=("Arial", 11, "bold")).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        
        self.sample_type_var = ctk.StringVar(value="non_standard")
        ctk.CTkRadioButton(type_frame, text="Standard (CS/SS)", variable=self.sample_type_var, 
                          value="standard", command=self.on_sample_type_change).grid(row=0, column=1, padx=10, sticky="w")
        ctk.CTkRadioButton(type_frame, text="Non-Standard", variable=self.sample_type_var, 
                          value="non_standard", command=self.on_sample_type_change).grid(row=0, column=2, padx=10, sticky="w")

        # --- 3. Standard Sample Input (For UI Only - Auto-detected from PDF) ---
        self.std_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        self.std_frame.pack(fill="x", pady=5)
        self.std_frame.columnconfigure((0, 1), weight=1)
        
        ctk.CTkLabel(self.std_frame, text="Standard Type (Will be auto-detected from PDF):", font=("Arial", 11)).grid(row=0, column=0, sticky="w", padx=10)
        self.std_type_var = ctk.StringVar(value="CS")
        std_type_frame = ctk.CTkFrame(self.std_frame, fg_color="transparent")
        std_type_frame.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 10))
        ctk.CTkRadioButton(std_type_frame, text="CS", variable=self.std_type_var, value="CS", state="disabled").pack(side="left", padx=5)
        ctk.CTkRadioButton(std_type_frame, text="SS", variable=self.std_type_var, value="SS", state="disabled").pack(side="left", padx=5)
        
        # Auto detection info
        auto_detect_label = ctk.CTkLabel(self.std_frame, text="PDF থেকে CS/SS auto-detect করা হবে", 
                                         font=("Arial", 10, "italic"), text_color="#FF5722")
        auto_detect_label.grid(row=1, column=1, sticky="e", padx=10, pady=(0, 10))
        
        # Initially hide standard frame
        self.std_frame.pack_forget()

        # --- 4. Logic / Settings Frame (For Non-Standard Files Only) ---
        self.logic_frame = ctk.CTkFrame(main_frame, fg_color="transparent") 
        self.logic_frame.pack(fill="x", pady=5)
        self.logic_frame.columnconfigure((0,1,2,3), weight=1)

        # Col 1: Component (Single / Multi)
        ctk.CTkLabel(self.logic_frame, text="Component Type", font=("Arial", 11, "bold")).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.comp_type_var = ctk.StringVar(value="single")
        ctk.CTkRadioButton(self.logic_frame, text="Single", variable=self.comp_type_var, value="single").grid(row=1, column=0, padx=10, sticky="w")
        ctk.CTkRadioButton(self.logic_frame, text="Multi", variable=self.comp_type_var, value="multi").grid(row=2, column=0, padx=10, sticky="w")

        # Col 2: Release Type (Immediate / Delayed / Extended)
        ctk.CTkLabel(self.logic_frame, text="Release Type", font=("Arial", 11, "bold")).grid(row=0, column=1, padx=10, pady=5, sticky="w")
        self.release_type_var = ctk.StringVar(value="immediate")
        ctk.CTkRadioButton(self.logic_frame, text="Immediate Release", variable=self.release_type_var, 
                          value="immediate", command=self.update_stage_dropdown).grid(row=1, column=1, padx=10, sticky="w")
        ctk.CTkRadioButton(self.logic_frame, text="Delayed Release", variable=self.release_type_var, 
                          value="delayed", command=self.update_stage_dropdown).grid(row=2, column=1, padx=10, sticky="w")
        ctk.CTkRadioButton(self.logic_frame, text="Extended Release", variable=self.release_type_var, 
                          value="extended", command=self.update_stage_dropdown).grid(row=3, column=1, padx=10, sticky="w")

        # Col 3: Medium (Present but Not Required)
        ctk.CTkLabel(self.logic_frame, text="Medium Name (Optional)", font=("Arial", 11, "bold")).grid(row=0, column=2, padx=10, pady=5, sticky="w")
        self.medium_name_entry = ctk.CTkEntry(self.logic_frame, placeholder_text="Acid / Buffer")
        self.medium_name_entry.grid(row=1, column=2, rowspan=3, padx=10, sticky="ew")

        # Col 4: Stage (Will be updated based on Release Type)
        ctk.CTkLabel(self.logic_frame, text="Select Stage", font=("Arial", 11, "bold")).grid(row=0, column=3, padx=10, pady=5, sticky="w")
        self.stage_var = ctk.StringVar(value="S1")
        self.stage_menu = ctk.CTkOptionMenu(self.logic_frame, variable=self.stage_var, values=[])
        self.stage_menu.grid(row=1, column=3, rowspan=3, padx=10, sticky="ew")
        
        # Initialize stage dropdown
        self.update_stage_dropdown()

        # --- 5. Action Buttons ---
        action_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        action_frame.pack(fill="x", pady=15)
        
        # Clear Form Button
        ctk.CTkButton(action_frame, text="Clear Form", command=self.clear_dissolution_form,
                      height=35, fg_color="#FF9800", hover_color="#F57C00", font=("Arial", 11, "bold"), width=120).pack(side="left", padx=(50, 10))
        
        # Select & Process Button
        ctk.CTkButton(action_frame, text="Select & Process", command=self.process_dissolution_pdf, 
                      height=35, fg_color="#2CC985", hover_color="#229A65", font=("Arial", 13, "bold")).pack(side="left", fill="x", expand=True, padx=10)

        # --- 6. Log (Resized to match Assay) ---
        log_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        log_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(log_frame, text="Processing Log:", font=("Arial", 11, "bold")).pack(anchor="w", padx=5)
        self.disso_log = ctk.CTkTextbox(log_frame, height=60, font=("Consolas", 11)) 
        self.disso_log.pack(fill="x", padx=5, pady=(0,5))

        # --- 7. Extracted Data Preview (Expanded to match Assay) ---
        diss_tree_label = ctk.CTkLabel(main_frame, text="Extracted Data Preview:", font=("Arial", 12, "bold"))
        diss_tree_label.pack(anchor="w", pady=(5, 0))

        self.diss_tree_container = ctk.CTkFrame(main_frame)
        self.diss_tree_container.pack(fill="both", expand=True, pady=5) 
        
        # Initialize the treeview with ALL columns immediately
        self._build_diss_treeview()

    def on_sample_type_change(self):
        """Handle sample type change (Standard/Non-Standard)"""
        sample_type = self.sample_type_var.get()
        
        if sample_type == "standard":
            # Show standard frame, hide logic frame
            self.std_frame.pack(fill="x", pady=5)
            self.logic_frame.pack_forget()
        else:
            # Show logic frame, hide standard frame
            self.std_frame.pack_forget()
            self.logic_frame.pack(fill="x", pady=5)

    def update_stage_dropdown(self, *args):
        """Update stage dropdown based on release type"""
        release_type = self.release_type_var.get()
        
        # Determine prefix and stage options
        if release_type == "immediate":
            # For Immediate Release: S1, S2, S3
            stage_options = ["S1", "S2", "S3"]
        elif release_type == "delayed":
            # For Delayed Release: V1, V2, V3
            stage_options = ["V1", "V2", "V3"]
        elif release_type == "extended":
            # For Extended Release: L1, L2, L3
            stage_options = ["L1", "L2", "L3"]
        else:
            stage_options = []
        
        # Update dropdown
        self.stage_menu.configure(values=stage_options)
        if stage_options:
            self.stage_var.set(stage_options[0])
        else:
            self.stage_var.set("")

    def clear_dissolution_form(self):
        """Clear all dissolution form fields"""
        # Clear entries
        self.disso_sample_id.delete(0, 'end')
        self.disso_user_id.delete(0, 'end')
        self.medium_name_entry.delete(0, 'end')
        
        # Reset radio buttons to default
        self.sample_type_var.set("non_standard")
        self.std_type_var.set("CS")
        self.comp_type_var.set("single")
        self.release_type_var.set("immediate")
        
        # Reset stage dropdown
        self.update_stage_dropdown()
        
        # Show appropriate frames based on sample type
        self.on_sample_type_change()
        
        # Clear log
        self.disso_log.delete("1.0", "end")
        
        # Clear treeview (check if treeview exists)
        if hasattr(self, 'diss_tree') and self.diss_tree:
            for item in self.diss_tree.get_children():
                self.diss_tree.delete(item)
        
        # Show success message
        self.disso_log.insert("end", "Form cleared successfully. Ready for new input.\n")
        messagebox.showinfo("Form Cleared", "Dissolution form has been cleared. You can now enter new data.")

    def init_db_tables(self):
        """Ensure all tables (Tab 1 & Tab 2) exist"""
        try:
            if os.path.exists(DB_CONFIG_FILE):
                with open(DB_CONFIG_FILE, "r") as f:
                    host, port, user, pwd, db = f.read().splitlines()
                conn = mysql.connector.connect(host=host, port=port, user=user, password=pwd, database=db)
                cursor = conn.cursor()
                
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS shimadzu_lc2050_results (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        machine_id VARCHAR(50), u_id VARCHAR(100), user_id VARCHAR(100), test_code VARCHAR(100), acquired_by VARCHAR(100),
                        sample_name_header VARCHAR(100), sample_id VARCHAR(100), tray VARCHAR(50), vial VARCHAR(50),
                        injection_volume VARCHAR(50), data_file VARCHAR(255), method_file VARCHAR(255), batch_file VARCHAR(255),
                        report_format_file VARCHAR(255), date_acquired VARCHAR(100), date_processed VARCHAR(100),
                        title VARCHAR(150), sample_name VARCHAR(150), sample_id_ind VARCHAR(150),
                        ret_time VARCHAR(50), area VARCHAR(50), tailing_factor VARCHAR(50), theoretical_plate VARCHAR(50),
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                """)
                
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS shimadzu_lc2050_multicom_raw (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        machine_id VARCHAR(50), u_id VARCHAR(100), user_id VARCHAR(100), test_code VARCHAR(100),
                        acquired_by VARCHAR(100), sample_name_header VARCHAR(100), sample_id VARCHAR(100), tray VARCHAR(50),
                        vial VARCHAR(50), injection_volume VARCHAR(50), data_file VARCHAR(255), method_file VARCHAR(255),
                        batch_file VARCHAR(255), report_format_file VARCHAR(255), date_acquired VARCHAR(100), date_processed VARCHAR(100),
                        compound_name VARCHAR(150), title VARCHAR(150), sample_name VARCHAR(150), sample_id_ind VARCHAR(150),
                        ret_time VARCHAR(50), area VARCHAR(50), tailing_factor VARCHAR(50), theoretical_plate VARCHAR(50),
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                """)

                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS shimadzu_dissolution_raw (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        machine_id VARCHAR(50), u_id VARCHAR(100), user_id VARCHAR(100), test_code VARCHAR(100),
                        acquired_by VARCHAR(100), sample_name_header VARCHAR(100), sample_id VARCHAR(100), 
                        tray VARCHAR(50), vial VARCHAR(50), injection_volume VARCHAR(50),
                        data_file VARCHAR(255), method_file VARCHAR(255), batch_file VARCHAR(255),
                        report_format_file VARCHAR(255), date_acquired VARCHAR(100), date_processed VARCHAR(100),
                        compound_name VARCHAR(150), title VARCHAR(150), sample_name VARCHAR(150), sample_id_ind VARCHAR(150),
                        ret_time VARCHAR(50), area VARCHAR(50), height VARCHAR(50), tailing_factor VARCHAR(50), theoretical_plate VARCHAR(50),
                        
                        component_type VARCHAR(50), process_type VARCHAR(50), medium_name VARCHAR(100),
                        stage VARCHAR(20), vessel_id VARCHAR(20),
                        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
                    )
                """)
                conn.commit()
                conn.close()
        except Exception as e:
            self.log_status(f"DB Init Error: {e}")

    # =========================================================================
    # TAB 1 ACTION: ASSAY EXTRACTION
    # =========================================================================
    def select_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        self.log_status(f"File dialog returned: {files}")
        if files:
            test_code = self.test_code_entry.get().strip()
            if test_code:
                self.save_test_code(test_code)
            
            sample_id = self.sample_id_entry.get().strip()
            user_id = self.user_id_entry.get().strip()
            if not sample_id or not user_id:
                messagebox.showwarning("Input Required", "Enter both sample ID (u_id) and User ID.")
                return

            import fitz
            self.log_status(f"PyMuPDF version: {fitz.__doc__}")

            for filepath in files:
                filename = os.path.basename(filepath)
                self.log_status(f"Processing: {filename}")
                try:
                    with open(filepath, 'rb') as f:
                        data = f.read()
                        doc = fitz.open(stream=data, filetype='pdf')
                        lines = [line.strip() for page in doc for line in page.get_text().splitlines() if line.strip()]

                        self.log_status(f"Extracted {len(lines)} lines")

                        if len(lines) > 0:
                            if self.mode_var.get() == "single":
                                self.extract_single(lines, self.machine_id_entry.get().strip(), sample_id, user_id)
                            else:
                                self.extract_multiple(lines, self.machine_id_entry.get().strip(), sample_id, user_id)
                        else:
                            self.log_status(f"No text extracted (image PDF?)")
                except Exception as e:
                    self.log_status(f"Error: {e}")

    # --------------------- Helpers (Assay) ---------------------
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
        self.log_status("Test code saved.")

    def save_config(self):
        machine_id = self.machine_id_entry.get().strip()
        test_code = self.test_code_entry.get().strip()
        with open(CONFIG_FILE, "w") as f:
            f.write(machine_id + "\n")
            f.write("dummy_path\n")
            f.write(test_code + "\n")
        self.log_status("Config saved.")

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                lines = f.read().splitlines()
                if len(lines) >= 1:
                    # Tab 1
                    self.machine_id_entry.configure(state="normal")
                    self.machine_id_entry.insert(0, lines[0])
                    self.machine_id_entry.configure(state="disabled")
                    # Tab 2
                    self.disso_machine_id_entry.configure(state="normal")
                    self.disso_machine_id_entry.insert(0, lines[0])
                    self.disso_machine_id_entry.configure(state="disabled")

                if len(lines) >= 3:
                    self.test_code_entry.set(lines[2])
                    
                    # Dissolution Dropdown
                    if lines[2] in ["10010", "10011"]:
                        self.disso_test_code.set(lines[2])

    def open_db_config(self):
        win = ctk.CTkToplevel(self)
        win.title("DB Config")
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
            messagebox.showinfo("Saved", "DB config saved.")

        ctk.CTkButton(win, text="Save", command=save).pack(pady=10)

    # --------------------- Assay Logic ---------------------
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

    # --------------------- Assay Multi ---------------------
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
            idx = len(lc) - 1 - lc[::-1].index(label_l) 
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
            except: return ""

        header_common = {
            "machine_id": machine_id, "u_id": sample_id, "user_id": user_id,
            "test_code": self.test_code_entry.get().strip(),
            "acquired_by": get_val("Acquired by"), "sample_name_header": get_val("Sample Name"),
            "sample_id": get_val("Sample ID"), "tray": int(get_val("Tray#") or 0),
            "vial": int(get_val("Vial#") or 0), "injection_volume": float(get_val("Injection Volume") or 0),
            "data_file": get_val("Data File"), "method_file": get_val("Method File"),
            "batch_file": get_val("Batch File"), "report_format_file": get_val("Report Format File"),
            "date_acquired": get_val("Date Acquired"), "date_processed": get_val("Date Processed")
        }

        compound_starts = self._find_compound_starts(lines)
        if not compound_starts:
            self.log_status("No compound headers found.")
            return

        rows_all = []
        for idx, (start_index, compound_name) in enumerate(compound_starts):
            end_index = len(lines)
            if idx + 1 < len(compound_starts):
                end_index = compound_starts[idx + 1][0]

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
            if title_idx_in_block is None: continue

            sublines = lines[title_idx_in_block:end_index]
            titles = self._sub_get_table_section(sublines, "Title")
            sample_names = self._sub_get_table_section(sublines, "Sample Name")
            sample_ids = self._sub_get_table_section(sublines, "Sample ID")
            ret_times = self._sub_get_table_section(sublines, "Ret. Time")
            areas = self._sub_get_table_section(sublines, "Area")
            tailing_factors = self._sub_get_table_section(sublines, "Tailing Factor")
            plates = self._sub_get_table_section(sublines, "Theoretical Plate") or self._sub_get_table_section(sublines, "Number of Theoretical Plate(USP)")

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

        if rows_all: self.insert_multi_db(rows_all)

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
                "compound_name",
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

    # =========================================================================
    # TAB 2 ACTION: DISSOLUTION EXTRACTION (UPDATED)
    # =========================================================================
    def process_dissolution_pdf(self):
        t_code = self.disso_test_code.get().strip() 
        s_id_entry = self.disso_sample_id.get().strip()
        u_id = self.disso_user_id.get().strip()

        if not t_code or not s_id_entry or not u_id:
            messagebox.showwarning("Error", "Missing Test Code, Sample ID, or User ID.")
            return
        
        m_id = self.disso_machine_id_entry.get().strip()
        sample_type = self.sample_type_var.get()
        
        # Clear previous log
        self.disso_log.delete("1.0", "end")
        
        if sample_type == "standard":
            # Standard file processing - stage will be auto-detected from PDF
            self.disso_log.insert("end", f"Processing Standard file (CS/SS will be auto-detected from PDF)...\n")
            self._process_standard_file(s_id_entry, u_id, t_code, m_id)
        else:
            # Non-Standard file processing - validate all selections
            comp_type = self.comp_type_var.get()
            release_type = self.release_type_var.get()
            medium = self.medium_name_entry.get().strip()
            stage_selected = self.stage_var.get()
            
            validation_errors = []
            
            if not comp_type:
                validation_errors.append("Component Type")
            
            if not release_type:
                validation_errors.append("Release Type")
            
            if not stage_selected:
                validation_errors.append("Stage")
            
            if validation_errors:
                messagebox.showwarning("Validation Error", 
                                      f"Please select the following for non-standard files:\n" +
                                      "\n".join([f"- {error}" for error in validation_errors]))
                return
            
            self.disso_log.insert("end", f"Processing Non-Standard file with selections:\n")
            self.disso_log.insert("end", f"Component: {comp_type}, Release: {release_type}\n")
            self.disso_log.insert("end", f"Medium: {medium}, Stage: {stage_selected}\n")
            
            # Process non-standard file
            self._process_non_standard_file(s_id_entry, u_id, t_code, m_id, comp_type, release_type, medium, stage_selected)

    def _detect_standard_type_from_pdf(self, lines):
        """
        Detect whether the PDF contains CS or SS from Sample ID field
        Returns: "CS" or "SS" based on PDF content
        """
        try:
            # Get Sample ID from PDF
            sample_id_from_pdf = ""
            for i, line in enumerate(lines):
                if "Sample ID" in line and i + 1 < len(lines):
                    sample_id_from_pdf = lines[i + 1].lstrip(": ").strip()
                    break
            
            # Check for CS or SS in Sample ID
            if sample_id_from_pdf:
                sample_id_upper = sample_id_from_pdf.upper()
                if "CS" in sample_id_upper:
                    return "CS"
                elif "SS" in sample_id_upper:
                    return "SS"
            
            # Also check in all lines for CS/SS patterns
            for line in lines:
                line_upper = line.upper()
                if " CS " in line_upper or line_upper.endswith(" CS") or line_upper.startswith("CS "):
                    return "CS"
                elif " SS " in line_upper or line_upper.endswith(" SS") or line_upper.startswith("SS "):
                    return "SS"
            
            # Check in filename or user input
            sample_id_entry = self.disso_sample_id.get().strip().upper()
            if "CS" in sample_id_entry:
                return "CS"
            elif "SS" in sample_id_entry:
                return "SS"
                
        except Exception as e:
            self.disso_log.insert("end", f"Error detecting standard type: {e}\n")
        
        # Default to CS if not detected
        return "CS"

    def _process_standard_file(self, sample_id, user_id, test_code, machine_id):
        """Process Standard files (CS or SS auto-detected from PDF)"""
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if not files: 
            return
        
        total_inserted = 0
        all_inserted_rows = []

        try:
            with open(DB_CONFIG_FILE, "r") as f:
                host, port, user, pwd, db = f.read().splitlines()
            conn = mysql.connector.connect(host=host, port=port, user=user, password=pwd, database=db)
            cursor = conn.cursor()

            for filepath in files:
                filename = os.path.basename(filepath)
                self.disso_log.insert("end", f"Processing: {filename}\n")
                
                with open(filepath, 'rb') as f:
                    doc = fitz.open(stream=f.read(), filetype='pdf')
                    lines = [l.strip() for page in doc for l in page.get_text().splitlines() if l.strip()]

                    # Detect CS or SS from PDF content
                    detected_std_type = self._detect_standard_type_from_pdf(lines)
                    self.disso_log.insert("end", f"Auto-detected Standard Type: {detected_std_type}\n")
                    
                    # Update UI radio button to show detected type
                    self.std_type_var.set(detected_std_type)

                    def get_val(lbl):
                        try:
                            idx = lines.index(lbl)
                            return lines[idx+1].lstrip(": ").strip()
                        except:
                            return ""

                    extracted_compound = ""
                    for line in lines:
                        if "Compound Name" in line:
                            parts = line.split(":", 1)
                            if len(parts) > 1:
                                extracted_compound = parts[1].strip()
                            break

                    # Header for Standard files
                    header = {
                        "machine_id": machine_id, 
                        "u_id": sample_id,   
                        "user_id": user_id, 
                        "test_code": test_code,
                        "acquired_by": get_val("Acquired by"), 
                        "sample_name_header": get_val("Sample Name"),
                        "sample_id": get_val("Sample ID"), 
                        "tray": get_val("Tray#"),          
                        "vial": get_val("Vial#"),          
                        "injection_volume": get_val("Injection Volume"), 
                        "data_file": get_val("Data File"), 
                        "method_file": get_val("Method File"),
                        "batch_file": get_val("Batch File"),
                        "report_format_file": get_val("Report Format File"), 
                        "date_acquired": get_val("Date Acquired"),
                        "date_processed": get_val("Date Processed"),         
                        "compound_name": extracted_compound,
                        # Standard-specific fields
                        "component_type": "",
                        "process_type": "",
                        "medium_name": "",
                        "stage": detected_std_type,  # Auto-detected "CS" or "SS"
                        "vessel_id": ""
                    }

                    def get_col(lbl):
                        try:
                            lc = [x.lower() for x in lines]
                            idx = len(lc) - 1 - lc[::-1].index(lbl.lower())
                            res = []
                            stop_headers = ["Title", "Sample Name", "Sample ID", "Ret. Time", "Area", "Height", 
                                          "Tailing Factor", "Theoretical Plate", "Theoretical Plate(USP)", 
                                          "Number of Theoretical Plate(USP)"]
                            for i in range(idx+1, len(lines)):
                                v = lines[i]
                                if v in stop_headers: break
                                if not v.startswith(":"): res.append(v)
                            return res
                        except: return []
                    
                    def get_table_section(label):
                        try:
                            idx = [i for i, val in enumerate(lines) if val.strip().lower() == label.lower()][-1]
                            result = []
                            stop_headers = ["Title", "Sample Name", "Sample ID", "Ret. Time", "Area", "Height", 
                                          "Tailing Factor", "Theoretical Plate", "Theoretical Plate(USP)"]
                            for i in range(idx + 1, len(lines)):
                                if lines[i] in stop_headers:
                                    break
                                if lines[i].strip() and not lines[i].startswith(":"):
                                    result.append(lines[i].lstrip(": ").strip())
                            return result
                        except:
                            return []

                    titles = get_col("Title")
                    ret_times = get_col("Ret. Time")
                    areas = get_col("Area")
                    heights = get_col("Height") 
                    tailing = get_col("Tailing Factor")
                    plates = get_col("Theoretical Plate") or get_col("Theoretical Plate(USP)") or get_col("Number of Theoretical Plate(USP)")
                    
                    pdf_sample_names_col = get_table_section("Sample Name")
                    pdf_sample_ids_col = get_table_section("Sample ID")

                    # For Standard files, take all rows
                    rows_to_insert = []
                    for i in range(max(len(titles), len(ret_times))):
                        t = titles[i] if i < len(titles) else ""
                        
                        row = header.copy()
                        row["title"] = t
                        row["ret_time"] = ret_times[i] if i < len(ret_times) else ""
                        row["area"] = areas[i] if i < len(areas) else ""
                        row["height"] = heights[i] if i < len(heights) else "" 
                        row["tailing_factor"] = tailing[i] if i < len(tailing) else ""
                        row["theoretical_plate"] = plates[i] if i < len(plates) else "" 
                        row["sample_name"] = pdf_sample_names_col[i] if i < len(pdf_sample_names_col) else ""
                        row["sample_id_ind"] = pdf_sample_ids_col[i] if i < len(pdf_sample_ids_col) else ""
                        
                        rows_to_insert.append(row)

                    cols = ["machine_id", "u_id", "user_id", "test_code", "acquired_by", "sample_name_header",
                            "sample_id", "tray", "vial", "injection_volume", "report_format_file", "date_processed",
                            "data_file", "method_file", "batch_file", "date_acquired",
                            "title", "sample_name", "ret_time", "area", "height", "tailing_factor", "theoretical_plate",
                            "component_type", "process_type", "medium_name", "stage", "vessel_id",
                            "compound_name", "sample_id_ind"] 

                    for r in rows_to_insert:
                        vals = tuple(r.get(c, "") for c in cols)
                        
                        cursor.execute(f"INSERT INTO shimadzu_dissolution_raw ({', '.join(cols)}, timestamp) VALUES ({', '.join(['%s']*len(cols))}, NOW())", vals)
                        total_inserted += 1
                        self.disso_log.insert("end", f"Saved {detected_std_type} Standard - Area: {r['area']}\n")
                        all_inserted_rows.append(r)

            conn.commit()
            conn.close()
            self._build_diss_treeview(all_inserted_rows)
            messagebox.showinfo("Success", f"Saved {total_inserted} Standard Rows (Detected: {detected_std_type}).")

        except Exception as e:
            messagebox.showerror("Error", f"Standard Processing Error: {str(e)}")

    def _process_non_standard_file(self, sample_id, user_id, test_code, machine_id, comp_type, release_type, medium, stage_selected):
        """Process Non-Standard files"""
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if not files: 
            return
        
        total_inserted = 0
        all_inserted_rows = []

        try:
            with open(DB_CONFIG_FILE, "r") as f:
                host, port, user, pwd, db = f.read().splitlines()
            conn = mysql.connector.connect(host=host, port=port, user=user, password=pwd, database=db)
            cursor = conn.cursor()

            for filepath in files:
                filename = os.path.basename(filepath)
                self.disso_log.insert("end", f"Processing: {filename}\n")
                
                with open(filepath, 'rb') as f:
                    doc = fitz.open(stream=f.read(), filetype='pdf')
                    lines = [l.strip() for page in doc for l in page.get_text().splitlines() if l.strip()]

                    def get_val(lbl):
                        try:
                            idx = lines.index(lbl)
                            return lines[idx+1].lstrip(": ").strip()
                        except:
                            return ""

                    extracted_compound = ""
                    for line in lines:
                        if "Compound Name" in line:
                            parts = line.split(":", 1)
                            if len(parts) > 1:
                                extracted_compound = parts[1].strip()
                            break

                    # Header for Non-Standard files
                    header = {
                        "machine_id": machine_id, 
                        "u_id": sample_id,   
                        "user_id": user_id, 
                        "test_code": test_code,
                        "acquired_by": get_val("Acquired by"), 
                        "sample_name_header": get_val("Sample Name"),
                        "sample_id": get_val("Sample ID"), 
                        "tray": get_val("Tray#"),          
                        "vial": get_val("Vial#"),          
                        "injection_volume": get_val("Injection Volume"), 
                        "data_file": get_val("Data File"), 
                        "method_file": get_val("Method File"),
                        "batch_file": get_val("Batch File"),
                        "report_format_file": get_val("Report Format File"), 
                        "date_acquired": get_val("Date Acquired"),
                        "date_processed": get_val("Date Processed"),         
                        "compound_name": extracted_compound,
                        # Non-Standard specific fields
                        "component_type": comp_type,
                        "process_type": release_type,
                        "medium_name": medium,
                        "stage": stage_selected,
                        "vessel_id": stage_selected
                    }

                    def get_col(lbl):
                        try:
                            lc = [x.lower() for x in lines]
                            idx = len(lc) - 1 - lc[::-1].index(lbl.lower())
                            res = []
                            stop_headers = ["Title", "Sample Name", "Sample ID", "Ret. Time", "Area", "Height", 
                                          "Tailing Factor", "Theoretical Plate", "Theoretical Plate(USP)", 
                                          "Number of Theoretical Plate(USP)"]
                            for i in range(idx+1, len(lines)):
                                v = lines[i]
                                if v in stop_headers: break
                                if not v.startswith(":"): res.append(v)
                            return res
                        except: return []
                    
                    def get_table_section(label):
                        try:
                            idx = [i for i, val in enumerate(lines) if val.strip().lower() == label.lower()][-1]
                            result = []
                            stop_headers = ["Title", "Sample Name", "Sample ID", "Ret. Time", "Area", "Height", 
                                          "Tailing Factor", "Theoretical Plate", "Theoretical Plate(USP)"]
                            for i in range(idx + 1, len(lines)):
                                if lines[i] in stop_headers:
                                    break
                                if lines[i].strip() and not lines[i].startswith(":"):
                                    result.append(lines[i].lstrip(": ").strip())
                            return result
                        except:
                            return []

                    titles = get_col("Title")
                    ret_times = get_col("Ret. Time")
                    areas = get_col("Area")
                    heights = get_col("Height") 
                    tailing = get_col("Tailing Factor")
                    plates = get_col("Theoretical Plate") or get_col("Theoretical Plate(USP)") or get_col("Number of Theoretical Plate(USP)")
                    
                    pdf_sample_names_col = get_table_section("Sample Name")
                    pdf_sample_ids_col = get_table_section("Sample ID")

                    # For Non-Standard files, skip Average rows
                    rows_to_insert = []
                    for i in range(max(len(titles), len(ret_times))):
                        t = titles[i] if i < len(titles) else ""
                        
                        # Skip average rows for non-standard files
                        if t in ["Average", "%RSD", "Standard Deviation", "Std. Dev."]: 
                            continue
                        
                        row = header.copy()
                        row["title"] = t
                        row["ret_time"] = ret_times[i] if i < len(ret_times) else ""
                        row["area"] = areas[i] if i < len(areas) else ""
                        row["height"] = heights[i] if i < len(heights) else "" 
                        row["tailing_factor"] = tailing[i] if i < len(tailing) else ""
                        row["theoretical_plate"] = plates[i] if i < len(plates) else "" 
                        row["sample_name"] = pdf_sample_names_col[i] if i < len(pdf_sample_names_col) else ""
                        row["sample_id_ind"] = pdf_sample_ids_col[i] if i < len(pdf_sample_ids_col) else ""
                        
                        rows_to_insert.append(row)

                    cols = ["machine_id", "u_id", "user_id", "test_code", "acquired_by", "sample_name_header",
                            "sample_id", "tray", "vial", "injection_volume", "report_format_file", "date_processed",
                            "data_file", "method_file", "batch_file", "date_acquired",
                            "title", "sample_name", "ret_time", "area", "height", "tailing_factor", "theoretical_plate",
                            "component_type", "process_type", "medium_name", "stage", "vessel_id",
                            "compound_name", "sample_id_ind"] 

                    for r in rows_to_insert:
                        vals = tuple(r.get(c, "") for c in cols)
                        
                        cursor.execute(f"INSERT INTO shimadzu_dissolution_raw ({', '.join(cols)}, timestamp) VALUES ({', '.join(['%s']*len(cols))}, NOW())", vals)
                        total_inserted += 1
                        self.disso_log.insert("end", f"Saved {stage_selected} - Area: {r['area']}\n")
                        all_inserted_rows.append(r)

            conn.commit()
            conn.close()
            self._build_diss_treeview(all_inserted_rows)
            messagebox.showinfo("Success", f"Saved {total_inserted} Non-Standard Rows.")

        except Exception as e:
            messagebox.showerror("Error", f"Non-Standard Processing Error: {str(e)}")

    def _build_diss_treeview(self, data=None):
        for w in self.diss_tree_container.winfo_children():
            w.destroy()
            
        # ALL columns extracted into DB for Dissolution
        diss_cols = [
            "machine_id", "u_id", "user_id", "test_code", "acquired_by", 
            "sample_name_header", "sample_id", "tray", "vial", "injection_volume", 
            "data_file", "method_file", "batch_file", "report_format_file", 
            "date_acquired", "date_processed", "compound_name", "title", 
            "sample_name", "sample_id_ind", "ret_time", "area", "height", 
            "tailing_factor", "theoretical_plate", "component_type", 
            "process_type", "medium_name", "stage", "vessel_id"
        ]
        
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Dissolution.Treeview", background="white", foreground="black", fieldbackground="white", rowheight=25)
        style.configure("Dissolution.Treeview.Heading", background="#E0E0E0", foreground="black", font=("Arial", 9, "bold"))
        
        self.diss_tree = ttk.Treeview(self.diss_tree_container, columns=diss_cols, show='headings', style="Dissolution.Treeview")
        for col in diss_cols:
            self.diss_tree.heading(col, text=col)
            # Adjust widths based on content type
            if col in ["machine_id", "u_id", "user_id", "test_code", "tray", "vial", "stage", "vessel_id"]:
                width = 80
            elif col in ["data_file", "method_file", "batch_file", "report_format_file", "compound_name"]:
                width = 150
            else:
                width = 100
            self.diss_tree.column(col, width=width, anchor="center")
        
        y_scroll = ttk.Scrollbar(self.diss_tree_container, orient="vertical", command=self.diss_tree.yview)
        x_scroll = ttk.Scrollbar(self.diss_tree_container, orient="horizontal", command=self.diss_tree.xview)
        self.diss_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.diss_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.diss_tree_container.rowconfigure(0, weight=1)
        self.diss_tree_container.columnconfigure(0, weight=1)

        if data:
            for i, row_data in enumerate(data):
                values = tuple(row_data.get(col, "") for col in diss_cols)
                self.diss_tree.insert("", "end", values=values, tags=("Dissolution.Treeview",))

    # --- Tab Animation Logic ---
    def tabview_callback(self):
        """Called whenever a tab is clicked/selected."""
        selected_tab_name = self.tabview.get()

        if selected_tab_name == self.current_tab_name_tracker:
            return 
            
        # If moving TO Dissolution FROM Assay
        if selected_tab_name == "Dissolution" and self.current_tab_name_tracker == "Assay":
             self.animate_tab_switch("Assay", "Dissolution")
        
        # If moving TO Assay FROM Dissolution (optional reverse animation)
        elif selected_tab_name == "Assay" and self.current_tab_name_tracker == "Dissolution":
             self.animate_tab_switch("Dissolution", "Assay")

        self.current_tab_name_tracker = selected_tab_name

    def animate_tab_switch(self, from_tab, to_tab):
        pass 

    # --------------------- logging & license ---------------------
    def log_status(self, msg):
        self.status_box.insert("end", msg+"\n")
        self.status_box.see("end")

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