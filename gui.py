import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.font as tkfont
import pandas as pd
import os
import datetime
import csv
import json
from decoder import UniversalJT808Decoder

# ตั้งค่า Theme ของ CustomTkinter
ctk.set_appearance_mode("Dark")  
ctk.set_default_color_theme("blue")  

# ==========================================
# --- MAIN APPLICATION (CUSTOM TKINTER) ---
# ==========================================
class ModernApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("GPS Decoder Pro (Ultimate CTk)")
        self.geometry("1450x900")

        # --- HEADER (THEME SWITCHER) ---
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill='x', padx=20, pady=(10, 0))

        ctk.CTkLabel(header_frame, text="GPS Decoder Pro", font=('Segoe UI', 18, 'bold'), text_color="#3b82f6").pack(side='left')

        self.theme_var = ctk.StringVar(value="Dark")
        theme_menu = ctk.CTkOptionMenu(header_frame, values=["Dark", "Light", "System"],
                                       variable=self.theme_var, command=self.change_theme, width=120)
        theme_menu.pack(side='right')

        # --- STYLE FOR TREEVIEW ---
        self.style = ttk.Style()
        self.style.theme_use("default")

        self.config_file = 'master_mapping_config.csv'
        self.decoder = UniversalJT808Decoder(self.config_file)
        self.results_df = None
        self.manual_results_df = None 
        self.filtered_df = None
        self.config_df = None
        self.last_loaded_paths = None 
        self.font_measure = tkfont.Font(family='Segoe UI', size=10)
        
        # Column Visibility State
        self.column_prefs_file = 'column_prefs.json'
        self.column_visibility_vars = {}
        self.current_all_columns = []
        
        # --- LAYOUT ---
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=20, pady=(5, 0))
        
        self.tab_manual = self.tabview.add("🔍 Manual Decode")
        self.tab_file = self.tabview.add("📂 File Process")
        self.tab_checksum = self.tabview.add("🛠️ Checksum Tool")
        self.tab_config = self.tabview.add("⚙️ Settings") 
        
        self.setup_ui()
        self.setup_checksum_ui()
        self.setup_config_editor() 

        # --- SYSTEM LOG ---
        log_frame = ctk.CTkFrame(self)
        log_frame.pack(side='bottom', fill='x', padx=20, pady=(10, 20))
        ctk.CTkLabel(log_frame, text="System Log:", font=('Segoe UI', 12, 'bold')).pack(anchor='w', padx=10, pady=(5,0))
        self.txt_log = ctk.CTkTextbox(log_frame, height=100, font=('Consolas', 12))
        self.txt_log.pack(fill='both', padx=10, pady=(0, 10))

        # บังคับให้โหลดสีและสไตล์ตาม Theme ปัจจุบันให้ครบทุกจุด
        self.change_theme(self.theme_var.get())

        if self.decoder.config is not None: self.log(f"System Ready: Loaded {self.config_file}", "success")
        else: self.log("⚠️ Warning: Configuration file missing.", "error")

    # ==========================================
    # --- DYNAMIC THEME ENGINE ---
    # ==========================================
    def change_theme(self, new_theme: str):
        ctk.set_appearance_mode(new_theme)
        # ตรวจสอบว่าโหมดปัจจุบันเป็นสว่างหรือมืด
        mode = ctk.get_appearance_mode()
        
        # 1. เปลี่ยนสีของตาราง (Treeview)
        if mode == "Light":
            self.style.configure("Treeview", background="#ffffff", foreground="#334155", rowheight=30, fieldbackground="#ffffff", borderwidth=0, font=('Segoe UI', 10))
            self.style.map('Treeview', background=[('selected', '#dbeafe')], foreground=[('selected', '#1e40af')])
            self.style.configure("Treeview.Heading", background="#e2e8f0", foreground="#1e293b", relief="flat", font=('Segoe UI', 10, 'bold'))
            self.style.map("Treeview.Heading", background=[('active', '#cbd5e1')])
            
            # ปรับสีแถว (Tag) สำหรับ Light Mode
            for t in [getattr(self, 'tree_manual', None), getattr(self, 'tree_file', None), getattr(self, 'tree_config', None)]:
                if t and t.winfo_exists():
                    t.tag_configure('even_row', background='#f8fafc', foreground='#334155')
                    t.tag_configure('error_row', background='#fee2e2', foreground='#b91c1c')
        else:
            self.style.configure("Treeview", background="#2b2b2b", foreground="white", rowheight=30, fieldbackground="#2b2b2b", borderwidth=0, font=('Segoe UI', 10))
            self.style.map('Treeview', background=[('selected', '#1f538d')], foreground=[('selected', 'white')])
            self.style.configure("Treeview.Heading", background="#3b3b3b", foreground="white", relief="flat", font=('Segoe UI', 10, 'bold'))
            self.style.map("Treeview.Heading", background=[('active', '#4b4b4b')])
            
            # ปรับสีแถว (Tag) สำหรับ Dark Mode
            for t in [getattr(self, 'tree_manual', None), getattr(self, 'tree_file', None), getattr(self, 'tree_config', None)]:
                if t and t.winfo_exists():
                    t.tag_configure('even_row', background='#323232', foreground='white')
                    t.tag_configure('error_row', background='#7f1d1d', foreground='white')

        # 2. เปลี่ยนสีของ System Log Textbox
        if hasattr(self, 'txt_log'):
            if mode == "Light":
                self.txt_log.tag_config('success', foreground='#15803d') # สีเขียวเข้ม
                self.txt_log.tag_config('error', foreground='#b91c1c')   # สีแดงเข้ม
                self.txt_log.tag_config('warning', foreground='#b45309') # สีส้มเข้ม
                self.txt_log.tag_config('info', foreground='#334155')    # สีเทาเข้ม
            else:
                self.txt_log.tag_config('success', foreground='#4ade80') 
                self.txt_log.tag_config('error', foreground='#f87171')   
                self.txt_log.tag_config('warning', foreground='#fcd34d')
                self.txt_log.tag_config('info', foreground='#e2e8f0')

        # 3. เปลี่ยนสีของผลลัพธ์ Checksum Textbox
        if hasattr(self, 'chk_result_textbox'):
            if mode == "Light":
                self.chk_result_textbox.configure(text_color="#047857") # เขียวเข้ม เพื่อให้อ่านง่ายบนพื้นขาว
            else:
                self.chk_result_textbox.configure(text_color="#6ee7b7") # เขียวสว่าง สำหรับพื้นดำ

    # ระบบ Undo/Redo (Ctrl+Z / Ctrl+Y) สำหรับ Textbox
    def apply_undo_redo(self, textbox):
        tb = textbox._textbox
        tb.configure(undo=True, maxundo=-1, autoseparators=True)
        def do_undo(event):
            try: tb.edit_undo()
            except: pass
            return "break"
        def do_redo(event):
            try: tb.edit_redo()
            except: pass
            return "break"
        tb.bind("<Control-z>", do_undo)
        tb.bind("<Control-y>", do_redo)

    # --- [CORE UTILS] ---
    def log(self, message, level="info"):
        if hasattr(self, 'txt_log'):
            self.txt_log.insert(tk.END, f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {message}\n", level)
            self.txt_log.see(tk.END)
        else: print(f"[LOG] {message}")

    def setup_ui(self):
        # --- MANUAL DECODE TAB ---
        top_bar = ctk.CTkFrame(self.tab_manual, fg_color="transparent")
        top_bar.pack(fill='x', pady=(0, 15))
        ctk.CTkLabel(top_bar, text="Paste Hex String (วาง Hex ได้ทุกรูปแบบ ระบบจะคัดแยกอัตโนมัติ):", font=('Segoe UI', 14, 'bold')).pack(anchor='w', pady=(0, 5))
        
        self.txt_manual = ctk.CTkTextbox(top_bar, height=80, font=('Consolas', 12))
        self.txt_manual.pack(fill='x', pady=(0, 10))
        self.apply_undo_redo(self.txt_manual)
        
        btn_bar_manual = ctk.CTkFrame(top_bar, fg_color="transparent")
        btn_bar_manual.pack(fill='x')
        
        ctk.CTkButton(btn_bar_manual, text="⚡ DECODE HEX", command=self.run_manual, font=('Segoe UI', 14, 'bold')).pack(side='left', padx=(0, 10))
        ctk.CTkButton(btn_bar_manual, text="🛠️ Checksum Tool", command=lambda: self.open_checksum_tool_from_tree(self.tree_manual), fg_color="#f59e0b", hover_color="#d97706", text_color="black").pack(side='left', padx=(0, 10))
        ctk.CTkButton(btn_bar_manual, text="💾 Export Report", command=lambda: self.export_data(self.manual_results_df), fg_color="#10b981", hover_color="#059669").pack(side='left')

        self.create_table(self.tab_manual, "manual")

        # --- FILE PROCESS TAB ---
        f_container = ctk.CTkFrame(self.tab_file, fg_color="transparent")
        f_container.pack(fill='both', expand=True)
        self.lbl_protocol = ctk.CTkLabel(f_container, text=f"Current Protocol: {os.path.basename(self.config_file)}", text_color="gray")
        self.lbl_protocol.pack(fill='x', pady=(0, 5))

        btn_bar = ctk.CTkFrame(f_container, fg_color="transparent")
        btn_bar.pack(fill='x', pady=(0, 10))
        ctk.CTkButton(btn_bar, text="⚙️ Change Protocol", command=self.change_protocol_file, fg_color="#64748b", hover_color="#475569").pack(side='left', padx=(0, 10))
        ctk.CTkButton(btn_bar, text="📂 Load CSV Data", command=self.select_file).pack(side='left', padx=(0, 10))
        ctk.CTkButton(btn_bar, text="🔄 Refresh Data", command=self.refresh_data, fg_color="#0ea5e9", hover_color="#0284c7").pack(side='left', padx=(0, 10))
        ctk.CTkButton(btn_bar, text="💾 Export Report", command=lambda: self.export_data(self.results_df), fg_color="#10b981", hover_color="#059669").pack(side='left')

        view_ctrl = ctk.CTkFrame(f_container)
        view_ctrl.pack(fill='x', pady=(0, 10))
        vc_inner = ctk.CTkFrame(view_ctrl, fg_color="transparent")
        vc_inner.pack(fill='x', padx=10, pady=10)

        ctk.CTkLabel(vc_inner, text="Start Row:").pack(side='left')
        self.var_f_start = ctk.StringVar(value="0")
        ctk.CTkEntry(vc_inner, textvariable=self.var_f_start, width=80).pack(side='left', padx=5)

        ctk.CTkLabel(vc_inner, text="Show Limit:").pack(side='left')
        self.var_f_limit = ctk.StringVar(value="1000")
        ctk.CTkEntry(vc_inner, textvariable=self.var_f_limit, width=80).pack(side='left', padx=5)

        ctk.CTkButton(vc_inner, text="Go / Update", command=self.update_file_table, width=100).pack(side='left', padx=10)
        ctk.CTkButton(vc_inner, text="<< Prev", command=self.prev_page_file, width=80, fg_color="#4b5563").pack(side='left', padx=(20, 5))
        ctk.CTkButton(vc_inner, text="Next >>", command=self.next_page_file, width=80, fg_color="#4b5563").pack(side='left', padx=5)
        
        ctk.CTkButton(vc_inner, text="🛠️ Checksum Tool", command=lambda: self.open_checksum_tool_from_tree(self.tree_file), width=120, fg_color="#f59e0b", hover_color="#d97706", text_color="black").pack(side='left', padx=(15, 0))
        
        self.lbl_file_info = ctk.CTkLabel(vc_inner, text="Total: 0 | Showing: 0-0", font=('Segoe UI', 12, 'bold'), text_color="#3b82f6")
        self.lbl_file_info.pack(side='right', padx=10)

        self.create_table(f_container, "file")

    def create_table(self, frame, name):
        table_frame = ctk.CTkFrame(frame)
        table_frame.pack(expand=True, fill='both')
        
        tree = ttk.Treeview(table_frame, show='headings')
        vsb = ctk.CTkScrollbar(table_frame, orientation="vertical", command=tree.yview)
        hsb = ctk.CTkScrollbar(table_frame, orientation="horizontal", command=tree.xview)
        
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        tree.bind("<Double-1>", lambda e: self.show_row_details(e, tree))
        
        if os.name == 'nt': tree.bind("<Button-3>", lambda event: self.show_advanced_column_menu(event, tree))
        else: tree.bind("<Button-2>", lambda event: self.show_advanced_column_menu(event, tree))
            
        if name == "file": self.tree_file = tree
        else: self.tree_manual = tree

    # --- [ADVANCED COLUMN VISIBILITY MANAGER] ---
    def load_column_prefs(self):
        if os.path.exists(self.column_prefs_file):
            try:
                with open(self.column_prefs_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except: pass
        return {}

    def save_column_prefs(self):
        prefs = {col: var.get() for col, var in self.column_visibility_vars.items()}
        try:
            with open(self.column_prefs_file, 'w', encoding='utf-8') as f:
                json.dump(prefs, f, ensure_ascii=False, indent=4)
        except Exception as e: print(f"Error saving column prefs: {e}")

    def apply_column_visibility(self):
        visible_cols = [col for col in self.current_all_columns if self.column_visibility_vars[col].get()]
        if len(visible_cols) == 0: visible_cols = [self.current_all_columns[0]]
            
        if hasattr(self, 'tree_file') and self.tree_file.winfo_exists():
            self.tree_file["displaycolumns"] = visible_cols
        if hasattr(self, 'tree_manual') and self.tree_manual.winfo_exists():
            self.tree_manual["displaycolumns"] = visible_cols

    def show_advanced_column_menu(self, event, tree):
        if not self.current_all_columns: return
        
        # ปรับสีเมนูให้เข้ากับธีม
        bg_col = "#2b2b2b" if ctk.get_appearance_mode() == "Dark" else "#ffffff"
        fg_col = "white" if ctk.get_appearance_mode() == "Dark" else "black"
        
        menu = tk.Menu(self, tearoff=0, font=('Segoe UI', 10), bg=bg_col, fg=fg_col)
        
        def show_all():
            for var in self.column_visibility_vars.values(): var.set(True)
            self.apply_column_visibility()
            self.save_column_prefs()

        def set_col_visibility(col, state):
            if not state: 
                visible_count = sum(1 for c in self.current_all_columns if self.column_visibility_vars[c].get())
                if visible_count <= 1:
                    messagebox.showwarning("คำเตือน", "ต้องแสดงอย่างน้อย 1 คอลัมน์ครับ")
                    return
            self.column_visibility_vars[col].set(state)
            self.apply_column_visibility()
            self.save_column_prefs()

        menu.add_command(label="👁️ แสดงคอลัมน์ทั้งหมด (Show All)", command=show_all)
        menu.add_separator()
        
        show_submenu = tk.Menu(menu, tearoff=0, font=('Segoe UI', 10), bg=bg_col, fg=fg_col)
        hide_submenu = tk.Menu(menu, tearoff=0, font=('Segoe UI', 10), bg=bg_col, fg=fg_col)
        
        menu.add_cascade(label="✅ เลือกโชว์คอลัมน์...", menu=show_submenu)
        menu.add_cascade(label="❌ เลือกซ่อนคอลัมน์...", menu=hide_submenu)
        
        visible_cols = [col for col in self.current_all_columns if self.column_visibility_vars[col].get()]
        hidden_cols = [col for col in self.current_all_columns if not self.column_visibility_vars[col].get()]
        
        if hidden_cols:
            for col in hidden_cols: show_submenu.add_command(label=f"โชว์: {col}", command=lambda c=col: set_col_visibility(c, True))
        else: show_submenu.add_command(label="(ไม่มีคอลัมน์ที่ซ่อนอยู่)", state="disabled")
            
        if visible_cols:
            for col in visible_cols: hide_submenu.add_command(label=f"ซ่อน: {col}", command=lambda c=col: set_col_visibility(c, False))
        else: hide_submenu.add_command(label="(ไม่มีคอลัมน์ที่แสดงอยู่)", state="disabled")

        try: menu.tk_popup(event.x_root, event.y_root)
        finally: menu.grab_release()

    def autosize_columns(self, tree, df, is_file_tab=False):
        if df is None: df = pd.DataFrame()
        all_config_cols = self.decoder.all_field_names 
        extra_cols = [c for c in df.columns if c not in all_config_cols]
        final_cols = all_config_cols + extra_cols
        
        self.current_all_columns = final_cols
        
        for t in [getattr(self, 'tree_file', None), getattr(self, 'tree_manual', None)]:
            if t and t.winfo_exists(): t['columns'] = final_cols
        
        saved_prefs = self.load_column_prefs()
        for col in final_cols:
            if col not in self.column_visibility_vars:
                self.column_visibility_vars[col] = tk.BooleanVar(value=saved_prefs.get(col, True))
                
        self.apply_column_visibility()
            
        for col in final_cols:
            max_width = 150 
            header_w = self.font_measure.measure(str(col)) + 40 
            if header_w > max_width: max_width = header_w
            if col in df.columns and len(df) > 0:
                sample_val = str(df[col].iloc[0])
                w = self.font_measure.measure(sample_val) + 40
                if w > max_width: max_width = w
            if max_width > 500: max_width = 500
            
            for t in [getattr(self, 'tree_file', None), getattr(self, 'tree_manual', None)]:
                if t and t.winfo_exists():
                    t.heading(col, text=col, anchor='center')
                    t.column(col, width=max_width, minwidth=100, anchor='center', stretch=False)
                    
        return final_cols

    def show_row_details(self, event, tree):
        item_id = tree.identify_row(event.y)
        if not item_id: return
        values = tree.item(item_id, 'values')
        cols = tree['columns']
        
        popup = ctk.CTkToplevel(self)
        popup.title("📄 Row Details")
        popup.geometry("600x700")
        popup.transient(self)
        
        txt = ctk.CTkTextbox(popup, font=('Consolas', 12))
        txt.pack(fill='both', expand=True, padx=20, pady=20)
        for c, v in zip(cols, values): 
            txt.insert(tk.END, f"{c:<30}: {v}\n")
            txt.insert(tk.END, "-"*60 + "\n")
        txt.configure(state='disabled')

    def run_manual(self):
        raw = self.txt_manual.get("1.0", tk.END).strip()
        if not raw: return
        decoded, st, _ = self.decoder.decode_raw(raw)
        self.manual_results_df = pd.DataFrame(decoded)
        self.display_manual(decoded)

    def display_manual(self, data):
        tree = self.tree_manual
        tree.delete(*tree.get_children())
        if not data: 
            self.log("No data found.", "error")
            return
        if data[0].get('_Status', '').startswith('❌'):
            self.log(f"Decode Failed: {data[0]['_Status']}", "error")
        else:
            self.log(f"Successfully decoded {len(data)} records.", "success")
            
        df = pd.DataFrame(data)
        final_cols = self.autosize_columns(tree, df, is_file_tab=False)
        for i, (_, r) in enumerate(df.iterrows()):
            tags = ['error_row'] if not str(r.get('_Status', 'OK')).startswith('OK') else (['even_row'] if i % 2 == 0 else [])
            tree.insert("", "end", values=[r.get(c,"") for c in final_cols], tags=tags)

    # --- [FILE FUNCTIONS] ---
    def select_file(self):
        paths = filedialog.askopenfilenames()
        if not paths: return
        self.last_loaded_paths = paths 
        self.process_batch_files(paths)

    def refresh_data(self, silent=False):
        self.decoder = UniversalJT808Decoder(self.config_file)
        if self.last_loaded_paths:
            self.process_batch_files(self.last_loaded_paths)
            if not silent: self.log("🔄 Data Refreshed.", "success")
        else:
            if not silent: self.log("⚠️ No file loaded to refresh.", "warning")

    def process_batch_files(self, paths):
        res = []
        for p in paths:
            try:
                if str(p).lower().endswith('.csv'):
                    df = pd.read_csv(p)
                    col = next((c for c in df.columns if c.lower() in ['raw-data','hex','raw']), None)
                    if col:
                        for _, r in df.iterrows():
                            d, _, _ = self.decoder.decode_raw(str(r[col]))
                            res.extend(d)
                else:
                    with open(p, 'r', encoding='utf-8') as f:
                        for line in f:
                            if '7E' in line:
                                d, _, _ = self.decoder.decode_raw(line)
                                res.extend(d)
            except Exception as e: 
                print(f"Error reading file {p}: {e}")
                
        self.results_df = pd.DataFrame(res)
        self.filtered_df = self.results_df 
        self.log(f"Batch Processed: {len(res)} rows loaded.", "success")
        self.update_file_table()

    def update_file_table(self):
        if self.results_df is None or self.results_df.empty: return
        try:
            start_idx = int(self.var_f_start.get()); limit = int(self.var_f_limit.get())
        except: start_idx = 0; limit = 1000; self.var_f_start.set("0")
        
        end_idx = start_idx + limit; total_rows = len(self.results_df)
        if start_idx < 0: start_idx = 0
        if end_idx > total_rows: end_idx = total_rows
        
        display_df = self.results_df.iloc[start_idx:end_idx]
        self.tree_file.delete(*self.tree_file.get_children())
        final_cols = self.autosize_columns(self.tree_file, display_df, is_file_tab=True)
        
        for i, (_, r) in enumerate(display_df.iterrows()):
            tags = ['error_row'] if not str(r.get('_Status', 'OK')).startswith('OK') else (['even_row'] if i % 2 == 0 else [])
            self.tree_file.insert("", "end", values=[r.get(c,"") for c in final_cols], tags=tags)
        self.lbl_file_info.configure(text=f"Total: {total_rows} | Showing: {start_idx}-{end_idx}")

    def next_page_file(self):
        try:
            current = int(self.var_f_start.get()); limit = int(self.var_f_limit.get())
            if self.results_df is not None and current + limit < len(self.results_df):
                self.var_f_start.set(str(current + limit)); self.update_file_table()
        except: pass

    def prev_page_file(self):
        try:
            current = int(self.var_f_start.get()); limit = int(self.var_f_limit.get())
            new_start = max(0, current - limit)
            self.var_f_start.set(str(new_start)); self.update_file_table()
        except: pass

    def export_data(self, target_df):
        if target_df is None or target_df.empty: 
            messagebox.showwarning("คำเตือน", "ไม่มีข้อมูลให้ Export ครับ")
            return
            
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel File", "*.xlsx"), ("CSV File", "*.csv")])
        if not file_path: return
        try:
            export_df = target_df.copy()
            non_empty = [c for c in export_df.columns if export_df[c].astype(str).str.strip().ne('').any()]
            keep = [c for c in self.decoder.all_field_names if c in non_empty]
            keep.extend([c for c in non_empty if c not in keep])
            export_df = export_df[keep]

            visible_cols = [col for col in self.current_all_columns if self.column_visibility_vars.get(col, tk.BooleanVar(value=True)).get()]
            if visible_cols: export_df = export_df[[c for c in visible_cols if c in export_df.columns]]

            for col in ['Device IMEI', 'Terminal ID', 'CCID', 'BSJ Serial No', 'Mobile Network Operator', 'Serial Number']:
                if col in export_df.columns: export_df[col] = export_df[col].astype(str)

            if file_path.lower().endswith('.csv'):
                export_df.to_csv(file_path, index=False, encoding='utf-8-sig', quoting=csv.QUOTE_ALL)
                self.log(f"Saved CSV: {os.path.basename(file_path)}", "success")
            else:
                export_df.to_excel(file_path, index=False)
                self.log(f"Saved Excel: {os.path.basename(file_path)}", "success")
        except Exception as e: self.log(f"Export Error: {e}", "error")

    # ==========================
    # --- CHECKSUM TOOL TAB ---
    # ==========================
    def setup_checksum_ui(self):
        container = ctk.CTkFrame(self.tab_checksum, fg_color="transparent")
        container.pack(fill='both', expand=True)
        
        left_panel = ctk.CTkFrame(container)
        left_panel.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        ctk.CTkLabel(left_panel, text="Hex Input (ลากเมาส์คลุมเพื่อดู ASCII | รองรับ Ctrl+Z/Y):", font=('Segoe UI', 14, 'bold')).pack(anchor='w', padx=10, pady=(10, 5))
        self.chk_hex_textbox = ctk.CTkTextbox(left_panel, height=200, font=('Consolas', 14))
        self.apply_undo_redo(self.chk_hex_textbox) 
        self.chk_hex_textbox.pack(fill='both', expand=True, padx=10, pady=(0, 5))
        self.chk_hex_textbox.tag_config("sync_hl", background="#eab308", foreground="black")
        
        cb_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        cb_frame.pack(fill='x', padx=10, pady=5)
        self.chk_remove_space_var = ctk.BooleanVar(value=True)
        self.chk_remove_7e_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(cb_frame, text="Remove Spaces", variable=self.chk_remove_space_var).pack(side='left', padx=(0, 15))
        ctk.CTkCheckBox(cb_frame, text="ตัด 7E/หา Checksum อัจฉริยะ (ทำเบื้องหลัง)", variable=self.chk_remove_7e_var).pack(side='left')
        
        ctk.CTkLabel(left_panel, text="ASCII Input (ลากเมาส์คลุมเพื่อดู Hex | รองรับ Ctrl+Z/Y):", font=('Segoe UI', 14, 'bold')).pack(anchor='w', padx=10, pady=(10, 5))
        self.chk_ascii_textbox = ctk.CTkTextbox(left_panel, height=100, font=('Consolas', 14))
        self.apply_undo_redo(self.chk_ascii_textbox) 
        self.chk_ascii_textbox.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        self.chk_ascii_textbox.tag_config("sync_hl", background="#eab308", foreground="black")
        
        btn_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        btn_frame.pack(fill='x', padx=10, pady=10)
        ctk.CTkButton(btn_frame, text="🧹 Clear", command=self.chk_clear_inputs, fg_color="#64748b").pack(side='left', fill='x', expand=True, padx=(0, 5))
        ctk.CTkButton(btn_frame, text="⚙️ Calculate (แปลง & ตรวจ)", command=self.chk_calculate).pack(side='left', fill='x', expand=True, padx=(5, 0))
        
        right_panel = ctk.CTkFrame(container)
        right_panel.pack(side='right', fill='both', expand=True, padx=(10, 0))
        
        ctk.CTkLabel(right_panel, text="Results & Verification:", font=('Segoe UI', 14, 'bold')).pack(anchor='w', padx=10, pady=(10, 5))
        self.chk_result_textbox = ctk.CTkTextbox(right_panel, font=('Consolas', 14))
        self.chk_result_textbox.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        self.chk_hex_textbox.bind("<B1-Motion>", self.sync_hex_selection)
        self.chk_hex_textbox.bind("<ButtonRelease-1>", self.sync_hex_selection)
        self.chk_ascii_textbox.bind("<B1-Motion>", self.sync_ascii_selection)
        self.chk_ascii_textbox.bind("<ButtonRelease-1>", self.sync_ascii_selection)

    def sync_hex_selection(self, event=None):
        self.chk_ascii_textbox.tag_remove("sync_hl", "1.0", tk.END)
        try:
            if self.chk_hex_textbox.tag_ranges(tk.SEL):
                first = self.chk_hex_textbox.index(tk.SEL_FIRST)
                last = self.chk_hex_textbox.index(tk.SEL_LAST)
                start_idx = len(self.chk_hex_textbox.get("1.0", first))
                end_idx = len(self.chk_hex_textbox.get("1.0", last))

                ascii_start = start_idx // 2
                ascii_end = end_idx // 2

                if ascii_end > ascii_start:
                    self.chk_ascii_textbox.tag_add("sync_hl", f"1.0+{ascii_start}c", f"1.0+{ascii_end}c")
        except: pass

    def sync_ascii_selection(self, event=None):
        self.chk_hex_textbox.tag_remove("sync_hl", "1.0", tk.END)
        try:
            if self.chk_ascii_textbox.tag_ranges(tk.SEL):
                first = self.chk_ascii_textbox.index(tk.SEL_FIRST)
                last = self.chk_ascii_textbox.index(tk.SEL_LAST)
                start_idx = len(self.chk_ascii_textbox.get("1.0", first))
                end_idx = len(self.chk_ascii_textbox.get("1.0", last))

                hex_start = start_idx * 2
                hex_end = end_idx * 2

                if hex_end > hex_start:
                    self.chk_hex_textbox.tag_add("sync_hl", f"1.0+{hex_start}c", f"1.0+{hex_end}c")
        except: pass

    def chk_hex_to_ascii(self, hex_string):
        try:
            bytes_object = bytes.fromhex(hex_string)
            result = ""
            for b in bytes_object:
                if 32 <= b <= 126: result += chr(b)
                else: result += "."
            return result
        except ValueError: return ""

    def chk_calculate_bcc(self, hex_string):
        try:
            bcc = 0
            for i in range(0, len(hex_string), 2): bcc ^= int(hex_string[i:i+2], 16)
            return f"{bcc:02X}"
        except ValueError: return "Error"

    def chk_crc16_ccitt_false(self, hex_string):
        try:
            data = bytes.fromhex(hex_string)
            crc = 0xFFFF
            for byte in data:
                crc ^= (byte << 8)
                for _ in range(8):
                    if crc & 0x8000: crc = (crc << 1) ^ 0x1021
                    else: crc <<= 1
                    crc &= 0xFFFF
            return f"{crc:04X}"
        except ValueError: return "Error"

    def chk_calculate(self):
        raw_input = self.chk_hex_textbox.get("1.0", "end-1c").strip()
        ascii_input = self.chk_ascii_textbox.get("1.0", "end-1c").strip()

        self.chk_result_textbox.configure(state=tk.NORMAL)
        self.chk_result_textbox.delete("1.0", tk.END)

        if raw_input:
            hex_str = raw_input
            if self.chk_remove_space_var.get(): 
                hex_str = hex_str.replace(" ", "").replace("\n", "").replace("\r", "")
                if hex_str != raw_input:
                    self.chk_hex_textbox.delete("1.0", tk.END)
                    self.chk_hex_textbox.insert("1.0", hex_str)
                    
            ascii_result = self.chk_hex_to_ascii(hex_str)
            self.chk_ascii_textbox.delete("1.0", tk.END)
            self.chk_ascii_textbox.insert("1.0", ascii_result)

            payload = hex_str.upper()
            original_checksum = "N/A"
            
            # --- อัลกอริทึมตัด Checksum ยืดหยุ่นขั้นเทพ (V2) ---
            if self.chk_remove_7e_var.get():
                if payload.startswith("7E") and payload.endswith("7E") and len(payload) >= 6:
                    original_checksum = payload[-4:-2] 
                    payload = payload[2:-4] 
                elif payload.startswith("7E") and not payload.endswith("7E"):
                    payload = payload[2:]
                    if len(payload) >= 2:
                        original_checksum = payload[-2:]
                        payload = payload[:-2]
                elif payload.endswith("7E") and not payload.startswith("7E"):
                    if len(payload) >= 4:
                        original_checksum = payload[-4:-2]
                        payload = payload[:-4]
                else:
                    if len(payload) >= 2:
                        original_checksum = payload[-2:]
                        payload = payload[:-2]

            bcc_result = self.chk_calculate_bcc(payload)
            crc16_result = self.chk_crc16_ccitt_false(payload)

            result_text = "--- Result & Verification ---\n\n"
            if self.chk_remove_7e_var.get() and original_checksum != "N/A":
                result_text += f"Original Checksum (ดึงมาจากแพ็กเกจ): {original_checksum}\n"
                
            result_text += f"Calculated BCC (XOR): {bcc_result}\n"
            
            if original_checksum != "N/A":
                match = "✅ ถูกต้อง (ตรงกัน)" if bcc_result == original_checksum else "❌ ผิดพลาด (ไม่ตรง)"
                result_text += f"สถานะการตรวจสอบ: {match}\n\n"
            else: result_text += "\n"
                
            result_text += f"CRC-16 (CCITT-FALSE): {crc16_result}\n"
            self.chk_result_textbox.insert(tk.END, result_text)
            
        elif ascii_input:
            try:
                hex_from_ascii = ascii_input.encode('utf-8').hex().upper()
                self.chk_hex_textbox.delete("1.0", tk.END)
                self.chk_hex_textbox.insert("1.0", hex_from_ascii)
                bcc_result = self.chk_calculate_bcc(hex_from_ascii)
                crc16_result = self.chk_crc16_ccitt_false(hex_from_ascii)
                result_text = "--- Result from ASCII Input ---\n\n"
                result_text += f"BCC (XOR): {bcc_result}\n\n"
                result_text += f"CRC-16 (CCITT-FALSE): {crc16_result}\n"
                self.chk_result_textbox.insert(tk.END, result_text)
            except Exception as e:
                self.chk_result_textbox.insert(tk.END, f"Error processing ASCII: {e}")

        self.chk_result_textbox.configure(state=tk.DISABLED)

    def chk_clear_inputs(self):
        self.chk_hex_textbox.delete("1.0", tk.END)
        self.chk_ascii_textbox.delete("1.0", tk.END)
        self.chk_hex_textbox.tag_remove("sync_hl", "1.0", tk.END)
        self.chk_ascii_textbox.tag_remove("sync_hl", "1.0", tk.END)
        self.chk_result_textbox.configure(state=tk.NORMAL)
        self.chk_result_textbox.delete("1.0", tk.END)
        self.chk_result_textbox.configure(state=tk.DISABLED)

    def open_checksum_tool_from_tree(self, tree):
        if not hasattr(tree, 'selection'): return
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("คำเตือน", "กรุณาคลิกเลือกแถวข้อมูลในตาราง 1 แถวก่อนครับ")
            return
            
        item = tree.item(selected[0])
        cols = list(tree['columns'])
        
        raw_hex = ""
        if 'Raw Hex Block' in cols:
            idx = cols.index('Raw Hex Block')
            raw_hex = str(item['values'][idx])
            
        if not raw_hex:
            messagebox.showwarning("คำเตือน", "ไม่พบข้อมูล Raw Hex ในแถวที่เลือกครับ")
            return
            
        self.chk_hex_textbox.delete("1.0", tk.END)
        self.chk_hex_textbox.insert("1.0", raw_hex)
        self.chk_calculate()
        self.tabview.set("🛠️ Checksum Tool")

    # ==========================
    # --- CONFIG EDITOR ---
    # ==========================
    def setup_config_editor(self):
        container = ctk.CTkFrame(self.tab_config, fg_color="transparent")
        container.pack(fill='both', expand=True)
        toolbar = ctk.CTkFrame(container, fg_color="transparent")
        toolbar.pack(fill='x', pady=(0, 10))
        ctk.CTkButton(toolbar, text="🔄 Reload File", command=self.load_config_data, fg_color="#64748b").pack(side='left', padx=(0, 5))
        ctk.CTkButton(toolbar, text="➕ Add Row (Insert)", command=self.add_config_row_inline).pack(side='left', padx=(0, 5))
        ctk.CTkButton(toolbar, text="❌ Delete Row", command=self.delete_config_row, fg_color="#ef4444", hover_color="#dc2626").pack(side='left', padx=(0, 5))
        ctk.CTkButton(toolbar, text="💾 Save System", command=self.save_config_data, fg_color="#10b981", hover_color="#059669").pack(side='right', padx=(5, 0))
        ctk.CTkButton(toolbar, text="📥 Import Backup", command=self.import_config_backup, fg_color="#f59e0b", hover_color="#d97706", text_color="black").pack(side='right', padx=(5, 0))
        ctk.CTkButton(toolbar, text="📤 Export Backup", command=self.export_config_backup, fg_color="#64748b").pack(side='right')
        
        columns = ['Category', 'SubID', 'FieldName', 'StartByte', 'Length', 'DataType', 'Scale']
        
        tree_frame = ctk.CTkFrame(container)
        tree_frame.pack(fill='both', expand=True)
        self.tree_config = ttk.Treeview(tree_frame, columns=columns, show='headings', selectmode='browse')
        vsb = ctk.CTkScrollbar(tree_frame, orientation="vertical", command=self.tree_config.yview)
        self.tree_config.configure(yscrollcommand=vsb.set)
        
        self.tree_config.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        
        for col in columns:
            self.tree_config.heading(col, text=col)
            self.tree_config.column(col, width=100, anchor='center')
        self.tree_config.column('FieldName', width=200, anchor='w')
        self.tree_config.bind("<Double-1>", self.on_tree_double_click)
        self.load_config_data()

    def load_config_data(self):
        try:
            if os.path.exists(self.config_file):
                self.config_df = pd.read_csv(self.config_file, dtype=str)
                self.config_df.fillna('', inplace=True)
                self.tree_config.delete(*self.tree_config.get_children())
                for i, row in self.config_df.iterrows():
                    vals = [row[c] for c in self.tree_config['columns']]
                    tags = ('even_row',) if i % 2 == 0 else ()
                    self.tree_config.insert("", "end", values=vals, tags=tags)
                self.log(f"Editor Loaded: {os.path.basename(self.config_file)}", "success")
            else: self.log("Config file not found.", "error")
        except Exception as e: self.log(f"Error loading config: {e}", "error")

    def change_protocol_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Protocol", "*.csv")], title="Select Protocol Configuration File")
        if not path: return
        self.config_file = path
        self.decoder = UniversalJT808Decoder(self.config_file)
        self.lbl_protocol.configure(text=f"Current Protocol: {os.path.basename(self.config_file)}")
        self.load_config_data()
        self.refresh_data(silent=True)
        self.log(f"Protocol Switched to: {os.path.basename(path)}", "success")

    def export_config_backup(self):
        try:
            rows = []
            columns = self.tree_config['columns']
            for item in self.tree_config.get_children():
                rows.append(self.tree_config.item(item)['values'])
            df_to_save = pd.DataFrame(rows, columns=columns)
            path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV File", "*.csv")])
            if not path: return
            df_to_save.to_csv(path, index=False)
            self.log(f"Config exported to {os.path.basename(path)}", "success")
        except Exception as e: messagebox.showerror("Export Error", str(e))

    def import_config_backup(self):
        try:
            path = filedialog.askopenfilename(filetypes=[("CSV File", "*.csv")])
            if not path: return
            new_df = pd.read_csv(path, dtype=str)
            new_df.fillna('', inplace=True)
            self.tree_config.delete(*self.tree_config.get_children())
            for i, row in new_df.iterrows():
                vals = [row[c] for c in self.tree_config['columns']]
                tags = ('even_row',) if i % 2 == 0 else ()
                self.tree_config.insert("", "end", values=vals, tags=tags)
            new_df.to_csv(self.config_file, index=False)
            self.decoder = UniversalJT808Decoder(self.config_file)
            self.refresh_data(silent=True)
            self.log(f"Imported rules from {os.path.basename(path)}", "success")
        except Exception as e: messagebox.showerror("Import Error", f"Invalid Config File: {e}")

    def on_tree_double_click(self, event):
        region = self.tree_config.identify("region", event.x, event.y)
        if region != "cell": return
        column = self.tree_config.identify_column(event.x)
        row_id = self.tree_config.identify_row(event.y)
        if not row_id: return
        col_idx = int(column.replace('#', '')) - 1
        col_name = self.tree_config['columns'][col_idx]
        x, y, w, h = self.tree_config.bbox(row_id, column)
        current_val = self.tree_config.item(row_id, 'values')[col_idx]
        
        if col_name == 'Category':
            base_cats = ['Header', 'Standard', 'Main', 'Container', 'Sub', 'Calculated']
            exist_cats = self.config_df['Category'].dropna().unique().tolist()
            all_cats = sorted(list(set(base_cats + exist_cats)))
            widget = ttk.Combobox(self.tree_config, values=all_cats) 
        elif col_name == 'DataType':
            base_dt = ['DWORD', 'WORD', 'BYTE', 'STRING', 'HEX', 'BCD', 'BIT', 'SIGNED_DWORD', 'SIGNED_WORD', 'INT']
            exist_dt = self.config_df['DataType'].dropna().unique().tolist()
            all_dt = sorted(list(set(base_dt + exist_dt)))
            widget = ttk.Combobox(self.tree_config, values=all_dt) 
        elif col_name == 'RefField':
            field_list = sorted(list(set(self.config_df['FieldName'].dropna().tolist())))
            widget = ttk.Combobox(self.tree_config, values=field_list)
        else:
            widget = ttk.Entry(self.tree_config)

        widget.place(x=x, y=y, width=w, height=h)
        widget.insert(0, str(current_val))
        widget.select_range(0, tk.END)
        widget.focus()

        def save_edit(event_inner=None):
            new_val = widget.get()
            current_values = list(self.tree_config.item(row_id, 'values'))
            current_values[col_idx] = new_val
            self.tree_config.item(row_id, values=current_values)
            widget.destroy()

        widget.bind("<Return>", save_edit)
        widget.bind("<FocusOut>", lambda e: save_edit()) 
        widget.bind("<Escape>", lambda e: widget.destroy())

    def add_config_row_inline(self):
        default_vals = ['Sub', '0000', 'New Field', '0', '4', 'DWORD', '1', '', '', '', '']
        selected = self.tree_config.selection()
        if selected:
            target_id = selected[-1]
            target_index = self.tree_config.index(target_id)
            new_item = self.tree_config.insert("", target_index + 1, values=default_vals)
        else:
            new_item = self.tree_config.insert("", "end", values=default_vals)
        self.tree_config.selection_set(new_item)
        self.tree_config.focus(new_item)
        self.tree_config.see(new_item)

    def delete_config_row(self):
        selected = self.tree_config.selection()
        if not selected: return
        if messagebox.askyesno("Confirm Delete", "Delete selected row(s)?"):
            for item in selected: self.tree_config.delete(item)

    def save_config_data(self):
        try:
            rows = []
            columns = self.tree_config['columns']
            for item in self.tree_config.get_children():
                rows.append(self.tree_config.item(item)['values'])
            new_df = pd.DataFrame(rows, columns=columns)
            new_df.to_csv(self.config_file, index=False)
            self.decoder = UniversalJT808Decoder(self.config_file)
            self.log("✅ Config saved and System updated.", "success")
            self.refresh_data(silent=True)
            messagebox.showinfo("Success", "Saved! Data has been refreshed.")
        except Exception as e: messagebox.showerror("Error", f"Failed to save: {e}")