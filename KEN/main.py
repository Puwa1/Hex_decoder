import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np

# ตั้งค่า Theme
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class AutoDecoderApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Auto CSV Decoder (Config Based)")
        self.geometry("1000x650")

        self.df_data = None   # เก็บข้อมูลดิบ
        self.df_rules = None  # เก็บกฎการแปลงค่า
        self.data_path = None

        # --- Layout ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0) # Header
        self.grid_rowconfigure(1, weight=1) # Content

        # --- Header Section ---
        self.header_frame = ctk.CTkFrame(self)
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=20)
        
        ctk.CTkLabel(self.header_frame, text="ระบบถอดรหัสอัตโนมัติด้วย Config CSV", 
                     font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)

        # --- Control Panel ---
        self.control_frame = ctk.CTkFrame(self)
        self.control_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        
        # 1. Load Data
        self.btn_data = ctk.CTkButton(self.control_frame, text="1. เลือกไฟล์ข้อมูล (Data CSV)", command=self.load_data_csv)
        self.btn_data.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        self.lbl_data = ctk.CTkLabel(self.control_frame, text="ยังไม่เลือกไฟล์")
        self.lbl_data.grid(row=0, column=1, padx=10, sticky="w")

        # 2. Load Rules
        self.btn_rules = ctk.CTkButton(self.control_frame, text="2. เลือกไฟล์กฎ (Rules CSV)", command=self.load_rules_csv)
        self.btn_rules.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.lbl_rules = ctk.CTkLabel(self.control_frame, text="ยังไม่เลือกไฟล์ (คอลัมน์: name, start, length, type, formula, condition)")
        self.lbl_rules.grid(row=1, column=1, padx=10, sticky="w")

        # Config Column Name
        ctk.CTkLabel(self.control_frame, text="ชื่อคอลัมน์ Hex ในไฟล์ข้อมูล:").grid(row=2, column=0, padx=20, pady=(10,0))
        self.entry_hex_col = ctk.CTkEntry(self.control_frame)
        self.entry_hex_col.grid(row=3, column=0, padx=20, pady=5, sticky="ew")
        self.entry_hex_col.insert(0, "raw-data")

        # 3. Process
        self.btn_process = ctk.CTkButton(self.control_frame, text="3. เริ่มการถอดรหัส (Auto Process)", 
                                         fg_color="green", hover_color="darkgreen", height=50,
                                         command=self.process_automation)
        self.btn_process.grid(row=4, column=0, columnspan=2, padx=20, pady=20, sticky="ew")

        # 4. Preview & Save
        self.textbox = ctk.CTkTextbox(self.control_frame, height=200)
        self.textbox.grid(row=5, column=0, columnspan=2, padx=20, pady=10, sticky="nsew")
        
        self.btn_save = ctk.CTkButton(self.control_frame, text="4. บันทึกผลลัพธ์ (Save CSV)", 
                                      command=self.save_csv, state="disabled")
        self.btn_save.grid(row=6, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

        self.control_frame.grid_columnconfigure(1, weight=1)
        self.control_frame.grid_rowconfigure(5, weight=1)

    def load_data_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if path:
            try:
                self.df_data = pd.read_csv(path)
                self.data_path = path
                self.lbl_data.configure(text=f"โหลดแล้ว: {path.split('/')[-1]} ({len(self.df_data)} แถว)")
                self.textbox.insert("end", f"-> Data Loaded: {len(self.df_data)} rows\n")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def load_rules_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if path:
            try:
                self.df_rules = pd.read_csv(path)
                # Check required columns
                required = {'name', 'start', 'length'}
                if not required.issubset(self.df_rules.columns):
                    messagebox.showerror("Error", f"ไฟล์กฎต้องมีคอลัมน์: {required}")
                    return
                
                self.lbl_rules.configure(text=f"โหลดแล้ว: {path.split('/')[-1]} ({len(self.df_rules)} กฎ)")
                self.textbox.insert("end", f"-> Rules Loaded: {len(self.df_rules)} rules found.\n")
                
                # Show rules preview
                self.textbox.insert("end", "--- ตัวอย่างกฎ ---\n")
                self.textbox.insert("end", self.df_rules.head().to_string() + "\n\n")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def process_automation(self):
        if self.df_data is None or self.df_rules is None:
            messagebox.showwarning("Warning", "กรุณาโหลดทั้งไฟล์ข้อมูลและไฟล์กฎก่อน")
            return

        hex_col = self.entry_hex_col.get()
        if hex_col not in self.df_data.columns:
            messagebox.showerror("Error", f"ไม่พบคอลัมน์ '{hex_col}' ในไฟล์ข้อมูล")
            return

        self.textbox.insert("end", "Running...\n")
        self.update()

        # สร้าง Dictionary เพื่อเก็บผลลัพธ์
        results = {name: [] for name in self.df_rules['name']}
        
        # วนลูปทุกแถวของข้อมูล (Data)
        for idx, row in self.df_data.iterrows():
            hex_str = str(row[hex_col]).strip().replace(" ", "")
            
            # --- Unescape JT808 (Optional but recommended) ---
            # ถ้าต้องการ unescape ให้เปิด comment ส่วนนี้
            # try:
            #     b = bytes.fromhex(hex_str).replace(b'\x7d\x02', b'\x7e').replace(b'\x7d\x01', b'\x7d')
            #     hex_str = b.hex().upper()
            # except: pass
            # ------------------------------------------------

            # วนลูปทุกกฎ (Rules) จากไฟล์ Config
            for _, rule in self.df_rules.iterrows():
                try:
                    # 1. Slice String
                    s = int(rule['start']) * 2
                    l = int(rule['length']) * 2
                    e = s + l
                    
                    val = None
                    if e <= len(hex_str):
                        raw_hex = hex_str[s:e]
                        
                        # 2. Convert Type
                        if rule['type'] == 'int':
                            val = int(raw_hex, 16)
                        elif rule['type'] == 'str':
                            val = bytes.fromhex(raw_hex).decode('gbk', errors='ignore')
                        elif rule['type'] == 'bcd':
                             val = raw_hex # BCD มักจะอ่านเป็น string ไปเลยแล้วค่อยไป format
                        else:
                            val = raw_hex # default

                        # ตัวแปร x ใช้สำหรับ eval()
                        x = val 

                        # 3. Check Condition (ถ้ามีเงื่อนไข)
                        condition_pass = True
                        if pd.notna(rule.get('condition')) and str(rule['condition']).strip() != "":
                            # ประเมินเงื่อนไข เช่น "x > 0"
                            try:
                                if not eval(str(rule['condition'])):
                                    condition_pass = False
                            except:
                                condition_pass = False # ถ้า error ถือว่าไม่ผ่าน

                        # 4. Apply Formula (ถ้าเงื่อนไขผ่าน หรือไม่มีเงื่อนไข)
                        if condition_pass:
                            if pd.notna(rule.get('formula')) and str(rule['formula']).strip() != "":
                                try:
                                    val = eval(str(rule['formula'])) # คำนวณสูตร
                                except Exception as err:
                                    val = f"Calc Err"
                        else:
                            val = None # หรือค่า Default กรณีไม่ผ่านเงื่อนไข

                    results[rule['name']].append(val)

                except Exception as ex:
                    results[rule['name']].append(f"Err: {ex}")

        # นำผลลัพธ์มารวมกับ DataFrame เดิม
        new_cols = pd.DataFrame(results)
        self.df_output = pd.concat([self.df_data, new_cols], axis=1)
        
        self.textbox.insert("end", "เสร็จสิ้น! ดูตัวอย่างผลลัพธ์:\n")
        self.textbox.insert("end", self.df_output.head().to_string() + "\n")
        self.btn_save.configure(state="normal")

    def save_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if path:
            self.df_output.to_csv(path, index=False)
            messagebox.showinfo("Saved", "บันทึกไฟล์เรียบร้อย")

if __name__ == "__main__":
    app = AutoDecoderApp()
    app.mainloop()