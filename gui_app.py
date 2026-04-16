import customtkinter as ctk
from tkinter import ttk, messagebox
import inventory_engine as engine
import scanner_interface as scanner
import qr_generator as gen
from PIL import Image, ImageTk
import pandas as pd
import os

class ResQUltimateAdmin(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("resQ Enterprise - Asset Studio v5.6(vishalshinde)")
        self.geometry("1400x850")
        ctk.set_appearance_mode("Dark")
        self.configure(fg_color="#0f0f12")

        # Variables
        self.selected_engineer = None
        self.current_scanned_id = None
        self.last_qr_path = None

        # 1. Setup Layout
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 2. Sidebar Navigation
        self.sidebar = ctk.CTkFrame(self, width=100, corner_radius=0, fg_color="#16161a")
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        ctk.CTkLabel(self.sidebar, text="RQ", font=("Arial", 28, "bold"), text_color="#3b82f6").pack(pady=40)
        
        # 3. Main Container and Tabs
        self.container = ctk.CTkFrame(self, fg_color="#1c1c21", corner_radius=20)
        self.container.grid(row=0, column=1, sticky="nsew", padx=25, pady=25)
        
        self.tabs = ctk.CTkTabview(self.container, fg_color="transparent")
        self.tabs.pack(fill="both", expand=True)
        
        # CRITICAL: Define Tabs before initializing UI content
        self.t_dash = self.tabs.add("Dashboard")
        self.t_ops = self.tabs.add("Operations")
        self.t_assets = self.tabs.add("Asset Manager")
        self.tabs._segmented_button.grid_forget() # Modern look (sidebar controlled)

        # 4. Create Sidebar Buttons
        self.create_nav_btn("📊", "Dashboard")
        self.create_nav_btn("🔧", "Operations")
        self.create_nav_btn("📁", "Asset Manager")

        # 5. Initialize Content
        self.init_dashboard_tab()
        self.init_operations_tab()
        self.init_assets_tab()
        
        self.tabs.set("Dashboard")
        self.refresh_all_data()

    def create_nav_btn(self, icon, name):
        btn = ctk.CTkButton(self.sidebar, text=icon, width=65, height=65, fg_color="transparent", 
                            font=("Arial", 24), hover_color="#2d2d35", command=lambda: self.tabs.set(name))
        btn.pack(pady=10)

    # --- 📁 ASSET MANAGER ---
    def init_assets_tab(self):
        self.asset_split = ctk.CTkFrame(self.t_assets, fg_color="transparent")
        self.asset_split.pack(fill="both", expand=True)

        self.reg_box = ctk.CTkFrame(self.asset_split, width=380, fg_color="#24242b", corner_radius=15)
        self.reg_box.pack(side="left", fill="y", padx=20, pady=20)
        
        ctk.CTkLabel(self.reg_box, text="Register New Asset", font=("Segoe UI", 18, "bold")).pack(pady=15)
        self.e_id = self.add_input(self.reg_box, "Article Number")
        self.e_name = self.add_input(self.reg_box, "Part Name")
        self.e_cp = self.add_input(self.reg_box, "Cost Price")
        self.e_sp = self.add_input(self.reg_box, "Selling Price")
        
        ctk.CTkLabel(self.reg_box, text="Category:", font=("Segoe UI", 12)).pack(pady=(5,0))
        self.reg_type_var = ctk.StringVar(value="OG")
        self.reg_type_toggle = ctk.CTkSegmentedButton(self.reg_box, values=["OG", "Local"], variable=self.reg_type_var)
        self.reg_type_toggle.pack(pady=10, padx=30, fill="x")
        
        ctk.CTkButton(self.reg_box, text="SAVE & GENERATE QR", height=45, fg_color="#3b82f6", font=("Segoe UI", 13, "bold"), command=self.save_article).pack(pady=10, padx=30, fill="x")

        # QR Preview Area
        self.qr_preview_frame = ctk.CTkFrame(self.reg_box, fg_color="#16161a", height=180, corner_radius=10)
        self.qr_preview_frame.pack(pady=10, padx=30, fill="x")
        self.qr_img_label = ctk.CTkLabel(self.qr_preview_frame, text="QR Preview Area")
        self.qr_img_label.pack(expand=True, pady=10)

        # Print Button
        self.btn_open_qr = ctk.CTkButton(self.reg_box, text="🖨️ OPEN TO PRINT", height=40, fg_color="#585b70", state="disabled", command=self.open_qr_file)
        self.btn_open_qr.pack(pady=10, padx=30, fill="x")

        # Catalog List
        self.master_box = ctk.CTkFrame(self.asset_split, fg_color="#16161a", corner_radius=15)
        self.master_box.pack(side="right", fill="both", expand=True, padx=20, pady=20)
        self.master_tree = ttk.Treeview(self.master_box, columns=("ID", "Name", "Stock"), show="headings")
        for c in ("ID", "Name", "Stock"): 
            self.master_tree.heading(c, text=c)
            self.master_tree.column(c, anchor="center")
        self.master_tree.pack(fill="both", expand=True, padx=10, pady=10)

    # --- 🔧 OPERATIONS ---
    def init_operations_tab(self):
        self.ops_split = ctk.CTkFrame(self.t_ops, fg_color="transparent")
        self.ops_split.pack(fill="both", expand=True)
        
        # Left Panel: Scanner
        self.scan_pnl = ctk.CTkFrame(self.ops_split, fg_color="#24242b", corner_radius=15)
        self.scan_pnl.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        
        self.status_lbl = ctk.CTkLabel(self.scan_pnl, text="READY TO SCAN", font=("Segoe UI", 20, "bold"), text_color="#585b70")
        self.status_lbl.pack(pady=40)
        
        ctk.CTkButton(self.scan_pnl, text="📸 ACTIVATE CAMERA", height=60, fg_color="#3b82f6", command=self.run_scan).pack(pady=20, padx=60, fill="x")
        
        ctk.CTkLabel(self.scan_pnl, text="Select Purchase Type:", font=("Segoe UI", 12)).pack(pady=(10,0))
        self.move_type_var = ctk.StringVar(value="Company")
        self.move_type_toggle = ctk.CTkSegmentedButton(self.scan_pnl, values=["Company", "Local"], variable=self.move_type_var)
        self.move_type_toggle.pack(pady=10, padx=60, fill="x")

        ctk.CTkButton(self.scan_pnl, text="📥 INWARD", height=45, fg_color="#10b981", command=lambda: self.execute_move("IN")).pack(pady=10, padx=60, fill="x")
        ctk.CTkButton(self.scan_pnl, text="📤 OUTWARD", height=45, fg_color="#f59e0b", command=lambda: self.execute_move("OUT")).pack(pady=10, padx=60, fill="x")
        
        # RIGHT PANEL: CHANGED TO SCROLLABLE FRAME
        self.eng_pnl = ctk.CTkScrollableFrame(self.ops_split, width=350, fg_color="#16161a", corner_radius=15, label_text="Select Engineer", label_font=("Segoe UI", 16, "bold"))
        self.eng_pnl.pack(side="right", fill="y", padx=20, pady=20)
        
        self.eng_btns = {}
        # This will now create all 21 buttons inside the scrollable area
        for eng in engine.ENGINEERS:
            b = ctk.CTkButton(self.eng_pnl, text=eng, height=45, fg_color="#2d2d35", command=lambda e=eng: self.set_eng(e))
            b.pack(pady=5, padx=10, fill="x")
            self.eng_btns[eng] = b
            
    # --- LOGIC ---
    def save_article(self):
        try:
            art, name = self.e_id.get(), self.e_name.get()
            cp, sp = float(self.e_cp.get()), float(self.e_sp.get())
            cat = self.reg_type_var.get()
            
            s, msg = engine.register_new_article(art, name, cp, sp, cat=cat)
            if s:
                gen.generate_article_qr(art, name)
                self.last_qr_path = f"Article_QRs/QR_{art}.png"
                img = Image.open(self.last_qr_path).resize((150, 150))
                self.tk_qr = ImageTk.PhotoImage(img)
                self.qr_img_label.configure(image=self.tk_qr, text="")
                
                # Enable Print Button
                self.btn_open_qr.configure(state="normal", fg_color="#10b981")
                
                messagebox.showinfo("OK", "Registered & QR Generated!"); self.refresh_all_data()
            else: messagebox.showerror("Error", msg)
        except: messagebox.showerror("Input Error", "Check Prices")

    def open_qr_file(self):
        if self.last_qr_path: os.startfile(os.path.abspath(self.last_qr_path))

    def execute_move(self, t):
        if not self.current_scanned_id:
            messagebox.showwarning("!", "Scan QR First")
            return
        p_type = self.move_type_var.get()
        s, m = engine.process_movement(self.current_scanned_id, t, purchase_type=p_type, engineer=self.selected_engineer)
        if s: messagebox.showinfo("OK", m); self.refresh_all_data()
        else: messagebox.showerror("Err", m)

    def run_scan(self):
        res, _ = scanner.activate_scanner()
        if res: self.current_scanned_id = res; self.status_lbl.configure(text=f"DETECTED: {res}", text_color="#3b82f6")

    def init_dashboard_tab(self):
        self.tree = ttk.Treeview(self.t_dash, columns=("ID", "Item", "Type", "Status", "User", "Date"), show="headings")
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col); self.tree.column(col, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=20, pady=20)

    def set_eng(self, n):
        for b in self.eng_btns.values(): b.configure(fg_color="#2d2d35")
        self.eng_btns[n].configure(fg_color="#3b82f6")
        self.selected_engineer = n

    def add_input(self, parent, placeholder):
        e = ctk.CTkEntry(parent, placeholder_text=placeholder, height=35, fg_color="#1c1c21", border_color="#3b82f6")
        e.pack(pady=5, padx=30, fill="x")
        return e

    def refresh_all_data(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        if os.path.exists(engine.DB_FILE):
            try:
                df_t = pd.read_excel(engine.DB_FILE, sheet_name='Transactions')
                df_m = pd.read_excel(engine.DB_FILE, sheet_name='Master')
                df = pd.merge(df_t, df_m[['Article_No', 'Part_Name']], on='Article_No', how='left')
                for _, r in df.iloc[::-1].iterrows():
                    self.tree.insert("", "end", values=(r['Article_No'], r['Part_Name'], r['Purchase_Type'], r['Status'], r['Engineer'], r['Date']))
            except: pass
        for i in self.master_tree.get_children(): self.master_tree.delete(i)
        if os.path.exists(engine.DB_FILE):
            try:
                df_m = pd.read_excel(engine.DB_FILE, sheet_name='Master')
                for _, r in df_m.iterrows():
                    self.master_tree.insert("", "end", values=(r['Article_No'], r['Part_Name'], r['Stock_Level']))
            except: pass

if __name__ == "__main__":
    app = ResQUltimateAdmin()
    app.mainloop()