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
        self.title("resQ Enterprise - Asset Studio v5.7(vishalshinde)")
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

        # 2. Sidebar
        self.sidebar = ctk.CTkFrame(self, width=100, corner_radius=0, fg_color="#16161a")
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        ctk.CTkLabel(self.sidebar, text="RQ", font=("Arial", 28, "bold"), text_color="#3b82f6").pack(pady=40)
        
        # 3. Container & Tabs
        self.container = ctk.CTkFrame(self, fg_color="#1c1c21", corner_radius=20)
        self.container.grid(row=0, column=1, sticky="nsew", padx=25, pady=25)
        self.tabs = ctk.CTkTabview(self.container, fg_color="transparent")
        self.tabs.pack(fill="both", expand=True)
        
        self.t_dash = self.tabs.add("Dashboard")
        self.t_ops = self.tabs.add("Operations")
        self.t_assets = self.tabs.add("Asset Manager")
        self.tabs._segmented_button.grid_forget()

        # 4. Navigation
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

    # --- 📊 DASHBOARD TAB ---
    def init_dashboard_tab(self):
        # Stats Cards
        self.stats_frame = ctk.CTkFrame(self.t_dash, fg_color="transparent")
        self.stats_frame.pack(fill="x", padx=30, pady=15)

        self.val_card = self.create_stat_card(self.stats_frame, "TOTAL INVENTORY VALUE", "₹ 0.00", "#1e1e2e")
        self.cnt_card = self.create_stat_card(self.stats_frame, "TOTAL ITEMS", "0", "#1e1e2e")

        self.tree_frame = ctk.CTkFrame(self.t_dash, fg_color="#16161a", corner_radius=15)
        self.tree_frame.pack(fill="both", expand=True, padx=25, pady=10)

        # DEFINED 6 COLUMNS
        self.tree = ttk.Treeview(self.tree_frame, columns=("ID", "Item", "Status", "User", "Charges", "Date"), show="headings")
        for col in ("ID", "Item", "Status", "User", "Charges", "Date"):
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=120)
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

    def create_stat_card(self, parent, title, val, color):
        card = ctk.CTkFrame(parent, fg_color=color, corner_radius=10, width=280, height=100)
        card.pack(side="left", padx=10)
        card.pack_propagate(False)
        ctk.CTkLabel(card, text=title, font=("Segoe UI", 11, "bold"), text_color="#3b82f6").pack(pady=(15, 0))
        lbl = ctk.CTkLabel(card, text=val, font=("Segoe UI", 24, "bold"))
        lbl.pack()
        return lbl

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
        
        ctk.CTkButton(self.reg_box, text="SAVE & GENERATE QR", height=45, fg_color="#3b82f6", command=self.save_article).pack(pady=15, padx=30, fill="x")

        self.qr_img_label = ctk.CTkLabel(self.reg_box, text="QR Preview Area")
        self.qr_img_label.pack(pady=10)

        self.btn_open_qr = ctk.CTkButton(self.reg_box, text="🖨️ OPEN TO PRINT", height=40, fg_color="#585b70", state="disabled", command=self.open_qr_file)
        self.btn_open_qr.pack(pady=10, padx=30, fill="x")

        # MASTER TABLE WITH 5 COLUMNS
        self.master_box = ctk.CTkFrame(self.asset_split, fg_color="#16161a", corner_radius=15)
        self.master_box.pack(side="right", fill="both", expand=True, padx=20, pady=20)
        self.master_tree = ttk.Treeview(self.master_box, columns=("ID", "Name", "CP", "SP", "Stock"), show="headings")
        for c in ("ID", "Name", "CP", "SP", "Stock"): 
            self.master_tree.heading(c, text=c)
            self.master_tree.column(c, anchor="center", width=90)
        self.master_tree.pack(fill="both", expand=True, padx=10, pady=10)

    # --- 🔧 OPERATIONS ---
    def init_operations_tab(self):
        self.ops_split = ctk.CTkFrame(self.t_ops, fg_color="transparent")
        self.ops_split.pack(fill="both", expand=True)
        self.scan_pnl = ctk.CTkFrame(self.ops_split, fg_color="#24242b", corner_radius=15)
        self.scan_pnl.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        self.status_lbl = ctk.CTkLabel(self.scan_pnl, text="READY TO SCAN", font=("Segoe UI", 20, "bold"))
        self.status_lbl.pack(pady=40)
        ctk.CTkButton(self.scan_pnl, text="📸 ACTIVATE CAMERA", height=60, fg_color="#3b82f6", command=self.run_scan).pack(pady=20, padx=60, fill="x")
        
        self.move_type_var = ctk.StringVar(value="Company")
        ctk.CTkSegmentedButton(self.scan_pnl, values=["Company", "Local"], variable=self.move_type_var).pack(pady=10, padx=60, fill="x")
        
        ctk.CTkButton(self.scan_pnl, text="📥 INWARD", height=45, fg_color="#10b981", command=lambda: self.execute_move("IN")).pack(pady=10, padx=60, fill="x")
        ctk.CTkButton(self.scan_pnl, text="📤 OUTWARD", height=45, fg_color="#f59e0b", command=lambda: self.execute_move("OUT")).pack(pady=10, padx=60, fill="x")
        
        self.eng_pnl = ctk.CTkScrollableFrame(self.ops_split, width=350, label_text="Select Engineer")
        self.eng_pnl.pack(side="right", fill="y", padx=20, pady=20)
        self.eng_btns = {}
        for eng in engine.ENGINEERS:
            b = ctk.CTkButton(self.eng_pnl, text=eng, height=40, fg_color="#2d2d35", command=lambda e=eng: self.set_eng(e))
            b.pack(pady=5, padx=10, fill="x")
            self.eng_btns[eng] = b

    # --- CORE REFRESH LOGIC (THE FIX) ---
    def refresh_all_data(self):
        # 1. Update Dashboard
        for i in self.tree.get_children(): self.tree.delete(i)
        if os.path.exists(engine.DB_FILE):
            try:
                df_t = pd.read_excel(engine.DB_FILE, sheet_name='Transactions')
                df_m = pd.read_excel(engine.DB_FILE, sheet_name='Master')
                df = pd.merge(df_t, df_m[['Article_No', 'Part_Name', 'SP']], on='Article_No', how='left')
                for _, r in df.iloc[::-1].iterrows():
                    # EXACTLY 6 VALUES FOR 6 DASHBOARD COLUMNS
                    self.tree.insert("", "end", values=(r['Article_No'], r['Part_Name'], r['Status'], r['Engineer'], f"₹{r['SP']}", r['Date']))
            except Exception as e: print(f"Dash Error: {e}")

        # 2. Update Asset Catalog
        for i in self.master_tree.get_children(): self.master_tree.delete(i)
        if os.path.exists(engine.DB_FILE):
            try:
                df_m = pd.read_excel(engine.DB_FILE, sheet_name='Master')
                for _, r in df_m.iterrows():
                    # EXACTLY 5 VALUES FOR 5 MASTER COLUMNS (Fixing the Stock/CP swap)
                    self.master_tree.insert("", "end", values=(
                        r['Article_No'], 
                        r['Part_Name'], 
                        f"₹{r['CP']}", 
                        f"₹{r['SP']}", 
                        int(r['Stock_Level'])
                    ))
                
                # Update Cards
                total_val = (df_m['Stock_Level'] * df_m['SP']).sum()
                self.val_card.configure(text=f"₹ {total_val:,.2f}")
                self.cnt_card.configure(text=str(int(df_m['Stock_Level'].sum())))
            except Exception as e: print(f"Master Error: {e}")

    # --- LOGIC HELPERS ---
    def save_article(self):
        try:
            art, name = self.e_id.get(), self.e_name.get()
            cp, sp = float(self.e_cp.get()), float(self.e_sp.get())
            s, msg = engine.register_new_article(art, name, cp, sp)
            if s:
                gen.generate_article_qr(art, name)
                self.last_qr_path = f"Article_QRs/QR_{art}.png"
                img = Image.open(self.last_qr_path).resize((130, 130))
                self.tk_qr = ImageTk.PhotoImage(img)
                self.qr_img_label.configure(image=self.tk_qr, text="")
                self.btn_open_qr.configure(state="normal", fg_color="#10b981")
                messagebox.showinfo("OK", "Registered!"); self.refresh_all_data()
            else: messagebox.showerror("Err", msg)
        except: messagebox.showerror("Err", "Prices must be numbers")

    def open_qr_file(self):
        if self.last_qr_path: os.startfile(os.path.abspath(self.last_qr_path))

    def execute_move(self, t):
        if not self.current_scanned_id: return messagebox.showwarning("!", "Scan QR First")
        s, m = engine.process_movement(self.current_scanned_id, t, purchase_type=self.move_type_var.get(), engineer=self.selected_engineer)
        if s: messagebox.showinfo("OK", m); self.refresh_all_data()

    def run_scan(self):
        res, _ = scanner.activate_scanner()
        if res: self.current_scanned_id = res; self.status_lbl.configure(text=f"DETECTED: {res}", text_color="#3b82f6")

    def set_eng(self, n):
        for b in self.eng_btns.values(): b.configure(fg_color="#2d2d35")
        self.eng_btns[n].configure(fg_color="#3b82f6"); self.selected_engineer = n

    def add_input(self, parent, placeholder):
        e = ctk.CTkEntry(parent, placeholder_text=placeholder, height=35, fg_color="#1c1c21", border_color="#3b82f6")
        e.pack(pady=5, padx=30, fill="x"); return e

if __name__ == "__main__":
    app = ResQUltimateAdmin()
    app.mainloop()