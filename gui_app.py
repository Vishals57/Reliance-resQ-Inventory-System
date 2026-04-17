import customtkinter as ctk
from tkinter import ttk, messagebox
from customtkinter import CTkInputDialog
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
        self.current_scanned_sr = None
        self.last_qr_path = None

        # Ensure DB exists + schema is up-to-date (adds new columns like In_Date/Out_Date)
        try:
            engine.initialize_db()
            engine.migrate_db()
        except Exception:
            pass

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
        self._apply_ttk_theme()
        self.refresh_all_data()

    def _apply_ttk_theme(self):
        """Simple professional dark style for ttk.Treeview widgets."""
        try:
            style = ttk.Style()
            try:
                style.theme_use("clam")
            except Exception:
                pass

            style.configure(
                "Treeview",
                background="#111114",
                fieldbackground="#111114",
                foreground="#e5e7eb",
                rowheight=28,
                bordercolor="#2a2a33",
                borderwidth=0,
                font=("Segoe UI", 10),
            )
            style.configure(
                "Treeview.Heading",
                background="#1b1b22",
                foreground="#9ca3af",
                relief="flat",
                font=("Segoe UI", 10, "bold"),
            )
            style.map("Treeview", background=[("selected", "#243b55")], foreground=[("selected", "#ffffff")])
        except Exception:
            pass

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

        # Right-side dashboard options
        self.dash_opts = ctk.CTkFrame(self.stats_frame, fg_color="transparent")
        self.dash_opts.pack(side="right", padx=10)
        self.highlight_closed_var = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(
            self.dash_opts,
            text="Highlight Closed (OUT) in green",
            variable=self.highlight_closed_var,
            command=self.refresh_all_data,
            onvalue=True,
            offvalue=False,
        ).pack(pady=(28, 0))

        # Search + filter
        self.dash_search = ctk.CTkEntry(self.dash_opts, placeholder_text="Search (ID / Item / Invoice / Engineer)")
        self.dash_search.pack(pady=(10, 0), padx=6)
        self.dash_search.bind("<KeyRelease>", lambda _e: self.refresh_all_data())

        self.dash_status_var = ctk.StringVar(value="ALL")
        ctk.CTkSegmentedButton(
            self.dash_opts,
            values=["ALL", "IN", "OUT"],
            variable=self.dash_status_var,
            command=lambda _v=None: self.refresh_all_data(),
        ).pack(pady=(10, 0), padx=6)

        self.tree_frame = ctk.CTkFrame(self.t_dash, fg_color="#16161a", corner_radius=15)
        self.tree_frame.pack(fill="both", expand=True, padx=25, pady=10)

        # Dashboard columns (show invoice + IN/OUT dates)
        self.tree = ttk.Treeview(
            self.tree_frame,
            columns=("ID", "Item", "Invoice", "Status", "User", "Charges", "IN Date", "OUT Date"),
            show="headings",
        )
        for col in ("ID", "Item", "Invoice", "Status", "User", "Charges", "IN Date", "OUT Date"):
            self.tree.heading(col, text=col)
            w = 160
            if col in ("Item", "Invoice"):
                w = 220
            if col in ("IN Date", "OUT Date"):
                w = 190
            self.tree.column(col, anchor="center", width=w)
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        # Row tags
        try:
            self.tree.tag_configure("closed", foreground="#22c55e")
        except Exception:
            pass

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

        self.btn_direct_print = ctk.CTkButton(self.reg_box, text="🖨️ DIRECT PRINT", height=40, fg_color="#585b70", state="disabled", command=self.direct_print_qr)
        self.btn_direct_print.pack(pady=(0, 10), padx=30, fill="x")

        # MASTER TABLE WITH 5 COLUMNS
        self.master_box = ctk.CTkFrame(self.asset_split, fg_color="#16161a", corner_radius=15)
        self.master_box.pack(side="right", fill="both", expand=True, padx=20, pady=20)
        self.master_tree = ttk.Treeview(self.master_box, columns=("ID", "Name", "CP", "SP", "Stock"), show="headings")
        for c in ("ID", "Name", "CP", "SP", "Stock"): 
            self.master_tree.heading(c, text=c)
            self.master_tree.column(c, anchor="center", width=140 if c in ("Name",) else 100)
        self.master_tree.pack(fill="both", expand=True, padx=10, pady=10)

    # --- 🔧 OPERATIONS ---
    def init_operations_tab(self):
        self.ops_split = ctk.CTkFrame(self.t_ops, fg_color="transparent")
        self.ops_split.pack(fill="both", expand=True)
        self.scan_pnl = ctk.CTkFrame(self.ops_split, fg_color="#24242b", corner_radius=15)
        self.scan_pnl.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        self.scan_pnl.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(self.scan_pnl, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=25, pady=(25, 10))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="Operations", font=("Segoe UI", 22, "bold"), text_color="#e5e7eb").grid(row=0, column=0, sticky="w")
        self.status_lbl = ctk.CTkLabel(
            header,
            text="READY TO SCAN",
            font=("Segoe UI", 12, "bold"),
            text_color="#9ca3af",
        )
        self.status_lbl.grid(row=1, column=0, sticky="w", pady=(6, 0))

        # Details card (fixes truncation: each field has its own row)
        self.details_card = ctk.CTkFrame(self.scan_pnl, fg_color="#1c1c21", corner_radius=14)
        self.details_card.grid(row=1, column=0, sticky="ew", padx=25, pady=(10, 15))
        self.details_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.details_card, text="Scanned Item Details", font=("Segoe UI", 14, "bold"), text_color="#e5e7eb").grid(
            row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(14, 10)
        )

        def _row(r, k):
            ctk.CTkLabel(self.details_card, text=k, font=("Segoe UI", 11, "bold"), text_color="#9ca3af").grid(
                row=r, column=0, sticky="w", padx=16, pady=6
            )
            v = ctk.CTkLabel(self.details_card, text="-", font=("Segoe UI", 12), text_color="#e5e7eb", anchor="w", justify="left", wraplength=650)
            v.grid(row=r, column=1, sticky="ew", padx=(0, 16), pady=6)
            return v

        self.v_art = _row(1, "Article Code")
        self.v_sr = _row(2, "Sr No")
        self.v_part = _row(3, "Part Name")
        self.v_inv = _row(4, "Invoice No")
        self.v_sp = _row(5, "Selling Price")

        actions = ctk.CTkFrame(self.scan_pnl, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", padx=25, pady=(0, 10))
        actions.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(actions, text="📸 ACTIVATE CAMERA", height=52, fg_color="#3b82f6", command=self.run_scan).grid(
            row=0, column=0, sticky="ew"
        )

        self.move_type_var = ctk.StringVar(value="Company")
        ctk.CTkSegmentedButton(actions, values=["Company", "Local"], variable=self.move_type_var).grid(
            row=1, column=0, sticky="ew", pady=(12, 0)
        )

        btn_row = ctk.CTkFrame(actions, fg_color="transparent")
        btn_row.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        btn_row.grid_columnconfigure(0, weight=1)
        btn_row.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(btn_row, text="📥 INWARD", height=44, fg_color="#10b981", command=lambda: self.execute_move("IN")).grid(
            row=0, column=0, sticky="ew", padx=(0, 8)
        )
        ctk.CTkButton(btn_row, text="📤 OUTWARD", height=44, fg_color="#f59e0b", command=lambda: self.execute_move("OUT")).grid(
            row=0, column=1, sticky="ew", padx=(8, 0)
        )

        quick = ctk.CTkFrame(actions, fg_color="transparent")
        quick.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        quick.grid_columnconfigure(0, weight=1)
        quick.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(quick, text="🧹 CLEAR", height=38, fg_color="#374151", command=self.clear_scan).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(quick, text="🖨️ REPRINT QR", height=38, fg_color="#585b70", command=self.direct_print_qr).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        
        self.eng_wrap = ctk.CTkFrame(self.ops_split, width=350, fg_color="transparent")
        self.eng_wrap.pack(side="right", fill="y", padx=20, pady=20)

        eng_hdr = ctk.CTkFrame(self.eng_wrap, fg_color="transparent")
        eng_hdr.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(eng_hdr, text="Engineers", font=("Segoe UI", 16, "bold"), text_color="#e5e7eb").pack(side="left")
        ctk.CTkButton(eng_hdr, text="+ Add", width=70, height=32, fg_color="#3b82f6", command=self.add_engineer_ui).pack(side="right", padx=(6, 0))
        ctk.CTkButton(eng_hdr, text="− Remove", width=90, height=32, fg_color="#ef4444", command=self.remove_engineer_ui).pack(side="right")

        self.eng_search = ctk.CTkEntry(self.eng_wrap, placeholder_text="Search engineer / BP / PPRR")
        self.eng_search.pack(fill="x", padx=6, pady=(0, 8))
        self.eng_search.bind("<KeyRelease>", lambda _e: self.refresh_engineers())

        self.eng_pnl = ctk.CTkScrollableFrame(self.eng_wrap, width=350, label_text="Select Engineer")
        self.eng_pnl.pack(fill="both", expand=True)
        self.eng_btns = {}
        self.refresh_engineers()

    def clear_scan(self):
        self.current_scanned_id = None
        self.current_scanned_sr = None
        self.status_lbl.configure(text="READY TO SCAN", text_color="#9ca3af")
        self.v_art.configure(text="-")
        self.v_sr.configure(text="-")
        self.v_part.configure(text="-")
        self.v_inv.configure(text="-")
        self.v_sp.configure(text="-")

    def refresh_engineers(self):
        try:
            for w in self.eng_pnl.winfo_children():
                w.destroy()
        except Exception:
            pass
        self.eng_btns = {}
        engineers = engine.get_engineers(active_only=True)
        q = ""
        try:
            q = (self.eng_search.get() or "").strip().lower()
        except Exception:
            q = ""

        for row in engineers:
            name = row.get("Engineer_Name", "")
            bp = row.get("BP_ID", "") or "-"
            pprr = row.get("PPRR_ID", "") or "-"
            label = f"{name}\nBP: {bp}   |   PPRR: {pprr}"

            if q:
                hay = f"{name} {bp} {pprr}".lower()
                if q not in hay:
                    continue

            b = ctk.CTkButton(
                self.eng_pnl,
                text=label,
                height=52,
                fg_color="#2d2d35",
                font=("Segoe UI", 12, "bold"),
                text_color="#e5e7eb",
                anchor="w",
                command=lambda e=name: self.set_eng(e),
            )
            b.pack(pady=5, padx=10, fill="x")
            self.eng_btns[name] = b

    def add_engineer_ui(self):
        try:
            name = (CTkInputDialog(text="Engineer Name", title="Add Engineer").get_input() or "").strip()
            if not name:
                return
            bp = (CTkInputDialog(text="BP ID (optional)", title="Add Engineer").get_input() or "").strip()
            pprr = (CTkInputDialog(text="PPRR ID (optional)", title="Add Engineer").get_input() or "").strip()
            ok, msg = engine.add_engineer(name, bp_id=bp, pprr_id=pprr)
            if ok:
                messagebox.showinfo("OK", msg)
                self.refresh_engineers()
            else:
                messagebox.showerror("Err", msg)
        except Exception as e:
            messagebox.showerror("Err", str(e))

    def remove_engineer_ui(self):
        # Remove currently selected engineer (simple + safe)
        if not self.selected_engineer:
            return messagebox.showwarning("!", "Select Engineer First")
        ok, msg = engine.remove_engineer(self.selected_engineer)
        if ok:
            messagebox.showinfo("OK", msg)
            self.selected_engineer = None
            self.refresh_engineers()
        else:
            messagebox.showerror("Err", msg)

    # --- CORE REFRESH LOGIC (THE FIX) ---
    def refresh_all_data(self):
        # 1. Update Dashboard
        for i in self.tree.get_children(): self.tree.delete(i)
        if os.path.exists(engine.DB_FILE):
            try:
                df_t = pd.read_excel(engine.DB_FILE, sheet_name='Transactions')
                df_m = pd.read_excel(engine.DB_FILE, sheet_name='Master')
                if "Sr_No" not in df_t.columns:
                    df_t["Sr_No"] = 1
                if "Tax_Invoice_No" not in df_t.columns:
                    df_t["Tax_Invoice_No"] = ""
                if "In_Date" not in df_t.columns:
                    df_t["In_Date"] = df_t["Date"] if "Date" in df_t.columns else ""
                if "Out_Date" not in df_t.columns:
                    df_t["Out_Date"] = df_t["Date"] if "Date" in df_t.columns else ""
                df = pd.merge(df_t, df_m[['Article_No', 'Part_Name', 'SP']], on='Article_No', how='left')

                # Dashboard filters
                q = ""
                try:
                    q = (self.dash_search.get() or "").strip().lower()
                except Exception:
                    q = ""
                status_filter = "ALL"
                try:
                    status_filter = (self.dash_status_var.get() or "ALL").strip().upper()
                except Exception:
                    status_filter = "ALL"

                for _, r in df.iloc[::-1].iterrows():
                    # EXACTLY 6 VALUES FOR 6 DASHBOARD COLUMNS
                    try:
                        sr = int(r.get("Sr_No", 1))
                    except Exception:
                        sr = 1
                    display_id = f"{r['Article_No']}-{sr}" if sr != 1 or (df_t[df_t["Article_No"] == r["Article_No"]]["Sr_No"].nunique() > 1) else str(r["Article_No"])
                    inv = r.get("Tax_Invoice_No", "")
                    inv = "-" if inv is None or str(inv).strip() == "" or str(inv) == "nan" else str(inv)
                    in_d = r.get("In_Date", "")
                    out_d = r.get("Out_Date", "")
                    in_d = "-" if in_d is None or str(in_d).strip() == "" or str(in_d) == "nan" else str(in_d)
                    out_d = "-" if out_d is None or str(out_d).strip() == "" or str(out_d) == "nan" else str(out_d)
                    status = str(r.get("Status", "") or "").strip().upper()
                    if status_filter in ("IN", "OUT") and status != status_filter:
                        continue

                    if q:
                        hay = f"{display_id} {r.get('Part_Name','')} {inv} {r.get('Engineer','')}".lower()
                        if q not in hay:
                            continue

                    tags = ()
                    if getattr(self, "highlight_closed_var", None) is not None:
                        if bool(self.highlight_closed_var.get()) and status == "OUT":
                            tags = ("closed",)

                    self.tree.insert(
                        "",
                        "end",
                        values=(
                            display_id,
                            r.get("Part_Name", ""),
                            inv,
                            status,
                            r.get("Engineer", ""),
                            f"₹{r.get('Charges', r.get('SP', ''))}",
                            in_d,
                            out_d,
                        ),
                        tags=tags,
                    )
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
                # One QR per Article_No, mobile-readable (full details available at registration time)
                # Invoice/Qty are not known at registration, so keep invoice as "-" and qty as 1.
                self.last_qr_path = gen.generate_box_qr(
                    art_no=art,
                    part_name=name,
                    tax_invoice_no="-",
                    sp=sp,
                    qty=1,
                    sr_list=None,
                )
                messagebox.showinfo("OK", "Registered!"); self.refresh_all_data()
                if self.last_qr_path and os.path.exists(self.last_qr_path):
                    img = Image.open(self.last_qr_path).resize((130, 130))
                    self.tk_qr = ImageTk.PhotoImage(img)
                    self.qr_img_label.configure(image=self.tk_qr, text="")
                    self.btn_open_qr.configure(state="normal", fg_color="#10b981")
                    self.btn_direct_print.configure(state="normal", fg_color="#3b82f6")
            else: messagebox.showerror("Err", msg)
        except: messagebox.showerror("Err", "Prices must be numbers")

    def open_qr_file(self):
        if self.last_qr_path: os.startfile(os.path.abspath(self.last_qr_path))

    def direct_print_qr(self):
        """
        Sends the QR PNG to the Windows default printer.
        Note: Windows uses the default associated app for PNG printing.
        """
        if not self.last_qr_path:
            return
        try:
            os.startfile(os.path.abspath(self.last_qr_path), "print")
        except Exception as e:
            messagebox.showerror("Print Error", f"Could not print.\n{e}")

    def execute_move(self, t):
        if not self.current_scanned_id: return messagebox.showwarning("!", "Scan QR First")

        # IN: allow multiple physical units via Qty -> Sr_No auto assignment
        if t.upper() == "IN":
            qty = 1
            try:
                d = CTkInputDialog(text="Enter Qty for inward (default 1)", title="Inward Qty")
                v = (d.get_input() or "").strip()
                qty = int(v) if v else 1
            except Exception:
                qty = 1
            if qty < 1:
                qty = 1

            tax_invoice_no = ""
            try:
                d2 = CTkInputDialog(text="Enter Tax Invoice No (optional)", title="Tax Invoice No")
                tax_invoice_no = (d2.get_input() or "").strip()
            except Exception:
                tax_invoice_no = ""

            # Predict Sr_No allocation (matches engine auto allocation)
            sr_list = engine.get_next_inward_sr_list(self.current_scanned_id, qty)

            s, m = engine.process_movement(
                self.current_scanned_id,
                t,
                purchase_type=self.move_type_var.get(),
                engineer=self.selected_engineer,
                qty=qty,
                tax_invoice_no=tax_invoice_no,
            )
            if s:
                # Generate ONE box QR (mobile-readable) for the whole inward
                details = engine.get_scan_details(self.current_scanned_id)
                part = details.get("Part_Name", "")
                sp = details.get("SP", None)
                box_path = gen.generate_box_qr(self.current_scanned_id, part, tax_invoice_no, sp, qty, sr_list=sr_list)
                if box_path:
                    self.last_qr_path = box_path
                    try:
                        img = Image.open(self.last_qr_path).resize((130, 130))
                        self.tk_qr = ImageTk.PhotoImage(img)
                        self.qr_img_label.configure(image=self.tk_qr, text="")
                        self.btn_open_qr.configure(state="normal", fg_color="#10b981")
                        self.btn_direct_print.configure(state="normal", fg_color="#3b82f6")
                    except Exception:
                        pass

                messagebox.showinfo("OK", m); self.refresh_all_data()
            else: messagebox.showerror("Err", m)
            return

        # OUT: if multiple units IN, ask Sr_No to outward specific unit
        if t.upper() == "OUT":
            if not self.selected_engineer:
                return messagebox.showwarning("!", "Select Engineer First")

            available = engine.get_available_sr_nos(self.current_scanned_id)
            sr_no = None
            if self.current_scanned_sr is not None and self.current_scanned_sr in available:
                sr_no = self.current_scanned_sr
            if len(available) > 1:
                try:
                    d = CTkInputDialog(
                        text=f"Multiple units IN.\nAvailable Sr No: {available}\nEnter Sr No to outward:",
                        title="Select Sr No",
                    )
                    v = (d.get_input() or "").strip()
                    sr_no = int(v) if v else (sr_no if sr_no is not None else available[0])  # auto-pick
                except Exception:
                    sr_no = sr_no if sr_no is not None else available[0]
            elif len(available) == 1:
                sr_no = available[0]

            s, m = engine.process_movement(
                self.current_scanned_id,
                t,
                purchase_type=self.move_type_var.get(),
                engineer=self.selected_engineer,
                sr_no=sr_no,
            )
            if s: messagebox.showinfo("OK", m); self.refresh_all_data()
            else: messagebox.showerror("Err", m)
            return

        s, m = engine.process_movement(self.current_scanned_id, t, purchase_type=self.move_type_var.get(), engineer=self.selected_engineer)
        if s: messagebox.showinfo("OK", m); self.refresh_all_data()

    def run_scan(self):
        res, _ = scanner.activate_scanner()
        if not res:
            return

        self.current_scanned_id = res
        self.current_scanned_sr = None

        d = engine.get_scan_details(res)
        in_units = d.get("in_units", []) or []

        # If multiple units are IN, allow choosing which Sr_No to show/use
        if len(in_units) > 1:
            available = [u["Sr_No"] for u in in_units]
            try:
                pick = CTkInputDialog(
                    text=f"Multiple units found IN.\nAvailable Sr No: {available}\nEnter Sr No to view/use (optional):",
                    title="Select Sr No",
                ).get_input()
                pick = (pick or "").strip()
                if pick:
                    self.current_scanned_sr = int(pick)
            except Exception:
                self.current_scanned_sr = None
        elif len(in_units) == 1:
            self.current_scanned_sr = int(in_units[0]["Sr_No"])

        # Pick unit info for display (selected Sr_No if present; otherwise first IN unit)
        unit = None
        if self.current_scanned_sr is not None:
            for u in in_units:
                if int(u["Sr_No"]) == int(self.current_scanned_sr):
                    unit = u
                    break
        if unit is None and in_units:
            unit = in_units[0]

        art = d.get("Article_No", res)
        part = d.get("Part_Name", "") or ""
        sp = d.get("SP", None)
        sp_txt = f"₹{sp}" if sp is not None and str(sp) != "nan" else "₹-"

        sr_txt = str(unit["Sr_No"]) if unit else "-"
        inv_txt = (unit.get("Tax_Invoice_No", "") if unit else "") or "-"

        self.status_lbl.configure(
            text="SCAN SUCCESSFUL",
            text_color="#10b981",
        )
        self.v_art.configure(text=str(art))
        self.v_sr.configure(text=str(sr_txt))
        self.v_part.configure(text=str(part) if part else "-")
        self.v_inv.configure(text=str(inv_txt) if inv_txt else "-")
        self.v_sp.configure(text=str(sp_txt))

    def set_eng(self, n):
        for b in self.eng_btns.values(): b.configure(fg_color="#2d2d35")
        self.eng_btns[n].configure(fg_color="#3b82f6"); self.selected_engineer = n

    def add_input(self, parent, placeholder):
        e = ctk.CTkEntry(parent, placeholder_text=placeholder, height=35, fg_color="#1c1c21", border_color="#3b82f6")
        e.pack(pady=5, padx=30, fill="x"); return e

if __name__ == "__main__":
    app = ResQUltimateAdmin()
    app.mainloop()