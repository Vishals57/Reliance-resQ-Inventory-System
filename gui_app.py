import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from customtkinter import CTkInputDialog
import inventory_engine as engine
import scanner_interface as scanner
import qr_generator as gen
from PIL import Image, ImageTk
import pandas as pd
import os
import threading
from datetime import datetime, timedelta
try:
    from tkcalendar import Calendar
    HAS_CALENDAR = True
except ImportError:
    HAS_CALENDAR = False

class ResQUltimateAdmin(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("resQ Enterprise - Asset Studio v5.8")
        self.geometry("1500x900")
        self.minsize(1200, 700)
        ctk.set_appearance_mode("Dark")
        self.configure(fg_color="#0a0a0e")
        
        # Color palette for better consistency
        self.colors = {
            "bg_primary": "#0a0a0e",
            "bg_secondary": "#16161a",
            "bg_tertiary": "#1f1f26",
            "accent": "#3b82f6",
            "accent_hover": "#2563eb",
            "success": "#10b981",
            "warning": "#f59e0b",
            "error": "#ef4444",
            "text_primary": "#f3f4f6",
            "text_secondary": "#d1d5db",
            "text_tertiary": "#9ca3af",
        }

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

        # 2. Sidebar with improved styling
        self.sidebar = ctk.CTkFrame(self, width=100, corner_radius=0, fg_color=self.colors["bg_secondary"])
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        # Logo with better styling
        logo_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_frame.pack(pady=30)
        ctk.CTkLabel(logo_frame, text="🎯", font=("Arial", 32), text_color=self.colors["accent"]).pack()
        ctk.CTkLabel(logo_frame, text="resQ", font=("Segoe UI", 14, "bold"), text_color=self.colors["accent"]).pack(pady=(5, 0))
        ctk.CTkLabel(logo_frame, text="by vishal", font=("Arial", 10), text_color=self.colors["accent"]).pack()

        # 3. Container & Tabs
        self.container = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        self.container.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        self.tabs = ctk.CTkTabview(self.container, fg_color="transparent")
        self.tabs.pack(fill="both", expand=True, padx=20, pady=20)
        
        self.t_dash = self.tabs.add("Dashboard")
        self.t_ops = self.tabs.add("Operations")
        self.t_assets = self.tabs.add("Asset Manager")
        self.t_billing = self.tabs.add("Billing / TCR")
        self.tabs._segmented_button.grid_forget()

        # 4. Navigation with current tab indicator
        self.nav_btns = {}
        self.create_nav_btn("📊", "Dashboard")
        self.create_nav_btn("🔧", "Operations")
        self.create_nav_btn("📁", "Asset Manager")
        self.create_nav_btn("📋", "Billing / TCR")

        # Backup/Export button always visible in sidebar
        ctk.CTkButton(
            self.sidebar,
            text="💾",
            width=65,
            height=65,
            fg_color="transparent",
            font=("Arial", 26),
            hover_color=self.colors["bg_tertiary"],
            command=self.backup_data_ui,
        ).pack(pady=12)
        ctk.CTkLabel(self.sidebar, text="v5.8", font=("Segoe UI", 10), text_color=self.colors["text_tertiary"]).pack(side="bottom", pady=(0, 20))

        # 5. Initialize Content
        self.init_dashboard_tab()
        self.init_operations_tab()
        self.init_assets_tab()
        self.init_billing_tab()
        
        self.tabs.set("Dashboard")
        self._apply_ttk_theme()
        self.refresh_all_data()

    def _apply_ttk_theme(self):
        """Professional dark theme for ttk.Treeview widgets with better visuals."""
        try:
            style = ttk.Style()
            try:
                style.theme_use("clam")
            except Exception:
                pass

            style.configure(
                "Treeview",
                background=self.colors["bg_secondary"],
                fieldbackground=self.colors["bg_secondary"],
                foreground=self.colors["text_primary"],
                rowheight=32,
                bordercolor=self.colors["bg_tertiary"],
                borderwidth=1,
                font=("Segoe UI", 10),
            )
            style.configure(
                "Treeview.Heading",
                background=self.colors["bg_tertiary"],
                foreground=self.colors["text_secondary"],
                relief="flat",
                font=("Segoe UI", 11, "bold"),
                padding=8,
            )
            style.map(
                "Treeview", 
                background=[("selected", self.colors["accent"]), ("alternate", self.colors["bg_tertiary"])],
                foreground=[("selected", "#ffffff")],
                borderwidth=[("selected", 1)]
            )
            style.configure("Treeview", rowheight=28)
        except Exception:
            pass

    def _parse_date(self, value):
        if value is None:
            return None
        if isinstance(value, (datetime,)):
            return value.date()
        s = str(value).strip()
        if s == "" or s.lower() in ("nan", "none"):
            return None
        for fmt in ("%d-%b-%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%b-%Y %H:%M:%S", "%d/%m/%Y"):
            try:
                return datetime.strptime(s, fmt).date()
            except Exception:
                continue
        return None

    def _format_date(self, value):
        dt = self._parse_date(value)
        if dt is None:
            return str(value).strip() if value is not None else ""
        return dt.strftime("%d-%b-%Y")

    def _show_date_picker(self, entry_widget, initial_value="", parent=None):
        """Simple reliable date picker with working calendar grid."""
        date_value = self._parse_date(initial_value) or datetime.now().date()
        selected_date = [date_value]
        
        if parent is None:
            parent = self
        
        dlg = tk.Toplevel(parent)
        dlg.title("Select Date")
        dlg.geometry("420x380")
        dlg.resizable(False, False)
        dlg.configure(bg="#16161a")
        dlg.transient(parent)
        dlg.grab_set()
        dlg.attributes('-topmost', True)
        
        current_month = [date_value.year, date_value.month]
        months_list = ["January", "February", "March", "April", "May", "June", 
                      "July", "August", "September", "October", "November", "December"]
        
        # Top frame with month/year and navigation
        top_frame = tk.Frame(dlg, bg="#3b82f6", height=60)
        top_frame.pack(fill="x", padx=0, pady=0)
        top_frame.pack_propagate(False)
        
        month_label = tk.Label(top_frame, text="", font=("Arial", 14, "bold"), 
                              fg="white", bg="#3b82f6")
        month_label.pack(pady=8)
        
        nav_frame = tk.Frame(top_frame, bg="#3b82f6")
        nav_frame.pack(fill="x", padx=10, pady=(0, 8))
        
        def update_month_label():
            month_label.config(text=f"{months_list[current_month[1]-1]} {current_month[0]}")
        
        def prev_month():
            if current_month[1] == 1:
                current_month[0] -= 1
                current_month[1] = 12
            else:
                current_month[1] -= 1
            update_month_label()
            draw_calendar()
        
        def next_month():
            if current_month[1] == 12:
                current_month[0] += 1
                current_month[1] = 1
            else:
                current_month[1] += 1
            update_month_label()
            draw_calendar()
        
        tk.Button(nav_frame, text="< Previous", bg="#1f1f26", fg="white", 
                 relief="flat", command=prev_month, width=12).pack(side="left", padx=3)
        tk.Button(nav_frame, text="Today", bg="#1f1f26", fg="white", 
                 relief="flat", command=lambda: set_today(), width=8).pack(side="left", padx=3)
        tk.Button(nav_frame, text="Next >", bg="#1f1f26", fg="white", 
                 relief="flat", command=next_month, width=12).pack(side="right", padx=3)
        
        # Selected date display
        selected_label = tk.Label(dlg, text=f"Selected: {selected_date[0].strftime('%d-%b-%Y')}", 
                                 font=("Arial", 10), fg="#f3f4f6", bg="#16161a")
        selected_label.pack(pady=8)
        
        # Calendar grid frame
        cal_frame = tk.Frame(dlg, bg="#16161a")
        cal_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Day headers
        day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        for col, day_name in enumerate(day_names):
            tk.Label(cal_frame, text=day_name, font=("Arial", 9, "bold"), 
                    fg="#3b82f6", bg="#1f1f26", width=5, height=2, relief="ridge").grid(row=0, column=col)
        
        day_buttons = {}
        
        def draw_calendar():
            # Clear old buttons
            for btn in day_buttons.values():
                btn.destroy()
            day_buttons.clear()
            
            year, month = current_month
            first_day = datetime(year, month, 1)
            last_day = (first_day + timedelta(days=32)).replace(day=1) - timedelta(days=1)
            start_weekday = first_day.weekday()
            
            row = 1
            col = start_weekday
            
            for day in range(1, last_day.day + 1):
                current_day_date = datetime(year, month, day).date()
                is_today = (current_day_date == datetime.now().date())
                is_selected = (current_day_date == selected_date[0])
                
                if is_today:
                    bg = "#3b82f6"
                    fg = "white"
                elif is_selected:
                    bg = "#059669"
                    fg = "white"
                else:
                    bg = "#1f1f26"
                    fg = "#f3f4f6"
                
                def make_select(d, y, m):
                    def select_day():
                        selected_date[0] = datetime(y, m, d).date()
                        selected_label.config(text=f"Selected: {selected_date[0].strftime('%d-%b-%Y')}")
                        confirm()
                    return select_day
                
                btn = tk.Button(cal_frame, text=str(day), font=("Arial", 10, "bold"),
                               bg=bg, fg=fg, width=5, height=2, relief="raised",
                               command=make_select(day, year, month), activebackground="#3b82f6",
                               activeforeground="white")
                btn.grid(row=row, column=col, padx=1, pady=1, sticky="nsew")
                day_buttons[day] = btn
                
                col += 1
                if col > 6:
                    col = 0
                    row += 1
        
        def set_today():
            today = datetime.now().date()
            selected_date[0] = today
            current_month[0] = today.year
            current_month[1] = today.month
            update_month_label()
            selected_label.config(text=f"Selected: {selected_date[0].strftime('%d-%b-%Y')}")
            draw_calendar()
        
        # Bottom buttons frame
        btn_frame = tk.Frame(dlg, bg="#16161a")
        btn_frame.pack(fill="x", padx=10, pady=10)
        
        def confirm():
            entry_widget.delete(0, "end")
            entry_widget.insert(0, selected_date[0].strftime("%d-%b-%Y"))
            dlg.destroy()
        
        def cancel():
            dlg.destroy()
        
        dlg.protocol("WM_DELETE_WINDOW", cancel)
        
        ctk.CTkButton(btn_frame, text="Cancel", fg_color="#4b5563", width=120, command=cancel).pack(side="right", padx=(5, 0))
        ctk.CTkButton(btn_frame, text="Confirm", fg_color="#10b981", width=120, command=confirm).pack(side="right")
        
        update_month_label()
        draw_calendar()
        
        dlg.focus_set()
        dlg.lift()
        dlg.wait_window()

    def create_nav_btn(self, icon, name):
        btn = ctk.CTkButton(
            self.sidebar,
            text=icon,
            width=70,
            height=70,
            fg_color=self.colors["bg_secondary"],
            text_color=self.colors["text_primary"],
            hover_color=self.colors["bg_tertiary"],
            border_width=1,
            border_color=self.colors["bg_tertiary"],
            corner_radius=20,
            font=("Arial", 26),
            command=lambda: self._set_tab(name)
        )
        btn.pack(pady=12)
        self.nav_btns[name] = btn
    
    def _set_tab(self, tab_name):
        """Set tab with visual feedback"""
        self.tabs.set(tab_name)
        self._update_nav_highlight()
    
    def _update_nav_highlight(self):
        """Update navigation button highlighting based on current tab"""
        current = self.tabs.get()
        for name, btn in self.nav_btns.items():
            if name == current:
                btn.configure(fg_color=self.colors["accent"], text_color="#ffffff", border_color=self.colors["accent"], border_width=1)
            else:
                btn.configure(fg_color=self.colors["bg_secondary"], text_color=self.colors["text_primary"], border_color=self.colors["bg_tertiary"], border_width=1)

    # --- 📊 DASHBOARD TAB ---
    def init_dashboard_tab(self):
        # Header section
        header = ctk.CTkFrame(self.t_dash, fg_color="transparent")
        header.pack(fill="x", padx=30, pady=(20, 10))
        ctk.CTkLabel(header, text="Dashboard", font=("Segoe UI", 24, "bold"), text_color=self.colors["text_primary"]).pack(side="left")
        ctk.CTkLabel(header, text="Real-time inventory overview", font=("Segoe UI", 12), text_color=self.colors["text_tertiary"]).pack(side="left", padx=(15, 0))
        
        # Right-side buttons
        btn_group = ctk.CTkFrame(header, fg_color="transparent")
        btn_group.pack(side="right")
        
        ctk.CTkButton(
            btn_group,
            text="📐 Reformat Excel",
            width=160,
            height=36,
            fg_color=self.colors["success"],
            hover_color="#059669",
            command=self.reformat_excel_ui,
        ).pack(side="right", padx=(0, 10))
        
        ctk.CTkButton(
            btn_group,
            text="🔁 Backup Data",
            width=140,
            height=36,
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent_hover"],
            command=self.backup_data_ui,
        ).pack(side="right")
        
        # Stats Cards
        self.stats_frame = ctk.CTkFrame(self.t_dash, fg_color=self.colors["bg_secondary"], corner_radius=20)
        self.stats_frame.pack(fill="x", padx=30, pady=(0, 20), ipady=10, ipadx=10)

        self.val_card = self.create_stat_card(self.stats_frame, "💰", "TOTAL VALUE", "₹ 0.00", self.colors["accent"])
        self.cnt_card = self.create_stat_card(self.stats_frame, "📦", "TOTAL ITEMS", "0", self.colors["success"])
        self.in_card = self.create_stat_card(self.stats_frame, "📥", "ASSETS IN", "0", self.colors["warning"])
        self.out_card = self.create_stat_card(self.stats_frame, "📤", "ASSETS OUT", "0", self.colors["error"])

        # Right-side dashboard options
        self.dash_opts = ctk.CTkFrame(self.stats_frame, fg_color="transparent")
        self.dash_opts.pack(side="right", padx=10, pady=8)
        self.highlight_closed_var = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(
            self.dash_opts,
            text="Highlight Completed",
            variable=self.highlight_closed_var,
            command=self.refresh_all_data,
            onvalue=True,
            offvalue=False,
        ).pack(pady=(28, 0))

        # Search + Filter Section (full width below stats)
        search_filter_container = ctk.CTkFrame(self.t_dash, fg_color="transparent")
        search_filter_container.pack(fill="x", padx=30, pady=(0, 15))
        search_filter_container.grid_columnconfigure(1, weight=1)
        
        # Search bar
        search_frame = ctk.CTkFrame(search_filter_container, fg_color=self.colors["bg_tertiary"], corner_radius=14)
        search_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=(0, 15))
        search_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(search_frame, text="🔍", font=("Arial", 14), text_color=self.colors["accent"]).grid(row=0, column=0, padx=(15, 10), pady=10)
        self.dash_search = ctk.CTkEntry(search_frame, placeholder_text="Search by ID, Item Name, Invoice, or Engineer", height=38)
        self.dash_search.grid(row=0, column=1, sticky="ew", padx=(0, 15), pady=10)
        self.dash_search.bind("<KeyRelease>", lambda _e: self.refresh_all_data())
        
        # Filter section
        filter_frame = ctk.CTkFrame(search_filter_container, fg_color="transparent")
        filter_frame.grid(row=0, column=2, sticky="ew")
        ctk.CTkLabel(filter_frame, text="Filter:", font=("Segoe UI", 12, "bold"), text_color=self.colors["text_primary"]).pack(side="left", padx=(0, 12))
        self.dash_status_var = ctk.StringVar(value="ALL")
        ctk.CTkSegmentedButton(
            filter_frame,
            values=["ALL", "IN", "OUT"],
            variable=self.dash_status_var,
            command=lambda _v=None: self.refresh_all_data(),
            fg_color=self.colors["bg_tertiary"],
            selected_color=self.colors["accent"],
        ).pack(side="left")

        self.tree_frame = ctk.CTkFrame(self.t_dash, fg_color=self.colors["bg_secondary"], corner_radius=12)
        self.tree_frame.pack(fill="both", expand=True, padx=30, pady=(0, 20))

        # Dashboard columns with improved styling
        self.tree = ttk.Treeview(
            self.tree_frame,
            columns=("ID", "Item", "Invoice", "Status", "User", "Charges", "IN Date", "OUT Date"),
            show="headings",
            height=20
        )
        for col in ("ID", "Item", "Invoice", "Status", "User", "Charges", "IN Date", "OUT Date"):
            self.tree.heading(col, text=col)
            w = 160
            if col in ("Item", "Invoice"):
                w = 220
            if col in ("IN Date", "OUT Date"):
                w = 190
            self.tree.column(col, anchor="center", width=w)
        
        scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)

        # Row tags for better visibility
        try:
            self.tree.tag_configure("closed", foreground="#10b981", background="#0f2818")
            self.tree.tag_configure("open", foreground="#f59e0b", background="#2d1f0a")
        except Exception:
            pass

    def create_stat_card(self, parent, icon, title, val, color):
        card = ctk.CTkFrame(parent, fg_color=self.colors["bg_tertiary"], corner_radius=12, width=280, height=110)
        card.pack(side="left", padx=12)
        card.pack_propagate(False)
        
        # Icon and title in header
        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(pady=(15, 5), padx=15, fill="x")
        ctk.CTkLabel(header, text=icon, font=("Arial", 20)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(header, text=title, font=("Segoe UI", 11, "bold"), text_color=self.colors["text_secondary"]).pack(side="left")
        
        # Value display
        lbl = ctk.CTkLabel(card, text=val, font=("Segoe UI", 22, "bold"), text_color=color)
        lbl.pack(pady=(0, 10))
        return lbl

    # --- 📁 ASSET MANAGER ---
    def init_assets_tab(self):
        self.asset_split = ctk.CTkFrame(self.t_assets, fg_color="transparent")
        self.asset_split.pack(fill="both", expand=True)

        self.reg_box = ctk.CTkFrame(self.asset_split, width=380, fg_color=self.colors["bg_secondary"], corner_radius=12)
        self.reg_box.pack(side="left", fill="y", padx=20, pady=20)
        
        header = ctk.CTkFrame(self.reg_box, fg_color="transparent")
        header.pack(pady=(20, 15), padx=20, fill="x")
        ctk.CTkLabel(header, text="📝", font=("Arial", 24)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(header, text="Register New Asset", font=("Segoe UI", 16, "bold"), text_color=self.colors["text_primary"]).pack(side="left")
        
        self.e_id = self.add_input(self.reg_box, "Article Number")
        self.e_name = self.add_input(self.reg_box, "Part Name")
        self.e_cp = self.add_input(self.reg_box, "Cost Price")
        self.e_sp = self.add_input(self.reg_box, "Selling Price")
        
        ctk.CTkButton(
            self.reg_box, 
            text="💾 SAVE & GENERATE QR", 
            height=45, 
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent_hover"],
            command=self.save_article
        ).pack(pady=15, padx=20, fill="x")

        self.qr_img_label = ctk.CTkLabel(self.reg_box, text="QR Preview", font=("Segoe UI", 11, "bold"), text_color=self.colors["text_tertiary"], text_color_disabled=self.colors["bg_tertiary"])
        self.qr_img_label.pack(pady=(15, 10))

        self.btn_open_qr = ctk.CTkButton(self.reg_box, text="🖨️ OPEN TO PRINT", height=40, fg_color=self.colors["success"], state="disabled", command=self.open_qr_file)
        self.btn_open_qr.pack(pady=8, padx=20, fill="x")

        self.btn_direct_print = ctk.CTkButton(self.reg_box, text="🖨️ DIRECT PRINT", height=40, fg_color=self.colors["accent"], state="disabled", command=self.direct_print_qr)
        self.btn_direct_print.pack(pady=(0, 20), padx=20, fill="x")

        # MASTER TABLE WITH 5 COLUMNS
        master_header = ctk.CTkFrame(self.asset_split, fg_color="transparent")
        master_header.pack(side="top", fill="x", padx=20, pady=(20, 10))
        master_header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(master_header, text="📊 Master Asset Catalog", font=("Segoe UI", 16, "bold"), text_color=self.colors["text_primary"]).pack(side="left")
        
        # Search bar for asset catalog
        search_frame = ctk.CTkFrame(self.asset_split, fg_color=self.colors["bg_tertiary"], corner_radius=12)
        search_frame.pack(side="top", fill="x", padx=20, pady=(0, 12))
        search_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(search_frame, text="🔍", font=("Arial", 14), text_color=self.colors["text_tertiary"]).pack(side="left", padx=(15, 10), pady=12)
        self.asset_search = ctk.CTkEntry(search_frame, placeholder_text="Search assets (ID / Name / Part)...", height=36)
        self.asset_search.pack(side="left", fill="x", expand=True, padx=(0, 15), pady=12)
        self.asset_search.bind("<KeyRelease>", lambda _e: self.refresh_all_data())
        
        self.master_box = ctk.CTkFrame(self.asset_split, fg_color=self.colors["bg_secondary"], corner_radius=12)
        self.master_box.pack(side="right", fill="both", expand=True, padx=20, pady=(0, 20))
        self.master_tree = ttk.Treeview(self.master_box, columns=("ID", "Name", "CP", "SP", "Stock"), show="headings", height=20)
        for c in ("ID", "Name", "CP", "SP", "Stock"): 
            self.master_tree.heading(c, text=c)
            self.master_tree.column(c, anchor="center", width=140 if c in ("Name",) else 100)
        
        scrollbar = ttk.Scrollbar(self.master_box, orient="vertical", command=self.master_tree.yview)
        self.master_tree.configure(yscroll=scrollbar.set)
        self.master_tree.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)

    # --- 📋 BILLING / TCR (post-OUTWARD) ---
    def init_billing_tab(self):
        wrap = ctk.CTkFrame(self.t_billing, fg_color="transparent")
        wrap.pack(fill="both", expand=True, padx=24, pady=20)

        top = ctk.CTkFrame(wrap, fg_color="transparent")
        top.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(
            top,
            text="Billing / TCR tracking",
            font=("Segoe UI", 22, "bold"),
            text_color=self.colors["text_primary"],
        ).pack(side="left")
        ctk.CTkButton(
            top,
            text="Refresh",
            width=100,
            height=34,
            fg_color=self.colors["accent"],
            command=self.refresh_billing_tree,
        ).pack(side="right")
        ctk.CTkButton(
            top,
            text="Generate Invoice",
            width=150,
            height=34,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            command=self.generate_engineer_invoice_pdf,
        ).pack(side="right", padx=(0, 10))
        ctk.CTkButton(
            top,
            text="Edit selected",
            width=120,
            height=34,
            fg_color=self.colors["success"],
            command=self.edit_billing_selected,
        ).pack(side="right", padx=(0, 10))

        # ===== TCR SEARCH SECTION =====
        search_frame = ctk.CTkFrame(wrap, fg_color=self.colors["bg_tertiary"], corner_radius=16)
        search_frame.pack(fill="x", padx=20, pady=(0, 14))
        
        # Header for search section
        search_header = ctk.CTkFrame(search_frame, fg_color="transparent")
        search_header.pack(fill="x", padx=20, pady=(12, 10))
        ctk.CTkLabel(search_header, text="🔍 TCR Search Filters", font=("Segoe UI", 13, "bold"),
                    text_color=self.colors["accent"]).pack(side="left")
        
        # First row: Search inputs
        search_inputs = ctk.CTkFrame(search_frame, fg_color="transparent")
        search_inputs.pack(fill="x", padx=20, pady=(0, 10))
        
        # Engineer search
        ctk.CTkLabel(search_inputs, text="Engineer:", font=("Segoe UI", 10, "bold"),
                    text_color=self.colors["text_secondary"]).pack(side="left", padx=(0, 5))
        self.tcr_search_engineer = ctk.CTkEntry(search_inputs, placeholder_text="Search engineer name...",
                                               height=32, width=180, fg_color=self.colors["bg_primary"])
        self.tcr_search_engineer.pack(side="left", padx=(0, 15))
        
        # Article search
        ctk.CTkLabel(search_inputs, text="Article No:", font=("Segoe UI", 10, "bold"),
                    text_color=self.colors["text_secondary"]).pack(side="left", padx=(0, 5))
        self.tcr_search_article = ctk.CTkEntry(search_inputs, placeholder_text="Search article code...",
                                              height=32, width=180, fg_color=self.colors["bg_primary"])
        self.tcr_search_article.pack(side="left", padx=(0, 15))
        
        # Status filter
        ctk.CTkLabel(search_inputs, text="Status:", font=("Segoe UI", 10, "bold"),
                    text_color=self.colors["text_secondary"]).pack(side="left", padx=(0, 5))
        self.tcr_status_var = ctk.StringVar(value="ALL")
        ctk.CTkSegmentedButton(
            search_inputs,
            values=["ALL", "OPEN", "IN_PROGRESS", "COMPLETED", "CLOSED"],
            variable=self.tcr_status_var,
            command=lambda: self.refresh_billing_tree(),
            fg_color=self.colors["bg_tertiary"],
            font=("Segoe UI", 9)
        ).pack(side="left", padx=(0, 10))
        
        # Search button
        ctk.CTkButton(search_inputs, text="🔎 SEARCH", width=120, height=32,
                     fg_color=self.colors["accent"], hover_color=self.colors["accent_hover"],
                     command=self.perform_tcr_search).pack(side="left", padx=(0, 5))
        
        # Clear button
        ctk.CTkButton(search_inputs, text="✕ CLEAR", width=100, height=32,
                     fg_color=self.colors["warning"], hover_color="#d97706",
                     command=self.clear_tcr_search).pack(side="left")

        ctk.CTkLabel(
            wrap,
            text=(
                "After OUTWARD, a row is created in the Excel sheet Service_Billing. When the engineer submits the bill, "
                "fill in TCR No and amounts here. Check “Money received” when payment is complete — the Status becomes CLOSED "
                "and the row turns green in Excel."
            ),
            font=("Segoe UI", 12),
            text_color=self.colors["text_tertiary"],
            wraplength=1100,
            justify="left",
        ).pack(anchor="w", pady=(0, 14))

        tree_frame = ctk.CTkFrame(wrap, fg_color=self.colors["bg_secondary"], corner_radius=16)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        cols = (
            "Sr",
            "Engineer",
            "Article",
            "Part Sr",
            "Outward",
            "Invoice No",
            "TCR No",
            "Money",
            "Status",
        )
        self.billing_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=22)
        widths = {
            "Sr": 50,
            "Engineer": 160,
            "Article": 90,
            "Part Sr": 70,
            "Outward": 150,
            "Invoice No": 130,
            "TCR No": 120,
            "Money": 60,
            "Status": 90,
        }
        for c in cols:
            self.billing_tree.heading(c, text=c)
            self.billing_tree.column(c, anchor="center", width=widths.get(c, 120))

        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.billing_tree.yview)
        self.billing_tree.configure(yscroll=sb.set)
        self.billing_tree.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        sb.pack(side="right", fill="y", pady=10)

        try:
            self.billing_tree.tag_configure("done", foreground="#22c55e")
            self.billing_tree.tag_configure("pending", foreground="#fbbf24")
        except Exception:
            pass

        self.billing_tree.bind("<Double-1>", lambda _e: self.edit_billing_selected())

    def refresh_billing_tree(self):
        if not getattr(self, "billing_tree", None):
            return
        for i in self.billing_tree.get_children():
            self.billing_tree.delete(i)
        try:
            df = engine.get_service_jobs()
        except Exception:
            return
        if df is None or df.empty:
            return

        # Get search criteria
        engineer_search = getattr(self, "tcr_search_engineer", None)
        article_search = getattr(self, "tcr_search_article", None)
        status_filter = getattr(self, "tcr_status_var", None)
        
        engineer_val = engineer_search.get().strip() if engineer_search else ""
        article_val = article_search.get().strip() if article_search else ""
        status_val = status_filter.get() if status_filter else "ALL"
        
        # Apply filters
        if engineer_val or article_val or (status_val and status_val.upper() != "ALL"):
            df = engine.filter_tcr_records(df, search_type=None, search_value=None, status_filter=status_val)
            
            if engineer_val:
                df["Engineer_Name"] = df["Engineer_Name"].astype(str).str.strip()
                df = df[df["Engineer_Name"].str.lower().str.contains(engineer_val.lower(), na=False)]
            
            if article_val:
                df["Article_No"] = df["Article_No"].astype(str).str.strip()
                df = df[df["Article_No"].str.contains(article_val, na=False)]
        
        def _money_show(val):
            if val is True or str(val).strip().lower() in ("true", "1", "yes", "y", "received", "paid"):
                return "Yes"
            return "No"

        for _, r in df.iloc[::-1].iterrows():
            st = str(r.get("Status", "") or "").strip().upper()
            tag = "done" if st == "CLOSED" else "pending"
            self.billing_tree.insert(
                "",
                "end",
                iid=str(r.get("Row_ID", "")),
                values=(
                    r.get("Sr_No", ""),
                    r.get("Engineer_Name", ""),
                    r.get("Article_No", ""),
                    r.get("Part_Sr_No", ""),
                    r.get("Outward_Date", ""),
                    r.get("Job_Card_Invoice_No", ""),
                    r.get("TCR_No", ""),
                    _money_show(r.get("Money_Received")),
                    r.get("Status", ""),
                ),
                tags=(tag,),
            )

    def edit_billing_selected(self):
        if not getattr(self, "billing_tree", None):
            return
        sel = self.billing_tree.selection()
        if not sel:
            return messagebox.showwarning("Billing", "Select a row first.")
        row_id = str(sel[0]).strip()
        if not row_id:
            return messagebox.showerror("Billing", "Invalid row selection.")

        df = engine.get_service_jobs()
        if df is None or df.empty:
            return messagebox.showerror("Billing", "No data.")
        m = df[df["Row_ID"].astype(str).str.strip() == row_id]
        if m.empty:
            return messagebox.showerror("Billing", "Row not found.")
        row = m.iloc[0].to_dict()

        win = ctk.CTkToplevel(self)
        win.title("TCR Tracking - Service Job Details")
        win.geometry("900x700")
        win.configure(fg_color=self.colors["bg_primary"])
        win.grab_set()
        win.attributes('-topmost', True)

        # Header with status indicator
        hdr = ctk.CTkFrame(win, fg_color=self.colors["bg_secondary"], corner_radius=10)
        hdr.pack(fill="x", padx=16, pady=(16, 8))

        status = str(row.get("Status", "OPEN")).strip().upper()
        status_color = self.colors["success"] if status == "CLOSED" else self.colors["warning"]

        ctk.CTkLabel(hdr, text="🔧 Service Job Details", font=("Segoe UI", 16, "bold"),
                    text_color=self.colors["text_primary"]).pack(side="left", pady=12)
        ctk.CTkLabel(hdr, text=f"Status: {status}", font=("Segoe UI", 12, "bold"),
                    text_color=status_color, fg_color=self.colors["bg_tertiary"],
                    corner_radius=8, padx=10, pady=4).pack(side="right", pady=12)

        # Progress indicator
        progress_frame = ctk.CTkFrame(win, fg_color=self.colors["bg_secondary"], corner_radius=8)
        progress_frame.pack(fill="x", padx=16, pady=(0, 16))

        steps = ["Issued", "In Progress", "Completed", "Paid"]
        current_step = 0
        if status == "CLOSED":
            money_received = row.get("Money_Received")
            if money_received is True or str(money_received).strip().lower() in ("true", "1", "yes", "y", "received", "paid"):
                current_step = 3  # Paid
            else:
                current_step = 2  # Completed but not paid
        elif str(row.get("TCR_No", "")).strip():
            current_step = 2  # Has TCR
        elif str(row.get("Bill_Amount", "")).strip():
            current_step = 1  # Has bill amount

        for i, step in enumerate(steps):
            color = self.colors["success"] if i <= current_step else self.colors["bg_tertiary"]
            text_color = "white" if i <= current_step else self.colors["text_secondary"]
            ctk.CTkLabel(progress_frame, text=step, font=("Segoe UI", 10, "bold"),
                        fg_color=color, text_color=text_color, corner_radius=6,
                        padx=8, pady=4).pack(side="left", padx=4, pady=8)

        # Main content area
        content = ctk.CTkScrollableFrame(win, fg_color=self.colors["bg_secondary"], corner_radius=10)
        content.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        entries = {}

        # Basic Information Section
        basic_frame = ctk.CTkFrame(content, fg_color=self.colors["bg_tertiary"], corner_radius=8)
        basic_frame.pack(fill="x", padx=10, pady=(10, 8))
        ctk.CTkLabel(basic_frame, text="📋 Basic Information", font=("Segoe UI", 14, "bold"),
                    text_color=self.colors["accent"]).pack(anchor="w", padx=12, pady=(12, 8))

        basic_info = [
            ("Job_Card_Invoice_No", "Invoice Number", True),
            ("Engineer_Name", "Engineer", False),
            ("Article_No", "Article Number", False),
            ("Part_Name_Part_No", "Part Name", False),
            ("Outward_Date", "Issue Date", False),
        ]

        for field, label, editable in basic_info:
            if not editable:
                val = str(row.get(field, "") or "").strip()
                if field in {"Outward_Date"} and val:
                    val = self._format_date(val)
                ctk.CTkLabel(basic_frame, text=f"{label}: {val}", font=("Segoe UI", 11),
                            text_color=self.colors["text_primary"]).pack(anchor="w", padx=12, pady=2)
            else:
                ctk.CTkLabel(basic_frame, text=label, font=("Segoe UI", 11, "bold"),
                            text_color=self.colors["text_secondary"]).pack(anchor="w", padx=12, pady=(8, 2))
                if field in {"Outward_Date"}:
                    frame = ctk.CTkFrame(basic_frame, fg_color="transparent")
                    frame.pack(fill="x", padx=12, pady=(0, 8))
                    e = ctk.CTkEntry(frame, height=34, fg_color=self.colors["bg_primary"], border_color=self.colors["accent"])
                    e.insert(0, self._format_date(row.get(field, "")))
                    e.pack(side="left", fill="x", expand=True)
                    ctk.CTkButton(frame, text="📅", width=40, fg_color=self.colors["accent"],
                                 command=lambda e=e, v=row.get(field, ""), w=win: self._show_date_picker(e, v, w)).pack(side="right", padx=(4, 0))
                else:
                    e = ctk.CTkEntry(basic_frame, height=34, fg_color=self.colors["bg_primary"], border_color=self.colors["accent"])
                    e.insert(0, str(row.get(field, "") or "").strip())
                    e.pack(fill="x", padx=12, pady=(0, 8))
                entries[field] = e

        # Service Details Section
        service_frame = ctk.CTkFrame(content, fg_color=self.colors["bg_tertiary"], corner_radius=8)
        service_frame.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkLabel(service_frame, text="🔧 Service Details", font=("Segoe UI", 14, "bold"),
                    text_color=self.colors["accent"]).pack(anchor="w", padx=12, pady=(12, 8))

        service_fields = [
            ("Material_Description", "Work Description", "Describe the service work performed"),
            ("Bill_Amount", "Service Charges (₹)", "Total amount charged for service"),
            ("Incentive", "Incentive Amount (₹)", "Additional incentive payment if applicable"),
        ]

        for field, label, help_text in service_fields:
            ctk.CTkLabel(service_frame, text=label, font=("Segoe UI", 11, "bold"),
                        text_color=self.colors["text_secondary"]).pack(anchor="w", padx=12, pady=(8, 2))
            if help_text:
                ctk.CTkLabel(service_frame, text=help_text, font=("Segoe UI", 9),
                            text_color=self.colors["text_tertiary"]).pack(anchor="w", padx=12, pady=(0, 2))

            e = ctk.CTkEntry(service_frame, height=34, fg_color=self.colors["bg_primary"], border_color=self.colors["accent"])
            val = str(row.get(field, "") or "").strip()
            e.insert(0, val)
            e.pack(fill="x", padx=12, pady=(0, 8))
            entries[field] = e

        # TCR Completion Section
        tcr_frame = ctk.CTkFrame(content, fg_color=self.colors["bg_tertiary"], corner_radius=8)
        tcr_frame.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkLabel(tcr_frame, text="✅ TCR Completion", font=("Segoe UI", 14, "bold"),
                    text_color=self.colors["accent"]).pack(anchor="w", padx=12, pady=(12, 8))

        tcr_fields = [
            ("TCR_No", "TCR Number", "Transaction Completion Report number"),
            ("TCR_Date", "TCR Date", "Date when TCR was submitted"),
            ("Warranty_Warranty_No", "Warranty Number", "Warranty reference if applicable"),
        ]

        for field, label, help_text in tcr_fields:
            ctk.CTkLabel(tcr_frame, text=label, font=("Segoe UI", 11, "bold"),
                        text_color=self.colors["text_secondary"]).pack(anchor="w", padx=12, pady=(8, 2))
            if help_text:
                ctk.CTkLabel(tcr_frame, text=help_text, font=("Segoe UI", 9),
                            text_color=self.colors["text_tertiary"]).pack(anchor="w", padx=12, pady=(0, 2))

            if field == "TCR_Date":
                frame = ctk.CTkFrame(tcr_frame, fg_color="transparent")
                frame.pack(fill="x", padx=12, pady=(0, 8))
                e = ctk.CTkEntry(frame, height=34, fg_color=self.colors["bg_primary"], border_color=self.colors["accent"])
                e.insert(0, self._format_date(row.get(field, "")))
                e.pack(side="left", fill="x", expand=True)
                ctk.CTkButton(frame, text="📅", width=40, fg_color=self.colors["accent"],
                             command=lambda e=e, v=row.get(field, ""), w=win: self._show_date_picker(e, v, w)).pack(side="right", padx=(4, 0))
            else:
                e = ctk.CTkEntry(tcr_frame, height=34, fg_color=self.colors["bg_primary"], border_color=self.colors["accent"])
                e.insert(0, str(row.get(field, "") or "").strip())
                e.pack(fill="x", padx=12, pady=(0, 8))
            entries[field] = e

        # Payment Section
        payment_frame = ctk.CTkFrame(content, fg_color=self.colors["bg_tertiary"], corner_radius=8)
        payment_frame.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkLabel(payment_frame, text="💰 Payment Status", font=("Segoe UI", 14, "bold"),
                    text_color=self.colors["accent"]).pack(anchor="w", padx=12, pady=(12, 8))

        # Money received checkbox
        money_var = ctk.BooleanVar(value=self._truthy_money(row.get("Money_Received")))
        cb = ctk.CTkCheckBox(
            payment_frame,
            text="✅ Payment Received - Mark as completed",
            variable=money_var,
            font=("Segoe UI", 12, "bold"),
            text_color=self.colors["text_primary"],
        )
        cb.pack(anchor="w", padx=12, pady=(8, 12))
        entries["Money_Received"] = money_var

        # Quick Actions
        actions_frame = ctk.CTkFrame(content, fg_color=self.colors["bg_tertiary"], corner_radius=8)
        actions_frame.pack(fill="x", padx=10, pady=(0, 16))
        ctk.CTkLabel(actions_frame, text="⚡ Quick Actions", font=("Segoe UI", 14, "bold"),
                    text_color=self.colors["accent"]).pack(anchor="w", padx=12, pady=(12, 8))

        def mark_completed():
            entries["TCR_No"].insert(0, f"TCR-{row.get('Sr_No', '')}")
            entries["TCR_Date"].delete(0, "end")
            entries["TCR_Date"].insert(0, datetime.now().strftime("%d-%b-%Y"))
            messagebox.showinfo("Quick Action", "TCR details auto-filled. Review and save.")

        def mark_paid():
            money_var.set(True)
            messagebox.showinfo("Quick Action", "Marked as paid. Don't forget to save!")

        btn_frame = ctk.CTkFrame(actions_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(btn_frame, text="Auto-fill TCR", fg_color=self.colors["warning"],
                     command=mark_completed).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_frame, text="Mark as Paid", fg_color=self.colors["success"],
                     command=mark_paid).pack(side="left")

        # Save/Cancel buttons
        btn_row = ctk.CTkFrame(win, fg_color="transparent")
        btn_row.pack(fill="x", padx=16, pady=(0, 16))
        ctk.CTkButton(btn_row, text="Cancel", fg_color="#4b5563", width=120, command=win.destroy).pack(side="right")
        ctk.CTkButton(btn_row, text="💾 Save Changes", fg_color=self.colors["success"], width=160,
                     command=lambda: self._save_billing_changes(win, row_id, entries)).pack(side="right", padx=(0, 10))

    def _truthy_money(self, val):
        if val is True:
            return True
        s = str(val).strip().lower()
        return s in ("true", "1", "yes", "y", "received", "paid")

    def _save_billing_changes(self, win, row_id, entries):
        updates = {}
        for k, w in entries.items():
            if k == "Money_Received":
                updates[k] = bool(w.get())
            else:
                updates[k] = (w.get() or "").strip()
        ok, msg = engine.update_service_job(row_id, updates)
        if ok:
            try:
                win.destroy()
            except Exception:
                pass
            self.refresh_billing_tree()
            self.refresh_all_data()
            messagebox.showinfo("Saved", msg)
        else:
            messagebox.showerror("Error", msg)

    # ===== TCR SEARCH METHODS =====
    def perform_tcr_search(self):
        """Perform TCR search based on selected filters"""
        engineer_val = (self.tcr_search_engineer.get() or "").strip()
        article_val = (self.tcr_search_article.get() or "").strip()
        
        # Trigger refresh with search criteria
        self.refresh_billing_tree()
        
        # Show feedback
        search_criteria = []
        if engineer_val:
            search_criteria.append(f"Engineer: {engineer_val}")
        if article_val:
            search_criteria.append(f"Article: {article_val}")
        
        status_val = (self.tcr_status_var.get() or "ALL").strip()
        if status_val != "ALL":
            search_criteria.append(f"Status: {status_val}")
        
        if search_criteria:
            messagebox.showinfo("TCR Search", f"Searching for:\n• {chr(10).join(['• ' + c for c in search_criteria])}")
        else:
            messagebox.showinfo("TCR Search", "Showing all records.")

    def clear_tcr_search(self):
        """Clear all search filters and reset to show all records"""
        self.tcr_search_engineer.delete(0, "end")
        self.tcr_search_article.delete(0, "end")
        self.tcr_status_var.set("ALL")
        self.refresh_billing_tree()
        messagebox.showinfo("Search Cleared", "All filters cleared. Showing all TCR records.")

    def generate_engineer_invoice_pdf(self):
        selected = self.billing_tree.selection()
        engineer_name = ""

        if selected:
            row_id = str(selected[0]).strip()
            if row_id:
                df = engine.get_service_jobs()
                if df is not None and not df.empty:
                    row = df[df["Row_ID"].astype(str).str.strip() == row_id]
                    if not row.empty:
                        engineer_name = str(row.iloc[0].get("Engineer_Name", "") or "").strip()

        if not engineer_name:
            engineer_name = (self.tcr_search_engineer.get() or "").strip()

        if not engineer_name:
            dialog = CTkInputDialog(text="Enter engineer name for the invoice", title="Generate Invoice PDF")
            engineer_name = str(dialog.get_input() or "").strip()

        if not engineer_name:
            return messagebox.showwarning("Invoice PDF", "Engineer name is required to generate the invoice.")

        default_name = f"Invoice_{engineer_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pdf"
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=default_name,
            title="Save Engineer Invoice PDF",
        )
        if not path:
            return

        ok, result = engine.generate_engineer_invoice_pdf(engineer_name, path)
        if ok:
            messagebox.showinfo("Invoice PDF", f"Invoice saved to:\n{result}")
        else:
            messagebox.showerror("Invoice PDF", result)

    # --- 🔧 OPERATIONS ---
    def init_operations_tab(self):
        self.ops_split = ctk.CTkFrame(self.t_ops, fg_color="transparent")
        self.ops_split.pack(fill="both", expand=True)
        
        self.scan_pnl = ctk.CTkFrame(self.ops_split, fg_color=self.colors["bg_secondary"], corner_radius=12)
        self.scan_pnl.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        self.scan_pnl.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(self.scan_pnl, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=25, pady=(25, 10))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="🔧 Operations", font=("Segoe UI", 22, "bold"), text_color=self.colors["text_primary"]).grid(row=0, column=0, sticky="w")
        self.status_lbl = ctk.CTkLabel(
            header,
            text="🟢 READY TO SCAN",
            font=("Segoe UI", 12, "bold"),
            text_color=self.colors["success"],
        )
        self.status_lbl.grid(row=1, column=0, sticky="w", pady=(6, 0))
        
        # Search bar for operations
        search_frame = ctk.CTkFrame(self.scan_pnl, fg_color=self.colors["bg_tertiary"], corner_radius=12)
        search_frame.grid(row=2, column=0, sticky="ew", padx=25, pady=(0, 15))
        search_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(search_frame, text="🔍", font=("Arial", 14), text_color=self.colors["text_tertiary"]).grid(row=0, column=0, padx=(15, 10), pady=12)
        self.ops_search = ctk.CTkEntry(search_frame, placeholder_text="Search transactions (ID/Item/Engineer)...", height=36)
        self.ops_search.grid(row=0, column=1, sticky="ew", padx=(0, 15), pady=12)
        self.ops_search.bind("<KeyRelease>", lambda _e: self.filter_operations())

        # Transaction History Panel
        history_frame = ctk.CTkFrame(self.scan_pnl, fg_color=self.colors["bg_tertiary"], corner_radius=12)
        history_frame.grid(row=3, column=0, sticky="ew", padx=25, pady=(0, 15))
        history_frame.grid_columnconfigure(0, weight=1)
        history_frame.grid_rowconfigure(1, weight=1)
        
        ctk.CTkLabel(history_frame, text="📜 Recent Transactions", font=("Segoe UI", 14, "bold"), text_color=self.colors["text_primary"]).grid(
            row=0, column=0, sticky="w", padx=16, pady=(14, 10)
        )
        
        # Transaction list
        self.ops_transaction_list = ctk.CTkScrollableFrame(history_frame, fg_color=self.colors["bg_secondary"], corner_radius=8)
        self.ops_transaction_list.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 16))
        self.ops_transaction_list.grid_columnconfigure(0, weight=1)
        
        # Details card with better styling
        self.details_card = ctk.CTkFrame(self.scan_pnl, fg_color=self.colors["bg_tertiary"], corner_radius=12)
        self.details_card.grid(row=1, column=0, sticky="ew", padx=25, pady=(10, 15))
        self.details_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.details_card, text="📋 Scanned Item Details", font=("Segoe UI", 14, "bold"), text_color=self.colors["text_primary"]).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(14, 10)
        )

        def _row(r, k, icon=""):
            label_text = f"{icon} {k}" if icon else k
            ctk.CTkLabel(self.details_card, text=label_text, font=("Segoe UI", 11, "bold"), text_color=self.colors["text_secondary"]).grid(
                row=r, column=0, sticky="w", padx=16, pady=8
            )
            v = ctk.CTkLabel(self.details_card, text="-", font=("Segoe UI", 12), text_color=self.colors["text_primary"], anchor="w", justify="left", wraplength=600)
            v.grid(row=r, column=1, sticky="ew", padx=(0, 16), pady=8)
            return v

        self.v_art = _row(1, "Article Code", "🏷️")
        self.v_sr = _row(2, "Sr No", "#️⃣")
        self.v_part = _row(3, "Part Name", "📦")
        self.v_inv = _row(4, "Invoice No", "📄")
        self.v_sp = _row(5, "Selling Price", "💵")

        actions = ctk.CTkFrame(self.scan_pnl, fg_color=self.colors["bg_tertiary"], corner_radius=16)
        actions.grid(row=2, column=0, sticky="ew", padx=25, pady=(0, 20))
        actions.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(actions, text="Scan & move parts quickly", font=("Segoe UI", 12, "bold"), text_color=self.colors["text_primary"]).grid(row=0, column=0, sticky="w", padx=18, pady=(16, 8))
        ctk.CTkLabel(actions, text="Use the camera to scan article QR codes and instantly issue or return items.", font=("Segoe UI", 10), text_color=self.colors["text_tertiary"], wraplength=520, justify="left").grid(row=1, column=0, sticky="w", padx=18)

        ctk.CTkButton(
            actions, 
            text="📸 ACTIVATE CAMERA", 
            height=52, 
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent_hover"],
            corner_radius=14,
            command=self.run_scan
        ).grid(row=2, column=0, sticky="ew", padx=18, pady=(16, 0))

        move_label = ctk.CTkLabel(actions, text="Purchase Type:", font=("Segoe UI", 11, "bold"), text_color=self.colors["text_secondary"])
        move_label.grid(row=3, column=0, sticky="w", padx=18, pady=(16, 6))
        
        self.move_type_var = ctk.StringVar(value="Company")
        ctk.CTkSegmentedButton(
            actions, 
            values=["🏢 Company", "🏪 Local"], 
            variable=self.move_type_var,
            fg_color=self.colors["bg_secondary"],
            selected_color=self.colors["accent"],
            text_color=self.colors["text_primary"],
            corner_radius=14,
        ).grid(row=4, column=0, sticky="ew", padx=18)

        btn_row = ctk.CTkFrame(actions, fg_color="transparent")
        btn_row.grid(row=5, column=0, sticky="ew", pady=(18, 0), padx=18)
        btn_row.grid_columnconfigure(0, weight=1)
        btn_row.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(
            btn_row, 
            text="📥 INWARD", 
            height=44, 
            fg_color=self.colors["success"],
            hover_color="#059669",
            command=lambda: self.execute_move("IN")
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(
            btn_row, 
            text="📤 OUTWARD", 
            height=44, 
            fg_color=self.colors["warning"],
            hover_color="#d97706",
            command=lambda: self.execute_move("OUT")
        ).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        quick = ctk.CTkFrame(actions, fg_color="transparent")
        quick.grid(row=6, column=0, sticky="ew", pady=(10, 18), padx=18)
        quick.grid_columnconfigure(0, weight=1)
        quick.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(quick, text="🧹 CLEAR", height=42, fg_color=self.colors["bg_secondary"], hover_color="#2d2d35", corner_radius=14, command=self.clear_scan).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(quick, text="🖨️ REPRINT QR", height=42, fg_color=self.colors["bg_secondary"], hover_color="#2d2d35", corner_radius=14, command=self.direct_print_qr).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        
        self.eng_wrap = ctk.CTkFrame(self.ops_split, width=350, fg_color=self.colors["bg_secondary"], corner_radius=20)
        self.eng_wrap.pack(side="right", fill="both", padx=20, pady=20)

        eng_hdr = ctk.CTkFrame(self.eng_wrap, fg_color="transparent")
        eng_hdr.pack(fill="x", pady=(20, 8), padx=20)
        ctk.CTkLabel(eng_hdr, text="👷 Engineers", font=("Segoe UI", 16, "bold"), text_color=self.colors["text_primary"]).pack(side="left")
        ctk.CTkButton(eng_hdr, text="✏️", width=40, height=32, fg_color=self.colors["warning"], hover_color="#d97706", command=self.edit_engineer_ui).pack(side="right", padx=(6, 0))
        ctk.CTkButton(eng_hdr, text="➕", width=40, height=32, fg_color=self.colors["success"], hover_color="#059669", command=self.add_engineer_ui).pack(side="right", padx=(6, 0))
        ctk.CTkButton(eng_hdr, text="➖", width=40, height=32, fg_color=self.colors["error"], hover_color="#dc2626", command=self.remove_engineer_ui).pack(side="right")

        self.eng_search = ctk.CTkEntry(self.eng_wrap, placeholder_text="🔍 Search engineer", height=36)
        self.eng_search.pack(fill="x", padx=20, pady=(0, 12))
        self.eng_search.bind("<KeyRelease>", lambda _e: self.refresh_engineers())

        self.eng_pnl = ctk.CTkScrollableFrame(self.eng_wrap, label_text="📋 Select Engineer", fg_color=self.colors["bg_secondary"], corner_radius=16)
        self.eng_pnl.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.eng_btns = {}
        self.refresh_engineers()

    def clear_scan(self):
        self.current_scanned_id = None
        self.current_scanned_sr = None
        self.status_lbl.configure(text="🟢 READY TO SCAN", text_color=self.colors["success"])
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

        if not engineers:
            ctk.CTkLabel(
                self.eng_pnl, 
                text="No engineers added yet", 
                font=("Segoe UI", 11), 
                text_color=self.colors["text_tertiary"]
            ).pack(pady=20)
            return

        for row in engineers:
            name = row.get("Engineer_Name", "")
            bp = row.get("BP_ID", "") or "-"
            pprr = row.get("PPRR_ID", "") or "-"
            label = f"{name}\n📌 BP: {bp}  |  {pprr}"

            if q:
                hay = f"{name} {bp} {pprr}".lower()
                if q not in hay:
                    continue

            b = ctk.CTkButton(
                self.eng_pnl,
                text=label,
                height=56,
                fg_color=self.colors["bg_secondary"],
                hover_color=self.colors["accent"],
                font=("Segoe UI", 11, "bold"),
                text_color=self.colors["text_primary"],
                anchor="w",
                command=lambda e=name: self.set_eng(e),
            )
            b.pack(pady=6, padx=10, fill="x")
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

    def edit_engineer_ui(self):
        if not self.selected_engineer:
            return messagebox.showwarning("Edit Engineer", "Select an engineer first.")
        try:
            bp = (CTkInputDialog(text="BP ID (optional) or leave empty to skip", title="Edit Engineer").get_input() or "").strip()
            pprr = (CTkInputDialog(text="PPRR ID (optional) or leave empty to skip", title="Edit Engineer").get_input() or "").strip()
            if bp == "" and pprr == "":
                return messagebox.showwarning("Edit Engineer", "Enter at least one value.")
            ok, msg = engine.add_engineer(self.selected_engineer, bp_id=bp, pprr_id=pprr)
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

    def backup_data_ui(self):
        try:
            ok, msg = engine.backup_database(include_qr_folder=True)
            if ok:
                messagebox.showinfo("Backup Created", f"Backup saved to:\n{msg}")
            else:
                messagebox.showerror("Backup Failed", msg)
        except Exception as e:
            messagebox.showerror("Backup Failed", str(e))

    def reformat_excel_ui(self):
        """Reformat Excel file with auto-sized columns"""
        try:
            ok, msg = engine.reformat_excel_file()
            if ok:
                messagebox.showinfo("Excel Reformatted", "✅ All columns have been auto-sized for better readability!\n\nYou can now open the Excel file to see the changes.")
            else:
                messagebox.showerror("Reformat Failed", msg)
        except Exception as e:
            messagebox.showerror("Reformat Failed", str(e))

    # --- CORE REFRESH LOGIC (THE FIX) ---
    def refresh_all_data(self):
        # 1. Update Dashboard
        for i in self.tree.get_children(): self.tree.delete(i)
        if os.path.exists(engine.DB_FILE):
            try:
                # Optimized: Read once with openpyxl engine for speed
                df_t = pd.read_excel(engine.DB_FILE, sheet_name='Transactions', engine='openpyxl')
                df_m = pd.read_excel(engine.DB_FILE, sheet_name='Master', engine='openpyxl')
                if "Sr_No" not in df_t.columns:
                    df_t["Sr_No"] = 1
                if "Tax_Invoice_No" not in df_t.columns:
                    df_t["Tax_Invoice_No"] = ""
                if "In_Date" not in df_t.columns:
                    df_t["In_Date"] = df_t["Date"] if "Date" in df_t.columns else ""
                if "Out_Date" not in df_t.columns:
                    df_t["Out_Date"] = df_t["Date"] if "Date" in df_t.columns else ""
                # Vectorized merge - faster than loop
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
                    in_d = self._format_date(in_d)
                    out_d = self._format_date(out_d)
                    in_d = "-" if in_d == "" else in_d
                    out_d = "-" if out_d == "" else out_d
                    status = str(r.get("Status", "") or "").strip().upper()
                    if status_filter in ("IN", "OUT") and status != status_filter:
                        continue

                    if q:
                        hay = f"{display_id} {r.get('Part_Name','')} {inv} {r.get('Engineer','')}".lower()
                        if q not in hay:
                            continue

                    # Add visual indicators for status
                    status_display = f"✓ {status}" if status == "IN" else f"✗ {status}"
                    
                    tags = ()
                    if getattr(self, "highlight_closed_var", None) is not None:
                        if bool(self.highlight_closed_var.get()) and status == "OUT":
                            tags = ("closed",)
                        elif status == "IN":
                            tags = ("open",)

                    self.tree.insert(
                        "",
                        "end",
                        values=(
                            display_id,
                            r.get("Part_Name", ""),
                            inv,
                            status_display,
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
                
                # Asset search filter
                asset_q = ""
                try:
                    asset_q = (self.asset_search.get() or "").strip().lower()
                except Exception:
                    asset_q = ""
                
                for _, r in df_m.iterrows():
                    # Apply search filter
                    if asset_q:
                        hay = f"{r['Article_No']} {r['Part_Name']}".lower()
                        if asset_q not in hay:
                            continue
                    
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
                
                # Count IN and OUT assets from transactions
                try:
                    in_count = len(df_t[df_t['Status'].astype(str).str.upper() == 'IN'])
                    out_count = len(df_t[df_t['Status'].astype(str).str.upper() == 'OUT'])
                    self.in_card.configure(text=str(in_count))
                    self.out_card.configure(text=str(out_count))
                except Exception as e:
                    pass
            except Exception as e: print(f"Master Error: {e}")

        try:
            self.refresh_billing_tree()
        except Exception:
            pass

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
            text="✅ SCAN SUCCESSFUL",
            text_color=self.colors["success"],
        )
        self.v_art.configure(text=str(art))
        self.v_sr.configure(text=str(sr_txt))
        self.v_part.configure(text=str(part) if part else "-")
        self.v_inv.configure(text=str(inv_txt) if inv_txt else "-")
        self.v_sp.configure(text=str(sp_txt))

    def set_eng(self, n):
        for b in self.eng_btns.values(): 
            b.configure(fg_color=self.colors["bg_secondary"])
        self.eng_btns[n].configure(fg_color=self.colors["accent"])
        self.selected_engineer = n

    def filter_operations(self):
        """Filter operations/transactions based on search input"""
        # This method filters the transaction history list
        # It's called whenever the user types in the ops_search field
        try:
            search_text = (self.ops_search.get() or "").strip().lower()
            # Refresh the transaction list with the search filter applied
            # The actual filtering is handled in refresh_all_data() when dash_search is checked
            self.refresh_all_data()
        except Exception as e:
            print(f"Filter operations error: {e}")

    def add_input(self, parent, placeholder):
        e = ctk.CTkEntry(
            parent, 
            placeholder_text=placeholder, 
            height=36, 
            fg_color=self.colors["bg_tertiary"], 
            border_color=self.colors["accent"],
            border_width=2
        )
        e.pack(pady=8, padx=20, fill="x")
        return e

if __name__ == "__main__":
    app = ResQUltimateAdmin()
    app.mainloop()