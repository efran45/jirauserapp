import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry  # pip install tkcalendar
import requests
from requests.auth import HTTPBasicAuth
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from datetime import datetime
import csv
import threading
import keyring
from dateutil import parser  # pip install python-dateutil
import webbrowser
import json
import time
from functools import partial

SERVICE_NAME = "jira_user_app"

# Configure requests session with retry logic
def create_session():
    session = requests.Session()
    retry = Retry(
        total=5,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["HEAD", "GET", "OPTIONS"]
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session

class JiraUserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Jira User & Group Management")
        self.root.geometry("1400x850")

        # ---------------- State ---------------- #
        self.jira_url = tk.StringVar()
        self.email = tk.StringVar()
        self.api_token = tk.StringVar()
        self.org_api_key = tk.StringVar()
        self.org_id = tk.StringVar()
        self.search_var = tk.StringVar()
        self.remember_creds = tk.BooleanVar(value=True)
        self.use_org_api = tk.BooleanVar(value=False)

        self.users_data = []
        self.groups_data = []
        self.groups_members = {}
        self.users_product_access = {}  # Store product access data
        self.products_data = {}  # Store products with their users
        self.current_view = "users"

        self.sort_column = None
        self.sort_reverse = False

        self.setup_ui()
        self.load_credentials()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------------- UI ---------------- #
    def setup_ui(self):
        print("=== Starting UI Setup ===")
        
        # Main container with notebook (tabs)
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Tab 1: Configuration
        config_tab = ttk.Frame(notebook, padding=15)
        notebook.add(config_tab, text="⚙️ Configuration")
        
        # Jira Configuration Section
        jira_frame = ttk.LabelFrame(config_tab, text="Jira Connection", padding=15)
        jira_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(jira_frame, text="Jira URL:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(jira_frame, textvariable=self.jira_url, width=60).grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=5)
        
        ttk.Label(jira_frame, text="Email:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(jira_frame, textvariable=self.email, width=60).grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=5)
        
        ttk.Label(jira_frame, text="Jira API Token:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(jira_frame, textvariable=self.api_token, width=60, show="*").grid(row=2, column=1, sticky="ew", padx=(10, 0), pady=5)
        
        jira_frame.columnconfigure(1, weight=1)
        
        # Organization API Section (collapsible)
        org_frame = ttk.LabelFrame(config_tab, text="Organization API (Optional - For Last Login Data)", padding=15)
        org_frame.pack(fill="x", pady=(0, 10))
        
        # Enable checkbox
        enable_frame = ttk.Frame(org_frame)
        enable_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Checkbutton(
            enable_frame,
            text="Enable Organization API",
            variable=self.use_org_api,
            command=self.toggle_org_api
        ).pack(side="left")
        
        ttk.Button(
            enable_frame,
            text="📖 Help",
            command=self.show_org_api_help,
            width=10
        ).pack(side="right")
        
        # Org API fields
        self.org_fields_frame = ttk.Frame(org_frame)
        self.org_fields_frame.pack(fill="x")
        
        ttk.Label(self.org_fields_frame, text="Org API Key:").grid(row=0, column=0, sticky="w", pady=5)
        self.org_api_entry_widget = ttk.Entry(self.org_fields_frame, textvariable=self.org_api_key, width=50, show="*")
        self.org_api_entry_widget.grid(row=0, column=1, sticky="ew", padx=(10, 10), pady=5)
        ttk.Button(self.org_fields_frame, text="Get Org ID", command=self.fetch_org_id_async, width=15).grid(row=0, column=2, pady=5)
        
        ttk.Label(self.org_fields_frame, text="Organization ID:").grid(row=1, column=0, sticky="w", pady=5)
        self.org_entry = ttk.Entry(self.org_fields_frame, textvariable=self.org_id, width=50, state="readonly")
        self.org_entry.grid(row=1, column=1, sticky="ew", padx=(10, 10), pady=5)
        
        self.org_fields_frame.columnconfigure(1, weight=1)
        
        # Settings
        settings_frame = ttk.Frame(config_tab)
        settings_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Checkbutton(
            settings_frame,
            text="Remember credentials (Jira URL, Email)",
            variable=self.remember_creds
        ).pack(side="left")
        
        # Action buttons
        action_frame = ttk.Frame(config_tab)
        action_frame.pack(fill="x")
        
        ttk.Button(action_frame, text="✓ Validate Token", command=self.validate_token_async, width=20).pack(side="left", padx=(0, 5))
        
        # Tab 2: Data (Users & Groups with sub-tabs)
        data_tab = ttk.Frame(notebook, padding=10)
        notebook.add(data_tab, text="👥 Data")
        
        # Tab 3: Products
        products_tab = ttk.Frame(notebook, padding=10)
        notebook.add(products_tab, text="📦 Products")
        
        # Setup Data tab (contains Users/Groups sub-tabs)
        self.setup_users_tab(data_tab)
        
        # Setup Products tab
        self.setup_products_tab(products_tab)
        
        # Store notebook reference for tab switching
        self.notebook = notebook
        
        print("=== UI Setup Complete ===")
    
    def setup_users_tab(self, parent):
        """Setup the Users tab with compact filters"""
        
        # Top action bar with fetch buttons back
        action_bar = ttk.Frame(parent)
        action_bar.pack(fill="x", pady=(0, 10))
        
        ttk.Button(action_bar, text="📥 Fetch Users", command=self.fetch_users_async, width=15).pack(side="left", padx=(0, 5))
        ttk.Button(action_bar, text="👥 Fetch Groups", command=self.fetch_groups_async, width=15).pack(side="left", padx=(0, 5))
        
        # Add Bulk Edit button (initially disabled)
        self.bulk_edit_btn = ttk.Button(action_bar, text="⚡ Bulk Edit", command=self.show_bulk_edit_dialog, width=15, state="disabled")
        self.bulk_edit_btn.pack(side="left", padx=(0, 5))
        
        ttk.Button(action_bar, text="💾 Export CSV", command=self.export_csv, width=15).pack(side="left", padx=(0, 5))
        
        # Separator
        ttk.Separator(action_bar, orient="vertical").pack(side="left", fill="y", padx=10)
        
        # Organization API checkbox
        ttk.Checkbutton(
            action_bar,
            text="Enable Org API",
            variable=self.use_org_api,
            command=self.toggle_org_api
        ).pack(side="left", padx=(0, 10))
        
        # Separator
        ttk.Separator(action_bar, orient="vertical").pack(side="left", fill="y", padx=10)
        
        # Status indicator
        self.status = ttk.Label(action_bar, text="Ready", foreground="blue", font=("", 10, "bold"))
        self.status.pack(side="left", padx=10)

        # Compact filters - all in ONE line
        filter_frame = ttk.LabelFrame(parent, text="🔍 Filters", padding=8)
        filter_frame.pack(fill="x", pady=(0, 10))
        
        # Single row with all filters
        filter_row = ttk.Frame(filter_frame)
        filter_row.pack(fill="x")
        
        # Search
        ttk.Label(filter_row, text="Search:").pack(side="left", padx=(0, 5))
        search_entry = ttk.Entry(filter_row, textvariable=self.search_var, width=25)
        search_entry.pack(side="left", padx=(0, 15))
        self.search_var.trace_add("write", lambda *_: self.filter_data())
        
        # Status
        ttk.Label(filter_row, text="Status:").pack(side="left", padx=(0, 5))
        self.status_filter = ttk.Combobox(filter_row, width=12, state="readonly")
        self.status_filter['values'] = ("All", "Active", "Inactive", "active", "inactive")
        self.status_filter.current(0)
        self.status_filter.pack(side="left", padx=(0, 15))
        self.status_filter.bind("<<ComboboxSelected>>", lambda e: self.filter_data())
        
        # Type
        ttk.Label(filter_row, text="Type:").pack(side="left", padx=(0, 5))
        self.type_filter = ttk.Combobox(filter_row, width=12, state="readonly")
        self.type_filter['values'] = ("All", "atlassian", "app", "customer")
        self.type_filter.current(0)
        self.type_filter.pack(side="left", padx=(0, 15))
        self.type_filter.bind("<<ComboboxSelected>>", lambda e: self.filter_data())
        
        # Date range
        ttk.Label(filter_row, text="Last Active:").pack(side="left", padx=(0, 5))
        
        self.date_from_picker = DateEntry(
            filter_row, 
            width=10, 
            background='darkblue',
            foreground='white', 
            borderwidth=2,
            date_pattern='yyyy-mm-dd'
        )
        self.date_from_picker.pack(side="left", padx=(0, 3))
        self.date_from_picker.delete(0, 'end')
        self.date_from_picker.bind("<<DateEntrySelected>>", lambda e: self.filter_data())
        self.date_from_picker.bind("<KeyRelease>", lambda e: self.filter_data())
        
        ttk.Label(filter_row, text="to").pack(side="left", padx=3)
        
        self.date_to_picker = DateEntry(
            filter_row, 
            width=10, 
            background='darkblue',
            foreground='white', 
            borderwidth=2,
            date_pattern='yyyy-mm-dd'
        )
        self.date_to_picker.pack(side="left", padx=(0, 15))
        self.date_to_picker.delete(0, 'end')
        self.date_to_picker.bind("<<DateEntrySelected>>", lambda e: self.filter_data())
        self.date_to_picker.bind("<KeyRelease>", lambda e: self.filter_data())
        
        # Clear button
        ttk.Button(filter_row, text="Clear", command=self.clear_filters, width=8).pack(side="left")

        # Progress bar (shared by both sub-tabs)
        self.progress = ttk.Progressbar(parent, mode='indeterminate')
        self.progress.pack(fill="x", pady=(10, 10))
        self.progress.pack_forget()

        # Data view - single tree shared by both Users and Groups sub-tabs
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill="both", expand=True, pady=(0, 5))

        # Add "select" column to columns tuple
        self.tree = ttk.Treeview(
            tree_frame,
            columns=("select", "name", "email", "id", "type", "status", "last_active"),
            show="tree headings"
        )

        # Configure select column
        self.tree.heading("select", text="☑", command=self.toggle_select_all)
        self.tree.column("select", width=30, minwidth=30, anchor="center", stretch=False)
        
        # Configure other columns with stretch enabled
        for col, w in zip(
            ("name", "email", "id", "type", "status", "last_active"),
            (200, 250, 200, 100, 80, 180)
        ):
            self.tree.heading(col, text=col.replace("_", " ").title(), command=partial(self.sort_by_column, col))
            self.tree.column(col, width=w, minwidth=w, stretch=True)
        
        # Track selected items
        self.selected_items = set()

        ysb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        
        # Bind both TreeviewOpen event and double-click for cross-platform compatibility
        self.tree.bind("<<TreeviewOpen>>", self.on_group_expand)
        self.tree.bind("<Double-Button-1>", self.on_item_double_click)
        
        self.tree.tag_configure("member", background="#f0f0f0")
        self.tree.tag_configure("product", background="#e8f4f8")
        
        # Add right-click context menu for users
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<Button-2>", self.show_context_menu)
        
        # Create context menu
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="🔗 Open User Profile in Jira", command=self.open_user_profile)
        self.context_menu.add_separator()
        
        # User Management Actions submenu
        user_mgmt_menu = tk.Menu(self.context_menu, tearoff=0)
        self.context_menu.add_cascade(label="👤 User Management", menu=user_mgmt_menu)
        user_mgmt_menu.add_command(label="🚫 Deactivate User", command=self.deactivate_user)
        user_mgmt_menu.add_command(label="✅ Reactivate User", command=self.reactivate_user)
        user_mgmt_menu.add_separator()
        user_mgmt_menu.add_command(label="👥 Add to Group...", command=self.add_user_to_group)
        user_mgmt_menu.add_command(label="➖ Remove from Group...", command=self.remove_user_from_group)
        user_mgmt_menu.add_separator()
        user_mgmt_menu.add_command(label="📦 Manage Product Access...", command=self.manage_product_access)
        
        self.context_menu.add_separator()
        self.context_menu.add_command(label="📋 Copy Account ID", command=self.copy_account_id)
        self.context_menu.add_command(label="📧 Copy Email", command=self.copy_email)
        
        # Bind click on tree to handle checkbox toggle
        self.tree.bind("<Button-1>", self.on_tree_click)
        
        # Force update to ensure window dimensions are known
        self.tree.update_idletasks()
        
        # Get the actual available width and redistribute
        self.adjust_column_widths()

        # Sub-notebook for Users/Groups tabs (below tree now, just for labels/organization)
        # Sub-notebook for Users/Groups tabs (below tree now, just for labels/organization)
        sub_tabs_frame = ttk.Frame(parent)
        sub_tabs_frame.pack(fill="x", pady=(0, 5))
        
        self.data_notebook = ttk.Notebook(sub_tabs_frame)
        self.data_notebook.pack(fill="x")
        
        # Users sub-tab (minimal - just a label)
        users_info = ttk.Frame(self.data_notebook)
        self.data_notebook.add(users_info, text="👤 Users View")
        
        # Groups sub-tab (minimal - just a label)
        groups_info = ttk.Frame(self.data_notebook)
        self.data_notebook.add(groups_info, text="👥 Groups View")
        
        # Bind tab change to switch view
        self.data_notebook.bind("<<NotebookTabChanged>>", self.on_data_tab_changed)
        
        # Footer with result count
        footer_frame = ttk.Frame(parent)
        footer_frame.pack(fill="x", pady=(5, 0))
        
        ttk.Separator(footer_frame, orient="horizontal").pack(fill="x", pady=(0, 5))
        
        self.result_count_label = ttk.Label(
            footer_frame, 
            text="No data loaded - Click 'Fetch Users' or 'Fetch Groups'", 
            font=("", 9),
            foreground="gray"
        )
        self.result_count_label.pack(side="left", padx=10)
    
    def setup_groups_tab(self, parent):
        """Setup the Groups tab - REMOVED, now handled by sub-tabs"""
        pass
    def on_data_tab_changed(self, event):
        """Handle switching between Users and Groups sub-tabs"""
        selected_tab = self.data_notebook.index(self.data_notebook.select())
        if selected_tab == 0:  # Users tab
            self.current_view = "users"
            self.display_current_data()
        elif selected_tab == 1:  # Groups tab
            self.current_view = "groups"
            self.display_current_data()
    
    def display_current_data(self):
        """Display data based on current view"""
        if self.current_view == "users":
            if self.users_data:
                self.filter_data()  # This will call display_users with filters
        elif self.current_view == "groups":
            if self.groups_data:
                self.display_groups()  # Show all groups
    
    def display_groups(self):
        """Display all groups (unfiltered)"""
        self.clear_tree()
        for g in self.groups_data:
            item = self.tree.insert(
                "",
                "end",
                values=("", g["name"], "", g["groupId"], "", g.get("memberCount", "")),
                tags=("group",)
            )
            self.tree.insert(item, "end", values=("", "Loading...", "", "", "", "", ""), tags=("placeholder",))
        
        # Update count
        if hasattr(self, 'result_count_label'):
            self.result_count_label.config(
                text=f"Showing {len(self.groups_data)} group(s)",
                foreground="green"
            )
        
        # Adjust column widths to fill window
        self.root.after(10, self.adjust_column_widths)
    
    
    def setup_groups_tab(self, parent):
        """Setup the Groups tab - REMOVED, now handled by sub-tabs"""
        pass
        
    def setup_products_tab(self, parent):
        # Top action bar
        action_bar = ttk.Frame(parent)
        action_bar.pack(fill="x", pady=(0, 10))
        
        ttk.Label(action_bar, text="ℹ️ Fetch users with Org API to see product access", 
                 font=("", 9), foreground="blue").pack(side="left", padx=10)
        
        ttk.Button(action_bar, text="📊 Analyze Products", command=self.analyze_products, width=20).pack(side="left", padx=(20, 5))
        ttk.Button(action_bar, text="💾 Export Products CSV", command=self.export_products_csv, width=20).pack(side="left", padx=(0, 5))
        
        # Separator
        ttk.Separator(action_bar, orient="vertical").pack(side="left", fill="y", padx=10)
        
        # Status indicator for products tab
        self.products_status = ttk.Label(action_bar, text="Ready", foreground="blue", font=("", 10, "bold"))
        self.products_status.pack(side="left", padx=10)
        
        # Search bar for products
        search_frame = ttk.LabelFrame(parent, text="🔍 Search Products", padding=10)
        search_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(search_frame, text="Search:").pack(side="left")
        self.products_search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.products_search_var, width=40)
        search_entry.pack(side="left", padx=(10, 0))
        self.products_search_var.trace_add("write", lambda *_: self.filter_products())
        
        ttk.Button(search_frame, text="Clear", command=self.clear_products_search, width=12).pack(side="left", padx=(20, 0))
        
        # Products tree view
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill="both", expand=True)
        
        self.products_tree = ttk.Treeview(
            tree_frame,
            columns=("product", "user_count", "url", "last_active"),
            show="tree headings"
        )
        
        for col, w in zip(
            ("product", "user_count", "url", "last_active"),
            (250, 120, 300, 180)
        ):
            self.products_tree.heading(col, text=col.replace("_", " ").title())
            self.products_tree.column(col, width=w)
        
        ysb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.products_tree.yview)
        xsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.products_tree.xview)
        self.products_tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        
        self.products_tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        
        self.products_tree.bind("<<TreeviewOpen>>", self.on_product_expand)
        self.products_tree.tag_configure("product", background="#e8f4f8")
        self.products_tree.tag_configure("user", background="#f0f0f0")
        
        # Footer
        footer_frame = ttk.Frame(parent)
        footer_frame.pack(fill="x", pady=(5, 0))
        
        ttk.Separator(footer_frame, orient="horizontal").pack(fill="x", pady=(0, 5))
        
        self.products_count_label = ttk.Label(
            footer_frame, 
            text="No products analyzed", 
            font=("", 9),
            foreground="gray"
        )
        self.products_count_label.pack(side="left", padx=10)
    
    def clear_products_search(self):
        self.products_search_var.set("")
        self.filter_products()
    
    def analyze_products(self):
        """Analyze product access from user data"""
        if not self.users_data:
            messagebox.showwarning("No Data", "Please fetch users first (with Org API enabled)")
            return
        
        if not self.use_org_api.get():
            messagebox.showwarning("Org API Required", "Please enable Organization API and fetch users again to see product access")
            return
        
        self.products_status.config(text="Analyzing products...", foreground="orange")
        
        # Build product dictionary
        products = {}
        
        for user in self.users_data:
            user_name = user.get("name", "")
            user_email = user.get("email", "") or "(No email)"
            user_id = user.get("account_id", "")
            user_status = user.get("account_status", "")
            
            product_access = user.get("product_access", [])
            
            for product in product_access:
                product_name = product.get("name", "Unknown")
                product_key = product.get("key", "")
                product_url = product.get("url", "")
                last_active = product.get("last_active", "")
                
                # Format last active
                if last_active:
                    try:
                        dt = parser.isoparse(last_active)
                        last_active_formatted = dt.strftime("%Y-%m-%d %H:%M:%S")
                    except:
                        last_active_formatted = last_active
                else:
                    last_active_formatted = "Never"
                
                # Create unique product identifier
                product_id = f"{product_name}|{product_url}"
                
                if product_id not in products:
                    products[product_id] = {
                        "name": product_name,
                        "key": product_key,
                        "url": product_url,
                        "users": []
                    }
                
                products[product_id]["users"].append({
                    "name": user_name,
                    "email": user_email,
                    "id": user_id,
                    "status": user_status,
                    "last_active": last_active_formatted
                })
        
        self.products_data = products
        self.display_products(products)
        
        self.products_status.config(text=f"{len(products)} products found", foreground="green")
        self.products_count_label.config(text=f"Showing {len(products)} product(s)", foreground="green")
    
    def display_products(self, products):
        """Display products in the tree"""
        # Clear existing
        for item in self.products_tree.get_children():
            self.products_tree.delete(item)
        
        if not products:
            self.products_tree.insert("", "end", values=("No products found", "", "", ""))
            return
        
        # Sort products by name
        sorted_products = sorted(products.items(), key=lambda x: x[1]["name"])
        
        for product_id, product_data in sorted_products:
            user_count = len(product_data["users"])
            
            # Get most recent last_active across all users
            last_actives = [u["last_active"] for u in product_data["users"] if u["last_active"] != "Never"]
            most_recent = max(last_actives) if last_actives else "Never"
            
            product_item = self.products_tree.insert(
                "",
                "end",
                values=(
                    f"📦 {product_data['name']}",
                    f"{user_count} users",
                    product_data['url'],
                    most_recent
                ),
                tags=("product",)
            )
            
            # Add placeholder for expansion
            self.products_tree.insert(product_item, "end", values=("Loading users...", "", "", ""), tags=("placeholder",))
    
    def on_product_expand(self, _):
        """Handle product expansion to show users"""
        item = self.products_tree.focus()
        if "product" not in self.products_tree.item(item, "tags"):
            return
        
        # Check if already loaded
        children = self.products_tree.get_children(item)
        if children and "placeholder" not in self.products_tree.item(children[0], "tags"):
            return
        
        # Clear placeholder
        self.products_tree.delete(*children)
        
        # Get product info from the item
        values = self.products_tree.item(item, "values")
        product_name = values[0].replace("📦 ", "")
        product_url = values[2]
        product_id = f"{product_name}|{product_url}"
        
        product_data = self.products_data.get(product_id)
        if not product_data:
            return
        
        # Display users sorted by name
        users = sorted(product_data["users"], key=lambda x: x["name"])
        
        for user in users:
            self.products_tree.insert(
                item,
                "end",
                values=(
                    f"  👤 {user['name']}",
                    user['status'],
                    user['email'],
                    user['last_active']
                ),
                tags=("user",)
            )
    
    def filter_products(self):
        """Filter products based on search term"""
        if not self.products_data:
            return
        
        term = self.products_search_var.get().lower()
        
        if not term:
            self.display_products(self.products_data)
            return
        
        # Filter products
        filtered = {
            pid: pdata for pid, pdata in self.products_data.items()
            if term in pdata["name"].lower() or 
               term in pdata["url"].lower() or
               any(term in u["name"].lower() or term in u["email"].lower() for u in pdata["users"])
        }
        
        self.display_products(filtered)
        self.products_count_label.config(text=f"Showing {len(filtered)} product(s)", foreground="green")
    
    def export_products_csv(self):
        """Export products and their users to CSV"""
        if not self.products_data:
            messagebox.showwarning("No Data", "Please analyze products first")
            return
        
        filename = f"jira_products_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        
        with open(filename, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Product Name", "Product URL", "User Name", "User Email", "User Status", "Last Active in Product"])
            
            for product_id, product_data in sorted(self.products_data.items(), key=lambda x: x[1]["name"]):
                for user in sorted(product_data["users"], key=lambda x: x["name"]):
                    writer.writerow([
                        product_data["name"],
                        product_data["url"],
                        user["name"],
                        user["email"],
                        user["status"],
                        user["last_active"]
                    ])
        
        messagebox.showinfo("Exported", f"Products exported to {filename}")
        self.products_status.config(text="Export complete", foreground="green")

    def toggle_org_api(self):
        if self.use_org_api.get():
            # Show org fields
            for child in self.org_fields_frame.winfo_children():
                child.configure(state="normal")
            self.org_entry.config(state="readonly")  # Keep ID readonly
        else:
            # Disable org fields
            for child in self.org_fields_frame.winfo_children():
                if isinstance(child, (ttk.Entry, ttk.Button, ttk.Combobox)):
                    child.configure(state="disabled")
    
    def show_org_api_help(self):
        help_text = """To use the Organization API, you need to create an Organization API Key:

1. Go to: https://admin.atlassian.com/
2. Select your organization
3. Click "Organization settings" in the left sidebar
4. Click "API keys"
5. Click "Create API key" (choose UNSCOPED key for best results)
6. Give it a name (e.g., "User Management")
7. Set expiration date (max 1 year)
8. Click "Create"
9. COPY THE KEY IMMEDIATELY (you can't see it again!)
10. Paste it in the "Org API Key" field above

Note: Use an UNSCOPED API key for full access to user endpoints.
"""
        messagebox.showinfo("How to Create Organization API Key", help_text)

    def clear_filters(self):
        self.search_var.set("")
        self.status_filter.current(0)
        self.type_filter.current(0)
        # Clear date pickers by setting them to empty
        self.date_from_picker.set_date(datetime.now())
        self.date_to_picker.set_date(datetime.now())
        # Delete the entries to make them empty
        self.date_from_picker.delete(0, 'end')
        self.date_to_picker.delete(0, 'end')
        self.filter_data()
    
    def update_column_visibility(self):
        """Update which columns are visible in the treeview"""
        visible_cols = [col for col, var in self.column_vars.items() if var.get()]
        
        if not visible_cols:
            # At least one column must be visible
            messagebox.showwarning("Warning", "At least one column must be visible")
            # Re-enable the last unchecked column
            for col, var in self.column_vars.items():
                if not var.get():
                    var.set(True)
                    break
            return
        
        self.tree.configure(displaycolumns=visible_cols)

    # ---------------- Credentials ---------------- #
    def load_credentials(self):
        self.jira_url.set(keyring.get_password(SERVICE_NAME, "jira_url") or "")
        self.email.set(keyring.get_password(SERVICE_NAME, "email") or "")
        self.org_id.set("")  # Don't load org_id from keyring
        self.api_token.set("")
        self.org_api_key.set("")

    def save_credentials(self):
        if self.remember_creds.get():
            keyring.set_password(SERVICE_NAME, "jira_url", self.jira_url.get())
            keyring.set_password(SERVICE_NAME, "email", self.email.get())
            # Don't save org_id
        else:
            try:
                keyring.delete_password(SERVICE_NAME, "jira_url")
                keyring.delete_password(SERVICE_NAME, "email")
            except:
                pass
        
        # Always try to remove org_id from keyring if it exists
        try:
            keyring.delete_password(SERVICE_NAME, "org_id")
        except:
            pass

    def on_close(self):
        self.save_credentials()
        self.root.destroy()

    # ---------------- Utilities ---------------- #
    def auth(self):
        return HTTPBasicAuth(self.email.get().strip(), self.api_token.get().strip())

    
    def adjust_column_widths(self):
        """Adjust column widths to fill available space proportionally"""
        try:
            # Get the tree's actual width
            tree_width = self.tree.winfo_width()
            
            # If width is too small (not yet rendered), skip
            if tree_width <= 1:
                return
            
            # Account for checkbox column (fixed 30px) and scrollbar (~20px)
            available_width = tree_width - 30 - 20
            
            # Define proportions for each column (should sum to 1.0)
            proportions = {
                "name": 0.20,        # 20%
                "email": 0.25,       # 25%
                "id": 0.20,          # 20%
                "type": 0.10,        # 10%
                "status": 0.08,      # 8%
                "last_active": 0.17  # 17%
            }
            
            # Calculate and set widths
            for col, proportion in proportions.items():
                width = int(available_width * proportion)
                self.tree.column(col, width=width)
        except Exception as e:
            # If anything goes wrong, silently continue
            pass
    
    def clear_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        # Clear selections when clearing tree
        self.selected_items.clear()
        self.update_bulk_edit_button()

    def clear_data(self):
        self.clear_tree()
        self.users_data = []
        self.groups_data = []
        self.groups_members = {}
        self.users_product_access = {}
        self.products_data = {}
        self.current_view = "users"
        self.tree.configure(show="headings")
        self.clear_filters()
        self.status.config(text="Cleared", foreground="blue")
        self.result_count_label.config(text="No results loaded", foreground="gray")
        
        # Clear products tree too
        for item in self.products_tree.get_children():
            self.products_tree.delete(item)
        self.products_count_label.config(text="No products analyzed", foreground="gray")

    # ---------------- Organization ID ---------------- #
    def fetch_org_id_async(self):
        threading.Thread(target=self.fetch_org_id, daemon=True).start()

    def fetch_org_id(self):
        org_api_key = self.org_api_key.get().strip()
        if not org_api_key:
            self.root.after(0, lambda: messagebox.showerror(
                "Error", 
                "Please enter your Organization API Key first.\n\nClick 'Help: Create Org API Key' for instructions."
            ))
            return
            
        self.root.after(0, lambda: self.status.config(text="Fetching organization ID...", foreground="orange"))
        session = create_session()
        try:
            r = session.get(
                "https://api.atlassian.com/admin/v1/orgs",
                headers={
                    "Accept": "application/json",
                    "Authorization": f"Bearer {org_api_key}"
                },
                timeout=30
            )
            r.raise_for_status()
            data = r.json()
            
            if data.get("data") and len(data["data"]) > 0:
                org = data["data"][0]
                org_id = org.get("id")
                org_name = org.get("attributes", {}).get("name", "Unknown")
                
                self.org_id.set(org_id)
                self.root.after(0, lambda: messagebox.showinfo(
                    "Organization Found", 
                    f"Organization: {org_name}\nID: {org_id}"
                ))
                self.root.after(0, lambda: self.status.config(text="Organization ID retrieved", foreground="green"))
            else:
                raise Exception("No organizations found for this account")
        except Exception as e:
            error_msg = f"Could not fetch org ID: {str(e)}"
            print(error_msg)
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            self.root.after(0, lambda: self.status.config(text="Failed to get org ID", foreground="red"))
        finally:
            session.close()

    # ---------------- Token Validation ---------------- #
    def validate_token_async(self):
        threading.Thread(target=self.validate_token, daemon=True).start()

    def validate_token(self):
        self.root.after(0, lambda: self.status.config(text="Validating token...", foreground="orange"))
        try:
            r = requests.get(
                f"{self.jira_url.get().rstrip('/')}/rest/api/3/myself",
                auth=self.auth(),
                headers={"Accept": "application/json"}
            )
            if r.status_code == 200:
                self.root.after(0, lambda: messagebox.showinfo("Success", "API token is valid"))
                self.root.after(0, lambda: self.status.config(text="Token valid", foreground="green"))
            else:
                raise Exception(r.text)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Invalid Token", str(e)))
            self.root.after(0, lambda: self.status.config(text="Token invalid", foreground="red"))
    
    def validate_and_fetch_all(self):
        """Validate token and automatically fetch users and groups"""
        threading.Thread(target=self._validate_and_fetch_all_thread, daemon=True).start()
    
    def _validate_and_fetch_all_thread(self):
        """Thread worker for validating and fetching all data"""
        self.root.after(0, lambda: self.status.config(text="Validating token...", foreground="orange"))
        try:
            # Validate token first
            r = requests.get(
                f"{self.jira_url.get().rstrip('/')}/rest/api/3/myself",
                auth=self.auth(),
                headers={"Accept": "application/json"}
            )
            if r.status_code != 200:
                raise Exception(r.text)
            
            self.root.after(0, lambda: self.status.config(text="Token valid! Loading data...", foreground="green"))
            
            # Auto-fetch users
            self.root.after(100, self.fetch_users_async)
            
            # Auto-fetch groups (after a small delay)
            self.root.after(1000, self.fetch_groups_async)
            
            # Switch to Users tab
            self.root.after(1500, lambda: self.notebook.select(1))  # Tab index 1 is Users
            
            self.root.after(2000, lambda: messagebox.showinfo(
                "Success", 
                "Token validated!\nUsers and groups are loading..."
            ))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Invalid Token", str(e)))
            self.root.after(0, lambda: self.status.config(text="Token invalid", foreground="red"))

    # ---------------- Async Wrappers ---------------- #
    def fetch_users_async(self):
        self.progress.pack(fill="x", padx=10, pady=(0,10))
        self.progress.start()
        threading.Thread(target=self.fetch_users, daemon=True).start()

    def fetch_groups_async(self):
        self.progress.pack(fill="x", padx=10, pady=(0,10))
        self.progress.start()
        threading.Thread(target=self.fetch_groups, daemon=True).start()

    # ---------------- Users ---------------- #
    def fetch_users(self):
        if self.use_org_api.get():
            self.fetch_users_org_api()
        else:
            self.fetch_users_standard_api()

    def fetch_users_standard_api(self):
        self.root.after(0, lambda: self.status.config(text="Fetching users (Standard API)...", foreground="orange"))
        self.current_view = "users"
        self.tree.configure(show="headings")

        users = []
        start = 0
        max_results = 1000
        
        session = create_session()

        try:
            page = 0
            while True:
                page += 1
                self.root.after(0, lambda p=page, u=len(users): self.status.config(
                    text=f"Fetching users page {p}... ({u} so far)", 
                    foreground="orange"
                ))
                
                print(f"Fetching page {page}, start={start}...")
                
                r = session.get(
                    f"{self.jira_url.get().rstrip('/')}/rest/api/3/users/search",
                    params={
                        "startAt": start, 
                        "maxResults": max_results
                    },
                    auth=self.auth(),
                    headers={"Accept": "application/json"},
                    timeout=30
                )
                r.raise_for_status()
                batch = r.json()
                
                if not batch:
                    break
                    
                print(f"Page {page}: got {len(batch)} users")
                users.extend(batch)
                start += max_results
                
                time.sleep(0.3)

            print(f"\nTotal users fetched: {len(users)}")
            
            self.users_data = users
            self.root.after(0, lambda: self.display_users(users))
            self.root.after(0, lambda: self.status.config(
                text=f"{len(users)} users loaded (no last login data available)", 
                foreground="orange"
            ))
        except Exception as e:
            error_msg = f"Error fetching users: {str(e)}"
            print(error_msg)
            import traceback
            traceback.print_exc()
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            self.root.after(0, lambda: self.status.config(text=error_msg, foreground="red"))
        finally:
            session.close()
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.progress.pack_forget())

    def fetch_users_org_api(self):
        org_id = self.org_id.get().strip()
        org_api_key = self.org_api_key.get().strip()
        
        if not org_id:
            self.root.after(0, lambda: messagebox.showerror(
                "Error", 
                "Please enter an Organization ID or click 'Get Org ID' first"
            ))
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.progress.pack_forget())
            return
        
        if not org_api_key:
            self.root.after(0, lambda: messagebox.showerror(
                "Error", 
                "Please enter your Organization API Key.\n\nClick 'Help: Create Org API Key' for instructions."
            ))
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.progress.pack_forget())
            return

        self.root.after(0, lambda: self.status.config(text="Fetching users (Org API)...", foreground="orange"))
        self.current_view = "users"
        self.tree.configure(show="headings")

        users = []
        cursor = None
        page = 0
        
        session = create_session()

        try:
            while True:
                page += 1
                self.root.after(0, lambda p=page: self.status.config(
                    text=f"Fetching users page {p}... ({len(users)} so far)", 
                    foreground="orange"
                ))
                
                params = {}
                if cursor:
                    params["cursor"] = cursor

                print(f"Fetching page {page}...")
                
                try:
                    r = session.get(
                        f"https://api.atlassian.com/admin/v1/orgs/{org_id}/users",
                        params=params,
                        headers={
                            "Accept": "application/json",
                            "Authorization": f"Bearer {org_api_key}"
                        },
                        timeout=30
                    )
                    r.raise_for_status()
                except requests.exceptions.Timeout:
                    print(f"Timeout on page {page}, retrying...")
                    time.sleep(2)
                    r = session.get(
                        f"https://api.atlassian.com/admin/v1/orgs/{org_id}/users",
                        params=params,
                        headers={
                            "Accept": "application/json",
                            "Authorization": f"Bearer {org_api_key}"
                        },
                        timeout=60
                    )
                    r.raise_for_status()
                
                data = r.json()
                
                if not users and data.get("data"):
                    print("\nDEBUG - First user from Org API:")
                    print(json.dumps(data["data"][0], indent=2))
                
                batch = data.get("data", [])
                if not batch:
                    break
                    
                users.extend(batch)
                print(f"Page {page}: got {len(batch)} users, total: {len(users)}")
                
                links = data.get("links", {})
                if links.get("next"):
                    next_url = links["next"]
                    if "cursor=" in next_url:
                        cursor = next_url.split("cursor=")[-1]
                        cursor = cursor.split("&")[0]
                    else:
                        break
                else:
                    break
                
                time.sleep(0.5)

            print(f"\nTotal users fetched: {len(users)}")
            
            self.users_data = users
            self.root.after(0, lambda: self.display_users_org(users))
            self.root.after(0, lambda: self.status.config(
                text=f"{len(users)} users loaded with last login data", 
                foreground="green"
            ))
        except Exception as e:
            error_msg = f"Error fetching users from Org API: {str(e)}"
            print(error_msg)
            import traceback
            traceback.print_exc()
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            self.root.after(0, lambda: self.status.config(text=error_msg, foreground="red"))
        finally:
            session.close()
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.progress.pack_forget())

    def display_users(self, users):
        self.clear_tree()
        # For standard API, we don't have product access data
        # So we won't make users expandable
        for u in users:
            last_active = "N/A (use Org API)"
            
            email = u.get("emailAddress", "")
            # Debug: print if email seems incomplete
            if email and "@" not in email:
                print(f"DEBUG - Incomplete email for user {u.get('displayName', 'Unknown')}: '{email}'")
                print(f"  Full user data: {u}")
            
            # Show "(No email)" if email is empty
            if not email:
                email = "(No email)"

            self.tree.insert(
                "",
                "end",
                values=(
                    "☐",  # Checkbox - unchecked by default
                    u.get("displayName", ""),
                    email,
                    u.get("accountId", ""),
                    u.get("accountType", ""),
                    "Active" if u.get("active") else "Inactive",
                    last_active
                ),
                tags=("user",)  # Add user tag for selection
            )
        
        # Update footer count
        self.result_count_label.config(
            text=f"Showing {len(users)} user(s)",
            foreground="green"
        )
        
        # Adjust column widths to fill window
        self.root.after(10, self.adjust_column_widths)

    def display_users_org(self, users):
        self.clear_tree()
        # Enable tree view for expandable users
        self.tree.configure(show="tree headings")
        
        for u in users:
            account_id = u.get("account_id", "")
            name = u.get("name", "")
            email = u.get("email", "")
            account_type = u.get("account_type", "")
            account_status = u.get("account_status", "")
            
            # Debug: print if email seems incomplete
            if email and "@" not in email:
                print(f"DEBUG - Incomplete email for user {name}: '{email}'")
                print(f"  Full user data: {u}")
            
            # Show "(No email)" if email is empty
            if not email:
                email = "(No email)"
            
            last_active = u.get("last_active", "")
            
            if last_active:
                try:
                    dt = parser.isoparse(last_active)
                    last_active = dt.strftime("%Y-%m-%d %H:%M:%S")
                except Exception as e:
                    print(f"Error parsing date for {name}: {e}, raw value: {last_active}")
            else:
                last_active = "Never logged in"

            # Insert user as expandable item
            user_item = self.tree.insert(
                "",
                "end",
                values=(
                    "☐",  # Checkbox - unchecked by default
                    name,
                    email,
                    account_id,
                    account_type,
                    account_status,
                    last_active
                ),
                tags=("user",)
            )
            
            # Store product access data for this user
            product_access = u.get("product_access", [])
            if product_access:
                self.users_product_access[account_id] = product_access
                # Add placeholder to make it expandable (7 values to match column count)
                self.tree.insert(user_item, "end", values=("", "Loading products...", "", "", "", "", ""), tags=("placeholder",))
        
        # Update footer count
        self.result_count_label.config(
            text=f"Showing {len(users)} user(s)",
            foreground="green"
        )
        
        # Adjust column widths to fill window
        self.root.after(10, self.adjust_column_widths)

    # ---------------- Groups ---------------- #
    def fetch_groups(self):
        self.root.after(0, lambda: self.status.config(text="Fetching groups...", foreground="orange"))
        self.current_view = "groups"
        self.tree.configure(show="tree headings")
        self.clear_tree()

        groups = []
        start = 0
        max_results = 50

        try:
            while True:
                r = requests.get(
                    f"{self.jira_url.get().rstrip('/')}/rest/api/3/group/bulk",
                    params={"startAt": start, "maxResults": max_results},
                    auth=self.auth(),
                    headers={"Accept": "application/json"}
                )
                r.raise_for_status()
                data = r.json()
                groups.extend(data.get("values", []))
                if data.get("isLast", True):
                    break
                start += max_results

            self.groups_data = groups

            def populate():
                for g in groups:
                    item = self.tree.insert(
                        "",
                        "end",
                        values=("", g["name"], "", g["groupId"], "", g.get("memberCount", "")),  # Empty checkbox for groups
                        tags=("group",)
                    )
                    self.tree.insert(item, "end", values=("", "Loading...", "", "", "", "", ""), tags=("placeholder",))
                self.status.config(text=f"{len(groups)} groups loaded", foreground="green")
                self.progress.stop()
                self.progress.pack_forget()
                
                # Update footer count
                self.result_count_label.config(
                    text=f"Showing {len(groups)} group(s)",
                    foreground="green"
                )

            self.root.after(0, populate)
        except Exception as e:
            error_msg = f"Error fetching groups: {str(e)}"
            print(error_msg)
            self.root.after(0, lambda: self.status.config(text=error_msg, foreground="red"))
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.progress.pack_forget())

    def on_group_expand(self, _):
        item = self.tree.focus()
        tags = self.tree.item(item, "tags")
        
        # Handle group expansion
        if "group" in tags:
            values = self.tree.item(item, "values")
            group_name = values[1] if len(values) > 1 else ""  # Index 1 because of checkbox column at index 0
            
            if not group_name or group_name in self.groups_members:
                return

            self.tree.delete(*self.tree.get_children(item))

            try:
                r = requests.get(
                    f"{self.jira_url.get().rstrip('/')}/rest/api/3/group/member",
                    params={
                        "groupname": group_name, 
                        "maxResults": 1000
                    },
                    auth=self.auth(),
                    headers={"Accept": "application/json"}
                )
                r.raise_for_status()

                members = r.json().get("values", [])
                self.groups_members[group_name] = members

                for m in members:
                    last_active = "N/A"

                    self.tree.insert(
                        item,
                        "end",
                        values=(
                            "",  # Empty checkbox for group members
                            m.get("displayName", ""),
                            m.get("emailAddress", ""),
                            m.get("accountId", ""),
                            m.get("accountType", ""),
                            "Active" if m.get("active") else "Inactive",
                            last_active
                        ),
                        tags=("member",)
                    )
            except Exception as e:
                print(f"Error loading group members: {e}")
        
        # Handle user expansion (show product access)
        elif "user" in tags:
            values = self.tree.item(item, "values")
            account_id = values[3] if len(values) > 3 else ""  # Index 3: checkbox(0), name(1), email(2), account_id(3)
            
            # Check if already loaded
            children = self.tree.get_children(item)
            if children and "placeholder" not in self.tree.item(children[0], "tags"):
                return
            
            # Clear placeholder
            self.tree.delete(*children)
            
            # Get product access for this user
            product_access = self.users_product_access.get(account_id, [])
            
            if not product_access:
                self.tree.insert(
                    item,
                    "end",
                    values=("", "No product access data", "", "", "", "", ""),  # Add checkbox column
                    tags=("product",)
                )
                return
            
            # Display each product
            for product in product_access:
                product_name = product.get("name", "Unknown")
                product_url = product.get("url", "")
                product_last_active = product.get("last_active", "")
                
                # Format last active date
                if product_last_active:
                    try:
                        dt = parser.isoparse(product_last_active)
                        product_last_active = dt.strftime("%Y-%m-%d %H:%M:%S")
                    except:
                        pass
                else:
                    product_last_active = "Never"
                
                self.tree.insert(
                    item,
                    "end",
                    values=(
                        "",  # Empty checkbox for product items
                        f"  📦 {product_name}",
                        "",
                        product_url,
                        "",
                        "",
                        product_last_active
                    ),
                    tags=("product",)
                )

    def on_item_double_click(self, event):
        """Handle double-click to expand/collapse items (for Windows compatibility)"""
        # Get the item that was clicked
        item = self.tree.identify_row(event.y)
        if not item:
            return
        
        # Check if the item has children or is expandable
        tags = self.tree.item(item, "tags")
        
        # If it's a group or user with expandable content
        if "group" in tags or "user" in tags:
            # Check current state
            if self.tree.item(item, "open"):
                # If already open, close it
                self.tree.item(item, open=False)
            else:
                # If closed, open it (this will trigger on_group_expand via TreeviewOpen)
                self.tree.item(item, open=True)
                # Also manually call on_group_expand for Windows compatibility
                self.on_group_expand(event)

    # ---------------- Sorting ---------------- #
    def sort_by_column(self, col):
        """Sort tree contents by column - improved for cross-platform compatibility"""
        # Get only top-level items (not children)
        items = [(self.tree.set(i, col), i) for i in self.tree.get_children("")]
        
        # Toggle sort direction
        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_reverse = False
            self.sort_column = col
        
        # Sort items
        try:
            items.sort(
                reverse=self.sort_reverse, 
                key=lambda x: (x[0].lower() if isinstance(x[0], str) else str(x[0]))
            )
        except Exception as e:
            print(f"Sort error: {e}")
            return
        
        # Reorder items in the tree
        for idx, (_, iid) in enumerate(items):
            self.tree.move(iid, "", idx)
        
        # Update column heading to show sort direction
        for column in ("name", "email", "id", "type", "status", "last_active"):
            heading_text = column.replace("_", " ").title()
            if column == col:
                # Add arrow indicator
                arrow = " ▼" if self.sort_reverse else " ▲"
                self.tree.heading(column, text=heading_text + arrow)
            else:
                self.tree.heading(column, text=heading_text)

    # ---------------- Filtering ---------------- #
    def filter_data(self):
        if self.current_view == "users":
            self.filter_users()
        elif self.current_view == "groups":
            self.filter_groups()
    
    def filter_users(self):
        term = self.search_var.get().lower()
        status_filter = self.status_filter.get()
        type_filter = self.type_filter.get()
        
        # Get dates from DateEntry widgets
        date_from_str = self.date_from_picker.get().strip()
        date_to_str = self.date_to_picker.get().strip()
        
        # Parse date filters
        date_from = None
        date_to = None
        try:
            if date_from_str:
                date_from = datetime.strptime(date_from_str, "%Y-%m-%d")
            if date_to_str:
                date_to = datetime.strptime(date_to_str, "%Y-%m-%d")
                date_to = date_to.replace(hour=23, minute=59, second=59)
        except ValueError:
            pass
        
        if self.use_org_api.get():
            filtered = []
            for u in self.users_data:
                # Search filter
                if term:
                    if not (term in (u.get("name", "").lower() or "") or
                           term in (u.get("email", "").lower() or "") or
                           term in (u.get("account_id", "").lower() or "")):
                        continue
                
                # Status filter
                if status_filter != "All":
                    user_status = u.get("account_status", "")
                    if user_status.lower() != status_filter.lower():
                        continue
                
                # Type filter
                if type_filter != "All":
                    user_type = u.get("account_type", "")
                    if user_type != type_filter:
                        continue
                
                # Date filter
                if date_from or date_to:
                    last_active = u.get("last_active", "")
                    if last_active:
                        try:
                            user_date = parser.isoparse(last_active)
                            if date_from and user_date < date_from:
                                continue
                            if date_to and user_date > date_to:
                                continue
                        except:
                            continue
                    else:
                        continue
                
                filtered.append(u)
            
            self.display_users_org(filtered)
        else:
            filtered = []
            for u in self.users_data:
                # Search filter
                if term:
                    if not (term in (u.get("displayName", "").lower() or "") or
                           term in (u.get("emailAddress", "").lower() or "") or
                           term in (u.get("accountId", "").lower() or "")):
                        continue
                
                # Status filter
                if status_filter != "All":
                    is_active = u.get("active", False)
                    if status_filter == "Active" and not is_active:
                        continue
                    if status_filter == "Inactive" and is_active:
                        continue
                
                # Type filter
                if type_filter != "All":
                    user_type = u.get("accountType", "")
                    if user_type != type_filter:
                        continue
                
                filtered.append(u)
            
            self.display_users(filtered)
    
    def filter_groups(self):
        # Get search term - if groups_search_var doesn't exist or is empty, show all
        term = ""
        if hasattr(self, 'groups_search_var'):
            term = self.groups_search_var.get().lower()
        
        if term:
            # Filter groups by search term
            filtered_groups = [
                g for g in self.groups_data
                if term in (g.get("name", "").lower() or "") or
                   term in (g.get("groupId", "").lower() or "")
            ]
        else:
            # Show all groups when no search term
            filtered_groups = self.groups_data
        
        self.clear_tree()
        for g in filtered_groups:
            item = self.tree.insert(
                "",
                "end",
                values=("", g["name"], "", g["groupId"], "", g.get("memberCount", "")),  # Empty checkbox for groups
                tags=("group",)
            )
            self.tree.insert(item, "end", values=("", "Loading...", "", "", "", "", ""), tags=("placeholder",))
        
        # Update result count label
        if hasattr(self, 'result_count_label'):
            self.result_count_label.config(
                text=f"Showing {len(filtered_groups)} group(s)",
                foreground="green"
            )
    
    # ---------------- Checkbox Selection ---------------- #
    def on_tree_click(self, event):
        """Handle clicks on the tree - toggle checkboxes if clicking select column"""
        region = self.tree.identify_region(event.x, event.y)
        column = self.tree.identify_column(event.x)
        
        # Check if clicked on the select column (#0 is tree column, #1 is first data column which is "select")
        if column == "#1":  # Select column
            item = self.tree.identify_row(event.y)
            if item:
                tags = self.tree.item(item, "tags")
                # Only allow selection on top-level user items (not groups, not children)
                if self.current_view == "users" and "user" in tags:
                    self.toggle_item_selection(item)
                    return "break"  # Prevent default behavior
    
    def toggle_item_selection(self, item):
        """Toggle selection state of an item"""
        if item in self.selected_items:
            self.selected_items.remove(item)
            self.tree.set(item, "select", "☐")
        else:
            self.selected_items.add(item)
            self.tree.set(item, "select", "☑")
        
        # Update bulk edit button state
        self.update_bulk_edit_button()
    
    def toggle_select_all(self):
        """Toggle selection of all visible items"""
        if self.current_view != "users":
            return
        
        # Get all top-level items with "user" tag
        all_items = [item for item in self.tree.get_children("") if "user" in self.tree.item(item, "tags")]
        
        if not all_items:
            return
        
        # If all are selected, deselect all. Otherwise, select all
        if len(self.selected_items) == len(all_items):
            # Deselect all
            for item in all_items:
                self.selected_items.discard(item)
                self.tree.set(item, "select", "☐")
        else:
            # Select all
            for item in all_items:
                self.selected_items.add(item)
                self.tree.set(item, "select", "☑")
        
        self.update_bulk_edit_button()
    
    def clear_all_selections(self):
        """Clear all selections"""
        for item in list(self.selected_items):
            self.tree.set(item, "select", "☐")
        self.selected_items.clear()
        self.update_bulk_edit_button()
    
    def update_bulk_edit_button(self):
        """Enable/disable bulk edit button based on selection"""
        if len(self.selected_items) > 0:
            self.bulk_edit_btn.config(state="normal")
        else:
            self.bulk_edit_btn.config(state="disabled")
    
    # ---------------- Context Menu ---------------- #
    def show_context_menu(self, event):
        """Show right-click context menu"""
        # Select the item under cursor
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            # Only show menu if we're in users view (not groups)
            if self.current_view == "users":
                try:
                    self.context_menu.tk_popup(event.x_root, event.y_root)
                finally:
                    self.context_menu.grab_release()
    
    def open_user_profile(self):
        """Open the selected user's profile in Jira user management"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        values = self.tree.item(item, "values")
        
        # Account ID is in index 2
        account_id = values[2] if len(values) > 2 else ""
        
        if not account_id or account_id == "(No email)":
            messagebox.showwarning("No Account ID", "Cannot open profile - no account ID available")
            return
        
        # Construct the user management URL
        jira_base = self.jira_url.get().rstrip('/')
        # Extract the site name from the URL (e.g., "hard-rock-digital" from "https://hard-rock-digital.atlassian.net")
        site_name = jira_base.replace("https://", "").replace("http://", "").split(".")[0]
        
        # The user management URL format
        profile_url = f"https://admin.atlassian.com/s/{site_name}/users/{account_id}"
        
        print(f"Opening user profile: {profile_url}")
        webbrowser.open(profile_url)
    
    def copy_account_id(self):
        """Copy the selected user's account ID to clipboard"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        values = self.tree.item(item, "values")
        account_id = values[2] if len(values) > 2 else ""
        
        if account_id:
            self.root.clipboard_clear()
            self.root.clipboard_append(account_id)
            self.status.config(text="Account ID copied to clipboard", foreground="blue")
    
    def copy_email(self):
        """Copy the selected user's email to clipboard"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        values = self.tree.item(item, "values")
        email = values[1] if len(values) > 1 else ""
        
        if email and email != "(No email)":
            self.root.clipboard_clear()
            self.root.clipboard_append(email)
            self.status.config(text="Email copied to clipboard", foreground="blue")
        else:
            messagebox.showinfo("No Email", "This user has no email address")
    
    # ---------------- User Management Actions ---------------- #
    def get_selected_user_info(self):
        """Helper to get selected user information"""
        selection = self.tree.selection()
        if not selection:
            return None
        
        item = selection[0]
        values = self.tree.item(item, "values")
        
        if len(values) < 4:  # Now need at least 4: checkbox + name + email + account_id
            return None
        
        return {
            "name": values[1],  # Shifted by 1 due to checkbox column
            "email": values[2],
            "account_id": values[3],
            "item": item
        }
    
    def deactivate_user(self):
        """Deactivate a user account"""
        user = self.get_selected_user_info()
        if not user:
            messagebox.showwarning("No Selection", "Please select a user")
            return
        
        # Confirm action
        confirm = messagebox.askyesno(
            "Confirm Deactivation",
            f"Are you sure you want to DEACTIVATE this user?\n\n"
            f"Name: {user['name']}\n"
            f"Email: {user['email']}\n"
            f"Account ID: {user['account_id']}\n\n"
            f"⚠️ This will revoke their access to Jira."
        )
        
        if not confirm:
            return
        
        threading.Thread(target=self._deactivate_user_thread, args=(user,), daemon=True).start()
    
    def _deactivate_user_thread(self, user):
        """Thread worker for deactivating user"""
        try:
            self.root.after(0, lambda: self.status.config(text=f"Deactivating {user['name']}...", foreground="orange"))
            
            # Note: Jira Cloud doesn't have a direct API to deactivate users
            # This requires the Organization API (admin.atlassian.com)
            if not self.org_api_key.get():
                self.root.after(0, lambda: messagebox.showerror(
                    "Organization API Required",
                    "User deactivation requires the Organization API.\n\n"
                    "Please enable and configure the Organization API in the Configuration tab."
                ))
                return
            
            # Deactivate via Organization API
            url = f"https://api.atlassian.com/users/{user['account_id']}/manage/lifecycle/disable"
            
            response = requests.post(
                url,
                headers={
                    "Authorization": f"Bearer {self.org_api_key.get()}",
                    "Accept": "application/json"
                },
                timeout=30
            )
            
            if response.status_code in [200, 204]:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Success",
                    f"User {user['name']} has been deactivated successfully."
                ))
                self.root.after(0, lambda: self.status.config(text="User deactivated", foreground="green"))
                # Refresh the user list
                self.root.after(0, self.fetch_users_async)
            else:
                raise Exception(f"API returned status {response.status_code}: {response.text}")
                
        except Exception as e:
            error_msg = f"Failed to deactivate user: {str(e)}"
            print(error_msg)
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            self.root.after(0, lambda: self.status.config(text="Deactivation failed", foreground="red"))
    
    def reactivate_user(self):
        """Reactivate a user account"""
        user = self.get_selected_user_info()
        if not user:
            messagebox.showwarning("No Selection", "Please select a user")
            return
        
        confirm = messagebox.askyesno(
            "Confirm Reactivation",
            f"Reactivate this user?\n\n"
            f"Name: {user['name']}\n"
            f"Email: {user['email']}\n"
            f"Account ID: {user['account_id']}"
        )
        
        if not confirm:
            return
        
        threading.Thread(target=self._reactivate_user_thread, args=(user,), daemon=True).start()
    
    def _reactivate_user_thread(self, user):
        """Thread worker for reactivating user"""
        try:
            self.root.after(0, lambda: self.status.config(text=f"Reactivating {user['name']}...", foreground="orange"))
            
            if not self.org_api_key.get():
                self.root.after(0, lambda: messagebox.showerror(
                    "Organization API Required",
                    "User reactivation requires the Organization API.\n\n"
                    "Please enable and configure the Organization API in the Configuration tab."
                ))
                return
            
            url = f"https://api.atlassian.com/users/{user['account_id']}/manage/lifecycle/enable"
            
            response = requests.post(
                url,
                headers={
                    "Authorization": f"Bearer {self.org_api_key.get()}",
                    "Accept": "application/json"
                },
                timeout=30
            )
            
            if response.status_code in [200, 204]:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Success",
                    f"User {user['name']} has been reactivated successfully."
                ))
                self.root.after(0, lambda: self.status.config(text="User reactivated", foreground="green"))
                self.root.after(0, self.fetch_users_async)
            else:
                raise Exception(f"API returned status {response.status_code}: {response.text}")
                
        except Exception as e:
            error_msg = f"Failed to reactivate user: {str(e)}"
            print(error_msg)
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            self.root.after(0, lambda: self.status.config(text="Reactivation failed", foreground="red"))
    
    def add_user_to_group(self):
        """Add selected user to a group"""
        user = self.get_selected_user_info()
        if not user:
            messagebox.showwarning("No Selection", "Please select a user")
            return
        
        # Fetch groups if not already loaded
        if not self.groups_data:
            self.status.config(text="Loading groups...", foreground="orange")
            self.fetch_groups()
            # Wait a moment for groups to load
            self.root.after(1500, lambda: self._show_group_selector(user, "add"))
            return
        
        self._show_group_selector(user, "add")
    
    def remove_user_from_group(self):
        """Remove selected user from a group"""
        user = self.get_selected_user_info()
        if not user:
            messagebox.showwarning("No Selection", "Please select a user")
            return
        
        if not self.groups_data:
            self.status.config(text="Loading groups...", foreground="orange")
            self.fetch_groups()
            self.root.after(1500, lambda: self._show_group_selector(user, "remove"))
            return
        
        self._show_group_selector(user, "remove")
    
    def _show_group_selector(self, user, action):
        """Show dialog to select a group for add/remove action"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"{'Add to' if action == 'add' else 'Remove from'} Group")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text=f"User: {user['name']} ({user['email']})", font=("", 10, "bold")).pack(pady=10)
        ttk.Label(dialog, text=f"Select a group to {'add user to' if action == 'add' else 'remove user from'}:").pack(pady=5)
        
        # Search box
        search_var = tk.StringVar()
        search_frame = ttk.Frame(dialog)
        search_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(search_frame, text="Search:").pack(side="left")
        search_entry = ttk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        # Listbox with scrollbar
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Populate groups
        group_names = [g["name"] for g in self.groups_data]
        group_names.sort()
        
        def update_list(*args):
            term = search_var.get().lower()
            listbox.delete(0, tk.END)
            for name in group_names:
                if term in name.lower():
                    listbox.insert(tk.END, name)
        
        search_var.trace("w", update_list)
        update_list()
        
        # Buttons
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill="x", padx=10, pady=10)
        
        def on_ok():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select a group")
                return
            
            group_name = listbox.get(selection[0])
            dialog.destroy()
            
            if action == "add":
                threading.Thread(target=self._add_user_to_group_thread, args=(user, group_name), daemon=True).start()
            else:
                threading.Thread(target=self._remove_user_from_group_thread, args=(user, group_name), daemon=True).start()
        
        ttk.Button(btn_frame, text="OK", command=on_ok, width=15).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy, width=15).pack(side="left")
        
        # Bind double-click
        listbox.bind("<Double-Button-1>", lambda e: on_ok())
    
    def _add_user_to_group_thread(self, user, group_name):
        """Thread worker for adding user to group"""
        try:
            self.root.after(0, lambda: self.status.config(text=f"Adding {user['name']} to {group_name}...", foreground="orange"))
            
            response = requests.post(
                f"{self.jira_url.get().rstrip('/')}/rest/api/3/group/user",
                params={"groupname": group_name},
                json={"accountId": user['account_id']},
                auth=self.auth(),
                headers={
                    "Accept": "application/json",
                    "Content-Type": "application/json"
                },
                timeout=30
            )
            
            if response.status_code in [200, 201]:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Success",
                    f"User {user['name']} added to group '{group_name}' successfully."
                ))
                self.root.after(0, lambda: self.status.config(text="User added to group", foreground="green"))
            else:
                raise Exception(f"API returned status {response.status_code}: {response.text}")
                
        except Exception as e:
            error_msg = f"Failed to add user to group: {str(e)}"
            print(error_msg)
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            self.root.after(0, lambda: self.status.config(text="Failed to add to group", foreground="red"))
    
    def _remove_user_from_group_thread(self, user, group_name):
        """Thread worker for removing user from group"""
        try:
            self.root.after(0, lambda: self.status.config(text=f"Removing {user['name']} from {group_name}...", foreground="orange"))
            
            response = requests.delete(
                f"{self.jira_url.get().rstrip('/')}/rest/api/3/group/user",
                params={
                    "groupname": group_name,
                    "accountId": user['account_id']
                },
                auth=self.auth(),
                headers={"Accept": "application/json"},
                timeout=30
            )
            
            if response.status_code in [200, 204]:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Success",
                    f"User {user['name']} removed from group '{group_name}' successfully."
                ))
                self.root.after(0, lambda: self.status.config(text="User removed from group", foreground="green"))
            else:
                raise Exception(f"API returned status {response.status_code}: {response.text}")
                
        except Exception as e:
            error_msg = f"Failed to remove user from group: {str(e)}"
            print(error_msg)
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            self.root.after(0, lambda: self.status.config(text="Failed to remove from group", foreground="red"))
    
    def manage_product_access(self):
        """Manage product access for a user"""
        user = self.get_selected_user_info()
        if not user:
            messagebox.showwarning("No Selection", "Please select a user")
            return
        
        messagebox.showinfo(
            "Product Access Management",
            "Product access management is done through the Atlassian Admin portal.\n\n"
            f"Opening the user profile for {user['name']}...\n\n"
            "You can manage product access in the 'Products' section of the user profile."
        )
        
        self.open_user_profile()
    
    # ---------------- Bulk Actions ---------------- #
    def show_bulk_edit_dialog(self):
        """Show bulk edit dialog for selected users"""
        if self.current_view != "users":
            messagebox.showinfo("Info", "Bulk edit is only available in Users view.")
            return
        
        if not self.selected_items:
            messagebox.showwarning("No Selection", "Please select users from the main view first.")
            return
        
        # Get selected user data
        selected_users = []
        for item in self.selected_items:
            values = self.tree.item(item, "values")
            if len(values) >= 4:
                selected_users.append({
                    "name": values[1],
                    "email": values[2],
                    "account_id": values[3],
                    "status": values[5] if len(values) > 5 else "Unknown"
                })
        
        if not selected_users:
            messagebox.showwarning("Error", "Could not retrieve user data for selected items.")
            return
        
        # Create dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Bulk Edit Users")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Header
        header_frame = ttk.Frame(dialog)
        header_frame.pack(fill="x", padx=20, pady=15)
        
        ttk.Label(
            header_frame, 
            text="⚡ Bulk Edit Users", 
            font=("", 16, "bold")
        ).pack()
        
        ttk.Label(
            header_frame,
            text=f"{len(selected_users)} user(s) selected",
            font=("", 10),
            foreground="gray"
        ).pack()
        
        # Selected users preview
        preview_frame = ttk.LabelFrame(dialog, text="📋 Selected Users", padding=10)
        preview_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        
        # Create scrollable list of selected users
        preview_list_frame = ttk.Frame(preview_frame)
        preview_list_frame.pack(fill="both", expand=True)
        
        preview_scrollbar = ttk.Scrollbar(preview_list_frame)
        preview_scrollbar.pack(side="right", fill="y")
        
        preview_listbox = tk.Listbox(
            preview_list_frame, 
            yscrollcommand=preview_scrollbar.set,
            height=8
        )
        preview_listbox.pack(side="left", fill="both", expand=True)
        preview_scrollbar.config(command=preview_listbox.yview)
        
        for user in selected_users:
            preview_listbox.insert(tk.END, f"• {user['name']} ({user['email']})")
        
        # Action selection
        action_frame = ttk.LabelFrame(dialog, text="🎯 Select Action", padding=15)
        action_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        action_var = tk.StringVar(value="deactivate")
        
        actions = [
            ("deactivate", "🚫 Deactivate Users", "Revoke access to Jira (requires Org API)"),
            ("reactivate", "✅ Reactivate Users", "Restore access to Jira (requires Org API)"),
            ("add_group", "➕ Add to Group", "Add users to a selected group"),
            ("remove_group", "➖ Remove from Group", "Remove users from a selected group"),
        ]
        
        for value, label, description in actions:
            frame = ttk.Frame(action_frame)
            frame.pack(fill="x", pady=2)
            
            ttk.Radiobutton(
                frame,
                text=label,
                variable=action_var,
                value=value
            ).pack(side="left")
            
            ttk.Label(
                frame,
                text=f"  ({description})",
                font=("", 8),
                foreground="gray"
            ).pack(side="left")
        
        # Warning message
        warning_frame = ttk.Frame(dialog)
        warning_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        warning_label = ttk.Label(
            warning_frame,
            text="⚠️ Warning: Bulk actions cannot be undone. Please verify your selection.",
            font=("", 9),
            foreground="red"
        )
        warning_label.pack()
        
        # Buttons
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        def execute_action():
            action = action_var.get()
            count = len(selected_users)
            
            # Close the bulk edit dialog
            dialog.destroy()
            
            # For group actions, show group selector first (which handles its own confirmation)
            if action in ["add_group", "remove_group"]:
                self._bulk_group_action(selected_users, action)
            else:
                # For direct actions (deactivate/reactivate), confirm first then execute
                action_names = {
                    "deactivate": "Deactivate",
                    "reactivate": "Reactivate"
                }
                
                confirm = messagebox.askyesno(
                    "Confirm Bulk Action",
                    f"Execute '{action_names.get(action, action)}' on {count} user(s)?\n\n"
                    f"⚠️ This action cannot be undone!"
                )
                
                if not confirm:
                    return
                
                # Execute direct action
                threading.Thread(
                    target=self._execute_bulk_action_thread, 
                    args=(selected_users, action), 
                    daemon=True
                ).start()
        
        def cancel_action():
            dialog.destroy()
        
        ttk.Button(
            btn_frame, 
            text="▶ Execute", 
            command=execute_action, 
            width=20
        ).pack(side="left", padx=5)
        
        ttk.Button(
            btn_frame, 
            text="✖ Cancel", 
            command=cancel_action, 
            width=20
        ).pack(side="right", padx=5)
    
    
    def _bulk_group_action(self, users, action):
        """Handle bulk add/remove from group - requires group selection"""
        if not self.groups_data:
            # Auto-fetch groups if not loaded
            self.status.config(text="Loading groups...", foreground="orange")
            self.fetch_groups()
            # Wait for groups to load, then retry
            self.root.after(1500, lambda: self._bulk_group_action(users, action))
            return
        
        # Show enhanced group selector dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Group for Bulk Action")
        dialog.geometry("500x450")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Header
        header_frame = ttk.Frame(dialog)
        header_frame.pack(fill="x", padx=20, pady=15)
        
        action_text = "Add to Group" if action == "add_group" else "Remove from Group"
        ttk.Label(
            header_frame, 
            text=f"👥 {action_text}", 
            font=("", 14, "bold")
        ).pack()
        
        ttk.Label(
            header_frame,
            text=f"Select which group to {action_text.lower()} for {len(users)} user(s)",
            font=("", 9),
            foreground="gray"
        ).pack()
        
        # Search box
        search_frame = ttk.LabelFrame(dialog, text="🔍 Search Groups", padding=10)
        search_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, font=("", 10))
        search_entry.pack(fill="x")
        search_entry.focus()
        
        # Group list
        list_frame = ttk.LabelFrame(dialog, text="📋 Available Groups", padding=10)
        list_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        
        # Listbox with scrollbar
        listbox_frame = ttk.Frame(list_frame)
        listbox_frame.pack(fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(listbox_frame)
        scrollbar.pack(side="right", fill="y")
        
        listbox = tk.Listbox(
            listbox_frame, 
            yscrollcommand=scrollbar.set,
            font=("", 10),
            activestyle="dotbox"
        )
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Store all group names sorted
        all_groups = sorted([g["name"] for g in self.groups_data])
        
        def update_list(*args):
            """Update listbox based on search term"""
            term = search_var.get().lower()
            listbox.delete(0, tk.END)
            
            for group_name in all_groups:
                if term in group_name.lower():
                    listbox.insert(tk.END, group_name)
            
            # Show count
            visible_count = listbox.size()
            list_frame.config(text=f"📋 Available Groups ({visible_count} shown)")
        
        # Bind search
        search_var.trace("w", update_list)
        update_list()  # Initial population
        
        # Info label
        info_label = ttk.Label(
            dialog,
            text="💡 Tip: Type to search, double-click or press Enter to select",
            font=("", 8),
            foreground="gray"
        )
        info_label.pack(padx=20, pady=(0, 10))
        
        # Buttons
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        def on_ok():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select a group")
                return
            
            group_name = listbox.get(selection[0])
            
            # Confirm the action
            confirm = messagebox.askyesno(
                "Confirm Bulk Group Action",
                f"{action_text} '{group_name}' for {len(users)} user(s)?\n\n"
                f"Users:\n" + "\n".join([f"  • {u['name']}" for u in users[:5]]) +
                (f"\n  ... and {len(users) - 5} more" if len(users) > 5 else "")
            )
            
            if not confirm:
                return
            
            dialog.destroy()
            
            # Execute the bulk action
            threading.Thread(
                target=self._execute_bulk_action_thread, 
                args=(users, action, group_name), 
                daemon=True
            ).start()
        
        def on_cancel():
            dialog.destroy()
        
        ttk.Button(
            btn_frame, 
            text="✓ Select Group", 
            command=on_ok, 
            width=20
        ).pack(side="left", padx=5)
        
        ttk.Button(
            btn_frame, 
            text="✖ Cancel", 
            command=on_cancel, 
            width=20
        ).pack(side="right", padx=5)
        
        # Bind double-click and Enter key
        listbox.bind("<Double-Button-1>", lambda e: on_ok())
        listbox.bind("<Return>", lambda e: on_ok())
        
        # Bind Escape to cancel
        dialog.bind("<Escape>", lambda e: on_cancel())
    
    def _execute_bulk_action_thread(self, users, action, group_name=None):
        """Execute bulk action on multiple users"""
        total = len(users)
        success_count = 0
        fail_count = 0
        
        self.root.after(0, lambda: self.status.config(text=f"Processing bulk action on {total} user(s)...", foreground="orange"))
        
        for i, user in enumerate(users, 1):
            try:
                self.root.after(0, lambda u=user, idx=i: self.status.config(
                    text=f"Processing {idx}/{total}: {u['name']}...", foreground="orange"
                ))
                
                if action == "deactivate":
                    url = f"https://api.atlassian.com/users/{user['account_id']}/manage/lifecycle/disable"
                    response = requests.post(url, headers={"Authorization": f"Bearer {self.org_api_key.get()}"}, timeout=30)
                
                elif action == "reactivate":
                    url = f"https://api.atlassian.com/users/{user['account_id']}/manage/lifecycle/enable"
                    response = requests.post(url, headers={"Authorization": f"Bearer {self.org_api_key.get()}"}, timeout=30)
                
                elif action == "add_group":
                    response = requests.post(
                        f"{self.jira_url.get().rstrip('/')}/rest/api/3/group/user",
                        params={"groupname": group_name},
                        json={"accountId": user['account_id']},
                        auth=self.auth(),
                        headers={"Accept": "application/json", "Content-Type": "application/json"},
                        timeout=30
                    )
                
                elif action == "remove_group":
                    response = requests.delete(
                        f"{self.jira_url.get().rstrip('/')}/rest/api/3/group/user",
                        params={"groupname": group_name, "accountId": user['account_id']},
                        auth=self.auth(),
                        headers={"Accept": "application/json"},
                        timeout=30
                    )
                
                if response.status_code in [200, 201, 204]:
                    success_count += 1
                else:
                    fail_count += 1
                    print(f"Failed for {user['name']}: {response.status_code} - {response.text}")
                
                # Small delay to avoid rate limiting
                time.sleep(0.5)
                
            except Exception as e:
                fail_count += 1
                print(f"Error processing {user['name']}: {str(e)}")
        
        # Show results
        result_msg = f"Bulk action completed:\n\n✓ Success: {success_count}\n✗ Failed: {fail_count}"
        self.root.after(0, lambda: messagebox.showinfo("Bulk Action Complete", result_msg))
        self.root.after(0, lambda: self.status.config(text=f"Bulk action complete: {success_count} success, {fail_count} failed", foreground="green"))
        
        # Refresh user list
        if action in ["deactivate", "reactivate"]:
            self.root.after(0, self.fetch_users_async)

    # ---------------- Export ---------------- #
    def export_csv(self):
        if self.current_view == "users":
            if not self.users_data:
                messagebox.showwarning("Warning", "No users to export.")
                return
            filename = f"jira_users_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            with open(filename, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Display Name", "Email", "Account ID", "Account Type", "Status", "Last Active"])
                
                if self.use_org_api.get():
                    for u in self.users_data:
                        last_active = u.get("last_active", "")
                        if last_active:
                            try:
                                dt = parser.isoparse(last_active)
                                last_active = dt.strftime("%Y-%m-%d %H:%M:%S")
                            except:
                                pass
                        writer.writerow([
                            u.get("name", ""),
                            u.get("email", ""),
                            u.get("account_id", ""),
                            u.get("account_type", ""),
                            u.get("account_status", ""),
                            last_active
                        ])
                else:
                    for u in self.users_data:
                        writer.writerow([
                            u.get("displayName", ""),
                            u.get("emailAddress", ""),
                            u.get("accountId", ""),
                            u.get("accountType", ""),
                            "Active" if u.get("active") else "Inactive",
                            "N/A (use Org API)"
                        ])
            messagebox.showinfo("Exported", f"Users exported to {filename}")

        elif self.current_view == "groups":
            if not self.groups_data:
                messagebox.showwarning("Warning", "No groups to export.")
                return
            filename = f"jira_groups_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            with open(filename, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Group Name", "Group ID", "Member Count", "Member Name", "Member Email", "Member ID", "Member Type", "Member Status", "Last Active"])
                for g in self.groups_data:
                    group_name = g.get("name", "")
                    group_id = g.get("groupId", "")
                    members = self.groups_members.get(group_name, [])
                    member_count = len(members)
                    if members:
                        for m in members:
                            writer.writerow([
                                group_name,
                                group_id,
                                member_count,
                                m.get("displayName", ""),
                                m.get("emailAddress", ""),
                                m.get("accountId", ""),
                                m.get("accountType", ""),
                                "Active" if m.get("active") else "Inactive",
                                "N/A"
                            ])
                    else:
                        writer.writerow([group_name, group_id, member_count, "", "", "", "", "", ""])
            messagebox.showinfo("Exported", f"Groups exported to {filename}")
    
    def export_groups_csv(self):
        """Export groups data to CSV"""
        if not self.groups_data:
            messagebox.showwarning("Warning", "No groups to export.")
            return
        filename = f"jira_groups_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        with open(filename, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Group Name", "Group ID", "Member Count"])
            for g in self.groups_data:
                writer.writerow([
                    g.get("name", ""),
                    g.get("groupId", ""),
                    g.get("memberCount", "")
                ])
        messagebox.showinfo("Exported", f"Groups exported to {filename}")

# ---------------- START ---------------- #
if __name__ == "__main__":
    root = tk.Tk()
    JiraUserApp(root)
    root.mainloop()
