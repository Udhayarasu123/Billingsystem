import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog, scrolledtext
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from num2words import num2words
from datetime import datetime
import os
import platform
import subprocess
import sys
import json
from tkinter import font as tkfont
import webbrowser
import pandas as pd
from PIL import Image, ImageTk
import sqlite3
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import qrcode
import threading

class BillingSystem:
    invoice_count = 0
    CONFIG_FILE = "billing_config.json"
    DB_FILE = "billing_database.db"
    THEMES = {
        "Default": {
            "primary": "#2c3e50",
            "secondary": "#3498db",
            "accent": "#e74c3c",
            "bg": "#f5f5f5",
            "text": "#333333"
        },
        "Dark": {
            "primary": "#1a1a1a",
            "secondary": "#4d4d4d",
            "accent": "#cc3300",
            "bg": "#333333",
            "text": "#ffffff"
        },
        "Ocean": {
            "primary": "#006994",
            "secondary": "#008080",
            "accent": "#ff6b6b",
            "bg": "#e6f7ff",
            "text": "#003366"
        },
        "Forest": {
            "primary": "#2e7d32",
            "secondary": "#689f38",
            "accent": "#d32f2f",
            "bg": "#f1f8e9",
            "text": "#1b5e20"
        }
    }

    def __init__(self, master):
        self.master = master
        self.master.title("Advanced Billing System")
        self.master.geometry("2400x1200")
        self.master.minsize(1200, 800)
        
        # Initialize database
        self.init_database()
        
        # Load configuration
        self.load_config()
        
        # Initialize variables
        self.invoice_number = self.get_last_invoice_number() + 1
        self.date = datetime.now().strftime("%d-%m-%Y")
        self.products = []  # For product history
        self.current_theme = "Default"
        
        # Setup UI
        self.setup_ui()
        self.apply_styles()
        self.calculate_totals()
        
        # Bind keyboard shortcuts
        self.setup_shortcuts()
        
        # Load product history if exists
        self.load_product_history()
        
        # Setup idle timer for auto-save
        self.setup_auto_save()

    def init_database(self):
        """Initialize SQLite database"""
        self.conn = sqlite3.connect(self.DB_FILE)
        self.cursor = self.conn.cursor()
        
        # Create tables if they don't exist
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS invoices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_number INTEGER,
                date TEXT,
                customer_name TEXT,
                customer_mobile TEXT,
                customer_place TEXT,
                customer_address TEXT,
                bill_type TEXT,
                subtotal REAL,
                sgst REAL,
                igst REAL,
                roundoff REAL,
                total REAL,
                pdf_path TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS invoice_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_id INTEGER,
                sno INTEGER,
                hsn TEXT,
                description TEXT,
                price REAL,
                quantity INTEGER,
                total REAL,
                FOREIGN KEY (invoice_id) REFERENCES invoices (id)
            )
        ''')
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                hsn TEXT UNIQUE,
                name TEXT,
                price REAL,
                category TEXT,
                last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS customers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                mobile TEXT UNIQUE,
                place TEXT,
                address TEXT,
                gstin TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        self.conn.commit()

    def get_last_invoice_number(self):
        """Get the last invoice number from database"""
        self.cursor.execute("SELECT MAX(invoice_number) FROM invoices")
        result = self.cursor.fetchone()
        return result[0] if result[0] is not None else 0

    def load_config(self):
        """Load configuration from file or use defaults"""
        defaults = {
              "primary_color": "#2c3e50",
            "secondary_color": "#3498db",
            "accent_color": "#e74c3c",
            "font_family": "Segoe UI",
            "font_size": 12,
            "company_name": "Sri Vetri Vinayaga Traders",
            "company_address": "Sengundhar Mahal, Thiruvika Nagar, Chinnasalem, Kallakurichi Dt-606201",
            "company_phone": "9080013157, 9942191481",
            "company_email": "vetrivinayagatraders@gmail.com",
            "company_website": "www.vetrivinayagatraders.com",
            "gstin": "33CUPPM4345DIZM",
            "bank_details": {
                "name": "City Union Bank",
                "account": "500101011688022",
                "ifsc": "CIUB0000561",
                "branch": "Chinnasalem"
            },
            "tax_rates": {
                "sgst": 9,
                "igst": 9
            },
            "auto_save": True,
            "auto_save_interval": 5,  # minutes
            "default_theme": "Default"
        }
        
        try:
            if os.path.exists(self.CONFIG_FILE):
                with open(self.CONFIG_FILE, 'r') as f:
                    self.config = json.load(f)
                # Merge with defaults for any missing keys
                for key, value in defaults.items():
                    if key not in self.config:
                        self.config[key] = value
            else:
                self.config = defaults
                
            # Set current theme
            self.current_theme = self.config.get("default_theme", "Default")
        except Exception as e:
            print(f"Error loading config: {e}")
            self.config = defaults

    def save_config(self):
        """Save configuration to file"""
        try:
            with open(self.CONFIG_FILE, 'w') as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            print(f"Error saving config: {e}")

    def setup_ui(self):
        """Setup the main user interface"""
        # Main container
        self.main_container = ttk.Frame(self.master)
        self.main_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Menu bar
        self.setup_menu()
        
        # Status bar
        self.setup_status_bar()
        
        # Notebook for multiple tabs
        self.notebook = ttk.Notebook(self.main_container)
        self.notebook.pack(fill="both", expand=True)
        
        # Invoice tab
        self.invoice_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.invoice_tab, text="New Invoice")
        
        # Reports tab
        self.reports_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.reports_tab, text="Reports", state="hidden")  # Initially hidden
        
        # Products tab
        self.products_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.products_tab, text="Products", state="hidden")  # Initially hidden
        
        # Customers tab
        self.customers_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.customers_tab, text="Customers", state="hidden")  # Initially hidden
        
        # Setup tabs
        self.setup_invoice_tab()
        self.setup_reports_tab()
        self.setup_products_tab()
        self.setup_customers_tab()
        
        # Bind notebook tab change event
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

    def setup_status_bar(self):
        """Setup the status bar"""
        self.status_bar = ttk.Frame(self.main_container)
        self.status_bar.pack(fill="x", pady=(0, 5))
        
        self.status_label = ttk.Label(
            self.status_bar, 
            text="Ready", 
            relief="sunken", 
            anchor="w"
        )
        self.status_label.pack(fill="x")
        
        self.auto_save_status = ttk.Label(
            self.status_bar, 
            text="Auto-save: ON" if self.config["auto_save"] else "Auto-save: OFF",
            relief="sunken",
            anchor="e",
            width=15
        )
        self.auto_save_status.pack(side="right")

    def setup_menu(self):
        """Setup the menu bar"""
        menubar = tk.Menu(self.master)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Invoice", command=self.new_invoice, accelerator="Ctrl+N")
        file_menu.add_command(label="Save Invoice", command=self.save_bill, accelerator="Ctrl+S")
        file_menu.add_command(label="Print Invoice", command=self.print_bill, accelerator="Ctrl+P")
        file_menu.add_command(label="Export to Excel", command=self.export_to_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_exit)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Clear Selected", command=self.clear_selected, accelerator="Del")
        edit_menu.add_command(label="Clear All", command=self.clear_all, accelerator="Ctrl+Del")
        edit_menu.add_separator()
        edit_menu.add_command(label="Find Invoice", command=self.find_invoice)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        
        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        
        # Theme submenu
        theme_menu = tk.Menu(view_menu, tearoff=0)
        for theme_name in self.THEMES:
            theme_menu.add_radiobutton(
                label=theme_name,
                command=lambda name=theme_name: self.change_theme(name),
                variable=tk.StringVar(value=self.current_theme),
                value=theme_name
            )
        
        view_menu.add_cascade(label="Theme", menu=theme_menu)
        view_menu.add_checkbutton(
            label="Auto-save", 
            variable=tk.BooleanVar(value=self.config["auto_save"]),
            command=self.toggle_auto_save
        )
        menubar.add_cascade(label="View", menu=view_menu)
        
        # Settings menu
        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="Appearance", command=self.appearance_settings)
        settings_menu.add_command(label="Company Info", command=self.company_settings)
        settings_menu.add_command(label="Tax Rates", command=self.tax_settings)
        settings_menu.add_command(label="Database", command=self.database_settings)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="User Guide", command=self.show_user_guide)
        help_menu.add_command(label="Check for Updates", command=self.check_for_updates)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        self.master.config(menu=menubar)

    def setup_invoice_tab(self):
        """Setup the invoice tab"""
        # Scrollable frame
        self.setup_scrollable_frame(self.invoice_tab)
        
        # Header section
        self.setup_header()
        
        # Customer information
        self.setup_customer_info()
        
        # Product entry
        self.setup_product_entry()
        
        # Product table
        self.setup_product_table()
        
        # Totals section
        self.setup_totals()
        
        # Footer buttons
        self.setup_footer()

    def setup_reports_tab(self):
        """Setup the reports tab"""
        # Notebook for different report types
        reports_notebook = ttk.Notebook(self.reports_tab)
        reports_notebook.pack(fill="both", expand=True)
        
        # Sales report
        sales_frame = ttk.Frame(reports_notebook)
        reports_notebook.add(sales_frame, text="Sales Report")
        
        # Date range selection
        date_frame = ttk.Frame(sales_frame)
        date_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(date_frame, text="From:").pack(side="left", padx=5)
        self.from_date = ttk.Entry(date_frame)
        self.from_date.pack(side="left", padx=5)
        
        ttk.Label(date_frame, text="To:").pack(side="left", padx=5)
        self.to_date = ttk.Entry(date_frame)
        self.to_date.pack(side="left", padx=5)
        
        ttk.Button(
            date_frame, 
            text="Generate Report", 
            command=self.generate_sales_report
        ).pack(side="right", padx=5)
        
        # Report display area
        self.report_text = scrolledtext.ScrolledText(
            sales_frame,
            wrap=tk.WORD,
            width=100,
            height=30
        )
        self.report_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Sales chart frame
        chart_frame = ttk.Frame(reports_notebook)
        reports_notebook.add(chart_frame, text="Sales Chart")
        
        self.sales_figure = plt.Figure(figsize=(6, 4), dpi=100)
        self.sales_ax = self.sales_figure.add_subplot(111)
        self.sales_canvas = FigureCanvasTkAgg(self.sales_figure, chart_frame)
        self.sales_canvas.get_tk_widget().pack(fill="both", expand=True)
        
        # Product report
        product_frame = ttk.Frame(reports_notebook)
        reports_notebook.add(product_frame, text="Product Report")
        
        ttk.Button(
            product_frame, 
            text="Generate Product Report", 
            command=self.generate_product_report
        ).pack(pady=5)
        
        self.product_report_text = scrolledtext.ScrolledText(
            product_frame,
            wrap=tk.WORD,
            width=100,
            height=30
        )
        self.product_report_text.pack(fill="both", expand=True, padx=5, pady=5)

    def setup_products_tab(self):
        """Setup the products tab"""
        # Top frame for controls
        control_frame = ttk.Frame(self.products_tab)
        control_frame.pack(fill="x", padx=5, pady=5)
        
        # Search box
        ttk.Label(control_frame, text="Search:").pack(side="left", padx=5)
        self.product_search = ttk.Entry(control_frame)
        self.product_search.pack(side="left", padx=5)
        self.product_search.bind("<KeyRelease>", self.search_products_in_db)
        
        # Buttons
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(side="right")
        
        ttk.Button(
            button_frame,
            text="Add Product",
            command=self.add_product_dialog
        ).pack(side="left", padx=2)
        
        ttk.Button(
            button_frame,
            text="Edit Product",
            command=self.edit_product_dialog
        ).pack(side="left", padx=2)
        
        ttk.Button(
            button_frame,
            text="Delete Product",
            command=self.delete_product
        ).pack(side="left", padx=2)
        
        ttk.Button(
            button_frame,
            text="Refresh",
            command=self.load_products_table
        ).pack(side="left", padx=2)
        
        # Products table
        self.products_table = ttk.Treeview(
            self.products_tab,
            columns=("ID", "HSN", "Name", "Price", "Category", "Last Updated"),
            show="headings"
        )
        
        # Configure columns
        self.products_table.heading("ID", text="ID", anchor="center")
        self.products_table.heading("HSN", text="HSN", anchor="center")
        self.products_table.heading("Name", text="Name", anchor="center")
        self.products_table.heading("Price", text="Price", anchor="center")
        self.products_table.heading("Category", text="Category", anchor="center")
        self.products_table.heading("Last Updated", text="Last Updated", anchor="center")
        
        self.products_table.column("ID", width=50, anchor="center")
        self.products_table.column("HSN", width=100, anchor="center")
        self.products_table.column("Name", width=250, anchor="w")
        self.products_table.column("Price", width=80, anchor="e")
        self.products_table.column("Category", width=100, anchor="center")
        self.products_table.column("Last Updated", width=120, anchor="center")
        
        # Add scrollbars
        y_scroll = ttk.Scrollbar(self.products_tab, orient="vertical", command=self.products_table.yview)
        x_scroll = ttk.Scrollbar(self.products_tab, orient="horizontal", command=self.products_table.xview)
        self.products_table.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        
        # Grid layout
        self.products_table.pack(side="left", fill="both", expand=True)
        y_scroll.pack(side="right", fill="y")
        x_scroll.pack(side="bottom", fill="x")
        
        # Load initial data
        self.load_products_table()

    def setup_customers_tab(self):
        """Setup the customers tab"""
        # Top frame for controls
        control_frame = ttk.Frame(self.customers_tab)
        control_frame.pack(fill="x", padx=10, pady=10)
        
        # Search box
        ttk.Label(control_frame, text="Search:").pack(side="left", padx=5)
        self.customer_search = ttk.Entry(control_frame)
        self.customer_search.pack(side="left", padx=5)
        self.customer_search.bind("<KeyRelease>", self.search_customers_in_db)
        
        # Buttons
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(side="right")
        
        ttk.Button(
            button_frame,
            text="Add Customer",
            command=self.add_customer_dialog
        ).pack(side="left", padx=2)
        
        ttk.Button(
            button_frame,
            text="Edit Customer",
            command=self.edit_customer_dialog
        ).pack(side="left", padx=2)
        
        ttk.Button(
            button_frame,
            text="Delete Customer",
            command=self.delete_customer
        ).pack(side="left", padx=2)
        
        ttk.Button(
            button_frame,
            text="Refresh",
            command=self.load_customers_table
        ).pack(side="left", padx=2)
        
        # Customers table
        self.customers_table = ttk.Treeview(
            self.customers_tab,
            columns=("ID", "Name", "Mobile", "Place", "Address", "GSTIN", "Created At"),
            show="headings"
        )
        
        # Configure columns
        self.customers_table.heading("ID", text="ID", anchor="center")
        self.customers_table.heading("Name", text="Name", anchor="center")
        self.customers_table.heading("Mobile", text="Mobile", anchor="center")
        self.customers_table.heading("Place", text="Place", anchor="center")
        self.customers_table.heading("Address", text="Address", anchor="center")
        self.customers_table.heading("GSTIN", text="GSTIN", anchor="center")
        self.customers_table.heading("Created At", text="Created At", anchor="center")
        
        self.customers_table.column("ID", width=50, anchor="center")
        self.customers_table.column("Name", width=150, anchor="w")
        self.customers_table.column("Mobile", width=100, anchor="center")
        self.customers_table.column("Place", width=100, anchor="center")
        self.customers_table.column("Address", width=200, anchor="w")
        self.customers_table.column("GSTIN", width=120, anchor="center")
        self.customers_table.column("Created At", width=120, anchor="center")
        
        # Add scrollbars
        y_scroll = ttk.Scrollbar(self.customers_tab, orient="vertical", command=self.customers_table.yview)
        x_scroll = ttk.Scrollbar(self.customers_tab, orient="horizontal", command=self.customers_table.xview)
        self.customers_table.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        
        # Grid layout
        self.customers_table.pack(side="left", fill="both", expand=True)
        y_scroll.pack(side="right", fill="y")
        x_scroll.pack(side="bottom", fill="x")
        
        # Load initial data
        self.load_customers_table()

    def setup_scrollable_frame(self, parent):
        """Setup the scrollable frame"""
        self.scrollable_frame = ttk.Frame(parent)
        self.scrollable_frame.pack(fill="both", expand=True)
        
        self.canvas = tk.Canvas(self.scrollable_frame, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.scrollable_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame_inner = ttk.Frame(self.canvas)
        
        self.scrollable_frame_inner.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame_inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack( fill="both", expand=True)
        
        # Bind mousewheel to scroll
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        """Handle mousewheel scrolling"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def setup_header(self):
        """Setup the header section"""
        header_frame = ttk.Frame(self.scrollable_frame_inner, style="Header.TFrame")
        header_frame.pack(fill="x", pady=(0, 10))
        
        # Company logo and name
        logo_frame = ttk.Frame(header_frame)
        logo_frame.pack(side="left", padx=100)
        
        try:
            # Try to load logo if exists
            if getattr(sys, 'frozen', False):
                logo_path = os.path.join(sys._MEIPASS, "logo.png")
            else:
                logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
                
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                img = img.resize((100, 100), Image.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(img)
                self.logo_label = ttk.Label(logo_frame, image=self.logo_img)
                self.logo_label.pack()
        except Exception as e:
            print(f"Error loading logo: {e}")
        
        # Company info
        company_frame = ttk.Frame(header_frame)
        company_frame.pack(side="left", expand=True, fill="x")
        
        self.company_name_label = ttk.Label(
            company_frame, 
            text=self.config["company_name"], 
            style="CompanyName.TLabel"
        )
        self.company_name_label.pack(anchor="center")
        
        self.address_label_1 = ttk.Label(
            company_frame, 
            text=self.config["company_address"],
            style="CompanyAddress.TLabel"
        )
        self.address_label_1.pack()
        
        self.address_label_2 = ttk.Label(
            company_frame, 
            text=f"Phone: {self.config['company_phone']} | Email: {self.config.get('company_email', '')}",
            style="CompanyAddress.TLabel"
        )
        self.address_label_2.pack()
        
        self.gstin_label = ttk.Label(
            company_frame, 
            text=f"GSTIN: {self.config['gstin']}",
            style="CompanyAddress.TLabel"
        )
        self.gstin_label.pack(anchor="w")
        
        # Invoice info
        invoice_frame = ttk.Frame(header_frame)
        invoice_frame.pack(side="right", padx=10)
        
        self.bill_type_var = tk.StringVar(value="Cash Bill")
        
        bill_type_frame = ttk.Frame(invoice_frame)
        bill_type_frame.pack(fill="x", pady=5)
        
        ttk.Radiobutton(
            bill_type_frame, 
            text="Cash Bill", 
            variable=self.bill_type_var, 
            value="Cash Bill",
            style="BillType.TRadiobutton"
        ).pack(side="left", padx=5)
        
        ttk.Radiobutton(
            bill_type_frame, 
            text="Credit Bill", 
            variable=self.bill_type_var, 
            value="Credit Bill",
            style="BillType.TRadiobutton"
        ).pack(side="left", padx=5)
        
        self.invoice_frame = ttk.Frame(invoice_frame)
        self.invoice_frame.pack(fill="x", pady=5)
        
        self.invoice_label = ttk.Label(
            self.invoice_frame, 
            text=f"Invoice No: {self.invoice_number:04d}",
            style="InvoiceInfo.TLabel"
        )
        self.invoice_label.pack(side="left")
        
        self.date_label = ttk.Label(
            self.invoice_frame, 
            text=f"Date: {self.date}",
            style="InvoiceInfo.TLabel"
        )
        self.date_label.pack(side="right")

    def setup_customer_info(self):
        """Setup customer information section"""
        customer_frame = ttk.LabelFrame(
            self.scrollable_frame_inner, 
            text="Customer Information", 
            style="Section.TLabelframe"
        )
        customer_frame.pack(fill="x", pady=10, padx=5)
        
        # Name
        ttk.Label(
            customer_frame, 
            text="Customer Name:", 
            style="Bold.TLabel"
        ).grid(row=0, column=0, padx=5, pady=5, sticky="e")
        
        self.name_entry = ttk.Entry(
            customer_frame, 
            font=(self.config["font_family"], self.config["font_size"])
        )
        self.name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="we")
        
        # Mobile
        ttk.Label(
            customer_frame, 
            text="Mobile Number:", 
            style="Bold.TLabel"
        ).grid(row=0, column=2, padx=5, pady=5, sticky="e")
        
        self.mobile_entry = ttk.Entry(
            customer_frame, 
            font=(self.config["font_family"], self.config["font_size"])
        )
        self.mobile_entry.grid(row=0, column=3, padx=5, pady=5, sticky="we")
        
        # Place
        ttk.Label(
            customer_frame, 
            text="Place:", 
            style="Bold.TLabel"
        ).grid(row=0, column=4, padx=5, pady=5, sticky="e")
        
        self.place_entry = ttk.Entry(
            customer_frame, 
            font=(self.config["font_family"], self.config["font_size"])
        )
        self.place_entry.grid(row=0, column=5, padx=5, pady=5, sticky="we")
        
        # Address (full width)
        ttk.Label(
            customer_frame, 
            text="Address:", 
            style="Bold.TLabel"
        ).grid(row=1, column=0, padx=5, pady=5, sticky="e")
        
        self.address_entry = ttk.Entry(
            customer_frame, 
            font=(self.config["font_family"], self.config["font_size"])
        )
        self.address_entry.grid(
            row=1, column=1, 
            columnspan=5, 
            padx=5, pady=5, 
            sticky="we"
        )
        
        # Configure grid weights
        for i in range(6):
            customer_frame.grid_columnconfigure(i, weight=1 if i in (1, 3, 5) else 0)

    def setup_product_entry(self):
        """Setup product entry section"""
        product_frame = ttk.LabelFrame(
            self.scrollable_frame_inner, 
            text="Add Product", 
            style="Section.TLabelframe"
        )
        product_frame.pack(fill="x", pady=10, padx=5)
        
        # HSN
        ttk.Label(
            product_frame, 
            text="HSN NO:", 
            style="Bold.TLabel"
        ).grid(row=0, column=0, padx=5, pady=5, sticky="e")
        
        self.product_id_entry = ttk.Combobox(
            product_frame, 
            font=(self.config["font_family"], self.config["font_size"]),
            values=[p["hsn"] for p in self.products] if self.products else []
        )
        self.product_id_entry.grid(row=0, column=1, padx=5, pady=5, sticky="we")
        self.product_id_entry.bind("<KeyRelease>", self.search_products)
        
        # Product Description
        ttk.Label(
            product_frame, 
            text="Product Description:", 
            style="Bold.TLabel"
        ).grid(row=0, column=2, padx=5, pady=5, sticky="e")
        
        self.product_name_entry = ttk.Combobox(
            product_frame, 
            font=(self.config["font_family"], self.config["font_size"]),
            values=[p["name"] for p in self.products] if self.products else []
        )
        self.product_name_entry.grid(row=0, column=3, padx=5, pady=5, sticky="we")
        self.product_name_entry.bind("<KeyRelease>", self.search_products)
        
        # Price
        ttk.Label(
            product_frame, 
            text="Price:", 
            style="Bold.TLabel"
        ).grid(row=0, column=4, padx=5, pady=5, sticky="e")
        
        self.price_entry = ttk.Entry(
            product_frame, 
            font=(self.config["font_family"], self.config["font_size"])
        )
        self.price_entry.grid(row=0, column=5, padx=5, pady=5, sticky="we")
        
        # Quantity
        ttk.Label(
            product_frame, 
            text="Quantity:", 
            style="Bold.TLabel"
        ).grid(row=1, column=0, padx=5, pady=5, sticky="e")
        
        self.quantity_entry = ttk.Entry(
            product_frame, 
            font=(self.config["font_family"], self.config["font_size"])
        )
        self.quantity_entry.grid(row=1, column=1, padx=5, pady=5, sticky="we")
        
        # Add button
        self.add_button = ttk.Button(
            product_frame, 
            text="Add Product", 
            command=self.add_to_table,
            style="Accent.TButton"
        )
        self.add_button.grid(row=1, column=2, columnspan=2, padx=5, pady=5, sticky="we")
        
        # Quick add buttons
        quick_add_frame = ttk.Frame(product_frame)
        quick_add_frame.grid(row=1, column=4, columnspan=2, padx=5, pady=5, sticky="we")
        
        ttk.Button(
            quick_add_frame, 
            text="+1", 
            command=lambda: self.quick_add_quantity(1),
            style="Small.TButton"
        ).pack(side="left", expand=True, fill="x", padx=2)
        
        ttk.Button(
            quick_add_frame, 
            text="+5", 
            command=lambda: self.quick_add_quantity(5),
            style="Small.TButton"
        ).pack(side="left", expand=True, fill="x", padx=2)
        
        ttk.Button(
            quick_add_frame, 
            text="+10", 
            command=lambda: self.quick_add_quantity(10),
            style="Small.TButton"
        ).pack(side="left", expand=True, fill="x", padx=2)
        
        # Configure grid weights
        for i in range(6):
            product_frame.grid_columnconfigure(i, weight=1 if i in (1, 3, 5) else 0)

    def setup_product_table(self):
        """Setup the product table"""
        table_frame = ttk.LabelFrame(
            self.scrollable_frame_inner, 
            text="Products", 
            style="Section.TLabelframe"
        )
        table_frame.pack(fill="both", expand=True, pady=10, padx=5)
        
        # Create treeview with scrollbars
        self.product_table = ttk.Treeview(
            table_frame, 
            columns=("S.No", "HSN", "Product Description", "Price", "Quantity", "Total"), 
            show="headings",
            style="Custom.Treeview"
        )
        
        # Configure columns
        self.product_table.heading("S.No", text="S.No", anchor="center")
        self.product_table.heading("HSN", text="HSN", anchor="center")
        self.product_table.heading("Product Description", text="Product Description", anchor="center")
        self.product_table.heading("Price", text="Price", anchor="center")
        self.product_table.heading("Quantity", text="Quantity", anchor="center")
        self.product_table.heading("Total", text="Total", anchor="center")
        
        self.product_table.column("S.No", width=50, anchor="center")
        self.product_table.column("HSN", width=100, anchor="center")
        self.product_table.column("Product Description", width=300, anchor="w")
        self.product_table.column("Price", width=100, anchor="e")
        self.product_table.column("Quantity", width=80, anchor="center")
        self.product_table.column("Total", width=120, anchor="e")
        
        # Add scrollbars
        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.product_table.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.product_table.xview)
        self.product_table.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        
        # Grid layout
        self.product_table.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        
        # Configure grid weights
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Bind events
        self.product_table.bind("<Delete>", lambda e: self.clear_selected())
        self.product_table.bind("<Double-1>", self.edit_selected_item)

    def setup_totals(self):
        """Setup the totals section"""
        totals_frame = ttk.LabelFrame(
            self.scrollable_frame_inner, 
            text="Invoice Totals", 
            style="Section.TLabelframe"
        )
        totals_frame.pack(fill="x", pady=10, padx=5)
        
        # Tax and total variables
        self.sgst_var = tk.StringVar()
        self.igst_var = tk.StringVar()
        self.roundoff_var = tk.StringVar()
        self.total_cost_var = tk.StringVar()
        self.grand_total_words_var = tk.StringVar()
        
        # Subtotal
        ttk.Label(
            totals_frame, 
            text="Subtotal:", 
            style="Bold.TLabel"
        ).grid(row=0, column=0, padx=5, pady=5, sticky="e")
        
        self.subtotal_var = tk.StringVar()
        ttk.Label(
            totals_frame, 
            textvariable=self.subtotal_var,
            style="Total.TLabel"
        ).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # SGST
        ttk.Label(
            totals_frame, 
            text=f"SGST ({self.config['tax_rates']['sgst']}%):", 
            style="Bold.TLabel"
        ).grid(row=1, column=0, padx=5, pady=5, sticky="e")
        
        ttk.Label(
            totals_frame, 
            textvariable=self.sgst_var,
            style="Total.TLabel"
        ).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        # IGST
        ttk.Label(
            totals_frame, 
            text=f"IGST ({self.config['tax_rates']['igst']}%):", 
            style="Bold.TLabel"
        ).grid(row=2, column=0, padx=5, pady=5, sticky="e")
        
        ttk.Label(
            totals_frame, 
            textvariable=self.igst_var,
            style="Total.TLabel"
        ).grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        # Roundoff
        ttk.Label(
            totals_frame, 
            text="Roundoff:", 
            style="Bold.TLabel"
        ).grid(row=3, column=0, padx=5, pady=5, sticky="e")
        
        ttk.Label(
            totals_frame, 
            textvariable=self.roundoff_var,
            style="Total.TLabel"
        ).grid(row=3, column=1, padx=5, pady=5, sticky="w")
        
        # Grand Total
        ttk.Label(
            totals_frame, 
            text="Grand Total:", 
            style="Bold.TLabel"
        ).grid(row=4, column=0, padx=5, pady=5, sticky="e")
        
        ttk.Label(
            totals_frame, 
            textvariable=self.total_cost_var,
            style="GrandTotal.TLabel"
        ).grid(row=4, column=1, padx=5, pady=5, sticky="w")
        
        # Amount in words
        ttk.Label(
            totals_frame, 
            text="Amount in words:", 
            style="Bold.TLabel"
        ).grid(row=5, column=0, padx=5, pady=5, sticky="ne")
        
        ttk.Label(
            totals_frame, 
            textvariable=self.grand_total_words_var,
            style="AmountWords.TLabel",
            wraplength=400
        ).grid(row=5, column=1, padx=5, pady=5, sticky="w")
        
        # Configure grid weights
        totals_frame.grid_columnconfigure(1, weight=1)

    def setup_footer(self):
        """Setup the footer buttons"""
        button_frame = ttk.Frame(self.scrollable_frame_inner)
        button_frame.pack(fill="x", pady=10)
        
        # Left side buttons
        left_frame = ttk.Frame(button_frame)
        left_frame.pack(side="left", fill="x", expand=True)
        
        self.save_button = ttk.Button(
            left_frame, 
            text="Save Bill", 
            command=self.save_bill,
            style="Accent.TButton"
        )
        self.save_button.pack(side="left", padx=5)
        
        self.print_button = ttk.Button(
            left_frame, 
            text="Print Bill", 
            command=self.print_bill,
            style="Accent.TButton"
        )
        self.print_button.pack(side="left", padx=5)
        
        self.export_button = ttk.Button(
            left_frame,
            text="Export",
            command=self.export_to_excel,
            style="Secondary.TButton"
        )
        self.export_button.pack(side="left", padx=5)
        
        # Right side buttons
        right_frame = ttk.Frame(button_frame)
        right_frame.pack(side="right", fill="x", expand=True)
        
        self.clear_selected_button = ttk.Button(
            right_frame, 
            text="Clear Selected", 
            command=self.clear_selected,
            style="Secondary.TButton"
        )
        self.clear_selected_button.pack(side="right", padx=5)
        
        self.clear_all_button = ttk.Button(
            right_frame, 
            text="Clear All", 
            command=self.clear_all,
            style="Secondary.TButton"
        )
        self.clear_all_button.pack(side="right", padx=5)
        
        # Center buttons
        center_frame = ttk.Frame(button_frame)
        center_frame.pack(side="left", fill="x", expand=True)
        
        self.discount_button = ttk.Button(
            center_frame, 
            text="Apply Discount", 
            command=self.apply_discount,
            style="Secondary.TButton"
        )
        self.discount_button.pack(side="left", padx=5)
        
        self.history_button = ttk.Button(
            center_frame, 
            text="Product History", 
            command=self.show_product_history,
            style="Secondary.TButton"
        )
        self.history_button.pack(side="left", padx=5)
        
        self.qr_button = ttk.Button(
            center_frame,
            text="Generate QR",
            command=self.generate_qr_code,
            style="Secondary.TButton"
        )
        self.qr_button.pack(side="left", padx=5)

    def apply_styles(self):
        """Configure ttk styles"""
        style = ttk.Style()
        
        # Get current theme colors
        theme_colors = self.THEMES.get(self.current_theme, self.THEMES["Default"])
        primary = theme_colors["primary"]
        secondary = theme_colors["secondary"]
        accent = theme_colors["accent"]
        bg_color = theme_colors["bg"]
        text_color = theme_colors["text"]
        
        # Update config with theme colors
        self.config.update({
            "primary_color": primary,
            "secondary_color": secondary,
            "accent_color": accent,
            "bg_color": bg_color,
            "text_color": text_color
        })
        
        # Configure theme
        style.theme_use('clam')  # 'clam', 'alt', 'default', 'classic'
        
        # General style
        style.configure(".", 
                       background=bg_color,
                       foreground=text_color,
                       font=(self.config["font_family"], self.config["font_size"]))
        
        # Frame styles
        style.configure("Header.TFrame", background=primary)
        style.configure("Section.TLabelframe", 
                       borderwidth=2, 
                       relief="groove",
                       labelmargins=10)
        style.configure("Section.TLabelframe.Label", 
                       foreground=primary,
                       font=(self.config["font_family"], self.config["font_size"], "bold"))
        
        # Label styles
        style.configure("CompanyName.TLabel", 
                       foreground="white",
                       font=(self.config["font_family"], 18, "bold"),
                       background=primary)
        style.configure("CompanyAddress.TLabel", 
                       foreground="white",
                       font=(self.config["font_family"], self.config["font_size"]),
                       background=primary)
        style.configure("InvoiceInfo.TLabel", 
                       foreground="white",
                       font=(self.config["font_family"], self.config["font_size"], "bold"),
                       background=primary)
        style.configure("Bold.TLabel", 
                       font=(self.config["font_family"], self.config["font_size"], "bold"))
        style.configure("Total.TLabel", 
                       font=(self.config["font_family"], self.config["font_size"], "bold"),
                       foreground=secondary)
        style.configure("GrandTotal.TLabel", 
                       font=(self.config["font_family"], self.config["font_size"]+2, "bold"),
                       foreground=accent)
        style.configure("AmountWords.TLabel", 
                       font=(self.config["font_family"], self.config["font_size"]),
                       foreground=text_color)
        
        # Button styles
        style.configure("TButton", 
                       padding=6,
                       relief="flat")
        style.configure("Accent.TButton", 
                       background=accent,
                       foreground="white",
                       font=(self.config["font_family"], self.config["font_size"], "bold"))
        style.map("Accent.TButton",
                 background=[("active", accent), ("disabled", "#cccccc")])
        style.configure("Secondary.TButton", 
                       background=secondary,
                       foreground="white",
                       font=(self.config["font_family"], self.config["font_size"]))
        style.map("Secondary.TButton",
                 background=[("active", secondary), ("disabled", "#cccccc")])
        style.configure("Small.TButton", 
                       font=(self.config["font_family"], self.config["font_size"]-2))
        
        # Entry styles
        style.configure("TEntry", 
                       padding=5,
                       relief="solid")
        
        # Treeview styles
        style.configure("Custom.Treeview",
                       font=(self.config["font_family"], self.config["font_size"]),
                       rowheight=25)
        style.configure("Custom.Treeview.Heading", 
                       font=(self.config["font_family"], self.config["font_size"], "bold"),
                       background=primary,
                       foreground="white",
                       relief="flat")
        style.map("Custom.Treeview", 
                 background=[("selected", secondary)])
        
        # Radiobutton styles
        style.configure("BillType.TRadiobutton", 
                       foreground="white",
                       background=primary,
                       font=(self.config["font_family"], self.config["font_size"], "bold"))
        
        # Configure the main window background
        self.master.configure(background=bg_color)
        
        # Update all widgets
        self.update_widget_styles()

    def update_widget_styles(self):
        """Update styles for all widgets"""
        # Update status bar
        self.status_label.config(style="TLabel")
        self.auto_save_status.config(style="TLabel")
        
        # Update all tabs
        for child in self.notebook.winfo_children():
            child.configure(style="TFrame")
            for widget in child.winfo_children():
                if isinstance(widget, ttk.Frame):
                    widget.configure(style="TFrame")

    def setup_shortcuts(self):
        """Setup keyboard shortcuts"""
        self.master.bind("<Control-n>", lambda e: self.new_invoice())
        self.master.bind("<Control-s>", lambda e: self.save_bill())
        self.master.bind("<Control-p>", lambda e: self.print_bill())
        self.master.bind("<Control-e>", lambda e: self.export_to_excel())
        self.master.bind("<Delete>", lambda e: self.clear_selected())
        self.master.bind("<Control-Delete>", lambda e: self.clear_all())
        self.master.bind("<Return>", lambda e: self.add_to_table())
        self.master.bind("<F1>", lambda e: self.show_about())
        self.master.bind("<F5>", lambda e: self.refresh_data())

    def setup_auto_save(self):
        """Setup auto-save functionality"""
        if self.config["auto_save"]:
            self.auto_save_timer = threading.Timer(
                self.config["auto_save_interval"] * 60,
                self.auto_save
            )
            self.auto_save_timer.daemon = True
            self.auto_save_timer.start()

    def auto_save(self):
        """Auto-save the current invoice"""
        if self.product_table.get_children():
            try:
                # Save to a temporary file
                temp_dir = os.path.join(os.path.expanduser("~"), "temp_bills")
                if not os.path.exists(temp_dir):
                    os.makedirs(temp_dir)
                
                temp_file = os.path.join(temp_dir, f"temp_invoice_{self.invoice_number}.json")
                self.save_invoice_data(temp_file)
                
                # Update status bar
                self.status_label.config(text=f"Auto-saved at {datetime.now().strftime('%H:%M:%S')}")
            except Exception as e:
                print(f"Auto-save error: {e}")
        
        # Reset timer
        self.setup_auto_save()

    def toggle_auto_save(self):
        """Toggle auto-save on/off"""
        self.config["auto_save"] = not self.config["auto_save"]
        self.auto_save_status.config(
            text="Auto-save: ON" if self.config["auto_save"] else "Auto-save: OFF"
        )
        
        if self.config["auto_save"]:
            self.setup_auto_save()
        else:
            if hasattr(self, 'auto_save_timer'):
                self.auto_save_timer.cancel()
        
        self.save_config()

    def add_to_table(self):
        """Add product to the table"""
        if not all([self.product_id_entry.get(), self.product_name_entry.get(), 
                   self.price_entry.get(), self.quantity_entry.get()]):
            messagebox.showwarning("Warning", "All fields must be filled.")
            return

        try:
            product_id = self.product_id_entry.get()
            product_name = self.product_name_entry.get()
            price = float(self.price_entry.get())
            quantity = int(self.quantity_entry.get())
            subtotal = price * quantity

            self.product_table.insert("", "end", values=(
                len(self.product_table.get_children()) + 1, 
                product_id, 
                product_name, 
                f"{price:.2f}", 
                quantity, 
                f"{subtotal:.2f}"
            ))

            # Add to product history if not already there
            self.add_to_product_history(product_id, product_name, price)

            # Clear entry fields
            self.product_id_entry.delete(0, tk.END)
            self.product_name_entry.delete(0, tk.END)
            self.price_entry.delete(0, tk.END)
            self.quantity_entry.delete(0, tk.END)
            self.product_id_entry.focus()

            self.calculate_totals()
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numbers for price and quantity")

    def calculate_totals(self):
        """Calculate invoice totals"""
        subtotal = 0.0
        for child in self.product_table.get_children():
            subtotal += float(self.product_table.item(child, "values")[5])

        sgst_rate = self.config["tax_rates"]["sgst"] / 100
        igst_rate = self.config["tax_rates"]["igst"] / 100
        
        sgst = subtotal * sgst_rate
        igst = subtotal * igst_rate
        grand_total = subtotal + sgst + igst
        roundoff = round(grand_total) - grand_total
        final_total = grand_total + roundoff

        self.subtotal_var.set(f"{subtotal:.2f}")
        self.sgst_var.set(f"{sgst:.2f}")
        self.igst_var.set(f"{igst:.2f}")
        self.roundoff_var.set(f"{roundoff:.2f}")
        self.total_cost_var.set(f"{final_total:.2f}")

        grand_total_in_words = num2words(final_total).title() + " Rupees Only"
        self.grand_total_words_var.set(grand_total_in_words)

    def clear_selected(self):
        """Clear selected items from the table"""
        selected_items = self.product_table.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "No items selected")
            return
            
        for selected_item in selected_items:
            self.product_table.delete(selected_item)
        self.reorder_sno()
        self.calculate_totals()

    def clear_all(self):
        """Clear all items from the table"""
        if not messagebox.askyesno("Confirm", "Are you sure you want to clear all items?"):
            return
            
        for child in self.product_table.get_children():
            self.product_table.delete(child)
        self.calculate_totals()

    def reorder_sno(self):
        """Reorder serial numbers in the table"""
        for i, child in enumerate(self.product_table.get_children(), 1):
            values = self.product_table.item(child, 'values')
            self.product_table.item(child, values=(i, *values[1:]))

    def save_bill(self):
        """Save the bill as PDF"""
        if not self.product_table.get_children():
            messagebox.showwarning("Warning", "No products added to the invoice")
            return
            
        default_filename = f"Invoice_{self.invoice_number:04d}_{self.date.replace('-', '')}.pdf"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=default_filename
        )
        
        if file_path:
            self.generate_pdf(file_path)
            self.save_invoice_to_db(file_path)
            self.invoice_number += 1
            self.invoice_label.config(text=f"Invoice No: {self.invoice_number:04d}")
            self.clear_all()
            return file_path
        return None

    def print_bill(self):
        """Print the bill"""
        file_path = self.save_bill()
        if file_path:
            self.open_pdf(file_path)

    def generate_pdf(self, file_path):
        """Generate PDF invoice"""
        c = canvas.Canvas(file_path, pagesize=A4)
        width, height = A4

        primary_color = colors.HexColor(self.config["primary_color"])
        secondary_color = colors.HexColor(self.config["secondary_color"])
        accent_color = colors.HexColor(self.config["accent_color"])

        try:
            # Get the path of the logo
            if getattr(sys, 'frozen', False):  # Check if the application is bundled by PyInstaller
                logo_path = os.path.join(sys._MEIPASS, "logo.png")
            else:
                logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
                
            if os.path.exists(logo_path):
                c.drawImage(logo_path, 40, height - 80, width=50, height=50)
        except Exception as e:
            print(f"Error loading logo for PDF: {e}")

        # Header
        c.setFont("Helvetica-Bold", 16)
        c.setFillColor(primary_color)
        c.drawCentredString(width / 2.0, height - 50, self.config["company_name"])
        
        c.setFont("Helvetica", 12)
        c.setFillColor(colors.black)
        c.drawCentredString(width / 2.0, height - 70, self.config["company_address"])
        c.drawCentredString(width / 2.0, height - 90, f"Phone: {self.config['company_phone']} | Email: {self.config.get('company_email', '')}")
        
        c.drawString(30, height - 110, f"GSTIN: {self.config['gstin']}")
        c.drawRightString(width - 30, height - 110, f"Date: {self.date}")
        c.drawRightString(width - 30, height - 130, f"Invoice No: {self.invoice_number:04d}")
        c.drawRightString(width - 30, height - 150, f"Bill Type: {self.bill_type_var.get()}")

        # Customer info
        c.setFont("Helvetica-Bold", 12)
        c.drawString(30, height - 180, "Customer Name: ")
        c.drawString(30, height - 200, "Mobile Number: ")
        c.drawString(30, height - 220, "Place: ")
        c.drawString(30, height - 240, "Address: ")

        c.setFont("Helvetica", 12)
        c.drawString(150, height - 180, self.name_entry.get())
        c.drawString(150, height - 200, self.mobile_entry.get())
        c.drawString(150, height - 220, self.place_entry.get())
        c.drawString(150, height - 240, self.address_entry.get())

        # Products table
        c.setFont("Helvetica-Bold", 12)
        y_position = height - 280
        
        # Table header
        table_header = [
            ["S.No", "HSN", "Product Description", "Price", "Quantity", "Total"]
        ]
        
        # Table data
        table_data = []
        for i, child in enumerate(self.product_table.get_children(), 1):
            values = self.product_table.item(child, "values")
            table_data.append([str(i)] + list(values[1:]))
        
        # Combine header and data
        table_data = table_header + table_data
        
        # Create table
        table = Table(table_data, colWidths=[40, 80, 250, 60, 60, 60])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), primary_color),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
            ('ALIGN', (2, 1), (2, -1), 'LEFT'),  # Product description left-aligned
            ('ALIGN', (3, 1), (-1, -1), 'RIGHT'),  # Numbers right-aligned
        ]))
        
        # Draw table
        table.wrapOn(c, width, height)
        table.drawOn(c, 30, y_position - len(table_data) * 20)
        
        # Totals
        y_position -= (len(table_data) * 50 + 60)
        
        # Create totals table
        totals_data = [
            ["Subtotal:", self.subtotal_var.get()],
            # [f"SGST ({self.config['tax_rates']['sgst']}%):", self.sgst_var.get()],
            # [f"IGST ({self.config['tax_rates']['igst']}%):", self.igst_var.get()],
            ["Roundoff:", self.roundoff_var.get()],
            ["Grand Total:", self.total_cost_var.get()]
        ]
        
        totals_table = Table(totals_data, colWidths=[150, 100])
        totals_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('TEXTCOLOR', (0, -1), (-1, -1), accent_color),
            ('FONTSIZE', (0, -1), (-1, -1), 14),
            ('LINEABOVE', (0, 0), (-1, 0), 1, colors.lightgrey),
            ('LINEABOVE', (0, -1), (-1, -1), 1, colors.black),
        ]))
        
        totals_table.wrapOn(c, width, height)
        totals_table.drawOn(c, width - 300, y_position)
        
        # Amount in words
        styles = getSampleStyleSheet()
        styleN = styles['Normal']
        styleN.wordWrap = 'CJK'
        
        words = Paragraph(f"<b>Amount in words:</b> {self.grand_total_words_var.get()}", styleN)
        words.wrapOn(c, 500, 300)
        words.drawOn(c, 30, y_position - 40)
        
        # Bank details
        bank_details = [
            f"Bank: {self.config['bank_details']['name']}",
            f"A/C No: {self.config['bank_details']['account']}",
            f"IFSC: {self.config['bank_details']['ifsc']}",
            f"Branch: {self.config['bank_details']['branch']}"
        ]
        
        y_position -= 100
        for i, detail in enumerate(bank_details):
            c.drawString(30, y_position - (i * 20), detail)
        
        # Footer
        c.drawRightString(width - 30, y_position - 50, f"For {self.config['company_name']}")
        c.drawRightString(width - 60, y_position - 70, "Seal and Signature")
        
        # QR Code
        qr_data = f"""
        Company: {self.config['company_name']}
        Invoice No: {self.invoice_number:04d}
        Date: {self.date}
        Customer: {self.name_entry.get()}
        Total: {self.total_cost_var.get()}
        """
        
        try:
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=4,
                border=4,
            )
            qr.add_data(qr_data)
            qr.make(fit=True)
            
            qr_img = qr.make_image(fill_color="black", back_color="white")
            qr_img_path = os.path.join(os.path.dirname(file_path), f"qr_{self.invoice_number}.png")
            qr_img.save(qr_img_path)
            
            # Draw QR code on PDF
            c.drawImage(qr_img_path, 30, y_position - 150, width=80, height=80)
            os.remove(qr_img_path)
        except Exception as e:
            print(f"Error generating QR code: {e}")
        
        c.showPage()
        c.save()

    def open_pdf(self, file_path):
        """Open PDF file with default viewer"""
        try:
            webbrowser.open(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open PDF: {e}")

    def new_invoice(self):
        """Create a new invoice"""
        if self.product_table.get_children() and not messagebox.askyesno(
            "New Invoice", 
            "Current invoice has items. Create new invoice anyway?"
        ):
            return
            
        self.clear_all()
        self.invoice_number += 1
        self.date = datetime.now().strftime("%d-%m-%Y")
        self.invoice_label.config(text=f"Invoice No: {self.invoice_number:04d}")
        self.date_label.config(text=f"Date: {self.date}")
        self.name_entry.focus()

    def edit_selected_item(self, event):
        """Edit selected item in the table"""
        selected_items = self.product_table.selection()
        if not selected_items:
            return
            
        selected_item = selected_items[0]
        values = self.product_table.item(selected_item, "values")
        
        # Create edit dialog
        edit_dialog = tk.Toplevel(self.master)
        edit_dialog.title("Edit Product")
        edit_dialog.transient(self.master)
        edit_dialog.grab_set()
        
        # Product ID
        ttk.Label(edit_dialog, text="HSN NO:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        product_id_entry = ttk.Entry(edit_dialog)
        product_id_entry.grid(row=0, column=1, padx=5, pady=5)
        product_id_entry.insert(0, values[1])
        
        # Product Name
        ttk.Label(edit_dialog, text="Product Description:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        product_name_entry = ttk.Entry(edit_dialog)
        product_name_entry.grid(row=1, column=1, padx=5, pady=5)
        product_name_entry.insert(0, values[2])
        
        # Price
        ttk.Label(edit_dialog, text="Price:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        price_entry = ttk.Entry(edit_dialog)
        price_entry.grid(row=2, column=1, padx=5, pady=5)
        price_entry.insert(0, values[3])
        
        # Quantity
        ttk.Label(edit_dialog, text="Quantity:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        quantity_entry = ttk.Entry(edit_dialog)
        quantity_entry.grid(row=3, column=1, padx=5, pady=5)
        quantity_entry.insert(0, values[4])
        
        # Save button
        def save_changes():
            try:
                new_values = (
                    values[0],  # Keep same S.No
                    product_id_entry.get(),
                    product_name_entry.get(),
                    f"{float(price_entry.get()):.2f}",
                    int(quantity_entry.get()),
                    f"{float(price_entry.get()) * int(quantity_entry.get()):.2f}"
                )
                self.product_table.item(selected_item, values=new_values)
                self.calculate_totals()
                edit_dialog.destroy()
            except ValueError:
                messagebox.showerror("Error", "Please enter valid numbers for price and quantity")
        
        ttk.Button(
            edit_dialog, 
            text="Save", 
            command=save_changes,
            style="Accent.TButton"
        ).grid(row=4, column=0, columnspan=2, pady=10)

    def quick_add_quantity(self, amount):
        """Quickly add quantity to the quantity field"""
        current = self.quantity_entry.get()
        try:
            new_quantity = int(current) + amount if current else amount
            self.quantity_entry.delete(0, tk.END)
            self.quantity_entry.insert(0, str(new_quantity))
        except ValueError:
            self.quantity_entry.delete(0, tk.END)
            self.quantity_entry.insert(0, str(amount))

    def apply_discount(self):
        """Apply discount to the invoice"""
        discount = simpledialog.askfloat(
            "Apply Discount", 
            "Enter discount amount (positive number):",
            parent=self.master,
            minvalue=0
        )
        
        if discount is not None:
            subtotal = float(self.subtotal_var.get())
            new_subtotal = subtotal - discount
            if new_subtotal < 0:
                messagebox.showerror("Error", "Discount cannot be greater than subtotal")
                return
                
            self.subtotal_var.set(f"{new_subtotal:.2f}")
            self.calculate_totals()

    def appearance_settings(self):
        """Open appearance settings dialog"""
        settings_dialog = tk.Toplevel(self.master)
        settings_dialog.title("Appearance Settings")
        settings_dialog.transient(self.master)
        settings_dialog.grab_set()
        
        # Theme selection
        ttk.Label(settings_dialog, text="Theme:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        theme_var = tk.StringVar(value=self.current_theme)
        theme_menu = ttk.OptionMenu(
            settings_dialog, 
            theme_var, 
            self.current_theme, 
            *self.THEMES.keys()
        )
        theme_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # Font family
        ttk.Label(settings_dialog, text="Font Family:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        font_family_entry = ttk.Combobox(settings_dialog, values=tkfont.families())
        font_family_entry.grid(row=1, column=1, padx=5, pady=5)
        font_family_entry.set(self.config["font_family"])
        
        # Font size
        ttk.Label(settings_dialog, text="Font Size:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        font_size_entry = ttk.Spinbox(settings_dialog, from_=8, to=24)
        font_size_entry.grid(row=2, column=1, padx=5, pady=5)
        font_size_entry.set(self.config["font_size"])
        
        # Auto-save
        auto_save_var = tk.BooleanVar(value=self.config["auto_save"])
        ttk.Checkbutton(
            settings_dialog, 
            text="Enable Auto-save", 
            variable=auto_save_var
        ).grid(row=3, column=0, columnspan=2, pady=5, sticky="w")
        
        # Auto-save interval
        ttk.Label(settings_dialog, text="Auto-save Interval (minutes):").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        auto_save_interval = ttk.Spinbox(settings_dialog, from_=1, to=60)
        auto_save_interval.grid(row=4, column=1, padx=5, pady=5)
        auto_save_interval.set(self.config["auto_save_interval"])
        
        def save_settings():
            self.current_theme = theme_var.get()
            self.config.update({
                "font_family": font_family_entry.get(),
                "font_size": int(font_size_entry.get()),
                "auto_save": auto_save_var.get(),
                "auto_save_interval": int(auto_save_interval.get()),
                "default_theme": self.current_theme
            })
            self.save_config()
            self.apply_styles()
            self.toggle_auto_save()  # Restart auto-save timer if needed
            settings_dialog.destroy()
            messagebox.showinfo("Success", "Appearance settings saved.")
        
        ttk.Button(
            settings_dialog, 
            text="Save", 
            command=save_settings,
            style="Accent.TButton"
        ).grid(row=5, column=0, columnspan=2, pady=10)

    def company_settings(self):
        """Open company settings dialog"""
        settings_dialog = tk.Toplevel(self.master)
        settings_dialog.title("Company Settings")
        settings_dialog.transient(self.master)
        settings_dialog.grab_set()
        
        # Company name
        ttk.Label(settings_dialog, text="Company Name:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        company_name_entry = ttk.Entry(settings_dialog)
        company_name_entry.grid(row=0, column=1, padx=5, pady=5)
        company_name_entry.insert(0, self.config["company_name"])
        
        # Company address
        ttk.Label(settings_dialog, text="Company Address:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        company_address_entry = ttk.Entry(settings_dialog)
        company_address_entry.grid(row=1, column=1, padx=5, pady=5)
        company_address_entry.insert(0, self.config["company_address"])
        
        # Company phone
        ttk.Label(settings_dialog, text="Company Phone:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        company_phone_entry = ttk.Entry(settings_dialog)
        company_phone_entry.grid(row=2, column=1, padx=5, pady=5)
        company_phone_entry.insert(0, self.config["company_phone"])
        
        # Company email
        ttk.Label(settings_dialog, text="Company Email:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        company_email_entry = ttk.Entry(settings_dialog)
        company_email_entry.grid(row=3, column=1, padx=5, pady=5)
        company_email_entry.insert(0, self.config.get("company_email", ""))
        
        # Company website
        ttk.Label(settings_dialog, text="Company Website:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        company_website_entry = ttk.Entry(settings_dialog)
        company_website_entry.grid(row=4, column=1, padx=5, pady=5)
        company_website_entry.insert(0, self.config.get("company_website", ""))
        
        # GSTIN
        ttk.Label(settings_dialog, text="GSTIN:").grid(row=5, column=0, padx=5, pady=5, sticky="e")
        gstin_entry = ttk.Entry(settings_dialog)
        gstin_entry.grid(row=5, column=1, padx=5, pady=5)
        gstin_entry.insert(0, self.config["gstin"])
        
        # Bank details
        ttk.Label(settings_dialog, text="Bank Name:").grid(row=6, column=0, padx=5, pady=5, sticky="e")
        bank_name_entry = ttk.Entry(settings_dialog)
        bank_name_entry.grid(row=6, column=1, padx=5, pady=5)
        bank_name_entry.insert(0, self.config["bank_details"]["name"])
        
        ttk.Label(settings_dialog, text="Account Number:").grid(row=7, column=0, padx=5, pady=5, sticky="e")
        account_entry = ttk.Entry(settings_dialog)
        account_entry.grid(row=7, column=1, padx=5, pady=5)
        account_entry.insert(0, self.config["bank_details"]["account"])
        
        ttk.Label(settings_dialog, text="IFSC Code:").grid(row=8, column=0, padx=5, pady=5, sticky="e")
        ifsc_entry = ttk.Entry(settings_dialog)
        ifsc_entry.grid(row=8, column=1, padx=5, pady=5)
        ifsc_entry.insert(0, self.config["bank_details"]["ifsc"])
        
        ttk.Label(settings_dialog, text="Branch:").grid(row=9, column=0, padx=5, pady=5, sticky="e")
        branch_entry = ttk.Entry(settings_dialog)
        branch_entry.grid(row=9, column=1, padx=5, pady=5)
        branch_entry.insert(0, self.config["bank_details"]["branch"])
        
        def save_settings():
            self.config.update({
                "company_name": company_name_entry.get(),
                "company_address": company_address_entry.get(),
                "company_phone": company_phone_entry.get(),
                "company_email": company_email_entry.get(),
                "company_website": company_website_entry.get(),
                "gstin": gstin_entry.get(),
                "bank_details": {
                    "name": bank_name_entry.get(),
                    "account": account_entry.get(),
                    "ifsc": ifsc_entry.get(),
                    "branch": branch_entry.get()
                }
            })
            self.save_config()
            
            # Update UI
            self.company_name_label.config(text=self.config["company_name"])
            self.address_label_1.config(text=self.config["company_address"])
            self.address_label_2.config(text=f"Phone: {self.config['company_phone']} | Email: {self.config.get('company_email', '')}")
            self.gstin_label.config(text=f"GSTIN: {self.config['gstin']}")
            
            settings_dialog.destroy()
            messagebox.showinfo("Success", "Company settings saved successfully.")
        
        ttk.Button(
            settings_dialog, 
            text="Save", 
            command=save_settings,
            style="Accent.TButton"
        ).grid(row=10, column=0, columnspan=2, pady=10)

    def tax_settings(self):
        """Open tax settings dialog"""
        settings_dialog = tk.Toplevel(self.master)
        settings_dialog.title("Tax Settings")
        settings_dialog.transient(self.master)
        settings_dialog.grab_set()
        
        # SGST
        ttk.Label(settings_dialog, text="SGST Rate (%):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        sgst_entry = ttk.Spinbox(settings_dialog, from_=0, to=100, increment=0.1)
        sgst_entry.grid(row=0, column=1, padx=5, pady=5)
        sgst_entry.set(self.config["tax_rates"]["sgst"])
        
        # IGST
        ttk.Label(settings_dialog, text="IGST Rate (%):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        igst_entry = ttk.Spinbox(settings_dialog, from_=0, to=100, increment=0.1)
        igst_entry.grid(row=1, column=1, padx=5, pady=5)
        igst_entry.set(self.config["tax_rates"]["igst"])
        
        def save_settings():
            self.config["tax_rates"].update({
                "sgst": float(sgst_entry.get()),
                "igst": float(igst_entry.get())
            })
            self.save_config()
            self.calculate_totals()
            settings_dialog.destroy()
            messagebox.showinfo("Success", "Tax settings saved successfully.")
        
        ttk.Button(
            settings_dialog, 
            text="Save", 
            command=save_settings,
            style="Accent.TButton"
        ).grid(row=2, column=0, columnspan=2, pady=10)

    def database_settings(self):
        """Open database settings dialog"""
        settings_dialog = tk.Toplevel(self.master)
        settings_dialog.title("Database Settings")
        settings_dialog.transient(self.master)
        settings_dialog.grab_set()
        
        # Backup button
        ttk.Button(
            settings_dialog,
            text="Backup Database",
            command=self.backup_database,
            style="Accent.TButton"
        ).grid(row=0, column=0, padx=5, pady=5)
        
        # Restore button
        ttk.Button(
            settings_dialog,
            text="Restore Database",
            command=self.restore_database,
            style="Secondary.TButton"
        ).grid(row=0, column=1, padx=5, pady=5)
        
        # Export data button
        ttk.Button(
            settings_dialog,
            text="Export Data to Excel",
            command=self.export_database_to_excel,
            style="Accent.TButton"
        ).grid(row=1, column=0, padx=5, pady=5)
        
        # Import data button
        ttk.Button(
            settings_dialog,
            text="Import Data from Excel",
            command=self.import_database_from_excel,
            style="Secondary.TButton"
        ).grid(row=1, column=1, padx=5, pady=5)
        
        # Status label
        self.db_status_label = ttk.Label(settings_dialog, text="")
        self.db_status_label.grid(row=2, column=0, columnspan=2, pady=5)

    def backup_database(self):
        """Backup the database to a file"""
        backup_file = filedialog.asksaveasfilename(
            defaultextension=".db",
            filetypes=[("Database files", "*.db"), ("All files", "*.*")],
            initialfile="billing_backup.db"
        )
        
        if backup_file:
            try:
                # Close current connection
                self.conn.close()
                
                # Copy file
                import shutil
                shutil.copy2(self.DB_FILE, backup_file)
                
                # Reopen connection
                self.conn = sqlite3.connect(self.DB_FILE)
                self.cursor = self.conn.cursor()
                
                self.db_status_label.config(text=f"Backup created: {backup_file}")
            except Exception as e:
                self.db_status_label.config(text=f"Backup failed: {str(e)}")
                # Try to reopen connection even if backup failed
                self.conn = sqlite3.connect(self.DB_FILE)
                self.cursor = self.conn.cursor()

    def restore_database(self):
        """Restore the database from a backup"""
        if not messagebox.askyesno("Confirm", "This will overwrite your current database. Continue?"):
            return
            
        backup_file = filedialog.askopenfilename(
            filetypes=[("Database files", "*.db"), ("All files", "*.*")]
        )
        
        if backup_file:
            try:
                # Close current connection
                self.conn.close()
                
                # Copy backup file
                import shutil
                shutil.copy2(backup_file, self.DB_FILE)
                
                # Reopen connection
                self.conn = sqlite3.connect(self.DB_FILE)
                self.cursor = self.conn.cursor()
                
                self.db_status_label.config(text=f"Database restored from: {backup_file}")
                messagebox.showinfo("Success", "Database restored successfully. Please restart the application.")
            except Exception as e:
                self.db_status_label.config(text=f"Restore failed: {str(e)}")
                # Try to reopen connection even if restore failed
                self.conn = sqlite3.connect(self.DB_FILE)
                self.cursor = self.conn.cursor()

    def export_database_to_excel(self):
        """Export database tables to Excel"""
        export_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="billing_data_export.xlsx"
        )
        
        if export_file:
            try:
                # Get all table names
                self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
                tables = [table[0] for table in self.cursor.fetchall()]
                
                # Create Excel writer
                with pd.ExcelWriter(export_file) as writer:
                    for table in tables:
                        # Read table data
                        df = pd.read_sql_query(f"SELECT * FROM {table}", self.conn)
                        # Write to Excel sheet
                        df.to_excel(writer, sheet_name=table, index=False)
                
                self.db_status_label.config(text=f"Data exported to: {export_file}")
                messagebox.showinfo("Success", "Database exported to Excel successfully.")
            except Exception as e:
                self.db_status_label.config(text=f"Export failed: {str(e)}")

    def import_database_from_excel(self):
        """Import data from Excel to database"""
        if not messagebox.askyesno("Confirm", "This will overwrite existing data in the database. Continue?"):
            return
            
        import_file = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if import_file:
            try:
                # Read Excel file
                excel_data = pd.ExcelFile(import_file)
                
                # Process each sheet (table)
                for sheet_name in excel_data.sheet_names:
                    df = excel_data.parse(sheet_name)
                    
                    # Delete existing data
                    self.cursor.execute(f"DELETE FROM {sheet_name}")
                    
                    # Insert new data
                    df.to_sql(sheet_name, self.conn, if_exists='append', index=False)
                
                self.conn.commit()
                self.db_status_label.config(text=f"Data imported from: {import_file}")
                messagebox.showinfo("Success", "Database imported from Excel successfully. Please refresh views.")
            except Exception as e:
                self.db_status_label.config(text=f"Import failed: {str(e)}")

    def show_about(self):
        """Show about dialog"""
        about_dialog = tk.Toplevel(self.master)
        about_dialog.title("About Billing System")
        about_dialog.transient(self.master)
        about_dialog.resizable(False, False)
        
        ttk.Label(
            about_dialog, 
            text="Advanced Billing System", 
            font=(self.config["font_family"], 16, "bold"),
            foreground=self.config["primary_color"]
        ).pack(pady=10)
        
        ttk.Label(
            about_dialog, 
            text="Version 3.0\n\nA comprehensive billing solution\nwith advanced features and customization",
            justify="center"
        ).pack(pady=5)
        
        ttk.Label(
            about_dialog, 
            text="Features:\n"
                 "- Invoice generation with GST\n"
                 "- Product and customer management\n"
                 "- Sales reporting and analytics\n"
                 "- Database backup and restore\n"
                 "- Customizable themes and appearance",
            justify="left"
        ).pack(pady=5)
        
        ttk.Label(
            about_dialog, 
            text=f"© {datetime.now().year} Draupathi IT Solutions",
            font=(self.config["font_family"], 10)
        ).pack(pady=10)
        
        ttk.Button(
            about_dialog, 
            text="Close", 
            command=about_dialog.destroy,
            style="Accent.TButton"
        ).pack(pady=10)

    def show_user_guide(self):
        """Show user guide in browser"""
        try:
            webbrowser.open("https://github.com/yourusername/billing-system/blob/main/docs/user_guide.md")
        except:
            messagebox.showinfo("User Guide", "Please check the documentation for user guide.")

    def check_for_updates(self):
        """Check for software updates"""
        # This would typically connect to a server to check for updates
        # For now, just show a message
        messagebox.showinfo("Check for Updates", "You are using the latest version.")

    def save_invoice_to_db(self, file_path):
        """Save invoice data to database"""
        try:
            # Save invoice
            self.cursor.execute('''
                INSERT INTO invoices (
                    invoice_number, date, customer_name, customer_mobile, 
                    customer_place, customer_address, bill_type, subtotal, 
                    sgst, igst, roundoff, total, pdf_path
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                self.invoice_number,
                self.date,
                self.name_entry.get(),
                self.mobile_entry.get(),
                self.place_entry.get(),
                self.address_entry.get(),
                self.bill_type_var.get(),
                float(self.subtotal_var.get()),
                float(self.sgst_var.get()),
                float(self.igst_var.get()),
                float(self.roundoff_var.get()),
                float(self.total_cost_var.get()),
                file_path
            ))
            
            # Get the invoice ID
            invoice_id = self.cursor.lastrowid
            
            # Save invoice items
            for child in self.product_table.get_children():
                values = self.product_table.item(child, "values")
                self.cursor.execute('''
                    INSERT INTO invoice_items (
                        invoice_id, sno, hsn, description, price, quantity, total
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    invoice_id,
                    values[0],
                    values[1],
                    values[2],
                    float(values[3]),
                    int(values[4]),
                    float(values[5])
                ))
            
            # Commit changes
            self.conn.commit()
            
            # Add customer to database if not exists
            if self.mobile_entry.get():
                self.cursor.execute('''
                    INSERT OR IGNORE INTO customers (name, mobile, place, address)
                    VALUES (?, ?, ?, ?)
                ''', (
                    self.name_entry.get(),
                    self.mobile_entry.get(),
                    self.place_entry.get(),
                    self.address_entry.get()
                ))
                self.conn.commit()
            
            return True
        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Database Error", f"Failed to save invoice: {str(e)}")
            return False

    def find_invoice(self):
        """Find and load an existing invoice"""
        find_dialog = tk.Toplevel(self.master)
        find_dialog.title("Find Invoice")
        find_dialog.transient(self.master)
        find_dialog.grab_set()
        
        ttk.Label(find_dialog, text="Search by:").grid(row=0, column=0, padx=5, pady=5)
        
        search_type = tk.StringVar(value="invoice_number")
        ttk.Radiobutton(
            find_dialog, 
            text="Invoice Number", 
            variable=search_type, 
            value="invoice_number"
        ).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Radiobutton(
            find_dialog, 
            text="Customer Mobile", 
            variable=search_type, 
            value="customer_mobile"
        ).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Label(find_dialog, text="Search Value:").grid(row=2, column=0, padx=5, pady=5)
        search_entry = ttk.Entry(find_dialog)
        search_entry.grid(row=2, column=1, padx=5, pady=5)
        
        results_frame = ttk.Frame(find_dialog)
        results_frame.grid(row=3, column=0, columnspan=2, padx=5, pady=5)
        
        results_tree = ttk.Treeview(
            results_frame,
            columns=("ID", "Invoice No", "Date", "Customer", "Mobile", "Total"),
            show="headings"
        )
        
        results_tree.heading("ID", text="ID")
        results_tree.heading("Invoice No", text="Invoice No")
        results_tree.heading("Date", text="Date")
        results_tree.heading("Customer", text="Customer")
        results_tree.heading("Mobile", text="Mobile")
        results_tree.heading("Total", text="Total")
        
        results_tree.column("ID", width=50, anchor="center")
        results_tree.column("Invoice No", width=100, anchor="center")
        results_tree.column("Date", width=100, anchor="center")
        results_tree.column("Customer", width=150, anchor="w")
        results_tree.column("Mobile", width=100, anchor="center")
        results_tree.column("Total", width=80, anchor="e")
        
        y_scroll = ttk.Scrollbar(results_frame, orient="vertical", command=results_tree.yview)
        results_tree.configure(yscrollcommand=y_scroll.set)
        
        results_tree.pack(side="left", fill="both", expand=True)
        y_scroll.pack(side="right", fill="y")
        
        def search_invoices():
            """Search invoices based on criteria"""
            for item in results_tree.get_children():
                results_tree.delete(item)
            
            query = f"SELECT id, invoice_number, date, customer_name, customer_mobile, total FROM invoices WHERE {search_type.get()} LIKE ? ORDER BY date DESC"
            self.cursor.execute(query, (f"%{search_entry.get()}%",))
            
            for row in self.cursor.fetchall():
                results_tree.insert("", "end", values=row)
        
        def load_invoice():
            """Load selected invoice"""
            selected = results_tree.selection()
            if not selected:
                return
                
            invoice_id = results_tree.item(selected[0], "values")[0]
            self.load_invoice_from_db(invoice_id)
            find_dialog.destroy()
        
        ttk.Button(
            find_dialog, 
            text="Search", 
            command=search_invoices
        ).grid(row=4, column=0, padx=5, pady=5)
        
        ttk.Button(
            find_dialog, 
            text="Load Selected", 
            command=load_invoice,
            style="Accent.TButton"
        ).grid(row=4, column=1, padx=5, pady=5)

    def load_invoice_from_db(self, invoice_id):
        """Load invoice from database"""
        try:
            # Get invoice details
            self.cursor.execute('''
                SELECT invoice_number, date, customer_name, customer_mobile, 
                       customer_place, customer_address, bill_type, subtotal, 
                       sgst, igst, roundoff, total 
                FROM invoices WHERE id = ?
            ''', (invoice_id,))
            invoice_data = self.cursor.fetchone()
            
            if not invoice_data:
                messagebox.showerror("Error", "Invoice not found")
                return
                
            # Clear current invoice
            self.clear_all()
            
            # Set invoice details
            self.invoice_number = invoice_data[0]
            self.date = invoice_data[1]
            self.invoice_label.config(text=f"Invoice No: {self.invoice_number:04d}")
            self.date_label.config(text=f"Date: {self.date}")
            
            # Set customer details
            self.name_entry.insert(0, invoice_data[2])
            self.mobile_entry.insert(0, invoice_data[3])
            self.place_entry.insert(0, invoice_data[4])
            self.address_entry.insert(0, invoice_data[5])
            self.bill_type_var.set(invoice_data[6])
            
            # Get invoice items
            self.cursor.execute('''
                SELECT sno, hsn, description, price, quantity, total 
                FROM invoice_items 
                WHERE invoice_id = ?
                ORDER BY sno
            ''', (invoice_id,))
            
            items = self.cursor.fetchall()
            
            # Add items to table
            for item in items:
                self.product_table.insert("", "end", values=item)
            
            # Set totals
            self.subtotal_var.set(f"{invoice_data[7]:.2f}")
            self.sgst_var.set(f"{invoice_data[8]:.2f}")
            self.igst_var.set(f"{invoice_data[9]:.2f}")
            self.roundoff_var.set(f"{invoice_data[10]:.2f}")
            self.total_cost_var.set(f"{invoice_data[11]:.2f}")
            
            # Update amount in words
            grand_total_in_words = num2words(float(invoice_data[11])).title() + " Rupees Only"
            self.grand_total_words_var.set(grand_total_in_words)
            
            messagebox.showinfo("Success", "Invoice loaded successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load invoice: {str(e)}")

    def add_to_product_history(self, hsn, name, price):
        """Add product to history if not already exists"""
        try:
            # Check if product already exists in database
            self.cursor.execute("SELECT id FROM products WHERE hsn = ?", (hsn,))
            if not self.cursor.fetchone():
                # Add to database
                self.cursor.execute('''
                    INSERT INTO products (hsn, name, price)
                    VALUES (?, ?, ?)
                ''', (hsn, name, price))
                self.conn.commit()
                
                # Add to local list
                self.products.append({
                    "hsn": hsn,
                    "name": name,
                    "price": price
                })
                
                # Update combobox values
                self.product_id_entry["values"] = [p["hsn"] for p in self.products]
                self.product_name_entry["values"] = [p["name"] for p in self.products]
        except Exception as e:
            print(f"Error adding product to database: {e}")

    def save_product_history(self):
        """Save product history to database"""
        # This is now handled by add_to_product_history which saves directly to database
        pass

    def load_product_history(self):
        """Load product history from database"""
        try:
            self.cursor.execute("SELECT hsn, name, price FROM products")
            self.products = []
            for row in self.cursor.fetchall():
                self.products.append({
                    "hsn": row[0],
                    "name": row[1],
                    "price": row[2]
                })
        except Exception as e:
            print(f"Error loading product history: {e}")

    def search_products(self, event):
        """Search products based on HSN or name"""
        search_term = event.widget.get().lower()
        
        if event.widget == self.product_id_entry:
            # Search by HSN
            results = [p["hsn"] for p in self.products if search_term in p["hsn"].lower()]
            self.product_id_entry["values"] = results
        elif event.widget == self.product_name_entry:
            # Search by name
            results = [p["name"] for p in self.products if search_term in p["name"].lower()]
            self.product_name_entry["values"] = results

    def show_product_history(self):
        """Show product history in a new window"""
        if not self.products:
            messagebox.showinfo("Product History", "No product history available")
            return
            
        history_window = tk.Toplevel(self.master)
        history_window.title("Product History")
        history_window.geometry("800x600")
        
        # Create treeview
        tree = ttk.Treeview(
            history_window, 
            columns=("HSN", "Product Name", "Price"), 
            show="headings",
            style="Custom.Treeview"
        )
        
        tree.heading("HSN", text="HSN")
        tree.heading("Product Name", text="Product Name")
        tree.heading("Price", text="Price")
        
        tree.column("HSN", width=150, anchor="center")
        tree.column("Product Name", width=400, anchor="w")
        tree.column("Price", width=150, anchor="e")
        
        # Add products
        for product in self.products:
            tree.insert("", "end", values=(product["hsn"], product["name"], f"{product['price']:.2f}"))
        
        # Add scrollbars
        y_scroll = ttk.Scrollbar(history_window, orient="vertical", command=tree.yview)
        x_scroll = ttk.Scrollbar(history_window, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        
        # Grid layout
        tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        
        # Button frame
        button_frame = ttk.Frame(history_window)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10, sticky="e")
        
        # Select button
        def select_product():
            selected = tree.selection()
            if not selected:
                return
                
            product = tree.item(selected[0], "values")
            self.product_id_entry.delete(0, tk.END)
            self.product_id_entry.insert(0, product[0])
            self.product_name_entry.delete(0, tk.END)
            self.product_name_entry.insert(0, product[1])
            self.price_entry.delete(0, tk.END)
            self.price_entry.insert(0, product[2])
            self.quantity_entry.focus()
            history_window.destroy()
        
        ttk.Button(
            button_frame, 
            text="Select", 
            command=select_product,
            style="Accent.TButton"
        ).pack(side="left", padx=5)
        
        # Close button
        ttk.Button(
            button_frame, 
            text="Close", 
            command=history_window.destroy,
            style="Secondary.TButton"
        ).pack(side="left", padx=5)
        
        # Configure grid weights
        history_window.grid_rowconfigure(0, weight=1)
        history_window.grid_columnconfigure(0, weight=1)

    def export_to_excel(self):
        """Export current invoice to Excel"""
        if not self.product_table.get_children():
            messagebox.showwarning("Warning", "No products added to the invoice")
            return
            
        default_filename = f"Invoice_{self.invoice_number:04d}_{self.date.replace('-', '')}.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_filename
        )
        
        if file_path:
            try:
                # Create DataFrame for products
                products_data = []
                for child in self.product_table.get_children():
                    values = self.product_table.item(child, "values")
                    products_data.append({
                        "S.No": values[0],
                        "HSN": values[1],
                        "Product Description": values[2],
                        "Price": float(values[3]),
                        "Quantity": int(values[4]),
                        "Total": float(values[5])
                    })
                
                df_products = pd.DataFrame(products_data)
                
                # Create DataFrame for totals
                totals_data = [{
                    "Subtotal": float(self.subtotal_var.get()),
                    f"SGST ({self.config['tax_rates']['sgst']}%)": float(self.sgst_var.get()),
                    f"IGST ({self.config['tax_rates']['igst']}%)": float(self.igst_var.get()),
                    "Roundoff": float(self.roundoff_var.get()),
                    "Grand Total": float(self.total_cost_var.get())
                }]
                
                df_totals = pd.DataFrame(totals_data)
                
                # Create Excel writer
                with pd.ExcelWriter(file_path) as writer:
                    # Write products sheet
                    df_products.to_excel(writer, sheet_name="Products", index=False)
                    
                    # Write totals sheet
                    df_totals.to_excel(writer, sheet_name="Totals", index=False)
                    
                    # Write invoice info sheet
                    invoice_info = {
                        "Invoice Number": [self.invoice_number],
                        "Date": [self.date],
                        "Customer Name": [self.name_entry.get()],
                        "Customer Mobile": [self.mobile_entry.get()],
                        "Customer Place": [self.place_entry.get()],
                        "Customer Address": [self.address_entry.get()],
                        "Bill Type": [self.bill_type_var.get()],
                        "Amount in Words": [self.grand_total_words_var.get()]
                    }
                    
                    df_info = pd.DataFrame(invoice_info)
                    df_info.to_excel(writer, sheet_name="Invoice Info", index=False)
                
                messagebox.showinfo("Success", f"Invoice exported to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export invoice: {str(e)}")

    def generate_qr_code(self):
        """Generate QR code for the current invoice"""
        if not self.product_table.get_children():
            messagebox.showwarning("Warning", "No products added to the invoice")
            return
            
        qr_data = f"""
        Company: {self.config['company_name']}
        Invoice No: {self.invoice_number:04d}
        Date: {self.date}
        Customer: {self.name_entry.get()}
        Total: {self.total_cost_var.get()}
        """
        
        try:
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4,
            )
            qr.add_data(qr_data)
            qr.make(fit=True)
            
            img = qr.make_image(fill_color="black", back_color="white")
            
            # Show QR code in a new window
            qr_window = tk.Toplevel(self.master)
            qr_window.title("Invoice QR Code")
            
            # Convert PIL image to Tkinter PhotoImage
            tk_img = ImageTk.PhotoImage(img)
            
            label = ttk.Label(qr_window, image=tk_img)
            label.image = tk_img  # Keep a reference
            label.pack(padx=10, pady=10)
            
            ttk.Button(
                qr_window,
                text="Save QR Code",
                command=lambda: self.save_qr_code(img),
                style="Accent.TButton"
            ).pack(pady=5)
            
            ttk.Button(
                qr_window,
                text="Close",
                command=qr_window.destroy,
                style="Secondary.TButton"
            ).pack(pady=5)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate QR code: {str(e)}")

    def save_qr_code(self, qr_img):
        """Save QR code to file"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png")],
            initialfile=f"qr_invoice_{self.invoice_number}.png"
        )
        
        if file_path:
            try:
                qr_img.save(file_path)
                messagebox.showinfo("Success", f"QR code saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save QR code: {str(e)}")

    def load_products_table(self):
        """Load products into the products table"""
        try:
            # Clear existing data
            for item in self.products_table.get_children():
                self.products_table.delete(item)
            
            # Load from database
            self.cursor.execute("SELECT id, hsn, name, price, category, last_updated FROM products")
            for row in self.cursor.fetchall():
                self.products_table.insert("", "end", values=row)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load products: {str(e)}")

    def search_products_in_db(self, event):
        """Search products in database"""
        search_term = self.product_search.get().lower()
        
        try:
            # Clear existing data
            for item in self.products_table.get_children():
                self.products_table.delete(item)
            
            # Search in database
            query = '''
                SELECT id, hsn, name, price, category, last_updated 
                FROM products 
                WHERE hsn LIKE ? OR name LIKE ? OR category LIKE ?
            '''
            self.cursor.execute(query, (f"%{search_term}%", f"%{search_term}%", f"%{search_term}%"))
            
            for row in self.cursor.fetchall():
                self.products_table.insert("", "end", values=row)
        except Exception as e:
            messagebox.showerror("Error", f"Search failed: {str(e)}")

    def add_product_dialog(self):
        """Show dialog to add a new product"""
        dialog = tk.Toplevel(self.master)
        dialog.title("Add Product")
        dialog.transient(self.master)
        dialog.grab_set()
        
        # HSN
        ttk.Label(dialog, text="HSN:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        hsn_entry = ttk.Entry(dialog)
        hsn_entry.grid(row=0, column=1, padx=5, pady=5)
        
        # Name
        ttk.Label(dialog, text="Name:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        name_entry = ttk.Entry(dialog)
        name_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # Price
        ttk.Label(dialog, text="Price:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        price_entry = ttk.Entry(dialog)
        price_entry.grid(row=2, column=1, padx=5, pady=5)
        
        # Category
        ttk.Label(dialog, text="Category:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        category_entry = ttk.Entry(dialog)
        category_entry.grid(row=3, column=1, padx=5, pady=5)
        
        def save_product():
            """Save the new product to database"""
            try:
                self.cursor.execute('''
                    INSERT INTO products (hsn, name, price, category)
                    VALUES (?, ?, ?, ?)
                ''', (
                    hsn_entry.get(),
                    name_entry.get(),
                    float(price_entry.get()),
                    category_entry.get()
                ))
                self.conn.commit()
                
                # Refresh products table
                self.load_products_table()
                
                # Add to local products list
                self.products.append({
                    "hsn": hsn_entry.get(),
                    "name": name_entry.get(),
                    "price": float(price_entry.get())
                })
                
                # Update combobox values
                self.product_id_entry["values"] = [p["hsn"] for p in self.products]
                self.product_name_entry["values"] = [p["name"] for p in self.products]
                
                dialog.destroy()
                messagebox.showinfo("Success", "Product added successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add product: {str(e)}")
        
        ttk.Button(
            dialog,
            text="Save",
            command=save_product,
            style="Accent.TButton"
        ).grid(row=4, column=0, columnspan=2, pady=10)

    def edit_product_dialog(self):
        """Show dialog to edit selected product"""
        selected = self.products_table.selection()
        if not selected:
            messagebox.showwarning("Warning", "No product selected")
            return
            
        product_id = self.products_table.item(selected[0], "values")[0]
        
        # Get product details from database
        self.cursor.execute("SELECT hsn, name, price, category FROM products WHERE id = ?", (product_id,))
        product = self.cursor.fetchone()
        
        if not product:
            messagebox.showerror("Error", "Product not found")
            return
            
        dialog = tk.Toplevel(self.master)
        dialog.title("Edit Product")
        dialog.transient(self.master)
        dialog.grab_set()
        
        # HSN
        ttk.Label(dialog, text="HSN:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        hsn_entry = ttk.Entry(dialog)
        hsn_entry.grid(row=0, column=1, padx=5, pady=5)
        hsn_entry.insert(0, product[0])
        
        # Name
        ttk.Label(dialog, text="Name:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        name_entry = ttk.Entry(dialog)
        name_entry.grid(row=1, column=1, padx=5, pady=5)
        name_entry.insert(0, product[1])
        
        # Price
        ttk.Label(dialog, text="Price:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        price_entry = ttk.Entry(dialog)
        price_entry.grid(row=2, column=1, padx=5, pady=5)
        price_entry.insert(0, product[2])
        
        # Category
        ttk.Label(dialog, text="Category:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        category_entry = ttk.Entry(dialog)
        category_entry.grid(row=3, column=1, padx=5, pady=5)
        category_entry.insert(0, product[3])
        
        def save_changes():
            """Save edited product to database"""
            try:
                self.cursor.execute('''
                    UPDATE products 
                    SET hsn = ?, name = ?, price = ?, category = ?
                    WHERE id = ?
                ''', (
                    hsn_entry.get(),
                    name_entry.get(),
                    float(price_entry.get()),
                    category_entry.get(),
                    product_id
                ))
                self.conn.commit()
                
                # Refresh products table
                self.load_products_table()
                
                # Update local products list
                for p in self.products:
                    if p["hsn"] == product[0]:
                        p["hsn"] = hsn_entry.get()
                        p["name"] = name_entry.get()
                        p["price"] = float(price_entry.get())
                        break
                
                # Update combobox values
                self.product_id_entry["values"] = [p["hsn"] for p in self.products]
                self.product_name_entry["values"] = [p["name"] for p in self.products]
                
                dialog.destroy()
                messagebox.showinfo("Success", "Product updated successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update product: {str(e)}")
        
        ttk.Button(
            dialog,
            text="Save",
            command=save_changes,
            style="Accent.TButton"
        ).grid(row=4, column=0, columnspan=2, pady=10)

    def delete_product(self):
        """Delete selected product"""
        selected = self.products_table.selection()
        if not selected:
            messagebox.showwarning("Warning", "No product selected")
            return
            
        product_id = self.products_table.item(selected[0], "values")[0]
        product_hsn = self.products_table.item(selected[0], "values")[1]
        
        if not messagebox.askyesno("Confirm", f"Delete product {product_hsn}?"):
            return
            
        try:
            self.cursor.execute("DELETE FROM products WHERE id = ?", (product_id,))
            self.conn.commit()
            
            # Refresh products table
            self.load_products_table()
            
            # Remove from local products list
            self.products = [p for p in self.products if p["hsn"] != product_hsn]
            
            # Update combobox values
            self.product_id_entry["values"] = [p["hsn"] for p in self.products]
            self.product_name_entry["values"] = [p["name"] for p in self.products]
            
            messagebox.showinfo("Success", "Product deleted successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete product: {str(e)}")

    def load_customers_table(self):
        """Load customers into the customers table"""
        try:
            # Clear existing data
            for item in self.customers_table.get_children():
                self.customers_table.delete(item)
            
            # Load from database
            self.cursor.execute("SELECT id, name, mobile, place, address, gstin, created_at FROM customers")
            for row in self.cursor.fetchall():
                self.customers_table.insert("", "end", values=row)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load customers: {str(e)}")

    def search_customers_in_db(self, event):
        """Search customers in database"""
        search_term = self.customer_search.get().lower()
        
        try:
            # Clear existing data
            for item in self.customers_table.get_children():
                self.customers_table.delete(item)
            
            # Search in database
            query = '''
                SELECT id, name, mobile, place, address, gstin, created_at 
                FROM customers 
                WHERE name LIKE ? OR mobile LIKE ? OR place LIKE ? OR gstin LIKE ?
            '''
            self.cursor.execute(query, 
                (f"%{search_term}%", f"%{search_term}%", f"%{search_term}%", f"%{search_term}%"))
            
            for row in self.cursor.fetchall():
                self.customers_table.insert("", "end", values=row)
        except Exception as e:
            messagebox.showerror("Error", f"Search failed: {str(e)}")

    def add_customer_dialog(self):
        """Show dialog to add a new customer"""
        dialog = tk.Toplevel(self.master)
        dialog.title("Add Customer")
        dialog.transient(self.master)
        dialog.grab_set()
        
        # Name
        ttk.Label(dialog, text="Name:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        name_entry = ttk.Entry(dialog)
        name_entry.grid(row=0, column=1, padx=5, pady=5)
        
        # Mobile
        ttk.Label(dialog, text="Mobile:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        mobile_entry = ttk.Entry(dialog)
        mobile_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # Place
        ttk.Label(dialog, text="Place:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        place_entry = ttk.Entry(dialog)
        place_entry.grid(row=2, column=1, padx=5, pady=5)
        
        # Address
        ttk.Label(dialog, text="Address:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        address_entry = ttk.Entry(dialog)
        address_entry.grid(row=3, column=1, padx=5, pady=5)
        
        # GSTIN
        ttk.Label(dialog, text="GSTIN:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        gstin_entry = ttk.Entry(dialog)
        gstin_entry.grid(row=4, column=1, padx=5, pady=5)
        
        def save_customer():
            """Save the new customer to database"""
            try:
                self.cursor.execute('''
                    INSERT INTO customers (name, mobile, place, address, gstin)
                    VALUES (?, ?, ?, ?, ?)
                ''', (
                    name_entry.get(),
                    mobile_entry.get(),
                    place_entry.get(),
                    address_entry.get(),
                    gstin_entry.get()
                ))
                self.conn.commit()
                
                # Refresh customers table
                self.load_customers_table()
                
                dialog.destroy()
                messagebox.showinfo("Success", "Customer added successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add customer: {str(e)}")
        
        ttk.Button(
            dialog,
            text="Save",
            command=save_customer,
            style="Accent.TButton"
        ).grid(row=5, column=0, columnspan=2, pady=10)

    def edit_customer_dialog(self):
        """Show dialog to edit selected customer"""
        selected = self.customers_table.selection()
        if not selected:
            messagebox.showwarning("Warning", "No customer selected")
            return
            
        customer_id = self.customers_table.item(selected[0], "values")[0]
        
        # Get customer details from database
        self.cursor.execute("SELECT name, mobile, place, address, gstin FROM customers WHERE id = ?", (customer_id,))
        customer = self.cursor.fetchone()
        
        if not customer:
            messagebox.showerror("Error", "Customer not found")
            return
            
        dialog = tk.Toplevel(self.master)
        dialog.title("Edit Customer")
        dialog.transient(self.master)
        dialog.grab_set()
        
        # Name
        ttk.Label(dialog, text="Name:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        name_entry = ttk.Entry(dialog)
        name_entry.grid(row=0, column=1, padx=5, pady=5)
        name_entry.insert(0, customer[0])
        
        # Mobile
        ttk.Label(dialog, text="Mobile:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        mobile_entry = ttk.Entry(dialog)
        mobile_entry.grid(row=1, column=1, padx=5, pady=5)
        mobile_entry.insert(0, customer[1])
        
        # Place
        ttk.Label(dialog, text="Place:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        place_entry = ttk.Entry(dialog)
        place_entry.grid(row=2, column=1, padx=5, pady=5)
        place_entry.insert(0, customer[2])
        
        # Address
        ttk.Label(dialog, text="Address:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        address_entry = ttk.Entry(dialog)
        address_entry.grid(row=3, column=1, padx=5, pady=5)
        address_entry.insert(0, customer[3])
        
        # GSTIN
        ttk.Label(dialog, text="GSTIN:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        gstin_entry = ttk.Entry(dialog)
        gstin_entry.grid(row=4, column=1, padx=5, pady=5)
        gstin_entry.insert(0, customer[4])
        
        def save_changes():
            """Save edited customer to database"""
            try:
                self.cursor.execute('''
                    UPDATE customers 
                    SET name = ?, mobile = ?, place = ?, address = ?, gstin = ?
                    WHERE id = ?
                ''', (
                    name_entry.get(),
                    mobile_entry.get(),
                    place_entry.get(),
                    address_entry.get(),
                    gstin_entry.get(),
                    customer_id
                ))
                self.conn.commit()
                
                # Refresh customers table
                self.load_customers_table()
                
                dialog.destroy()
                messagebox.showinfo("Success", "Customer updated successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update customer: {str(e)}")
        
        ttk.Button(
            dialog,
            text="Save",
            command=save_changes,
            style="Accent.TButton"
        ).grid(row=5, column=0, columnspan=2, pady=10)

    def delete_customer(self):
        """Delete selected customer"""
        selected = self.customers_table.selection()
        if not selected:
            messagebox.showwarning("Warning", "No customer selected")
            return
            
        customer_id = self.customers_table.item(selected[0], "values")[0]
        customer_name = self.customers_table.item(selected[0], "values")[1]
        
        if not messagebox.askyesno("Confirm", f"Delete customer {customer_name}?"):
            return
            
        try:
            self.cursor.execute("DELETE FROM customers WHERE id = ?", (customer_id,))
            self.conn.commit()
            
            # Refresh customers table
            self.load_customers_table()
            
            messagebox.showinfo("Success", "Customer deleted successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete customer: {str(e)}")

    def generate_sales_report(self):
        """Generate sales report for selected date range"""
        from_date = self.from_date.get()
        to_date = self.to_date.get()
        
        if not from_date or not to_date:
            messagebox.showwarning("Warning", "Please enter both from and to dates")
            return
            
        try:
            # Get sales data
            query = '''
                SELECT date, invoice_number, customer_name, total 
                FROM invoices 
                WHERE date BETWEEN ? AND ?
                ORDER BY date
            '''
            self.cursor.execute(query, (from_date, to_date))
            sales_data = self.cursor.fetchall()
            
            if not sales_data:
                self.report_text.delete(1.0, tk.END)
                self.report_text.insert(tk.END, "No sales data found for the selected period")
                return
                
            # Calculate totals
            total_sales = sum(row[3] for row in sales_data)
            total_invoices = len(sales_data)
            
            # Format report
            report = f"Sales Report from {from_date} to {to_date}\n"
            report += "=" * 50 + "\n\n"
            report += f"{'Date':<12}{'Invoice No':<12}{'Customer':<30}{'Amount':>10}\n"
            report += "-" * 64 + "\n"
            
            for row in sales_data:
                report += f"{row[0]:<12}{row[1]:<12}{row[2][:28]:<30}{row[3]:>10.2f}\n"
            
            report += "\n" + "=" * 50 + "\n"
            report += f"Total Invoices: {total_invoices}\n"
            report += f"Total Sales: {total_sales:.2f}\n"
            
            # Display report
            self.report_text.delete(1.0, tk.END)
            self.report_text.insert(tk.END, report)
            
            # Generate chart
            self.generate_sales_chart(sales_data)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report: {str(e)}")

    def generate_sales_chart(self, sales_data):
        """Generate sales chart from sales data"""
        try:
            # Clear previous chart
            self.sales_ax.clear()
            
            # Prepare data
            dates = [row[0] for row in sales_data]
            amounts = [row[3] for row in sales_data]
            
            # Create bar chart
            self.sales_ax.bar(dates, amounts, color=self.config["secondary_color"])
            self.sales_ax.set_title("Sales by Date")
            self.sales_ax.set_xlabel("Date")
            self.sales_ax.set_ylabel("Amount")
            
            # Rotate x-axis labels for better readability
            for label in self.sales_ax.get_xticklabels():
                label.set_rotation(45)
                label.set_ha('right')
            
            # Redraw canvas
            self.sales_canvas.draw()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate chart: {str(e)}")

    def generate_product_report(self):
        """Generate product sales report"""
        try:
            # Get product sales data
            query = '''
                SELECT p.hsn, p.name, SUM(ii.quantity) as total_quantity, 
                       SUM(ii.total) as total_sales
                FROM products p
                JOIN invoice_items ii ON p.hsn = ii.hsn
                GROUP BY p.hsn, p.name
                ORDER BY total_sales DESC
            '''
            self.cursor.execute(query)
            product_data = self.cursor.fetchall()
            
            if not product_data:
                self.product_report_text.delete(1.0, tk.END)
                self.product_report_text.insert(tk.END, "No product sales data found")
                return
                
            # Calculate totals
            total_quantity = sum(row[2] for row in product_data)
            total_sales = sum(row[3] for row in product_data)
            
            # Format report
            report = "Product Sales Report\n"
            report += "=" * 50 + "\n\n"
            report += f"{'HSN':<10}{'Product Name':<30}{'Qty Sold':>10}{'Total Sales':>15}\n"
            report += "-" * 65 + "\n"
            
            for row in product_data:
                report += f"{row[0]:<10}{row[1][:28]:<30}{row[2]:>10}{row[3]:>15.2f}\n"
            
            report += "\n" + "=" * 50 + "\n"
            report += f"Total Quantity Sold: {total_quantity}\n"
            report += f"Total Sales: {total_sales:.2f}\n"
            
            # Display report
            self.product_report_text.delete(1.0, tk.END)
            self.product_report_text.insert(tk.END, report)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report: {str(e)}")

    def change_theme(self, theme_name):
        """Change application theme"""
        self.current_theme = theme_name
        self.apply_styles()
        self.config["default_theme"] = theme_name
        self.save_config()

    def refresh_data(self):
        """Refresh all data views"""
        current_tab = self.notebook.tab(self.notebook.select(), "text")
        
        if current_tab == "Products":
            self.load_products_table()
        elif current_tab == "Customers":
            self.load_customers_table()
        elif current_tab == "Reports":
            self.generate_sales_report()
            self.generate_product_report()
        
        self.status_label.config(text="Data refreshed")

    def on_tab_changed(self, event):
        """Handle tab change event"""
        tab = self.notebook.tab(self.notebook.select(), "text")
        
        if tab == "Reports":
            # Generate reports when tab is selected
            self.generate_sales_report()
            self.generate_product_report()
        elif tab == "Products":
            self.load_products_table()
        elif tab == "Customers":
            self.load_customers_table()

    def on_exit(self):
        """Handle application exit"""
        # Cancel auto-save timer if running
        if hasattr(self, 'auto_save_timer'):
            self.auto_save_timer.cancel()
        
        # Close database connection
        self.conn.close()
        
        # Close the application
        self.master.quit()

if __name__ == "__main__":
    root = tk.Tk()
    
    # Set window icon
    try:
        if getattr(sys, 'frozen', False):
            icon_path = os.path.join(sys._MEIPASS, "icon.ico")
        else:
            icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
            
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except: 
        pass
    
    app = BillingSystem(root)
    root.protocol("WM_DELETE_WINDOW", app.on_exit)
    root.mainloop() 