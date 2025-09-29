import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import pandas as pd
import os
from datetime import datetime, timedelta
import calendar

STOCK_FILE = "stock1.xlsx"
ORDERS_FILE = "orders1.xlsx"


# Ensure Excel files exist
# Ensure Excel files exist with proper date formatting
# Ensure Excel files exist with proper date formatting
def init_files():
    if not os.path.exists(STOCK_FILE):
        df = pd.DataFrame(
            columns=["piece_id", "product_name", "length_m", "date_added", "seller_price", "unit_cost", "profit",
                     "status", "sold_date", "order_id"])
        # Save with explicit datetime format
        try:
            with pd.ExcelWriter(STOCK_FILE, engine='openpyxl', datetime_format='YYYY-MM-DD HH:MM:SS') as writer:
                df.to_excel(writer, index=False)
        except:
            # Fallback if openpyxl not available
            df.to_excel(STOCK_FILE, index=False)

    if not os.path.exists(ORDERS_FILE):
        df = pd.DataFrame(
            columns=["order_id", "order_date", "customer_name", "address", "phone1", "phone2", "city", "item_name",
                     "length_m", "qty", "total_unit_cost", "total_seller_price", "profit_total", "allocated_piece_ids",
                     "status"])
        # Save with explicit datetime format
        try:
            with pd.ExcelWriter(ORDERS_FILE, engine='openpyxl', datetime_format='YYYY-MM-DD HH:MM:SS') as writer:
                df.to_excel(writer, index=False)
        except:
            # Fallback if openpyxl not available
            df.to_excel(ORDERS_FILE, index=False)


# Load product names from stock
def get_product_names():
    if os.path.exists(STOCK_FILE):
        df = pd.read_excel(STOCK_FILE)
        if "product_name" in df.columns:
            return sorted(df["product_name"].dropna().unique().tolist())
    return []


class StockApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("LED Strip Stock & Orders App")
        self.geometry("1400x800")

        # Initialize product_names here
        self.product_names = get_product_names()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.stock_tab = ttk.Frame(self.notebook)
        self.orders_tab = ttk.Frame(self.notebook)
        self.summary_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.stock_tab, text="Stock Management")
        self.notebook.add(self.orders_tab, text="Order Management")
        self.notebook.add(self.summary_tab, text="Summary Report")

        self.create_stock_tab()
        self.create_orders_tab()
        self.create_summary_tab()

        # Auto-refresh product names in comboboxes
        self.update_product_comboboxes()

        # Bind tab change event to refresh summary
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

    def on_tab_change(self, event):
        """Refresh summary tab when it is selected"""
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 2:  # Summary tab is index 2
            self.update_summary()

    def update_product_comboboxes(self):
        # Refresh product names from file
        self.product_names = get_product_names()

        # Update comboboxes only if they exist
        if hasattr(self, 'product_cb'):
            self.product_cb['values'] = self.product_names

        if hasattr(self, 'order_product_cb'):
            self.order_product_cb['values'] = self.product_names

        # Update filter comboboxes too - check if they exist first
        if hasattr(self, 'filter_product_cb'):
            self.filter_product_cb['values'] = self.product_names

        if hasattr(self, 'order_filter_product_cb'):
            self.order_filter_product_cb['values'] = self.product_names

        if hasattr(self, 'summary_product_cb'):
            self.summary_product_cb['values'] = self.product_names

    def create_stock_tab(self):
        # Main frame with top-bottom split
        main_frame = ttk.Frame(self.stock_tab)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Top frame for Add Stock and Filters
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill="x", pady=5)

        # Left frame for Add Stock
        left_frame = ttk.LabelFrame(top_frame, text="Add New Stock")
        left_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        # Right frame for Filters and Removal
        right_frame = ttk.LabelFrame(top_frame, text="Stock Filters & Removal")
        right_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)

        # Add Stock Frame (Left side)
        add_frame = ttk.Frame(left_frame)
        add_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Product input with auto-suggestion
        ttk.Label(add_frame, text="Product Name:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.product_var = tk.StringVar()
        self.product_cb = ttk.Combobox(add_frame, textvariable=self.product_var, width=30)
        self.product_cb.grid(row=0, column=1, padx=5, pady=5)
        self.product_cb.bind('<KeyRelease>', self.update_product_suggestions)
        self.product_cb.bind('<Return>', lambda e: self.length_cb.focus())

        ttk.Label(add_frame, text="Length (m):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.length_var = tk.StringVar()
        self.length_cb = ttk.Combobox(add_frame, textvariable=self.length_var, values=[5, 10, 15, 20, 30], width=15)
        self.length_cb.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.length_cb.bind('<Return>', lambda e: self.pcs_entry.focus())

        ttk.Label(add_frame, text="PCS:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.pcs_entry = ttk.Entry(add_frame, width=15)
        self.pcs_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.pcs_entry.bind('<Return>', lambda e: self.unit_cost_entry.focus())

        ttk.Label(add_frame, text="Unit Cost:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.unit_cost_entry = ttk.Entry(add_frame, width=15)
        self.unit_cost_entry.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        self.unit_cost_entry.bind('<Return>', lambda e: self.seller_price_entry.focus())

        ttk.Label(add_frame, text="Seller Price:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.seller_price_entry = ttk.Entry(add_frame, width=15)
        self.seller_price_entry.grid(row=1, column=3, padx=5, pady=5, sticky="w")
        self.seller_price_entry.bind('<Return>', lambda e: self.add_stock())

        # Clear form button (NEW)
        clear_btn = ttk.Button(add_frame, text="Clear Form", command=self.clear_stock_form)
        clear_btn.grid(row=2, column=2, padx=5, pady=5, sticky="e")

        add_btn = ttk.Button(add_frame, text="Add to Stock", command=self.add_stock)
        add_btn.grid(row=2, column=3, padx=5, pady=5, sticky="e")

        # Filter and Remove Frame (Right side)
        filter_remove_frame = ttk.Frame(right_frame)
        filter_remove_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Filter Frame
        filter_frame = ttk.Frame(filter_remove_frame)
        filter_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(filter_frame, text="Product:").grid(row=0, column=0, padx=5, pady=5)
        self.filter_product_var = tk.StringVar()
        self.filter_product_cb = ttk.Combobox(filter_frame, textvariable=self.filter_product_var, width=20)
        self.filter_product_cb['values'] = self.product_names
        self.filter_product_cb.grid(row=0, column=1, padx=5, pady=5)
        self.filter_product_cb.bind('<<ComboboxSelected>>', lambda e: self.apply_stock_filters())

        ttk.Label(filter_frame, text="Length:").grid(row=0, column=2, padx=5, pady=5)
        self.filter_length_var = tk.StringVar()
        self.filter_length_cb = ttk.Combobox(filter_frame, textvariable=self.filter_length_var,
                                             values=["All", "5", "10", "15", "20", "30"], width=10)
        self.filter_length_cb.set("All")
        self.filter_length_cb.grid(row=0, column=3, padx=5, pady=5)
        self.filter_length_cb.bind('<<ComboboxSelected>>', lambda e: self.apply_stock_filters())

        ttk.Label(filter_frame, text="Status:").grid(row=0, column=4, padx=5, pady=5)
        self.filter_status_var = tk.StringVar()
        self.filter_status_cb = ttk.Combobox(filter_frame, textvariable=self.filter_status_var,
                                             values=["All", "IN_STOCK", "SOLD", "REMOVED"], width=12)
        self.filter_status_cb.set("IN_STOCK")
        self.filter_status_cb.grid(row=0, column=5, padx=5, pady=5)
        self.filter_status_cb.bind('<<ComboboxSelected>>', lambda e: self.apply_stock_filters())

        # Date Range Filter (NEW)
        ttk.Label(filter_frame, text="Date Range:").grid(row=1, column=0, padx=5, pady=5)
        self.filter_date_var = tk.StringVar()
        self.filter_date_cb = ttk.Combobox(filter_frame, textvariable=self.filter_date_var,
                                           values=["All", "Today", "Last 7 Days", "This Month", "Last Month",
                                                   "Last 3 Months", "Last 6 Months", "Last 12 Months"], width=15)
        self.filter_date_cb.set("All")
        self.filter_date_cb.grid(row=1, column=1, padx=5, pady=5)
        self.filter_date_cb.bind('<<ComboboxSelected>>', lambda e: self.apply_stock_filters())

        clear_filter_btn = ttk.Button(filter_frame, text="Clear Filters", command=self.clear_stock_filters)
        clear_filter_btn.grid(row=1, column=6, padx=5, pady=5)

        # Remove Stock Frame
        remove_frame = ttk.Frame(filter_remove_frame)
        remove_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(remove_frame, text="Select items to remove:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        remove_btn = ttk.Button(remove_frame, text="Remove Selected", command=self.remove_selected_stock)
        remove_btn.grid(row=0, column=1, padx=5, pady=5, sticky="e")

        # Bottom frame for Stock Preview (LARGE TABLE)
        bottom_frame = ttk.LabelFrame(main_frame, text="Stock Preview")
        bottom_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Stock table view with scrollbar - LARGER SIZE
        table_frame = ttk.Frame(bottom_frame)
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Add scrollbar to stock table
        scrollbar = ttk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.stock_table = ttk.Treeview(table_frame,
                                        columns=("piece_id", "product_name", "length_m", "date_added", "seller_price",
                                                 "unit_cost", "profit", "status"),
                                        show="headings", height=20, yscrollcommand=scrollbar.set)  # Increased height
        scrollbar.config(command=self.stock_table.yview)

        # Define columns
        columns = [("piece_id", "Piece ID", 120),
                   ("product_name", "Product Name", 150),
                   ("length_m", "Length (m)", 80),
                   ("date_added", "Date Added", 120),
                   ("seller_price", "Sell Price", 80),
                   ("unit_cost", "Cost", 80),
                   ("profit", "Profit", 80),
                   ("status", "Status", 100)]

        for col_id, heading, width in columns:
            self.stock_table.heading(col_id, text=heading)
            self.stock_table.column(col_id, width=width, anchor="center")

        self.stock_table.pack(fill="both", expand=True)

        # Summary frame for totals (NEW)
        summary_frame = ttk.Frame(bottom_frame)
        summary_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(summary_frame, text="Total Sell Price:", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=10,
                                                                                            pady=5, sticky="e")
        self.total_sell_price_var = tk.StringVar(value="Rs. 0.00")
        ttk.Label(summary_frame, textvariable=self.total_sell_price_var, font=("Arial", 10)).grid(row=0, column=1,
                                                                                                  padx=10, pady=5,
                                                                                                  sticky="w")

        ttk.Label(summary_frame, text="Total Cost:", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=10, pady=5,
                                                                                      sticky="e")
        self.total_cost_var = tk.StringVar(value="Rs. 0.00")
        ttk.Label(summary_frame, textvariable=self.total_cost_var, font=("Arial", 10)).grid(row=0, column=3, padx=10,
                                                                                            pady=5, sticky="w")

        ttk.Label(summary_frame, text="Total Profit:", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=10,
                                                                                        pady=5, sticky="e")
        self.total_profit_var = tk.StringVar(value="Rs. 0.00")
        ttk.Label(summary_frame, textvariable=self.total_profit_var, font=("Arial", 10)).grid(row=0, column=5, padx=10,
                                                                                              pady=5, sticky="w")

        ttk.Label(summary_frame, text="Total Quantity:", font=("Arial", 10, "bold")).grid(row=1, column=0, padx=10,
                                                                                          pady=5, sticky="e")
        self.total_quantity_var = tk.StringVar(value="0")
        ttk.Label(summary_frame, textvariable=self.total_quantity_var, font=("Arial", 10)).grid(row=1, column=1,
                                                                                                padx=10, pady=5,
                                                                                                sticky="w")

        ttk.Label(summary_frame, text="Profit Percentage:", font=("Arial", 10, "bold")).grid(row=1, column=2, padx=10,
                                                                                             pady=5, sticky="e")
        self.profit_percentage_var = tk.StringVar(value="0.00%")
        ttk.Label(summary_frame, textvariable=self.profit_percentage_var, font=("Arial", 10)).grid(row=1, column=3,
                                                                                                   padx=10, pady=5,
                                                                                                   sticky="w")

        # Load initial stock data
        self.load_stock()

    def clear_stock_form(self):
        """Clear all input fields in the stock form"""
        self.product_var.set("")
        self.length_var.set("")
        self.pcs_entry.delete(0, tk.END)
        self.unit_cost_entry.delete(0, tk.END)
        self.seller_price_entry.delete(0, tk.END)

    def update_product_suggestions(self, event):
        typed = self.product_var.get().lower()
        if typed == '':
            self.product_cb['values'] = self.product_names
        else:
            matches = [name for name in self.product_names if typed in name.lower()]
            self.product_cb['values'] = matches

    def apply_stock_filters(self):
        product_filter = self.filter_product_var.get()
        length_filter = self.filter_length_var.get()
        status_filter = self.filter_status_var.get()
        date_filter = self.filter_date_var.get()

        df = pd.read_excel(STOCK_FILE)

        # Apply filters
        if product_filter and product_filter != "All":
            df = df[df["product_name"] == product_filter]

        if length_filter and length_filter != "All":
            df = df[df["length_m"] == int(length_filter)]

        if status_filter and status_filter != "All":
            df = df[df["status"] == status_filter]

        # Apply date filter
        if date_filter and date_filter != "All":
            try:
                # Convert dates safely
                df["date_added"] = pd.to_datetime(df["date_added"], errors='coerce', format='mixed')
                today = datetime.now().date()

                # Drop rows with invalid dates
                df = df.dropna(subset=['date_added'])

                if date_filter == "Today":
                    df = df[df["date_added"].dt.date == today]
                elif date_filter == "Last 7 Days":
                    seven_days_ago = today - timedelta(days=7)
                    df = df[df["date_added"].dt.date >= seven_days_ago]
                elif date_filter == "This Month":
                    df = df[df["date_added"].dt.month == today.month]
                elif date_filter == "Last Month":
                    last_month = today.month - 1 if today.month > 1 else 12
                    df = df[df["date_added"].dt.month == last_month]
                elif date_filter == "Last 3 Months":
                    three_months_ago = today - timedelta(days=90)
                    df = df[df["date_added"].dt.date >= three_months_ago]
                elif date_filter == "Last 6 Months":
                    six_months_ago = today - timedelta(days=180)
                    df = df[df["date_added"].dt.date >= six_months_ago]
                elif date_filter == "Last 12 Months":
                    twelve_months_ago = today - timedelta(days=365)
                    df = df[df["date_added"].dt.date >= twelve_months_ago]
            except Exception as e:
                messagebox.showerror("Error", f"Date filter error: {str(e)}")
                return

        # Display filtered results
        for row in self.stock_table.get_children():
            self.stock_table.delete(row)

        for _, row in df.iterrows():
            self.stock_table.insert("", "end", values=(
                row["piece_id"], row["product_name"], row["length_m"],
                row["date_added"], row["seller_price"], row["unit_cost"],
                row["profit"], row["status"]
            ))

        # Update totals
        self.update_stock_totals(df)

    def update_stock_totals(self, df):
        """Update the total sell price, cost and profit for displayed items"""
        total_sell_price = df["seller_price"].sum()
        total_cost = df["unit_cost"].sum()
        total_profit = df["profit"].sum()
        total_quantity = len(df)

        # Calculate profit percentage
        profit_percentage = (total_profit / total_cost * 100) if total_cost > 0 else 0

        self.total_sell_price_var.set(f"Rs. {total_sell_price:,.2f}")
        self.total_cost_var.set(f"Rs. {total_cost:,.2f}")
        self.total_profit_var.set(f"Rs. {total_profit:,.2f}")
        self.total_quantity_var.set(f"{total_quantity:,}")
        self.profit_percentage_var.set(f"{profit_percentage:.2f}%")

    def clear_stock_filters(self):
        self.filter_product_var.set("")
        self.filter_length_var.set("All")
        self.filter_status_var.set("IN_STOCK")
        self.filter_date_var.set("All")
        self.load_stock()

    def add_stock(self):
        product = self.product_var.get().strip()
        length = self.length_var.get().strip()
        pcs = self.pcs_entry.get().strip()
        unit_cost = self.unit_cost_entry.get().strip()
        seller_price = self.seller_price_entry.get().strip()

        if not product or not length or not pcs or not unit_cost or not seller_price:
            messagebox.showerror("Error", "All fields are required")
            return

        try:
            pcs = int(pcs)
            unit_cost = float(unit_cost)
            seller_price = float(seller_price)
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numbers")
            return

        df = pd.read_excel(STOCK_FILE)
        date_added = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows = []
        for i in range(pcs):
            piece_id = f"{product}_{length}m_{len(df) + i + 1}"
            profit = seller_price - unit_cost
            rows.append(
                [piece_id, product, length, date_added, seller_price, unit_cost, profit, "IN_STOCK", None, None])

        new_df = pd.DataFrame(rows, columns=df.columns)
        df = pd.concat([df, new_df], ignore_index=True)
        df.to_excel(STOCK_FILE, index=False)
        messagebox.showinfo("Success", f"Added {pcs} pcs of {product} ({length}m)")

        # Clear input fields
        self.clear_stock_form()

        # Refresh data
        self.load_stock()
        self.update_product_comboboxes()

        # Auto-update summary if it's visible
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 2:  # Summary tab is index 2
            self.update_summary()

            # Save with proper date formatting
        with pd.ExcelWriter(STOCK_FILE, engine='openpyxl', datetime_format='YYYY-MM-DD HH:MM:SS') as writer:
            df.to_excel(writer, index=False)

    def remove_selected_stock(self):
        selected_items = self.stock_table.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select items to remove")
            return

        result = messagebox.askyesno("Confirm", f"Are you sure you want to remove {len(selected_items)} item(s)?")
        if not result:
            return

        df = pd.read_excel(STOCK_FILE)
        pieces_to_remove = []

        for item in selected_items:
            piece_id = self.stock_table.item(item)['values'][0]
            pieces_to_remove.append(piece_id)

        # Mark as REMOVED instead of actually deleting
        df.loc[df['piece_id'].isin(pieces_to_remove), 'status'] = 'REMOVED'
        df.to_excel(STOCK_FILE, index=False)

        messagebox.showinfo("Success", f"Removed {len(selected_items)} item(s)")
        self.load_stock()

        # Auto-update summary if it's visible
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 2:  # Summary tab is index 2
            self.update_summary()

    def load_stock(self):
        for row in self.stock_table.get_children():
            self.stock_table.delete(row)

        # Read with explicit date formatting
        df = pd.read_excel(STOCK_FILE)

        # Convert date columns with proper formatting
        date_columns = ['date_added', 'sold_date']
        for col in date_columns:
            if col in df.columns:
                # First try to convert with specific format, then fallback to coerce
                df[col] = pd.to_datetime(df[col], format='%Y-%m-%d %H:%M:%S', errors='coerce')

        # Apply current filter
        status_filter = self.filter_status_var.get() if hasattr(self, 'filter_status_var') else "IN_STOCK"
        if status_filter and status_filter != "All":
            df = df[df['status'] == status_filter]
        else:
            df = df[df['status'] == 'IN_STOCK']

        for _, row in df.iterrows():
            # Format dates for display
            date_added = row["date_added"].strftime("%Y-%m-%d %H:%M") if pd.notna(row["date_added"]) else "N/A"
            sold_date = row["sold_date"].strftime("%Y-%m-%d %H:%M") if pd.notna(row["sold_date"]) else "N/A"

            self.stock_table.insert("", "end", values=(
                row["piece_id"], row["product_name"], row["length_m"],
                date_added, row["seller_price"], row["unit_cost"],
                row["profit"], row["status"]
            ))

        # Update totals
        self.update_stock_totals(df)

    def create_orders_tab(self):
        # Main frame
        main_frame = ttk.Frame(self.orders_tab)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Top frame for Place New Order (NEW LAYOUT)
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill="x", pady=5)

        # Left frame for customer details (2 columns)
        customer_frame = ttk.LabelFrame(top_frame, text="Customer Details")
        customer_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        # Customer details in 2 columns
        cust_col1 = ttk.Frame(customer_frame)
        cust_col1.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        cust_col2 = ttk.Frame(customer_frame)
        cust_col2.pack(side="right", fill="both", expand=True, padx=5, pady=5)

        ttk.Label(cust_col1, text="Customer Name:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.cust_name = ttk.Entry(cust_col1, width=25)
        self.cust_name.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.cust_name.bind('<Return>', lambda e: self.cust_address.focus())

        ttk.Label(cust_col1, text="Address:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.cust_address = ttk.Entry(cust_col1, width=25)
        self.cust_address.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.cust_address.bind('<Return>', lambda e: self.cust_phone1.focus())

        ttk.Label(cust_col1, text="Phone 1:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.cust_phone1 = ttk.Entry(cust_col1, width=25)
        self.cust_phone1.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.cust_phone1.bind('<Return>', lambda e: self.cust_phone2.focus())

        ttk.Label(cust_col2, text="Phone 2:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.cust_phone2 = ttk.Entry(cust_col2, width=25)
        self.cust_phone2.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.cust_phone2.bind('<Return>', lambda e: self.cust_city.focus())

        ttk.Label(cust_col2, text="City:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.cust_city = ttk.Entry(cust_col2, width=25)
        self.cust_city.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.cust_city.bind('<Return>', lambda e: self.order_product_cb.focus())

        # Clear form button for customer details (NEW)
        clear_cust_btn = ttk.Button(cust_col2, text="Clear Form", command=self.clear_customer_form)
        clear_cust_btn.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        # Right frame for order items
        items_frame = ttk.LabelFrame(top_frame, text="Order Items")
        items_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)

        # Item entry
        item_entry_frame = ttk.Frame(items_frame)
        item_entry_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(item_entry_frame, text="Item Name:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.order_product_var = tk.StringVar()
        self.order_product_cb = ttk.Combobox(item_entry_frame, textvariable=self.order_product_var, width=20)
        self.order_product_cb.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.order_product_cb.bind('<KeyRelease>', self.update_order_product_suggestions)
        self.order_product_cb.bind('<Return>', lambda e: self.order_length_cb.focus())

        ttk.Label(item_entry_frame, text="Length (m):").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.order_length_var = tk.StringVar()
        self.order_length_cb = ttk.Combobox(item_entry_frame, textvariable=self.order_length_var,
                                            values=[5, 10, 15, 20, 30],
                                            width=10)
        self.order_length_cb.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        self.order_length_cb.bind("<<ComboboxSelected>>", self.update_availability)
        self.order_length_cb.bind('<Return>', lambda e: self.qty_entry.focus())

        self.availability_lbl = ttk.Label(item_entry_frame, text="Available: 0")
        self.availability_lbl.grid(row=0, column=4, padx=5, pady=5)

        ttk.Label(item_entry_frame, text="Qty:").grid(row=0, column=5, padx=5, pady=5, sticky="e")
        self.qty_var = tk.IntVar(value=1)
        self.qty_entry = ttk.Entry(item_entry_frame, textvariable=self.qty_var, width=10)
        self.qty_entry.grid(row=0, column=6, padx=5, pady=5, sticky="w")

        # Add trace to update availability when quantity changes - FIX FOR REAL-TIME UPDATE
        self.qty_var.trace('w', self.on_qty_change)

        self.qty_entry.bind('<Return>', lambda e: self.add_order_item())

        add_item_btn = ttk.Button(item_entry_frame, text="Add Item", command=self.add_order_item)
        add_item_btn.grid(row=0, column=7, padx=5, pady=5)

        # Order items table - WITHOUT REMOVE COLUMN
        items_table_frame = ttk.Frame(items_frame)
        items_table_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Scrollbar for items table
        items_scrollbar = ttk.Scrollbar(items_table_frame)
        items_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Remove the "remove" column from the table
        self.order_items_table = ttk.Treeview(items_table_frame,
                                              columns=("product", "length", "qty", "price", "total"),
                                              show="headings", height=5, yscrollcommand=items_scrollbar.set)
        items_scrollbar.config(command=self.order_items_table.yview)

        # Define columns for order items - WITHOUT THE REMOVE COLUMN
        item_columns = [("product", "Product", 150),
                        ("length", "Length", 80),
                        ("qty", "Qty", 60),
                        ("price", "Price", 80),
                        ("total", "Total", 100)]

        for col_id, heading, width in item_columns:
            self.order_items_table.heading(col_id, text=heading)
            self.order_items_table.column(col_id, width=width, anchor="center")

        self.order_items_table.pack(fill="both", expand=True)

        # Order total - MAKE IT LARGER
        total_frame = ttk.Frame(items_frame)
        total_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(total_frame, text="Order Total:", font=("Arial", 12, "bold")).grid(row=0, column=0, padx=5, pady=5,
                                                                                     sticky="e")
        self.order_total_var = tk.StringVar(value="Rs. 0.00")
        ttk.Label(total_frame, textvariable=self.order_total_var, font=("Arial", 12, "bold")).grid(row=0, column=1,
                                                                                                   padx=5, pady=5,
                                                                                                   sticky="w")

        # Remove Added Item button - BELOW THE ORDER TOTAL
        remove_item_btn = ttk.Button(total_frame, text="Remove Selected Item", command=self.remove_selected_order_item)
        remove_item_btn.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        # Place order button
        order_btn_frame = ttk.Frame(items_frame)
        order_btn_frame.pack(fill="x", padx=5, pady=5)

        order_btn = ttk.Button(order_btn_frame, text="Place Order", command=self.place_order)
        order_btn.pack(side="right", padx=5, pady=5)

        # Bottom frame for Order History & Filters (FULL WIDTH)
        history_frame = ttk.LabelFrame(main_frame, text="Order History & Filters")
        history_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Order Filters with Return button on top right
        filter_top_frame = ttk.Frame(history_frame)
        filter_top_frame.pack(fill="x", padx=10, pady=5)

        # Filters on left
        filter_left_frame = ttk.Frame(filter_top_frame)
        filter_left_frame.pack(side="left", fill="x", expand=True)

        # Return button on right (NEW POSITION with new name)
        return_btn = ttk.Button(filter_top_frame, text="Cancel or Return Selected Order",
                                command=self.cancel_or_return_order)
        return_btn.pack(side="right", padx=5, pady=5)

        # Filter controls
        order_filter_frame = ttk.Frame(filter_left_frame)
        order_filter_frame.pack(fill="x", pady=5)

        ttk.Label(order_filter_frame, text="Customer:").grid(row=0, column=0, padx=5, pady=5)
        self.order_filter_customer_var = tk.StringVar()
        self.order_filter_customer_entry = ttk.Entry(order_filter_frame, textvariable=self.order_filter_customer_var,
                                                     width=20)
        self.order_filter_customer_entry.grid(row=0, column=1, padx=5, pady=5)
        self.order_filter_customer_entry.bind('<KeyRelease>', lambda e: self.apply_order_filters())

        ttk.Label(order_filter_frame, text="Product:").grid(row=0, column=2, padx=5, pady=5)
        self.order_filter_product_var = tk.StringVar()
        self.order_filter_product_cb = ttk.Combobox(order_filter_frame, textvariable=self.order_filter_product_var,
                                                    width=15)
        self.order_filter_product_cb['values'] = self.product_names
        self.order_filter_product_cb.grid(row=0, column=3, padx=5, pady=5)
        self.order_filter_product_cb.bind('<<ComboboxSelected>>', lambda e: self.apply_order_filters())

        ttk.Label(order_filter_frame, text="Date Range:").grid(row=0, column=4, padx=5, pady=5)
        self.order_filter_date_var = tk.StringVar()
        self.order_filter_date_cb = ttk.Combobox(order_filter_frame, textvariable=self.order_filter_date_var,
                                                 values=["All", "Today", "Last 7 Days", "This Month", "Last Month",
                                                         "Last 3 Months", "Last 6 Months", "Last 12 Months"],
                                                 width=15)
        self.order_filter_date_cb.set("All")
        self.order_filter_date_cb.grid(row=0, column=5, padx=5, pady=5)
        self.order_filter_date_cb.bind('<<ComboboxSelected>>', lambda e: self.apply_order_filters())

        ttk.Label(order_filter_frame, text="Status:").grid(row=0, column=6, padx=5, pady=5)
        self.order_filter_status_var = tk.StringVar()
        self.order_filter_status_cb = ttk.Combobox(order_filter_frame, textvariable=self.order_filter_status_var,
                                                   values=["All", "ACTIVE", "CANCELLED", "RETURNED"], width=12)
        self.order_filter_status_cb.set("ACTIVE")
        self.order_filter_status_cb.grid(row=0, column=7, padx=5, pady=5)
        self.order_filter_status_cb.bind('<<ComboboxSelected>>', lambda e: self.apply_order_filters())

        clear_order_filter_btn = ttk.Button(order_filter_frame, text="Clear Filters", command=self.clear_order_filters)
        clear_order_filter_btn.grid(row=0, column=8, padx=5, pady=5)

        # Orders history with scrollbar
        table_frame = ttk.Frame(history_frame)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        scrollbar = ttk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.orders_table = ttk.Treeview(table_frame,
                                         columns=("order_id", "order_date", "customer_name", "city", "item_name",
                                                  "length_m", "qty", "total_seller_price", "status"),
                                         show="headings", height=15, yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.orders_table.yview)

        # Define columns
        columns = [("order_id", "Order ID", 120),
                   ("order_date", "Order Date", 120),
                   ("customer_name", "Customer", 150),
                   ("city", "City", 100),
                   ("item_name", "Item", 150),
                   ("length_m", "Length", 80),
                   ("qty", "Qty", 60),
                   ("total_seller_price", "Total", 100),
                   ("status", "Status", 80)]

        for col_id, heading, width in columns:
            self.orders_table.heading(col_id, text=heading)
            self.orders_table.column(col_id, width=width, anchor="center")

        self.orders_table.pack(fill="both", expand=True)

        # Order summary frame (NEW)
        order_summary_frame = ttk.Frame(history_frame)
        order_summary_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(order_summary_frame, text="Total Price:", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=10,
                                                                                             pady=5, sticky="e")
        self.orders_total_price_var = tk.StringVar(value="Rs. 0.00")
        ttk.Label(order_summary_frame, textvariable=self.orders_total_price_var, font=("Arial", 10)).grid(row=0,
                                                                                                          column=1,
                                                                                                          padx=10,
                                                                                                          pady=5,
                                                                                                          sticky="w")

        ttk.Label(order_summary_frame, text="Total Pcs:", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=10,
                                                                                           pady=5, sticky="e")
        self.orders_total_pcs_var = tk.StringVar(value="0")
        ttk.Label(order_summary_frame, textvariable=self.orders_total_pcs_var, font=("Arial", 10)).grid(row=0, column=3,
                                                                                                        padx=10, pady=5,
                                                                                                        sticky="w")

        # Initialize order items list
        self.order_items = []

        # Load initial orders
        self.load_orders()

        # Bind double click to edit price
        self.order_items_table.bind("<Double-1>", self.edit_order_item_price)

    def edit_order_item_price(self, event):
        """Edit the price of an order item by double-clicking on it"""
        item = self.order_items_table.identify_row(event.y)
        column = self.order_items_table.identify_column(event.x)

        if not item or column != "#4":  # Only allow editing in the Price column (column 4)
            return

        # Get current values
        current_values = self.order_items_table.item(item, "values")

        # Check if we have enough values
        if len(current_values) < 4:
            return

        try:
            current_price = float(current_values[3].replace("Rs. ", "").replace(",", ""))
        except ValueError:
            current_price = 0.0

        # Create a popup window for editing
        popup = tk.Toplevel(self)
        popup.title("Edit Price")
        popup.geometry("300x150")
        popup.transient(self)
        popup.grab_set()

        ttk.Label(popup, text=f"Edit Price for {current_values[0]} ({current_values[1]}m):").pack(pady=10)

        new_price_var = tk.StringVar(value=str(current_price))
        price_entry = ttk.Entry(popup, textvariable=new_price_var, width=15)
        price_entry.pack(pady=5)
        price_entry.focus()

        def save_new_price():
            try:
                new_price = float(new_price_var.get())
                if new_price <= 0:
                    raise ValueError

                # Update the item in the order_items list
                item_index = self.order_items_table.index(item)
                if 0 <= item_index < len(self.order_items):
                    self.order_items[item_index]["price"] = new_price
                    self.order_items[item_index]["total"] = new_price * self.order_items[item_index]["qty"]

                    # Update the table
                    self.order_items_table.item(item, values=(
                        self.order_items[item_index]["product"],
                        self.order_items[item_index]["length"],
                        self.order_items[item_index]["qty"],
                        f"Rs. {new_price:,.2f}",
                        f"Rs. {self.order_items[item_index]['total']:,.2f}",
                        "❌"
                    ))

                    # Update the order total
                    self.update_order_total()

                popup.destroy()
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid price")

        ttk.Button(popup, text="Save", command=save_new_price).pack(pady=10)
        price_entry.bind('<Return>', lambda e: save_new_price())

    def on_qty_change(self, *args):
        """Update availability when quantity changes"""
        try:
            # Only update if we have a valid quantity
            if self.qty_var.get() >= 0:
                self.update_availability()
        except tk.TclError:
            # Handle invalid input (non-numeric)
            pass

    def clear_customer_form(self):
        """Clear all customer input fields"""
        self.cust_name.delete(0, tk.END)
        self.cust_address.delete(0, tk.END)
        self.cust_phone1.delete(0, tk.END)
        self.cust_phone2.delete(0, tk.END)
        self.cust_city.delete(0, tk.END)

    def cancel_or_return_order(self):
        selected_item = self.orders_table.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select an order to cancel or return")
            return

        order_id = self.orders_table.item(selected_item[0])['values'][0]
        order_status = self.orders_table.item(selected_item[0])['values'][8] if len(
            self.orders_table.item(selected_item[0])['values']) > 8 else "ACTIVE"

        if order_status in ["CANCELLED", "RETURNED"]:
            messagebox.showwarning("Warning", f"This order has already been {order_status.lower()}")
            return

        # Ask for action type
        action = messagebox.askquestion("Select Action",
                                        f"Do you want to CANCEL order {order_id}?\n\nClick 'Yes' to CANCEL\nClick 'No' to RETURN",
                                        icon='question')

        if action == 'yes':
            # Cancel order
            confirm = messagebox.askyesno("Confirm Cancellation",
                                          f"Are you sure you want to CANCEL order {order_id}?\n\nThis will mark the order as CANCELLED and restock the items.")
            if confirm:
                self.process_order_action(order_id, "CANCELLED")
        elif action == 'no':
            # Return order
            confirm = messagebox.askyesno("Confirm Return",
                                          f"Are you sure you want to RETURN order {order_id}?\n\nThis will mark the order as RETURNED and restock the items.")
            if confirm:
                self.process_order_action(order_id, "RETURNED")

    def process_order_action(self, order_id, action):
        """Process order cancellation or return"""
        # Update order status
        df_orders = pd.read_excel(ORDERS_FILE)
        df_orders.loc[df_orders['order_id'] == order_id, 'status'] = action
        df_orders.to_excel(ORDERS_FILE, index=False)

        # Get all allocated piece IDs for this order
        allocated_pieces = df_orders.loc[df_orders['order_id'] == order_id, 'allocated_piece_ids']

        # Update stock status back to IN_STOCK for ALL pieces in this order
        df_stock = pd.read_excel(STOCK_FILE)

        for piece_ids_str in allocated_pieces:
            if pd.isna(piece_ids_str):
                continue

            piece_ids = piece_ids_str.split(',')
            for piece_id in piece_ids:
                df_stock.loc[df_stock['piece_id'] == piece_id, 'status'] = 'IN_STOCK'
                df_stock.loc[df_stock['piece_id'] == piece_id, 'sold_date'] = None
                df_stock.loc[df_stock['piece_id'] == piece_id, 'order_id'] = None

        df_stock.to_excel(STOCK_FILE, index=False)

        messagebox.showinfo("Success",
                            f"Order {order_id} {action.lower()} successfully. All items added back to stock.")

        # Refresh data
        self.load_orders()
        self.load_stock()
        self.update_product_comboboxes()

        # Auto-update summary if it's visible
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 2:  # Summary tab is index 2
            self.update_summary()

    def remove_selected_order_item(self):
        """Remove the selected item from the order items table"""
        selected_item = self.order_items_table.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select an item to remove")
            return

        # Get the index of the selected item
        item_index = self.order_items_table.index(selected_item[0])

        # Remove from the order items list
        if 0 <= item_index < len(self.order_items):
            self.order_items.pop(item_index)

        # Remove from the table
        self.order_items_table.delete(selected_item[0])

        # Update the order total
        self.update_order_total()

        # Update availability after removing item
        self.update_availability()

        messagebox.showinfo("Success", "Item removed from order")

    def add_order_item(self):
        product = self.order_product_var.get().strip()
        length = self.order_length_var.get().strip()
        qty = self.qty_var.get()

        if not product or not length:
            messagebox.showerror("Error", "Please select product and length")
            return

        try:
            length = int(length)
            qty = int(qty)
            if qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Please enter valid quantity and length")
            return

        # Calculate already added quantity for this product in current order
        already_added_qty = 0
        for item in self.order_items:
            if item["product"] == product and item["length"] == length:
                already_added_qty += item["qty"]

        # Check availability (considering already added items)
        df = pd.read_excel(STOCK_FILE)
        total_available = len(df[
                                  (df["product_name"] == product) & (df["length_m"] == length) & (
                                          df["status"] == "IN_STOCK")])

        actually_available = total_available - already_added_qty

        if qty > actually_available:
            messagebox.showerror("Error",
                                 f"Only {actually_available} pcs available (after considering already added items)")
            return

        # Get price from first available piece
        available_pieces = df[
            (df["product_name"] == product) & (df["length_m"] == length) & (df["status"] == "IN_STOCK")]

        if len(available_pieces) == 0:
            messagebox.showerror("Error", "No available pieces found")
            return

        price = available_pieces.iloc[0]["seller_price"]
        total = price * qty

        # Add to order items list
        self.order_items.append({
            "product": product,
            "length": length,
            "qty": qty,
            "price": price,
            "total": total
        })

        # Update order items table - WITHOUT REMOVE COLUMN
        self.order_items_table.insert("", "end",
                                      values=(product, length, qty, f"Rs. {price:,.2f}", f"Rs. {total:,.2f}"))

        # Update order total
        self.update_order_total()

        # Update availability after adding item
        self.update_availability()

        # Clear the input fields after adding
        self.order_product_var.set("")
        self.order_length_var.set("")
        self.qty_var.set(1)
        self.order_product_cb.focus()

    def update_order_total(self):
        total = sum(item["total"] for item in self.order_items)
        self.order_total_var.set(f"Rs. {total:,.2f}")

    def apply_order_filters(self):
        customer_filter = self.order_filter_customer_var.get().lower()
        product_filter = self.order_filter_product_var.get()
        date_filter = self.order_filter_date_var.get()
        status_filter = self.order_filter_status_var.get()

        df = pd.read_excel(ORDERS_FILE) if os.path.exists(ORDERS_FILE) else pd.DataFrame()
        if df.empty:
            return

        # Apply filters
        if customer_filter:
            df = df[df["customer_name"].str.lower().str.contains(customer_filter, na=False)]

        if product_filter and product_filter != "All":
            df = df[df["item_name"] == product_filter]

        if date_filter and date_filter != "All":
            # Handle date conversion safely
            try:
                # Try different date formats
                df["order_date"] = pd.to_datetime(df["order_date"], errors='coerce', format='mixed')
                today = datetime.now().date()

                # Drop rows with invalid dates
                df = df.dropna(subset=['order_date'])

                if date_filter == "Today":
                    df = df[df["order_date"].dt.date == today]
                elif date_filter == "Last 7 Days":
                    seven_days_ago = today - timedelta(days=7)
                    df = df[df["order_date"].dt.date >= seven_days_ago]
                elif date_filter == "This Month":
                    df = df[df["order_date"].dt.month == today.month]
                elif date_filter == "Last Month":
                    last_month = today.month - 1 if today.month > 1 else 12
                    df = df[df["order_date"].dt.month == last_month]
                elif date_filter == "Last 3 Months":
                    three_months_ago = today - timedelta(days=90)
                    df = df[df["order_date"].dt.date >= three_months_ago]
                elif date_filter == "Last 6 Months":
                    six_months_ago = today - timedelta(days=180)
                    df = df[df["order_date"].dt.date >= six_months_ago]
                elif date_filter == "Last 12 Months":
                    twelve_months_ago = today - timedelta(days=365)
                    df = df[df["order_date"].dt.date >= twelve_months_ago]
            except Exception as e:
                messagebox.showerror("Error", f"Date filter error: {str(e)}")
                return

        if status_filter and status_filter != "All":
            if "status" not in df.columns:
                df["status"] = "ACTIVE"  # Default status for old orders
            df = df[df["status"] == status_filter]

        # Display filtered results
        for row in self.orders_table.get_children():
            self.orders_table.delete(row)

        for _, row in df.iterrows():
            self.orders_table.insert("", "end", values=(
                row["order_id"],
                row["order_date"].strftime("%Y-%m-%d %H:%M") if hasattr(row["order_date"], 'strftime') else str(
                    row["order_date"]),
                row["customer_name"],
                row["city"],
                row["item_name"],
                row["length_m"],
                row["qty"],
                row["total_seller_price"],
                row.get("status", "ACTIVE")  # Default to ACTIVE if status column doesn't exist
            ))

        # Update order summary
        self.update_orders_summary(df)

    def update_orders_summary(self, df):
        """Update the total price and total pieces for displayed orders"""
        total_price = df["total_seller_price"].sum()
        total_pcs = df["qty"].sum()

        self.orders_total_price_var.set(f"Rs. {total_price:,.2f}")
        self.orders_total_pcs_var.set(f"{total_pcs:,}")

    def clear_order_filters(self):
        self.order_filter_customer_var.set("")
        self.order_filter_product_var.set("")
        self.order_filter_date_var.set("All")
        self.order_filter_status_var.set("ACTIVE")
        self.load_orders()

    def update_order_product_suggestions(self, event):
        typed = self.order_product_var.get().lower()
        if typed == '':
            self.order_product_cb['values'] = self.product_names
        else:
            matches = [name for name in self.product_names if typed in name.lower()]
            self.order_product_cb['values'] = matches

    def update_availability(self, event=None):
        """Update available quantity display with color coding - considering already added items"""
        product = self.order_product_var.get().strip()
        length = self.order_length_var.get().strip()

        if not product or not length:
            self.availability_lbl.config(text="Available: 0", foreground="black")
            return

        try:
            length = int(length)
            qty = self.qty_var.get()
        except (ValueError, tk.TclError):
            self.availability_lbl.config(text="Available: 0", foreground="black")
            return

        # Calculate already added quantity for this product in current order
        already_added_qty = 0
        for item in self.order_items:
            if item["product"] == product and item["length"] == length:
                already_added_qty += item["qty"]

        # හැකි තාක් quickly stock check කිරීමට
        try:
            df = pd.read_excel(STOCK_FILE)
            total_available = len(
                df[(df["product_name"] == product) & (df["length_m"] == length) & (df["status"] == "IN_STOCK")])

            # Calculate actually available (total minus already added)
            actually_available = total_available - already_added_qty

            # Show warning if requested quantity exceeds available stock
            if qty > actually_available:
                self.availability_lbl.config(text=f"Available: {actually_available} (Insufficient!)", foreground="red")
            else:
                self.availability_lbl.config(text=f"Available: {actually_available}", foreground="black")
        except Exception as e:
            self.availability_lbl.config(text="Error reading stock", foreground="red")

    def place_order(self):
        if not self.order_items:
            messagebox.showerror("Error", "Please add at least one item to the order")
            return

        name = self.cust_name.get().strip()
        address = self.cust_address.get().strip()
        phone1 = self.cust_phone1.get().strip()
        phone2 = self.cust_phone2.get().strip()
        city = self.cust_city.get().strip()

        if not name or not address or not phone1 or not city:
            messagebox.showerror("Error", "Customer details are required")
            return

        order_id = f"ORD_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        order_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        df_stock = pd.read_excel(STOCK_FILE)
        order_rows = []
        allocated_pieces = []

        for item in self.order_items:
            product = item["product"]
            length = item["length"]
            qty = item["qty"]
            custom_price = item["price"]  # Use the custom price from order items

            # Get available pieces for this item
            available_pieces = df_stock[
                (df_stock["product_name"] == product) & (df_stock["length_m"] == length) & (
                        df_stock["status"] == "IN_STOCK")].head(qty)

            if len(available_pieces) < qty:
                messagebox.showerror("Error", f"Only {len(available_pieces)} pcs available for {product} ({length}m)")
                return

            # Update stock to SOLD with custom price
            df_stock.loc[available_pieces.index, "status"] = "SOLD"
            df_stock.loc[available_pieces.index, "sold_date"] = order_date
            df_stock.loc[available_pieces.index, "order_id"] = order_id
            df_stock.loc[available_pieces.index, "seller_price"] = custom_price
            df_stock.loc[available_pieces.index, "profit"] = custom_price - df_stock.loc[
                available_pieces.index, "unit_cost"]

            # Calculate totals
            total_cost = available_pieces["unit_cost"].sum()
            total_price = custom_price * qty
            total_profit = total_price - total_cost

            # Add to order rows
            order_rows.append({
                "order_id": order_id,
                "order_date": order_date,
                "customer_name": name,
                "address": address,
                "phone1": phone1,
                "phone2": phone2,
                "city": city,
                "item_name": product,
                "length_m": length,
                "qty": qty,
                "total_unit_cost": total_cost,
                "total_seller_price": total_price,
                "profit_total": total_profit,
                "allocated_piece_ids": ",".join(available_pieces["piece_id"].tolist()),
                "status": "ACTIVE"
            })

            # Add to allocated pieces
            allocated_pieces.extend(available_pieces["piece_id"].tolist())

        # Save updated stock
        df_stock.to_excel(STOCK_FILE, index=False)

        # Save order(s)
        odf = pd.read_excel(ORDERS_FILE) if os.path.exists(ORDERS_FILE) else pd.DataFrame(columns=[
            "order_id", "order_date", "customer_name", "address", "phone1", "phone2", "city", "item_name",
            "length_m", "qty", "total_unit_cost", "total_seller_price", "profit_total", "allocated_piece_ids", "status"
        ])

        for order_row in order_rows:
            odf = pd.concat([odf, pd.DataFrame([order_row])], ignore_index=True)

        odf.to_excel(ORDERS_FILE, index=False)

        messagebox.showinfo("Success", f"Order {order_id} placed successfully!")

        # Clear form
        self.clear_customer_form()
        self.order_product_var.set("")
        self.order_length_var.set("")
        self.qty_var.set(1)
        self.availability_lbl.config(text="Available: 0")

        # Clear order items
        for item in self.order_items_table.get_children():
            self.order_items_table.delete(item)
        self.order_items = []
        self.update_order_total()

        # Refresh data
        self.load_orders()
        self.load_stock()
        self.update_product_comboboxes()

        # Auto-update summary if it's visible
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 2:  # Summary tab is index 2
            self.update_summary()

        # Save orders with proper date formatting
        with pd.ExcelWriter(ORDERS_FILE, engine='openpyxl', datetime_format='YYYY-MM-DD HH:MM:SS') as writer:
            odf.to_excel(writer, index=False)

    def load_orders(self):
        for row in self.orders_table.get_children():
            self.orders_table.delete(row)

        if os.path.exists(ORDERS_FILE):
            # Read with explicit date formatting
            df = pd.read_excel(ORDERS_FILE)

            # Convert order_date with proper formatting - handle mixed types
            if 'order_date' in df.columns:
                # First convert all to string to handle mixed types
                df['order_date'] = df['order_date'].astype(str)

                # Now convert to datetime with multiple format attempts
                df['order_date'] = pd.to_datetime(
                    df['order_date'],
                    errors='coerce',
                    format='mixed',  # Try multiple formats
                    dayfirst=False  # Use month-first format (MM/DD/YYYY)
                )

                # Drop rows with invalid dates
                df = df.dropna(subset=['order_date'])

            # Show latest orders first (only if we have valid dates)
            if not df.empty and 'order_date' in df.columns:
                df = df.sort_values("order_date", ascending=False)

            for _, row in df.head(100).iterrows():
                # Handle date display safely
                order_date = row["order_date"]
                if hasattr(order_date, 'strftime'):
                    formatted_date = order_date.strftime("%Y-%m-%d %H:%M")
                else:
                    # If it's still a string, try to format it
                    try:
                        formatted_date = pd.to_datetime(str(order_date)).strftime("%Y-%m-%d %H:%M")
                    except:
                        formatted_date = str(order_date)

                self.orders_table.insert("", "end", values=(
                    row["order_id"],
                    formatted_date,
                    row["customer_name"],
                    row["city"],
                    row["item_name"],
                    row["length_m"],
                    row["qty"],
                    row["total_seller_price"],
                    row.get("status", "ACTIVE")
                ))

        # Update order summary
        if not df.empty:
            self.update_orders_summary(df)

    def create_summary_tab(self):
        main_frame = ttk.Frame(self.summary_tab)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Filter Frame
        filter_frame = ttk.LabelFrame(main_frame, text="Summary Filters")
        filter_frame.pack(fill="x", pady=5)

        ttk.Label(filter_frame, text="Date Range:").grid(row=0, column=0, padx=5, pady=5)
        self.summary_date_var = tk.StringVar()
        self.summary_date_cb = ttk.Combobox(filter_frame, textvariable=self.summary_date_var,
                                            values=["All Time", "Today", "Last 7 Days", "This Month", "Last Month",
                                                    "Last 3 Months", "Last 6 Months", "Last 12 Months", "Custom Range"],
                                            width=15)
        self.summary_date_cb.set("All Time")
        self.summary_date_cb.grid(row=0, column=1, padx=5, pady=5)
        self.summary_date_cb.bind('<<ComboboxSelected>>', self.on_summary_date_change)

        ttk.Label(filter_frame, text="Product:").grid(row=0, column=2, padx=5, pady=5)
        self.summary_product_var = tk.StringVar()
        self.summary_product_cb = ttk.Combobox(filter_frame, textvariable=self.summary_product_var, width=15)
        self.summary_product_cb['values'] = self.product_names
        self.summary_product_cb.grid(row=0, column=3, padx=5, pady=5)
        self.summary_product_cb.bind('<<ComboboxSelected>>', lambda e: self.update_summary())

        ttk.Label(filter_frame, text="Length:").grid(row=0, column=4, padx=5, pady=5)
        self.summary_length_var = tk.StringVar()
        self.summary_length_cb = ttk.Combobox(filter_frame, textvariable=self.summary_length_var,
                                              values=["All", "5", "10", "15", "20", "30"], width=10)
        self.summary_length_cb.set("All")
        self.summary_length_cb.grid(row=0, column=5, padx=5, pady=5)
        self.summary_length_cb.bind('<<ComboboxSelected>>', lambda e: self.update_summary())

        # Sales Report Button (කලින් විදිහට)
        sales_report_btn = ttk.Button(filter_frame, text="Sales Report", command=self.show_sales_report)
        sales_report_btn.grid(row=0, column=6, padx=5, pady=5)

        # Summary Frame
        summary_frame = ttk.LabelFrame(main_frame, text="Business Summary")
        summary_frame.pack(fill="both", expand=True, pady=5)

        # Create a canvas and scrollbar for the summary tab
        canvas = tk.Canvas(summary_frame)
        scrollbar = ttk.Scrollbar(summary_frame, orient="vertical", command=canvas.yview)
        self.scrollable_summary_frame = ttk.Frame(canvas)

        self.scrollable_summary_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_summary_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Initial summary update
        self.update_summary()

    def on_summary_date_change(self, event):
        if self.summary_date_var.get() == "Custom Range":
            self.select_custom_date_range()
        else:
            self.update_summary()

    def select_custom_date_range(self):
        # Create a popup for custom date range selection
        popup = tk.Toplevel(self)
        popup.title("Select Custom Date Range")
        popup.geometry("300x150")
        popup.transient(self)
        popup.grab_set()

        ttk.Label(popup, text="Start Date (YYYY-MM-DD):").pack(pady=5)
        start_date_entry = ttk.Entry(popup)
        start_date_entry.pack(pady=5)

        ttk.Label(popup, text="End Date (YYYY-MM-DD):").pack(pady=5)
        end_date_entry = ttk.Entry(popup)
        end_date_entry.pack(pady=5)

        def apply_custom_range():
            start_date = start_date_entry.get()
            end_date = end_date_entry.get()

            try:
                # Validate dates
                datetime.strptime(start_date, "%Y-%m-%d")
                datetime.strptime(end_date, "%Y-%m-%d")

                self.custom_start_date = start_date
                self.custom_end_date = end_date
                self.update_summary()
                popup.destroy()
            except ValueError:
                messagebox.showerror("Error", "Please enter valid dates in YYYY-MM-DD format")

        ttk.Button(popup, text="Apply", command=apply_custom_range).pack(pady=10)

    def show_sales_report(self):
        # Create a new window for sales report
        report_window = tk.Toplevel(self)
        report_window.title("Sales Report")
        report_window.geometry("1000x600")

        # Frame for report controls
        control_frame = ttk.Frame(report_window)
        control_frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(control_frame, text="Year:").grid(row=0, column=0, padx=5, pady=5)
        year_var = tk.StringVar(value=str(datetime.now().year))
        year_cb = ttk.Combobox(control_frame, textvariable=year_var,
                               values=[str(y) for y in range(2020, datetime.now().year + 1)], width=10)
        year_cb.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(control_frame, text="Product:").grid(row=0, column=2, padx=5, pady=5)
        product_var = tk.StringVar()
        product_cb = ttk.Combobox(control_frame, textvariable=product_var,
                                  values=self.product_names, width=20)
        product_cb.grid(row=0, column=3, padx=5, pady=5)

        def generate_report():
            year = int(year_var.get())
            product_filter = product_var.get() if product_var.get() else None

            report_df = self.generate_sales_report_data(year, product_filter)
            self.display_sales_report(report_df, report_text, year, product_filter)

        generate_btn = ttk.Button(control_frame, text="Generate Report", command=generate_report)
        generate_btn.grid(row=0, column=4, padx=5, pady=5)

        # Text area for report
        report_text = scrolledtext.ScrolledText(report_window, wrap=tk.WORD, font=("Courier New", 10))
        report_text.pack(fill="both", expand=True, padx=10, pady=10)

        # Generate initial report for current year
        initial_report = self.generate_sales_report_data(datetime.now().year)
        self.display_sales_report(initial_report, report_text, datetime.now().year)

    def generate_sales_report_data(self, year, product_filter=None):
        # Load orders data
        if not os.path.exists(ORDERS_FILE):
            return pd.DataFrame()

        df_orders = pd.read_excel(ORDERS_FILE)

        # Filter by year and status (only ACTIVE orders for sales report)
        df_orders['order_date'] = pd.to_datetime(df_orders['order_date'], errors='coerce')
        df_orders = df_orders[df_orders['order_date'].dt.year == year]
        df_orders = df_orders[df_orders['status'] == 'ACTIVE']

        # Apply product filter if specified
        if product_filter:
            df_orders = df_orders[df_orders['item_name'] == product_filter]

        # Group by month
        df_orders['month'] = df_orders['order_date'].dt.month
        df_orders['year'] = df_orders['order_date'].dt.year

        # Create a summary dataframe - only for months with orders
        report_data = []

        # Get unique months with orders, sorted in descending order (newest first)
        months_with_orders = sorted(df_orders['month'].unique(), reverse=True)

        for month in months_with_orders:
            month_orders = df_orders[df_orders['month'] == month]

            total_sales = month_orders['total_seller_price'].sum()
            total_cost = month_orders['total_unit_cost'].sum()
            total_profit = month_orders['profit_total'].sum()
            total_quantity = month_orders['qty'].sum()

            profit_percentage = (total_profit / total_cost * 100) if total_cost > 0 else 0

            report_data.append({
                'month': month,
                'month_name': calendar.month_name[month],
                'total_sales': total_sales,
                'total_cost': total_cost,
                'total_profit': total_profit,
                'total_quantity': total_quantity,
                'profit_percentage': profit_percentage
            })

        # Add yearly total if there are any orders
        if not df_orders.empty:
            yearly_sales = df_orders['total_seller_price'].sum()
            yearly_cost = df_orders['total_unit_cost'].sum()
            yearly_profit = df_orders['profit_total'].sum()
            yearly_quantity = df_orders['qty'].sum()
            yearly_profit_pct = (yearly_profit / yearly_cost * 100) if yearly_cost > 0 else 0

            report_data.append({
                'month': 13,
                'month_name': 'YEARLY TOTAL',
                'total_sales': yearly_sales,
                'total_cost': yearly_cost,
                'total_profit': yearly_profit,
                'total_quantity': yearly_quantity,
                'profit_percentage': yearly_profit_pct
            })

        return pd.DataFrame(report_data)

    def display_sales_report(self, report_df, report_text, year, product_filter=None):
        report_text.delete(1.0, tk.END)

        if report_df.empty:
            report_text.insert(tk.END, "No sales data available for the selected criteria.")
            return

        # Header
        filter_info = f" for {product_filter}" if product_filter else " for All Products"
        report_text.insert(tk.END, f"SALES REPORT - {year}{filter_info}\n")
        report_text.insert(tk.END, "=" * 80 + "\n\n")

        # Column headers
        header = f"{'Month':<15} {'Sales (Rs)':>15} {'Cost (Rs)':>15} {'Profit (Rs)':>15} {'Qty':>10} {'Profit %':>10}\n"
        report_text.insert(tk.END, header)
        report_text.insert(tk.END, "-" * 80 + "\n")

        # Monthly data - already sorted in descending order by month
        for _, row in report_df.iterrows():
            if row['month'] == 13:  # Yearly total
                report_text.insert(tk.END, "-" * 80 + "\n")
                line = f"{row['month_name']:<15} {row['total_sales']:>15,.2f} {row['total_cost']:>15,.2f} {row['total_profit']:>15,.2f} {row['total_quantity']:>10,.0f} {row['profit_percentage']:>10.2f}%\n"
                report_text.insert(tk.END, line)
                report_text.insert(tk.END, "=" * 80 + "\n")
            else:
                line = f"{row['month_name']:<15} {row['total_sales']:>15,.2f} {row['total_cost']:>15,.2f} {row['total_profit']:>15,.2f} {row['total_quantity']:>10,.0f} {row['profit_percentage']:>10.2f}%\n"
                report_text.insert(tk.END, line)

        # Add some insights at the end
        report_text.insert(tk.END, "\nINSIGHTS:\n")
        report_text.insert(tk.END, "=" * 80 + "\n")

        # Find best and worst months (excluding yearly total)
        monthly_data = report_df[report_df['month'] != 13]

        if not monthly_data.empty:
            best_month = monthly_data.loc[monthly_data['total_profit'].idxmax()]
            worst_month = monthly_data.loc[monthly_data['total_profit'].idxmin()]

            report_text.insert(tk.END,
                               f"• Best Month: {best_month['month_name']} (Profit: Rs. {best_month['total_profit']:,.2f})\n")
            report_text.insert(tk.END,
                               f"• Worst Month: {worst_month['month_name']} (Profit: Rs. {worst_month['total_profit']:,.2f})\n")

            # Average monthly profit
            avg_profit = monthly_data['total_profit'].mean()
            report_text.insert(tk.END, f"• Average Monthly Profit: Rs. {avg_profit:,.2f}\n")

        # Add yearly total insights if available
        yearly_data = report_df[report_df['month'] == 13]
        if not yearly_data.empty:
            yearly_row = yearly_data.iloc[0]
            report_text.insert(tk.END, f"• Yearly Profit Margin: {yearly_row['profit_percentage']:.2f}%\n")

    def update_summary(self):
        # Clear previous summary
        for widget in self.scrollable_summary_frame.winfo_children():
            widget.destroy()

        # Load data with filters
        df_stock = pd.read_excel(STOCK_FILE)
        df_orders = pd.read_excel(ORDERS_FILE) if os.path.exists(ORDERS_FILE) else pd.DataFrame()

        # Apply date filter to orders safely
        date_filter = self.summary_date_var.get()
        product_filter = self.summary_product_var.get()
        length_filter = self.summary_length_var.get()

        if not df_orders.empty and date_filter != "All Time":
            try:
                # Convert dates safely
                df_orders["order_date"] = pd.to_datetime(df_orders["order_date"], errors='coerce', format='mixed')
                df_orders = df_orders.dropna(subset=['order_date'])

                today = datetime.now().date()

                if date_filter == "Today":
                    df_orders = df_orders[df_orders["order_date"].dt.date == today]
                elif date_filter == "Last 7 Days":
                    seven_days_ago = today - timedelta(days=7)
                    df_orders = df_orders[df_orders["order_date"].dt.date >= seven_days_ago]
                elif date_filter == "This Month":
                    df_orders = df_orders[df_orders["order_date"].dt.month == today.month]
                    df_orders = df_orders[df_orders["order_date"].dt.year == today.year]
                elif date_filter == "Last Month":
                    last_month = today.month - 1 if today.month > 1 else 12
                    last_month_year = today.year if today.month > 1 else today.year - 1
                    df_orders = df_orders[df_orders["order_date"].dt.month == last_month]
                    df_orders = df_orders[df_orders["order_date"].dt.year == last_month_year]
                elif date_filter == "Last 3 Months":
                    three_months_ago = today - timedelta(days=90)
                    df_orders = df_orders[df_orders["order_date"].dt.date >= three_months_ago]
                elif date_filter == "Last 6 Months":
                    six_months_ago = today - timedelta(days=180)
                    df_orders = df_orders[df_orders["order_date"].dt.date >= six_months_ago]
                elif date_filter == "Last 12 Months":
                    twelve_months_ago = today - timedelta(days=365)
                    df_orders = df_orders[df_orders["order_date"].dt.date >= twelve_months_ago]
                elif date_filter == "Custom Range" and hasattr(self, 'custom_start_date') and hasattr(self,
                                                                                                      'custom_end_date'):
                    start_date = datetime.strptime(self.custom_start_date, "%Y-%m-%d").date()
                    end_date = datetime.strptime(self.custom_end_date, "%Y-%m-%d").date()
                    df_orders = df_orders[(df_orders["order_date"].dt.date >= start_date) &
                                          (df_orders["order_date"].dt.date <= end_date)]
            except Exception as e:
                messagebox.showerror("Error", f"Date filter error: {str(e)}")
                return

        # Apply product filter
        if product_filter:
            df_stock = df_stock[df_stock["product_name"] == product_filter]
            if not df_orders.empty:
                df_orders = df_orders[df_orders["item_name"] == product_filter]

        # Apply length filter
        if length_filter and length_filter != "All":
            df_stock = df_stock[df_stock["length_m"] == int(length_filter)]
            if not df_orders.empty:
                df_orders = df_orders[df_orders["length_m"] == int(length_filter)]

        # Calculate metrics
        total_instock = len(df_stock[df_stock["status"] == "IN_STOCK"])
        total_sold = len(df_stock[df_stock["status"] == "SOLD"])
        total_removed = len(df_stock[df_stock["status"] == "REMOVED"])

        total_cost = df_stock["unit_cost"].sum()
        seller_value = df_stock["seller_price"].sum()
        total_profit = seller_value - total_cost

        # Calculate profit percentage
        profit_percentage = (total_profit / total_cost * 100) if total_cost > 0 else 0

        # Only count ACTIVE orders for profit and revenue
        if not df_orders.empty:
            if "status" in df_orders.columns:
                active_orders = df_orders[df_orders["status"] == "ACTIVE"]
                total_order_profit = active_orders["profit_total"].sum()
                total_revenue = active_orders["total_seller_price"].sum()
                total_order_cost = active_orders["total_unit_cost"].sum()
                cancelled_orders = df_orders[df_orders["status"] == "CANCELLED"]
                returned_orders = df_orders[df_orders["status"] == "RETURNED"]
                total_cancelled = len(cancelled_orders)
                total_returned = len(returned_orders)
            else:
                total_order_profit = df_orders["profit_total"].sum()
                total_revenue = df_orders["total_seller_price"].sum()
                total_order_cost = df_orders["total_unit_cost"].sum()
                total_cancelled = 0
                total_returned = 0
        else:
            total_order_profit = 0
            total_revenue = 0
            total_order_cost = 0
            total_cancelled = 0
            total_returned = 0

        total_orders = len(df_orders) if not df_orders.empty else 0
        active_orders_count = total_orders - total_cancelled - total_returned

        # Calculate order profit percentage
        order_profit_percentage = (total_order_profit / total_order_cost * 100) if total_order_cost > 0 else 0

        # Calculate additional business metrics
        inventory_turnover = total_revenue / seller_value if seller_value > 0 else 0
        cancellation_rate = (total_cancelled / total_orders * 100) if total_orders > 0 else 0
        return_rate = (total_returned / total_orders * 100) if total_orders > 0 else 0

        # Display metrics - Enhanced Business Summary
        # Header
        ttk.Label(self.scrollable_summary_frame, text="📊 BUSINESS PERFORMANCE DASHBOARD",
                  font=("Arial", 14, "bold"), foreground="darkblue").grid(row=0, column=0, columnspan=4, pady=15)

        # Key Metrics Section
        metrics_frame = ttk.LabelFrame(self.scrollable_summary_frame, text="🚀 Key Performance Indicators")
        metrics_frame.grid(row=1, column=0, columnspan=4, padx=10, pady=10, sticky="ew")

        # Row 1
        ttk.Label(metrics_frame, text="Total Revenue:", font=("Arial", 11, "bold")).grid(row=0, column=0, sticky="w",
                                                                                         padx=15, pady=8)
        ttk.Label(metrics_frame, text=f"Rs. {total_revenue:,.2f}", font=("Arial", 11, "bold"), foreground="green").grid(
            row=0, column=1, sticky="w", padx=15, pady=8)

        ttk.Label(metrics_frame, text="Total Order Cost:", font=("Arial", 11, "bold")).grid(row=0, column=2, sticky="w",
                                                                                            padx=15, pady=8)
        ttk.Label(metrics_frame, text=f"Rs. {total_order_cost:,.2f}", font=("Arial", 11, "bold")).grid(
            row=0, column=3, sticky="w", padx=15, pady=8)

        ttk.Label(metrics_frame, text="Order Profit %:", font=("Arial", 11, "bold")).grid(row=0, column=4, sticky="w",
                                                                                          padx=15, pady=8)
        order_profit_color = "darkgreen" if order_profit_percentage >= 20 else "orange" if order_profit_percentage >= 10 else "red"
        ttk.Label(metrics_frame, text=f"{order_profit_percentage:.2f}%", font=("Arial", 11, "bold"),
                  foreground=order_profit_color).grid(row=0, column=5, sticky="w", padx=15, pady=8)

        # Row 2
        ttk.Label(metrics_frame, text="Total Profit:", font=("Arial", 11, "bold")).grid(row=1, column=0, sticky="w",
                                                                                        padx=15, pady=8)
        ttk.Label(metrics_frame, text=f"Rs. {total_order_profit:,.2f}", font=("Arial", 11, "bold"),
                  foreground="darkgreen").grid(
            row=1, column=1, sticky="w", padx=15, pady=8)

        ttk.Label(metrics_frame, text="Inventory Turnover:", font=("Arial", 11, "bold")).grid(row=1, column=2,
                                                                                              sticky="w", padx=15,
                                                                                              pady=8)
        turnover_color = "darkgreen" if inventory_turnover >= 4 else "orange" if inventory_turnover >= 2 else "red"
        ttk.Label(metrics_frame, text=f"{inventory_turnover:.2f}x", font=("Arial", 11), foreground=turnover_color).grid(
            row=1, column=3, sticky="w", padx=15, pady=8)

        ttk.Label(metrics_frame, text="Active Orders:", font=("Arial", 11, "bold")).grid(row=1, column=4, sticky="w",
                                                                                         padx=15, pady=8)
        ttk.Label(metrics_frame, text=f"{active_orders_count}", font=("Arial", 11)).grid(row=1, column=5, sticky="w",
                                                                                         padx=15, pady=8)

        # Stock Analysis Section
        stock_frame = ttk.LabelFrame(self.scrollable_summary_frame, text="📦 Stock Analysis")
        stock_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        ttk.Label(stock_frame, text="IN_STOCK Items:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w",
                                                                                        padx=10, pady=5)
        ttk.Label(stock_frame, text=f"{total_instock}", font=("Arial", 10)).grid(row=0, column=1, sticky="w", padx=10,
                                                                                 pady=5)

        ttk.Label(stock_frame, text="SOLD Items:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w",
                                                                                    padx=10, pady=5)
        ttk.Label(stock_frame, text=f"{total_sold}", font=("Arial", 10)).grid(row=1, column=1, sticky="w", padx=10,
                                                                              pady=5)

        ttk.Label(stock_frame, text="REMOVED Items:", font=("Arial", 10, "bold")).grid(row=2, column=0, sticky="w",
                                                                                       padx=10, pady=5)
        ttk.Label(stock_frame, text=f"{total_removed}", font=("Arial", 10)).grid(row=2, column=1, sticky="w", padx=10,
                                                                                 pady=5)

        ttk.Label(stock_frame, text="Inventory Value:", font=("Arial", 10, "bold")).grid(row=3, column=0, sticky="w",
                                                                                         padx=10, pady=5)
        ttk.Label(stock_frame, text=f"Rs. {seller_value:,.2f}", font=("Arial", 10)).grid(row=3, column=1, sticky="w",
                                                                                         padx=10, pady=5)

        ttk.Label(stock_frame, text="Inventory Cost:", font=("Arial", 10, "bold")).grid(row=4, column=0, sticky="w",
                                                                                        padx=10, pady=5)
        ttk.Label(stock_frame, text=f"Rs. {total_cost:,.2f}", font=("Arial", 10)).grid(row=4, column=1, sticky="w",
                                                                                       padx=10, pady=5)

        ttk.Label(stock_frame, text="Inventory Profit:", font=("Arial", 10, "bold")).grid(row=5, column=0, sticky="w",
                                                                                          padx=10, pady=5)
        ttk.Label(stock_frame, text=f"Rs. {total_profit:,.2f}", font=("Arial", 10),
                  foreground="green" if total_profit > 0 else "red").grid(row=5, column=1, sticky="w", padx=10, pady=5)

        ttk.Label(stock_frame, text="Profit %:", font=("Arial", 10, "bold")).grid(row=6, column=0, sticky="w",
                                                                                  padx=10, pady=5)
        profit_pct_color = "green" if profit_percentage > 0 else "red"
        ttk.Label(stock_frame, text=f"{profit_percentage:.2f}%", font=("Arial", 10),
                  foreground=profit_pct_color).grid(row=6, column=1, sticky="w", padx=10, pady=5)

        # Order Analysis Section
        order_frame = ttk.LabelFrame(self.scrollable_summary_frame, text="📋 Order Analysis")
        order_frame.grid(row=2, column=2, columnspan=2, padx=10, pady=10, sticky="nsew")

        ttk.Label(order_frame, text="Total Orders:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w",
                                                                                      padx=10, pady=5)
        ttk.Label(order_frame, text=f"{total_orders}", font=("Arial", 10)).grid(row=0, column=1, sticky="w", padx=10,
                                                                                pady=5)

        ttk.Label(order_frame, text="Cancellation Rate:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w",
                                                                                           padx=10, pady=5)
        cancel_color = "red" if cancellation_rate > 5 else "orange" if cancellation_rate > 2 else "darkgreen"
        ttk.Label(order_frame, text=f"{cancellation_rate:.2f}%", font=("Arial", 10), foreground=cancel_color).grid(
            row=1, column=1, sticky="w", padx=10, pady=5)

        ttk.Label(order_frame, text="Return Rate:", font=("Arial", 10, "bold")).grid(row=2, column=0, sticky="w",
                                                                                     padx=10, pady=5)
        return_color = "red" if return_rate > 5 else "orange" if return_rate > 2 else "darkgreen"
        ttk.Label(order_frame, text=f"{return_rate:.2f}%", font=("Arial", 10), foreground=return_color).grid(row=2,
                                                                                                             column=1,
                                                                                                             sticky="w",
                                                                                                             padx=10,
                                                                                                             pady=5)

        ttk.Label(order_frame, text="Success Rate:", font=("Arial", 10, "bold")).grid(row=3, column=0, sticky="w",
                                                                                      padx=10, pady=5)
        success_rate = 100 - cancellation_rate - return_rate
        success_color = "darkgreen" if success_rate >= 90 else "orange" if success_rate >= 80 else "red"
        ttk.Label(order_frame, text=f"{success_rate:.2f}%", font=("Arial", 10), foreground=success_color).grid(row=3,
                                                                                                               column=1,
                                                                                                               sticky="w",
                                                                                                               padx=10,
                                                                                                               pady=5)

        # Performance Insights Section
        insights_frame = ttk.LabelFrame(self.scrollable_summary_frame, text="💡 Performance Insights")
        insights_frame.grid(row=3, column=0, columnspan=4, padx=10, pady=10, sticky="ew")

        insights_text = scrolledtext.ScrolledText(insights_frame, height=6, wrap=tk.WORD, font=("Arial", 9))
        insights_text.pack(fill="both", expand=True, padx=10, pady=10)

        # Generate insights
        insights = []

        # Order profitability insights
        if order_profit_percentage > 25:
            insights.append("💰 Excellent order profitability! Your pricing strategy is working well.")
        elif order_profit_percentage > 15:
            insights.append("💰 Good order profitability. Maintain your current pricing.")
        else:
            insights.append("⚠️  Low order profitability. Consider reviewing your pricing strategy.")

        # Inventory insights
        if seller_value > total_cost * 2:
            insights.append("📈 Excellent inventory profitability!")
        elif seller_value > total_cost * 1.5:
            insights.append("📈 Good inventory profitability.")
        else:
            insights.append("⚠️  Low inventory profitability. Consider reviewing your pricing strategy.")

        if total_instock < 20:
            insights.append("📦 Low inventory levels! Consider restocking to avoid stockouts.")
        elif total_instock > 100:
            insights.append("📦 High inventory levels. Consider promotions to reduce stock.")

        # Business insights
        if inventory_turnover > 6:
            insights.append("🚀 High inventory turnover! Your stock is moving quickly.")
        elif inventory_turnover > 3:
            insights.append("📊 Healthy inventory turnover. Maintain current levels.")
        else:
            insights.append("⚠️  Low inventory turnover. Consider promotions or reviewing slow-moving items.")

        if cancellation_rate > 10:
            insights.append("❌ High cancellation rate. Review order processing and customer service.")
        elif cancellation_rate > 5:
            insights.append("⚠️  Moderate cancellation rate. Monitor order fulfillment process.")

        if return_rate > 8:
            insights.append("❌ High return rate. Check product quality and customer expectations.")
        elif return_rate > 3:
            insights.append("⚠️  Moderate return rate. Ensure accurate product descriptions.")

        if not insights:
            insights.append("📈 Business performance is stable. Continue monitoring key metrics.")

        insights_text.insert(tk.END, "\n".join(insights))
        insights_text.config(state=tk.DISABLED)

        # Configure grid weights for proper resizing
        for i in range(4):
            self.scrollable_summary_frame.columnconfigure(i, weight=1)
            metrics_frame.columnconfigure(i, weight=1)
            stock_frame.columnconfigure(i, weight=1)
            order_frame.columnconfigure(i, weight=1)


if __name__ == "__main__":
    init_files()
    app = StockApp()
    app.mainloop()
