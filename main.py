import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pandas as pd
import os
from datetime import datetime, timedelta

STOCK_FILE = "stock.xlsx"
ORDERS_FILE = "orders.xlsx"


# Ensure Excel files exist
def init_files():
    if not os.path.exists(STOCK_FILE):
        df = pd.DataFrame(
            columns=["piece_id", "product_name", "length_m", "date_added", "seller_price", "unit_cost", "profit",
                     "status", "sold_date", "order_id"])
        df.to_excel(STOCK_FILE, index=False)
    if not os.path.exists(ORDERS_FILE):
        df = pd.DataFrame(
            columns=["order_id", "order_date", "customer_name", "address", "phone1", "phone2", "city", "item_name",
                     "length_m", "qty", "total_unit_cost", "total_seller_price", "profit_total", "allocated_piece_ids"])
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

        ttk.Label(add_frame, text="Length (m):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.length_var = tk.StringVar()
        self.length_cb = ttk.Combobox(add_frame, textvariable=self.length_var, values=[5, 10, 15, 20, 30], width=15)
        self.length_cb.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(add_frame, text="PCS:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.pcs_entry = ttk.Entry(add_frame, width=15)
        self.pcs_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(add_frame, text="Unit Cost:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.unit_cost_entry = ttk.Entry(add_frame, width=15)
        self.unit_cost_entry.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        ttk.Label(add_frame, text="Seller Price:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.seller_price_entry = ttk.Entry(add_frame, width=15)
        self.seller_price_entry.grid(row=1, column=3, padx=5, pady=5, sticky="w")

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

        clear_filter_btn = ttk.Button(filter_frame, text="Clear Filters", command=self.clear_stock_filters)
        clear_filter_btn.grid(row=0, column=6, padx=5, pady=5)

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

        # Load initial stock data
        self.load_stock()

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

        df = pd.read_excel(STOCK_FILE)

        # Apply filters
        if product_filter and product_filter != "All":
            df = df[df["product_name"] == product_filter]

        if length_filter and length_filter != "All":
            df = df[df["length_m"] == int(length_filter)]

        if status_filter and status_filter != "All":
            df = df[df["status"] == status_filter]

        # Display filtered results
        for row in self.stock_table.get_children():
            self.stock_table.delete(row)

        for _, row in df.iterrows():
            self.stock_table.insert("", "end", values=(
                row["piece_id"], row["product_name"], row["length_m"],
                row["date_added"], row["seller_price"], row["unit_cost"],
                row["profit"], row["status"]
            ))

    def clear_stock_filters(self):
        self.filter_product_var.set("")
        self.filter_length_var.set("All")
        self.filter_status_var.set("IN_STOCK")
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
        self.pcs_entry.delete(0, tk.END)
        self.unit_cost_entry.delete(0, tk.END)
        self.seller_price_entry.delete(0, tk.END)

        # Refresh data
        self.load_stock()
        self.update_product_comboboxes()

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

        df.to_excel(STOCK_FILE, index=False)

        messagebox.showinfo("Success", f"Removed {len(selected_items)} item(s)")
        self.load_stock()

    def load_stock(self):
        for row in self.stock_table.get_children():
            self.stock_table.delete(row)

        df = pd.read_excel(STOCK_FILE)
        # Apply current filter
        status_filter = self.filter_status_var.get() if hasattr(self, 'filter_status_var') else "IN_STOCK"
        if status_filter and status_filter != "All":
            df = df[df['status'] == status_filter]
        else:
            df = df[df['status'] == 'IN_STOCK']

        for _, row in df.iterrows():
            self.stock_table.insert("", "end", values=(
                row["piece_id"], row["product_name"], row["length_m"],
                row["date_added"], row["seller_price"], row["unit_cost"],
                row["profit"], row["status"]
            ))

    def create_orders_tab(self):
        # Main frame
        main_frame = ttk.Frame(self.orders_tab)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Left frame for order form
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        # Right frame for order history
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)

        # Order form
        form = ttk.LabelFrame(left_frame, text="Place New Order")
        form.pack(fill="both", expand=True, padx=5, pady=5)

        # Customer details
        cust_frame = ttk.Frame(form)
        cust_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(cust_frame, text="Customer Name:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.cust_name = ttk.Entry(cust_frame, width=25)
        self.cust_name.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(cust_frame, text="Address:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.cust_address = ttk.Entry(cust_frame, width=25)
        self.cust_address.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(cust_frame, text="Phone 1:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.cust_phone1 = ttk.Entry(cust_frame, width=25)
        self.cust_phone1.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(cust_frame, text="Phone 2:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.cust_phone2 = ttk.Entry(cust_frame, width=25)
        self.cust_phone2.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(cust_frame, text="City:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.cust_city = ttk.Entry(cust_frame, width=25)
        self.cust_city.grid(row=4, column=1, padx=5, pady=5, sticky="w")

        # Order items frame
        items_frame = ttk.LabelFrame(form, text="Order Items")
        items_frame.pack(fill="x", padx=5, pady=5)

        # Item entry
        item_entry_frame = ttk.Frame(items_frame)
        item_entry_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(item_entry_frame, text="Item Name:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.order_product_var = tk.StringVar()
        self.order_product_cb = ttk.Combobox(item_entry_frame, textvariable=self.order_product_var, width=20)
        self.order_product_cb.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.order_product_cb.bind('<KeyRelease>', self.update_order_product_suggestions)

        ttk.Label(item_entry_frame, text="Length (m):").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.order_length_var = tk.StringVar()
        self.order_length_cb = ttk.Combobox(item_entry_frame, textvariable=self.order_length_var,
                                            values=[5, 10, 15, 20, 30],
                                            width=10)
        self.order_length_cb.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        self.order_length_cb.bind("<<ComboboxSelected>>", self.update_availability)

        self.availability_lbl = ttk.Label(item_entry_frame, text="Available: 0")
        self.availability_lbl.grid(row=0, column=4, padx=5, pady=5)

        ttk.Label(item_entry_frame, text="Qty:").grid(row=0, column=5, padx=5, pady=5, sticky="e")
        self.qty_var = tk.IntVar(value=1)
        self.qty_entry = ttk.Entry(item_entry_frame, textvariable=self.qty_var, width=10)
        self.qty_entry.grid(row=0, column=6, padx=5, pady=5, sticky="w")

        add_item_btn = ttk.Button(item_entry_frame, text="Add Item", command=self.add_order_item)
        add_item_btn.grid(row=0, column=7, padx=5, pady=5)

        # Order items table
        items_table_frame = ttk.Frame(items_frame)
        items_table_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Scrollbar for items table
        items_scrollbar = ttk.Scrollbar(items_table_frame)
        items_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.order_items_table = ttk.Treeview(items_table_frame,
                                              columns=("product", "length", "qty", "price", "total"),
                                              show="headings", height=5, yscrollcommand=items_scrollbar.set)
        items_scrollbar.config(command=self.order_items_table.yview)

        # Define columns for order items
        item_columns = [("product", "Product", 150),
                        ("length", "Length", 80),
                        ("qty", "Qty", 60),
                        ("price", "Price", 80),
                        ("total", "Total", 100)]

        for col_id, heading, width in item_columns:
            self.order_items_table.heading(col_id, text=heading)
            self.order_items_table.column(col_id, width=width, anchor="center")

        self.order_items_table.pack(fill="both", expand=True)

        # Order total
        total_frame = ttk.Frame(items_frame)
        total_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(total_frame, text="Order Total:", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5,
                                                                                     sticky="e")
        self.order_total_var = tk.StringVar(value="Rs. 0.00")
        ttk.Label(total_frame, textvariable=self.order_total_var, font=("Arial", 10, "bold")).grid(row=0, column=1,
                                                                                                   padx=5, pady=5,
                                                                                                   sticky="w")

        # Place order button
        order_btn_frame = ttk.Frame(form)
        order_btn_frame.pack(fill="x", padx=5, pady=5)

        order_btn = ttk.Button(order_btn_frame, text="Place Order", command=self.place_order)
        order_btn.pack(side="right", padx=5, pady=5)

        # Order history section
        history_frame = ttk.LabelFrame(right_frame, text="Order History & Filters")
        history_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Order Filters
        order_filter_frame = ttk.Frame(history_frame)
        order_filter_frame.pack(fill="x", padx=10, pady=5)

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
                                                 values=["All", "Today", "Last 7 Days", "This Month", "Last Month"],
                                                 width=15)
        self.order_filter_date_cb.set("All")
        self.order_filter_date_cb.grid(row=0, column=5, padx=5, pady=5)
        self.order_filter_date_cb.bind('<<ComboboxSelected>>', lambda e: self.apply_order_filters())

        clear_order_filter_btn = ttk.Button(order_filter_frame, text="Clear Filters", command=self.clear_order_filters)
        clear_order_filter_btn.grid(row=0, column=6, padx=5, pady=5)

        # Orders history with scrollbar
        table_frame = ttk.Frame(history_frame)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        scrollbar = ttk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.orders_table = ttk.Treeview(table_frame,
                                         columns=("order_id", "order_date", "customer_name", "city", "item_name",
                                                  "length_m", "qty", "total_seller_price"),
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
                   ("total_seller_price", "Total", 100)]

        for col_id, heading, width in columns:
            self.orders_table.heading(col_id, text=heading)
            self.orders_table.column(col_id, width=width, anchor="center")

        self.orders_table.pack(fill="both", expand=True)

        # Initialize order items list
        self.order_items = []

        # Load initial orders
        self.load_orders()

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

        # Check availability
        df = pd.read_excel(STOCK_FILE)
        available_pieces = df[
            (df["product_name"] == product) & (df["length_m"] == length) & (df["status"] == "IN_STOCK")]

        if len(available_pieces) < qty:
            messagebox.showerror("Error", f"Only {len(available_pieces)} pcs available")
            return

        # Get price from first available piece
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

        # Update order items table
        self.order_items_table.insert("", "end",
                                      values=(product, length, qty, f"Rs. {price:,.2f}", f"Rs. {total:,.2f}"))

        # Update order total
        self.update_order_total()

        # Clear item fields
        self.order_product_var.set("")
        self.order_length_var.set("")
        self.qty_var.set(1)
        self.availability_lbl.config(text="Available: 0")

    def update_order_total(self):
        total = sum(item["total"] for item in self.order_items)
        self.order_total_var.set(f"Rs. {total:,.2f}")

    def apply_order_filters(self):
        customer_filter = self.order_filter_customer_var.get().lower()
        product_filter = self.order_filter_product_var.get()
        date_filter = self.order_filter_date_var.get()

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
            except Exception as e:
                messagebox.showerror("Error", f"Date filter error: {str(e)}")
                return

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
                row["total_seller_price"]
            ))

    def clear_order_filters(self):
        self.order_filter_customer_var.set("")
        self.order_filter_product_var.set("")
        self.order_filter_date_var.set("All")
        self.load_orders()

    def update_order_product_suggestions(self, event):
        typed = self.order_product_var.get().lower()
        if typed == '':
            self.order_product_cb['values'] = self.product_names
        else:
            matches = [name for name in self.product_names if typed in name.lower()]
            self.order_product_cb['values'] = matches

    def update_availability(self, event=None):
        product = self.order_product_var.get().strip()
        length = self.order_length_var.get().strip()
        if not product or not length:
            self.availability_lbl.config(text="Available: 0")
            return

        try:
            length = int(length)
        except ValueError:
            self.availability_lbl.config(text="Available: 0")
            return

        df = pd.read_excel(STOCK_FILE)
        available = len(df[(df["product_name"] == product) & (df["length_m"] == length) & (df["status"] == "IN_STOCK")])
        self.availability_lbl.config(text=f"Available: {available}")

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

            # Get available pieces for this item
            available_pieces = df_stock[
                (df_stock["product_name"] == product) & (df_stock["length_m"] == length) & (
                        df_stock["status"] == "IN_STOCK")].head(qty)

            if len(available_pieces) < qty:
                messagebox.showerror("Error", f"Only {len(available_pieces)} pcs available for {product} ({length}m)")
                return

            # Update stock to SOLD
            df_stock.loc[available_pieces.index, "status"] = "SOLD"
            df_stock.loc[available_pieces.index, "sold_date"] = order_date
            df_stock.loc[available_pieces.index, "order_id"] = order_id

            # Calculate totals
            total_cost = available_pieces["unit_cost"].sum()
            total_price = available_pieces["seller_price"].sum()
            total_profit = available_pieces["profit"].sum()

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
                "allocated_piece_ids": ",".join(available_pieces["piece_id"].tolist())
            })

            # Add to allocated pieces
            allocated_pieces.extend(available_pieces["piece_id"].tolist())

        # Save updated stock
        df_stock.to_excel(STOCK_FILE, index=False)

        # Save order(s)
        odf = pd.read_excel(ORDERS_FILE) if os.path.exists(ORDERS_FILE) else pd.DataFrame(columns=[
            "order_id", "order_date", "customer_name", "address", "phone1", "phone2", "city", "item_name",
            "length_m", "qty", "total_unit_cost", "total_seller_price", "profit_total", "allocated_piece_ids"
        ])

        for order_row in order_rows:
            odf = pd.concat([odf, pd.DataFrame([order_row])], ignore_index=True)

        odf.to_excel(ORDERS_FILE, index=False)

        messagebox.showinfo("Success", f"Order {order_id} placed successfully!")

        # Clear form
        self.cust_name.delete(0, tk.END)
        self.cust_address.delete(0, tk.END)
        self.cust_phone1.delete(0, tk.END)
        self.cust_phone2.delete(0, tk.END)
        self.cust_city.delete(0, tk.END)
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

    def load_orders(self):
        for row in self.orders_table.get_children():
            self.orders_table.delete(row)

        if os.path.exists(ORDERS_FILE):
            df = pd.read_excel(ORDERS_FILE)
            # Show latest orders first
            df = df.sort_values("order_date", ascending=False)
            for _, row in df.head(100).iterrows():
                # Handle date display safely
                order_date = row["order_date"]
                if hasattr(order_date, 'strftime'):
                    formatted_date = order_date.strftime("%Y-%m-%d %H:%M")
                else:
                    formatted_date = str(order_date)

                self.orders_table.insert("", "end", values=(
                    row["order_id"],
                    formatted_date,
                    row["customer_name"],
                    row["city"],
                    row["item_name"],
                    row["length_m"],
                    row["qty"],
                    row["total_seller_price"]
                ))
        for item in selected_items:
            piece_id = self.stock_table.item(item)['values'][0]
            pieces_to_remove.append(piece_id)

        # Mark as REMOVED instead of actually deleting
        df.loc[df['piece_id'].isin(pieces_to_remove), 'status'] = 'REMOVED'


    def create_summary_tab(self):
        main_frame = ttk.Frame(self.summary_tab)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Filter Frame
        filter_frame = ttk.LabelFrame(main_frame, text="Summary Filters")
        filter_frame.pack(fill="x", pady=5)

        ttk.Label(filter_frame, text="Date Range:").grid(row=0, column=0, padx=5, pady=5)
        self.summary_date_var = tk.StringVar()
        self.summary_date_cb = ttk.Combobox(filter_frame, textvariable=self.summary_date_var,
                                            values=["All Time", "Today", "This Week", "This Month", "Last Month",
                                                    "This Year"], width=15)
        self.summary_date_cb.set("All Time")
        self.summary_date_cb.grid(row=0, column=1, padx=5, pady=5)
        self.summary_date_cb.bind('<<ComboboxSelected>>', lambda e: self.update_summary())

        ttk.Label(filter_frame, text="Product:").grid(row=0, column=2, padx=5, pady=5)
        self.summary_product_var = tk.StringVar()
        self.summary_product_cb = ttk.Combobox(filter_frame, textvariable=self.summary_product_var, width=15)
        self.summary_product_cb['values'] = self.product_names
        self.summary_product_cb.grid(row=0, column=3, padx=5, pady=5)
        self.summary_product_cb.bind('<<ComboboxSelected>>', lambda e: self.update_summary())

        # Summary Frame ONLY - Sales details table REMOVED
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

        # Sales details table REMOVED as requested

        # Initial summary update
        self.update_summary()

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

        if not df_orders.empty and date_filter != "All Time":
            try:
                # Convert dates safely
                df_orders["order_date"] = pd.to_datetime(df_orders["order_date"], errors='coerce', format='mixed')
                df_orders = df_orders.dropna(subset=['order_date'])

                today = datetime.now().date()

                if date_filter == "Today":
                    df_orders = df_orders[df_orders["order_date"].dt.date == today]
                elif date_filter == "This Week":
                    start_of_week = today - timedelta(days=today.weekday())
                    df_orders = df_orders[df_orders["order_date"].dt.date >= start_of_week]
                elif date_filter == "This Month":
                    df_orders = df_orders[df_orders["order_date"].dt.month == today.month]
                elif date_filter == "Last Month":
                    last_month = today.month - 1 if today.morrow > 1 else 12
                    df_orders = df_orders[df_orders["order_date"].dt.month == last_month]
                elif date_filter == "This Year":
                    df_orders = df_orders[df_orders["order_date"].dt.year == today.year]
            except Exception as e:
                messagebox.showerror("Error", f"Date filter error: {str(e)}")
                return

        # Apply product filter
        if product_filter:
            df_stock = df_stock[df_stock["product_name"] == product_filter]
            if not df_orders.empty:
                df_orders = df_orders[df_orders["item_name"] == product_filter]

        # Calculate metrics
        total_instock = len(df_stock[df_stock["status"] == "IN_STOCK"])
        total_sold = len(df_stock[df_stock["status"] == "SOLD"])
        total_removed = len(df_stock[df_stock["status"] == "REMOVED"])

        total_cost = df_stock["unit_cost"].sum()
        seller_value = df_stock["seller_price"].sum()
        profit = df_orders["profit_total"].sum() if not df_orders.empty else 0
        total_revenue = df_orders["total_seller_price"].sum() if not df_orders.empty else 0
        total_orders = len(df_orders) if not df_orders.empty else 0

        # Display metrics
        ttk.Label(self.scrollable_summary_frame, text="STOCK SUMMARY", font=("Arial", 12, "bold")).grid(row=0, column=0,
                                                                                                        columnspan=2,
                                                                                                        pady=10)

        ttk.Label(self.scrollable_summary_frame, text="IN_STOCK:", font=("Arial", 10, "bold")).grid(row=1, column=0,
                                                                                                    sticky="w", padx=10,
                                                                                                    pady=5)
        ttk.Label(self.scrollable_summary_frame, text=f"{total_instock} items").grid(row=1, column=1, sticky="w",
                                                                                     padx=10, pady=5)

        ttk.Label(self.scrollable_summary_frame, text="SOLD:", font=("Arial", 10, "bold")).grid(row=2, column=0,
                                                                                                sticky="w", padx=10,
                                                                                                pady=5)
        ttk.Label(self.scrollable_summary_frame, text=f"{total_sold} items").grid(row=2, column=1, sticky="w", padx=10,
                                                                                  pady=5)

        ttk.Label(self.scrollable_summary_frame, text="REMOVED:", font=("Arial", 10, "bold")).grid(row=3, column=0,
                                                                                                   sticky="w", padx=10,
                                                                                                   pady=5)
        ttk.Label(self.scrollable_summary_frame, text=f"{total_removed} items").grid(row=3, column=1, sticky="w",
                                                                                     padx=10, pady=5)

        ttk.Label(self.scrollable_summary_frame, text="FINANCIAL SUMMARY", font=("Arial", 12, "bold")).grid(row=4,
                                                                                                            column=0,
                                                                                                            columnspan=2,
                                                                                                            pady=10)

        ttk.Label(self.scrollable_summary_frame, text="Total Inventory Cost:", font=("Arial", 10, "bold")).grid(row=5,
                                                                                                                column=0,
                                                                                                                sticky="w",
                                                                                                                padx=10,
                                                                                                                pady=5)
        ttk.Label(self.scrollable_summary_frame, text=f"Rs. {total_cost:,.2f}").grid(row=5, column=1, sticky="w",
                                                                                     padx=10, pady=5)

        ttk.Label(self.scrollable_summary_frame, text="Total Seller Value:", font=("Arial", 10, "bold")).grid(row=6,
                                                                                                              column=0,
                                                                                                              sticky="w",
                                                                                                              padx=10,
                                                                                                              pady=5)
        ttk.Label(self.scrollable_summary_frame, text=f"Rs. {seller_value:,.2f}").grid(row=6, column=1, sticky="w",
                                                                                       padx=10, pady=5)

        ttk.Label(self.scrollable_summary_frame, text="Total Revenue:", font=("Arial", 10, "bold")).grid(row=7,
                                                                                                         column=0,
                                                                                                         sticky="w",
                                                                                                         padx=10,
                                                                                                         pady=5)
        ttk.Label(self.scrollable_summary_frame, text=f"Rs. {total_revenue:,.2f}").grid(row=7, column=1, sticky="w",
                                                                                        padx=10, pady=5)

        ttk.Label(self.scrollable_summary_frame, text="Total Profit Earned:", font=("Arial", 10, "bold")).grid(row=8,
                                                                                                               column=0,
                                                                                                               sticky="w",
                                                                                                               padx=10,
                                                                                                               pady=5)
        ttk.Label(self.scrollable_summary_frame, text=f"Rs. {profit:,.2f}").grid(row=8, column=1, sticky="w", padx=10,
                                                                                 pady=5)

        ttk.Label(self.scrollable_summary_frame, text="Total Orders:", font=("Arial", 10, "bold")).grid(row=9,
                                                                                                        column=0,
                                                                                                        sticky="w",
                                                                                                        padx=10,
                                                                                                        pady=5)
        ttk.Label(self.scrollable_summary_frame, text=f"{total_orders} orders").grid(row=9, column=1, sticky="w",
                                                                                     padx=10, pady=5)

        # Sales details table REMOVED as requested


if __name__ == "__main__":
    init_files()
    app = StockApp()
    app.mainloop()