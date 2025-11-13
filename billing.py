"""
Billing Software - All-in-One Themed Edition with Printer Select (Single file)
- Simple text receipt printing (chosen)
- Printer combobox (Windows enumeration when available)
- Themes, product suggestions, PDF save/print, read/merge PDFs
Requirements:
    pip install openpyxl reportlab PyPDF2 pywin32
Note:
    win32print is optional (Windows). Program runs on other OSes but printing features may be limited.
Author: Generated for user
"""

# ------------------- Imports -------------------
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os
import sys
import tempfile

from openpyxl import Workbook, load_workbook

# win32print used for printer enumeration & printing on Windows (optional)
try:
    import win32print
except Exception:
    win32print = None

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from PyPDF2 import PdfReader, PdfWriter

# ------------------- Configuration -------------------
PRODUCT_FILE = "product_master.xlsx"
BILLS_FILE = "bills.xlsx"
CUSTOMER_FILE = "customers.xlsx"
BILL_COUNTER_FILE = "bill_counter.txt"
PDF_FOLDER = "bills_pdf"

SHOP_NAME = "P.MUTHUGANESAN NADAR TEXTILE AND READYMADE"
SHOP_ADDRESS = "103b kamachi amman sanathi street east raja veethi kanchipuram"
SHOP_PHONE = "04447791355 / 9944369227"

DEFAULT_LOW_STOCK_THRESHOLD = 100
DEFAULT_GST_PERCENT = 18.0

# ------------------- Helpers -------------------
def ensure_pdf_folder(folder=PDF_FOLDER):
    if not os.path.exists(folder):
        os.makedirs(folder)
    return folder

def get_next_bill_no():
    """Read bill counter file, increment, save and return (string, int)."""
    if os.path.exists(BILL_COUNTER_FILE):
        with open(BILL_COUNTER_FILE, "r") as f:
            try:
                last_no = int(f.read().strip())
            except:
                last_no = 0
    else:
        last_no = 0
    next_no = last_no + 1
    with open(BILL_COUNTER_FILE, "w") as f:
        f.write(str(next_no))
    return f"BILL-{next_no:06d}", next_no

def get_next_pdf_filename():
    ensure_pdf_folder()
    existing = [f for f in os.listdir(PDF_FOLDER) if f.endswith(".pdf")]
    if not existing:
        return os.path.join(PDF_FOLDER, "001.pdf")
    numbers = [int(os.path.splitext(f)[0]) for f in existing if os.path.splitext(f)[0].isdigit()]
    next_num = max(numbers) + 1 if numbers else 1
    return os.path.join(PDF_FOLDER, f"{next_num:03d}.pdf")

def read_pdf(file_path):
    if not os.path.exists(file_path):
        messagebox.showerror("Error", f"{file_path} does not exist")
        return ""
    reader = PdfReader(file_path)
    text = ""
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text

def merge_pdfs(pdf_list, output_file):
    writer = PdfWriter()
    for pdf_path in pdf_list:
        if not os.path.exists(pdf_path):
            continue
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            writer.add_page(page)
    with open(output_file, "wb") as f:
        writer.write(f)
    messagebox.showinfo("Success", f"Merged PDFs saved as {output_file}")

# ------------------- Themes -------------------
THEMES = {
    "Light": {
        "bg": "#f7f9fc",
        "panel": "#eaf0fb",
        "accent": "#1976d2",
        "text": "#1f2937",
        "tree_even": "#ffffff",
        "tree_odd": "#f3f6fb",
        "tree_low": "#ffefef",
        "preview_bg": "#ffffff",
        "font": ("Segoe UI", 10),
        "mono": ("Consolas", 10)
    },
    "Dark": {
        "bg": "#0f1724",
        "panel": "#0b1220",
        "accent": "#1e88e5",
        "text": "#e6eef8",
        "tree_even": "#0b1220",
        "tree_odd": "#0f1626",
        "tree_low": "#3b0f0f",
        "preview_bg": "#0b1220",
        "font": ("Segoe UI", 10),
        "mono": ("Consolas", 10)
    },
    "Warm": {
        "bg": "#fcfbf8",
        "panel": "#f6efe6",
        "accent": "#ef6c00",
        "text": "#3b2f2f",
        "tree_even": "#fffaf5",
        "tree_odd": "#fff3e0",
        "tree_low": "#ffddd2",
        "preview_bg": "#fffaf0",
        "font": ("Segoe UI", 10),
        "mono": ("Consolas", 10)
    }
}
THEME_ORDER = ["Light", "Dark", "Warm"]

# ------------------- BillingApp -------------------
class BillingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Billing Software")
        try:
            self.root.state("zoomed")
        except:
            pass

        # Core data
        self.sub_total = 0.0
        self.gst_percent = tk.DoubleVar(value=DEFAULT_GST_PERCENT)
        self.total = 0.0
        self.paid_amount = tk.DoubleVar(value=0.0)
        self.due_amount = tk.DoubleVar(value=0.0)
        self.low_stock_threshold = DEFAULT_LOW_STOCK_THRESHOLD
        self.products = self.load_products()
        self.items_in_bill = []

        # Theme
        self.current_theme_name = "Light"
        self.current_theme = THEMES[self.current_theme_name]

        # Style
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except:
            pass

        # Build UI
        self.build_ui()

        # Shortcuts
        self.root.bind("<Delete>", lambda e: self.delete_item())
        self.root.bind("<Control-d>", lambda e: self.delete_item())
        self.root.bind("<Control-p>", lambda e: self.show_print_tab())
        self.root.bind("<Control-n>", lambda e: self.new_client())
        self.root.bind("<Control-c>", lambda e: self.clear_all())

        # Auto-save stub
        self.auto_save_interval = 5000
        self.root.after(self.auto_save_interval, self.auto_save_bill)

    # ---------------- UI build ----------------
    def build_ui(self):
        self.root.configure(bg=self.current_theme["bg"])

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Main & Print frames
        self.main_frame = tk.Frame(self.notebook, bg=self.current_theme["bg"])
        self.print_frame = tk.Frame(self.notebook, bg=self.current_theme["bg"])
        self.notebook.add(self.main_frame, text="🧾 Billing")
        self.notebook.add(self.print_frame, text="🖨️ Print Bill")

        # Main frame sections
        self.build_customer_frame()
        self.build_item_frame()
        self.build_toolbar()
        self.build_treeview()
        self.build_bottom_frame()

        # Print frame
        self.build_print_preview_with_printer_select()

        # Apply theme
        self.apply_theme(self.current_theme_name)

    def build_customer_frame(self):
        cust_frame = tk.LabelFrame(self.main_frame, text="Customer Details", padx=10, pady=10)
        cust_frame.pack(fill="x", padx=10, pady=8)
        tk.Label(cust_frame, text="Name:").grid(row=0, column=0, sticky="w")
        self.cust_name = tk.Entry(cust_frame, width=30)
        self.cust_name.grid(row=0, column=1, padx=6)
        self.cust_name.insert(0, "Customer")
        tk.Label(cust_frame, text="Phone:").grid(row=0, column=2, sticky="w")
        self.cust_phone = tk.Entry(cust_frame, width=20)
        self.cust_phone.grid(row=0, column=3, padx=6)
        ttk.Button(cust_frame, text="Save Customer", command=self.save_customer).grid(row=0, column=4, padx=12)

    def build_item_frame(self):
        item_frame = tk.LabelFrame(self.main_frame, text="Add Item", padx=10, pady=10)
        item_frame.pack(fill="x", padx=10, pady=6)
        tk.Label(item_frame, text="Item:").grid(row=0, column=0, sticky="w")
        self.item_name = tk.Entry(item_frame, width=40)
        self.item_name.grid(row=0, column=1, padx=6)
        self.item_name.bind("<KeyRelease>", self.show_suggestions)
        self.item_name.bind("<Down>", self.move_down)
        self.item_name.bind("<Up>", self.move_up)
        self.item_name.bind("<Escape>", lambda e: self.suggestion_box.grid_remove())
        self.item_name.bind("<Return>", self.enter_in_item_name)

        self.suggestion_box = tk.Listbox(item_frame, height=5)
        self.suggestion_box.grid(row=1, column=1, sticky="w", padx=6)
        self.suggestion_box.bind("<<ListboxSelect>>", self.fill_from_suggestion)
        self.suggestion_box.bind("<Return>", self.fill_from_suggestion)
        self.suggestion_box.grid_remove()

        tk.Label(item_frame, text="MRP:").grid(row=0, column=2, sticky="w")
        self.item_mrp = tk.Entry(item_frame, width=10)
        self.item_mrp.grid(row=0, column=3, padx=6)
        tk.Label(item_frame, text="Rate:").grid(row=0, column=4, sticky="w")
        self.item_rate = tk.Entry(item_frame, width=10)
        self.item_rate.grid(row=0, column=5, padx=6)
        tk.Label(item_frame, text="Disc%:").grid(row=0, column=6, sticky="w")
        self.item_discount = tk.Entry(item_frame, width=8)
        self.item_discount.grid(row=0, column=7, padx=6)
        tk.Label(item_frame, text="Qty:").grid(row=0, column=8, sticky="w")
        self.item_qty = tk.Entry(item_frame, width=6)
        self.item_qty.grid(row=0, column=9, padx=6)
        self.item_qty.bind("<Return>", lambda e: self.add_item())
        ttk.Button(item_frame, text="Add Item", command=self.add_item).grid(row=0, column=10, padx=12)

    def build_toolbar(self):
        toolbar = tk.Frame(self.main_frame, pady=6)
        toolbar.pack(fill="x", padx=10)
        ttk.Button(toolbar, text="Add (Enter)", command=self.add_item).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Delete (Del)", command=self.delete_item).pack(side="left", padx=4)
        ttk.Button(toolbar, text="New Client (Ctrl+N)", command=self.new_client).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Print (Ctrl+P)", command=self.show_print_tab).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Remove Low Stock", command=self.delete_low_stock_items).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Read PDF", command=self.open_read_pdf).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Merge PDFs", command=self.open_merge_pdf).pack(side="left", padx=4)

        right_frame = tk.Frame(toolbar)
        right_frame.pack(side="right")
        tk.Label(right_frame, text="LowStockThreshold:").pack(side="left", padx=6)
        self.threshold_entry = tk.Entry(right_frame, width=6)
        self.threshold_entry.pack(side="left")
        self.threshold_entry.insert(0, str(self.low_stock_threshold))
        ttk.Button(right_frame, text="Set Threshold", command=self.set_low_stock_threshold).pack(side="left", padx=6)
        tk.Label(right_frame, text="GST %:").pack(side="left", padx=6)
        tk.Entry(right_frame, textvariable=self.gst_percent, width=6).pack(side="left")
        tk.Label(right_frame, text="Theme:").pack(side="left", padx=(12,4))
        self.theme_var = tk.StringVar(value=self.current_theme_name)
        ttk.OptionMenu(right_frame, self.theme_var, self.current_theme_name, *THEMES.keys(), command=self.on_theme_select).pack(side="left")
        ttk.Button(right_frame, text="Cycle Theme", command=self.cycle_theme).pack(side="left", padx=6)

    def build_treeview(self):
        columns = ("Item","MRP","Rate","Discount","Qty","Total")
        self.tree = ttk.Treeview(self.main_frame, columns=columns, show="headings", height=14)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=300 if col=="Item" else 90, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)
        self.tree.tag_configure("evenrow")
        self.tree.tag_configure("oddrow")
        self.tree.tag_configure("lowstock")

    def build_bottom_frame(self):
        bottom_frame = tk.Frame(self.main_frame)
        bottom_frame.pack(fill="x", padx=10, pady=(0,10))
        ttk.Button(bottom_frame, text="Clear All (Ctrl+C)", command=self.clear_all).pack(side="left")
        self.sub_label = tk.Label(bottom_frame, text="SubTotal: 0.00", font=self.current_theme.get("font"))
        self.sub_label.pack(side="right", padx=10)
        self.gst_label = tk.Label(bottom_frame, text="GST: 0.00", font=self.current_theme.get("font"))
        self.gst_label.pack(side="right", padx=10)
        self.total_label = tk.Label(bottom_frame, text="Total: 0.00", font=self.current_theme.get("font"))
        self.total_label.pack(side="right", padx=10)
        self.item_count_label = tk.Label(bottom_frame, text="Items: 0", font=self.current_theme.get("font"))
        self.item_count_label.pack(side="right", padx=10)
        tk.Label(bottom_frame, text="Paid:").pack(side="right", padx=6)
        paid_entry = tk.Entry(bottom_frame, textvariable=self.paid_amount, width=10)
        paid_entry.pack(side="right")
        paid_entry.bind("<KeyRelease>", lambda e: self.update_due())
        tk.Label(bottom_frame, text="Due:").pack(side="right", padx=6)
        self.due_label = tk.Label(bottom_frame, text="0.00", font=self.current_theme.get("font"))
        self.due_label.pack(side="right", padx=6)

    def build_print_preview_with_printer_select(self):
        self.bill_preview = tk.Text(self.print_frame, font=self.current_theme.get("mono"), width=90, height=25)
        self.bill_preview.pack(fill="both", expand=True, padx=10, pady=(10,6))
        printer_frame = tk.Frame(self.print_frame)
        printer_frame.pack(fill="x", padx=10, pady=(0,6))
        tk.Label(printer_frame, text="Select Printer:").pack(side="left", padx=(0,6))
        self.printer_var = tk.StringVar()
        self.printer_combo = ttk.Combobox(printer_frame, textvariable=self.printer_var, state="readonly", width=60)
        self.printer_combo.pack(side="left", padx=(0,6))
        ttk.Button(printer_frame, text="Refresh Printers", command=self.refresh_printer_list).pack(side="left", padx=(6,0))
        btn_frame = tk.Frame(self.print_frame)
        btn_frame.pack(pady=(0,10))
        ttk.Button(btn_frame, text="Print", command=self.print_bill_to_printer).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Cancel", command=lambda: self.notebook.select(self.main_frame)).pack(side="left", padx=6)

    # ---------------- Theme handlers ----------------
    def apply_theme(self, theme_name):
        if theme_name not in THEMES:
            return
        self.current_theme_name = theme_name
        self.current_theme = THEMES[theme_name]
        t = self.current_theme
        try:
            self.root.configure(bg=t["bg"])
        except:
            pass
        try:
            self.style.configure("TNotebook", background=t["bg"])
            self.style.configure("TNotebook.Tab", font=t["font"], padding=[10,6])
            self.style.configure("TFrame", background=t["bg"])
            self.style.configure("TLabel", background=t["panel"], font=t["font"], foreground=t["text"])
            self.style.configure("TButton", font=t["font"])
        except:
            pass
        for widget in (self.main_frame, self.print_frame):
            try:
                widget.configure(bg=t["bg"])
            except:
                pass
        for child in self.main_frame.winfo_children():
            try:
                child.configure(bg=t["panel"])
            except:
                pass
        try:
            self.suggestion_box.configure(bg=t["tree_even"], fg=t["text"], font=t["font"])
        except:
            pass
        try:
            self.style.configure("Treeview", font=t["font"], foreground=t["text"])
            self.tree.tag_configure("evenrow", background=t["tree_even"])
            self.tree.tag_configure("oddrow", background=t["tree_odd"])
            self.tree.tag_configure("lowstock", background=t["tree_low"])
        except:
            pass
        for lbl in (self.sub_label, self.gst_label, self.total_label, self.item_count_label, self.due_label):
            try:
                lbl.configure(bg=t["bg"], fg=t["text"], font=t["font"])
            except:
                pass
        try:
            self.bill_preview.configure(bg=t["preview_bg"], fg=t["text"], font=t["mono"])
        except:
            pass
        entries = [self.cust_name, self.cust_phone, self.item_name, self.item_mrp, self.item_rate, self.item_discount, self.item_qty, self.threshold_entry]
        for e in entries:
            try:
                bg_color = "white" if theme_name == "Light" else t["panel"]
                e.configure(bg=bg_color, fg=t["text"], font=t["font"])
            except:
                pass
        try:
            self.theme_var.set(self.current_theme_name)
        except:
            pass
        self.refresh_tree()

    def cycle_theme(self):
        idx = THEME_ORDER.index(self.current_theme_name)
        next_idx = (idx + 1) % len(THEME_ORDER)
        self.apply_theme(THEME_ORDER[next_idx])

    def on_theme_select(self, selected):
        self.apply_theme(selected)

    # ---------------- Printers ----------------
    def refresh_printer_list(self):
        printers = []
        default_printer = None
        if win32print:
            try:
                raw = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
                for p in raw:
                    try:
                        pname = p[2]
                    except:
                        pname = str(p)
                    printers.append(pname)
                try:
                    default_printer = win32print.GetDefaultPrinter()
                except:
                    default_printer = None
            except:
                printers = []
                default_printer = None
        if not printers:
            if default_printer:
                printers = [default_printer]
            else:
                printers = ["<No printers found>"]
        try:
            self.printer_combo["values"] = printers
            if default_printer and default_printer in printers:
                self.printer_var.set(default_printer)
            else:
                self.printer_var.set(printers[0])
        except:
            pass

    # ---------------- Products / Suggestions ----------------
    def load_products(self):
        products = {}
        if os.path.exists(PRODUCT_FILE):
            try:
                wb = load_workbook(PRODUCT_FILE)
                ws = wb.active
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and row[0]:
                        products[str(row[0]).lower()] = {"mrp": row[1] or 0, "rate": row[2] or 0, "discount": row[3] or 0}
            except:
                products = {}
        return products

    def show_suggestions(self, event):
        val = self.item_name.get().lower()
        if not val:
            self.suggestion_box.grid_remove()
            return
        matches = [p for p in self.products if val in p]
        if matches:
            self.suggestion_box.delete(0, tk.END)
            for m in matches[:20]:
                self.suggestion_box.insert(tk.END, m)
            self.suggestion_box.grid()
        else:
            self.suggestion_box.grid_remove()

    def fill_from_suggestion(self, event=None):
        sel = self.suggestion_box.curselection()
        if sel:
            val = self.suggestion_box.get(sel)
            self.item_name.delete(0, tk.END)
            self.item_name.insert(0, val)
            prod = self.products.get(val.lower())
            if prod:
                self.item_mrp.delete(0, tk.END); self.item_mrp.insert(0, prod["mrp"])
                self.item_rate.delete(0, tk.END); self.item_rate.insert(0, prod["rate"])
                self.item_discount.delete(0, tk.END); self.item_discount.insert(0, prod["discount"])
            self.suggestion_box.grid_remove()
            self.item_qty.focus()

    def move_down(self, event):
        if self.suggestion_box.size() > 0:
            self.suggestion_box.focus()
            self.suggestion_box.selection_set(0)
            self.suggestion_box.activate(0)

    def move_up(self, event):
        if self.suggestion_box.size() > 0:
            self.suggestion_box.focus()
            self.suggestion_box.selection_set(self.suggestion_box.size()-1)
            self.suggestion_box.activate(self.suggestion_box.size()-1)

    def enter_in_item_name(self, event):
        if self.suggestion_box.winfo_ismapped():
            self.fill_from_suggestion()
        else:
            self.add_item()

    # ---------------- Billing actions ----------------
    def add_item(self):
        try:
            name = self.item_name.get().strip()
            mrp = float(self.item_mrp.get() or 0)
            rate = float(self.item_rate.get() or 0)
            disc = float(self.item_discount.get() or 0)
            qty = int(self.item_qty.get() or 0)
        except Exception:
            messagebox.showerror("Error","Invalid item details")
            return
        if not name:
            messagebox.showerror("Error","Item name required")
            return
        total = round(rate * qty * (1 - disc/100), 2)
        self.items_in_bill.append({"name":name,"mrp":mrp,"rate":rate,"discount":disc,"qty":qty,"total":total})
        self.update_totals()
        self.refresh_tree()
        self.item_name.delete(0, tk.END)
        self.item_mrp.delete(0, tk.END)
        self.item_rate.delete(0, tk.END)
        self.item_discount.delete(0, tk.END)
        self.item_qty.delete(0, tk.END)

    def delete_item(self):
        sel = self.tree.selection()
        if not sel:
            return
        try:
            index = int(sel[0])
        except:
            for s in sel:
                try:
                    self.tree.delete(s)
                except:
                    pass
            return
        if 0 <= index < len(self.items_in_bill):
            del self.items_in_bill[index]
        self.update_totals()
        self.refresh_tree()

    def clear_all(self):
        self.items_in_bill.clear()
        self.update_totals()
        self.refresh_tree()
        try:
            self.cust_name.delete(0, tk.END)
            self.cust_name.insert(0, "Customer")
            self.cust_phone.delete(0, tk.END)
        except:
            pass

    def new_client(self):
        self.clear_all()

    def update_totals(self):
        self.sub_total = sum(i["total"] for i in self.items_in_bill)
        gst_amt = round(self.sub_total * self.gst_percent.get() / 100, 2)
        self.total = self.sub_total + gst_amt
        self.sub_label.config(text=f"SubTotal: {self.sub_total:.2f}")
        self.gst_label.config(text=f"GST: {gst_amt:.2f}")
        self.total_label.config(text=f"Total: {self.total:.2f}")
        self.item_count_label.config(text=f"Items: {len(self.items_in_bill)}")
        self.update_due()

    def update_due(self):
        paid = self.paid_amount.get()
        due = self.total - paid
        self.due_label.config(text=f"{due:.2f}")

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for idx, item in enumerate(self.items_in_bill):
            tag = "lowstock" if item["qty"] <= self.low_stock_threshold else ("evenrow" if idx%2==0 else "oddrow")
            self.tree.insert("", tk.END, iid=str(idx), values=(item["name"], item["mrp"], item["rate"], item["discount"], item["qty"], item["total"]), tags=(tag,))

    # ---------------- Customer persistence ----------------
    def save_customer(self):
        name = self.cust_name.get().strip()
        phone = self.cust_phone.get().strip()
        if not name or not phone:
            return
        wb = Workbook()
        if os.path.exists(CUSTOMER_FILE):
            wb = load_workbook(CUSTOMER_FILE)
        ws = wb.active
        ws.append([name, phone, str(datetime.now())])
        wb.save(CUSTOMER_FILE)
        messagebox.showinfo("Saved", "Customer saved")

    def delete_low_stock_items(self):
        self.items_in_bill = [i for i in self.items_in_bill if i["qty"] > self.low_stock_threshold]
        self.update_totals()
        self.refresh_tree()

    def set_low_stock_threshold(self):
        try:
            self.low_stock_threshold = int(self.threshold_entry.get())
            messagebox.showinfo("Threshold","Low stock threshold updated")
            self.refresh_tree()
        except:
            messagebox.showerror("Error","Invalid threshold")

    # ---------------- PDF saving ----------------
    def print_and_save_pdf(self):
        bill_no, _ = get_next_bill_no()
        pdf_file = get_next_pdf_filename()
        try:
            c = canvas.Canvas(pdf_file, pagesize=A4)
            c.setFont("Helvetica-Bold", 16)
            c.drawString(60, 800, SHOP_NAME)
            c.setFont("Helvetica", 10)
            c.drawString(60, 785, SHOP_ADDRESS)
            c.drawString(60, 770, SHOP_PHONE)
            c.setFont("Helvetica", 10)
            c.drawString(400, 740, f"Bill No: {bill_no}")
            c.drawString(400, 725, f"Date: {datetime.now().strftime('%d-%m-%Y %H:%M')}")
            data = [["Item","MRP","Rate","Disc%","Qty","Total"]]
            for it in self.items_in_bill:
                data.append([it["name"], it["mrp"], it["rate"], it["discount"], it["qty"], it["total"]])
            table = Table(data, colWidths=[200,50,50,50,40,60])
            table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.5,colors.black),
                                       ("FONT",(0,0),(-1,0),"Helvetica-Bold")]))
            table.wrapOn(c, 40, 600)
            table.drawOn(c, 40, 600 - len(data)*18)
            y = 600 - len(data)*18 - 20
            c.drawString(40, y, f"SubTotal: {self.sub_total:.2f}")
            y -= 14
            gst_amt = self.sub_total * self.gst_percent.get() / 100
            c.drawString(40, y, f"GST: {gst_amt:.2f}")
            y -= 14
            c.drawString(40, y, f"Total: {self.total:.2f}")
            c.save()
            messagebox.showinfo("PDF Saved", f"Bill saved as {pdf_file}")
        except Exception as e:
            messagebox.showerror("PDF Error", f"Failed to create PDF: {e}")

    def open_read_pdf(self):
        file = filedialog.askopenfilename(filetypes=[("PDF Files","*.pdf")])
        if file:
            text = read_pdf(file)
            win = tk.Toplevel(self.root)
            win.title("PDF Content")
            txt = tk.Text(win, width=100, height=30)
            txt.pack(fill="both", expand=True)
            txt.insert(tk.END, text)

    def open_merge_pdf(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files","*.pdf")])
        if files:
            out = filedialog.asksaveasfilename(defaultextension=".pdf")
            if out:
                merge_pdfs(files, out)

    # ---------------- Print preview & printing (TEXT receipt) ----------------
    def show_print_tab(self):
        self.refresh_printer_list()
        self.refresh_print_bill()
        self.notebook.select(self.print_frame)

    def refresh_print_bill(self):
        # Original behavior: increments bill counter (keeps as-is)
        self.bill_preview.delete("1.0", tk.END)
        bill_no, _ = get_next_bill_no()
        lines = []
        lines.append(f"{SHOP_NAME}\n")
        lines.append(f"{SHOP_ADDRESS}\n")
        lines.append(f"{SHOP_PHONE}\n")
        lines.append("="*60 + "\n")
        lines.append(f"Bill No: {bill_no:<10} Date: {datetime.now().strftime('%d-%m-%Y %H:%M')}\n")
        lines.append("="*60 + "\n")
        lines.append(f"{'Item':20}{'MRP':>6}{'Rate':>7}{'Disc%':>7}{'Qty':>6}{'Total':>9}\n")
        lines.append("-"*60 + "\n")
        for i in self.items_in_bill:
            lines.append(f"{i['name'][:20]:20}{i['mrp']:>6.2f}{i['rate']:>7.2f}{i['discount']:>7.2f}{i['qty']:>6}{i['total']:>9.2f}\n")
        lines.append("-"*60 + "\n")
        lines.append(f"SubTotal: {self.sub_total:.2f}\n")
        gst_amt = self.sub_total * self.gst_percent.get() / 100
        lines.append(f"GST: {gst_amt:.2f}\n")
        lines.append(f"Total: {self.total:.2f}\n")
        paid = self.paid_amount.get()
        due = self.total - paid
        lines.append(f"Paid: {paid:.2f}  Due: {due:.2f}\n")
        lines.append("="*60 + "\n")
        self.bill_preview.insert(tk.END, "".join(lines))

    def print_bill_to_printer(self):
        """Create a temporary PDF from the text preview and send to selected printer (or open)."""
        text = self.bill_preview.get("1.0", tk.END)
        if not text.strip():
            messagebox.showerror("Error", "Nothing to print")
            return

        # Build temporary PDF (simple text layout to preserve alignment)
        try:
            tmp_pdf = tempfile.mktemp(suffix=".pdf")
            c = canvas.Canvas(tmp_pdf, pagesize=A4)
            width, height = A4
            margin = 40
            y = height - margin

            # Header
            c.setFont("Helvetica-Bold", 14)
            c.drawString(margin, y, SHOP_NAME)
            y -= 18
            c.setFont("Helvetica", 10)
            c.drawString(margin, y, SHOP_ADDRESS)
            y -= 14
            c.drawString(margin, y, SHOP_PHONE)
            y -= 20
            c.setFont("Helvetica", 9)
            c.drawString(margin, y, "-" * 95)
            y -= 16

            # Write lines with wrapping/paging
            lines = text.splitlines()
            c.setFont("Courier", 9)
            line_height = 12
            max_lines = int((y - margin) / line_height) or 40
            idx = 0
            while idx < len(lines):
                for _ in range(max_lines):
                    if idx >= len(lines):
                        break
                    line = lines[idx]
                    # naive wrap at ~100 chars
                    while len(line) > 0:
                        piece = line[:110]
                        c.drawString(margin, y, piece)
                        y -= line_height
                        line = line[110:]
                        if y < margin + line_height:
                            break
                    if y < margin + line_height:
                        break
                    idx += 1
                if idx < len(lines):
                    c.showPage()
                    y = height - margin
                    c.setFont("Courier", 9)
            c.save()
        except Exception as e:
            messagebox.showerror("PDF Error", f"Failed to create PDF for printing: {e}")
            return

        # Determine selected printer
        try:
            selected = self.printer_var.get()
            if selected == "<No printers found>":
                selected = None
        except:
            selected = None

        # Windows printing via ShellExecute
        if win32print and selected:
            try:
                win32print.ShellExecute(0, "print", tmp_pdf, f'/d:"{selected}"', ".", 0)
                messagebox.showinfo("Printed", f"Bill sent to printer: {selected}")
                return
            except Exception as e:
                messagebox.showwarning("Print Error", f"Could not print to '{selected}': {e}\nTrying default printer...")

        if win32print:
            try:
                default = win32print.GetDefaultPrinter()
                win32print.ShellExecute(0, "print", tmp_pdf, f'/d:"{default}"', ".", 0)
                messagebox.showinfo("Printed", f"Bill sent to default printer: {default}")
                return
            except Exception:
                # fallback: open the PDF for manual printing
                try:
                    os.startfile(tmp_pdf)
                    messagebox.showinfo("PDF Ready", f"PDF opened for printing: {tmp_pdf}")
                except Exception:
                    messagebox.showinfo("PDF Saved", f"PDF saved to: {tmp_pdf}")
                return

        # Non-windows fallback: open PDF so user can print manually
        try:
            if sys.platform == "darwin":
                os.system(f'open "{tmp_pdf}"')
            elif os.name == "nt":
                os.startfile(tmp_pdf)
            else:
                os.system(f'xdg-open "{tmp_pdf}"')
            messagebox.showinfo("PDF Ready", f"PDF saved and opened for manual printing:\n{tmp_pdf}")
        except Exception:
            messagebox.showinfo("PDF Saved", f"PDF saved to: {tmp_pdf}\nPlease open and print manually.")

    # ---------------- Auto-save stub ----------------
    def auto_save_bill(self):
        # Placeholder for future auto-save; preserved as a stub.
        self.root.after(self.auto_save_interval, self.auto_save_bill)

# ---------------- Run App ----------------
if __name__ == "__main__":
    root = tk.Tk()
    app = BillingApp(root)
    app.apply_theme(app.current_theme_name)
    root.mainloop()
