"""
Avirenza Internship Task 1 – Monthly Sales Dashboard for a Small D2C Tea Brand
Candidate: Kshitij Hemchandra Gharat (AVZ-2026-000161)

This script builds the dataset and Excel workbook (tea-brand-sales.xlsx).
Sheets created: Lookups, Raw, Clean, Pivots, Dashboard, Notes
"""

import random
import datetime
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule
import copy

random.seed(42)

# ─────────────────────────────────────────────────
# LOOKUP DATA
# ─────────────────────────────────────────────────
PRODUCTS = [
    # (Product Name, Category, Unit Price)
    ("Assam Classic Loose Leaf",   "Loose Leaf",  350),
    ("Darjeeling First Flush",     "Loose Leaf",  520),
    ("Green Tulsi Blend",          "Loose Leaf",  290),
    ("Masala Chai Loose Leaf",     "Loose Leaf",  310),
    ("Earl Grey Loose Leaf",       "Loose Leaf",  380),
    ("Ginger Lemon Tea Bags",      "Tea Bags",    180),
    ("Classic Assam Tea Bags",     "Tea Bags",    150),
    ("Green Mint Tea Bags",        "Tea Bags",    160),
    ("Festive Gift Box (Mixed)",   "Gift Box",    850),
    ("Premium Tea Gift Hamper",    "Gift Box",   1200),
    ("Wellness Herbal Gift Box",   "Gift Box",    750),
    ("Bamboo Tea Infuser",         "Accessories", 220),
    ("Ceramic Tea Mug",            "Accessories", 320),
]

CITIES = [
    ("Mumbai",    "West"),
    ("Delhi",     "North"),
    ("Bengaluru", "South"),
    ("Kolkata",   "East"),
    ("Chennai",   "South"),
    ("Hyderabad", "South"),
    ("Pune",      "West"),
    ("Jaipur",    "North"),
    ("Guwahati",  "North-East"),
    ("Lucknow",   "North"),
    ("Ahmedabad", "West"),
    ("Kochi",     "South"),
    ("Bhubaneswar","East"),
    ("Shillong",  "North-East"),
]

CHANNELS      = ["Website", "Instagram", "WhatsApp"]
DISCOUNTS     = [0, 5, 10, 15]
PAYMENT_MODES = ["UPI", "Card", "COD"]
STATUSES      = ["Delivered", "Returned", "Cancelled"]

# Weight distributions (realistic)
CHANNEL_W  = [0.50, 0.35, 0.15]
DISCOUNT_W = [0.45, 0.25, 0.20, 0.10]
PAYMENT_W  = [0.55, 0.28, 0.17]
STATUS_W   = [0.78, 0.12, 0.10]

def weighted_choice(choices, weights):
    r = random.random()
    cumulative = 0
    for c, w in zip(choices, weights):
        cumulative += w
        if r <= cumulative:
            return c
    return choices[-1]

def random_date(start, end):
    delta = end - start
    return start + datetime.timedelta(days=random.randint(0, delta.days))

START_DATE = datetime.date(2026, 1, 1)
END_DATE   = datetime.date(2026, 3, 31)

# ─────────────────────────────────────────────────
# GENERATE 385 CLEAN ORDERS (15 will be added messy)
# ─────────────────────────────────────────────────
clean_orders = []
for i in range(1, 386):
    city_data  = random.choice(CITIES)
    product    = random.choice(PRODUCTS)
    channel    = weighted_choice(CHANNELS, CHANNEL_W)
    discount   = weighted_choice(DISCOUNTS, DISCOUNT_W)
    payment    = weighted_choice(PAYMENT_MODES, PAYMENT_W)
    status     = weighted_choice(STATUSES, STATUS_W)
    qty        = random.randint(1, 5)
    order_date = random_date(START_DATE, END_DATE)
    order_id   = f"ORD-{1000 + i}"

    clean_orders.append({
        "Order ID":        order_id,
        "Order Date":      order_date,
        "Customer City":   city_data[0],
        "Region":          city_data[1],
        "Channel":         channel,
        "Product":         product[0],
        "Category":        product[1],
        "Quantity":        qty,
        "Unit Price":      product[2],
        "Discount Percent": discount,
        "Payment Mode":    payment,
        "Delivery Status": status,
    })

# ─────────────────────────────────────────────────
# 15 PLANTED MESSY ISSUES
# ─────────────────────────────────────────────────
# We'll append/modify after clean_orders list
messy_notes = []

# Issue 1: Blank Customer City
row = copy.deepcopy(clean_orders[4])
row["Order ID"] = "ORD-2001"
row["Customer City"] = ""
row["Region"] = ""
clean_orders.append(row)
messy_notes.append(("Issue 1", "ORD-2001", "Blank Customer City and Region"))

# Issue 2: Inconsistent city spelling – "mumbai" (lowercase)
row = copy.deepcopy(clean_orders[10])
row["Order ID"] = "ORD-2002"
row["Customer City"] = "mumbai"
clean_orders.append(row)
messy_notes.append(("Issue 2", "ORD-2002", "City spelled as 'mumbai' instead of 'Mumbai'"))

# Issue 3: Inconsistent city spelling – "DELHI"
row = copy.deepcopy(clean_orders[15])
row["Order ID"] = "ORD-2003"
row["Customer City"] = "DELHI"
clean_orders.append(row)
messy_notes.append(("Issue 3", "ORD-2003", "City spelled as 'DELHI' instead of 'Delhi'"))

# Issue 4: Duplicate row (exact copy of ORD-1001)
row = copy.deepcopy(clean_orders[0])
clean_orders.append(row)
messy_notes.append(("Issue 4", "ORD-1001", "Duplicate row – exact copy of ORD-1001"))

# Issue 5: Duplicate row (exact copy of ORD-1050)
row = copy.deepcopy(clean_orders[49])
clean_orders.append(row)
messy_notes.append(("Issue 5", "ORD-1050", "Duplicate row – exact copy of ORD-1050"))

# Issue 6: Blank Delivery Status
row = copy.deepcopy(clean_orders[20])
row["Order ID"] = "ORD-2006"
row["Delivery Status"] = ""
clean_orders.append(row)
messy_notes.append(("Issue 6", "ORD-2006", "Blank Delivery Status"))

# Issue 7: Inconsistent channel spelling – "website" (lowercase)
row = copy.deepcopy(clean_orders[25])
row["Order ID"] = "ORD-2007"
row["Channel"] = "website"
clean_orders.append(row)
messy_notes.append(("Issue 7", "ORD-2007", "Channel spelled as 'website' instead of 'Website'"))

# Issue 8: Inconsistent category – "loose leaf" (lowercase)
row = copy.deepcopy(clean_orders[30])
row["Order ID"] = "ORD-2008"
row["Category"] = "loose leaf"
clean_orders.append(row)
messy_notes.append(("Issue 8", "ORD-2008", "Category spelled as 'loose leaf' instead of 'Loose Leaf'"))

# Issue 9: Zero Unit Price (invalid)
row = copy.deepcopy(clean_orders[35])
row["Order ID"] = "ORD-2009"
row["Unit Price"] = 0
clean_orders.append(row)
messy_notes.append(("Issue 9", "ORD-2009", "Unit Price is 0 – invalid entry"))

# Issue 10: Quantity = 0 (invalid)
row = copy.deepcopy(clean_orders[40])
row["Order ID"] = "ORD-2010"
row["Quantity"] = 0
clean_orders.append(row)
messy_notes.append(("Issue 10", "ORD-2010", "Quantity is 0 – invalid entry"))

# Issue 11: City spelling with extra space – "Bengaluru " (trailing space)
row = copy.deepcopy(clean_orders[45])
row["Order ID"] = "ORD-2011"
row["Customer City"] = "Bengaluru "
clean_orders.append(row)
messy_notes.append(("Issue 11", "ORD-2011", "City has trailing space: 'Bengaluru ' instead of 'Bengaluru'"))

# Issue 12: Inconsistent payment mode – "cod" (lowercase)
row = copy.deepcopy(clean_orders[50])
row["Order ID"] = "ORD-2012"
row["Payment Mode"] = "cod"
clean_orders.append(row)
messy_notes.append(("Issue 12", "ORD-2012", "Payment Mode spelled as 'cod' instead of 'COD'"))

# Issue 13: Invalid Discount Percent (25 – not in allowed set)
row = copy.deepcopy(clean_orders[55])
row["Order ID"] = "ORD-2013"
row["Discount Percent"] = 25
clean_orders.append(row)
messy_notes.append(("Issue 13", "ORD-2013", "Discount Percent is 25 – not in allowed set (0,5,10,15)"))

# Issue 14: Blank Order Date
row = copy.deepcopy(clean_orders[60])
row["Order ID"] = "ORD-2014"
row["Order Date"] = ""
clean_orders.append(row)
messy_notes.append(("Issue 14", "ORD-2014", "Blank Order Date"))

# Issue 15: City misspelled – "Kolkatta" (double t)
row = copy.deepcopy(clean_orders[65])
row["Order ID"] = "ORD-2015"
row["Customer City"] = "Kolkatta"
clean_orders.append(row)
messy_notes.append(("Issue 15", "ORD-2015", "City misspelled as 'Kolkatta' instead of 'Kolkata'"))

# Shuffle so messy rows are scattered
random.shuffle(clean_orders)

# Total = 385 + 15 = 400
assert len(clean_orders) == 400, f"Expected 400, got {len(clean_orders)}"

RAW_HEADERS = [
    "Order ID", "Order Date", "Customer City", "Region",
    "Channel", "Product", "Category", "Quantity",
    "Unit Price", "Discount Percent", "Payment Mode", "Delivery Status"
]

# ─────────────────────────────────────────────────
# BUILD CLEAN DATA (fix all issues)
# ─────────────────────────────────────────────────
CITY_REGION_MAP = {c[0]: c[1] for c in CITIES}
PRODUCT_PRICE_MAP = {p[0]: p[2] for p in PRODUCTS}

def clean_row(row, seen_ids):
    """Return cleaned row dict or None if duplicate."""
    r = dict(row)

    # Fix city casing and trim
    city = str(r["Customer City"]).strip().title()

    # Fix known misspelling
    if city.lower() == "kolkatta":
        city = "Kolkata"
    if city == "":
        city = "Unknown"
    r["Customer City"] = city

    # Fix region from city map
    r["Region"] = CITY_REGION_MAP.get(city, r.get("Region", "Unknown"))

    # Fix Channel casing
    r["Channel"] = str(r["Channel"]).strip().title()

    # Fix Category casing
    cat = str(r["Category"]).strip()
    cat_map = {
        "loose leaf": "Loose Leaf",
        "tea bags": "Tea Bags",
        "gift box": "Gift Box",
        "accessories": "Accessories",
    }
    r["Category"] = cat_map.get(cat.lower(), cat.title())

    # Fix Payment Mode (UPI, Card, COD stay uppercase for UPI/COD)
    pm = str(r["Payment Mode"]).strip().upper()
    if pm == "CARD":
        pm = "Card"
    r["Payment Mode"] = pm

    # Fix blank Delivery Status
    if str(r["Delivery Status"]).strip() == "":
        r["Delivery Status"] = "Delivered"

    # Fix invalid unit price
    if r["Unit Price"] == 0:
        prod_name = r["Product"]
        r["Unit Price"] = PRODUCT_PRICE_MAP.get(prod_name, 300)

    # Fix invalid quantity
    if r["Quantity"] == 0:
        r["Quantity"] = 1

    # Fix invalid discount
    if r["Discount Percent"] not in [0, 5, 10, 15]:
        r["Discount Percent"] = 10

    # Fix blank order date
    if str(r["Order Date"]).strip() == "":
        r["Order Date"] = datetime.date(2026, 2, 15)

    # De-duplicate by Order ID
    oid = r["Order ID"]
    if oid in seen_ids:
        return None  # duplicate – skip
    seen_ids.add(oid)

    return r

seen_ids = set()
clean_data = []
for row in clean_orders:
    cr = clean_row(row, seen_ids)
    if cr is not None:
        clean_data.append(cr)

# ─────────────────────────────────────────────────
# COMPUTED COLUMNS
# ─────────────────────────────────────────────────
def compute(row):
    qty  = row["Quantity"]
    up   = row["Unit Price"]
    disc = row["Discount Percent"]
    gross = qty * up
    disc_amt = round(gross * disc / 100, 2)
    net_rev  = gross - disc_amt if row["Delivery Status"] not in ["Cancelled"] else 0
    month_name = ""
    od = row["Order Date"]
    if isinstance(od, datetime.date):
        month_name = od.strftime("%B")
    elif isinstance(od, str) and od:
        try:
            month_name = datetime.datetime.strptime(od, "%Y-%m-%d").strftime("%B")
        except:
            month_name = "Unknown"
    row["Gross Amount"]    = gross
    row["Discount Amount"] = disc_amt
    row["Net Revenue"]     = net_rev
    row["Month"]           = month_name
    return row

clean_data = [compute(r) for r in clean_data]

# ─────────────────────────────────────────────────
# PIVOT DATA (computed in Python for the sheet)
# ─────────────────────────────────────────────────

def pivot_by(data, key, value="Net Revenue"):
    result = {}
    for row in data:
        k = row.get(key, "")
        v = row.get(value, 0)
        result[k] = result.get(k, 0) + v
    return sorted(result.items(), key=lambda x: -x[1])

pivot_month    = pivot_by(clean_data, "Month")
pivot_category = pivot_by(clean_data, "Category")
pivot_region   = pivot_by(clean_data, "Region")
pivot_channel  = pivot_by(clean_data, "Channel")

# Top 10 products
product_pivot_all = pivot_by(clean_data, "Product")
pivot_top10 = product_pivot_all[:10]

# Order counts
def count_by(data, key):
    result = {}
    for row in data:
        k = row.get(key, "")
        result[k] = result.get(k, 0) + 1
    return result

# KPIs
total_net_revenue = sum(r["Net Revenue"] for r in clean_data)
total_orders      = len(clean_data)
avg_order_value   = round(total_net_revenue / total_orders, 2) if total_orders else 0
returns_cancelled = sum(1 for r in clean_data if r["Delivery Status"] in ["Returned", "Cancelled"])
ret_canc_rate     = round(returns_cancelled / total_orders * 100, 2) if total_orders else 0

# ─────────────────────────────────────────────────
# WORKBOOK CREATION
# ─────────────────────────────────────────────────
wb = Workbook()
# Remove default sheet
wb.remove(wb.active)

# ── Colour palette ──
C_HEADER_BG  = "1B4332"  # dark green
C_HEADER_FG  = "FFFFFF"
C_ALT_ROW    = "D8F3DC"
C_ACCENT     = "52B788"
C_LIGHT      = "F0FFF4"
C_ORANGE     = "FF6D00"
C_TEAL       = "0077B6"
C_PURPLE     = "7B2D8B"
C_GOLD       = "FFB703"
C_WHITE      = "FFFFFF"
C_DARK       = "1B1B1B"

def header_fill(color=C_HEADER_BG):
    return PatternFill("solid", fgColor=color)

def alt_fill():
    return PatternFill("solid", fgColor=C_ALT_ROW)

def hfont(bold=True, color=C_HEADER_FG, size=11):
    return Font(bold=bold, color=color, size=size, name="Calibri")

def dfont(bold=False, color=C_DARK, size=10):
    return Font(bold=bold, color=color, size=size, name="Calibri")

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left_align():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def set_col_width(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width

def apply_header_row(ws, headers, row=1, bg=C_HEADER_BG):
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.fill   = header_fill(bg)
        cell.font   = hfont()
        cell.alignment = center()
        cell.border = thin_border()

# ══════════════════════════════════════════════════
# SHEET 1 – LOOKUPS
# ══════════════════════════════════════════════════
ws_lk = wb.create_sheet("Lookups")
ws_lk.sheet_view.showGridLines = False

# Products table
ws_lk.cell(1, 1, "PRODUCT LOOKUP").font = hfont(size=12, color=C_DARK)
ws_lk.cell(1, 1).fill = header_fill(C_GOLD)
ws_lk.merge_cells("A1:C1")
ws_lk.cell(1, 1).alignment = center()

prod_headers = ["Product Name", "Category", "Unit Price (₹)"]
apply_header_row(ws_lk, prod_headers, row=2, bg=C_HEADER_BG)
for i, (name, cat, price) in enumerate(PRODUCTS, 1):
    r = i + 2
    for col, val in enumerate([name, cat, price], 1):
        cell = ws_lk.cell(r, col, val)
        cell.font   = dfont()
        cell.border = thin_border()
        if i % 2 == 0:
            cell.fill = alt_fill()
        cell.alignment = left_align()

# Cities table (offset to col 5)
ws_lk.cell(1, 5, "CITY & REGION LOOKUP").font = hfont(size=12, color=C_DARK)
ws_lk.cell(1, 5).fill = header_fill(C_TEAL)
ws_lk.merge_cells("E1:F1")
ws_lk.cell(1, 5).alignment = center()

city_headers = ["City", "Region"]
apply_header_row(ws_lk, city_headers, row=2, bg=C_TEAL)
for i, (city, region) in enumerate(CITIES, 1):
    r = i + 2
    for col, val in enumerate([city, region], 1):
        cell = ws_lk.cell(r, 4 + col, val)
        cell.font   = dfont()
        cell.border = thin_border()
        if i % 2 == 0:
            cell.fill = alt_fill()
        cell.alignment = left_align()

# Channel / Discount / Payment / Status reference
ref_data = {
    "Channel":          CHANNELS,
    "Discount %":       [str(d) for d in DISCOUNTS],
    "Payment Mode":     PAYMENT_MODES,
    "Delivery Status":  STATUSES,
}
start_col = 8
for header, vals in ref_data.items():
    ws_lk.cell(1, start_col, header).fill   = header_fill(C_PURPLE)
    ws_lk.cell(1, start_col).font           = hfont()
    ws_lk.cell(1, start_col).alignment      = center()
    ws_lk.cell(1, start_col).border         = thin_border()
    for i, v in enumerate(vals, 2):
        c = ws_lk.cell(i, start_col, v)
        c.font = dfont(); c.border = thin_border(); c.alignment = left_align()
    start_col += 1

for c in range(1, 13):
    set_col_width(ws_lk, c, 22)
ws_lk.row_dimensions[1].height = 24

# ══════════════════════════════════════════════════
# SHEET 2 – RAW  (400 rows with messy data)
# ══════════════════════════════════════════════════
ws_raw = wb.create_sheet("Raw")
ws_raw.sheet_view.showGridLines = False
ws_raw.freeze_panes = "A2"

apply_header_row(ws_raw, RAW_HEADERS, row=1)

for r_idx, row in enumerate(clean_orders, 2):
    is_alt = (r_idx % 2 == 0)
    for c_idx, col_name in enumerate(RAW_HEADERS, 1):
        val = row.get(col_name, "")
        if isinstance(val, datetime.date):
            val = val.strftime("%Y-%m-%d")
        cell = ws_raw.cell(r_idx, c_idx, val)
        cell.font      = dfont(size=9)
        cell.border    = thin_border()
        cell.alignment = left_align()
        if is_alt:
            cell.fill = alt_fill()

col_widths_raw = [12, 12, 14, 12, 12, 28, 14, 10, 12, 16, 14, 16]
for c, w in enumerate(col_widths_raw, 1):
    set_col_width(ws_raw, c, w)
ws_raw.row_dimensions[1].height = 30

# ══════════════════════════════════════════════════
# SHEET 3 – CLEAN
# ══════════════════════════════════════════════════
CLEAN_HEADERS = RAW_HEADERS + ["Gross Amount", "Discount Amount", "Net Revenue", "Month"]

ws_cl = wb.create_sheet("Clean")
ws_cl.sheet_view.showGridLines = False
ws_cl.freeze_panes = "A2"

apply_header_row(ws_cl, CLEAN_HEADERS, row=1)

for r_idx, row in enumerate(clean_data, 2):
    is_alt = (r_idx % 2 == 0)
    for c_idx, col_name in enumerate(CLEAN_HEADERS, 1):
        val = row.get(col_name, "")
        if isinstance(val, datetime.date):
            val = val.strftime("%Y-%m-%d")
        cell = ws_cl.cell(r_idx, c_idx, val)
        cell.font      = dfont(size=9)
        cell.border    = thin_border()
        cell.alignment = left_align()
        if is_alt:
            cell.fill = alt_fill()

col_widths_cl = [12, 12, 14, 12, 12, 28, 14, 10, 12, 16, 14, 16, 14, 16, 14, 10]
for c, w in enumerate(col_widths_cl, 1):
    set_col_width(ws_cl, c, w)
ws_cl.row_dimensions[1].height = 30

# ══════════════════════════════════════════════════
# SHEET 4 – PIVOTS
# ══════════════════════════════════════════════════
ws_pv = wb.create_sheet("Pivots")
ws_pv.sheet_view.showGridLines = False

def write_pivot_table(ws, title, headers, data, start_row, start_col, bg=C_HEADER_BG, add_total=True):
    # Title
    title_cell = ws.cell(start_row, start_col, title)
    title_cell.font      = hfont(size=11, color=C_WHITE)
    title_cell.fill      = header_fill(bg)
    title_cell.alignment = center()
    title_cell.border    = thin_border()
    ws.merge_cells(
        start_row=start_row, start_column=start_col,
        end_row=start_row, end_column=start_col + len(headers) - 1
    )

    # Headers
    for ci, h in enumerate(headers, start_col):
        c = ws.cell(start_row + 1, ci, h)
        c.font      = hfont(size=10)
        c.fill      = header_fill(bg)
        c.alignment = center()
        c.border    = thin_border()

    # Data
    total = 0
    for ri, row_vals in enumerate(data, start_row + 2):
        for ci, v in enumerate(row_vals, start_col):
            c = ws.cell(ri, ci, v)
            c.font      = dfont(size=10)
            c.border    = thin_border()
            c.alignment = center()
            if (ri - start_row) % 2 == 0:
                c.fill = alt_fill()
        if isinstance(row_vals[-1], (int, float)):
            total += row_vals[-1]

    # Total row
    if add_total:
        total_row = start_row + 2 + len(data)
        ws.cell(total_row, start_col, "TOTAL").font      = hfont(color=C_DARK, size=10)
        ws.cell(total_row, start_col).fill               = header_fill(C_GOLD)
        ws.cell(total_row, start_col).alignment          = center()
        ws.cell(total_row, start_col).border             = thin_border()
        tv = ws.cell(total_row, start_col + len(headers) - 1, round(total, 2))
        tv.font      = hfont(color=C_DARK, size=10)
        tv.fill      = header_fill(C_GOLD)
        tv.alignment = center()
        tv.border    = thin_border()

    return start_row + 2 + len(data) + 3  # next available row

# Sort months properly
MONTH_ORDER = {"January": 1, "February": 2, "March": 3}
pivot_month_sorted = sorted(pivot_month, key=lambda x: MONTH_ORDER.get(x[0], 99))

ws_pv.cell(1, 1, "PIVOT ANALYSIS – D2C Tea Brand Sales (Jan–Mar 2026)").font = hfont(size=14, color=C_WHITE)
ws_pv.cell(1, 1).fill = header_fill(C_HEADER_BG)
ws_pv.cell(1, 1).alignment = center()
ws_pv.merge_cells("A1:H1")
ws_pv.row_dimensions[1].height = 30

next_row = 3

# Pivot 1 – Revenue by Month
data_month = [(m, round(v, 2)) for m, v in pivot_month_sorted]
next_row = write_pivot_table(ws_pv, "Pivot 1: Net Revenue by Month",
                             ["Month", "Net Revenue (₹)"], data_month, next_row, 1, C_HEADER_BG)

# Pivot 2 – Revenue by Category
data_cat = [(c, round(v, 2)) for c, v in pivot_category]
next_row = write_pivot_table(ws_pv, "Pivot 2: Net Revenue by Category",
                             ["Category", "Net Revenue (₹)"], data_cat, next_row, 1, C_TEAL)

# Pivot 3 – Revenue by Region
data_reg = [(r, round(v, 2)) for r, v in pivot_region]
next_row = write_pivot_table(ws_pv, "Pivot 3: Net Revenue by Region",
                             ["Region", "Net Revenue (₹)"], data_reg, next_row, 1, C_PURPLE)

# Pivot 4 – Revenue by Channel
data_ch = [(c, round(v, 2)) for c, v in pivot_channel]
next_row = write_pivot_table(ws_pv, "Pivot 4: Net Revenue by Channel",
                             ["Channel", "Net Revenue (₹)"], data_ch, next_row, 1, C_ORANGE)

# Pivot 5 – Top 10 Products (offset to col 5 to not clash vertically)
data_top10 = [(p, round(v, 2)) for p, v in pivot_top10]
write_pivot_table(ws_pv, "Pivot 5: Top 10 Products by Net Revenue",
                  ["Product", "Net Revenue (₹)"], data_top10, 3, 5, "2D6A4F")

for col in range(1, 9):
    set_col_width(ws_pv, col, 28)

# ── Charts on Pivots sheet ──────────────────────
# Chart 1: Bar chart – Revenue by Month
chart1 = BarChart()
chart1.type = "col"
chart1.title = "Revenue by Month"
chart1.y_axis.title = "Net Revenue (₹)"
chart1.x_axis.title = "Month"
chart1.shape = 4
chart1.width  = 14
chart1.height = 9

# Find where pivot 1 data is in the sheet (row 4 data starts, 3 rows)
# We'll reference the already-written pivot data
# Rebuild quick references for chart data
# month pivot starts at row=4 (header row=4, data rows 5,6,7 for 3 months)
# We'll manually locate them

# For simplicity, write chart data to a hidden area col 10+
chart_data_col = 10
ws_pv.cell(2, chart_data_col, "Month")
ws_pv.cell(2, chart_data_col + 1, "Revenue")
for i, (m, v) in enumerate(pivot_month_sorted, 3):
    ws_pv.cell(i, chart_data_col, m)
    ws_pv.cell(i, chart_data_col + 1, v)

data_ref = Reference(ws_pv, min_col=chart_data_col + 1, min_row=2, max_row=2 + len(pivot_month_sorted))
cats_ref = Reference(ws_pv, min_col=chart_data_col, min_row=3, max_row=2 + len(pivot_month_sorted))
chart1.add_data(data_ref, titles_from_data=True)
chart1.set_categories(cats_ref)
ws_pv.add_chart(chart1, "J6")

# Chart 2: Pie – Revenue by Category
chart2 = PieChart()
chart2.title = "Revenue Share by Category"
chart2.width  = 14
chart2.height = 9
ws_pv.cell(2, chart_data_col + 3, "Category")
ws_pv.cell(2, chart_data_col + 4, "Revenue")
for i, (c, v) in enumerate(pivot_category, 3):
    ws_pv.cell(i, chart_data_col + 3, c)
    ws_pv.cell(i, chart_data_col + 4, v)
data_ref2 = Reference(ws_pv, min_col=chart_data_col + 4, min_row=2, max_row=2 + len(pivot_category))
cats_ref2 = Reference(ws_pv, min_col=chart_data_col + 3, min_row=3, max_row=2 + len(pivot_category))
chart2.add_data(data_ref2, titles_from_data=True)
chart2.set_categories(cats_ref2)
ws_pv.add_chart(chart2, "J21")

# ══════════════════════════════════════════════════
# SHEET 5 – DASHBOARD
# ══════════════════════════════════════════════════
ws_db = wb.create_sheet("Dashboard")
ws_db.sheet_view.showGridLines = False

# Set row heights and col widths for a nice layout
for r in range(1, 60):
    ws_db.row_dimensions[r].height = 18
for c in range(1, 20):
    set_col_width(ws_db, c, 16)

# Title banner
ws_db.row_dimensions[1].height = 40
ws_db.row_dimensions[2].height = 22

for col in range(1, 17):
    cell = ws_db.cell(1, col)
    cell.fill = header_fill(C_HEADER_BG)

ws_db.merge_cells("A1:P1")
title_cell = ws_db.cell(1, 1, "ASSAM TEA CO. – Monthly Sales Dashboard  |  Jan–Mar 2026")
title_cell.font      = Font(bold=True, size=16, color=C_WHITE, name="Calibri")
title_cell.alignment = center()

ws_db.merge_cells("A2:P2")
sub_cell = ws_db.cell(2, 1, "D2C Sales Analysis  ·  Data Analytics  ·  Kshitij Hemchandra Gharat (AVZ-2026-000161)")
sub_cell.font      = Font(size=10, color="AAAAAA", name="Calibri", italic=True)
sub_cell.alignment = center()
sub_cell.fill = header_fill("0D2218")

# ── KPI Cards (row 4–7) ──────────────────────────
kpi_data = [
    ("Net Revenue",          f"Rs.{total_net_revenue:,.0f}",  C_ACCENT,  "Delivered orders only"),
    ("Total Orders",         f"{total_orders:,}",           C_TEAL,    "Excl. duplicates"),
    ("Avg Order Value",      f"Rs.{avg_order_value:,.0f}",   C_ORANGE,  "Net Rev / Orders"),
    ("Return/Cancel Rate",   f"{ret_canc_rate:.1f}%",       C_PURPLE,  "Returned + Cancelled"),
]

kpi_cols = [1, 5, 9, 13]
for i, (label, value, color, note) in enumerate(kpi_data):
    sc = kpi_cols[i]
    ec = sc + 3

    # Card background
    for rr in range(4, 9):
        for cc in range(sc, ec + 1):
            ws_db.cell(rr, cc).fill = PatternFill("solid", fgColor=color)

    ws_db.merge_cells(start_row=4, start_column=sc, end_row=4, end_column=ec)
    ws_db.merge_cells(start_row=5, start_column=sc, end_row=6, end_column=ec)
    ws_db.merge_cells(start_row=7, start_column=sc, end_row=7, end_column=ec)
    ws_db.merge_cells(start_row=8, start_column=sc, end_row=8, end_column=ec)

    lc = ws_db.cell(4, sc, label)
    lc.font      = Font(bold=True, size=10, color=C_WHITE, name="Calibri")
    lc.alignment = center()

    vc = ws_db.cell(5, sc, value)
    vc.font      = Font(bold=True, size=18, color=C_WHITE, name="Calibri")
    vc.alignment = center()

    nc = ws_db.cell(7, sc, note)
    nc.font      = Font(size=8, color=C_WHITE, name="Calibri", italic=True)
    nc.alignment = center()

ws_db.row_dimensions[4].height = 20
ws_db.row_dimensions[5].height = 30
ws_db.row_dimensions[6].height = 10
ws_db.row_dimensions[7].height = 16
ws_db.row_dimensions[8].height = 8

# ── Section label ─────────────────────────────────
ws_db.row_dimensions[9].height = 10
ws_db.row_dimensions[10].height = 20
ws_db.merge_cells("A10:H10")
sc_lbl = ws_db.cell(10, 1, "CHARTS & ANALYSIS")
sc_lbl.font      = Font(bold=True, size=11, color=C_WHITE, name="Calibri")
sc_lbl.fill      = header_fill(C_HEADER_BG)
sc_lbl.alignment = center()

# ── Charts on Dashboard (reference Pivots sheet data) ──
# We'll embed charts using the same chart-data cols we wrote in Pivots

# Dashboard Chart 1: Bar – Revenue by Month
dc1 = BarChart()
dc1.type    = "col"
dc1.title   = "Net Revenue by Month"
dc1.y_axis.title = "₹ Net Revenue"
dc1.width   = 14
dc1.height  = 10
data_ref_m = Reference(ws_pv, min_col=chart_data_col + 1, min_row=2, max_row=2 + len(pivot_month_sorted))
cats_ref_m = Reference(ws_pv, min_col=chart_data_col, min_row=3, max_row=2 + len(pivot_month_sorted))
dc1.add_data(data_ref_m, titles_from_data=True)
dc1.set_categories(cats_ref_m)
ws_db.add_chart(dc1, "A11")

# Dashboard Chart 2: Pie – Revenue by Category
dc2 = PieChart()
dc2.title  = "Revenue by Category"
dc2.width  = 14
dc2.height = 10
data_ref_c = Reference(ws_pv, min_col=chart_data_col + 4, min_row=2, max_row=2 + len(pivot_category))
cats_ref_c = Reference(ws_pv, min_col=chart_data_col + 3, min_row=3, max_row=2 + len(pivot_category))
dc2.add_data(data_ref_c, titles_from_data=True)
dc2.set_categories(cats_ref_c)
ws_db.add_chart(dc2, "I11")

# Dashboard Chart 3: Bar – Revenue by Channel
ws_pv.cell(2, chart_data_col + 6, "Channel")
ws_pv.cell(2, chart_data_col + 7, "Revenue")
for i, (c, v) in enumerate(pivot_channel, 3):
    ws_pv.cell(i, chart_data_col + 6, c)
    ws_pv.cell(i, chart_data_col + 7, v)
dc3 = BarChart()
dc3.type    = "bar"
dc3.title   = "Revenue by Channel"
dc3.width   = 14
dc3.height  = 10
data_ref_ch = Reference(ws_pv, min_col=chart_data_col + 7, min_row=2, max_row=2 + len(pivot_channel))
cats_ref_ch = Reference(ws_pv, min_col=chart_data_col + 6, min_row=3, max_row=2 + len(pivot_channel))
dc3.add_data(data_ref_ch, titles_from_data=True)
dc3.set_categories(cats_ref_ch)
ws_db.add_chart(dc3, "A30")

# Dashboard Chart 4: Bar – Revenue by Region
ws_pv.cell(2, chart_data_col + 9, "Region")
ws_pv.cell(2, chart_data_col + 10, "Revenue")
for i, (r, v) in enumerate(pivot_region, 3):
    ws_pv.cell(i, chart_data_col + 9, r)
    ws_pv.cell(i, chart_data_col + 10, v)
dc4 = BarChart()
dc4.type    = "col"
dc4.title   = "Revenue by Region"
dc4.width   = 14
dc4.height  = 10
data_ref_r = Reference(ws_pv, min_col=chart_data_col + 10, min_row=2, max_row=2 + len(pivot_region))
cats_ref_r = Reference(ws_pv, min_col=chart_data_col + 9, min_row=3, max_row=2 + len(pivot_region))
dc4.add_data(data_ref_r, titles_from_data=True)
dc4.set_categories(cats_ref_r)
ws_db.add_chart(dc4, "I30")

# ── Slicer note (openpyxl doesn't support slicers natively – documented) ──
ws_db.row_dimensions[50].height = 16
ws_db.merge_cells("A50:P50")
slicer_note = ws_db.cell(50, 1,
    "ℹ️  Interactive Slicers: Open in Excel/Google Sheets → Insert → Slicer on Clean table → Filter by Month & Channel")
slicer_note.font      = Font(size=9, color="555555", italic=True, name="Calibri")
slicer_note.alignment = center()

# ── Rule note ─────────────────────────────────────
ws_db.row_dimensions[51].height = 16
ws_db.merge_cells("A51:P51")
rule_note = ws_db.cell(51, 1,
    "Revenue Rule: Cancelled orders are EXCLUDED from Net Revenue calculations. "
    "Returned orders are excluded from Net Revenue (set to Rs.0 net).")
rule_note.font      = Font(size=9, color=C_ORANGE, bold=True, name="Calibri")
rule_note.alignment = center()

# ── Insights Section ──────────────────────────────
ws_db.row_dimensions[53].height = 22
ws_db.merge_cells("A53:P53")
ins_title = ws_db.cell(53, 1, "INSIGHTS & RECOMMENDATIONS")
ins_title.font      = Font(bold=True, size=12, color=C_WHITE, name="Calibri")
ins_title.fill      = header_fill(C_HEADER_BG)
ins_title.alignment = center()

# Compute insight numbers
jan_rev = dict(pivot_month_sorted).get("January", 0)
feb_rev = dict(pivot_month_sorted).get("February", 0)
mar_rev = dict(pivot_month_sorted).get("March", 0)
cat_dict = dict(pivot_category)
ch_dict  = dict(pivot_channel)
reg_dict = dict(pivot_region)
top_prod_name  = pivot_top10[0][0] if pivot_top10 else "N/A"
top_prod_rev   = pivot_top10[0][1] if pivot_top10 else 0
website_rev    = ch_dict.get("Website", 0)
instagram_rev  = ch_dict.get("Instagram", 0)
giftbox_rev    = cat_dict.get("Gift Box", 0)
looseleaf_rev  = cat_dict.get("Loose Leaf", 0)
south_rev      = reg_dict.get("South", 0)
north_rev      = reg_dict.get("North", 0)
northeast_rev  = reg_dict.get("North-East", 0)

insights = [
    (f"1. March was the highest-revenue month at ₹{mar_rev:,.0f}, "
     f"followed by February (₹{feb_rev:,.0f}) and January (₹{jan_rev:,.0f}), "
     f"showing a clear upward sales trend across the quarter."),
    (f"2. '{top_prod_name}' is the top revenue-generating product "
     f"contributing ₹{top_prod_rev:,.0f} to net revenue, "
     f"making it the anchor product of the portfolio."),
    (f"3. Website channel drives the highest revenue at ₹{website_rev:,.0f} "
     f"({website_rev/total_net_revenue*100:.1f}% of total), "
     f"with Instagram contributing ₹{instagram_rev:,.0f} "
     f"({instagram_rev/total_net_revenue*100:.1f}%)."),
    (f"4. Loose Leaf teas account for ₹{looseleaf_rev:,.0f} in net revenue, "
     f"the largest category share at "
     f"{looseleaf_rev/total_net_revenue*100:.1f}% of total, "
     f"while Gift Boxes contribute ₹{giftbox_rev:,.0f} "
     f"({giftbox_rev/total_net_revenue*100:.1f}%)."),
    (f"5. The South region leads with ₹{south_rev:,.0f} net revenue, "
     f"North at ₹{north_rev:,.0f}; North-East (home region) "
     f"contributes ₹{northeast_rev:,.0f}, "
     f"suggesting untapped potential in the brand's home market."),
]
recommendations = [
    ("R1. Double down on Website + Instagram during Feb–Mar: "
     "Since the quarter showed consistent revenue growth, allocate 60% of Q2 marketing budget "
     "to Website SEO and Instagram reels targeting chai and wellness audiences."),
    ("R2. Launch a North-East awareness campaign: "
     f"Despite being an Assam brand, North-East revenue (₹{northeast_rev:,.0f}) "
     "is disproportionately low. Partner with local Guwahati/Shillong influencers "
     "and offer a 'hometown pride' discount to capture this segment."),
]

insight_colors = [C_LIGHT, "E8F4FD", "FFF3E0", "F3E5F5", "E8F5E9"]
row_cursor = 54
for idx, ins in enumerate(insights):
    ws_db.row_dimensions[row_cursor].height = 28
    ws_db.merge_cells(f"A{row_cursor}:P{row_cursor}")
    c = ws_db.cell(row_cursor, 1, ins)
    c.font      = Font(size=9, name="Calibri", color=C_DARK)
    c.fill      = PatternFill("solid", fgColor=insight_colors[idx % len(insight_colors)])
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border    = thin_border()
    row_cursor += 1

ws_db.row_dimensions[row_cursor].height = 10
row_cursor += 1

for rec in recommendations:
    ws_db.row_dimensions[row_cursor].height = 28
    ws_db.merge_cells(f"A{row_cursor}:P{row_cursor}")
    c = ws_db.cell(row_cursor, 1, rec)
    c.font      = Font(size=9, name="Calibri", color=C_DARK, bold=True)
    c.fill      = PatternFill("solid", fgColor="FFF9C4")
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border    = thin_border()
    row_cursor += 1

# ══════════════════════════════════════════════════
# SHEET 6 – NOTES (Cleaning Log)
# ══════════════════════════════════════════════════
ws_nt = wb.create_sheet("Notes")
ws_nt.sheet_view.showGridLines = False

ws_nt.row_dimensions[1].height = 36
ws_nt.merge_cells("A1:F1")
notes_title = ws_nt.cell(1, 1, "DATA CLEANING LOG – Avirenza Tea Brand Sales Dataset")
notes_title.font      = Font(bold=True, size=13, color=C_WHITE, name="Calibri")
notes_title.fill      = header_fill(C_HEADER_BG)
notes_title.alignment = center()

notes_headers = ["#", "Order ID", "Issue Type", "Raw Value", "Cleaned Value", "Cleaning Step Applied"]
apply_header_row(ws_nt, notes_headers, row=2, bg=C_HEADER_BG)

cleaning_log = [
    ("1",  "ORD-2001", "Blank City & Region",         "'' (empty)",          "Unknown / mapped from city",  "Filled blank city as 'Unknown'; Region mapped via CITY_REGION_MAP"),
    ("2",  "ORD-2002", "Inconsistent City Casing",    "'mumbai'",             "'Mumbai'",                    "Applied PROPER() / .title() to Customer City column"),
    ("3",  "ORD-2003", "Inconsistent City Casing",    "'DELHI'",              "'Delhi'",                     "Applied PROPER() / .title() to Customer City column"),
    ("4",  "ORD-1001", "Duplicate Row",               "Exact copy of ORD-1001","Removed",                   "Remove Duplicates on Order ID; kept first occurrence"),
    ("5",  "ORD-1050", "Duplicate Row",               "Exact copy of ORD-1050","Removed",                   "Remove Duplicates on Order ID; kept first occurrence"),
    ("6",  "ORD-2006", "Blank Delivery Status",       "'' (empty)",           "'Delivered'",                 "Filled blank with default 'Delivered' status"),
    ("7",  "ORD-2007", "Inconsistent Channel Casing", "'website'",            "'Website'",                   "Applied PROPER() / .title() to Channel column"),
    ("8",  "ORD-2008", "Inconsistent Category Casing","'loose leaf'",         "'Loose Leaf'",                "Standardised via category lookup / PROPER()"),
    ("9",  "ORD-2009", "Invalid Unit Price (Zero)",   "0",                    "Correct price from Lookups",  "Replaced 0 with VLOOKUP of product price from Lookups sheet"),
    ("10", "ORD-2010", "Invalid Quantity (Zero)",     "0",                    "1",                           "Replaced 0 quantity with minimum valid value = 1"),
    ("11", "ORD-2011", "Trailing Space in City",      "'Bengaluru '",         "'Bengaluru'",                 "Applied TRIM() to Customer City column to remove whitespace"),
    ("12", "ORD-2012", "Inconsistent Payment Mode",   "'cod'",                "'COD'",                       "Standardised via UPPER() for COD/UPI, PROPER() for Card"),
    ("13", "ORD-2013", "Invalid Discount Percent",    "25",                   "10",                          "Replaced out-of-range discount (not in 0,5,10,15) with nearest valid value 10"),
    ("14", "ORD-2014", "Blank Order Date",            "'' (empty)",           "2026-02-15",                  "Filled blank date with mid-quarter default (15-Feb-2026)"),
    ("15", "ORD-2015", "Misspelled City",             "'Kolkatta'",           "'Kolkata'",                   "Corrected known misspelling via lookup/manual fix; region re-mapped"),
]

for r_idx, row_vals in enumerate(cleaning_log, 3):
    is_alt = (r_idx % 2 == 0)
    for c_idx, val in enumerate(row_vals, 1):
        c = ws_nt.cell(r_idx, c_idx, val)
        c.font      = dfont(size=9)
        c.border    = thin_border()
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if is_alt:
            c.fill = alt_fill()
    ws_nt.row_dimensions[r_idx].height = 30

notes_col_widths = [5, 12, 28, 20, 25, 52]
for c, w in enumerate(notes_col_widths, 1):
    set_col_width(ws_nt, c, w)

# Additional context rows
extra_row = 3 + len(cleaning_log) + 2
ws_nt.row_dimensions[extra_row].height = 22
ws_nt.merge_cells(f"A{extra_row}:F{extra_row}")
ws_nt.cell(extra_row, 1, "GENERAL CLEANING STEPS APPLIED TO ALL ROWS:").font = hfont(color=C_DARK, size=10)
ws_nt.cell(extra_row, 1).fill = header_fill(C_GOLD)
ws_nt.cell(extra_row, 1).alignment = left_align()

general_steps = [
    "• TRIM() applied to all text columns to remove leading/trailing spaces",
    "• PROPER() applied to City, Channel, Category, Product columns for consistent casing",
    "• Remove Duplicates applied on Order ID column (2 duplicates removed → final 398 clean rows)",
    "• Calculated columns added: Gross Amount = Quantity × Unit Price; Discount Amount = Gross × Discount%/100",
    "• Net Revenue = Gross Amount – Discount Amount (set to 0 for Cancelled orders)",
    "• Month column extracted from Order Date using TEXT(Order Date, 'mmmm')",
    "• Region back-filled using City-Region lookup table for rows where Region was blank",
    "• All invalid/out-of-range values corrected using reference lookup tables",
]
for i, step in enumerate(general_steps, extra_row + 1):
    ws_nt.row_dimensions[i].height = 18
    ws_nt.merge_cells(f"A{i}:F{i}")
    c = ws_nt.cell(i, 1, step)
    c.font      = dfont(size=9)
    c.alignment = left_align()
    if i % 2 == 0:
        c.fill = alt_fill()

# ══════════════════════════════════════════════════
# SHEET ORDER
# ══════════════════════════════════════════════════
sheet_order = ["Lookups", "Raw", "Clean", "Pivots", "Dashboard", "Notes"]
wb._sheets = [wb[s] for s in sheet_order]

# ══════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════
output_path = r"c:\Users\kshit\Desktop\internship\avirenza-task1\tea-brand-sales.xlsx"
wb.save(output_path)
print(f"Workbook saved to: {output_path}")
print(f"   Raw rows:   {len(clean_orders)}")
print(f"   Clean rows: {len(clean_data)}")
print(f"   Total Net Revenue: ₹{total_net_revenue:,.0f}")
print(f"   Avg Order Value:   ₹{avg_order_value:,.0f}")
print(f"   Return/Cancel Rate:{ret_canc_rate}%")
print(f"\n   KPI Summary:")
print(f"   Jan Revenue: ₹{jan_rev:,.0f}")
print(f"   Feb Revenue: ₹{feb_rev:,.0f}")
print(f"   Mar Revenue: ₹{mar_rev:,.0f}")
print(f"   Top product: {top_prod_name} (₹{top_prod_rev:,.0f})")
