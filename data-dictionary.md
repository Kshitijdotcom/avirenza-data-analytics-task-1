# Data Dictionary – Assam Tea Co. Sales Dataset

**Project:** Monthly Sales Dashboard for a Small D2C Tea Brand  
**Candidate:** Kshitij Hemchandra Gharat (AVZ-2026-000161)  
**Dataset Period:** January 2026 – March 2026  
**Total Raw Orders:** 400 | **Clean Orders (post-deduplication):** 398

---

## Sheet: Raw

The Raw sheet contains 400 sample orders built to represent a realistic D2C tea brand export.  
15 data quality issues are deliberately introduced (see Notes sheet for the full log).

| Column | Data Type | Description | How Built | Allowed Values |
|---|---|---|---|---|
| Order ID | Text | Unique identifier for each order | Sequential from ORD-1001 to ORD-1385; ORD-2001 to ORD-2015 used for rows with planted issues | Unique string |
| Order Date | Date | Date the order was placed | Random date spread across Jan–Mar 2026 | 2026-01-01 to 2026-03-31 |
| Customer City | Text | City the order was shipped to | Selected from a 14-city lookup table covering major Indian cities | Mumbai, Delhi, Bengaluru, Kolkata, Chennai, Hyderabad, Pune, Jaipur, Guwahati, Lucknow, Ahmedabad, Kochi, Bhubaneswar, Shillong |
| Region | Text | Geographical region of the city | Looked up from city-region mapping table | North, South, East, West, North-East |
| Channel | Text | Sales channel through which order was placed | Weighted distribution: Website 50%, Instagram 35%, WhatsApp 15% | Website, Instagram, WhatsApp |
| Product | Text | Name of the product ordered | Selected from a 13-product lookup table | See Lookups sheet for full list |
| Category | Text | Product category | Derived from the product lookup | Loose Leaf, Tea Bags, Gift Box, Accessories |
| Quantity | Integer | Number of units ordered | Random value between 1 and 5 | 1, 2, 3, 4, 5 |
| Unit Price | Integer | Price per unit in Indian Rupees (₹) | Taken from product lookup table | ₹150 to ₹1,200 depending on product |
| Discount Percent | Integer | Discount applied on the order | Weighted distribution: 0%→45%, 5%→25%, 10%→20%, 15%→10% | 0, 5, 10, 15 |
| Payment Mode | Text | Mode of payment used | Weighted distribution: UPI 55%, Card 28%, COD 17% | UPI, Card, COD |
| Delivery Status | Text | Final delivery status of the order | Weighted distribution: Delivered 78%, Returned 12%, Cancelled 10% | Delivered, Returned, Cancelled |

---

## Sheet: Clean

The Clean sheet is a copy of Raw with all 15 data quality issues resolved.  
Four additional computed columns are appended.

| Column | Data Type | Description | Formula / Logic |
|---|---|---|---|
| Order ID | Text | Same as Raw – unique order identifier | Carried from Raw after deduplication |
| Order Date | Date | Cleaned order date | Blank dates filled with `2026-02-15` (mid-quarter default) |
| Customer City | Text | Cleaned city name | `TRIM(PROPER(Customer City))` · Known misspellings corrected manually |
| Region | Text | Region mapped from cleaned city | `VLOOKUP(Customer City, Lookups!E:F, 2, FALSE)` |
| Channel | Text | Cleaned channel | `PROPER(TRIM(Channel))` |
| Product | Text | Same as Raw | Carried from Raw |
| Category | Text | Cleaned category | Standardised via category lookup / `PROPER()` |
| Quantity | Integer | Cleaned quantity | Zero values replaced with `1` |
| Unit Price | Integer | Cleaned unit price | Zero values replaced via `VLOOKUP(Product, Lookups!A:C, 3, FALSE)` |
| Discount Percent | Integer | Cleaned discount | Out-of-range values replaced with `10` (nearest valid) |
| Payment Mode | Text | Cleaned payment mode | `UPPER()` for UPI/COD; `PROPER()` for Card |
| Delivery Status | Text | Cleaned status | Blanks filled with `"Delivered"` |
| **Gross Amount** | Currency (₹) | Revenue before discount | `= Quantity × Unit Price` |
| **Discount Amount** | Currency (₹) | Amount discounted | `= Gross Amount × Discount Percent / 100` |
| **Net Revenue** | Currency (₹) | Actual revenue earned | `= Gross Amount − Discount Amount` (set to `0` for Cancelled orders) |
| **Month** | Text | Month of the order | `= TEXT(Order Date, "mmmm")` |

> **Revenue Rule:** Cancelled orders contribute ₹0 to Net Revenue. This rule is stated on the Dashboard.

---

## Sheet: Lookups

Reference tables used during data generation and cleaning.

| Table | Columns | Purpose |
|---|---|---|
| Product Lookup | Product Name, Category, Unit Price | Maps each product to its category and standard price |
| City & Region Lookup | City, Region | Maps each city to its geographic region |
| Channel Reference | Channel | Lists valid channel values |
| Discount Reference | Discount % | Lists allowed discount values (0, 5, 10, 15) |
| Payment Mode Reference | Payment Mode | Lists valid payment modes |
| Delivery Status Reference | Delivery Status | Lists valid delivery statuses |

### Products (13 products across 4 categories)

| Product | Category | Unit Price (₹) |
|---|---|---|
| Assam Classic Loose Leaf | Loose Leaf | 350 |
| Darjeeling First Flush | Loose Leaf | 520 |
| Green Tulsi Blend | Loose Leaf | 290 |
| Masala Chai Loose Leaf | Loose Leaf | 310 |
| Earl Grey Loose Leaf | Loose Leaf | 380 |
| Ginger Lemon Tea Bags | Tea Bags | 180 |
| Classic Assam Tea Bags | Tea Bags | 150 |
| Green Mint Tea Bags | Tea Bags | 160 |
| Festive Gift Box (Mixed) | Gift Box | 850 |
| Premium Tea Gift Hamper | Gift Box | 1,200 |
| Wellness Herbal Gift Box | Gift Box | 750 |
| Bamboo Tea Infuser | Accessories | 220 |
| Ceramic Tea Mug | Accessories | 320 |

---

## Sheet: Pivots

Five pivot tables summarising the clean data.

| Pivot | Rows | Values | Purpose |
|---|---|---|---|
| Pivot 1 | Month | Net Revenue (₹) | Trend analysis across Jan, Feb, Mar |
| Pivot 2 | Category | Net Revenue (₹) | Which category earns most |
| Pivot 3 | Region | Net Revenue (₹) | Geographic performance |
| Pivot 4 | Channel | Net Revenue (₹) | Website vs Instagram vs WhatsApp |
| Pivot 5 | Product (Top 10) | Net Revenue (₹) | Best-performing individual SKUs |

---

## Sheet: Notes

A full cleaning log listing all 15 data quality issues planted in the Raw sheet,  
the raw value, the corrected value, and the cleaning technique applied.

---

## Sheet: Dashboard

One-page summary with:
- **4 KPI Cards:** Net Revenue, Total Orders, Average Order Value, Return/Cancellation Rate  
- **4 Charts:** Revenue by Month (bar), Revenue by Category (pie), Revenue by Channel (bar), Revenue by Region (bar)  
- **Insights & Recommendations:** 5 data-backed observations + 2 actionable business recommendations  
- **Filter Note:** Instructions to add Month and Channel slicers in Excel/Google Sheets

---

## Messy Issues Planted (summary)

| # | Issue Type |
|---|---|
| 1 | Blank Customer City and Region |
| 2 | City in lowercase ('mumbai') |
| 3 | City in uppercase ('DELHI') |
| 4 | Duplicate row (ORD-1001) |
| 5 | Duplicate row (ORD-1050) |
| 6 | Blank Delivery Status |
| 7 | Channel in lowercase ('website') |
| 8 | Category in lowercase ('loose leaf') |
| 9 | Unit Price = 0 (invalid) |
| 10 | Quantity = 0 (invalid) |
| 11 | City with trailing space ('Bengaluru ') |
| 12 | Payment Mode in lowercase ('cod') |
| 13 | Discount Percent = 25 (not in allowed set) |
| 14 | Blank Order Date |
| 15 | City misspelled ('Kolkatta' → 'Kolkata') |
