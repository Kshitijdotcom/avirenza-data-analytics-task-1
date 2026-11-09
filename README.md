# Monthly Sales Dashboard – Assam Tea Co. (D2C Tea Brand)

**Avirenza Internship Task 1 | Data Analytics**  
**Candidate:** Kshitij Hemchandra Gharat (AVZ-2026-000161)  
**Due:** 4 December 2026

> **Google Sheets Link:** [View Dataset on Google Sheets](https://docs.google.com/spreadsheets/d/1Jifp278QPRqny1jtW8uju4d7qUJWQl8XHR5p7xlIGwE/edit?usp=sharing)

---

## What I Built

For this task I created a sample dataset of 400 orders (January–March 2026) for a fictional D2C tea brand from Assam, cleaned it, analysed it with pivot tables, and built a one-page sales dashboard. I also built an interactive HTML version of the dashboard that opens directly in any browser.

### Workbook sheets:
| Sheet | Contents |
|---|---|
| **Raw** | 400 orders with 15 deliberately introduced data quality issues |
| **Lookups** | Reference tables for products, cities, regions, channels and valid values |
| **Clean** | 398 cleaned rows + 4 calculated columns (Gross Amount, Discount Amount, Net Revenue, Month) |
| **Pivots** | 5 summary tables covering revenue by month, category, region, channel and top 10 products |
| **Dashboard** | One-page view with 4 KPI cards, 4 charts, 2 slicers, insights and recommendations |
| **Notes** | Cleaning log listing all 15 issues — original value, corrected value and technique used |

---

## Files in This Repository

```
avirenza-task1/
├── tea-brand-sales.xlsx      ← Main workbook (all 6 sheets)
├── dashboard.html            ← Interactive visual dashboard (open in any browser)
├── dashboard.pdf             ← Dashboard sheet exported as PDF
├── data-dictionary.md        ← Description of every column and how it was built
├── generate_workbook.py      ← Python script used to build the dataset and workbook
├── README.md                 ← This file
└── screenshots/
    ├── dashboard.png         ← Screenshot of the Dashboard sheet
    └── notes.png             ← Screenshot of the Notes sheet
```

---

## Tools Used

| Tool | Purpose |
|---|---|
| Python 3.14 | Building the dataset and workbook programmatically |
| openpyxl 3.1.5 | Writing and formatting the Excel file |
| pandas 2.3.3 | Data checks and aggregations |
| Chart.js 4.4.0 | Charts in the HTML dashboard |
| HTML + CSS Grid | Dashboard layout |
| Microsoft Excel | Adding slicers and exporting the PDF |

---

## How to Open

### Excel Workbook
1. Download `tea-brand-sales.xlsx`
2. Open in **Microsoft Excel** or **Google Sheets** (File → Import)
3. Navigate the sheets using the tabs at the bottom: Lookups → Raw → Clean → Pivots → Dashboard → Notes

### HTML Dashboard
1. Download `dashboard.html`
2. Double-click — it opens in any browser (Chrome, Edge, Firefox)
3. No internet or installation needed

### Slicers (Month & Channel filters)
- **Excel:** Click any cell in the Clean table → Insert → Slicer → tick Month and Channel → OK
- **Google Sheets:** Select the Clean data range → Data → Add a slicer → set column to Month / Channel

---

## Dashboard KPIs (Jan–Mar 2026)

| KPI | Value |
|---|---|
| Net Revenue | ₹4,37,689 |
| Total Orders | 398 |
| Average Order Value | ₹1,100 |
| Return/Cancellation Rate | 23.1% |

> **Revenue Rule:** Cancelled orders are excluded from Net Revenue (counted as Rs. 0). This is stated on the Dashboard.

---

## Key Insights

1. **January was the highest-revenue month** at ₹1,55,536, with February close behind at ₹1,54,376. March dipped to ₹1,27,776.
2. **Premium Tea Gift Hamper** is the best-performing product at ₹69,180 net revenue for the quarter.
3. **Website brought in 51.2%** of total revenue, with Instagram contributing a strong 34.8%.
4. **Loose Leaf and Gift Box** categories together made up about 79% of net revenue.
5. **The North-East region** — the brand's home — underperformed compared to South and North, which points to an untapped opportunity.

## Recommendations

- **R1:** Put more marketing spend behind the Website and Instagram channels in Q2, since they already drive the majority of sales.
- **R2:** Run a North-East focused campaign using local influencers in Guwahati and Shillong to grow the brand in its home region.

---

## Data Cleaning Summary (15 Issues Fixed)

| # | Issue | Fix Applied |
|---|---|---|
| 1 | Blank Customer City and Region | Filled with 'Unknown'; Region re-mapped from lookup |
| 2 | City in lowercase ('mumbai') | PROPER() to correct casing |
| 3 | City in uppercase ('DELHI') | PROPER() to correct casing |
| 4 | Duplicate row (ORD-1001) | Duplicate removed |
| 5 | Duplicate row (ORD-1050) | Duplicate removed |
| 6 | Blank Delivery Status | Filled with default 'Delivered' |
| 7 | Channel in lowercase ('website') | PROPER() applied |
| 8 | Category in lowercase ('loose leaf') | Standardised via lookup |
| 9 | Unit Price = 0 | Replaced using VLOOKUP from Lookups sheet |
| 10 | Quantity = 0 | Replaced with minimum valid value of 1 |
| 11 | Trailing space in city name ('Bengaluru ') | TRIM() applied |
| 12 | Payment Mode in lowercase ('cod') | UPPER() applied |
| 13 | Discount % = 25 (not in allowed set) | Replaced with nearest valid value of 10 |
| 14 | Blank Order Date | Filled with mid-quarter date (15-Feb-2026) |
| 15 | City misspelled ('Kolkatta') | Corrected to 'Kolkata'; Region re-mapped |

---

## Dashboard Screenshot

![Dashboard](screenshots/dashboard.png)

## Notes Sheet Screenshot

![Notes Sheet](screenshots/notes.png)

---

## What I Learned

- How to build a realistic sample dataset from scratch using Python and control value distributions
- Why data validation matters — even small issues like trailing spaces or wrong casing can break pivot results
- How to structure a multi-sheet workbook so raw data, clean data, reference tables and analysis stay separate
- How to build a visual dashboard using HTML and Chart.js that works offline in any browser
- How to turn numbers into specific, actionable insights rather than just describing the data
