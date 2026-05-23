# 🛒 Olist Brazilian E-Commerce — Data Analysis & Excel Dashboard

A comprehensive data analysis project built on the [Olist public dataset](https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce), covering **100,000+ orders** placed on Brazil's largest marketplace between **2016 and 2018**.

---

## 📊 Project Overview

This project merges 9 relational CSV tables into a single master fact table, cleans and standardises the data, and delivers a fully formatted, multi-tab **Excel workbook** with pivot analyses, live formulas, and an executive dashboard — all built programmatically in Python.

| Metric | Value |
|---|---|
| Total Revenue | R$ 13,279,837 |
| Total Delivered Orders | 96,478 |
| Average Order Value | R$ 137.75 |
| Average Review Score | 4.09 / 5.0 |
| Median Delivery Time | 12 days |

---

## 📁 Repository Structure

```
olist-ecommerce-analysis/
│
├── build_workbook.py               # Python script to generate the Excel workbook
├── Olist_Ecommerce_Analysis.xlsx   # Final deliverable — open directly in Excel
└── README.md
```

---

## 📂 Excel Workbook Structure

| Tab | Contents |
|---|---|
| **README Cover** | Project description, table of contents, data dictionary |
| **Clean Data** | 5,000-row cleaned fact table as an Excel Table with filters |
| **Analysis & Pivots** | 3 pivot summaries with live SUMIFS / IFERROR formulas |
| **Dashboard** | 5 KPI cards + 4 native Excel charts |

---

## 🔍 Key Insights

- **São Paulo dominates** — SP alone accounts for ~42% of total revenue, with the Southeast (SP + RJ + MG) generating over 60% combined.
- **Revenue grew rapidly** through 2017–2018 after a near-zero 2016 — classic marketplace hockey-stick growth.
- **Customer satisfaction is strong** — 57% of reviews are 5-star, though ~11% are 1-star, often tied to late deliveries.
- **Credit card is king** — ~74% of payments are by credit card; boleto (bank slip) is a distant second at ~19%.
- **Top categories:** Bed & Bath, Health & Beauty, and Computers & Accessories lead by revenue.

---

## 🛠️ How to Reproduce

### Requirements

```bash
pip install pandas openpyxl numpy
```

### Data

Download the dataset from Kaggle:  
👉 [Brazilian E-Commerce Public Dataset by Olist](https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce)

Extract all CSVs into a `data/` folder in the project root.

### Run

```bash
python build_workbook.py
```

The output file `Olist_Ecommerce_Analysis.xlsx` will be written to `/mnt/user-data/outputs/` (adjust the path in the script as needed for your environment).

---

## 📦 Data Sources

| File | Description |
|---|---|
| `olist_orders_dataset.csv` | Order lifecycle and timestamps |
| `olist_order_items_dataset.csv` | Line items with price and freight |
| `olist_order_payments_dataset.csv` | Payment type and value |
| `olist_order_reviews_dataset.csv` | Customer review scores |
| `olist_customers_dataset.csv` | Customer city and state |
| `olist_products_dataset.csv` | Product attributes and category |
| `olist_sellers_dataset.csv` | Seller location data |
| `product_category_name_translation.csv` | Portuguese → English category names |

---

## 📄 License

Dataset: [CC BY-NC-SA 4.0](https://creativecommons.org/licenses/by-nc-sa/4.0/) — Olist  
Code: MIT
