# 🍔 McDonald's India — Menu Nutrition Analysis

An end-to-end data analyst portfolio project that automates the full pipeline from raw CSV data to a professionally formatted Excel report — using **Python**, **SQL**, and **Excel (openpyxl)**.

---

## 📌 Project Overview

This project analyzes the nutritional content of McDonald's India's menu (141 items across 7 categories). The pipeline is fully automated: one Python script cleans the data, loads it into a SQLite database, runs analytical SQL queries, and generates a formatted multi-sheet Excel report.

**Skills demonstrated:** Data cleaning · Feature engineering · SQL aggregations & filtering · Excel automation · Data storytelling

---

## 🗂️ Project Structure

```
mcdonalds-india-nutrition/
│
├── India_Menu.csv          # Raw dataset (source: Kaggle)
├── pipeline.py             # Main automation script
├── mcdonalds.db            # SQLite database 
└── McDonalds_India_Nutrition_Analysis.xlsx  # Final Excel report 
```

---

## ⚙️ Tech Stack

| Layer   | Tool / Library         | Purpose                              |
|---------|------------------------|--------------------------------------|
| Python  | `pandas`               | Data loading, cleaning, engineering  |
| SQL     | `sqlite3`              | Database creation & analytical queries |
| Excel   | `openpyxl`             | Formatted multi-sheet report output  |

**Python version:** 3.8+  
**Dependencies:** `pandas`, `openpyxl`

---

## 🚀 How to Run

**1. Clone the repository**
```bash
git clone https://github.com/your-username/mcdonalds-india-nutrition.git
cd mcdonalds-india-nutrition
```

**2. Install dependencies**
```bash
pip install pandas openpyxl
```

**3. Run the pipeline**
```bash
python pipeline.py
```

The script will print progress logs and generate `mcdonalds.db` and the Excel report automatically.

---

## 🔄 Pipeline Stages

### Stage 1 — Python: Data Cleaning & Feature Engineering
- Loads `India_Menu.csv` using pandas (141 rows, 13 columns)
- Strips whitespace from column names and string values
- Parses serving size strings (e.g., `"168 g"`) to numeric values
- Engineers 3 new features:
  - `High_Calorie` — flag for items > 500 kcal
  - `High_Sodium` — flag for items > 800 mg sodium
  - `Protein_Per_100kcal` — protein efficiency ratio

### Stage 2 — SQL: Load & Query via SQLite
Cleaned data is loaded into a local SQLite database. Five analytical queries are executed:

| # | Query | Purpose |
|---|-------|---------|
| 1 | Category Nutrition Summary | AVG/MAX calories, protein, fat, sodium per category |
| 2 | Top 10 Highest Calorie Items | Ranked by `Energy (kCal)` DESC |
| 3 | Top 10 Highest Protein Items | Ranked by `Protein (g)` with efficiency ratio |
| 4 | High Sodium Alert | Items where `Sodium > 800mg` |
| 5 | Healthiest Picks | Items where calories < 400, protein > 8g, sodium < 700mg |

**Sample SQL (Query 5 — Healthiest Picks):**
```sql
SELECT "Menu Category", "Menu Items",
       "Energy (kCal)", "Protein (g)", "Sodium (mg)",
       ROUND(Protein_Per_100kcal, 2) AS Protein_per_100kcal
FROM menu
WHERE "Energy (kCal)" < 400
  AND "Protein (g)" > 8
  AND "Sodium (mg)" < 700
ORDER BY Protein_Per_100kcal DESC
LIMIT 10;
```

### Stage 3 — Excel: Automated Formatted Report
A 7-sheet Excel workbook is generated with McDonald's brand colors, zebra-striped tables, KPI cards, and embedded bar charts.

| Sheet | Content |
|-------|---------|
| 📊 Dashboard | 8 KPI cards + full category nutrition summary table |
| 🔥 High Calorie | Top 10 highest calorie items + bar chart |
| 💪 High Protein | Top 10 highest protein items + bar chart |
| ⚠️ High Sodium | 21 flagged items with color-coded severity levels |
| ✅ Healthiest Picks | Smart-filtered best options from the menu |
| 📋 Raw Data | Full 141-row cleaned dataset |
| 🗄️ SQL Queries | All 5 SQL queries for reference |

---

## 📊 Key Findings

- **141 items** across 7 menu categories
- **14 high-calorie items** exceed 500 kcal per serving
- **21 high-sodium items** exceed 800 mg sodium — 40%+ of the daily recommended limit in one meal
- The **Gourmet Menu** has the highest average calorie count per item
- The **McCafe & Beverages** categories are the safest for low-sodium choices
- Only **10 items** meet all three "healthy" criteria: < 400 kcal, > 8g protein, < 700mg sodium

---

## 📁 Dataset

**Source:** [McDonald's India Menu Nutrition Facts — Kaggle](https://www.kaggle.com/datasets/deepcontractor/mcdonalds-india-menu-nutrition-facts/data)

**Columns:** Menu Category, Menu Items, Per Serve Size, Energy (kCal), Protein (g), Total fat (g), Sat Fat (g), Trans fat (g), Cholesterols (mg), Total carbohydrate (g), Total Sugars (g), Added Sugars (g), Sodium (mg)

---

## 💡 What This Project Demonstrates

- Writing clean, modular Python for data pipelines
- Using SQL (`GROUP BY`, `ORDER BY`, `WHERE`, `LIMIT`, `ROUND`, `AVG`, `MAX`) for analysis
- Automating Excel report generation with formatting, charts, and multiple sheets
- End-to-end thinking: from raw data → insights → presentation
- Portfolio-ready documentation and code structure

---

## 📬 Contact

**Robina Shakwa**
https://www.linkedin.com/in/robina-shakwa-b4519b3b5/
