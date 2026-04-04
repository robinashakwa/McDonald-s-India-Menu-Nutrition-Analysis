# 🍔 McDonald's India Menu Nutrition Analysis

### End-to-End Data Analytics Pipeline (Python → SQL → Excel)

---

## 📌 Project Overview

This project analyzes the nutritional composition of McDonald's India menu items and builds a complete **data analytics pipeline** that transforms raw data into actionable insights and a professional Excel dashboard.

The goal is to simulate a **real-world data analyst workflow**, where data is cleaned, analyzed, and presented in a structured report for decision-making.

---

## 🎯 Problem Statement

Can customers make healthier food choices at McDonald's?

This project aims to:

* Identify high-calorie and high-sodium menu items
* Analyze nutritional trends across categories
* Detect healthier food options based on defined criteria
* Deliver insights in a structured dashboard format

---

## ⚙️ Tech Stack

* **Python (Pandas, NumPy)** → Data Cleaning & Feature Engineering
* **SQL (SQLite)** → Data Analysis & Querying
* **Excel (OpenPyXL)** → Automated Dashboard & Reporting

---

## 🔄 Data Pipeline

```text
Raw Dataset (CSV)
        ↓
Python (Cleaning + Feature Engineering)
        ↓
SQLite (Analytical Queries)
        ↓
Excel (Dashboard + Report Generation)
```

---

## 🧹 Data Cleaning & Feature Engineering

* Cleaned column names and handled text inconsistencies
* Extracted numeric values from serving size
* Created new analytical features:

  * `High_Calorie` → Items above 500 kcal
  * `High_Sodium` → Items above 800 mg sodium
  * `Protein_Per_100kcal` → Protein efficiency metric

---

## 🧠 Key Analysis Performed

### 📊 Category-Level Insights

* Average calories, protein, fat, carbs, and sodium
* Identification of calorie-dense categories

### 🔥 High-Calorie Items

* Top 10 items with highest calorie content
* Helps identify unhealthy choices

### 💪 High-Protein Items

* Ranked by protein content
* Evaluated protein efficiency (per 100 kcal)

### ⚠️ High Sodium Alert

* Flagged items exceeding safe sodium thresholds
* Compared against WHO daily recommendations

### ✅ Healthiest Picks

* Filtered items based on:

  * Low calories (<400 kcal)
  * High protein (>8g)
  * Low sodium (<700 mg)

---

## 📈 Output

The pipeline generates a fully formatted Excel report with:

* 📊 Dashboard (KPIs + category summary)
* 🔥 High-calorie analysis with charts
* 💪 High-protein analysis with charts
* ⚠️ High sodium alerts
* ✅ Healthy food recommendations
* 📋 Cleaned raw dataset
* 🗄️ SQL query reference

---

## 💡 Key Insights

* Certain menu categories dominate calorie and sodium intake
* Many items exceed recommended sodium levels in a single serving
* Only a limited subset of items qualify as “healthy” based on defined criteria
* Protein efficiency varies significantly across menu items

---

## 🚀 How to Run

```bash
pip install -r requirements.txt
python pipeline.py
```

The script will generate the Excel report automatically.

---

## 📂 Project Structure

```text
McDonalds-Nutrition-Analysis/
│
├── pipeline.py
├── India_Menu.csv
├── requirements.txt
├── outputs/
│   └── McDonalds_India_Nutrition_Analysis.xlsx
└── README.md
```

## 🧠 Key Learning

This project demonstrates:

* End-to-end data analysis workflow
* Data cleaning and feature engineering
* Writing analytical SQL queries
* Automating report generation
* Translating data into business insights

## Dataset source: https://www.kaggle.com/datasets/deepcontractor/mcdonalds-india-menu-nutrition-facts/data
---
project ongoing........................................................
