"""
McDonald's India Menu Nutrition Analysis
Automated Pipeline: Python → SQL → Excel
"""

import pandas as pd
import sqlite3
import os
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint

# ─────────────────────────────────────────────
# STEP 1: PYTHON — Load & Clean Data
# ─────────────────────────────────────────────

print("=" * 55)
print("  McDonald's India Menu — Nutrition Analysis Pipeline")
print("=" * 55)

CSV_PATH = os.path.join(os.path.dirname(__file__), "India_Menu.csv")
df = pd.read_csv(CSV_PATH)

print(f"\n[PYTHON] Loaded {len(df)} rows, {len(df.columns)} columns")

# Clean column names
df.columns = df.columns.str.strip()

# Strip whitespace from string columns
for col in df.select_dtypes(include="object").columns:
    df[col] = df[col].str.strip()

# Parse serve size to numeric grams/ml
df["Serve_Size_Num"] = (
    df["Per Serve Size"]
    .str.extract(r"([\d.]+)")
    .astype(float)
)

# Flag high-sodium items (>800mg = concerning)
df["High_Sodium"] = df["Sodium (mg)"] > 800

# Flag high-calorie items (>500 kcal)
df["High_Calorie"] = df["Energy (kCal)"] > 500

# Protein efficiency ratio
df["Protein_Per_100kcal"] = (df["Protein (g)"] / df["Energy (kCal)"] * 100).round(2)

print(f"[PYTHON] Cleaning done — {df['High_Sodium'].sum()} high-sodium items, "
      f"{df['High_Calorie'].sum()} high-calorie items")


# ─────────────────────────────────────────────
# STEP 2: SQL — Load into SQLite & Query
# ─────────────────────────────────────────────

DB_PATH = os.path.join(os.path.dirname(__file__), "mcdonalds.db")
conn = sqlite3.connect(DB_PATH)
cursor = conn.cursor()

# Load cleaned data into SQL table
df.to_sql("menu", conn, if_exists="replace", index=False)
print(f"\n[SQL] Loaded data into SQLite database → {DB_PATH}")

# Query 1: Category-level nutrition summary
q1 = """
SELECT
    "Menu Category",
    COUNT(*) AS Item_Count,
    ROUND(AVG("Energy (kCal)"), 1)     AS Avg_Calories,
    ROUND(MAX("Energy (kCal)"), 1)     AS Max_Calories,
    ROUND(AVG("Protein (g)"), 1)       AS Avg_Protein_g,
    ROUND(AVG("Total fat (g)"), 1)     AS Avg_Fat_g,
    ROUND(AVG("Sodium (mg)"), 1)       AS Avg_Sodium_mg,
    ROUND(AVG("Total carbohydrate (g)"), 1) AS Avg_Carbs_g
FROM menu
GROUP BY "Menu Category"
ORDER BY Avg_Calories DESC
"""
df_cat = pd.read_sql(q1, conn)
print(f"[SQL] Query 1 — Category Summary: {len(df_cat)} categories")

# Query 2: Top 10 highest calorie items
q2 = """
SELECT
    "Menu Category",
    "Menu Items",
    "Energy (kCal)",
    "Protein (g)",
    "Total fat (g)",
    "Sodium (mg)"
FROM menu
ORDER BY "Energy (kCal)" DESC
LIMIT 10
"""
df_top10_cal = pd.read_sql(q2, conn)
print(f"[SQL] Query 2 — Top 10 Highest Calorie Items fetched")

# Query 3: Top 10 highest protein items
q3 = """
SELECT
    "Menu Category",
    "Menu Items",
    "Protein (g)",
    "Energy (kCal)",
    ROUND(Protein_Per_100kcal, 2) AS Protein_per_100kcal
FROM menu
ORDER BY "Protein (g)" DESC
LIMIT 10
"""
df_top10_prot = pd.read_sql(q3, conn)
print(f"[SQL] Query 3 — Top 10 Highest Protein Items fetched")

# Query 4: High sodium items
q4 = """
SELECT
    "Menu Category",
    "Menu Items",
    "Sodium (mg)",
    "Energy (kCal)"
FROM menu
WHERE High_Sodium = 1
ORDER BY "Sodium (mg)" DESC
"""
df_sodium = pd.read_sql(q4, conn)
print(f"[SQL] Query 4 — High Sodium Items: {len(df_sodium)} found")

# Query 5: Healthiest items (low cal, good protein, low sodium)
q5 = """
SELECT
    "Menu Category",
    "Menu Items",
    "Energy (kCal)",
    "Protein (g)",
    "Sodium (mg)",
    ROUND(Protein_Per_100kcal, 2) AS Protein_per_100kcal
FROM menu
WHERE "Energy (kCal)" < 400
  AND "Protein (g)" > 8
  AND "Sodium (mg)" < 700
ORDER BY Protein_Per_100kcal DESC
LIMIT 10
"""
df_healthy = pd.read_sql(q5, conn)
print(f"[SQL] Query 5 — Healthiest Items: {len(df_healthy)} found")

conn.close()


# ─────────────────────────────────────────────
# STEP 3: EXCEL — Build Formatted Report
# ─────────────────────────────────────────────

print("\n[EXCEL] Building formatted Excel report...")

wb = Workbook()

# ── Color palette (McDonald's branded) ──
GOLD        = "FFC72C"
RED         = "DA291C"
DARK_RED    = "9B1C14"
WHITE       = "FFFFFF"
LIGHT_GRAY  = "F5F5F5"
MID_GRAY    = "D9D9D9"
DARK_GRAY   = "404040"
GREEN       = "2E7D32"
ORANGE      = "E65100"
BLUE        = "1565C0"

def hdr_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def thin_border():
    s = Side(style="thin", color="BDBDBD")
    return Border(left=s, right=s, top=s, bottom=s)

def apply_table_header(ws, row, headers, bg=RED, fg=WHITE, bold=True):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = Font(bold=bold, color=fg, name="Arial", size=10)
        c.fill = hdr_fill(bg)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin_border()

def apply_data_row(ws, row, values, bg=None, number_fmts=None):
    for col, val in enumerate(values, 1):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(name="Arial", size=9, color=DARK_GRAY)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border()
        if bg:
            c.fill = hdr_fill(bg)
        if number_fmts and col - 1 < len(number_fmts) and number_fmts[col - 1]:
            c.number_format = number_fmts[col - 1]

def zebra_bg(row_idx):
    return LIGHT_GRAY if row_idx % 2 == 0 else WHITE

def set_col_widths(ws, widths):
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w


# ════════════════════════════════════════════
# SHEET 1: Dashboard / Summary
# ════════════════════════════════════════════
ws1 = wb.active
ws1.title = "📊 Dashboard"
ws1.sheet_view.showGridLines = False

# Title banner
ws1.merge_cells("A1:H1")
t = ws1["A1"]
t.value = "🍔  McDonald's India — Menu Nutrition Analysis"
t.font = Font(bold=True, size=18, color=WHITE, name="Arial")
t.fill = hdr_fill(RED)
t.alignment = Alignment(horizontal="center", vertical="center")
ws1.row_dimensions[1].height = 42

ws1.merge_cells("A2:H2")
sub = ws1["A2"]
sub.value = "Automated Python · SQL · Excel Pipeline  |  141 Menu Items  |  7 Categories"
sub.font = Font(size=10, color=WHITE, name="Arial", italic=True)
sub.fill = hdr_fill(DARK_RED)
sub.alignment = Alignment(horizontal="center", vertical="center")
ws1.row_dimensions[2].height = 22

ws1.row_dimensions[3].height = 10  # spacer

# ── KPI Cards (row 4-7) ──
kpis = [
    ("Total Items", len(df), RED),
    ("Menu Categories", df["Menu Category"].nunique(), DARK_RED),
    ("High-Calorie Items\n(>500 kcal)", int(df["High_Calorie"].sum()), ORANGE),
    ("High-Sodium Items\n(>800mg)", int(df["High_Sodium"].sum()), ORANGE),
    ("Avg Calories", f"{df['Energy (kCal)'].mean():.0f} kcal", BLUE),
    ("Avg Protein", f"{df['Protein (g)'].mean():.1f} g", GREEN),
    ("Max Calories", f"{df['Energy (kCal)'].max():.0f} kcal", DARK_RED),
    ("Avg Sodium", f"{df['Sodium (mg)'].mean():.0f} mg", BLUE),
]

cols_kpi = [1, 2, 3, 4, 5, 6, 7, 8]
for i, (label, val, color) in enumerate(kpis):
    col = cols_kpi[i]
    ws1.merge_cells(start_row=4, start_column=col, end_row=5, end_column=col)
    lc = ws1.cell(row=4, column=col, value=label)
    lc.font = Font(bold=True, size=8, color=WHITE, name="Arial")
    lc.fill = hdr_fill(color)
    lc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws1.merge_cells(start_row=6, start_column=col, end_row=7, end_column=col)
    vc = ws1.cell(row=6, column=col, value=val)
    vc.font = Font(bold=True, size=13, color=color, name="Arial")
    vc.fill = hdr_fill(LIGHT_GRAY)
    vc.alignment = Alignment(horizontal="center", vertical="center")
    vc.border = thin_border()

ws1.row_dimensions[4].height = 24
ws1.row_dimensions[5].height = 10
ws1.row_dimensions[6].height = 28
ws1.row_dimensions[7].height = 12

ws1.row_dimensions[8].height = 14  # spacer

# ── Category Summary Table ──
ws1.merge_cells("A9:H9")
ct = ws1["A9"]
ct.value = "Nutrition Summary by Category"
ct.font = Font(bold=True, size=12, color=WHITE, name="Arial")
ct.fill = hdr_fill(GOLD)
ct.alignment = Alignment(horizontal="left", vertical="center")
ws1.row_dimensions[9].height = 26

cat_headers = [
    "Menu Category", "Items", "Avg Calories\n(kcal)",
    "Max Calories\n(kcal)", "Avg Protein\n(g)",
    "Avg Fat\n(g)", "Avg Carbs\n(g)", "Avg Sodium\n(mg)"
]
apply_table_header(ws1, 10, cat_headers, bg=DARK_GRAY)
ws1.row_dimensions[10].height = 34

for i, row in df_cat.iterrows():
    r = 11 + i
    bg = zebra_bg(i)
    vals = [
        row["Menu Category"], row["Item_Count"],
        row["Avg_Calories"], row["Max_Calories"],
        row["Avg_Protein_g"], row["Avg_Fat_g"],
        row["Avg_Carbs_g"], row["Avg_Sodium_mg"]
    ]
    apply_data_row(ws1, r, vals, bg=bg)
    ws1.cell(r, 1).alignment = Alignment(horizontal="left", vertical="center")
    ws1.cell(r, 1).font = Font(bold=True, size=9, name="Arial", color=DARK_GRAY)
    ws1.row_dimensions[r].height = 18

set_col_widths(ws1, [24, 8, 14, 14, 14, 10, 10, 13])


# ════════════════════════════════════════════
# SHEET 2: Top 10 Calorie Items
# ════════════════════════════════════════════
ws2 = wb.create_sheet("🔥 High Calorie")
ws2.sheet_view.showGridLines = False

ws2.merge_cells("A1:F1")
t2 = ws2["A1"]
t2.value = "🔥  Top 10 Highest Calorie Menu Items"
t2.font = Font(bold=True, size=14, color=WHITE, name="Arial")
t2.fill = hdr_fill(ORANGE)
t2.alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[1].height = 36

ws2.merge_cells("A2:F2")
sub2 = ws2["A2"]
sub2.value = "Items exceeding 500 kcal per serving are flagged as high-calorie"
sub2.font = Font(italic=True, size=9, color=DARK_GRAY, name="Arial")
sub2.alignment = Alignment(horizontal="center")
ws2.row_dimensions[2].height = 18

hdrs2 = ["Category", "Menu Item", "Calories\n(kcal)", "Protein\n(g)", "Fat\n(g)", "Sodium\n(mg)"]
apply_table_header(ws2, 4, hdrs2, bg=ORANGE)
ws2.row_dimensions[4].height = 34

for i, row in df_top10_cal.iterrows():
    r = 5 + i
    bg = zebra_bg(i)
    vals = [
        row["Menu Category"], row["Menu Items"],
        row["Energy (kCal)"], row["Protein (g)"],
        row["Total fat (g)"], row["Sodium (mg)"]
    ]
    apply_data_row(ws2, r, vals, bg=bg)
    ws2.cell(r, 2).alignment = Alignment(horizontal="left", vertical="center")
    ws2.cell(r, 3).font = Font(bold=True, size=9, color=ORANGE, name="Arial")
    ws2.row_dimensions[r].height = 18

set_col_widths(ws2, [18, 32, 12, 11, 10, 12])

# Bar chart — Calories
chart2 = BarChart()
chart2.type = "bar"
chart2.title = "Top 10 Items by Calories"
chart2.y_axis.title = "Calories (kcal)"
chart2.style = 10
chart2.width = 18
chart2.height = 12
data_ref = Reference(ws2, min_col=3, min_row=4, max_row=14)
cats_ref = Reference(ws2, min_col=2, min_row=5, max_row=14)
chart2.add_data(data_ref, titles_from_data=True)
chart2.set_categories(cats_ref)
chart2.series[0].graphicalProperties.solidFill = ORANGE
ws2.add_chart(chart2, "A16")


# ════════════════════════════════════════════
# SHEET 3: Top Protein Items
# ════════════════════════════════════════════
ws3 = wb.create_sheet("💪 High Protein")
ws3.sheet_view.showGridLines = False

ws3.merge_cells("A1:F1")
t3 = ws3["A1"]
t3.value = "💪  Top 10 Highest Protein Menu Items"
t3.font = Font(bold=True, size=14, color=WHITE, name="Arial")
t3.fill = hdr_fill(GREEN)
t3.alignment = Alignment(horizontal="center", vertical="center")
ws3.row_dimensions[1].height = 36

ws3.merge_cells("A2:F2")
sub3 = ws3["A2"]
sub3.value = "Protein per 100 kcal = protein efficiency metric (higher = better quality protein source)"
sub3.font = Font(italic=True, size=9, color=DARK_GRAY, name="Arial")
sub3.alignment = Alignment(horizontal="center")
ws3.row_dimensions[2].height = 18

hdrs3 = ["Category", "Menu Item", "Protein\n(g)", "Calories\n(kcal)", "Protein /\n100 kcal"]
apply_table_header(ws3, 4, hdrs3, bg=GREEN)
ws3.row_dimensions[4].height = 34

for i, row in df_top10_prot.iterrows():
    r = 5 + i
    bg = zebra_bg(i)
    vals = [
        row["Menu Category"], row["Menu Items"],
        row["Protein (g)"], row["Energy (kCal)"],
        row["Protein_per_100kcal"]
    ]
    apply_data_row(ws3, r, vals, bg=bg)
    ws3.cell(r, 2).alignment = Alignment(horizontal="left", vertical="center")
    ws3.cell(r, 3).font = Font(bold=True, size=9, color=GREEN, name="Arial")
    ws3.row_dimensions[r].height = 18

set_col_widths(ws3, [18, 34, 11, 12, 14])

# Bar chart — Protein
chart3 = BarChart()
chart3.type = "bar"
chart3.title = "Top 10 Items by Protein Content"
chart3.y_axis.title = "Protein (g)"
chart3.style = 10
chart3.width = 18
chart3.height = 12
data_ref3 = Reference(ws3, min_col=3, min_row=4, max_row=14)
cats_ref3 = Reference(ws3, min_col=2, min_row=5, max_row=14)
chart3.add_data(data_ref3, titles_from_data=True)
chart3.set_categories(cats_ref3)
chart3.series[0].graphicalProperties.solidFill = GREEN
ws3.add_chart(chart3, "A16")


# ════════════════════════════════════════════
# SHEET 4: High Sodium Alert
# ════════════════════════════════════════════
ws4 = wb.create_sheet("⚠️ High Sodium")
ws4.sheet_view.showGridLines = False

ws4.merge_cells("A1:D1")
t4 = ws4["A1"]
t4.value = f"⚠️  High Sodium Items  (Sodium > 800 mg)  —  {len(df_sodium)} items flagged"
t4.font = Font(bold=True, size=13, color=WHITE, name="Arial")
t4.fill = hdr_fill(DARK_RED)
t4.alignment = Alignment(horizontal="center", vertical="center")
ws4.row_dimensions[1].height = 34

ws4.merge_cells("A2:D2")
sub4 = ws4["A2"]
sub4.value = "WHO recommends <2000mg sodium/day. Items exceeding 800mg in a single serving are flagged."
sub4.font = Font(italic=True, size=9, color=DARK_GRAY, name="Arial")
sub4.alignment = Alignment(horizontal="center")
ws4.row_dimensions[2].height = 18

hdrs4 = ["Category", "Menu Item", "Sodium\n(mg)", "Calories\n(kcal)"]
apply_table_header(ws4, 4, hdrs4, bg=DARK_RED)
ws4.row_dimensions[4].height = 34

for i, row in df_sodium.iterrows():
    r = 5 + i
    bg = zebra_bg(i)
    vals = [
        row["Menu Category"], row["Menu Items"],
        row["Sodium (mg)"], row["Energy (kCal)"]
    ]
    apply_data_row(ws4, r, vals, bg=bg)
    ws4.cell(r, 2).alignment = Alignment(horizontal="left", vertical="center")
    # Color-code sodium level
    sodium_val = row["Sodium (mg)"]
    cell = ws4.cell(r, 3)
    if sodium_val > 1200:
        cell.font = Font(bold=True, size=9, color=DARK_RED, name="Arial")
    elif sodium_val > 1000:
        cell.font = Font(bold=True, size=9, color=ORANGE, name="Arial")
    ws4.row_dimensions[r].height = 18

set_col_widths(ws4, [18, 34, 13, 13])


# ════════════════════════════════════════════
# SHEET 5: Healthiest Picks
# ════════════════════════════════════════════
ws5 = wb.create_sheet("✅ Healthiest Picks")
ws5.sheet_view.showGridLines = False

ws5.merge_cells("A1:F1")
t5 = ws5["A1"]
t5.value = "✅  Healthiest Menu Picks  (<400 kcal, >8g Protein, <700mg Sodium)"
t5.font = Font(bold=True, size=13, color=WHITE, name="Arial")
t5.fill = hdr_fill(GREEN)
t5.alignment = Alignment(horizontal="center", vertical="center")
ws5.row_dimensions[1].height = 34

ws5.merge_cells("A2:F2")
sub5 = ws5["A2"]
sub5.value = "Items that balance low calories, adequate protein and controlled sodium — ideal choices"
sub5.font = Font(italic=True, size=9, color=DARK_GRAY, name="Arial")
sub5.alignment = Alignment(horizontal="center")
ws5.row_dimensions[2].height = 18

hdrs5 = ["Category", "Menu Item", "Calories\n(kcal)", "Protein\n(g)", "Sodium\n(mg)", "Protein /\n100 kcal"]
apply_table_header(ws5, 4, hdrs5, bg=GREEN)
ws5.row_dimensions[4].height = 34

for i, row in df_healthy.iterrows():
    r = 5 + i
    bg = zebra_bg(i)
    vals = [
        row["Menu Category"], row["Menu Items"],
        row["Energy (kCal)"], row["Protein (g)"],
        row["Sodium (mg)"], row["Protein_per_100kcal"]
    ]
    apply_data_row(ws5, r, vals, bg=bg)
    ws5.cell(r, 2).alignment = Alignment(horizontal="left", vertical="center")
    ws5.cell(r, 6).font = Font(bold=True, size=9, color=GREEN, name="Arial")
    ws5.row_dimensions[r].height = 18

set_col_widths(ws5, [18, 32, 12, 11, 12, 14])


# ════════════════════════════════════════════
# SHEET 6: Raw Data (cleaned)
# ════════════════════════════════════════════
ws6 = wb.create_sheet("📋 Raw Data")
ws6.sheet_view.showGridLines = False

ws6.merge_cells("A1:M1")
t6 = ws6["A1"]
t6.value = "📋  Full Cleaned Dataset — McDonald's India Menu (141 Items)"
t6.font = Font(bold=True, size=13, color=WHITE, name="Arial")
t6.fill = hdr_fill(DARK_GRAY)
t6.alignment = Alignment(horizontal="center", vertical="center")
ws6.row_dimensions[1].height = 30

export_cols = [
    "Menu Category", "Menu Items", "Per Serve Size",
    "Energy (kCal)", "Protein (g)", "Total fat (g)",
    "Sat Fat (g)", "Trans fat (g)", "Cholesterols (mg)",
    "Total carbohydrate (g)", "Total Sugars (g)",
    "Added Sugars (g)", "Sodium (mg)"
]
apply_table_header(ws6, 2, export_cols, bg=DARK_GRAY)
ws6.row_dimensions[2].height = 28

for i, row in df[export_cols].iterrows():
    r = 3 + i
    bg = zebra_bg(i)
    vals = list(row)
    apply_data_row(ws6, r, vals, bg=bg)
    ws6.cell(r, 1).alignment = Alignment(horizontal="left", vertical="center")
    ws6.cell(r, 2).alignment = Alignment(horizontal="left", vertical="center")
    ws6.row_dimensions[r].height = 16

raw_widths = [20, 32, 14, 13, 11, 12, 11, 11, 15, 18, 14, 14, 12]
set_col_widths(ws6, raw_widths)


# ════════════════════════════════════════════
# SHEET 7: SQL Queries Reference
# ════════════════════════════════════════════
ws7 = wb.create_sheet("🗄️ SQL Queries")
ws7.sheet_view.showGridLines = False

ws7.merge_cells("A1:B1")
t7 = ws7["A1"]
t7.value = "🗄️  SQL Queries Used in This Analysis"
t7.font = Font(bold=True, size=13, color=WHITE, name="Arial")
t7.fill = hdr_fill(BLUE)
t7.alignment = Alignment(horizontal="center", vertical="center")
ws7.row_dimensions[1].height = 32

sql_entries = [
    ("Category Nutrition Summary",
     'SELECT "Menu Category", COUNT(*) AS Item_Count,\n'
     'ROUND(AVG("Energy (kCal)"), 1) AS Avg_Calories,\n'
     'ROUND(MAX("Energy (kCal)"), 1) AS Max_Calories,\n'
     'ROUND(AVG("Protein (g)"), 1) AS Avg_Protein_g,\n'
     'ROUND(AVG("Total fat (g)"), 1) AS Avg_Fat_g,\n'
     'ROUND(AVG("Sodium (mg)"), 1) AS Avg_Sodium_mg\n'
     'FROM menu\n'
     'GROUP BY "Menu Category"\n'
     'ORDER BY Avg_Calories DESC'),
    ("Top 10 Highest Calorie Items",
     'SELECT "Menu Category", "Menu Items", "Energy (kCal)",\n'
     '"Protein (g)", "Total fat (g)", "Sodium (mg)"\n'
     'FROM menu\n'
     'ORDER BY "Energy (kCal)" DESC\n'
     'LIMIT 10'),
    ("Top 10 Highest Protein Items",
     'SELECT "Menu Category", "Menu Items", "Protein (g)",\n'
     '"Energy (kCal)", ROUND(Protein_Per_100kcal, 2) AS Protein_per_100kcal\n'
     'FROM menu\n'
     'ORDER BY "Protein (g)" DESC\n'
     'LIMIT 10'),
    ("High Sodium Alert (>800mg)",
     'SELECT "Menu Category", "Menu Items",\n'
     '"Sodium (mg)", "Energy (kCal)"\n'
     'FROM menu\n'
     'WHERE High_Sodium = 1\n'
     'ORDER BY "Sodium (mg)" DESC'),
    ("Healthiest Picks Filter",
     'SELECT "Menu Category", "Menu Items",\n'
     '"Energy (kCal)", "Protein (g)", "Sodium (mg)",\n'
     'ROUND(Protein_Per_100kcal, 2) AS Protein_per_100kcal\n'
     'FROM menu\n'
     'WHERE "Energy (kCal)" < 400\n'
     '  AND "Protein (g)" > 8\n'
     '  AND "Sodium (mg)" < 700\n'
     'ORDER BY Protein_Per_100kcal DESC\n'
     'LIMIT 10'),
]

row_ptr = 3
for title, query in sql_entries:
    ws7.cell(row_ptr, 1, value=f"▶  {title}").font = Font(
        bold=True, size=10, color=WHITE, name="Arial Narrow")
    ws7.cell(row_ptr, 1).fill = hdr_fill(BLUE)
    ws7.cell(row_ptr, 1).alignment = Alignment(vertical="center")
    ws7.row_dimensions[row_ptr].height = 22
    row_ptr += 1

    for line in query.split("\n"):
        c = ws7.cell(row_ptr, 1, value=line)
        c.font = Font(name="Courier New", size=9, color=DARK_GRAY)
        c.fill = hdr_fill(LIGHT_GRAY)
        c.alignment = Alignment(vertical="center", indent=1)
        ws7.row_dimensions[row_ptr].height = 16
        row_ptr += 1

    row_ptr += 1  # blank line between queries

ws7.column_dimensions["A"].width = 80


# ─────────────────────────────────────────────
# Save final output
# ─────────────────────────────────────────────

OUT_PATH = "/mnt/user-data/outputs/McDonalds_India_Nutrition_Analysis.xlsx"
wb.save(OUT_PATH)

print(f"\n[EXCEL] Workbook saved → {OUT_PATH}")
print("\n✅  Pipeline complete! 6 sheets generated:")
print("   1. 📊 Dashboard — KPIs + Category Summary")
print("   2. 🔥 High Calorie — Top 10 + Bar Chart")
print("   3. 💪 High Protein — Top 10 + Bar Chart")
print("   4. ⚠️  High Sodium — Alert table")
print("   5. ✅ Healthiest Picks — Filtered recommendations")
print("   6. 📋 Raw Data — Full cleaned dataset")
print("   7. 🗄️  SQL Queries — Reference sheet")
print("=" * 55)
