# 📊 Retail Sales & Customer Analytics Dashboard (Excel & Power Query)

> An interactive dual-view Excel analytics dashboard tracking revenue trends, regional performance, and customer purchasing patterns. Built using Power Query for data cleansing and type enforcement, Pivot Tables for multidimensional aggregation, and VBA for dynamic filter resets.

---

## 🖥️ Dashboard Previews

### 1. Sales Performance Dashboard
![Sales Dashboard](assets/Sales%20Dashboard.png)

### 2. Customer Analytics Dashboard
![Customer Dashboard](assets/Customer%20Dashboard.png)

---

## 📌 Key Metrics & Business Questions Answered

### 1. Key Performance Indicators (KPIs)
* **Total Orders:** 9,994 transactions processed across the entire timeline.
* **Total Revenue:** $2,297,138.90 in cumulative sales.
* **Total Profit:** $286,397.02 generated overall.
* **Total Customers:** 793 unique buyers served.
* **Total Units Sold:** 1,862 distinct items moved across categories.

### 2. Business Questions Answered

#### 📈 Sales Performance Analysis
* **What are the peak sales periods during the year?**
  * Tracks seasonal fluctuations across all 12 months, identifying revenue acceleration in Q3 and Q4 (notably in September and November).
* **Which sales quarters contribute the most to annual revenue?**
  * Evaluates revenue breakdown across Q1, Q2, Q3, and Q4 to measure quota progression.
* **How is revenue distributed across geographical regions?**
  * Uses a radar breakdown to compare revenue contribution across the **West**, **East**, **Central**, and **South** regions.

#### 👥 Customer & Market Demographics
* **Which states represent the highest market penetration?**
  * Highlights top customer volume hubs, led by **California**, **New York**, and **Texas**.
* **Which regions hold the largest customer base?**
  * Indicates that the **West** and **East** regions maintain the highest concentration of active buyers.
* **What product sub-categories attract the most buyers?**
  * Identifies top product categories by customer count, showing highest demand in **Binders**, **Paper**, and **Phones**.
* **How has customer acquisition evolved over time?**
  * Visualizes the multi-year distribution of transactions and active buyers across four consecutive years (2014–2017).

---

## ⚙️ ETL Transformation Workflow (Power Query)

1. **Source Data Ingestion:**
   * Loaded raw transactional retail records into Excel Power Query.

2. **Data Cleaning & Type Enforcement:**
   * Formatted `Order Date` and `Ship Date` fields to standard Date formats (`YYYY-MM-DD`).
   * Cast `Order ID`, `Customer ID`, `State`, `Region`, and `Sub-Category` fields to **Text**.
   * Cast `Sales` and `Profit` fields to **Currency / Decimal** values.
   * Cast `Quantity` and `Discount` fields to **Integer / Percentage** values.

3. **Handling Missing & Irregular Values:**
   * Inspected for null entries and empty strings in customer identifiers and location records.
   * Cleaned leading/trailing whitespaces across categorical columns.

4. **Feature Engineering:**
   * Extracted `Year` (2014–2017), `Quarter` (Q1–Q4), and `Month Name` from `Order Date` for timeline segmentation.
   * Loaded transformed dataset directly into the Data Model for Pivot Table analysis.

---

## 🗂️ Data Dictionary

| Column Name | Data Type | Description |
| :--- | :--- | :--- |
| `Order ID` | Text | Unique identifier for each sales transaction |
| `Order Date` | Date | Date when the purchase was placed |
| `Ship Date` | Date | Dispatch date of the order |
| `Customer ID` | Text | Unique identifier for the client |
| `State` | Text | Destination state of the order |
| `Region` | Text | Operational territory (West, East, Central, South) |
| `Sub-Category` | Text | Classification of items (e.g., Binders, Paper, Phones) |
| `Sales` | Decimal | Total gross revenue generated |
| `Profit` | Decimal | Net profit or loss recorded |

---

## 💻 Automation & Interactivity (VBA Macro)

A custom VBA macro is assigned to the **Clear All Filters** button to reset all connected timeline and regional slicers simultaneously:

```vba
Sub ClearAllSlicers()
    ' Loops through all slicer caches in the workbook and clears active filters
    Dim slcr As SlicerCache
    On Error Resume Next
    For Each slcr In ActiveWorkbook.SlicerCaches
        slcr.ClearManualFilter
    Next slcr
    On Error GoTo 0
End Sub
