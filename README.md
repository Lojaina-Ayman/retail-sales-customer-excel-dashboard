# 📊 Retail Sales & Customer Analytics Dashboard (Excel)

> An interactive dual-view Excel dashboard tracking business revenue, regional demand, and customer purchasing behavior. Built using Power Query for ETL, Pivot Tables for multidimensional aggregation, and VBA for dynamic filter resets.

---

## 🖥️ Dashboard Previews

### 1. Sales Performance Dashboard
![Sales Dashboard](assets/Sales%20Dashboard.png)

### 2. Customer Analytics Dashboard
![Customer Dashboard](assets/Customer%20Dashboard.png)

---

## 📌 Key Metrics & Business Questions Answered

### 1. Key Business Metrics (KPIs)
* **Total Orders:** 9,994 transactions processed across the entire timeline.
* **Total Revenue:** $2,297,138.90 in cumulative sales.
* **Total Profit:** $286,397.02 generated overall.
* **Total Customers:** 793 unique customers served.
* **Total Units/Items Sold:** 1,862 distinct units moved across categories.

---

### 2. Business Questions Answered

#### 📈 Sales Performance Analysis
* **What are the peak sales periods?**
  * Evaluates monthly fluctuations across all 12 months, highlighting volume acceleration in September and November.
* **Which quarters drive the most revenue?**
  * Compares total sales volume across Q1, Q2, Q3, and Q4.
* **How does revenue vary geographically?**
  * Leverages a radar visual to assess regional performance across West, East, Central, and South.

#### 👥 Customer & Market Analysis
* **Which states hold the highest market penetration?**
  * Identifies primary customer volume hubs, led by California, New York, and Texas.
* **Which regions contain the largest customer base?**
  * Displays buyer volume concentration, showing the strongest presence in the West and East territories.
* **What product sub-categories attract the most buyers?**
  * Highlights high-demand categories by unique customer count, with Binders, Paper, and Phones leading.
* **How has customer activity trended over time?**
  * Tracks customer transaction volume across 2014, 2015, 2016, and 2017.

---

## 🛠️ Implementation Highlights

* **Power Query ETL:** Formatted dates, standardized data types, and extracted timeline dimensions (Year, Quarter, Month).
* **Pivot Modeling:** Built interconnected summaries driving all charts and KPI cards.
* **Interactive Navigation:** Seamless switching between dedicated **Sales** and **Customer** dashboards.
* **Filter Automation:** Custom macro linked to the **"Clear All Filters"** button to reset slicers instantly.

---

## ⚙️ Filter Reset Macro

```vba
Sub ClearAllSlicers()
    Dim slcr As SlicerCache
    For Each slcr In ActiveWorkbook.SlicerCaches
        slcr.ClearManualFilter
    Next slcr
End Sub
