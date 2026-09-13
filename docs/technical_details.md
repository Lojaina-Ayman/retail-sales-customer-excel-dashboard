# Technical Details & ETL Documentation

---

## 1. Power Query ETL Workflow

* **Data Type Enforcement:**
  * Cleaned and converted date fields (`Order Date`, `Ship Date`) into standardized Date formats (`YYYY-MM-DD`).
  * Enforced string/text format on identifiers (`Order ID`, `Customer ID`) and categorical columns (`Region`, `State`, `Sub-Category`).
  * Converted monetary and numerical figures (`Sales`, `Profit`) into standard decimal currency types.
* **Feature Engineering:**
  * Extracted temporal attributes from `Order Date`:
    * Year (`2014`, `2015`, `2016`, `2017`)
    * Quarter (`Q1`, `Q2`, `Q3`, `Q4`)
    * Month (`Jan` - `Dec`)
* **Data Hygiene:**
  * Trimmed leading and trailing whitespaces.
  * Checked for and removed blank or invalid transaction rows before loading into Pivot Tables.

---

## 2. VBA Automation (Clear All Filters Macro)

### Where to Find / Add This in Excel:
1. Press `ALT + F11` inside Excel to open the **VBA Editor**.
2. Click **Insert** > **Module**.
3. Paste the code below into the module window.
4. Return to your dashboard, right-click your **Clear All Filters** shape/button, choose **Assign Macro**, and select `ClearAllSlicers`.

```vba
Sub ClearAllSlicers()
    ' Loops through all slicers in the workbook and clears active selections
    Dim slcr As SlicerCache
    On Error Resume Next
    For Each slcr In ActiveWorkbook.SlicerCaches
        slcr.ClearManualFilter
    Next slcr
    On Error GoTo 0
End Sub
