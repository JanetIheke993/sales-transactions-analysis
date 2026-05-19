# sales-transactions-analysis
Sales transaction analysis project in Excel using Pivot Tables, Pivot Charts, XLOOKUP, and dashboards to uncover revenue trends, product performance, regional sales insights, and business patterns.
Sales Transactions Analysis Project Documentation
Project title: Sales Transactions
Project Objective
Clean, analyze, and visualize the sales transaction dataset to produce actionable business insights. Deliverables include a cleaned dataset, summary tables by Region/Product, pivot tables and pivot charts, conditional formatting to flag exceptions, and a short set of insights to guide decision-makers.
Data Preparation and Cleaning
Overview: Ensure the dataset has consistent formats, no duplicates, and numeric values for measures. All cleaning steps were performed on a sheet named Cleaning.
Steps performed
1. Remove duplicate transactions
Selected the full dataset (Ctrl+A) → Data → Remove Duplicates.
Used TransactionID as the unique key.
Result: No duplicate TransactionIDs found.
2. Standardize Date column
Selected the Date column → Data → Text to Columns → Delimited → Finish (this forces Excel to reinterpret the column where necessary).
Formatted the column as Date via Home → Number Format → More Number Formats → Custom (e.g., dd-mmm-yyyy) to ensure consistent display and correct underlying date serial values.
3. Convert Sales Amount to numeric values
Problem observed: Sales values contained currency symbols and thousands separators (e.g., $18,370) which Excel may treat as text.
Solution options used:
Select Sales column → Data → Text to Columns → Finish (this typically converts recognisable currency strings to numbers). Then set Number Format to Currency.
4. Validate data types
Verified that the Date column contains Excel date serials and Sales contains numeric values by using ISNUMBER() checks and sorting.
Analytical Methods and Calculations
All analysis was carried out on a new sheet named Analysis to keep raw data intact.
A. Total Sales per Region
Approach: Generated a list of unique Regions and used SUMIF for aggregation.
Formula used:
=SUMIF(Cleaning!$C:$C, $A3, Cleaning!$F:$F)
Cleaning!$C:$C — Region column in the Cleaning sheet.
$A3 — Region name in the Analysis sheet.
Cleaning!$F:$F — Sales column in the Cleaning sheet.
B. Pivot Table: Sales by Region and Product Category
Creation steps:
1. Click a single cell within the cleaned data.
2. Insert → PivotTable → New Worksheet (or existing sheet in the workbook).
3. Drag Region to Rows, Product Category to Columns, and Sales Amount to Values.
4. Set Values to Sum of Sales and format as Currency.
C. Average Sales Amount per Product
Approach: List unique products and use AVERAGEIF.
Formula used:
10341.42964
Cleaning!$E:$E — Product Name column.
Analysis!J3 — Product name reference in the Analysis sheet.
Cleaning!$F:$F — Sales column.
D. Lookup Sales Amount for a Transaction
Task: Retrieve the Sales amount for TransactionID = 1050.
Formula (VLOOKUP):
15200
Better modern alternative (XLOOKUP):
#NAME?
E. Sales Category Classification
Business rule:
High if Sales > 10,000
Medium if Sales between 5,000 and 10,000 (inclusive)
Low if Sales < 5,000
Formula implemented (IFS):
Low
Note: This formula assumes F2 is numeric.

Visualizations
A. Clustered Column Chart — Total Sales by Region
Source: The aggregated table created with SUMIF or the regional PivotTable.
Insert → Charts → Clustered Column.
Chart refinements: axis titles, data labels, and currency formatting on the value axis.
B. Pivot Chart — Total Sales by Product Category across Regions
Created from the PivotTable described above: Insert → PivotChart.
Chosen chart type: Clustered Column for easy comparison across categories and regions.
C. Conditional Formatting — Flagging High-Value Transactions
Objective: Highlight transactions where Sales Amount > 15,000.
Common mistake observed: Applying a conditional formula like =$F2>15000 will fail if column F contains text values (e.g., "$18,370").
Robust approach used:
Convert Sales column to numeric (preferred) OR use a formula-based rule that converts text to a numeric value before evaluating:
#VALUE!
Apply rule via Home → Conditional Formatting → New Rule → Use a formula to determine which cells to format.
Set the Applies to range to the full data rows (e.g., =$A$2:$H$12001).


5. Key Findings and Insights
•      West region achieved the highest total sales, while the East region recorded the lowest.
•      Electronics emerged as the top-performing product category, driving most of the revenue.
•      High-value transactions contributed the largest share of total sales (₦92.47M).
•      Smartphones led as the best-selling product across all regions.
•      A decline in sales trend was observed over time, indicating a need for renewed marketing or sales strategies.


6. Challenges and How They Were Addressed
Non-numeric Sales values — Fixed by converting currency-formatted text into numeric values using Text to Columns or the VALUE+SUBSTITUTE formula.
Date inconsistencies — Resolved via Text to Columns and consistent custom date formatting.
Large dataset handling — Performance was maintained by using Excel Tables (recommended) and pivot tables rather than heavy volatile formulas.


7. Appendix — Important formulas and quick references
SUMIF by region: =SUMIF(Cleaning!$C:$C, $A3, Cleaning!$F:$F)
AVERAGEIF by product: =AVERAGEIF(Cleaning!$E:$E, Analysis!J3, Cleaning!$F:$F)
IFS sales category: =IFS(F2<5000, "Low", F2<=10000, "Medium", F2>10000, "High")
VLOOKUP example: =VLOOKUP(1050, Cleaning!$A:$F, 6, FALSE)
XLOOKUP (preferred): =XLOOKUP(1050, Cleaning!$A:$A, Cleaning!$F:$F, "Not Found")
Conditional formatting numeric conversion rule: =VALUE(SUBSTITUTE(SUBSTITUTE($F2, "$", ""), ",", ""))>15000


9. Recommendations for next steps
1. Load the cleaned dataset into Power BI or Tableau for interactive dashboards and filters (slicers) across multiple visualizations.
2. Add time-series analysis (monthly trends, seasonality) and forecasts.
3. Implement cohort or RFM analysis to identify high-value customers and retention opportunities.
4. Consider storing the cleaned dataset as a CSV or in a database for reuse and reproducibility.

