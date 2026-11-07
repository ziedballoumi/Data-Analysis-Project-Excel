# Sales Analysis Project --- Excel & Power Query

## 1. Project Overview

This project demonstrates an end-to-end sales data analysis workflow
using **Microsoft Excel** and **Power Query**. It starts with a
deliberately messy sales dataset, cleans and prepares the data, then
uses Excel analysis tools and charts to answer business questions
through a set of interactive dashboards.

**Main objective:** turn raw, inconsistent sales data into a reliable
dataset and a clear sales analysis report that supports business
decisions.

## 2. Dataset Overview

The practice dataset contains approximately **600 rows** and the
following original columns:

  Column         Description
  -------------- --------------------------------
  `Order_ID`     Identifier of an order
  `Order_Date`   Date associated with the order
  `Customer`     Customer name or identifier
  `Region`       Sales region
  `Product`      Product name
  `Category`     Product category
  `Unit_Price`   Price per unit
  `Quantity`     Number of units sold

A calculated column, `Revenue`, is added during preparation:

`Revenue = Unit_Price × Quantity`

The dataset is intended for learning and portfolio practice. Any
generated or simulated values should not be presented as verified
real-world company data.

## 3. Raw Dataset Problems

The raw dataset was intentionally designed to include common
data-quality issues:

-   **Missing values:** some fields contain blank or null values,
    particularly `Product` and `Unit_Price`.
-   **Duplicate records:** repeated rows or repeated order records may
    exist and need investigation.
-   **Inconsistent text:** product, customer, category, or region names
    may use inconsistent capitalization.
-   **Extra spaces:** leading, trailing, or repeated spaces can cause
    values that should match to be treated as different.
-   **Inconsistent date formats:** dates may appear in different textual
    formats and need to be converted to a consistent Date type.
-   **Inconsistent price formatting:** `Unit_Price` may contain currency
    symbols, text, or values stored as text rather than numbers.
-   **Invalid numeric values:** `Quantity` and/or `Unit_Price` may
    contain zero, negative, error, or otherwise implausible values that
    require validation.

Not every issue should be fixed automatically. For example, repeated
`Order_ID` values can be valid when an order contains multiple products.
Duplicate removal should therefore be based on the correct row-level
business key, not on `Order_ID` alone.

## 4. Data Cleaning with Power Query

The cleaning workflow is performed in Power Query. Keep the original
source unchanged and apply transformations in a query so that the
process can be refreshed.

### Step 1 --- Connect to the source

1.  Open Excel.
2.  Select **Data → Get Data** and import the workbook or table.
3.  Choose **Transform Data** to open Power Query Editor.
4.  Check that the first row is used as headers and that column names
    are correct.

### Step 2 --- Set appropriate data types

Set the columns to suitable types:

-   `Order_ID`, `Customer`, `Region`, `Product`, and `Category`: Text
    (unless the identifier needs another type).
-   `Order_Date`: Date.
-   `Unit_Price`: Decimal Number or Fixed Decimal Number, depending on
    the required precision.
-   `Quantity`: Whole Number.

For mixed date formats, use **Change Type → Using Locale** where
appropriate. Select the locale that matches the source date format.
Verify the converted dates rather than assuming the conversion
succeeded.

### Step 3 --- Standardize text

For relevant text columns:

1.  Select the column.
2.  Use **Transform → Format → Trim** to remove leading and trailing
    spaces.
3.  Use **Transform → Format → Clean** to remove non-printing
    characters.
4.  Apply a consistent capitalization convention, such as **Capitalize
    Each Word** or **UPPERCASE**, where appropriate.
5.  Check for spelling variants that formatting alone cannot fix.

Apply the same standardization to lookup/reference tables before merging
them.

### Step 4 --- Handle duplicate records

1.  Inspect repeated rows and repeated identifiers.
2.  Determine what uniquely identifies a sales line.
3.  Remove duplicates only when the records are genuinely duplicate
    transactions.
4.  Do not remove rows solely because `Order_ID` repeats: one order may
    contain several products or line items.

Keep a record of the number of rows before and after deduplication.

### Step 5 --- Handle missing `Product` values

A missing product can sometimes be recovered from other reliable fields,
but only when those fields identify one unambiguous product.

A practical lookup approach:

1.  Create or use a reference table with the known combination of
    `Region`, `Category`, and `Unit_Price` mapped to `Product`.
2.  Standardize the text and data types in both tables.
3.  In Power Query, select **Home → Merge Queries**.
4.  Match the raw table to the reference table using the relevant
    columns: `Region`, `Category`, and `Unit_Price`.
5.  Use a **Left Outer** join to retain all original sales rows.
6.  Expand the matched `Product` field.
7.  Use the lookup value only where the original `Product` is null.
8.  Review unmatched rows and ambiguous matches manually.

**Important:** this method is valid only if the combination of region,
category, and unit price maps to one product. If multiple products share
that combination, do not guess; leave the value unresolved or
investigate another source.

### Step 6 --- Handle missing `Unit_Price` values

A missing price can be recovered using a product-price reference table
when each product has a known, appropriate price.

1.  Create a reference table containing `Product` and its valid
    reference price.
2.  Ensure product names are standardized in both tables.
3.  Merge the sales query with the reference table using `Product`.
4.  Use a **Left Outer** join.
5.  Expand the reference price column.
6.  Add a custom column that keeps the original price when present and
    uses the reference price only when the original is null.

Example Power Query custom-column expression (adjust names to match the
actual query):

``` powerquery
if [Unit_Price] = null
then [Reference_Price]
else [Unit_Price]
```

7.  Rename the resulting field to `Final_Unit_Price` or replace the
    original after validation.
8.  Review products without a reference price.

If a product can have different prices depending on date, region, or
another condition, the lookup key must include those fields. Do not
apply one fixed price when the business rules do not support it.

### Step 7 --- Validate numeric fields

-   Convert `Unit_Price` to a numeric type after removing or
    interpreting currency symbols and text.
-   Convert `Quantity` to a whole number.
-   Inspect errors, nulls, zeros, and negative values.
-   Decide whether zero or negative values are invalid or represent
    legitimate returns, corrections, or adjustments. Do not delete them
    without understanding the business meaning.
-   Replace errors with null only when that is an appropriate documented
    treatment; then decide how to handle the resulting nulls.

### Step 8 --- Create the Revenue column

After price and quantity have been cleaned, add a Custom Column in Power
Query:

``` powerquery
[Final_Unit_Price] * [Quantity]
```

Name the new column `Revenue` and set its type to Decimal Number or
Fixed Decimal Number.

If the original price column was retained and validated, use
`[Unit_Price]` instead.

### Step 9 --- Final quality checks and load

Before loading the data:

-   Confirm column names and data types.
-   Check that required fields are populated or that unresolved nulls
    are documented.
-   Check for remaining errors and unexpected values.
-   Reconcile row counts before and after transformations.
-   Check that `Revenue` is calculated correctly on sample records.
-   Load the cleaned result to an Excel table, for example
    `Cleaned_Data`.

## 5. Business Analysis Questions

The report is organized around four areas of analysis.

### Dashboard 1 --- Sales Performance Overview

**Objective:** understand overall sales performance and its evolution
over time.

Business questions:

1.  What is the total revenue?
2.  How many distinct orders were recorded?
3.  What is the total quantity sold?
4.  What is the average order value?
5.  Which month generated the highest revenue?
6.  How does the number of orders change over time?

Suggested KPIs and charts:

-   KPI cards: Total Revenue, Distinct Orders, Total Quantity Sold,
    Average Order Value.
-   Line chart: Revenue by month.
-   Column or bar chart: Number of orders by month.

Useful definitions:

-   **Total Revenue:** sum of `Revenue`.
-   **Distinct Orders:** distinct count of `Order_ID`, if an order can
    have multiple line items.
-   **Total Quantity Sold:** sum of `Quantity` (subject to the treatment
    of returns).
-   **Average Order Value:** total revenue divided by distinct orders.

### Dashboard 2 --- Product & Category Analysis

**Objective:** identify the categories and products that contribute most
to sales.

Business questions:

1.  Which category generates the highest revenue?
2.  Which category has the highest sales volume?
3.  Which five products generate the highest revenue?
4.  Which five products have the highest quantity sold?
5.  Which products generate the lowest revenue?

Suggested charts:

-   Bar chart: Revenue by category.
-   Doughnut chart: Quantity sold by category.
-   Horizontal bar chart: Top 5 products by revenue.
-   Horizontal bar chart: Top 5 products by quantity sold.

### Dashboard 3 --- Regional Sales Analysis

**Objective:** compare sales performance across regions.

Business questions:

1.  Which region generates the highest revenue?
2.  Which region has the highest number of orders?
3.  Which region has the highest sales volume?
4.  Which region contributes the least to total revenue?

Suggested charts:

-   Bar or column chart: Revenue by region.
-   Doughnut chart: Distinct orders by region.
-   Bar chart: Quantity sold by region.

### Dashboard 4 --- Customer Analysis

**Objective:** identify valuable customers and understand customer
purchasing activity.

Business questions:

1.  Who are the top 10 customers by revenue?
2.  Which customers place the highest number of orders?
3.  How much revenue do the top customers contribute?
4.  Which customers generate the lowest revenue?

Suggested charts:

-   Horizontal bar chart: Top 10 customers by revenue.
-   Bar chart: Number of distinct orders by top customers.
-   Doughnut chart: Revenue contribution of selected top customers
    versus other customers.

## 6. Excel Visualization and Report Design

The final deliverable is an Excel workbook containing the cleaned data
and four dashboard worksheets.

Suggested workbook structure:

  Worksheet                 Contents
  ------------------------- ----------------------------------------------
  `Cleaned_Data`            Final cleaned dataset and calculated Revenue
  `Dashboard_1_Sales`       Sales KPIs, monthly revenue, monthly orders
  `Dashboard_2_Products`    Category and product performance charts
  `Dashboard_3_Regions`     Regional revenue, orders, and quantity
  `Dashboard_4_Customers`   Top customers and customer contribution

Recommended Excel features:

-   **Excel Tables:** keep the cleaned source structured and expandable.
-   **PivotTables:** summarize revenue, orders, and quantity by date,
    product, category, region, and customer.
-   **PivotCharts:** create charts connected to PivotTables.
-   **Slicers:** filter by Region, Category, and Product.
-   **Timeline:** filter by Order_Date.
-   **Conditional Formatting:** highlight important values in supporting
    tables.

For interactive filtering across several PivotTables, connect slicers to
the relevant PivotTables through **Slicer → Report Connections**. This
works when the PivotTables share a compatible data source/cache.

### Visual design guidelines

-   Use a consistent navy blue, white, and light-grey palette.
-   Place KPI cards near the top of each dashboard.
-   Arrange charts in a balanced two-column layout.
-   Use clear, descriptive chart titles.
-   Remove unnecessary chart borders and worksheet gridlines.
-   Keep number formats consistent (currency, whole numbers,
    percentages).
-   Use data labels only when they improve readability.
-   Keep slicers and the date timeline in consistent positions.

## 7. Suggested Excel Calculations

Assuming the cleaned data is stored in an Excel Table named `SalesData`:

**Total Revenue**

``` excel
=SUM(SalesData[Revenue])
```

**Total Quantity Sold**

``` excel
=SUM(SalesData[Quantity])
```

**Distinct Orders**

``` excel
=COUNTA(UNIQUE(SalesData[Order_ID]))
```

**Average Order Value**

``` excel
=SUM(SalesData[Revenue])/COUNTA(UNIQUE(SalesData[Order_ID]))
```

**Quantity sold for Electronics**

``` excel
=SUMIF(SalesData[Category],"Electronics",SalesData[Quantity])
```

**Number of rows in the North region**

``` excel
=COUNTIF(SalesData[Region],"North")
```

The `COUNTIF` example counts rows, not necessarily distinct orders. Use
a PivotTable with Distinct Count or an appropriate unique-count formula
when the question asks for the number of orders.

## 8. Project Workflow Summary

1.  Import the raw Excel dataset into Power Query.
2.  Inspect the data and identify quality problems.
3.  Set correct data types.
4.  Standardize text and date formats.
5.  Investigate and remove genuine duplicate records.
6.  Recover missing Product and Unit_Price values using validated
    reference tables where possible.
7.  Validate numeric values and errors.
8.  Calculate Revenue.
9.  Load the cleaned dataset into Excel.
10. Build PivotTables and PivotCharts.
11. Create the four dashboards.
12. Add slicers and a date timeline.
13. Review the results and document unresolved data-quality limitations.

## 9. Final Deliverables

The completed project should contain:

-   The original raw dataset.
-   The Power Query cleaning workflow.
-   The cleaned Excel dataset.
-   Four Excel dashboards.
-   A sales analysis report answering the business questions.
-   This README documenting the project.

## 10. Important Notes and Limitations

-   The dataset is a practice dataset and may contain simulated
    information.
-   The example chart values and KPI placeholders used during planning
    are illustrative; calculate final results from the actual cleaned
    dataset.
-   A lookup can only recover missing information when the reference
    mapping is reliable and unambiguous.
-   Repeated order IDs are not automatically duplicates: one order may
    contain multiple product lines.
-   The meaning of zero or negative quantities and prices must be
    checked before changing or removing those records.
-   The distinct order count must be used when an order can appear on
    multiple rows.

------------------------------------------------------------------------

**Project focus:** Data Cleaning · Power Query · Excel · PivotTables ·
PivotCharts · KPI Analysis · Business Intelligence
