# Retail Sales Analysis Dashboard

## Project Overview

This project focuses on analysing an Online Retail dataset using **Microsoft Excel and Power BI**.

The objective is to understand sales performance, customer behaviour, product performance and revenue trends through an interactive Power BI dashboard.

The dashboard helps answer questions such as:

- What is the total revenue generated?
- How many unique orders and customers are there?
- Which products generate the highest revenue?
- Which countries contribute the most revenue?
- How does revenue change over time?
- Which customers contribute the most revenue?

## Tools & Technologies

- **Microsoft Excel** – Data inspection, cleaning and preparation.
- **Power BI Desktop** – Data visualisation and dashboard development.
- **DAX** – KPI calculations and measures.

## Dataset Description

The dataset contains online retail transaction records. Each row represents a product line within an invoice. Therefore, the same invoice number can appear multiple times when an order contains different products.

| Column | Description |
|---|---|
| InvoiceNo | Unique invoice or transaction identifier |
| StockCode | Product identifier |
| Description | Product name or description |
| Quantity | Number of units purchased |
| InvoiceDate | Date of transaction |
| InvoiceTime | Time of transaction |
| UnitPrice | Price per unit |
| CUSTOMERid | Customer identifier |
| Country | Customer's country |
| Revenue | Revenue generated from the transaction line |

**Power BI table name:** `Sheet1`

## Excel Data Cleaning & Preparation

The following operations were performed or considered during data preparation.

### 1. Extract Date from DateTime

Extracted the date from the original date-time value.

```excel
=INT(E2)
```

Formatted the result as `DD-MM-YYYY`.

### 2. Extract Time from DateTime

Extracted the time from the original date-time value.

```excel
=MOD(E2,1)
```

Formatted the result as `HH:MM:SS`.

### 3. Calculate Revenue

Revenue is calculated by multiplying Quantity by UnitPrice.

```excel
=D2*G2
```

The calculated revenue should be checked against the original Revenue column.

### 4. Description Column

The Description column was inspected, including unusual and unclear product descriptions. It was retained for product-level analysis rather than removed.

### 5. Data Quality Considerations

- Repeated invoice numbers are expected because one invoice can contain multiple product lines.
- Negative quantities may represent returned products or reversals.
- Invoice numbers beginning with `C` may indicate cancelled transactions.
- Zero UnitPrice records should be investigated before removal.
- Blank CustomerIDs may affect customer-level analysis.
- The original dataset should be preserved before making cleaning changes.

## Power BI Dashboard

The dashboard contains two report pages.

### Page 1 – Sales Overview

This page provides an overview of the overall retail sales performance.

**KPI Cards**

| KPI | Description |
|---|---|
| Total Revenue | Total revenue generated from transactions |
| Total Customers | Number of distinct customer IDs |
| Total Orders | Number of unique invoice numbers |
| Total Products | Number of distinct product codes |

**Visualisations**

| Chart Name | Chart Type | Purpose |
|---|---|---|
| Sum of Revenue by Month | Line Chart | Analyses monthly revenue trends |
| Total Revenue by Description | Bar Chart | Identifies products generating the highest revenue |
| Total Revenue by Country | Bar Chart | Compares revenue across countries |
| Total Revenue by Year and Quarter | Column Chart | Compares revenue across different quarters and years |

**Slicers and Navigation**

- Country slicer
- Year slicer
- Clear All Slicers button
- Page navigation buttons

### Page 2 – Customer and Product Details

This page focuses on customer-level and product-level revenue analysis.

**Visualisations**

| Chart Name | Visual Type | Purpose |
|---|---|---|
| Customer and Description Breakdown | Decomposition Tree | Explores revenue by customer and product description |
| Total Revenue by CUSTOMERid | Pie Chart | Shows the revenue contribution of customer categories |

The decomposition tree allows users to explore revenue by different dimensions, such as CustomerID and Description.

The customer pie chart also displays a blank customer category where CustomerID is missing.

**Slicers and Navigation**

- Country slicer
- Year slicer
- Clear All Slicers button
- Page navigation buttons



## DAX Measures

The following DAX measures were used or prepared for KPI calculations.

### 1. Total Revenue

```dax
Total Revenue =
SUM(Sheet1[Revenue])
```

Calculates the total revenue by adding all values in the Revenue column.

### 2. Total Orders

```dax
Total Orders =
DISTINCTCOUNT(Sheet1[InvoiceNo])
```

Counts unique invoice numbers rather than individual transaction rows.

### 3. Total Customers

```dax
Total Customers =
DISTINCTCOUNT(Sheet1[CUSTOMERid])
```

Counts the distinct customer IDs.

To exclude blank CustomerIDs:

```dax
Total Customers Excluding Blank =
CALCULATE(
    DISTINCTCOUNT(Sheet1[CUSTOMERid]),
    Sheet1[CUSTOMERid] <> BLANK()
)
```

### 4. Total Products

```dax
Total Products =
DISTINCTCOUNT(Sheet1[StockCode])
```

Counts the number of unique product codes.

### 5. Total Quantity Sold

```dax
Total Quantity Sold =
CALCULATE(
    SUM(Sheet1[Quantity]),
    Sheet1[Quantity] > 0
)
```

Calculates the total quantity using only positive quantities.

### 6. Average Order Value

```dax
Average Order Value =
DIVIDE(
    [Total Revenue],
    [Total Orders],
    0
)
```

Calculates the average revenue generated per order.

### 7. Total Returned Units

```dax
Total Returned Units =
CALCULATE(
    SUMX(
        Sheet1,
        ABS(Sheet1[Quantity])
    ),
    Sheet1[Quantity] < 0
)
```

Calculates the total returned units by converting negative quantities into positive values.

### 8. Return Rate

```dax
Return Rate =
DIVIDE(
    [Total Returned Units],
    [Total Quantity Sold] + [Total Returned Units],
    0
)
```

Calculates the percentage of returned units relative to total sold and returned units.

Format this measure as a percentage in Power BI.

## Important DAX Notes

- All the above calculations are DAX measures, not calculated columns.
- Measures respond dynamically to report filters and slicers.
- Total Revenue is a direct sum of the supplied Revenue column.
- The revenue measure may include cancellation or return-related amounts if they exist in the source data.
- Average Order Value uses Total Revenue divided by Total Orders.
- Return Rate is based on positive and negative quantities.

## Key Business Insights

The dashboard is designed to help users:

1. Monitor overall retail revenue.
2. Identify the products generating the highest revenue.
3. Compare sales performance across countries.
4. Understand monthly and quarterly revenue trends.
5. Identify customers contributing the most revenue.
6. Explore customer and product revenue relationships.

## Interview Explanation

**Project Introduction:**

"I developed an Online Retail Sales Analysis dashboard using Microsoft Excel and Power BI. I first inspected and prepared the transaction data in Excel, including extracting date and time values and checking revenue calculations.

I then used DAX measures to calculate key performance indicators such as Total Revenue, Total Orders, Total Customers and Total Products.

I created interactive visualisations to analyse monthly revenue trends, product performance, country-wise revenue, quarterly sales and customer contributions. I also added slicers and page navigation to make the dashboard interactive and easier to explore."

## Future Improvements

Possible improvements for future versions:

- Month-over-month revenue analysis.
- Dedicated returns and cancellation analysis.
- Customer segmentation.
- Repeat customer analysis.
- Top 10 product and customer rankings.
- More detailed sales performance analysis.

## Project Summary

| Category | Details |
|---|---|
| Project | Retail Sales Analysis |
| Domain | Retail / E-commerce |
| Tools | Microsoft Excel, Power BI |
| Language | DAX |
| Data Preparation | Excel |
| Dashboard | Power BI |
| Report Pages | 2 |
| Main Analysis | Revenue, Orders, Customers, Products, Countries and Trends |

**Note:** The KPI values and charts should be interpreted using the actual report filters and source data. The screenshots show rounded dashboard values, not independently recalculated figures.
