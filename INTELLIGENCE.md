<div align="center">

# 📊 DATA INTELLIGENCE
### Sales Analysis • Customer Behaviour • Product Performance

**An Excel-based Data Analytics Project**

![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Project](https://img.shields.io/badge/Project-Data%20Analytics-2563EB?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Completed-16A34A?style=for-the-badge)

**Author:** Abhishek Choudhary  
**Workbook:** `data intelligence.xlsx`

</div>

---

## 📌 1. Project Overview

The **Data Intelligence** project analyses sales transaction data to understand product demand, customer purchasing behaviour, regional performance, payment preferences and sales patterns.

The project uses Microsoft Excel features such as PivotTables, PivotCharts, slicers, timelines, formulas, descriptive statistics and regression analysis to turn raw transaction records into useful business information.

## 🎯 2. Project Objectives

- Analyse overall sales and transaction performance.
- Identify the most frequently purchased products.
- Find repeat customers and understand customer behaviour.
- Compare product categories and regional sales.
- Examine payment-method and customer-segment patterns.
- Use PivotTables, slicers and timelines for interactive analysis.
- Apply statistical and regression analysis where appropriate.
- Present findings as clear business insights and recommendations.

## ❓ 3. Problem Statements

1. Which products are purchased most frequently?
2. Which products generate the highest sales value?
3. Which customers appear most frequently in the transaction data?
4. Which product category contributes the most to sales?
5. Which region records the highest sales value?
6. Which payment method is used most often?
7. How do sales transactions vary over time?
8. Can the dataset be explored interactively using PivotTables, slicers and timelines?
9. What do descriptive statistics and the regression output reveal about the data?

## 🗂️ 4. Dataset Description

The workbook contains a transaction dataset with **250 records** and fields including:

| Field | Description |
|---|---|
| Transaction_ID | Unique transaction reference |
| Date | Transaction date |
| Customer_ID | Customer reference |
| Customer_Name | Customer name |
| Product_ID | Product reference |
| Product_Name | Purchased product |
| Category | Product category |
| Quantity | Number of units purchased |
| Unit_Price | Price per unit |
| Payment_Method | Payment method used |
| Region | Sales region |
| Customer_Segment | Basic, Standard or Premium |
| Customer_Since | Customer start date |
| Total_Amount | Transaction value field |

**Coverage in the supplied data:** 11 April 2024 to 11 April 2025; 50 distinct customer IDs; 10 product names; 3 product categories; 5 regions.

## 🧰 5. Tools and Excel Features Used

- **PivotTables and PivotCharts** — summarise and compare transaction data.
- **Slicers** — filter PivotTable results interactively.
- **Timeline** — filter transaction records by date.
- **Conditional Formatting** — highlight important values and performance indicators.
- **Descriptive Statistics** — review measures such as mean, median, standard deviation, minimum and maximum.
- **Regression Analysis** — examine the relationship represented by the regression output in the workbook.
- **Excel formulas** — calculate and identify useful information.

### Formula examples

| Purpose | Example formula |
|---|---|
| Calculate transaction amount | `=Quantity_Cell*Unit_Price_Cell` |
| Count a product's occurrences | `=COUNTIF(Product_Range,Product_Cell)` |
| Find the largest frequency | `=MAX(Frequency_Range)` |
| Return the most frequent product from a frequency table | `=INDEX(Product_List,MATCH(MAX(Frequency_List),Frequency_List,0))` |
| Find unique names (Excel 365) | `=UNIQUE(Name_Range)` |
| Filter repeated names (Excel 365) | `=UNIQUE(FILTER(Name_Range,COUNTIF(Name_Range,Name_Range)>1))` |
| Total sales | `=SUM(Total_Amount_Range)` |
| Average transaction value | `=AVERAGE(Total_Amount_Range)` |

Replace the example range names with the actual cell ranges in your worksheet.

## 📈 6. Key Findings and Insights

The following findings are calculated from the 250 transaction rows in the supplied workbook. Transaction value is calculated as **Quantity × Unit_Price** so that the first record, whose `Total_Amount` cell is blank, is included consistently.

| Metric | Finding |
|---|---:|
| Transactions | 250 |
| Distinct customers | 50 |
| Product names | 10 |
| Total units sold | 753 |
| Calculated sales value | 229,192.47 |
| Average transaction value | 916.77 |
| Highest transaction value | 4,499.95 |
| Lowest transaction value | 59.99 |

### 🛒 Product insights

- **Bookshelf** is the most frequently appearing product, with **35 transactions**.
- **Laptop** has the highest calculated sales value at **67,499.25**.
- **Smartphone** is close behind, contributing **67,199.04** in calculated sales value.
- Frequency and revenue tell different stories: the most frequently purchased product is not the product with the highest sales value.

### 🧩 Category and region insights

- **Electronics** contributes the largest calculated sales value: **171,756.05**.
- **East** is the highest-sales region at **59,288.39**.
- **North** contributes **50,808.31**.
- These values describe this dataset and should not be treated as forecasts.

### 👥 Customer and payment insights

- **Mark Carter** appears most frequently, with **12 transactions**.
- **Cash** is the most frequently recorded payment method, with **76 transactions**.
- The workbook contains **50 distinct customer IDs**.

## 🧠 7. Regression Analysis

The `REGRESSION` worksheet reports:

| Statistic | Workbook result |
|---|---:|
| Multiple R | 0.8354 |
| R Square | 0.6979 |
| Adjusted R Square | 0.6967 |
| Standard Error | 151.6460 |
| Observations | 249 |
| F-statistic | 570.7310 |
| Significance F | 3.7426E-66 |

The reported R Square is approximately **0.698**, meaning the fitted model accounts for about **69.8% of the variation in its dependent variable** in the analysed sample.

**Interpretation note:** The regression output does not clearly identify the X-variable and Y-variable names in its visible summary. Therefore, this report does not assign a business meaning to the coefficient or claim that one specific factor causes another. Confirm the original regression input ranges before making that interpretation.

## 💡 8. Business Recommendations

1. **Inventory planning:** Review stock levels for Bookshelf, the most frequently appearing product.
2. **Revenue strategy:** Study Laptop and Smartphone pricing, availability and promotion performance because they lead sales value.
3. **Regional planning:** Investigate what is driving the East region's higher sales and whether successful practices can be applied elsewhere.
4. **Customer retention:** Review repeat-purchase behaviour and consider suitable loyalty offers.
5. **Payment experience:** Keep cash payment available while monitoring demand for digital payment methods.
6. **Ongoing monitoring:** Use PivotTable filters, slicers and timelines to review performance by product, region, segment and date.

## 📊 9. Workbook Structure

| Worksheet | Purpose |
|---|---|
| `Data intelligence dashboard` | Transaction-level dataset and analysis fields |
| `PIVOT TABLE` | PivotTable-related analysis and supporting calculations |
| `TIMELINE AND SLICER` | Customer/date summary for interactive filtering |
| `REGRESSION` | Regression statistics and ANOVA output |

## ⚠️ 10. Data Notes and Limitations

- One transaction row has a blank value in the `Total_Amount` column. For the headline sales calculations in this README, its value was recomputed using `Quantity × Unit_Price`.
- The regression worksheet contains 249 observations, while the transaction table contains 250 records. The regression sample may therefore exclude one record.
- The regression variable labels are not sufficiently clear in the visible output to support a specific predictor/outcome interpretation.
- These results describe the supplied workbook only. They do not establish causation or predict future sales.

## 🏁 11. Conclusion

This project demonstrates how Excel can be used to convert transaction-level data into meaningful business insights. The analysis identifies Bookshelf as the most frequently appearing product, Laptop as the highest-value product by calculated sales, Electronics as the leading category and East as the highest-sales region.

Using PivotTables, charts, slicers, timelines, formulas and statistical tools makes the data easier to explore and communicate. The findings can support inventory reviews, product planning, customer-retention activities and regional performance monitoring.

---

<div align="center">

### ✨ DATA → ANALYSIS → INSIGHTS → DECISIONS

**Project created by Abhishek Choudhary**  
*Data Intelligence | Microsoft Excel*

</div>
