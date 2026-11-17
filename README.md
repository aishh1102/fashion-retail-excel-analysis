# Fashion Retail Sales Analysis (Excel)

An end-to-end data analysis project in Excel: cleaning a messy multi-table dataset, building calculated columns, analysing sales with PivotTables, and presenting the findings in a dashboard with business recommendations.

![Dashboard](images/dashboard.png)

## Business problem

A fashion retailer with 7 physical stores and an online store wants to understand where its money comes from and where it leaks. This project answers:

1. Which months, categories and stores drive revenue and profit?
2. Which products and channels have the highest return rates?
3. Do bigger discounts increase sales, or only reduce profit?
4. Which customer groups matter most?

## Key results

| Metric | Value |
|---|---|
| Total revenue (after discounts) | ₹14.27M |
| Total profit | ₹6.97M |
| Profit margin | 48.9% |
| Overall return rate | 8.0% |

## Key findings

1. **Festive season drives sales.** October and November revenue is about double a normal month (Oct 2025: ₹1.07M vs Jun 2025: ₹0.45M).
2. **Sarees earn the most and are returned the least.** Sarees bring in ₹3.79M with a 4.5% return rate. Dresses have the highest return rate at 13.1%.
3. **Online sells the most but returns the most.** The online store is the top seller (₹3.72M, 26% of revenue), but its return rate is 17.1% against 4.7% in physical stores.
4. **Discounts cut profit without lifting units per order.** Average units per line stay near 1.2 at every discount level, while profit margin falls from 54.3% (no discount) to 20.1% (31%+ discount).
5. **A few loyal customers carry the business.** The 50 Platinum customers (5% of all customers) bring in about 30% of revenue; the 580 Bronze customers bring in 12%.

## Recommendations

- Stock up and launch campaigns before October.
- Reduce discounts above 20%, since they cost margin and do not add units.
- Investigate online returns, especially size issues on dresses, and improve size guides.
- Protect and grow the Platinum and Gold customer groups with loyalty rewards.

## Dataset

A **synthetic (computer-generated) practice dataset**, not real company data. It covers Jan 2024 to Dec 2025 and has 5 related tables:

| Table | Rows | Description |
|---|---|---|
| Sales | 5,025 | One row per order line |
| Customers | 1,020 | Demographics, city, loyalty tier |
| Products | 200 | Category, brand, fabric, MRP, cost |
| Stores | 8 | 7 physical stores and 1 online store |
| Returns | 401 | Returned items with reason and refund |

Relationships: `Sales.customer_id → Customers`, `Sales.product_id → Products`, `Sales.store_id → Stores`, `Returns.sale_id → Sales`.

The data was deliberately messy, so cleaning was part of the project.

## Data cleaning (15 issues found and logged)

Every issue, with the number of rows affected and the decision taken, is recorded in the **Issues Log** sheet.

| Table | Issue | Rows | Fix |
|---|---|---|---|
| Customers | Duplicate customers | 20 | Removed duplicates |
| Customers | Mixed gender codes (F, male, trailing spaces) | 115 | Standardised with a formula |
| Customers | Same city spelled differently (Bangalore/Bengaluru, Cochin/Kochi) | 47 | Standardised to 13 cities |
| Customers | Missing or impossible ages (0, 112, 150) | 45 | Blanked; missing ages not invented |
| Customers | Missing acquisition channel | 20 | Labelled "Unknown" |
| Products | Inconsistent category capitalisation and spaces | 9 | Standardised to 7 categories |
| Products | Missing colour | 10 | Labelled "Unknown" |
| Products | Cost price higher than MRP | 3 | Replaced with category average cost ratio × MRP (estimate) |
| Sales | Duplicate sales rows | 25 | Removed duplicates |
| Sales | Dates stored as text | 12 | Converted to real dates |
| Sales | Inconsistent payment labels (upi, CC, NetBanking) | 245 | Standardised to 6 methods |
| Sales | Missing unit price | 50 | Looked up from Products with XLOOKUP |
| Sales | Missing discount | 75 | Assumed 0% |
| Returns | Missing return reason | 16 | Labelled "Not Specified" |
| Returns | Return date before order date | 3 | Flagged; kept in refund totals, excluded from time-to-return analysis |

## Calculated columns (Sales table)

- `net_revenue = quantity × unit_price × (1 − discount)`
- `profit = net_revenue − quantity × unit_cost`
- Lookups from other tables with XLOOKUP: category, store name, loyalty tier
- `order_month`, `is_returned`, `channel` (Online / Store), `discount_band`

## Dashboard

The Dashboard sheet shows four headline numbers (revenue, profit, margin, return rate) and 7 charts, each built from its own PivotTable:

1. Revenue and profit by category
2. Monthly revenue trend
3. Revenue and profit by store
4. Return rate by category
5. Return rate: online vs store
6. Profit margin by discount level
7. Revenue by loyalty tier

## Excel skills used

XLOOKUP, SUMPRODUCT, SWITCH, IF, SUMIFS, COUNTIF, TEXT, Remove Duplicates, data validation checks, PivotTables, charts, number formatting and dashboard layout. Built in Excel for the web.

## Assumptions and limitations

- The data is synthetic. Patterns such as the festive-season peak and the effect of discounts reflect how the data was generated, so treat the findings as a demonstration of method, not real business conclusions.
- Missing discounts were assumed to be 0%.
- Three impossible cost prices were replaced with a category-based estimate.
- Customers with missing ages were left blank, so average-age figures use only customers with a valid age.

## Files

- `fashion_retail_project.xlsx`: full workbook (data, cleaning columns, Issues Log, pivots, Dashboard, Insights)
- `dashboard.png`: dashboard screenshot

---
Author: Aishwarya M
