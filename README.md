# DecodeLabs Data Analytics Track | Project 2: Exploratory Data Analysis (EDA) 🔍

**Author:** Sandhiya Kumari  
**Program:** Industrial Training Program | Batch 2026 | DecodeLabs  
**Project Goal:** Understand patterns, trends and distributions in the cleaned orders dataset, and turn them into business insight  

---

## Executive Summary

**Problem statement:** What do 1,200 e-commerce orders (Jan 2023 to Jun 2025, $1.26M total value) say about business performance, and which patterns, risks and outliers deserve action?

**Methodology (IPO framework):**
- **Input:** `Cleaned_Dataset.xlsx` from Project 1 (0 duplicates, 0 missing values, verified formats)
- **Process:** descriptive statistics (count, mean, median, five-number summary), outlier screening with IQR and Z-score, trend analysis, group comparisons (Kruskal-Wallis, chi-square), Pearson correlation
- **Output:** 7 charts, a formula-driven statistics workbook and a full PDF report

## Key Findings

| # | Headline | Evidence |
| :--- | :--- | :--- |
| 1 | **41% of order value was cancelled or returned** | 497 of 1,200 orders; $519,674 of $1,264,762. Only 19% ($242,600) is delivered. |
| 2 | **First-half revenue fell about 10% in each of the last two years** | H1 revenue: $286.5k (2023), $257.1k (2024), $231.9k (2025). Down 19.1% overall: orders -8.3%, mean order value -11.7%. |
| 3 | **A typical order is $824, not $1,054** | Order value is right-skewed (skewness 0.89); the mean overstates the typical order by 28%. Use the median. |
| 4 | **No segment moves order value** | Product, payment method, referral channel, coupon use and weekday: all p > 0.4. Coupon and non-coupon orders have the same median order ($820 vs $834). |
| 5 | **99% of customers bought only once** | 1,178 of 1,189 customers. |
| 6 | **8 unusually large orders are valid bulk orders** | Flagged by IQR (> $3,330), missed by Z-score. Quantity = 5, each equals Quantity x UnitPrice; kept as signal. |

![Order value distribution](charts/01_order_value_distribution.png)

### What the data does not show (and why that matters)

- **The revenue trend is a warning sign, not a proven trend.** The month-by-month linear trend is not statistically significant (p = 0.11) because monthly swings are large. The H1-to-H1 declines are consistent two years running.
- **June may be a seasonal peak, but only as a hypothesis.** It was the top H1 month in 2024 and 2025, third in 2023.
- **Correlations are structural.** TotalPrice = Quantity x UnitPrice, so its correlation with those columns (0.72 and 0.62) is arithmetic, not a discovery. No causation is claimed.
- **Several fields are suspiciously even** (order status, product, channel, weekday). Confirm them against the source system before big decisions.

## Recommendations

1. **Investigate fulfilment first:** add cancellation and return reason codes. Each 1-point cut in cancelled/returned value is worth about $12.6k over the 30-month period.
2. **Report median order value** in dashboards.
3. **A/B test coupons** before expanding them: 74% of orders already carry one, with no basket-size lift.
4. **Launch a repeat-purchase programme** and track repeat rate (currently 0.9% of customers).
5. **Monitor the H1 revenue decline** and re-test the trend as more months arrive.

## Charts

| | |
| :--- | :--- |
| ![Outliers](charts/02_outliers_boxplot.png) | ![Trend](charts/03_monthly_revenue_trend.png) |
| ![H1 YoY](charts/04_h1_year_over_year.png) | ![Status](charts/05_order_status.png) |
| ![Segments](charts/06_segment_comparison.png) | ![Correlation](charts/07_correlation_heatmap.png) |

## Repository Structure

```text
.
├── Cleaned_Dataset.xlsx      # Input: output of Project 1
├── eda_analysis.py           # Reproducible EDA pipeline (stats, charts, workbook)
├── EDA_Report_Project2.pdf   # Full report with executive summary and recommendations
├── EDA_Statistics.xlsx       # Statistics workbook (382 live formulas)
├── eda_stats.json            # Every number used in the report
├── charts/                   # 7 PNG charts
└── README.md                 # This file
```

`EDA_Statistics.xlsx` sheets: **Data**, **Descriptive_Stats** (mean, median, count, five-number summary, IQR and Z-score outlier counts), **By_Category**, **Monthly_Trend**, **Correlation**, **Outliers**, **Stat_Tests**. Values are live Excel formulas; the hypothesis-test p-values are static (computed in SciPy) and shown in blue.

## How to Reproduce

Requirements: Python 3 with `pandas`, `numpy`, `scipy`, `matplotlib` and `openpyxl`.

```bash
pip install pandas numpy scipy matplotlib openpyxl
python eda_analysis.py Cleaned_Dataset.xlsx .
```

This writes `charts/`, `EDA_Statistics.xlsx` and `eda_stats.json` into the current folder. Formula results in the workbook appear once it is opened in Excel (or recalculated in LibreOffice).

## Tools

Python (pandas, SciPy, Matplotlib, openpyxl) for analysis and charts; Excel formulas for the statistics workbook; ReportLab for the PDF report.

**Skills demonstrated:** descriptive statistics, outlier detection (IQR vs Z-score), trend and segment analysis, correlation analysis, hypothesis testing, data storytelling and business-impact framing.

---

*DecodeLabs | Professional Standard*
