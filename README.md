# My Spending

A spending dashboard built from a Rocket Money transaction export.

```
pip install pandas openpyxl
python3 build.py path/to/All_Transactions.xlsx output/spending_dashboard.html
```

Open `output/spending_dashboard.html` in a browser. It is one self-contained file that works offline.

## What the build does

1. **Cleans the export**: removes duplicate rows from repeated imports and from the same card being linked twice, and flips Fidelity's sign convention.
2. **Classifies every transaction** with its own rules (`build.py` → `R`) into 42 categories and 8 groups, and cleans up vendor names.
3. **Excludes non-spending**: credit-card payments, transfers between your own accounts, investing, income, and rewards.
4. **Embeds the data** into `template.html`.

## Rent + car note switch
A switch at the top hides rent and car-loan payments from every chart except **Spending by year**. That chart always shows them as their own stacked segment, next to everything else. Top category, top vendor, vendor loyalty and biggest months always leave them out. Your choice is remembered in the browser.

## Dashboard tab
Spending by year (plus a same-period comparison for the current year), year over year by month, month by month, month-over-month change, spending velocity (monthly, quarterly and yearly pace), categories, category mix, a category-by-year heatmap, top vendors, a vendor directory and a searchable transaction list. Filter by period and by category group.

## Deep dive tab
Key findings, income vs. spending, essentials vs. discretionary, what grew and what shrank, weekly rhythm, the payday effect, food habits, small purchases, subscriptions, fees and interest, seasonality, purchase size, vendor loyalty, peak months, and recommendations.

> Your transaction export and the generated HTML are git-ignored on purpose. This repository is public.
