---
name: weekly-demand-exploration
description: "Use when exploring food demand by week, summarizing trends, identifying high-demand meal categories or items, comparing fulfilment centers, calculating demand KPIs, or preparing an evidence-based weekly operations review from train.csv."
argument-hint: "Specify the demand question or operational segment to analyze"
---

# Weekly Demand Exploration

## When to use

- A user asks how orders change over time or which weeks are unusually high or low.
- A user asks which meals/categories account for most demand or which centers lead.
- A user needs demand KPIs or an operational review based on this project's weekly order data.

## Procedure

1. Read `notebooks/analysis.ipynb` and identify the existing `train`, `df`, `weekly`, and enrichment steps before adding calculations.
2. Use `train.csv` as the source of historical order counts. Enrich with meal and center dimensions only after checking unique keys.
3. Report the week range, row count, and total orders. Aggregate order totals by week, category, meal, or center as appropriate.
4. Distinguish total demand from per-record averages and account for differences in center counts and observed coverage when comparing center types.
5. For Pareto summaries, sort descending by orders and calculate cumulative share from the total orders; state the threshold and number of meals.
6. Keep missing center-meal-week observations missing. Do not infer zero sales from an absent row or call recorded orders the unconstrained customer demand.
7. Add a chart only when it answers the question; give it a finding-oriented title and save it under `reports/figures/`.
8. Report the numeric result and its source cell/table. State limits such as anonymized locations, week numbers without dates, and possible stock constraints.
9. Run the notebook from the first cell through the end and verify all output checks.

## Avoid

- Changing raw CSV files, inventing dates, or treating test rows as labeled demand.
- Inferring efficiency or causation from total orders alone.
- Calling estimated sales revenue profit.
