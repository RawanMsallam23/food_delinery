---
description: "Use when changing this project's data analysis notebook, metrics, statistical comparisons, charts, cleaning decisions, or client-facing conclusions."
applyTo: "notebooks/**/*.ipynb"
---
# Data Analysis Rules

- Preserve source CSVs and report every merge's row count and order-total checks.
- Treat a missing center-meal-week observation as unknown, not zero.
- Keep observed outliers unless a documented data-quality reason supports exclusion.
- Label revenue as an estimate; it is not profit and excludes costs.
- Use matched center-meal comparisons for promotion analysis and report Holm-adjusted p-values for multiple tests.
- Describe observational differences as associations, never causal effects.
- Re-run all notebook cells and verify generated outputs after edits.
