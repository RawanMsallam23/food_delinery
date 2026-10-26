---
name: promotion-impact-testing
description: "Use when evaluating, testing, comparing, or reporting the relationship between email promotions, homepage features, discounts, and food orders in this project; use a controlled experiment for causal impact."
argument-hint: "Specify promotion channel, outcome, and comparison period"
---

# Promotion Impact Testing

## When to use

- A user asks whether email or homepage promotion is associated with order volume.
- A user asks how to test a discount or promotion, estimate lift, compare channels, or report promotional results.

## Procedure

1. Read the promotion analysis in `notebooks/analysis.ipynb`. Keep email, homepage, and discount measures distinct.
2. For historical comparisons, compare promoted and non-promoted observations only within center-meal pairs that have both conditions.
3. Report group sample sizes, medians, relative median difference, Mann–Whitney two-sided p-value, and Holm-adjusted p-value for the three planned promotion tests.
4. For discount relationship, calculate the documented discount fraction `(base_price - checkout_price) / base_price` and Spearman's rho; do not call it causal lift.
5. Explain selection bias, concurrent campaigns, repeated observations, stock constraints, and the absence of costs/margins. Statistical significance does not imply business significance.
6. For causal evidence, recommend a randomized or defensible staggered test. Pre-specify the assignment unit, control/treatment, period, primary outcome, sample-size/power approach, and guardrails.
7. Measure incremental orders and contribution margin after campaign and ingredient costs; track stockouts, waste, and customer cancellations as guardrails.
8. Avoid contamination between treatment and control; consider center-level assignment when customer or operational spillover is plausible.
9. State that `test.csv` has no `num_orders` and cannot validate promotional outcomes until labels become available.
10. Re-run relevant notebook cells and check updated charts and outputs.

## Avoid

- Claiming promotions caused higher demand from these observational comparisons.
- Comparing unrelated meal-center mixtures without a matched or otherwise justified design.
- Treating a p-value below a threshold as proof of useful campaign profitability.
