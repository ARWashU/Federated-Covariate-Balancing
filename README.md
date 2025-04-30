# Covariate Balancing via Constrained Optimization

This repository provides a comprehensive Python-based pipeline for **covariate balancing using constrained optimization**, with visualization and reporting integrated into Excel and PowerPoint formats. The tool is ideal for estimating treatment effects like ATE, ATT, and ATC in observational studies while ensuring covariate balance between treatment and control groups.

---

## 📌 Features

- Simulate or read datasets with treatment assignment
- Automatically find feasible constraint parameters (`δ`)
- Perform constrained optimization (SLSQP) to generate balancing weights
- Validate and visualize covariate balance pre- and post-optimization
- Calculate treatment effects: ATE, ATT, ATC
- Export results to:
  - Excel (`.xlsx`) — tabular summaries
  - PowerPoint (`.pptx`) — visualizations

## 🚀 How to Use
1. Simulate Data (Default)
The script simulates a dataset of patients with covariates, treatment assignment, and outcome.

Run it directly: python Covariate Balancer v2.p


