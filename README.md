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

Run it directly: python Covariate Balancer v2.py

2. To Use Simulated Data:
X,D,Y = read_in_raw_data("file_name")

📘 What the Script Does
Data Preparation
simulate_data(): Simulates synthetic covariates, treatment, and outcome data.

read_in_raw_data(): Reads actual dataset with treatment and outcome.

Delta Tuning
get_delta_that_works(): Finds the smallest delta (starting from 0.05) that satisfies all covariate balance constraints.

Optimization
Uses SLSQP (Sequential Least Squares Programming) to find weights that balance covariates between treatment/control groups under linear constraints.

Validation
check_constraints(): Confirms constraints are satisfied using final weights.

Reporting
Exports summary tables to Excel

Plots weights, covariate balance, and treatment effects to PowerPoint

📊 Visual Outputs

Plot	Description
Weights Distribution	Visualizes individual patient weights
Mean Differences (Before/After)	Shows covariate balance improvements
Standardized Mean Differences	Key diagnostic for covariate balance
ATE/ATT/ATC Bar Chart	Summarizes estimated treatment effects
🧪 Treatment Effects Explained
ATE (Average Treatment Effect): Average effect if all were treated vs. none

ATT (on the Treated): Effect on treated group vs. if they were not treated

ATC (on the Controls): Effect on control group if they had been treated

🔧 Customization

🔢 Number of Covariates/Patients: Adjustable in simulate_data()

⚙️ Optimization Solver: Currently using SLSQP, structure allows adding others like trust-constr, PuLP, etc.

📄 Reporting Format: Exports results to .xlsx and .pptx, easy to customize

✅ Output Examples
After a successful run, you'll get:

Output_Consolidated.xlsx:

Constraint checks

Covariate summaries

Patient-level weights

Output_Consolidated.pptx:

Covariate balance visuals

Standardized Mean Differences (SMD)

Treatment effects (ATE, ATT, ATC)


