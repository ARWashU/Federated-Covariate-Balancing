#!/usr/bin/env python
# coding: utf-8

# In[2]:


import numpy as np
import pandas as pd
import pulp
import matplotlib.pyplot as plt
import time
import io
import os
import sys
import warnings
from scipy.optimize import minimize
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill

"""Simulates dataset for a given number of patients and covariates."""
"""This is a substitute in the absence of an actual dataset"""
def simulate_data(num_patients, num_covariates, ATE=2, seed=42):

    global df  # simulated data stored in the global variable for easier manipulation
    
    np.random.seed(seed)
    
    # Generate a random set of covariates
    X = np.random.normal(0, 1, (num_patients, num_covariates))
    
    # Generate random propensity scores
    beta = np.random.normal(0, 1, num_covariates)
    ps = 1 / (1 + np.exp(-np.dot(X, beta)))
    
    # Assign patients to Treatment or Control groups
    D = np.random.binomial(1, p=ps)
    
    # Simulate outputs
    Y = ATE * D + np.dot(X, np.random.normal(0, 1, num_covariates)) + np.random.normal(0, 1, num_patients)

    # Sort the data by treatment assignment
    # All treatment group members come first, followed by all control group members
    # This is just a convenience step to make subsequent calculations simpler
    data = np.column_stack((X, D, Y))
    data = data[data[:, -2].argsort()[::-1]]
    columns = [f'covariate{i+1}' for i in range(X.shape[1])] + ['D', 'Y']

    # Store the data in a dataframe for subsequent manipulation
    df = pd.DataFrame(data, columns=columns)    

    # Re-extract X, D, Y after the sorting
    X = df.iloc[:, :-2].values
    D = df['D'].values
    Y = df['Y'].values
    
    return X, D, Y

"""Reads in dataset from a CSV file"""
def read_in_raw_data (file_name):

    global df # read-in data stored in the global variable for easier manipulation

    data = pd.read_csv(file_name)
    covariate_columns = [col for col in data.iloc[:, 1:].columns if col not in ['School_meal', 'BMI']]
    X = data[covariate_columns].values
    D = data['School_meal'].values
    Y = data['BMI'].values
    
    # Sort the data by treatment assignment
    # All treatment group members come first, followed by all control group members
    # This is just a convenience step to make subsequent calculations simpler
    data = np.column_stack((X, D, Y))
    data = data[data[:, -2].argsort()[::-1]]
    columns = [f'covariate{i+1}' for i in range(X.shape[1])] + ['D', 'Y']

    # Store the data in a dataframe for subsequent manipulation
    df = pd.DataFrame(data, columns=columns)

    # Re-extract X, D, Y after the sorting
    X = df.iloc[:, :-2].values
    D = df['D'].values
    Y = df['Y'].values
    
    return X, D, Y

"""Calculate and store commonly used data in global variables"""
def prepare_data (data):
    
    # prepare_data initializes the following global variables

    global covariates
    global num_patients, num_covariates
    global treatment_assignments
    global treatment_size, control_size
    global mean_treatment_covariates, mean_control_covariates

    X,_,_ = data
    num_patients, num_covariates = X.shape
    treatment_assignments = df['D'].astype(int)
    treatment_size = np.sum(treatment_assignments)
    control_size = num_patients - treatment_size
    
    covariates = df[[f'covariate{i+1}' for i in range(num_covariates)]]
    mean_treatment_covariates = covariates[treatment_assignments == 1].mean()
    mean_control_covariates = covariates[treatment_assignments == 0].mean()

    print("Total patients: ", num_patients, "Total covariates: ", num_covariates)
    print("Treatment group: ", treatment_size, "Control group: ", num_patients - treatment_size)
        

"""Define equality and inequality constraints for the PuLP solver."""
def define_constraints_pulp(delta):

    A_eq = np.zeros((2, num_patients))
    b_eq = np.array([1, 1])
    A_eq[0, :treatment_size] = 1
    A_eq[1, treatment_size:] = 1

    A_ub = []
    b_ub = []

    for i in range(num_covariates):
        covariate_values = covariates.iloc[:, i]
        mean_control_covariate = mean_control_covariates[i]

        A_ub.append(np.concatenate([covariate_values[:treatment_size], np.zeros(control_size)]))
        b_ub.append(delta + mean_control_covariate)
        A_ub.append(np.concatenate([-covariate_values[:treatment_size], np.zeros(control_size)]))
        b_ub.append(delta - mean_control_covariate)

        mean_treatment_covariate = mean_treatment_covariates[i]

        A_ub.append(np.concatenate([np.zeros(treatment_size), covariate_values[treatment_size:]]))
        b_ub.append(delta + mean_treatment_covariate)
        A_ub.append(np.concatenate([np.zeros(treatment_size), -covariate_values[treatment_size:]]))
        b_ub.append(delta - mean_treatment_covariate)

    return np.array(A_eq), b_eq, np.array(A_ub), np.array(b_ub), treatment_size, control_size

"""Checks feasibility of the constraints using PuLP."""
def check_feasibility_with_pulp(A_eq, b_eq, A_ub, b_ub):

    model = pulp.LpProblem("FeasibilityCheck", pulp.LpMinimize)
    weights = pulp.LpVariable.dicts("weights", range(num_patients), lowBound=1e-10, upBound=1-1e-10, cat='Continuous')

    model += pulp.lpSum([weights[i] for i in range(treatment_size)]) == 1
    model += pulp.lpSum([weights[i] for i in range(treatment_size, num_patients)]) == 1

    for i in range(len(A_ub)):
        model += pulp.lpSum([A_ub[i, j] * weights[j] for j in range(num_patients)]) <= b_ub[i]

    model += 0
    model.solve(pulp.PULP_CBC_CMD(msg=True))

    if pulp.LpStatus[model.status] == 'Optimal':
        feasible_weights = np.array([weights[i].varValue for i in range(num_patients)])
        return True, feasible_weights
    else:
        return False, None

"""Define objective function, equality and inequality constraints for the SLSQP and Trust-constr solvers."""
def objective(weights):
    return np.var(weights)

def objective_grad(weights):
    return np.ones(len(weights))

def objective_hess(weights):
    return np.zeros((len(weights), len(weights)))

def constraint_sum_weights_treatment(weights):
    return np.sum(weights[:treatment_size]) - 1

def constraint_sum_weights_treatment_grad(weights):
    grad = np.zeros(len(weights))
    grad[:treatment_size] = 1
    return grad

def constraint_sum_weights_control(weights):
    return np.sum(weights[treatment_size:]) - 1

def constraint_sum_weights_control_grad(weights):
    grad = np.zeros(len(weights))
    grad[treatment_size:] = 1
    return grad

def constraint_cov_diff_treatment(weights, covariate_idx, direction):
    weighted_mean_treatment_cov = np.sum(weights[:treatment_size] * covariates.iloc[:treatment_size, covariate_idx])
    if direction == 'upper':
        return delta - (weighted_mean_treatment_cov - mean_control_covariates[covariate_idx])
    else:
        return delta + (weighted_mean_treatment_cov - mean_control_covariates[covariate_idx])

def constraint_cov_diff_control(weights, covariate_idx, direction):
    weighted_mean_control_cov = np.sum(weights[treatment_size:] * covariates.iloc[treatment_size:, covariate_idx])
    if direction == 'upper':
        return delta - (weighted_mean_control_cov - mean_treatment_covariates[covariate_idx])
    else:
        return delta + (weighted_mean_control_cov - mean_treatment_covariates[covariate_idx])

def constraint_cov_diff_treatment_grad(weights, covariate_idx, direction):
    grad = np.zeros(len(weights))
    grad[:treatment_size] = -covariates.iloc[:treatment_size, covariate_idx] if direction == 'upper' else covariates.iloc[:treatment_size, covariate_idx]
    return grad

def constraint_cov_diff_control_grad(weights, covariate_idx, direction):
    grad = np.zeros(len(weights))
    grad[treatment_size:] = -covariates.iloc[treatment_size:, covariate_idx] if direction == 'upper' else covariates.iloc[treatment_size:, covariate_idx]
    return grad

def define_constraints_slsqp_trust ():
   
    constraints = [
        {'type': 'eq', 'fun': constraint_sum_weights_treatment, 'jac': constraint_sum_weights_treatment_grad},
        {'type': 'eq', 'fun': constraint_sum_weights_control, 'jac': constraint_sum_weights_control_grad}
    ]

    for i in range(num_covariates):
        constraints.append({'type': 'ineq', 'fun': constraint_cov_diff_treatment, 'args': (i, 'upper'), 'jac': constraint_cov_diff_treatment_grad})
        constraints.append({'type': 'ineq', 'fun': constraint_cov_diff_treatment, 'args': (i, 'lower'), 'jac': constraint_cov_diff_treatment_grad})
        constraints.append({'type': 'ineq', 'fun': constraint_cov_diff_control, 'args': (i, 'upper'), 'jac': constraint_cov_diff_control_grad})
        constraints.append({'type': 'ineq', 'fun': constraint_cov_diff_control, 'args': (i, 'lower'), 'jac': constraint_cov_diff_control_grad})

    return constraints


"""Solves the optimization problem using SLSQP."""
# Define callback function
iteration = [0]  # To keep track of iterations

def callback(weights):
    iteration[0] += 1
    print(f"\rIteration {iteration[0]}", end='')
    sys.stdout.flush()

def run_slsqp_optimization(initial_weights, options, bounds, constraints):
    start_time = time.time()
    result = minimize(
        fun=objective,
        constraints=constraints,
        x0=initial_weights,
        method='SLSQP',
        jac=objective_grad,
        bounds=bounds,
        callback = callback,
        options=options
    )
    end_time = time.time()
    return result, end_time - start_time

    
'''A class to capture print statements'''
class PrintCapture(io.StringIO):
    def __init__(self):
        super().__init__()

    def __enter__(self):
        self._stdout = sys.stdout
        sys.stdout = self
        return self

    def __exit__(self, *args):
        sys.stdout = self._stdout

'''A function to write captured print statements to an Excel file'''

# Tracker for output file
file_erased_xl = False

def write_to_excel(content, filename, sheet_name='Constraint Check', method=''):
    global file_erased_xl

    # Add ".pptx" to the filename if it doesn't already end with it
    if not filename.endswith(".xlsx"):
        filename += ".xlsx"
    
    # Erase the file if it exists and this is the first call in the current execution
    if os.path.exists(filename) and not file_erased_xl:
        os.remove(filename)
        file_erased_xl = True

    # Create a new Excel file or load the existing one
    if os.path.exists(filename):
        writer = pd.ExcelWriter(filename, mode='a', engine='openpyxl', if_sheet_exists='replace')
    else:
        writer =  pd.ExcelWriter(filename, mode='w', engine='openpyxl')
        file_erased_xl = True

    with writer as writer:
        # Split the content into lines
        lines = content.split('\n')
        # Remove empty lines
        lines = [line for line in lines if line.strip() != '']
        # Create a DataFrame from the lines
        df = pd.DataFrame(lines, columns=['Output'])
        df.to_excel(writer, sheet_name=f'{sheet_name}_{method}', index=False)
        # Apply color to the sheet tab based on the method
        workbook = writer.book
        sheet = workbook[f'{sheet_name}_{method}']
        tab_color = 'FF0000' if method == 'PuLP' else '0000FF' if method == 'SLSQP' else '00FF00'
        sheet.sheet_properties.tabColor = tab_color

"""Checks if the optimized weights satisfy the constraints and writes the results to an Excel file."""      
def check_constraints(weights, filename, method):

    if weights is None:
        print("No weights available. Optimization failed.")
        return

    with PrintCapture() as pc:
        print(f"Sum of weights in treatment group: {np.sum(weights[:treatment_size])}")
        print(f"Sum of weights in control group: {np.sum(weights[treatment_size:])}")
        print(f"Objective function value: {np.sum(weights)}")

        for i in range(num_covariates):
            weighted_mean_treatment_cov = np.sum(weights[:treatment_size] * covariates.iloc[:treatment_size, i])
            weighted_mean_control_cov = np.sum(weights[treatment_size:] * covariates.iloc[treatment_size:, i])
            print(f"Constraint check for covariate {i+1} (treatment): {np.abs(weighted_mean_treatment_cov - mean_control_covariates[i]) <= delta} (", np.abs(weighted_mean_treatment_cov - mean_control_covariates[i]), "<=", delta, ")")
            print(f"Constraint check for covariate {i+1} (control): {np.abs(weighted_mean_control_cov - mean_treatment_covariates[i]) <= delta} (", np.abs(weighted_mean_control_cov - mean_treatment_covariates[i]), "<=", delta, ")")

    # Add ".xlsx" to the filename if it doesn't already end with it
    if not filename.endswith(".xlsx"):
        filename += ".xlsx"
            
    # Write the captured print statements to the specified Excel file
    write_to_excel(pc.getvalue(), filename, method=method)

'''A function to create slides in a powerpoint file of visualizations'''
# Tracker for output file
file_erased_ppt = False

def save_plot_to_ppt(plt, heading, filename, method):
    global file_erased_ppt
    
    # Add ".pptx" to the filename if it doesn't already end with it
    if not filename.endswith(".pptx"):
        filename += ".pptx"
    
    # Erase the file if it exists and this is the first call in the current execution
    if os.path.exists(filename) and not file_erased_ppt:
        os.remove(filename)
        file_erased_ppt = True

    # Create a new presentation or load the existing one
    if os.path.exists(filename):
        prs = Presentation(filename)
    else:
        prs = Presentation()
        file_erased_ppt = True
    
    # Save the plot to a BytesIO object
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png')
    img_stream.seek(0)  # Rewind the stream to the beginning
    
    # Add a slide with a title and content layout
    slide_layout = prs.slide_layouts[5]  # Choose a layout (5 is title and content)
    slide = prs.slides.add_slide(slide_layout)
    title_shape = slide.shapes.title
    title_shape.text = f"{heading} ({method})"

    # Format the title to fit in one line
    for paragraph in title_shape.text_frame.paragraphs:
        paragraph.font.size = Pt(24)
        paragraph.alignment = PP_ALIGN.CENTER
    
    # Calculate the space available for the plot
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    top_margin = Inches(1.5)  # Space for the title
    left_margin = Inches(0.5)
    right_margin = Inches(0.5)
    bottom_margin = Inches(0.5)
    
    img_width = slide_width - left_margin - right_margin
    img_height = slide_height - top_margin - bottom_margin
    
    # Add the plot image to the slide
    slide.shapes.add_picture(img_stream, left_margin, top_margin, width=img_width, height=img_height)      
    prs.save(filename)

"""Main function to plot the graphs and save them to Excel and PowerPoint."""
def plot_graph(weights, file, method):
    mean_diff_before_after = calculate_mean_differences(weights)
    save_covariate_summary_to_excel(mean_diff_before_after, file, method)
    save_patient_summary_to_excel(weights, file, method)
    plot_and_save_weights_distribution(weights, file, method)
    plot_and_save_mean_diff_before_after(mean_diff_before_after, delta, file, method)

    # Calculate and plot standardized mean differences
    smd_before, smd_after = calculate_standardized_mean_differences(weights)
    plot_standardized_mean_differences(smd_before, smd_after, file, method)

"""Calculate mean differences before and after optimization."""
def calculate_mean_differences(weights):
    weighted_mean_treatment_covariates = np.dot(weights[:treatment_size], covariates.iloc[:treatment_size])
    weighted_mean_control_covariates = np.dot(weights[treatment_size:], covariates.iloc[treatment_size:])
    mean_diff_before_balancing_treatment = mean_treatment_covariates - mean_control_covariates
    mean_diff_before_balancing_control = mean_control_covariates - mean_treatment_covariates
    mean_diff_after_balancing_treatment = weighted_mean_treatment_covariates - mean_control_covariates
    mean_diff_after_balancing_control = weighted_mean_control_covariates - mean_treatment_covariates
    
    return {
        'weighted_treatment': weighted_mean_treatment_covariates,
        'weighted_control': weighted_mean_control_covariates,        
        'before_treatment': mean_diff_before_balancing_treatment,
        'before_control': mean_diff_before_balancing_control,
        'after_treatment': mean_diff_after_balancing_treatment,
        'after_control': mean_diff_after_balancing_control
    }

"""Calculate standardized mean differences before and after optimization."""
def calculate_standardized_mean_differences(weights):
    weighted_mean_treatment_covariates = np.dot(weights[:treatment_size], covariates.iloc[:treatment_size])
    weighted_mean_control_covariates = np.dot(weights[treatment_size:], covariates.iloc[treatment_size:])
    
    total = len(weights)
    ts = treatment_size
    cs = total - ts
    
    pooled_sd_before = np.sqrt((covariates[treatment_assignments == 1].var()*(ts-1) + covariates[treatment_assignments == 0].var()*(cs-1))/(ts+cs-2))
    squared_diffs_treatment = np.array([(covariates.iloc[:ts,j] - weighted_mean_treatment_covariates[j])**2 for j in range(num_covariates)])
    squared_diffs_control = np.array([(covariates.iloc[ts:,j] - weighted_mean_control_covariates[j])**2 for j in range(num_covariates)])
    pooled_sd_after = np.sqrt(np.average(squared_diffs_treatment,axis=1, weights= weights[:ts])+np.average(squared_diffs_control,axis=1, weights= weights[ts:]))
    
    smd_before_balancing = [np.abs(mean_treatment_covariates[j] - mean_control_covariates[j]) / pooled_sd_before[j]for j in range(num_covariates)]
    smd_after_balancing = [np.abs(weighted_mean_control_covariates[j] - mean_treatment_covariates[j]) / pooled_sd_after[j]for j in range(num_covariates)]
    
    return smd_before_balancing, smd_after_balancing

"""Plot standardized mean differences before and after optimization."""
def plot_standardized_mean_differences(smd_before, smd_after, file, method):
    covariates = [f'covariate{i+1}' for i in range(num_covariates)]
    y_positions = np.arange(len(covariates))  # Positions for the covariates on the y-axis

    plt.figure(figsize=(10, 8))

    # Plot SMD before balancing
    plt.plot(smd_before, y_positions, 'o-', color='blue', label='SMD Before', alpha=0.6)

    # Plot SMD after balancing
    plt.plot(smd_after, y_positions, 'o-', color='red', label='SMD After', alpha=0.6)

    # Formatting
    plt.yticks(y_positions, covariates)
    plt.xlabel('Standardized Mean Difference')
    plt.ylabel('Covariates')
    plt.title('Standardized Mean Differences Before vs After Weighting')
    plt.axvline(x=0, color='black', linestyle='--', linewidth=1)
    plt.grid(True, axis='x')
    plt.legend()

    save_plot_to_ppt(plt, "Standardized Mean Differences Before vs After Weighting", file, method)
    plt.close()

"""Save covariate summary to Excel."""
def save_covariate_summary_to_excel(mean_diff_before_after, file, method):
    selected_covariates = list(range(num_covariates))
    covariate_table_data = {
        'Mean in Treatment Group Before Balancing': mean_treatment_covariates.values[selected_covariates],
        'Weighted Mean in Treatment Group After Balancing': mean_diff_before_after['weighted_treatment'][selected_covariates],
        'Mean in Control Group Before Balancing': mean_control_covariates.values[selected_covariates],
        'Weighted Mean in Control Group After Balancing': mean_diff_before_after['weighted_control'][selected_covariates],
        'Abs Diff (Treatment Before - Control Before)': np.abs(mean_diff_before_after['before_treatment'][selected_covariates]),
        'Abs Diff (Weighted Treatment - Control)': np.abs(mean_diff_before_after['after_treatment'][selected_covariates]),
        'Abs Diff (Weighted Control - Treatment)': np.abs(mean_diff_before_after['after_control'][selected_covariates])
    }
    covariate_table_df = pd.DataFrame(covariate_table_data, index=[f'covariate{i+1}' for i in selected_covariates]).transpose()
    
    write_excel(file, f'Covariate Summary_{method}', covariate_table_df, method)

"""Save patient summary to Excel."""
def save_patient_summary_to_excel(weights, file, method):
    treatment_or_control = ['Treatment' if i == 1 else 'Control' for i in df['D']]
    patient_table_data = {
        'In Treatment or Control': treatment_or_control,
        'Weight': weights,
        'Covariate1': df.iloc[:, 0],
        'Covariate2': df.iloc[:, 1],
        **{f'Covariate{i+1}': df.iloc[:, i] for i in range(num_covariates)}
    }
    patient_table_df = pd.DataFrame(patient_table_data)
    
    write_excel(file, f'Patient Summary_{method}', patient_table_df, method)

"""Write DataFrame to Excel with specified formatting."""
def write_excel(file, sheet_name, df, method):
    if not file.endswith(".xlsx"):
        file += ".xlsx"
    with pd.ExcelWriter(file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name)
        workbook = writer.book
        sheet = workbook[sheet_name]
        tab_color = 'FF0000' if method == 'PuLP' else '0000FF' if method == 'SLSQP' else '00FF00'
        sheet.sheet_properties.tabColor = tab_color
        adjust_column_widths(sheet)

"""Adjust column widths for better readability in Excel."""
def adjust_column_widths(sheet):
    for column_cells in sheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = length

"""Plot and save the distribution of weights."""
def plot_and_save_weights_distribution(weights, file, method):

    plt.figure(figsize=(12, 6))
    indices = np.arange(len(weights))
    plt.scatter(indices[:treatment_size], weights[:treatment_size], label='Treatment Group', color='b', alpha=0.6)
    plt.scatter(indices[treatment_size:], weights[treatment_size:], label='Control Group', color='r', alpha=0.6)
    plt.xlabel('Patient Index')
    plt.ylabel('Weight')
    plt.title('Weights Assigned to Each Patient')
    plt.legend()
    plt.tight_layout()
    plt.xlim(0, len(weights) - 1)
    
    if not file.endswith(".pptx"):
        file += ".pptx"
    
    save_plot_to_ppt(plt, "Distribution of Weights", file, method)
    plt.close()

"""Plot and save mean differences before and after optimization."""
def plot_and_save_mean_diff_before_after(mean_diff_before_after, delta, file, method):

    plt.figure(figsize=(12, 12))
    plt.scatter(mean_diff_before_after['before_treatment'], mean_diff_before_after['after_treatment'], color='blue', edgecolor='blue', label='Treatment Group')
    plt.scatter(mean_diff_before_after['before_control'], mean_diff_before_after['after_control'], color='red', edgecolor='red', label='Control Group')
    plt.xlabel('Mean Difference Before Optimization')
    plt.ylabel('Mean Difference After Optimization')
    plt.title('Mean Difference Before vs After Optimization')
    plt.legend()
    plt.xlim(-1, 1)
    plt.ylim(-2 * delta, 2 * delta)
    plt.grid(False)
    plt.axhline(y=-delta, color='black', linestyle='-', linewidth=2)
    plt.axhline(y=delta, color='black', linestyle='-', linewidth=2)
    
    save_plot_to_ppt(plt, "Impact of covariate balancing using constrained optimization", file, method)
    plt.close()

'''Plot ATE, ATT, and ATC'''
def plot_treatment_effects(ATE, ATT, ATC, file, method):
    effects = ['ATE', 'ATT', 'ATC']
    values = [ATE, ATT, ATC]

    plt.figure(figsize=(10, 6))
    plt.bar(effects, values, color=['blue', 'green', 'red','yellow'])
    plt.xlabel('Effect Type')
    plt.ylabel('Value')
    plt.title('Treatment Effects: ATE, ATT, and ATC')
    save_plot_to_ppt(plt, "Treatment Effects: ATE, ATT, and ATC", file, method)
    plt.close()
    
'''For the dataset, find a delta that works. Use increments of 0.05 till feasibility is established.'''
def get_delta_that_works ():
    
    # Start with delta = 0.05
    delta = 0.05
    
    #Iterate up to 10 times to find a delta that works
    max_iterations = 10
    
    feasible = False
    iteration = 0

    while not feasible and iteration < max_iterations:
        print(f"\nChecking feasibility with delta = {delta:.2f}. ", end='')
        A_eq, b_eq, A_ub, b_ub, treatment_size, control_size = define_constraints_pulp(delta)
        num_patients = treatment_size + control_size
        feasible, feasible_weights = check_feasibility_with_pulp(A_eq, b_eq, A_ub, b_ub)
    
        if not feasible:
            print("Not feasible")
            delta += 0.05
            iteration += 1

    print(f"\n{'Feasible' if feasible else 'Not Feasible'} with delta = {delta:.2f}")
    
    if (not feasible):
        print ("Max iterations to find a feasible delta reached but no feasible delta found.")
        return 0
    else:
        return delta

'''Calculate ATT, ATC, and ATE'''
def treatment_effect_pooled(data, weights, outfile = "Output_Consolidated",method="SLSQP"):
    
    X, D, Y = data
    # Calculate mean treatment and control outputs
    mean_treatment_output = np.sum(Y[:treatment_size]) / treatment_size
    mean_control_output = np.sum(Y[treatment_size:]) / control_size
    
    # Calculate weighted mean treatment output
    treatment_weights = weights[:treatment_size]
    weighted_mean_treatment_output = np.sum(Y[:treatment_size] * treatment_weights) / np.sum(treatment_weights)
    
    # Calculate weighted mean control output
    control_weights = weights[treatment_size:]
    weighted_mean_control_output = np.sum(Y[treatment_size:] * control_weights) / np.sum(control_weights)
    
    # Calculate ATT
    # ATT = Average Treatment Effect on the Treated
    # ATT = Mean output of treatment group - Output of treatment group as if they had not been treated
    ATT = mean_treatment_output - weighted_mean_control_output

    # Calculate ATC
    # ATT = Average Treatment Effect on the Controlled
    # ATT = Output of control group as if they had been treated - Mean output of control group
    ATC = weighted_mean_treatment_output - mean_control_output 

    # Calculate ATE
    ATE = ATT * (treatment_size / num_patients) + ATC * (control_size / num_patients)
    
    # Call the plot function
    plot_treatment_effects(ATE, ATT, ATC, outfile, method)


'''For the dataset, run the constrained optimization and generate balancing weights.'''
def covariate_balancing_pooled (data, outfile = "Output_Consolidated",method="SLSQP"):
    
    global delta
    
    result = False
 
    prepare_data (data)

    delta = get_delta_that_works()
    if (delta ==0):
        print ("Exiting.")
        sys.exit();

    # SLSQP Optimization
    print(f"\nRunning {method} Optimization...")
    
    initial_weights = np.ones(num_patients) / num_patients

    options = {'maxiter': 1000}
    bounds = [(1e-10, 1-1e-10) for _ in range(num_patients)]
    constraints = define_constraints_slsqp_trust ()

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        result_slsqp, slsqp_time = run_slsqp_optimization(initial_weights, options, bounds, constraints)

    results_summary = {
        'SLSQP': {'time': slsqp_time, 'success': result_slsqp.success, 'message': result_slsqp.message}
    }

    optimized_weights_slsqp = result_slsqp.x

    if result_slsqp.success:
        print("\nSLSQP Optimization successful!")
        print (results_summary)
        
        print(f"Sum of weights in treatment group: {np.sum(optimized_weights_slsqp[:treatment_size])}")
        print(f"Sum of weights in control group: {np.sum(optimized_weights_slsqp[treatment_size:])}")
        print(f"Objective function value: {np.sum(optimized_weights_slsqp)}")
        
        print(f"Preparing output to {outfile}...", end='')
        check_constraints(optimized_weights_slsqp, outfile, method)   
        plot_graph(optimized_weights_slsqp, outfile, method)       
        result = True
    else:
        print(f"\n{method} Optimization failed.")
        
    return optimized_weights_slsqp, result
    

'''Main execution'''

#X, D, Y = read_in_raw_data("nhanes_bmi.csv")
X, D, Y = simulate_data (2330,11)
data = [X, D, Y]

optimized_weights, result = covariate_balancing_pooled (data)

if (result == True):
    treatment_effect_pooled(data, optimized_weights)    
    print ("Done.\nOutputs written.")

else:
    print ("Balancing did not work.\nExiting.")
    sys.exit();





