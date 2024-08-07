import pandas as pd
import numpy as np
from scipy.optimize import curve_fit
import matplotlib.pyplot as plt

# Logistic function definition
def logistic_func(x, L, k, x0):
    return L / (1 + np.exp(-k * (x - x0)))

# Function to fit logistic curve and generate equation string
def fit_logistic_curve(x_data, y_data):
    popt, _ = curve_fit(logistic_func, x_data, y_data, maxfev=10000)
    L, k, x0 = popt
    equation = f"y = {L:.4f} / (1 + exp(-{k:.4f} * (x - {x0:.4f})))"
    return popt, equation

# Function to plot the fitted logistic curve
def plot_logistic_fit(x_data, y_data, popt, label):
    plt.scatter(x_data, y_data, label=f'Data (q = {label})')
    x_fit = np.linspace(min(x_data), max(x_data), 100)
    y_fit = logistic_func(x_fit, *popt)
    plt.plot(x_fit, y_fit, label=f'Logistic Fit (q = {label})')

# Main script
if __name__ == "__main__":
    # File path
    excel_file_path = 'sample.xlsx'  # Update this path to your actual Excel file

    # Step 1: Read the data
    df = pd.read_excel(excel_file_path)

    # Assuming the Excel file has columns 'x', 'R', 'q', 'lambda', and 'y'
    # Group the data by 'q' value
    curves = df.groupby('q')

    # Step 2-4: Fit logistic function and print equations
    for q_value, group in curves:
        x_data = group['x'].values
        y_data = group['y'].values
        try:
            popt, equation = fit_logistic_curve(x_data, y_data)
            print(f"Fitted parameters for q = {q_value}: {popt}")
            print(f"Equation for q = {q_value}: {equation}")
            plot_logistic_fit(x_data, y_data, popt, q_value)
        except Exception as e:
            print(f"Could not fit logistic function for q = {q_value}: {e}")

    # Step 5: Plot all curves with fitted logistic functions
    plt.xlabel('x')
    plt.ylabel('y')
    plt.title('Logistic Fit for Each q Value')
    plt.legend()
    plt.grid(True)
    plt.show()
