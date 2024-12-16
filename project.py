import pandas as pd
import numpy as np
import openpyxl
import matplotlib
matplotlib.use('Agg')  # Use 'Agg' backend for non-interactive plotting
import matplotlib.pyplot as plt
from scipy.stats import linregress
from scipy.integrate import simpson


# Function to find max peaks and corresponding wavelengths
def find_max_peaks(df):
    peaks = {}
    peak_wavelengths = {}
    for column in df.columns[1:]:
        max_value = df[column].max()
        max_wavelength = df["Wavelength"][df[column].idxmax()]
        peaks[column] = max_value
        peak_wavelengths[column] = max_wavelength
    return peaks, peak_wavelengths


# Function to calculate degradation percentages (starting from dark amount)
def calculate_degradation_percentages(peaks):
    dark_peak = peaks["Dark"]  # Starting point: dark amount after adsorption
    degradation_percentages = {}
    for time_point, peak_value in peaks.items():
        if time_point not in ["Initial Solution", "Dark"]:
            degradation_percentages[time_point] = ((dark_peak - peak_value) / dark_peak) * 100
    return degradation_percentages


# Function to calculate degradation kinetics (assuming first-order reaction)
def calculate_kinetics(peaks, time_intervals):
    dark_peak = peaks["Dark"]  # Starting point: dark amount after adsorption
    remaining_peaks = [peaks[key] for key in peaks if key not in ["Initial Solution", "Dark"]]
    ln_remaining = np.log(remaining_peaks / dark_peak)

    # Linear regression for kinetics
    slope, intercept, r_value, p_value, std_err = linregress(time_intervals, ln_remaining)
    kinetics_data = pd.DataFrame({
        "Time (min)": time_intervals,
        "-ln(C/C0)": ln_remaining
    })
    return kinetics_data, slope, r_value ** 2  # Slope is the rate constant, r^2 is the fit quality


# Function to calculate Area Under the Curve (AUC)
def calculate_auc(df):
    auc_values = {}
    for column in df.columns[1:]:
        auc = simpson(df[column], x=df["Wavelength"])
        auc_values[column] = auc
    return auc_values


# Plot spectra for visual comparison
def plot_spectra(df):
    plt.figure(figsize=(10, 6))
    for column in df.columns[1:]:
        plt.plot(df["Wavelength"], df[column], label=column)
    plt.xlabel("Wavelength (nm)")
    plt.ylabel("Absorbance")
    plt.title("UV-Vis Absorbance Spectra")
    plt.legend()
    plt.grid()
    plt.savefig("spectra_comparison.png")  # Save the plot



# Main script
file_path = r"C:\Users\S. MR. B\PycharmProjects\pythonProject1\venv\Book2.xlsx"  # Corrected file path
df = pd.read_excel(file_path)

# Extract time intervals from column names (e.g., "Y min", "2Y min")
time_intervals = [int(col.split()[0].replace("Y", "").strip()) for col in df.columns[2:] if "min" in col]

# Find max peaks and wavelengths
peaks, peak_wavelengths = find_max_peaks(df)

# Calculate degradation percentages
degradation_percentages = calculate_degradation_percentages(peaks)

# Calculate degradation kinetics
kinetics_data, rate_constant, r_squared = calculate_kinetics(peaks, time_intervals)

# Calculate AUC
auc_values = calculate_auc(df)

# Plot spectra
plot_spectra(df)

# Save results to Excel
degradation_data = pd.DataFrame({
    "Condition": list(degradation_percentages.keys()),
    "Degradation (%)": list(degradation_percentages.values())
})

# Peak wavelength shifts
peak_shift_data = pd.DataFrame({
    "Condition": list(peak_wavelengths.keys()),
    "Max Absorbance (Peak)": list(peaks.values()),
    "Peak Wavelength (nm)": list(peak_wavelengths.values())
})

# AUC data
auc_data = pd.DataFrame({
    "Condition": list(auc_values.keys()),
    "Area Under Curve (AUC)": list(auc_values.values())
})

# Save all results to an Excel file
with pd.ExcelWriter("degradation_results.xlsx") as writer:
    degradation_data.to_excel(writer, sheet_name="Degradation Percentages", index=False)
    kinetics_data.to_excel(writer, sheet_name="Kinetics Data", index=False)
    peak_shift_data.to_excel(writer, sheet_name="Peak Wavelengths", index=False)
    auc_data.to_excel(writer, sheet_name="AUC Data", index=False)

# Save degradation rate and R^2 value
rate_summary = pd.DataFrame({
    "Rate Constant (k)": [rate_constant],
    "R-squared": [r_squared]
})
with pd.ExcelWriter("degradation_results.xlsx", mode="a") as writer:
    rate_summary.to_excel(writer, sheet_name="Kinetics Summary", index=False)

print("Results saved to 'degradation_results.xlsx'")

