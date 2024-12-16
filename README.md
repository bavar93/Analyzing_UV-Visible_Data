# Analyzing_UV-Visible_Data

This project analyzes UV-Vis absorbance spectra data, calculates degradation percentages, kinetics, and generates plots for visual comparison. The results are saved to an Excel file.

## Features
- **Identify Maximum Absorbance Peaks**: The script identifies the maximum absorbance peaks for each condition and their corresponding wavelengths.
- **Calculate Degradation Percentages**: It calculates degradation percentages based on the initial conditions, focusing on changes from the dark amount after adsorption.
- **Determine Degradation Kinetics**: The script calculates degradation kinetics assuming a first-order reaction. It performs a linear regression to determine the rate constant and the quality of the fit (R-squared value).
- **Compute Area Under the Curve (AUC)**: It calculates the area under the curve for each condition to quantify the overall absorbance.
- **Generate Plots for Visual Comparison**: The script generates and saves UV-Vis absorbance spectra plots for visual comparison.

## Prerequisites
Make sure you have the following packages installed before running the script:
- pandas
- numpy
- openpyxl
- matplotlib
- scipy

## Usage

1. Place the UV-Vis absorbance spectra data in an Excel file (`Book2.xlsx`) and update the file path in the script if necessary.
2. Run the script:

    ```bash
    python project.py
    ```

3. The results, including degradation percentages, kinetics data, peak wavelengths, and AUC values, will be saved in `degradation_results.xlsx`.
4. The UV-Vis absorbance spectra plot will be saved as `spectra_comparison.png`.

## License
This project is licensed under the MIT License 





