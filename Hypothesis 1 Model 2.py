import pandas as pd
import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor
from tabulate import tabulate

# Load the dataset from Excel
data = pd.read_excel("C:/Users/cjfar/ECON 498H (Honors Thesis)/Second Draft/Honors Thesis Dataset 2.0.xlsx", sheet_name="Table")

# Clean column names by removing spaces, special characters, and hidden characters
data.columns = data.columns.str.replace(' ', '_').str.replace('\xa0', '').str.strip()

# Optionally, print the cleaned column names to verify
print("Cleaned column names in the dataset:")
print(data.columns.tolist())
print()
print()
print()

# Define dependent and independent variables
X = data[[  
    'House_of_Representatives_DV',
    'Average_Inflation_Rate',  # Adjusted to match cleaned column names
    'Unemployment_Rate',       # Adjusted to match cleaned column names
    '_Production_Volume:_Economic_Activity:_Industry_(Except_Construction)',  # Preserved initial underscore
    'Federal_Budget_DV',
    'S&P_500_Index_Return',
    'Government_Consumption_Expenditures_and_Gross_Investment(GCEA)',
]]

# Define dependent variable
Y = data['Real_GDP_Growth_Rate']

# Add a constant (intercept) to the independent variables
X = sm.add_constant(X)

# Run the regression model
model = sm.OLS(Y, X).fit()

# Prepare the regression results for tabulation
results = {
    "Variable": model.params.index,
    "Coefficient": model.params.round(3).values,  # Round coefficients to 3 decimal places
    "P-value": model.pvalues.round(3).values  # Round p-values to 3 decimal places
}

# Convert the results dictionary to a DataFrame for easier tabulation
results_df = pd.DataFrame(results)

# Prepare the R-squared value (rounded)
r_squared = round(model.rsquared, 3)

# Create the first table for the Strength of Model Metric
strength_of_model_table = [
    ["Metric", "Value", "Range"],
    ["R-Squared", r_squared, "0 to 1"]
]

# Create the second table for the regression results
regression_results_table = results_df.values.tolist()
regression_results_table.insert(0, ["Variable", "Coefficient", "P-value"])  # Add headers to regression results table

# Get the condition number from the model and round it
condition_number = round(model.condition_number, 3)

# Create a table for the condition number
condition_number_table = [
    ["Metric", "Value"],
    ["Condition Number", condition_number]
]

# Calculate VIF for each independent variable
vif_data = pd.DataFrame()
vif_data["Variable"] = X.columns
vif_data["VIF"] = [variance_inflation_factor(X.values, i) for i in range(X.shape[1])]

# Filter VIF values greater than 10 for easier interpretation
vif_data["VIF"] = vif_data["VIF"].round(3)
vif_results_table = vif_data.values.tolist()
vif_results_table.insert(0, ["Variable", "VIF"])  # Add headers to VIF results table

# Calculate the width of each table based on the longest row length including headers
strength_of_model_table_width = max(len(str(cell)) for row in strength_of_model_table for cell in row) + 4
regression_results_table_width = max(len(str(cell)) for row in regression_results_table for cell in row) + 4
condition_number_table_width = max(len(str(cell)) for row in condition_number_table for cell in row) + 4
vif_results_table_width = max(len(str(cell)) for row in vif_results_table for cell in row) + 4

print()
print()
print()
# Print the first table (Strength of Model Metric)
print(tabulate(strength_of_model_table, headers="firstrow", tablefmt="pretty"))
print()
print()
print()

# Print the second table (Calculated Coefficients)
print("\n")
print(tabulate(regression_results_table, headers="firstrow", tablefmt="pretty"))
print()
print()
print()

# Print the third table (Condition Number)
print("\n")
print(tabulate(condition_number_table, headers="firstrow", tablefmt="pretty"))
print()
print()
print()

# Print the fourth table (VIF Results)
print("\nVIF Results:")
print(tabulate(vif_results_table, headers="firstrow", tablefmt="pretty"))
print()
print()
print()
