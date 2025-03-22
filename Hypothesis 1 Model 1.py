import pandas as pd
import statsmodels.api as sm
import os

# Load the dataset from Excel
data = pd.read_excel("C:/Users/cjfar/ECON 498H (Honors Thesis)/Second Draft/Honors Thesis Dataset 2.0.xlsx", sheet_name="Table")

# Clean column names by removing spaces, special characters, and hidden characters
data.columns = data.columns.str.replace(' ', '_').str.replace('\xa0', '').str.strip()

# Optionally, print the cleaned column names to verify
print("Cleaned column names in the dataset:")
print(data.columns.tolist())
print("\n" * 3)

# Define dependent and independent variables
X = data[[  
    'House_of_Representatives_DV',
    'Net_Exports_of_Goods_and_Services',
    'Average_Inflation_Rate',
    'Unemployment_Rate',
    '_Production_Volume:_Economic_Activity:_Industry_(Except_Construction)',
    'Production_Volume:_Economic_Activity:_Manufacturing',
    'Federal_Budget_DV',
    'S&P_500_Index_Return',
    'Dow_Jones_Index_Return',
    'Government_Consumption_Expenditures_and_Gross_Investment(GCEA)',
    'Gross_Private_Domestic_Investment_(GPDIA)',
    'Consumption_of_Fixed_Capital_(CFC)'
]]

# Define dependent variable
Y = data['Real_GDP_Growth_Rate']

# Add a constant (intercept) to the independent variables
X = sm.add_constant(X)

# Run the regression model
model = sm.OLS(Y, X).fit()

# Create a DataFrame with regression results
results_df = pd.DataFrame({
    "Variable": model.params.index,
    "Coefficient": model.params.round(3).values,
    "P-value": model.pvalues.round(3).values
})

# Create the Strength of Model Metric DataFrame
r_squared = round(model.rsquared, 3)
strength_df = pd.DataFrame({"Metric": ["R-Squared"], "Value": [r_squared], "Range": ["0 to 1"]})

# Create the Condition Number DataFrame
condition_number = round(model.condition_number, 3)
condition_df = pd.DataFrame({"Metric": ["Condition Number"], "Value": [condition_number]})

# Save results to a single Excel file with multiple sheets
output_path = "C:/Users/cjfar/ECON 498H (Honors Thesis)/Third Draft/Python Code 3.0/Hypothesis 1/Model 1"
os.makedirs(output_path, exist_ok=True)
output_file = os.path.join(output_path, "model_1_results.xlsx")

with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    results_df.to_excel(writer, sheet_name='Regression Results', index=False)
    strength_df.to_excel(writer, sheet_name='R-Squared', index=False)
    condition_df.to_excel(writer, sheet_name='Condition Number', index=False)

# Confirm that results were saved
print("All results have been saved to 'model_1_results.xlsx' in the specified directory.")
