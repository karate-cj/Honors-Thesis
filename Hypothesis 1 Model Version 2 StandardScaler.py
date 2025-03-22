import pandas as pd
import statsmodels.api as sm
from sklearn.preprocessing import StandardScaler
from statsmodels.stats.outliers_influence import variance_inflation_factor

# Load the dataset from Excel
data = pd.read_excel("C:/Users/cjfar/ECON 498H (Honors Thesis)/Second Draft/Honors Thesis Dataset 2.0.xlsx", sheet_name="Table")

# Clean column names by removing spaces, special characters, and hidden characters
data.columns = data.columns.str.replace(' ', '_').str.replace('\xa0', '').str.strip()

# Remove leading underscores in column names for cleaner output
data.rename(columns=lambda x: x.lstrip('_'), inplace=True)

# Define dependent and independent variables
X = data[[  
    'House_of_Representatives_DV',
    'Average_Inflation_Rate',
    'Unemployment_Rate',
    'Production_Volume:_Economic_Activity:_Industry_(Except_Construction)',
    'Federal_Budget_DV',
    'S&P_500_Index_Return',
    'Government_Consumption_Expenditures_and_Gross_Investment(GCEA)',
]]

# Define dependent variable
Y = data['Real_GDP_Growth_Rate']

# Add a constant (intercept) to the independent variables
X = sm.add_constant(X)

# Standardize independent variables (excluding the constant term)
scaler = StandardScaler()
X_scaled_values = scaler.fit_transform(X.iloc[:, 1:])  # Exclude constant term
X_scaled_df = pd.DataFrame(X_scaled_values, columns=X.columns[1:])  # Keep original column names
X_scaled_df = pd.concat([X[['const']].reset_index(drop=True), X_scaled_df], axis=1)  # Add constant back

# Run the regression model with standardized variables
model_scaled = sm.OLS(Y, X_scaled_df).fit()

# Prepare the regression results including standard errors
results_df = pd.DataFrame({
    "Variable": model_scaled.params.index,
    "Coefficient": model_scaled.params.round(3).values,
    "Standard Error": model_scaled.bse.round(3).values,
    "P-value": model_scaled.pvalues.round(3).values
})

# Prepare model statistics
r_squared = round(model_scaled.rsquared, 3)
condition_number = round(model_scaled.condition_number, 3)

# Calculate VIF for each independent variable (standardized variables)
vif_data = pd.DataFrame()
vif_data["Variable"] = X_scaled_df.columns
vif_data["VIF"] = [variance_inflation_factor(X_scaled_df.values, i) for i in range(X_scaled_df.shape[1])]
vif_data["VIF"] = vif_data["VIF"].round(3)

# Create a Pandas Excel writer
output_path = "C:/Users/cjfar/ECON 498H (Honors Thesis)/Third Draft/Python Code 3.0/Hypothesis 1/Model 2/Regression_Results.xlsx"
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Write the regression results to the "Regression Results" sheet
    results_df.to_excel(writer, sheet_name="Regression Results", index=False)
    
    # Write the VIF results to the "VIF Results" sheet
    vif_data.to_excel(writer, sheet_name="VIF Results", index=False)
    
    # Create a dataframe for the R-squared and Condition Number and add them to the same sheet
    model_strength_df = pd.DataFrame({
        "Metric": ["R-Squared", "Condition Number"],
        "Value": [r_squared, condition_number]
    })
    model_strength_df.to_excel(writer, sheet_name="Regression Results", index=False, startrow=len(results_df) + 2)

print(f"Results have been successfully saved to {output_path}")
