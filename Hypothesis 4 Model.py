import pandas as pd
import statsmodels.api as sm
from sklearn.preprocessing import StandardScaler
from statsmodels.stats.outliers_influence import variance_inflation_factor

# Load the dataset from Excel
data = pd.read_excel("C:/Users/cjfar/ECON 498H (Honors Thesis)/Second Draft/Honors Thesis Dataset 2.0.xlsx", sheet_name="Table")

# Clean column names
data.columns = data.columns.str.replace(' ', '_').str.replace('\xa0', '').str.strip()
data.rename(columns=lambda x: x.lstrip('_'), inplace=True)

# Define independent and dependent variables
X = data[[  
    'Government_Type_DV',
    'Average_Inflation_Rate',
    'Unemployment_Rate',
    'Production_Volume:_Economic_Activity:_Industry_(Except_Construction)',
    'Federal_Budget_DV',
    'S&P_500_Index_Return',
    'Government_Consumption_Expenditures_and_Gross_Investment(GCEA)',
]]
Y = data['Real_GDP_Growth_Rate']

# Add a constant to independent variables
X = sm.add_constant(X)

# Standardize independent variables (excluding constant)
scaler = StandardScaler()
X_scaled_values = scaler.fit_transform(X.iloc[:, 1:])  # Exclude constant
X_scaled_df = pd.DataFrame(X_scaled_values, columns=X.columns[1:])
X_scaled_df = pd.concat([X[['const']].reset_index(drop=True), X_scaled_df], axis=1)

# Run regression model
model_scaled = sm.OLS(Y, X_scaled_df).fit()

# Create DataFrame for regression results
results_df = pd.DataFrame({
    "Variable": model_scaled.params.index,
    "Coefficient": model_scaled.params.round(3).values,
    "Standard Error": model_scaled.bse.round(3).values,
    "P-value": model_scaled.pvalues.round(3).values
})

# Create DataFrame for model strength metrics
metrics_df = pd.DataFrame({
    "Metric": ["R-Squared", "Condition Number"],
    "Value": [round(model_scaled.rsquared, 3), round(model_scaled.condition_number, 3)]
})

# Calculate VIF
vif_data = pd.DataFrame({
    "Variable": X_scaled_df.columns,
    "VIF": [variance_inflation_factor(X_scaled_df.values, i) for i in range(X_scaled_df.shape[1])]
})
vif_data["VIF"] = vif_data["VIF"].round(3)

# Export all data to a single Excel file
output_path = "C:/Users/cjfar/ECON 498H (Honors Thesis)/Third Draft/Python Code 3.0/Hypothesis 4/Regression_Results.xlsx"
with pd.ExcelWriter(output_path) as writer:
    results_df.to_excel(writer, sheet_name="Regression Results", index=False)
    metrics_df.to_excel(writer, sheet_name="Model Strength Metrics", index=False)
    vif_data.to_excel(writer, sheet_name="VIF Results", index=False)

