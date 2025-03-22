import pandas as pd
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.tools.tools import add_constant

# Load your dataset
data = pd.read_excel("C:/Users/cjfar/ECON 498H (Honors Thesis)/Second Draft/Honors Thesis Dataset 2.0.xlsx", sheet_name="Table")

# Clean column names
data.columns = data.columns.str.replace(r'\s+', '_', regex=True).str.replace('\xa0', '').str.strip()
data.columns = data.columns.str.lstrip('_')

# Define selected columns
selected_columns = [
    'House_of_Representatives_DV', 'Net_Exports_of_Goods_and_Services', 'Average_Inflation_Rate',
    'Unemployment_Rate', 'Production_Volume:_Economic_Activity:_Industry_(Except_Construction)',
    'Production_Volume:_Economic_Activity:_Manufacturing', 'Federal_Budget_DV', 'S&P_500_Index_Return',
    'Dow_Jones_Index_Return', 'Government_Consumption_Expenditures_and_Gross_Investment_(GCEA)',
    'Gross_Private_Domestic_Investment_(GPDIA)', 'Consumption_of_Fixed_Capital_(CFC)', 'Real_GDP_Growth_Rate'
]

# Check for missing columns
missing_columns = [col for col in selected_columns if col not in data.columns]
if missing_columns:
    print("Missing columns:", missing_columns)
else:
    filtered_data = data[selected_columns]
    filtered_data_with_constant = add_constant(filtered_data)

    # Calculate VIF
    vif_data = pd.DataFrame()
    vif_data['Variable'] = filtered_data_with_constant.columns
    vif_data['VIF'] = [variance_inflation_factor(filtered_data_with_constant.values, i) 
                       for i in range(filtered_data_with_constant.shape[1])]

    # Filter high VIF
    high_vif = vif_data[vif_data['VIF'] > 10].sort_values(by='VIF', ascending=False)
    high_vif_vars = [var for var in high_vif['Variable'].tolist() if var != 'const']

    # Correlation matrix for high VIF variables
    correlation_matrix = filtered_data[high_vif_vars].corr()
    correlation_threshold = 0.8
    high_correlation_pairs = [
        (correlation_matrix.columns[i], correlation_matrix.columns[j], round(correlation_matrix.iloc[i, j], 3))
        for i in range(len(correlation_matrix.columns))
        for j in range(i) if abs(correlation_matrix.iloc[i, j]) > correlation_threshold
    ]

    # Exclude specific variables
    excluded_variables = [
        'Production_Volume:_Economic_Activity:_Manufacturing', 'Consumption_of_Fixed_Capital_(CFC)',
        'Gross_Private_Domestic_Investment_(GPDIA)', 'Net_Exports_of_Goods_and_Services', 'Dow_Jones_Index_Return'
    ]
    filtered_data_excluded = filtered_data.drop(columns=excluded_variables)
    filtered_data_excluded_with_constant = add_constant(filtered_data_excluded)

    # VIF for excluded variables dataset
    vif_data_excluded = pd.DataFrame()
    vif_data_excluded['Variable'] = filtered_data_excluded_with_constant.columns
    vif_data_excluded['VIF'] = [variance_inflation_factor(filtered_data_excluded_with_constant.values, i) 
                                for i in range(filtered_data_excluded_with_constant.shape[1])]
    high_vif_excluded = vif_data_excluded[vif_data_excluded['VIF'] > 10].sort_values(by='VIF', ascending=False)
    high_vif_vars_excluded = [var for var in high_vif_excluded['Variable'].tolist() if var != 'const']

    # Correlation matrix for new high VIF variables
    correlation_matrix_excluded = filtered_data_excluded[high_vif_vars_excluded].corr()
    high_correlation_pairs_excluded = [
        (correlation_matrix_excluded.columns[i], correlation_matrix_excluded.columns[j], round(correlation_matrix_excluded.iloc[i, j], 3))
        for i in range(len(correlation_matrix_excluded.columns))
        for j in range(i) if abs(correlation_matrix_excluded.iloc[i, j]) > correlation_threshold
    ]

    # Export results to Excel
    with pd.ExcelWriter("VIF_Results.xlsx") as writer:
        vif_data.to_excel(writer, sheet_name="VIF All Variables", index=False)
        high_vif.to_excel(writer, sheet_name="High VIF Variables", index=False)
        vif_data_excluded.to_excel(writer, sheet_name="VIF Excluded Variables", index=False)
        high_vif_excluded.to_excel(writer, sheet_name="High VIF Excluded", index=False)
        pd.DataFrame(high_correlation_pairs, columns=["Variable 1", "Variable 2", "Correlation"]).to_excel(writer, sheet_name="High Correlation", index=False)
        pd.DataFrame(high_correlation_pairs_excluded, columns=["Variable 1", "Variable 2", "Correlation"]).to_excel(writer, sheet_name="High Correlation Excluded", index=False)

