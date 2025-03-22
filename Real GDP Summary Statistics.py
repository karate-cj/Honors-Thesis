import pandas as pd
import numpy as np
from scipy.stats import skew, kurtosis

# Load the dataset from Excel
data = pd.read_excel("C:/Users/cjfar/ECON 498H (Honors Thesis)/Second Draft/Honors Thesis Dataset 2.0.xlsx", sheet_name="Table")

# Clean column names by removing spaces, special characters, and hidden characters
data.columns = data.columns.str.replace(' ', '_').str.replace('\xa0', '').str.strip()

# Calculate descriptive statistics for the Real_GDP column
real_gdp = data['Real_GDP']

# Calculate additional statistics
summary_stats = {
    "Statistic": [
        "Count",
        "Mean",
        "Standard Deviation",
        "Minimum",
        "25th Percentile",
        "Median (50th Percentile)",
        "75th Percentile",
        "Maximum",
        "Variance",
        "Skewness",
        "Kurtosis"
    ],
    "Value": [
        round(real_gdp.count(), 3),  # Count has no $
        round(real_gdp.mean(), 3),  # Mean
        round(real_gdp.std(), 3),  # Standard Deviation
        round(real_gdp.min(), 3),  # Minimum
        round(real_gdp.quantile(0.25), 3),  # 25th Percentile
        round(real_gdp.median(), 3),  # Median
        round(real_gdp.quantile(0.75), 3),  # 75th Percentile
        round(real_gdp.max(), 3),  # Maximum
        round(real_gdp.var(), 3),  # Variance
        round(skew(real_gdp, nan_policy='omit'), 3),  # Skewness
        round(kurtosis(real_gdp, nan_policy='omit'), 3)  # Kurtosis
    ]
}

# Convert the summary stats dictionary into a DataFrame
summary_stats_df = pd.DataFrame(summary_stats)

# Save the results to an Excel file
output_path = "C:/Users/cjfar/ECON 498H (Honors Thesis)/Third Draft/Python Code 3.0/Summary Statistics/summary_stats.xlsx"
summary_stats_df.to_excel(output_path, index=False)

print(f"Summary statistics saved to {output_path}")


