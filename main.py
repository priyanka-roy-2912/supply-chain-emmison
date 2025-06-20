import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt

from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.preprocessing import StandardScaler
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error, r2_score
import joblib

# Path to Excel file
excel_file = "SupplyChainEmissionFactorsforUSIndustriesCommodities.xlsx"

# Define year range
years = range(2010, 2017)

# Read sample sheet for testing (optional)
# df_1 = pd.read_excel(excel_file, sheet_name=f'{years[0]}_Detail_Commodity')
# df_2 = pd.read_excel(excel_file, sheet_name=f'{years[0]}_Detail_Industry')

# Process all sheets and combine data
all_data = []

for year in years:
    try:
        df_com = pd.read_excel(excel_file, sheet_name=f'{year}_Detail_Commodity')
        df_ind = pd.read_excel(excel_file, sheet_name=f'{year}_Detail_Industry')

        # Add metadata
        df_com['Source'] = 'Commodity'
        df_ind['Source'] = 'Industry'
        df_com['Year'] = df_ind['Year'] = year

        # Clean column names
        df_com.columns = df_com.columns.str.strip()
        df_ind.columns = df_ind.columns.str.strip()

        # Rename for consistency
        df_com.rename(columns={
            'Commodity Code': 'Code',
            'Commodity Name': 'Name'
        }, inplace=True)

        df_ind.rename(columns={
            'Industry Code': 'Code',
            'Industry Name': 'Name'
        }, inplace=True)

        # Combine and append
        all_data.append(pd.concat([df_com, df_ind], ignore_index=True))

    except Exception as e:
        print(f"Error processing year {year}: {e}")

# Merge all years into one DataFrame
df = pd.concat(all_data, ignore_index=True)

# Preview
print("\nFirst 10 rows:")
print(df.head(10))

# Summary
print(f"\nTotal rows combined: {len(df)}")
print("\nColumns:")
print(df.columns)

# Missing values
print("\nMissing values per column:")
print(df.isnull().sum())
