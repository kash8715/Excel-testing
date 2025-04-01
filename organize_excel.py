import pandas as pd
import numpy as np

# Read the original Excel file
df = pd.read_excel('DHS1.xlsx')

# Create a new DataFrame with organized data
organized_data = []

# Extract performance indicators
performance_indicators = [
    "Average number of individuals in adult families in shelters per day",
    "Average number of families with children in shelters per day",
    "Average number of individuals in families with children in shelters per day",
    "Average number of single adults in shelters per day",
    "Adult families entering the DHS shelter services system"
]

# Create a clean DataFrame for performance indicators
performance_df = pd.DataFrame({
    'Performance Indicator': performance_indicators,
    'FY20': [5177, 11719, 36548, 16866, 1118],
    'FY21': [4186, 9823, 30212, 18012, 528],
    'FY22': [3130, 8505, 25969, 16465, 598],
    'FY23': [5119, 12749, 40915, 20162, 777],
    'FY24': [4749, 18652, 61103, 20468, 1479],
    'Target FY24': ['↓', '↓', '↓', '↓', '↓'],
    'Target FY25': ['↓', '↓', '↓', '↓', '↓'],
    '5-Year Trend': ['Neutral', 'Up', 'Up', 'Up', 'Up'],
    'Desired Direction': ['Down', 'Down', 'Down', 'Down', 'Down']
})

# Create a clean DataFrame for budget information
budget_data = [
    ['101 - Administration', 30.8, 35.9, 'All'],
    ['102 - Street Programs', 9.9, 10.9, '3a'],
    ['Other Than Personal Services - Total', 3381.4, 3841.7, 'All'],
    ['200 - Shelter Intake and Program', 3049.9, 3471.9, 'All'],
    ['201 - Administration', 30.8, 34.1, 'All'],
    ['202 - Street Programs', 300.7, 335.7, '3a'],
    ['Agency Total', 3540.4, 4017.9, np.nan]
]

budget_df = pd.DataFrame(budget_data, columns=['Category', 'FY20', 'FY21', 'Type'])

# Create Excel writer object
with pd.ExcelWriter('organized_dhs_data.xlsx') as writer:
    # Write performance indicators to first sheet
    performance_df.to_excel(writer, sheet_name='Performance Indicators', index=False)
    
    # Write budget information to second sheet
    budget_df.to_excel(writer, sheet_name='Budget Information', index=False)
    
    # Auto-adjust column widths
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for idx, col in enumerate(performance_df if sheet_name == 'Performance Indicators' else budget_df):
            max_length = max(
                performance_df[col].astype(str).apply(len).max() if sheet_name == 'Performance Indicators'
                else budget_df[col].astype(str).apply(len).max(),
                len(str(col))
            )
            worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2

print("Organized Excel file has been created as 'organized_dhs_data.xlsx'") 