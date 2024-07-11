import pandas as pd

# Define the working folder path
working_folder = 'D:/GDMS/Results/Tino data processing REC1 July 2024/ICPMS vs GMDS/ICPMS files/'

# Ensure the folder path ends with a separator
if not working_folder.endswith('/'):
    working_folder += '/'
    
# Define the function to convert non-numerical values
def convert_to_number(value):
    if isinstance(value, str) and ('<' in value or '=' in value):
        return float(value.lstrip('<=').strip())
    try:
        return float(value)
    except ValueError:
        return value

# Load and clean each ICPMS file
def load_and_clean_icpms(file_name, condition):
    file_path = working_folder + file_name
    df = pd.read_excel(file_path, header=7)
    df = df.dropna(how='all', axis=0)  # Drop empty rows
    df.columns = ['Element', 'Element Concentration (ppmw)'] + list(df.columns[2:])
    df = df[['Element', 'Element Concentration (ppmw)']].dropna(subset=['Element'])
    df['Condition'] = condition
    df['Element Concentration (ppmw)'] = df['Element Concentration (ppmw)'].apply(convert_to_number)
    return df

# Define file names and conditions
file_names_icpms = [
    'REC1 unleached ICPMS.xlsx'
    'REC1 HNO3 ICPMS.xlsx',
    'REC1 HCL ICPMS.xlsx',
    'REC1 H2SO4 ICPMS.xlsx',
]

conditions_icpms = ['unleached', 'HNO3', "HCL", "H2SO4"]

# Load and clean data
cleaned_dfs_icpms = [load_and_clean_icpms(file, cond) for file, cond in zip(file_names_icpms, conditions_icpms)]

# Combine all ICPMS dataframes into a single dataframe
combined_df_icpms = pd.concat(cleaned_dfs_icpms, ignore_index=True)

# Save the combined dataframe to a single Excel file
combined_file_path_icpms = working_folder + 'combined_ICPMS_cleaned.xlsx'
combined_df_icpms.to_excel(combined_file_path_icpms, index=False)

# Display the combined dataframe
print(combined_df_icpms.head())
