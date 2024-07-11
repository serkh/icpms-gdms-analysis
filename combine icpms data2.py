import os
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
    file_path = os.path.join(working_folder, file_name)
    print(f"Loading file: {file_path}")  # Debug print statement
    df = pd.read_excel(file_path, header=0)
    df = df.dropna(how='all', axis=0)  # Drop empty rows
    df.columns = ['Element', 'Element Concentration (ppmw)'] + list(df.columns[2:])
    df['Element Concentration (ppmw)'] = df['Element Concentration (ppmw)'].apply(convert_to_number)
    df = df.dropna(subset=['Element'])  # Drop rows where 'Element' is NaN after conversion
    df['Condition'] = condition
    print(f"File {file_name} loaded with {df.shape[0]} rows and {df['Element'].nunique()} unique elements.")  # Debug print statement
    return df

# Define file names and conditions
file_names_icpms = [
    'REC1 unleached ICPMS.xlsx',
    'REC1 HNO3 ICPMS.xlsx',
    'REC1 HCL ICPMS.xlsx',
    'REC1 H2SO4 ICPMS.xlsx',
]

conditions_icpms = ['unleached', 'HNO3', 'HCL', 'H2SO4']

# Load and clean data
cleaned_dfs_icpms = []
for file, cond in zip(file_names_icpms, conditions_icpms):
    cleaned_df = load_and_clean_icpms(file, cond)
    cleaned_dfs_icpms.append(cleaned_df)
    # Debugging: Print unique elements found in each file
    print(f"After loading {file}, elements found: {cleaned_df['Element'].unique()}")

# Combine all ICPMS dataframes into a single dataframe
combined_df_icpms = pd.concat(cleaned_dfs_icpms, ignore_index=True)

# Save the combined dataframe to a single Excel file
combined_file_path_icpms = os.path.join(working_folder, 'combined_ICPMS_cleaned5.xlsx')
combined_df_icpms.to_excel(combined_file_path_icpms, index=False)

# Display the combined dataframe and check the number of elements
print(f"Combined dataframe has {combined_df_icpms['Element'].nunique()} unique elements.")
print(combined_df_icpms.head())

# Debugging: Print all unique elements in the combined dataframe
print(f"All unique elements in combined dataframe: {combined_df_icpms['Element'].unique()}")
