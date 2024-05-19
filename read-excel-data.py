import os
import pandas as pd
from tqdm import tqdm

folder_path = r'C:\Users\DwainBrowne\SnapSuite\Sales - Leads - Leads\Q3 - 2023'

# Get all Excel and CSV files in the folder_path
files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.csv'))]

# Define a function to standardize headers for Excel files
def standardize_excel_headers(df):
    standardized_headers = {
        'Business Nam': 'BusinessName',
        'Business Name': 'BusinessName',
        'Number of Em': 'NumberOfEmployees',
        'Number of Employees': 'NumberOfEmployees',
        'Contact Persol': 'ContactPerson',
        'Contact Person': 'ContactPerson',
        'First Name': 'FirstName',
        'Corporate Ema': 'CorporateEmail',
        'Corporate Email': 'CorporateEmail',
        'Email': 'Email',
        'Generic Email': 'Email',
        'Website': 'Website',
        'Phone': 'Phone',
        'Phone Type': 'PhoneType',
        'Street Address': 'StreetAddress',
        'Zip Code': 'ZipCode',
        'State': 'State',
        'City': 'City',
        'Id': 'Id'  # Include if needed
    }
    df.rename(columns=standardized_headers, inplace=True)
    return df

# Define a function to standardize headers for CSV files
def standardize_csv_headers(df):
    standardized_headers = {
        'Industry': 'Industry',
        'Team Size': 'TeamSize',
        'Revenue Range': 'RevenueRange',
        'Total Funding': 'TotalFunding',
        'Work Email #1': 'Email',
        'Work Email #2': 'Email',
        'Work Email #3': 'Email',
        'Work Email #4': 'Email',
        'Work Email #5': 'Email',
        'Work Email #6': 'Email',
        'Work Email #7': 'Email',
        'Direct Email #1': 'Email',
        'Direct Email #2': 'Email',
        'Direct Email #3': 'Email',
        'Direct Email #4': 'Email',
        'Phone #1': 'Phone',
        'Phone #2': 'Phone',
        'Phone #3': 'Phone',
        'Phone #4': 'Phone',
        'Phone #5': 'Phone',
        'Phone #6': 'Phone',
        'Phone #7': 'Phone',
        'Phone #8': 'Phone'
    }
    df.rename(columns=standardized_headers, inplace=True)
    return df

# Function to deduplicate column names
def deduplicate_columns(df):
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols[cols == dup].index.values.tolist()] = [dup + '_' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols
    return df

# Function to merge email columns
def merge_emails(df):
    email_columns = [col for col in df.columns if 'Email' in col]
    df['Email'] = df[email_columns].apply(lambda row: ', '.join(row.dropna().astype(str)), axis=1)
    df.drop(columns=[col for col in email_columns if col != 'Email'], inplace=True)
    return df

dataframes = []
row_counts = {}

# Read each file, standardize headers, merge emails, and append to the list
print("Reading and processing files...")
for file in tqdm(files):
    try:
        file_path = os.path.join(folder_path, file)
        if file.endswith('.xlsx'):
            df = pd.read_excel(file_path)
            df = standardize_excel_headers(df)
        elif file.endswith('.csv'):
            df = pd.read_csv(file_path)
            df = standardize_csv_headers(df)
        df = deduplicate_columns(df)
        df = merge_emails(df)
        row_counts[file] = len(df)
        dataframes.append(df)
        print(f"Successfully processed {file} with {len(df)} rows")
    except Exception as e:
        print(f"Error processing {file}: {e}")

# Filter out empty DataFrames or those with all NA columns
dataframes = [df for df in dataframes if not df.empty and not df.isna().all().all()]

# Ensure all DataFrames have the same columns
all_columns = set(col for df in dataframes for col in df.columns)
for df in dataframes:
    for col in all_columns:
        if col not in df.columns:
            df[col] = pd.NA  # Add missing columns with NA values

# Concatenate all dataframes into one
print("Merging dataframes...")
try:
    merged_df = pd.concat(dataframes, ignore_index=True)
except Exception as e:
    print(f"Error during merging: {e}")

# Remove duplicates
merged_df = merged_df.drop_duplicates()

# Ensure there are no spaces in the final headers
merged_df.columns = [col.replace(' ', '') for col in merged_df.columns]

# Save the merged dataframe to a new excel file
output_path = os.path.join(folder_path, '2024_excel_and_csv_merged_output_006.xlsx')
try:
    merged_df.to_excel(output_path, index=False)
    print(f"Merged file saved to {output_path}")
    print(f"Total rows successfully merged: {len(merged_df)}")
except Exception as e:
    print(f"Error saving the merged file: {e}")

# Print row counts for each file
for file, count in row_counts.items():
    print(f"{file}: {count} rows")
