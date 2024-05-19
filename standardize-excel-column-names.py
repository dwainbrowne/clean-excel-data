"""
This script processes an Excel file to standardize column names and merge columns with similar meanings. 
It reads the Excel file, maps columns to standard names, merges columns with the same standardized names by selecting the first non-null value, and saves the cleaned data to a new Excel file.

The script performs the following steps:
1. Defines a function to standardize column names based on common names.
2. Reads the Excel file into a pandas DataFrame.
3. Applies the standardization function to all column names.
4. Merges columns with the same standardized names by choosing the first non-null value.
5. Splits "ContactPerson" into "FirstName" and "LastName" if applicable.
6. Cleans the "LastName" field to remove job titles in parentheses.
7. Saves the cleaned DataFrame to a new Excel file.

To use the script:
1. Replace 'path_to_your_excel_file.xlsx' with the full path to the Excel file you want to process.
2. Replace 'path_to_save_cleaned_file.xlsx' with the desired path for the cleaned Excel file.

Prerequisites:
- Python 3.x
- pandas library (install using `pip install pandas`)
- tqdm library (install using `pip install tqdm`)
"""

import pandas as pd
from tqdm import tqdm
import re

# Define a function to standardize column names
def standardize_column_names(column):
    column = column.strip().lower()
    if column in ['businessname', 'companyname', 'companyname.1', 'companyname.2', 'company','businessname.1']:
        return 'BusinessName'
    elif column in ['numberofemployees', 'numberofemployees.', 'teamsize']:
        return 'NumberOfEmployees'
    elif column in ['contactperson', 'fullname', 'name', 'contactperson.1', 'commercialcleaningservice']:
        return 'ContactPerson'
    elif column in ['firstname', 'firstname.1']:
        return 'FirstName'
    elif column in ['lastname']:
        return 'LastName'
    elif column == 'email':
        return 'Email'
    elif column in ['website', 'companywebsite']:
        return 'Website'
    elif column in ['phone', 'phone_1', 'phone_2', 'phone_3', 'phone_4', 'phone_5', 'phone_6', 'phone_7', 'companyphone']:
        return 'Phone'
    elif column in ['phonetype', 'phonetype.1', 'clearoutphonelinetype']:
        return 'PhoneType'
    elif column in ['streetaddress', 'streetaddress.1']:
        return 'StreetAddress'
    elif column in ['zipcode', 'zipcode.1']:
        return 'ZipCode'
    elif column == 'state':
        return 'State'
    elif column in ['city', 'city.1','location']:
        return 'City'
    elif column == 'prospectlinkedinurl':
        return 'LinkedInURL'
    elif column in ['facebookprofile','facebookprofile.1','companyfacebook']:
        return 'FacebookProfile'
    elif column in ['occupation','position','jobtitle']:
        return 'JobTitle'
    elif column in ['position', 'prospectposition']:
        return 'Position'
    elif column == 'linkedinurl':
        return 'LinkedInURL'
    elif column == 'industry':
        return 'Industry'
    elif column == 'revenue':
        return 'RevenueRange'
    elif column == 'birthday':
        return 'Birthday'
    elif column == 'location':
        return 'Location'
    elif column == 'lastknowsoftware':
        return 'LastKnownSoftware'
    elif column == 'totalfunding':
        return 'TotalFunding'
    elif column == 'clearoutphonecountryname':
        return 'CountryName'
    elif column.startswith('clearoutphone'):
        return column.replace('clearoutphone', '').capitalize()
    else:
        return column

# Clean up last name field to remove job titles in parentheses
def clean_last_name(name):
    return re.sub(r'\s*\(.*?\)\s*', '', name)

# File paths
file_path = r'C:\Users\DwainBrowne\SnapSuite\Sales - Leads - Leads\Q3 - 2023\2024_excel_and_csv_merged_output_007.xlsx'
new_file_path = r'C:\Users\DwainBrowne\SnapSuite\Sales - Leads - Leads\Q3 - 2023\2024_standardized_excel_column_cleaned_file_001.xlsx'

try:
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Standardize column names with progress tracking
    print("Standardizing column names...")
    df.columns = [standardize_column_names(col) for col in tqdm(df.columns, desc="Standardizing columns")]

    # Merge columns with the same standardized names by choosing the first non-null value with progress tracking
    print("Merging columns...")
    merged_df = df.T.groupby(level=0).first().T

    # Split "ContactPerson" into "FirstName" and "LastName"
    if 'ContactPerson' in merged_df.columns:
        try:
            contact_split = merged_df['ContactPerson'].str.split(n=1, expand=True)
            merged_df['FirstName'] = contact_split[0]
            merged_df['LastName'] = contact_split[1].fillna('')  # Handle cases where there is no last name
            
            # Clean up the last name to remove job titles in parentheses
            merged_df['LastName'] = merged_df['LastName'].apply(clean_last_name)
            
            merged_df.drop(columns=['ContactPerson'], inplace=True)
        except Exception as split_error:
            print(f"Error splitting 'ContactPerson': {split_error}")
            merged_df['FirstName'] = merged_df['ContactPerson']
            merged_df['LastName'] = ''

    # Save the cleaned data to a new Excel file
    print("Saving cleaned data to new Excel file...")
    merged_df.to_excel(new_file_path, index=False)

    print(f"Data has been cleaned and saved to {new_file_path}")
except PermissionError as e:
    print(f"PermissionError: {e}. Please ensure the file is not open and you have the necessary permissions.")
except Exception as e:
    print(f"An error occurred: {e}")
