import pandas as pd
import re

# Set the file path to your Excel file
file_path = r'C:\Users\DwainBrowne\SnapSuite\Sales - Leads - Leads\Q3 - 2023\2024_standardized_excel_column_cleaned_file_001.xlsx'

# Load the Excel file into a DataFrame
df = pd.read_excel(file_path)

# Time zone mappings for all U.S. states
time_zone_mapping = {
    'Eastern': ['CT', 'DE', 'FL', 'GA', 'IN', 'KY', 'ME', 'MD', 'MA', 'MI', 'NH', 'NJ', 'NY', 'NC', 'OH', 'PA', 'RI', 'SC', 'TN', 'VT', 'VA', 'WV'],
    'Central': ['AL', 'AR', 'IL', 'IA', 'KS', 'KY', 'LA', 'MN', 'MS', 'MO', 'OK', 'SD', 'TN', 'TX', 'WI'],
    'Mountain': ['AZ', 'CO', 'ID', 'MT', 'NM', 'UT', 'WY'],
    'Pacific': ['CA', 'NV', 'OR', 'WA'],
    'Alaska': ['AK'],
    'Hawaii': ['HI']
}

# Area code mappings to time zones (partial list for example)
area_code_mapping = {
    'Eastern': ['212', '315', '347', '516', '518', '607', '631', '716', '718', '845', '914'],
    'Central': ['205', '251', '256', '334', '938', '479', '501', '870'],
    'Mountain': ['303', '719', '970'],
    'Pacific': ['209', '213', '310', '323', '408', '415', '424', '510', '530', '559', '562', '619', '626', '650', '661', '707', '714', '760', '805', '818', '831', '858', '909', '916', '925', '949'],
    'Alaska': ['907'],
    'Hawaii': ['808']
}

# Generate time slots in 15-minute increments from 9am to 1pm EST
time_slots = [f'{hour}:{minute:02d}am' for hour in range(9, 12) for minute in range(0, 60, 15)] + \
             [f'12:{minute:02d}pm' for minute in range(0, 60, 15)] + \
             [f'1:{minute:02d}pm' for minute in range(0, 15, 15)]

def get_time_zone_from_state(state):
    """
    Determine the time zone based on the state.
    """
    for tz, states in time_zone_mapping.items():
        if state in states:
            return tz
    return None

def get_time_zone_from_area_code(area_code):
    """
    Determine the time zone based on the phone number area code.
    """
    for tz, area_codes in area_code_mapping.items():
        if area_code in area_codes:
            return tz
    return None

def best_time_to_call(state, phone):
    """
    Determine the best time to call based on state or phone area code.
    """
    time_zone = get_time_zone_from_state(state)
    
    if time_zone is None and phone:
        phone = str(phone)
        area_codes = re.findall(r'\d+', phone)
        if area_codes:
            area_code = area_codes[0][:3]
            time_zone = get_time_zone_from_area_code(area_code)
    
    if time_zone == 'Eastern':
        return time_slots[0]  # 9:00am
    elif time_zone == 'Central':
        return time_slots[4]  # 10:00am
    elif time_zone == 'Mountain':
        return time_slots[8]  # 11:00am
    elif time_zone == 'Pacific':
        return time_slots[12]  # 12:00pm
    elif time_zone == 'Alaska':
        return time_slots[16]  # 1:00pm
    elif time_zone == 'Hawaii':
        return time_slots[16]  # 1:00pm
    return '9:00am'  # Default to 9:00am Eastern if time zone not found

# Add progress update
total_rows = len(df)
print(f"Total rows to process: {total_rows}")

# Add a new column 'BestTimeToCall' based on the state or area code
df['BestTimeToCall'] = df.apply(lambda row: best_time_to_call(row['State'], row['Phone']), axis=1)

# Save the updated DataFrame back to a new Excel file
output_file_path = r'C:\Users\DwainBrowne\SnapSuite\Sales - Leads - Leads\Q3 - 2023\updated_file.xlsx'
df.to_excel(output_file_path, index=False)

print(f'The updated file has been saved to {output_file_path}')
