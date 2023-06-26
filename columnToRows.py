import pandas as pd
import numpy as np

# Read the Excel file
df = pd.read_excel('filepath/formatMePls.xlsx', header=None, engine='openpyxl')

df = df.dropna(how='all')  # Drop completely empty rows
df = df.reset_index(drop=True)  # Reset index

# Initialize variables
records = []
record = {}
record_fields = ["Name", "Title", "Organization"]
common_titles = ["MD", "Doctor", "Professor", "Fellow", "Resident", "Associate Professor", 
                 "Senior Clinical Scientist", "Academic Clinical Lecturer"]  # Add more common titles here

# Iterate over DataFrame
for i in range(len(df)):
    cell = df.loc[i, 0]
    
    if pd.isna(cell):
        continue  # Skip row if it is NaN

    # Check if a new record should start
    if len(cell) == 2 and i+1 < len(df) and (len(df.loc[i+1, 0].split()) > 1 and df.loc[i+1, 0] != df.loc[i+1, 0].upper() or i+2 == len(df) or len(df.loc[i+2, 0]) == 2) and cell not in ["Dr", "MD"]:  
        if record:  # If previous record exists, add it to records
            records.append(record)
        record = {'Initials': cell}  # Start new record
        next_field_index = 0  # Reset the index for the next field
    else:  # If we're in the middle of a record
        if next_field_index < len(record_fields):  # If there are still fields left to fill
            # If it's the third line and it's not a common title, treat it as an organization
            if next_field_index == 1 and cell not in common_titles:
                record["Organization"] = cell
                next_field_index += 2
            else:
                record[record_fields[next_field_index]] = cell
                next_field_index += 1

# Append last record
if record:
    records.append(record)

# Create a new DataFrame with structured data
df = pd.DataFrame(records)

df.to_excel('reformatedFinal3.xlsx', index=False)
