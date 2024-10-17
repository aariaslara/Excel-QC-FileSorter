import os
import pandas as pd
from openpyxl import load_workbook

# Define paths
mypath = "your_directory_here"
result_path = "result_file_here.xlsx"
pass_path = "pass_folder_here"
fail_path = "fail_folder_here"

# List of all filenames in the directory
filenamelist = os.listdir(mypath)
results = []

# Function to check the first sheet
def check_sheet1(sheet):
    # Define control labels and acceptable values
    controls = {"control1", "control2", "control3"}
    list_of_valid_values = ["Pass", "N/A"]  # Acceptable values
    
    # Iterate through each row in the sheet
    for row, col in sheet.iterrows():
        metrics1 = col[0]  # First metric to check
        metrics2 = col[1]  # Second metric to check
        pass_na1 = col[2]  # Result for the first metric
        pass_na2 = col[3]  # Result for the second metric
        
        # Check if the metric is a control and its result is invalid
        if metrics1 in controls:
            if pass_na1 not in list_of_valid_values:
                return False
        if metrics2 in controls:
            if pass_na2 not in list_of_valid_values:
                return False
    
    # Additional check for value in the second row, 6th column
    value_to_check = sheet.iloc[1, 5]
    value_to_check = float(value_to_check)
    if not value_to_check < 0.050:  # Ensure the value is below 0.050
        return False
    
    return True  # Return True if all controls pass

# Function to check the second sheet
def check_sheet2(sheet):
    controls = {"controlA", "controlB", "controlC", "controlD"}  # List of controls to check

    # Iterate over the rows
    for index, row in sheet.iterrows():
        sample = row[1]  # First column: sample name
        coverage = row[2]  # Coverage column
        pass_fail = row[3]  # Pass/Fail column 1
        pass_fail2 = row[4]  # Pass/Fail column 2

        # Check if the sample is a control, and validate its coverage and status
        if sample in controls:
            coverage = float(coverage)
            # Ensure coverage is above 800 and both pass_fail and pass_fail2 are either "Pass" or empty
            if coverage < 800.0 or (pass_fail not in {"Pass", ""} or pass_fail2 not in {"Pass", ""}):
                return False

    return True  # Return True if all controls pass

# Function to check the third sheet
def check_sheet3(sheet):
    phrase1 = "expected_phrase_1"
    phrase2 = "expected_phrase_2"

    # Iterate over the rows, looking for specific phrases
    for index, row in sheet.iterrows():
        # Check if phrase1 is part of the current row's first column value
        if phrase1 in row[0]:
            # Check if the next row contains phrase2
            if index + 1 < len(sheet) and sheet.iloc[index + 1, 0] == phrase2:
                return True  # Both conditions are satisfied

    return False  # Return False if no match is found

# Iterate through the filenames in the directory
for filename in filenamelist:
    if filename.lower().endswith(".xlsx"):  # Process only Excel files
        file_path = os.path.join(mypath, filename)
        # Load the Excel file into a dictionary of DataFrames, one for each sheet
        workbook = pd.read_excel(file_path, header=None, sheet_name=None, na_filter=False)
        
        # Access the specific sheets
        sheet1 = workbook['Sheet1']
        sheet2 = workbook['Sheet2']
        
        # Check if the 'QC messages' sheet exists
        if 'Sheet3' in workbook:
            sheet3 = workbook['Sheet3']
            # Check all sheets for pass criteria
            if check_sheet1(sheet1) and check_sheet2(sheet2) and check_sheet3(sheet3):
                results.append([filename[8:12], filename[13:23], 'Pass'])  # Append 'Pass' result
                print(results)
                # Move file to 'Pass' folder
                src_path = os.path.join(mypath, filename)
                desc_path = os.path.join(pass_path, filename)
                os.rename(src_path, desc_path)
            else:
                results.append([filename[8:12], filename[13:23], 'Fail'])  # Append 'Fail' result
                print(results)
                # Move file to 'Fail' folder
                src_path = os.path.join(mypath, filename)
                desc_path = os.path.join(fail_path, filename)
                os.rename(src_path, desc_path)
        else:
            # If 'QC messages' sheet does not exist, only check sheets 1 and 2
            if check_sheet1(sheet1) and check_sheet2(sheet2):
                results.append([filename[8:12], filename[13:23], 'Pass'])  # Append 'Pass' result
                print(results)
                # Move file to 'Pass' folder
                src_path = os.path.join(mypath, filename)
                desc_path = os.path.join(pass_path, filename)
                os.rename(src_path, desc_path)
            else:
                results.append([filename[8:12], filename[13:23], 'Fail'])  # Append 'Fail' result
                print(results)
                # Move file to 'Fail' folder
                src_path = os.path.join(mypath, filename)
                desc_path = os.path.join(fail_path, filename)
                os.rename(src_path, desc_path)

# Save the results to Excel
workbook = load_workbook(result_path)
worksheet = workbook['Sheet1']

# Append each result to the worksheet
for result in results:
    worksheet.append(result)

# Save the updated workbook
workbook.save(result_path)
