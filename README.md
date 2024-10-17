Excel-QC-FileSorter

## Overview

Excel-QC-FileSorter is a Python script that automates the quality control (QC) process for lab batch files in Excel format. The script evaluates specific criteria across multiple sheets in each Excel file, moves the files to 'Pass' or 'Fail' folders based on the results, and logs the results into an Excel report.

## Features

Processes all Excel files within a specified directory.
Checks multiple QC criteria across three sheets (Sheet1, Sheet2, Sheet3).
Moves files to 'Pass' or 'Fail' folders based on QC results.
Logs results (Pass/Fail) into a separate results Excel file.

## Prerequisites

Python 3.x

Required packages:

pandas

openpyxl

You can install the required packages using the following command:

pip install pandas openpyxl

## File Structure

Input Directory: Contains the Excel files to be processed.

Pass Folder: Stores files that pass all QC checks.

Fail Folder: Stores files that fail one or more QC checks.

Results File: A designated Excel file that logs the filenames and their pass/fail statuses.

## Usage

Clone the repository:

git clone https://github.com/yourusername/Excel-QC-FileSorter.git

Navigate to the directory and modify the paths in the script to point to your working directory, pass/fail folders, and results file:

mypath = "your_directory_here"

result_path = "result_file_here.xlsx"

pass_path = "pass_folder_here"

fail_path = "fail_folder_here"

## Run the script:

python script_name.py

The script will process all Excel files in the specified directory, move them to the appropriate folders, and log the results.

## Quality Control Criteria

Sheet1:
Checks specific controls for valid "Pass" or "N/A" status.
Ensures a value in the second row, 6th column is below 0.050.

Sheet2:
Validates that coverage for controls is above 800 and that Pass/Fail statuses are either "Pass" or empty.

Sheet3 (optional):
Searches for specific phrases across rows to determine pass/fail status.

## Example Log

The log in the result Excel file will look like this:

File ID	Timestamp	Result

1234	2024-10-17	Pass

5678	2024-10-17	Fail
