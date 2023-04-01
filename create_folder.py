from pathlib import Path
import os, sys
import xlwings as xw

os.chdir(sys.path[0])

def max_value(file_path):
    wb = xw.Book(file_path) # Load the Excel workbook

    max_values = [] # Create a list to hold the max values from each sheet

    for sheet in wb.sheets: # Loop through each sheet and find the max value in column A
        column_a = sheet.range('A2').expand('down').value
        max_values.append(int(max(column_a)))
    
    max_value = max(max_values) # Find the maximum value in the list of max values

    return max_value

max_val = max_value('./Körlevél_HMKE.xlsm')

# Create output directory
Company1_DIR = Path.cwd() / "company1"
Company1_DIR.mkdir(exist_ok=True)

NEW_DIR = Path.cwd() / "company1" / f"{max_val + 1}-company1-2023"
NEW_DIR.mkdir(exist_ok=True)

# Type new_max_value to the first empty cell

wb = xw.Book(r'./Körlevél_HMKE.xlsm')

ws = wb.sheets[0] # 

column_a = ws.range('A2').expand('down')

ws[(f'A{len(column_a) + 2}')].value = max_val + 1