import xlwings as xw

def max_value(file_path):
    wb = xw.Book(file_path) # Load the Excel workbook

    max_values = [] # Create a list to hold the max values from each sheet

    for sheet in wb.sheets: # Loop through each sheet and find the max value in column A
        column_a = sheet.range('A2').expand('down').value
        max_values.append(int(max(column_a)))
    
    max_value = max(max_values) # Find the maximum value in the list of max values

    return max_value

max_val = max_value('test.xlsx')

# Print the result
print(f"The maximum value in column A across all sheets is: {max_val}")