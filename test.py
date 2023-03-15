import xlwings as xw

# Load the Excel workbook
wb = xw.Book('test.xlsx')

# Create a list to hold the max values from each sheet
max_values = []

# Loop through each sheet and find the max value in column A
for sheet in wb.sheets:
    column_a = sheet.range('A2').expand('down').value
    max_values.append(int(max(column_a)))

# Find the maximum value in the list of max values
max_value = max(max_values)

print(f"The maximum value in column A across all sheets is: {max_value}")
