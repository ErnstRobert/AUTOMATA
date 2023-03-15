import xlwings as xw

wb = xw.Book(r'c:/Users/I575327/Documents/Áram projekt/suncollector/Körlevél_HMKE.xlsm')

ws = wb.sheets[0]

column_a = ws.range('A2').expand('down')

print(len(column_a))