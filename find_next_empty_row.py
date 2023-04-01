import xlwings as xw

wb = xw.Book(r'c:suncollector/Körlevél_HMKE.xlsm')

ws = wb.sheets[0] # 0= magányszemély, 1= jogi személy

column_a = ws.range('A2').expand('down')

print(len(column_a))

ws[(f'A{len(column_a) + 2}')].value = max_val + 1