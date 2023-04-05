from pathlib import Path
import os, sys
import xlwings as xw

os.chdir(sys.path[0])

# Create output directory
Company1_DIR = Path.cwd() / "company1"
Company1_DIR.mkdir(exist_ok=True)

max_val = 8

NEW_DIR = Path.cwd() / "company1" / f"{max_val+1}-company1-2023"
NEW_DIR.mkdir(exist_ok=True)

# Type new_max_value to the first empty cell

wb = xw.Book(r'./Körlevél_HMKE.xlsm')

ws = wb.sheets[0] # 

column_a = ws.range('B2').expand('down')

ws[(f'A{len(column_a) + 2}')].value = max_val + 1