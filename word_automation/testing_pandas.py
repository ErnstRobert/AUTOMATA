from pathlib import Path
import pandas as pd
import re
import xlwings as xw

current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd
excel_filepath = current_dir / "word_automation.xlsm"

excel_file = pd.ExcelFile(excel_filepath)
wb = xw.Book(excel_filepath)

input_munkalapok = excel_file.sheet_names[1:3]

munkalap = {}
for sheet_name in input_munkalapok:
    munkalap[sheet_name] = excel_file.parse(sheet_name)
munkalap

usecols = {
  "Magánszemély": "U:U",
  "Jogi személy": "V:V"
}

iktato = {}
for sheet_name in input_munkalapok:
    iktato[sheet_name] = excel_file.parse(
        sheet_name, usecols=usecols.get(sheet_name, None)
    )
iktato

osszes_iktato = pd.concat(iktato.values(), ignore_index=True)
osszes_munkalap = pd.concat(munkalap.values(), ignore_index=True)

sht = wb.sheets("Összesítő")

sht.range("A1").options(index=False).value = osszes_munkalap