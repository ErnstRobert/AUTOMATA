import os, sys
import xlwings as xw
from docxtpl import DocxTemplate

os.chdir(sys.path[0])

def main():
    wb = xw.Book.caller()
    sht_panel = wb.sheets['PANEL']
    doc = DocxTemplate('./hmke.docx')

    context = sht_panel.range('A2').options(dict, expand='table', numbers=int).value
    output_name = f"proba_{context['IKTATÓ']}.docx"
    doc.render(context)
    doc.save(output_name)

if __name__ == '__main__':
    xw.Book('./word_automation.xlsm').set_mock_caller()
    main()
