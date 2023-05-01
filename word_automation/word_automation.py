import PySimpleGUI as sg
from pathlib import Path
import pandas as pd
import os, sys
import xlwings as xw
from docxtpl import DocxTemplate

os.chdir(sys.path[0])

def new_max_value(file_path="./word_automation.xlsm"):
    df_osszesito = pd.read_excel(file_path, ["Összesítő"])
    return len(df_osszesito["Összesítő"]["IKTATÓ"]) + 1

def priv_max_value(file_path="./word_automation.xlsm"):
    df_priv = pd.read_excel(file_path, ["Magánszemély"])
    return len(df_priv["Magánszemély"]["Irányítószám"]) + 1

def corp_max_value(file_path="./word_automation.xlsm"):
    df_corp = pd.read_excel(file_path, ["Jogi személy"])
    return len(df_corp["Jogi személy"]["Irányítószám"]) + 1

def new_dir(cegnev):
    Company1_DIR = Path.cwd() / f"{cegnev}"
    Company1_DIR.mkdir(exist_ok=True)
    NEW_DIR = Path.cwd() / f"{cegnev}" / f"{new_max_value()}-{cegnev}-2023"
    NEW_DIR.mkdir(exist_ok=True)
    file_name = Path("./word_automation.xlsm")
    wb = xw.Book(file_name)
    sh = wb.sheets("Összesítő")
    sh.range(f"U{new_max_value() + 1}").value = f"{new_max_value()}-{cegnev}-2023"
    sh.range(f"AH{new_max_value() + 1}").value = f"{NEW_DIR}"
    wb.save("./word_automation.xlsm")
    return NEW_DIR 

def load_data(cust_file_path):
    cust_file_name= Path(cust_file_path)
    wb_cust = xw.Book(cust_file_name)
    wb_main = xw.Book("./word_automation.xlsm")
    sht_priv = wb_main.sheets("Magánszemély")
    sht_corp = wb_main.sheets("Jogi személy")

    if wb_cust.sheets("magánszemély").range("D2").value == "Magánszemély":
        sht_cust = wb_cust.sheets("magánszemély")
        sht_priv.range(f"A{priv_max_value() + 1}").value = sht_cust.range("A2:T2").value
        sht_priv.range(f"U{priv_max_value() + 1}").value = cust_file_path.split("/")[-2]
        sht_priv.range(f"V{priv_max_value() + 1}").value = sht_cust.range("V2:AC2").value
        wb_main.save("./word_automation.xlsm")
        priv_max_value()
        return
    else:
        sht_cust = wb_cust.sheets("Jogi személy")
        sht_corp.range(f"A{corp_max_value() + 1}").value = sht_cust.range("A2:U2").value
        sht_corp.range(f"V{corp_max_value() + 1}").value = cust_file_path.split("/")[-2]
        sht_corp.range(f"W{corp_max_value() + 1}").value = sht_cust.range("W2:AD2").value
        wb_main.save("./word_automation.xlsm")
        corp_max_value()
        return
    
def main():
    wb = xw.Book.caller()
    sht_panel = wb.sheets['PANEL']
    doc = DocxTemplate('hmke.docx')
    plan1 = DocxTemplate('1fázis.docx')
    plan3 = DocxTemplate('3fázis.docx')

    context = sht_panel.range('A2').options(dict, expand='table', numbers=int).value
    print(context)
    output_name = f"{context['PATH_TO_DIR']}\proba_{context['IKTATÓ']}.docx"
    output_plan = f"{context['PATH_TO_DIR']}\plan_{context['IKTATÓ']}.docx"
    doc.render(context)
    doc.save(output_name)
    if sht_panel.range('B22').value == "1 fázis":
        plan1.render(context)
        plan1.save(output_plan)
    else:
        plan3.render(context)
        plan3.save(output_plan)

layout = [
    [sg.Text("Válassz céget:"), sg.OptionMenu(values = ["cég1", "cég2", "cég3"], key="-CEG_NEV-"), sg.Button("Új mappa létrehozása")],
    [sg.Text("Ügyfél adatok beolvasása:"), sg.Input(key="-IN-"), sg.FileBrowse()],
    [sg.Button("Ügyfél beolvasás összesítőbe"), sg.Button("Export wordbe"), sg.Exit()],
]

window = sg.Window("Excelmagic", layout)

while True:
    event, values = window.read()
    
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    print(event, values)
    if event == "Új mappa létrehozása":
        # new_dir(values["-CEG_NEV-"])                              <-------- !!!!!!!!!! ha nem kell felugró ablak !!!!!!!!!!!!!!!!!
        sg.PopupScrolled(f"Új mappa: {new_dir(values['-CEG_NEV-'])}")
    if event == "Ügyfél beolvasás összesítőbe":
        load_data(values["-IN-"])
        priv_max_value()
        corp_max_value()
    if event == "Export wordbe":
        if __name__ == '__main__':
            xw.Book("word_automation.xlsm").set_mock_caller()
            main()

window.close()
