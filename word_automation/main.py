import PySimpleGUI as sg
from pathlib import Path
import pandas as pd
import xlwings as xw
import os, sys

os.chdir(sys.path[0])

def new_max_value(file_path="./word_automation.xlsm"):
    df_osszesito = pd.read_excel(file_path, ["Összesítő"])
    return len(df_osszesito["Összesítő"]["IKTATÓ"]) + 1
print(new_max_value())


def new_dir(cegnev):
    Company1_DIR = Path.cwd() / f"{cegnev}"
    Company1_DIR.mkdir(exist_ok=True)
    NEW_DIR = Path.cwd() / f"{cegnev}" / f"{new_max_value()}-{cegnev}-2023"
    NEW_DIR.mkdir(exist_ok=True)
    file_name = Path("./word_automation.xlsm")
    wb = xw.Book(file_name)
    sh = wb.sheets("Összesítő")
    sh.range(f"U{new_max_value() + 1}").value = f"{new_max_value()}-{cegnev}-2023"
    wb.save("./word_automation.xlsm")
    return

layout = [
    [sg.Text("Válassz céget:"), sg.OptionMenu(values = ["cég1", "cég2", "cég3"], key="-CEG_NEV-"), sg.Button("Új mappa")],
    [sg.Text("Ügyfél adatok beolvasása:"), sg.Input(key="-IN-"), sg.FileBrowse()],
    [sg.Button("Export wordbe"), sg.Exit()],
]

window = sg.Window("Excelmagic", layout)

while True:
    event, values = window.read()
    print(event, values)
    if event == sg.WIN_CLOSED or event == "Exit":
        break

    if event == "Új mappa":
        new_dir(values["-CEG_NEV-"])
    if event == "Új ügyfél beolvasás":
        print(event, values)
    if event == "Export wordbe":
        print(event, values)

window.close()