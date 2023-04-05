import PySimpleGUI as sg
from pathlib import Path
import pandas as pd
import xlwings as xw
import os, sys

os.chdir(sys.path[0])

def new_max_value(file_path="./word_automation.xlsm"):
    df_osszesito = pd.read_excel(file_path, ["Összesítő"])
    return len(df_osszesito["Összesítő"]["Irányítószám"])
print(new_max_value())

def new_dir():
    NEW_DIR = Path.cwd() / f"COMP1" / f"{new_max_value()}-COMP1-2023"   #<---- inputból kell a cégnév
    NEW_DIR.mkdir(exist_ok=True)
    return 

layout = [
    [sg.Text("Cég név:"), sg.Input(key="cegnev", do_not_clear=False)],
    [sg.Button("Új mappa"), sg.Button("Új ügyfél beolvasás"), sg.Button("Export wordbe"), sg.Exit()],
]

window = sg.Window("Excelmagic", layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "Új mappa":
        new_dir() 
    if event == "Új ügyfél beolvasás":
        print(event, values)
    if event == "Export wordbe":
        print(event, values)

window.close()