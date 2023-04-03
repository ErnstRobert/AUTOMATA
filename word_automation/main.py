import PySimpleGUI as sg


layout = [
    [sg.Button("Új ügyfél mappa")],
    [sg.Text("Cégnév:"), sg.Input(key="cegnev", do_not_clear=False)],
    [sg.Button("Összesítő beolvasás"), sg.Button("Export wordbe"), sg.Exit()],
]

window = sg.Window("Excelmagic", layout)

#def validate(values):
    is_valid = True
    values_invalid = []

    if len(values["cegnev"]) == 0:
        values_invalid.append("cegnev")
        is_valid = False

    result = [is_valid, values_invalid]

    return result

#def gen_error_message(values_invalid):
    error_message = ""
    for value_invalid in values_invalid:
        error_message = ('\nInvalid' + ":" + value_invalid)
    return

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "Új ügyfél mappa":
        print(event, values)
    if event == "Összesítő beolvasás":
        print(event, values)
    if event == "Export wordbe":
        print(event, values)

window.close()