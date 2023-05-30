import PySimpleGUI as Py
from converter import exceltopdf

label1 = Py.Text("Select Excel file to be converted:")
input1 = Py.Input()
choose_button1 = Py.FilesBrowse("Choose", key="files")
label2 = Py.Text("Select destination folder:")
input2 = Py.Input()
choose_button2 = Py.FolderBrowse("Choose", key="folder")

Convert_button = Py.Button("Convert")
output = Py.Text(key="output")
window = Py.Window("Excel to PDf",
                   layout=[[label1, input1, choose_button1],
                           [label2, input2, choose_button2],
                           [Convert_button, output]])
while True:
    event, values = window.read()
    filePath = values['files'].split(";")
    destination = values['folder']
    if Convert_button:
        pdf = exceltopdf(filePath, destination)
        if pdf:
            window["output"].update(value="Conversion Completed")
        else:
            window["output"].update(value="Please select Excel files", text_color='Red')
    elif event == Py.WIN_CLOSED:
        break

window.close()
