import PySimpleGUI as sg
from PyPDF2 import PdfFileReader
import os.path
import xlsxwriter

inputPath = None
outputPath = "C://"

l = ["Please click Browse to choose the Path and then Import to Load the variables.."]

layout = [
    [sg.Text("Please select PDF Form")],
    [
        sg.In(size=(25, 1), enable_events=True, key="-INPUT-"),
        sg.FileBrowse(target="-INPUT-"),
        sg.Button("Import Form"),
    ],
    [sg.Text("List of variables in the form")],
    [sg.Listbox(l, size=(70, 10), key="listBox")],
    [sg.Text("Please select Excel file to Export")],
    [
        sg.In(size=(25, 1), enable_events=True, key="-OUTPUT-"),
        sg.FileBrowse(),
        sg.Button("Export Data"),
    ],
]

window = sg.Window("PDF Form to Excel", layout).Finalize()

while True:
    event, values = window.read()
    if event == "-INPUT-":
        inputPath = values["-INPUT-"]

    if event == "Import Form":
        l = []
        pdf_reader = PdfFileReader(open(inputPath, "rb"))
        dictionary = pdf_reader.getFormTextFields()  # returns a python dictionary
        for key in dictionary:
            newVal = [key + " : " + str(dictionary[key])]
            l.append(newVal)

        window.Element("listBox").Update(values=l)

    if event == "-OUTPUT-":
        outputPath = values["-OUTPUT-"]

    if event == "Export Data":

        workbook = xlsxwriter.Workbook("exportTest.xlsx")
        worksheet = workbook.add_worksheet()
        row = 0

        for key in dictionary:
            keyToWrite = str(key)
            valueToWrite = str(dictionary[key])
            worksheet.write(1, row, keyToWrite)
            worksheet.write(2, row, valueToWrite)
            row = row + 1
        workbook.close()

    if event == sg.WIN_CLOSED:
        break


window.close()
