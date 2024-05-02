import os
import openpyxl as xl
from openpyxl import load_workbook

humber = 'C:/Users/Julie/Desktop/hv/humber.xlsx'
valleywood = 'C:/Users/Julie/Desktop/hv/valleywood.xlsx'

hbook = load_workbook(filename=humber)
vbook = load_workbook(filename=valleywood)

def copysheet(book, name):
    hmain = book.worksheets[0]
    vmain = book.worksheets[1]
    hsheet = hbook.create_sheet(name)
    vsheet = vbook.create_sheet(name)

    for row in hmain:
        for cell in row:
            hsheet[cell.coordinate].value = cell.value

    for row in vmain:
        for cell in row:
            vsheet[cell.coordinate].value = cell.value

    hbook.save(humber)
    vbook.save(valleywood)

path = 'C:/Users/Julie/Desktop/hv/all'
spreadsheets = os.listdir(path)
# print(spreadsheets)

for wb in spreadsheets:
    specific = path + '/' + wb
    book = load_workbook(filename=specific)
    copysheet(book, wb)


# for file in spreadsheets:
#     file_name = f"{file}"
#     wb = load_workbook(filename=file_name)
#     copysheet(wb)

# for spreadsheet in spreadsheets:
#     copysheet(spreadsheet)

#print("hello world")
