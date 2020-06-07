#! python3
# fileNamesToExcel.py - takes the file names in a folder and outputs them
# in an excel file

import openpyxl, os

#TODO: get all filenames in directory in a list
full_names = os.listdir()
folder = os.path.basename(os.getcwd())
wb = openpyxl.Workbook()
sheet = wb.active
split_names = []
#todo: split filenames along hyphen
for name in full_names:
    if name.endswith('.pdf'):
        temp = name.rstrip('.pdf').split('-')
        split_names.append(temp)
#todo: put each piece of filename in a cell in a row
for r, row in enumerate(split_names):
    for c, col in enumerate(split_names[r]):
        sheet.cell(column=c+1,row=r+1).value = split_names[r][c]


wb.save(folder+'.xlsx')