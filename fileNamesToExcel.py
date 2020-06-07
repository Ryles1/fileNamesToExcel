#! python3
# fileNamesToExcel.py - takes the file names in a folder and outputs them
# in an excel file

import openpyxl, os

#let user pick file extension

ext = input('Enter file extension to search: ')
ext = '.' + ext

# get all filenames in directory in a list
full_names = os.listdir()
folder = os.path.basename(os.getcwd())
wb = openpyxl.Workbook()
sheet = wb.active
split_names = []
#: split filenames along hyphen
for name in full_names:
    if name.endswith(ext):
        temp = name.rstrip(ext).split('-')
        split_names.append(temp)
# put each piece of filename in a cell in a row
for r, row in enumerate(split_names):
    for c, col in enumerate(split_names[r]):
        sheet.cell(column=c+1,row=r+1).value = split_names[r][c]


wb.save(folder+'.xlsx')