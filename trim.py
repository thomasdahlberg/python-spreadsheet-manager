import openpyxl
import os
import sys

file_name = sys.argv[1]
sheet_name = 'sheet1'
sheet = None
wb = None

if file_name:
    wb = openpyxl.load_workbook(filename=f"{file_name}.xlsx")

if sheet_name:
    sheet = wb[sheet_name]

sheet.delete_rows(idx=1, amount=1)

data = sheet.rows

csv_file = open(f"{file_name}.csv", "w+")

for row in data:
    l = list(row)
    for i in range(len(l)):
        if i == len(l) - 1:
            csv_file.write(str(l[i].value))
        else:
            csv_file.write(str(l[i].value) + ',')
    csv_file.write('\n')

csv_file.close()
wb.save(filename=f"{file_name}.xlsx")
