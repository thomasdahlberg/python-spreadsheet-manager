import csv
import os
import sys
import openpyxl


def get_excel_doc():
    file_obj = {}
    os.chdir('./work')
    files = os.listdir()
    file_map = {}
    count = 1
    for file in files:
        if file[-5:] == '.xlsx':
            file_map[count] = file
            count += 1
    doc_range = range(1, len(file_map)+1)
    print(f"Excel Documents: {file_map}")
    get_doc = input("Pick document number: ")
    while int(get_doc) not in doc_range:
        print("Not a valid document")
        get_doc = input("Pick document number: ")
    file_obj["wb"] = openpyxl.load_workbook(f"{file_map[int(get_doc)]}")
    file_obj["name"] = f"{file_map[int(get_doc)]}"
    return file_obj
