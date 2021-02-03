import sys
import csv

IMPOSSIBLE_DATES = [
    ['2/29', '2/28'],
    ['2/30', '2/28'],
    ['2/31', '2/28'],
    ['4/31', '4/30'],
    ['6/31', '6/30'],
    ['9/31', '9/30'],
    ['11/31', '11/30'],
]


# def has_num(str):
#     return any(i.isdigit() for i in str)


# def has_letter(str):
#     lowercase_str = str.lower()
#     return lowercase_str.islower()


file_name = sys.argv[1]

tsv_file = open(file_name)

read_tsv = csv.reader(tsv_file, delimiter="\t")

row_cache = []

for row in read_tsv:
    row_cache.append(row)

for row in row_cache:
    for cell in range(len(row)):
        if "/" in row[cell]:
            for date in IMPOSSIBLE_DATES:
                row[cell] = row[cell].replace(date[0], date[1])
    # print(row)
csv_file = open(f"{file_name}", "w+")

for row in row_cache:
    for i in range(len(row)):
        if i == len(row) - 1:
            csv_file.write(str(row[i]))
        else:
            csv_file.write(str(row[i]) + '\t')
    csv_file.write('\n')

csv_file.close()

print('done!')
