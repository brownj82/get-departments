from openpyxl import load_workbook

file = '/mnt/c/Users/brownj82/my_stuff/nonsensitive_files/clients/helpful-stuff/subdepartment_list.xlsx'

workbook = load_workbook(filename=file)
sheet = workbook.active

departments = []
mbu = int(input('mbu:'))

for row in sheet.iter_rows(min_row=4, min_col=5, max_col=7,values_only=True):
    if row[0] == mbu:
        departments.append(row[2])

print(set(departments))
