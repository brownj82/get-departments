from openpyxl import load_workbook

file = '/mnt/c/Users/brownj82/my_stuff/nonsensitive_files/clients/helpful-stuff/subdepartment_list.xlsx'

workbook = load_workbook(filename=file)
sheet = workbook.active

departments = []
mbu = int(input('MBU: '))
mbu_column_number = int(input('MBU Column Number: '))
department_column_number = int(input('Department Column Number: '))

for row in sheet.iter_rows(min_row = mbu_column_number - 1, min_col = mbu_column_number, max_col = department_column_number,values_only=True):
    if row[0] == mbu:
        departments.append(row[2])

print(set(departments))
