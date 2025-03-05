from openpyxl import load_workbook
import tkinter as tk
from tkinter import ttk
import pyperclip

# Get departments
def getDepartment():
    workbook = load_workbook(filename = './subdepartment_list.xlsx')
    sheet = workbook.active
    departments = []
    mbu_value = mbu.get()
    for row in sheet.iter_rows(min_row = 4, min_col = 5, max_col = 7,values_only = True):
        if row[0] == mbu_value:
            departments.append(row[2])
    result = ', '.join(map(str, set(departments)))
    pyperclip.copy(result)
    print(result)
    department_label.configure(text = f'The following departments were copied to the clipboard:\n{result}')

# Create window
root = tk.Tk()
root.title('MBU to Department')
root.geometry('600x500+50+50')

# Create frame
content = ttk.Frame(root)
content.grid(column = 0, row = 0)

# Create mbu frame
mbu = tk.IntVar()
mbu_label = ttk.Label(content, text = 'MBU #').grid(column = 0, row = 1, sticky = 'w')
mbu_entry = ttk.Entry(content, textvariable = mbu).grid(column = 1, row = 1, columnspan = 2, sticky = 'w')
mbu_search_button = ttk.Button(content, text = 'Search', command = getDepartment).grid(column = 3, row = 1, sticky = 'w')

# Create department frame
department_label = ttk.Label(content, text = 'Department')
department_label.grid(column = 0, row = 2, columnspan = 5, rowspan = 5, sticky = 'W')

root.mainloop()
