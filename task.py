import string
import xlrd
import cursor
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def num_to_excel_col(n):
    if n < 1:
        raise ValueError("Number must be positive")
    result = ""
    while True:
        if n > 26:
            n, r = divmod(n - 1, 26)
            result = chr(r + ord('A')) + result
        else:
            return chr(n + ord('A') - 1) + result


def excel_col_to_num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


def print_data(data):
    print("\n")

    max_length = [0]
    for idx in range(len(data[0])):
        max_length.append(0)

    for row in data:
        idx = 0
        for item in row:
            idx = idx + 1
            if len(str(item)) > max_length[idx]:
                max_length[idx] = len(str(item))

    print('    ', end='')
    for idx in range(len(data[0])):
        add_spaces = '   '
        for i in range(max_length[idx + 1] - len(str(num_to_excel_col(idx + 1)))):
            add_spaces = add_spaces + ' '
        print(num_to_excel_col(idx + 1), end=add_spaces)
    print("\n")

    row_idx = 0
    for row in data:
        row_idx = row_idx + 1
        print(row_idx, end='   ')
        idx = 0
        for item in row:
            idx = idx + 1
            add_spaces = '   '
            for i in range(max_length[idx] - len(str(item))):
                add_spaces = add_spaces + ' '
            print(item, end=add_spaces)
        print("")


def repace_data(data, old_string, new_string):
    for row in range(len(data)):
        for item in range(len(data[0])):
            data[row][item] = str(data[row][item]).replace(old_string, new_string)
    return data


def swap_data(data, first_cell, second_cell):
    col1 = first_cell[0]
    row1 = 0
    col2 = second_cell[0]
    row2 = 0

    for i in range(1, len(first_cell)):
        if first_cell[i].isdigit():
            row1 = int(first_cell[i:])
        else:
            col1 = col1+first_cell[i]

    for i in range(1, len(second_cell)):
        if second_cell[i].isdigit():
            row2 = int(second_cell[i:])
        else:
            col2 = col2+second_cell[i]

    tmp = data[row1-1][excel_col_to_num(col1)-1]
    data[row1 - 1][excel_col_to_num(col1) - 1] = data[row2-1][excel_col_to_num(col2)-1]
    data[row2 - 1][excel_col_to_num(col2) - 1] = tmp

    return data


Tk().withdraw()
filename = askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
print('\nYou loaded: ', filename)

wb = xlrd.open_workbook(filename)
sheet = wb.sheet_by_index(0)
all_data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]


choice = ''
while choice is not '0':
    print("\n1] Print data")
    print("2] Remove dublicate rows")
    print("3] Swap cells")
    print("4] Replace string")
    print("0] Exit")
    cursor.show()
    choice = str(input())
    cursor.hide()

    if choice == '1':
        print_data(all_data)

    if choice == '2':
        data_no_dublicates = []
        for elem in all_data:
            if elem not in data_no_dublicates:
                data_no_dublicates.append(elem)
        all_data = data_no_dublicates
        print_data(all_data)

    if choice == '3':
        cell_one = str(input("Enter cell one (example A1): "))
        cell_two = str(input("Enter cell two (example A2): "))
        data_replaced = swap_data(all_data, cell_one, cell_two)
        print_data(data_replaced)

    if choice == '4':
        old = str(input("Enter string to be replaced: "))
        new = str(input("Enter string to be replace with: "))
        data_replaced = repace_data(all_data, old, new)
        print_data(data_replaced)


# # Remove dublicates
# data_no_dublicates = []
# for elem in all_data:
#     if elem not in data_no_dublicates:
#         data_no_dublicates.append(elem)
# print_data(data_no_dublicates)
#
# # Replace some data
# data_replaced = repace_data(data_no_dublicates, 'Ivanov', 'Dimitrov')
# print_data(data_replaced)
#
# # Swap some data
# data_swapped = swap_data(data_replaced, 'A3', 'C4')
# print_data(data_swapped)
