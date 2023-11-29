import openpyxl


book = openpyxl.open('data_owners.xlsx', read_only = True)

sheet = book.active

new_book = openpyxl.Workbook()

new_sheet = new_book.active
new_row = 1
for row in range(2, 5):
    number_room = sheet[row][0].value
    fio = sheet[row][1].value
    room_area = sheet[row][2].value
    variant_own = sheet[row][3].value
    part_own = sheet[row][4].value
    document_number = sheet[row][5].value
    new_sheet[f'A{new_row}'].value = number_room
    new_sheet[f'B{new_row}'].value = fio
    new_sheet[f'C{new_row}'].value = float(room_area)
    new_sheet[f'D{new_row}'].value = variant_own
    if type(part_own) == str:
        a, b = part_own.split("/")
        part_own = int(a) / int(b)
    new_sheet[f'E{new_row}'].value = part_own * float(room_area)
    new_sheet[f'F{new_row}'].value = document_number
    new_row += 1
    new_book.save(f'protocol{new_row}.xlsx')
    new_book.close()






# def print_hi(name):
#     print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.
#
#
# # Press the green button in the gutter to run the script.
# if __name__ == '__main__':
#     print_hi('PyCharm')


