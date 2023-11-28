import openpyxl

book = openpyxl.open('data_owners.xlsx', read_only = True)

sheet = book.active

for row in range(1, sheet.max_row + 1):
    number_room = sheet[row][0].value
    fio = sheet[row][1].value
    room_area = sheet[row][2].value
    variant_own = sheet[row][3].value
    part_own = sheet[row][4].value
    document_number = sheet[row][5].value
    print(number_room, fio, room_area, variant_own, part_own, document_number)







# def print_hi(name):
#     print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.
#
#
# # Press the green button in the gutter to run the script.
# if __name__ == '__main__':
#     print_hi('PyCharm')


