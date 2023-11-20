# import openpyxl as xl
#
# wb = xl.load_workbook('transactions.xlsx')
#
# sheet = wb['Sheet1']
#
# cell = sheet['c4']
#
# cell = sheet.cell(4, 3)
#
#
# for row in range(2, sheet.max_row + 1):
#     cell = sheet.cell(row, 3)
#     corrected_price = cell.value * 0.9
#     corrected_price_cell = sheet.cell(row, 4)
#     corrected_price_cell.value = corrected_price
#
# corrected_price_cell_heading = sheet.cell(1, 4)
# corrected_price_cell_heading.value = 'Corrected price'
# print(corrected_price_cell_heading.value)
#
#
# wb.save('transactions2.xlsx')

import openpyxl as xl

wb = xl.load_workbook('transactions.xlsx')

sheet = wb['Sheet1']

for row in range(2, sheet.max_row + 1):
    cell = sheet.cell(row, 3)
    corrected_price = cell.value * 0.9
    corrected_price_cell = sheet.cell(row, 4)
    corrected_price_cell.value = corrected_price
    print(corrected_price_cell.value)

corrected_price_cell_heading = sheet.cell(1, 4)

corrected_price_cell_heading.value = 'corrected price'

wb.save('transaction3.xlsx')

