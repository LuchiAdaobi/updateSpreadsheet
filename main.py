import openpyxl as xl
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.marker import DataPoint

# cell = sheet['c4']
# cell = sheet.cell(4, 3)




wb = xl.load_workbook('transactions.xlsx')

sheet = wb['Sheet1']

for row in range(2, sheet.max_row + 1):
    cell = sheet.cell(row, 3)
    corrected_price = cell.value * 0.9
    corrected_price_cell = sheet.cell(row, 4)
    corrected_price_cell.value = corrected_price

corrected_price_cell_heading = sheet.cell(1, 4)

corrected_price_cell_heading.value = 'corrected price'

values = Reference(sheet,
          min_row=2,
          max_row=sheet.max_row,
          min_col=4,
          max_col=4)

chart = BarChart()
chart.add_data(values)
sheet.add_chart(chart, 'e2')

slices = [DataPoint(idx=i) for i in range(3)]
sliceOne, sliceTwo, sliceThree = slices
sliceOne.graphicalProperties.solidFill = "02c212"
sliceTwo.graphicalProperties.solidFill = "d62b00"
sliceThree.graphicalProperties.solidFill = "f5ad31"
chart.series[0].data_points = slices



wb.save('transaction3.xlsx')

