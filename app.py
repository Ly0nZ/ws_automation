import openpyxl as oxl
from pathlib import Path

def update_workbook(filename):
    workbook = oxl.load_workbook(filename)
    sheet = workbook['Sheet1']

    for row in range(3, sheet.max_row + 1):
        cell = sheet.cell(row, 3)
        commission = cell.value * 0.1

        new_cell = sheet.cell(row, 4)
        new_cell.value = commission
    
    values = Reference(sheet,
                       min_row=3,
                       max_row=sheet.max_row,
                       min_col=4,
                       max_col=4)
    
    chart = BarChart()
    chart.add_data(values)
    sheet.add_chart(chart, 'F3')
    workbook.save(filename)

path = Path("sales")
for file in path.glob('*.xlsx'):
    update_workbook(file)
