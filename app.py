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
    
    workbook.save(filename)

path = Path("sales")
for file in path.glob('*.xlsx'):
    process_workbook(file)
