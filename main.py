from openpyxl.workbook import Workbook
from openpyxl import load_workbook


# wb = Workbook()
wb = load_workbook('cell.xlsx')
ws = wb.active
#grab a specific cell
# print(f'{ws['A3'].value} : {ws['B3'].value}')


#grab a specific row
# column_a = ws['3']
# for cell in column_a:
#     print(f'{cell.value}\n')

#grab a range
range = ws['C6':'D23']
for cell in range:
    for x in cell:
        print(f'{x.value}')

