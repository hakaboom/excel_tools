from excel_tools._xls import XlsReader
from excel_tools._xlsx import XlsxReader
from excel_tools import read_xl

a = read_xl(name='新建.xlsx')
a.max_column = 6

b = read_xl(name='新建.xls')
b.max_column = 5
# a.max_row = 11


for cell in a.get_row(1, start_colx=1):
    print(cell.column, cell.row, cell.value)

for cell in b.get_row(1, start_colx=1):
    print(cell.column, cell.row, cell.value)
