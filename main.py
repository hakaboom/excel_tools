from excel_tools._xls import XlsReader
from excel_tools._xlsx import XlsxReader
from excel_tools import read_xl

a = read_xl(name='新建.xlsx')
a.max_column = 5

b = read_xl(name='新建.xls')
b.max_column = 5
# a.max_row = 11


print(a.get_row_value(1, start_colx=1, end_colx=2))
print(b.get_row_value(1, start_colx=1, end_colx=2))
#
# for v in a._sheet.row_slice(1 - 1):
#     print(v.ctype)