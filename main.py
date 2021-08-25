from excel_tools._xls import XlsReader
from excel_tools._xlsx import XlsxReader
a = XlsxReader(name='道具表.xlsx')
a.max_column = 11

print(a.get_row_value(1))