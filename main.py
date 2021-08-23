from excel_tools._xls import XlsReader
from excel_tools._xlsx import XlsxReader
a = XlsxReader(name='道具表.xlsx')
print(a._sheet.max_column)