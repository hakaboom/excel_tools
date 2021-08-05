from excel_tools._xls import XlsReader
from excel_tools._xlsx import XlsxReader
a = XlsReader(name='新建.xls')
print(a.get_row_value(2))
