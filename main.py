from excel_tools._xls import XlsReader
from excel_tools._xlsx import XlsxReader
# TODO: 统一不同模块对excel的获取索引规范
a = XlsReader(name='新建.xls')
print(a.get_cell(1, 1))


b = XlsxReader(name='新建.xlsx')
print(b.get_row(1))