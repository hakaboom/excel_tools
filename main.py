from excel_tools._xls import XlsReader
from excel_tools._xlsx import XlsxReader

import os
from typing import Union
WORK_PATH = 'C:\\Users\\Administrator\\Desktop\\敏感词替换'
SensitiveWords_Path = 'C:\\Users\\Administrator\\Desktop\\敏感词.xls'
IGNORE_LINES = 5


def read_head(xl: Union[XlsReader, XlsxReader]):
    """
    读取表格表头

    Returns:
        表名, 表格列数
    """
    ret = xl.get_row_value(1, 0, 2)
    return ret[0], int(ret[1])


def read_end_tag(xl: Union[XlsReader, XlsxReader]):
    """
    读取表格#END_TAG#标记

    Returns:
        #END_TAG# 所在的行数
    """
    col_value_list = xl.get_col_value(1)
    if index := col_value_list.index('#END_TAG#'):
        return index
    raise ValueError()


class LoadAllExcel(object):
    def __init__(self):
        self.xls = {}
        for root, dirs, files in os.walk(WORK_PATH):
            for file in files:
                if os.path.splitext(file)[1] == '.xls':
                    self.xls[file] = XlsReader(path=WORK_PATH, name=file)
                elif os.path.splitext(file)[1] == '.xlsx':
                    self.xls[file] = XlsxReader(name=file, path=WORK_PATH)


def LoadBlackStr():
    """
    读取获取屏蔽字表格

    Returns:
        屏蔽字列表
    """
    xl = XlsReader(name=SensitiveWords_Path)
    return xl.get_col_value(1, 1, read_end_tag(xl))


AllExcel = LoadAllExcel()
SensitiveWords = LoadBlackStr()


for excel_name, excel_work in AllExcel.xls.items():
    _, col_len = read_head(excel_work)
    row_len = read_end_tag(excel_work)
    # print(f'{excel_name}, 最大列数={col_len}, 最大行数={row_len}')
    # 遍历表中所有元素
    for col_index in range(col_len):
        for row_index in range(IGNORE_LINES, row_len):
            cell = excel_work.get_cell(rowx=row_index + 1, cols=col_index + 1)
            for word in SensitiveWords:
                if word in str(cell.value):
                    print(f'{excel_name}, {cell}, 触发屏蔽字:{word}')
