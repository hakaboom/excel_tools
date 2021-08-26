# -*- coding: utf-8 -*-
import xlrd
from xlrd.sheet import Cell as XLCell

import os
from typing import Optional, Union, Generator, List, Any

from excel_tools.read.base_reader import ExcelBaseObject
from excel_tools.utils import Cell


class XlsReader(ExcelBaseObject):
    def __init__(self, sheet: Union[int, str, None] = 0, *args, **kwargs):
        """
        xls读取类,仅可以读取 .xls格式的表格
        注意：所有对表的索引都从1,1开始。sheet:A1=(row=1,col=1)

        Args:
            name(str): 文件名
            path(str): 文件路径
            sheet(int,str): 需要读取的工作表(表名或表索引)
        """
        super(XlsReader, self).__init__(*args, **kwargs)
        self.set_sheet(sheet)

    def open_xl(self):
        self._xl = xlrd.open_workbook(os.path.join(self._path, self._name), formatting_info=True)

    def set_sheet(self, sheet: Union[int, str, None] = 0) -> None:
        if isinstance(sheet, str):
            self._sheet = self._xl.sheet_by_name(sheet)
        elif isinstance(sheet, int):
            self._sheet = self._xl.sheet_by_index(sheet)
        else:
            self._sheet = self._xl.sheet_by_index(0)

    @property
    def sheet_name(self) -> str:
        """
        当前读取中的工作表名称

        Returns:
            工作表名称
        """
        return self._sheet.name

    @property
    def max_row(self) -> int:
        if not hasattr(self, '_max_row'):
            setattr(self, '_max_row',  self._sheet.nrows)

        return getattr(self, '_max_row')

    @max_row.setter
    def max_row(self, rowx: int) -> None:
        setattr(self, '_max_row', rowx)

    @property
    def max_column(self) -> int:
        if not hasattr(self, '_max_col'):
            setattr(self, '_max_col',  self._sheet.ncols)

        return getattr(self, '_max_col')

    @max_column.setter
    def max_column(self, colx: int) -> None:
        setattr(self, '_max_col', colx)

    def get_rows(self) -> Generator[List[XLCell], Any, None]:
        return self._sheet.get_rows()

    def get_row(self, rowx: int) -> List[XLCell]:
        rowx = rowx - 1

        return [Cell(row=rowx, column=col + 1, worksheet=self._sheet, value=v.value, ctype=v.ctype)
                for col, v in enumerate(self._sheet.row(rowx)) if col < self.max_column]

    def get_row_len(self, rowx: int) -> int:
        return len(self.get_row(rowx=rowx))

    def get_row_value(self, rowx: int, start_colx: Optional[int] = 0, end_colx: int = None) -> List[Any]:
        return [v.value for v in self.get_row(rowx=rowx)][start_colx:end_colx]

    def get_col(self, colx: int, start_rowx: Optional[int] = 0, end_rowx: Optional[int] = None) -> List[Cell]:
        return [Cell(worksheet=self._sheet, row=row + 1 + start_rowx, column=colx)
                for row, value in enumerate(self.get_col_value(colx, start_rowx, end_rowx))]

    def get_col_value(self, colx: int, start_rowx: Optional[int] = 0, end_rowx: Optional[int] = None) -> List[Any]:
        return self._sheet.col_values(colx - 1, start_rowx, end_rowx)

    def get_cell(self, rowx: int, colx: int) -> Cell:
        cell = self._sheet.cell(rowx - 1, colx - 1)
        return Cell(worksheet=self._sheet, row=rowx, column=colx, value=cell.value, ctype=cell.ctype)

    def get_cell_value(self, rowx: int, colx: int) -> Any:
        return self.get_cell(rowx, colx).value