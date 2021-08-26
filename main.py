from excel_tools.write._xlsx import XlsxWrite
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, NamedStyle

__all__ = ['xlStyle']


xlStyle = NamedStyle(name='xlStyle')
_thin = Side(border_style='thin', color='ffffff')
xlStyle.border = Border(left=_thin, top=_thin, right=_thin, bottom=_thin)
xlStyle.alignment = Alignment(horizontal='center', vertical='center')


a = XlsxWrite(name='new.xlsx')
a.write_value_to_row([1, 2, 3], 1, start_colx=2)
a.save()
