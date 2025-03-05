import os
from itertools import chain
from typing import List, Type, Tuple

from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from win32api import RGB

from utils import DataType, FilterCriteria, SortOrder, Settings
from .formatters import DataFormatter
from .formatters import BaseFormatter
from .macros import Button, ApplyFiltersMacro, SortMacro, RemoveFiltersMacro
from .macros.button_injector import ButtonInjector
from utils.utils import remove


class Spreadsheet:
    __slots__ = ["data", "settings", "formatters"]

    def __init__(self, data: DataType, settings: Settings, formatters: List[Tuple[Type[BaseFormatter], str]]):
        self.data = data
        self.settings = settings
        self.formatters = formatters

    def export(self, titles: List[str], data: DataType):
        wb = Workbook()
        wb.remove(wb.active)
        column = len(data) * 2
        target = self.settings.fileName.replace('.xlsx', '.xlsm')
        remove(target, os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        sheet_titles = []
        for formatter, title in self.formatters:
            form = formatter(self.settings, data, titles)
            sheet = wb.create_sheet(title)
            form.format(sheet)
            if isinstance(form, DataFormatter):
                sheet_titles.append(title)
        wb.save(self.settings.fileName)
        for i, sheet_name in enumerate(sheet_titles):
            ButtonInjector(i + 1, self.settings.fileName if i == 0 else f"{i - 1}{target}", [
                Button('A1', 'Apply Filters',
                       ApplyFiltersMacro(column, FilterCriteria.GREATER_THAN_ZERO, sheet_name),
                       back_color=RGB(18, 230, 89),
                       fore_color=RGB(18, 230, 89)),
                Button('A2', 'Remove Filters',
                       RemoveFiltersMacro(column, sheet_name),
                       back_color=RGB(230, 64, 18),
                       fore_color=RGB(230, 64, 18)),
                *chain(*[[Button(f'{col}1', '↑', SortMacro(col, SortOrder.DESCENDING, sheet_name)),
                          Button(f'{col}2', '↓', SortMacro(col, SortOrder.ASCENDING, sheet_name))]
                         for col in [get_column_letter(x) for x in range(4, column + 2, 2)]])
            ]).save(f"{i}{target}")
            remove(f"{i - 1}{target}")
        os.rename(f"{len(sheet_titles) - 1}{target}", target)
        remove(self.settings.fileName)
