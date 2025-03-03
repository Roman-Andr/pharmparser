from typing import List, Type

from openpyxl.workbook import Workbook

from .formatters import SheetFormatter
from .macros.ButtonInjector import run
from ..utils import DataType
from ..utils import Settings


class Spreadsheet:
    __slots__ = ["data", "settings", "formatters"]

    def __init__(self, data: DataType, settings: Settings, formatters: List[Type[SheetFormatter]]):
        self.data = data
        self.settings = settings
        self.formatters = formatters

    def export(self, titles: List[str], data: DataType):
        wb = Workbook()
        wb.remove(wb.active)
        for formatter in self.formatters:
            form = formatter(self.settings, data, titles)
            form.format(wb.create_sheet(form.title))
        wb.save(self.settings.fileName)
        run(len(data) * 2)
