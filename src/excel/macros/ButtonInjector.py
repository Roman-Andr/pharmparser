import os
import shutil
from itertools import chain

import pythoncom
import win32com.client as win32
from openpyxl.utils import get_column_letter
from win32api import RGB
from win32com.client import CDispatch

from src.utils import FilterCriteria
from src.utils import SortOrder
from . import *


class ButtonInjector:
    def __init__(self, file_path, buttons):
        self.file_path = file_path
        pythoncom.CoInitialize()
        self.excel: CDispatch = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = False
        self.workbook = None
        self.worksheets = None
        self.buttons = []

        self.open_workbook()
        self.buttons.extend(buttons)

    def open_workbook(self):
        self.workbook = self.excel.Workbooks.Open(os.path.abspath(self.file_path))
        self.worksheet = self.workbook.Sheets(1)

    def close_workbook(self):
        if self.workbook:
            self.workbook.Close()
        self.excel.Quit()

    def save(self, new_file_path):
        self.generate_vba_code()
        if self.workbook:
            self.workbook.SaveAs(os.path.abspath(new_file_path), FileFormat=52)
        self.close_workbook()

    def generate_vba_code(self):
        for button in self.buttons:
            button.create(self.worksheet)
            if self.workbook:
                module = self.workbook.VBProject.VBComponents.Add(1)
                module.CodeModule.AddFromString(button.macro.get_code().strip())


def run(column):
    file_path = 'data.xlsx'
    target = 'data.xlsm'

    cache_dir = os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py')
    if os.path.exists(cache_dir):
        shutil.rmtree(cache_dir)

    if os.path.exists(target):
        os.remove(target)

    Macro.start_row, Macro.end_row = 3, 100000
    end_column = column

    buttons = [
        Button('A1', 'Apply Filters',
               ApplyFiltersMacro(end_column, FilterCriteria.GREATER_THAN_ZERO),
               back_color=RGB(18, 230, 89),
               fore_color=RGB(18, 230, 89)),
        Button('A2', 'Remove Filters',
               RemoveFiltersMacro(end_column),
               back_color=RGB(230, 64, 18),
               fore_color=RGB(230, 64, 18)),
        *chain(*[[Button(f'{col}1', '↑', SortMacro(col, SortOrder.DESCENDING)),
                  Button(f'{col}2', '↓', SortMacro(col, SortOrder.ASCENDING))]
                 for col in [get_column_letter(x) for x in range(4, end_column + 2, 2)]])
    ]

    injector = ButtonInjector(file_path, buttons)
    injector.save(target)

    if os.path.exists(file_path):
        os.remove(file_path)
