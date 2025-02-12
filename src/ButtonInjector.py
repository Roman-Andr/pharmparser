import os
from typing import List

import pythoncom
import win32com.client as win32
from openpyxl.utils import get_column_letter
from win32com.client import CDispatch

from Button import Button
from FilterCriteria import FilterCriteria
from Macro import Macro, ApplyFiltersMacro, RemoveFiltersMacro, SortMacro
from SortOrder import SortOrder


class ButtonInjector:
    def __init__(self, file_path, *buttons):
        self.file_path = file_path
        pythoncom.CoInitialize()
        self.excel: CDispatch = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = False
        self.workbook = None
        self.worksheets = None
        self.buttons = []

        self.open_workbook()
        self.add_buttons(*buttons)

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

    def add_buttons(self, *buttons):
        for button in buttons:
            self.buttons.append(button)

    def generate_vba_code(self):
        for button in self.buttons:
            button.create(self.worksheet)
            if self.workbook:
                module = self.workbook.VBProject.VBComponents.Add(1)
                module.CodeModule.AddFromString(button.macro.get_code().strip())


def run(column):
    file_path = 'data.xlsx'
    target = 'data.xlsm'

    if os.path.exists(target):
        os.remove(target)

    Macro.start_row, Macro.end_row = 3, 100000
    end_column = column

    buttons = [
        Button('A1', 'Apply Filters',
               ApplyFiltersMacro(end_column, FilterCriteria.GREATER_THAN_ZERO)),
        Button('A2', 'Remove Filters',
               RemoveFiltersMacro(end_column))
    ]
    columns = [get_column_letter(x) for x in range(4, end_column + 2, 2)]
    for col in columns:
        buttons.append(Button(f'{col}2', '↓',
                              SortMacro(col, SortOrder.ASCENDING)))
        buttons.append(Button(f'{col}1', '↑',
                              SortMacro(col, SortOrder.DESCENDING)))

    injector = ButtonInjector(file_path, *buttons)
    injector.save(target)
