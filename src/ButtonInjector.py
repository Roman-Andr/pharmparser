import os
from enum import Enum

import pythoncom
import win32com.client as win32
from openpyxl.utils import get_column_letter

from Macro import Macro, ApplyFiltersMacro, RemoveFiltersMacro, SortMacro


class FilterCriteria(Enum):
    GREATER_THAN_ZERO = ">0"
    LESS_THAN_ZERO = "<0"
    GREATER_THAN_OR_EQUAL_ZERO = ">=0"
    LESS_THAN_OR_EQUAL_ZERO = "<=0"
    EQUAL_ZERO = "=0"


class SortOrder(Enum):
    ASCENDING = "xlAscending"
    DESCENDING = "xlDescending"


class ButtonInjector:
    def __init__(self, file_path, *buttons):
        self.file_path = file_path
        pythoncom.CoInitialize()
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = False
        self.workbook = None
        self.worksheet = None
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


class Button:
    def __init__(self, cell_address, caption, macro, back_color=None, fore_color=None):
        self.cell_address = cell_address
        self.caption = caption
        self.macro = macro
        self.back_color = back_color
        self.fore_color = fore_color
        self.button_name = None

    def create(self, worksheet):
        cell = worksheet.Range(self.cell_address)
        left = cell.Left
        top = cell.Top
        width = cell.Width
        height = cell.Height

        button = worksheet.Buttons().Add(left, top, width, height)
        button.Caption = self.caption
        button.OnAction = self.macro.name
        self.button_name = button.Name
        self.macro.add_position_code(self.generate_position_code(), self.restore_position_code())

    def generate_position_code(self):
        id_name = self.button_name.replace(' ', '')
        return f"""
        Dim btn{id_name} As Button
        Set btn{id_name} = ActiveSheet.Buttons("{self.button_name}")
        Dim btn{id_name}Left As Double
        Dim btn{id_name}Top As Double
        Dim btn{id_name}Width As Double
        Dim btn{id_name}Height As Double
        btn{id_name}Left = btn{id_name}.Left
        btn{id_name}Top = btn{id_name}.Top
        btn{id_name}Width = btn{id_name}.Width
        btn{id_name}Height = btn{id_name}.Height
        """

    def restore_position_code(self):
        id_name = self.button_name.replace(' ', '')
        return f"""
        btn{id_name}.Left = btn{id_name}Left
        btn{id_name}.Top = btn{id_name}Top
        btn{id_name}.Width = btn{id_name}Width
        btn{id_name}.Height = btn{id_name}Height
        """


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
