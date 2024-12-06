import os
from abc import ABC, abstractmethod
from enum import Enum

import pythoncom
from openpyxl.utils import get_column_letter

import win32com.client as win32


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


class Macro(ABC):
    start_row = 3
    end_row = 100

    def __init__(self, name, code_template):
        self.name = name
        self.code_template = code_template
        self.position_codes = []

    def add_position_code(self, position_code, restore_code):
        self.position_codes.append((position_code, restore_code))

    @abstractmethod
    def get_code(self):
        pass


class ApplyFiltersMacro(Macro):
    def __init__(self, start_column, end_column, criteria):
        self.start_column = start_column
        self.end_column = end_column
        self.criteria = criteria
        code_template = """
        Sub ApplyFilters()
            Application.ScreenUpdating = False
            {position_code_block}
            ActiveSheet.AutoFilterMode = False
            Dim col As Integer
            For col = {start_column} To {end_column} Step 2
                ActiveSheet.Range("{data_range}").AutoFilter Field:=col, Criteria1:="{criteria}"
            Next col
            {restore_code_block}
            Application.ScreenUpdating = True
        End Sub
        """
        super().__init__('ApplyFilters', code_template)

    def get_code(self):
        position_code_block = "\n".join([code for code, _ in self.position_codes])
        restore_code_block = "\n".join([code for _, code in self.position_codes])
        data_range = f"A{Macro.start_row}:T{Macro.end_row}"
        return self.code_template.format(
            position_code_block=position_code_block,
            restore_code_block=restore_code_block,
            start_column=self.start_column,
            end_column=self.end_column,
            criteria=self.criteria.value,
            data_range=data_range
        )


class RemoveFiltersMacro(Macro):
    def __init__(self, start_column, end_column):
        self.start_column = start_column
        self.end_column = end_column
        code_template = """
        Sub RemoveFilters()
            Application.ScreenUpdating = False
            {position_code_block}
            If ActiveSheet.AutoFilterMode Then
                Dim col As Integer
                For col = {start_column} To {end_column} Step 2
                    ActiveSheet.Range("{data_range}").AutoFilter Field:=col
                Next col
            End If
            ActiveSheet.Range("{data_range}").Sort Key1:=ActiveSheet.Columns("A"), Order1:=xlAscending, Header:=xlYes
            {restore_code_block}
            Application.ScreenUpdating = True
        End Sub
        """
        super().__init__('RemoveFilters', code_template)

    def get_code(self):
        position_code_block = "\n".join([code for code, _ in self.position_codes])
        restore_code_block = "\n".join([code for _, code in self.position_codes])
        data_range = f"A{Macro.start_row}:T{Macro.end_row}"
        return self.code_template.format(
            position_code_block=position_code_block,
            restore_code_block=restore_code_block,
            start_column=self.start_column,
            end_column=self.end_column,
            data_range=data_range
        )


class SortMacro(Macro):
    def __init__(self, column, sort_order):
        self.column = column
        self.sort_order = sort_order
        code_template = """
        Sub Sort{sort_name}{column}()
            Application.ScreenUpdating = False
            {position_code_block}
            ActiveSheet.Range("{data_range}").Sort Key1:=ActiveSheet.Columns("{column}"), Order1:={sort_order}, Header:=xlYes
            {restore_code_block}
            Application.ScreenUpdating = True
        End Sub
        """
        super().__init__(f'Sort{sort_order.name}{column}', code_template)

    def get_code(self):
        position_code_block = "\n".join([code for code, _ in self.position_codes])
        restore_code_block = "\n".join([code for _, code in self.position_codes])
        data_range = f"A{Macro.start_row}:T{Macro.end_row}"
        return self.code_template.format(
            position_code_block=position_code_block,
            restore_code_block=restore_code_block,
            column=self.column,
            sort_order=self.sort_order.value,
            sort_name=self.sort_order.name,
            data_range=data_range
        )


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

        # if self.back_color:
        # button.ShapeRange.Fill.ForeColor.RGB = self.back_color
        # if self.fore_color:
        #     button.ShapeRange.TextFrame.Characters().Font.Color = self.fore_color

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
               ApplyFiltersMacro(4, end_column, FilterCriteria.GREATER_THAN_ZERO),
               back_color=0x19CF1F,
               fore_color=0x19CF1F),
        Button('A2', 'Remove Filters',
               RemoveFiltersMacro(4, end_column),
               back_color=0xE81737,
               fore_color=0xE81737)
    ]

    columns = [get_column_letter(x) for x in range(4, end_column + 2, 2)]
    for col in columns:
        buttons.append(Button(f'{col}2', '↓',
                              SortMacro(col, SortOrder.ASCENDING),
                              back_color=0x19CF1F,
                              fore_color=0x19CF1F))
        buttons.append(Button(f'{col}1', '↑',
                              SortMacro(col, SortOrder.DESCENDING),
                              back_color=0xE81737,
                              fore_color=0xE81737))

    injector = ButtonInjector(file_path, *buttons)
    injector.save(target)
