import os

import pythoncom
import win32com.client as win32
from win32com.client import CDispatch


class ButtonInjector:
    def __init__(self, id, file_path, buttons):
        self.file_path = file_path
        pythoncom.CoInitialize()
        self.excel: CDispatch = win32.Dispatch('Excel.Application')
        self.excel.Visible = False
        self.workbook = None
        self.worksheets = None
        self.buttons = []

        self.open_workbook(id)
        self.buttons.extend(buttons)

    def open_workbook(self, id):
        self.workbook = self.excel.Workbooks.Open(os.path.abspath(self.file_path))
        self.worksheet = self.workbook.Sheets(id)

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
        sheet_name = self.worksheet.Name
        for button in self.buttons:
            button.macro.sheet_name = sheet_name
            button.create(self.worksheet)
            if self.workbook:
                module = self.workbook.VBProject.VBComponents.Add(1)
                module.CodeModule.AddFromString(button.macro.get_code().strip())
