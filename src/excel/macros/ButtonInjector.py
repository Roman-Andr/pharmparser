import os

import pythoncom
import win32com.client as win32
from win32com.client import CDispatch


class ButtonInjector:
    def __init__(self, file_path, buttons):
        self.file_path = file_path
        pythoncom.CoInitialize()
        self.excel: CDispatch = win32.Dispatch('Excel.Application')
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
