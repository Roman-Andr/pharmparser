from abc import ABC

from openpyxl.utils import get_column_letter


class Macro(ABC):
    start_col = 4
    start_row = 3
    end_row = 100000

    def __init__(self, name, end_column):
        self.name = name
        self.data_range = f"A{Macro.start_row}:{get_column_letter(end_column)}{Macro.end_row}"
        self.position_codes = []

    def add_position_code(self, position_code, restore_code):
        self.position_codes.append((position_code, restore_code))


class SortMacro(Macro):
    def __init__(self, column, end_column, sort_order, sheet_name):
        self.column = column
        self.sort_order = sort_order
        super().__init__(f'Sort{sort_order.name}{column}_{sheet_name}', end_column)
        self.code = f"""
        Sub Sort{sort_order.name}{column}_{sheet_name}()
            Application.ScreenUpdating = False
            {"\n".join([code for code, _ in self.position_codes])}
            ActiveSheet.Range("{self.data_range}").Sort Key1:=ActiveSheet.Columns("{self.column}"), Order1:={sort_order.value}, Header:=xlYes
            {"\n".join([code for _, code in self.position_codes])}
            Application.ScreenUpdating = True
        End Sub
        """


class RemoveFiltersMacro(Macro):
    def __init__(self, end_column, sheet_name):
        self.start_column = Macro.start_col
        super().__init__(f'RemoveFilters_{sheet_name}', end_column)
        self.code = f"""
        Sub RemoveFilters_{sheet_name}()
            Application.ScreenUpdating = False
            {"\n".join([code for code, _ in self.position_codes])}
            If ActiveSheet.AutoFilterMode Then
                Dim col As Integer
                For col = {self.start_column} To {end_column} Step 2
                    ActiveSheet.Range("{self.data_range}").AutoFilter Field:=col
                Next col
            End If
            ActiveSheet.Range("{self.data_range}").Sort Key1:=ActiveSheet.Columns("A"), Order1:=xlAscending, Header:=xlYes
            {"\n".join([code for _, code in self.position_codes])}
            Application.ScreenUpdating = True
        End Sub
        """


class ApplyFiltersMacro(Macro):
    def __init__(self, end_column, criteria, sheet_name):
        self.start_column = Macro.start_col
        self.criteria = criteria
        super().__init__(f'ApplyFilters_{sheet_name}', end_column)
        self.code = f"""
        Sub ApplyFilters_{sheet_name}()
            Application.ScreenUpdating = False
            {"\n".join([code for code, _ in self.position_codes])}
            ActiveSheet.AutoFilterMode = False
            Dim col As Integer
            For col = {self.start_column} To {end_column} Step 2
                ActiveSheet.Range("{self.data_range}").AutoFilter Field:=col, Criteria1:="{criteria.value}"
            Next col
            {"\n".join([code for _, code in self.position_codes])}
            Application.ScreenUpdating = True
        End Sub
        """
