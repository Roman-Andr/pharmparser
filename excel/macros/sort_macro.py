from openpyxl.utils import get_column_letter

from .macro import Macro


class SortMacro(Macro):
    def __init__(self, column, end_column, sort_order, sheet_name=""):
        self.column = column
        self.end_column = end_column
        self.sort_order = sort_order
        self.sheet_name = sheet_name
        code_template = """
        Sub Sort{sort_name}{column}_{sheet_name}()
            Application.ScreenUpdating = False
            {position_code_block}
            ActiveSheet.Range("{data_range}").Sort Key1:=ActiveSheet.Columns("{column}"), Order1:={sort_order}, Header:=xlYes
            {restore_code_block}
            Application.ScreenUpdating = True
        End Sub
        """
        super().__init__(f'Sort{sort_order.name}{column}_{sheet_name}', code_template)

    def get_code(self):
        position_code_block = "\n".join([code for code, _ in self.position_codes])
        restore_code_block = "\n".join([code for _, code in self.position_codes])
        data_range = f"A{Macro.start_row}:{get_column_letter(self.end_column)}{Macro.end_row}"
        return self.code_template.format(
            position_code_block=position_code_block,
            restore_code_block=restore_code_block,
            column=self.column,
            sort_order=self.sort_order.value,
            sort_name=self.sort_order.name,
            data_range=data_range,
            sheet_name=self.sheet_name
        )
