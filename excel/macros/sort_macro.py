from .macro import Macro


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
