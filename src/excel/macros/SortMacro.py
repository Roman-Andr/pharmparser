from .Macro import Macro


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
