from Macro import Macro


class ApplyFiltersMacro(Macro):
    def __init__(self, end_column, criteria):
        self.start_column = Macro.start_col
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