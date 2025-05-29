from .macro import Macro


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
