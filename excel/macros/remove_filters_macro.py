from .macro import Macro


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
