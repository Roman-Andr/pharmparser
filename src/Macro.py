from abc import abstractmethod, ABC


class Macro(ABC):
    start_col = 4
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


class RemoveFiltersMacro(Macro):
    def __init__(self, end_column):
        self.start_column = Macro.start_col
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
