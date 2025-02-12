class Button:
    def __init__(self, cell_address, caption, macro, back_color=None, fore_color=None):
        self.cell_address = cell_address
        self.caption = caption
        self.macro = macro
        self.back_color = back_color
        self.fore_color = fore_color
        self.button_name = None

    def create(self, worksheet):
        cell = worksheet.Range(self.cell_address)
        left = cell.Left
        top = cell.Top
        width = cell.Width
        height = cell.Height

        button = worksheet.Buttons().Add(left, top, width, height)
        button.Caption = self.caption
        button.OnAction = self.macro.name
        self.button_name = button.Name
        self.macro.add_position_code(self.generate_position_code(), self.restore_position_code())

    def generate_position_code(self):
        id_name = self.button_name.replace(' ', '')
        return f"""
        Dim btn{id_name} As Button
        Set btn{id_name} = ActiveSheet.Buttons("{self.button_name}")
        Dim btn{id_name}Left As Double
        Dim btn{id_name}Top As Double
        Dim btn{id_name}Width As Double
        Dim btn{id_name}Height As Double
        btn{id_name}Left = btn{id_name}.Left
        btn{id_name}Top = btn{id_name}.Top
        btn{id_name}Width = btn{id_name}.Width
        btn{id_name}Height = btn{id_name}.Height
        """

    def restore_position_code(self):
        id_name = self.button_name.replace(' ', '')
        return f"""
        btn{id_name}.Left = btn{id_name}Left
        btn{id_name}.Top = btn{id_name}Top
        btn{id_name}.Width = btn{id_name}Width
        btn{id_name}.Height = btn{id_name}Height
        """
