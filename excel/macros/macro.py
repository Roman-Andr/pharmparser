from abc import abstractmethod, ABC

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
