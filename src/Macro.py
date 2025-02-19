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
