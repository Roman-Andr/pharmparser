from abc import ABC, abstractmethod

from openpyxl.worksheet.worksheet import Worksheet

from ...utils import DataType, Settings


class BaseFormatter(ABC):
    __slots__ = ["data", "settings", "title", "titles"]

    def __init__(self, settings: Settings, data: DataType, titles: list[str]):
        self.settings = settings
        self.title = None
        self.data = data
        self.titles = titles

    @abstractmethod
    def format(self, sheet: Worksheet):
        pass
