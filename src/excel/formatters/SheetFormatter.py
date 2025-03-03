from abc import ABC, abstractmethod
from typing import List

from openpyxl.worksheet.worksheet import Worksheet

from src.utils import DataType
from src.utils import Settings


class SheetFormatter(ABC):
    __slots__ = ["settings", "titles", "data", "title"]

    def __init__(self, settings: Settings, data: DataType, titles: List[str]):
        self.settings = settings
        self.title = None
        self.data = data
        self.titles = titles

    @abstractmethod
    def format(self, sheet: Worksheet):
        pass
