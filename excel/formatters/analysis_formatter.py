import string
from itertools import chain
from typing import List

from numpy import mean
from openpyxl.worksheet.worksheet import Worksheet

from utils import DataType, Settings
from .base_formatter import BaseFormatter


class AnalysisFormatter(BaseFormatter):
    def __init__(self, settings: Settings, data: DataType, titles: List[str]):
        super().__init__(settings, data, titles)

    def format(self, ws: Worksheet):
        title = self.titles[0]
        names = set(chain(*[list(self.data[x].keys()) for x in self.titles]))
        ws.column_dimensions["A"].width = self.settings.colWidth
        for x in string.ascii_uppercase[1::]:
            ws.column_dimensions[x].width = self.settings.cellWidth
        grid = [
            [title],
            ["Асортимент", len(self.data[title])],
            ["Средний асортимент конкурентов",
             mean([len(self.data[name]) for name in self.titles if name != title])],
            ["Позиций ниже всех", sum(
                1 for item, price in self.data[title].items()
                if all(price < self.data.get(competitor, {}).get(item, float('-inf'))
                       for competitor in self.titles if competitor != title)
            )],
            ["Уникальных позиций", sum(
                1 for item in self.data[title]
                if all(item not in self.data.get(competitor, {})
                       for competitor in self.titles if competitor != title)
            )],
            ["", "Асортимент", "Дороже", "Дешевле", "Уникальных"],
            *[
                [name,
                 len(self.data.get(name, {})),
                 sum(
                     1 for item in self.data[title]
                     if item in self.data.get(name, {}) and self.data[title][item] < self.data[name][item]),
                 sum(
                     1 for item in self.data[title]
                     if item in self.data.get(name, {}) and self.data[title][item] > self.data[name][item]),
                 sum(
                     1 for item in self.data[name]
                     if item not in self.data.get(title, {}))
                 ]
                for name in self.titles if name != title
            ]
        ]
        [ws.append(x) for x in grid]
