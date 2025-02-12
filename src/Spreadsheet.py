from typing import List, Type

from openpyxl.workbook import Workbook

from ButtonInjector import run
from SheetFormatter import SheetFormatter
from datatypes import DataType
from settings import Settings


class Spreadsheet:
    def __init__(self, data: DataType, settings: Settings, formatters: List[Type[SheetFormatter]]):
        self.data = data
        self.settings = settings
        self.formatters = formatters

    def export(self, titles: List[str], data: DataType):
        wb = Workbook()
        wb.remove(wb.active)
        for formatter in self.formatters:
            form = formatter(self.settings, data, titles)
            form.format(wb.create_sheet(form.title))
        wb.save(self.settings.fileName)
        run(len(data) * 2)

    # def excel_export(self, codes: List[int], titles: List[str], data: DataType):
    #     offset = 2
    #     wb = Workbook()
    #     grid = [*([] for _ in range(offset)),
    #             ["Название", titles[0]] +
    #             list(chain(*[[titles[i + 1], "Разница"] for i, x in enumerate(titles) if x != titles[-1]]))]
    #     names = sorted(list(set(chain(*[list(data[x].keys()) for x in titles]))), key=lambda k: k.lower())
    #     for x in names:
    #         prices = []
    #         for y in titles:
    #             price1, price2 = data[titles[0]].get(x, "Нет"), data[y].get(x, "Нет")
    #             prices.append(price2)
    #             if (price2 == "Нет" or price1 == "Нет") and y != codes[0]:
    #                 prices.append(0)
    #             elif y != titles[0]:
    #                 prices.append(float(f"{float(f'{(float(price2) - float(price1)):.2f}'):+}"))
    #         row = [x] + prices
    #         grid.append(row)
    #     ws = wb.active
    #     ws.title = self.settings.data_title
    #     for x in string.ascii_uppercase:
    #         ws.column_dimensions[x].width = self.settings.cellWidth
    #     for x in string.ascii_uppercase[3::2]:
    #         ws.column_dimensions[x].width = self.settings.diffWidth
    #     ws.column_dimensions["A"].width = self.settings.colWidth
    #
    #     for x in string.ascii_uppercase[3::2]:
    #         red_cell, green_cell = PatternFill(bgColor=self.settings.red), PatternFill(bgColor=self.settings.green)
    #         dxf_red, dxf_green = DifferentialStyle(fill=red_cell), DifferentialStyle(fill=green_cell)
    #         rule_less = Rule("cellIs", operator="lessThan", formula=["0"], dxf=dxf_red)
    #         rule_higher = Rule("cellIs", operator="greaterThan", formula=["0"], dxf=dxf_green)
    #         [ws.conditional_formatting.add(f"{x}{2 + offset}:{x}{len(grid)}", rule) for rule in
    #          (rule_less, rule_higher)]
    #
    #     ws.auto_filter.ref = f"A{1 + offset}:{get_column_letter(len(grid[0 + offset]))}{len(grid)}"
    #
    #     [ws.append(x) for x in grid]
    #
    #     # ws = wb.create_sheet(title=self.settings.title)
    #     # grid = [
    #     #     ["Асортимент", ],
    #     #     ["Позиций ниже всех", ],
    #     # ]
    #
    #     wb.save(self.settings.fileName)
    #     run(len(grid[0 + offset]))
