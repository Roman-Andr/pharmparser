import json
import os
import string
from itertools import chain
from multiprocessing import Pool
from threading import Thread
from typing import List, Tuple, Callable

import requests
from bs4 import BeautifulSoup
from openpyxl.formatting import Rule
from openpyxl.styles import PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.workbook import Workbook


class ParserEngine:
    config = "config.json"
    green = "19CF1F"
    red = "E81737"
    title = "Анализ"
    fileName = "data.xlsx"
    colWidth = 45
    cellWidth = 13
    diffWidth = 9

    @classmethod
    def start(cls, entries: List[Tuple[str, int]], done: Callable):
        thread = Thread(target=cls.process, args=(entries, done))
        thread.start()

    @staticmethod
    def save(data):
        with open("data.json", "w", encoding="utf-8") as outfile:
            json.dump(data, outfile)

    @classmethod
    def load(cls):
        if not os.path.exists(cls.config):
            return None
        with open(cls.config, "r") as infile:
            return json.load(infile)

    @staticmethod
    def download(target: int, limit: int = 5000):
        request = requests.post(
            "https://tabletka.by/ajax-request/reload-pharmacy-price",
            headers={
                "Accept": "application/json, text/javascript, */*; q=0.01",
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                "X-Requested-With": "XMLHttpRequest",
                "sec-ch-ua-mobile": "?0",
                "Sec-Fetch-Site": "same-origin",
                "Sec-Fetch-Mode": "cors",
                "Sec-Fetch-Dest": "empty",
                "host": "tabletka.by",
            },
            data={
                "_csrf": "Your _csrt",
                "id": target,
                "page": "0",
                "sort": "name",
                "sort_type": "asc"
            },
            cookies={
                # Your cookies
                "lim-result": str(limit)
            }
        )
        return json.loads(request.content)["data"]

    @staticmethod
    def parse(file: str):
        soup = BeautifulSoup(file, 'lxml')
        names = [x.text + ", " + y.text for x, y in zip(
            soup.select("div[class=tooltip-info-header] > a"),
            soup.select("span[class=form-title]")
        )]
        prices = [x.text.strip().rstrip(" р.").lstrip("от ") for x in
                  soup.select("div[class=tooltip-info-header] > span[class=price-value]")]
        entry = {name: float(price) for name, price in zip(names, prices)}
        return entry

    @classmethod
    def process(cls, entries: List[Tuple[str, int]], done: Callable):
        codes = [y for x, y in entries]
        titles = [x for x, y in entries]
        with Pool(len(codes)) as pool:
            pages = pool.map(cls.download, codes)
        with Pool(len(codes)) as pool:
            parse_res = pool.map(cls.parse, pages)
        data = dict(zip(codes, parse_res))
        data["titles"] = {y: x for x, y in entries}
        cls.save(data)
        cls.excel_export(codes, titles, data)
        done()

    @classmethod
    def excel_export(cls, codes: List[int], titles: List[str], data):
        wb = Workbook()
        grid = [
            ["Название", titles[0]] +
            list(chain(*[[titles[i + 1], "Разница"] for i, x in enumerate(codes) if x != codes[-1]]))]
        names = sorted(list(set(chain(*[list(data[x].keys()) for x in codes]))), key=lambda k: k[0])
        for x in names:
            prices = []
            for y in codes:
                price1, price2 = data[codes[0]].get(x, "Нет"), data[y].get(x, "Нет")
                prices.append(price2)
                if (price2 == "Нет" or price1 == "Нет") and y != codes[0]:
                    prices.append(0)
                elif y != codes[0]:
                    prices.append(float(f"{float(f'{(float(price2) - float(price1)):.2f}'):+}"))
            row = [x] + prices
            grid.append(row)
        ws = wb.active
        ws.title = cls.title
        for x in string.ascii_uppercase:
            ws.column_dimensions[x].width = cls.cellWidth
        for x in string.ascii_uppercase[3::2]:
            ws.column_dimensions[x].width = cls.diffWidth
        ws.column_dimensions["A"].width = cls.colWidth

        for x in string.ascii_uppercase[3::2]:
            red_cell, green_cell = PatternFill(bgColor=cls.red), PatternFill(bgColor=cls.green)
            dxf_red, dxf_green = DifferentialStyle(fill=red_cell), DifferentialStyle(fill=green_cell)
            rule_less = Rule("cellIs", operator="lessThan", formula=["0"], dxf=dxf_red)
            rule_higher = Rule("cellIs", operator="greaterThan", formula=["0"], dxf=dxf_green)
            [ws.conditional_formatting.add(f"{x}2:{x}{len(grid)}", rule) for rule in (rule_less, rule_higher)]

        ws.auto_filter.ref = f"A1:{string.ascii_uppercase[len(grid[0]) - 1]}{len(grid)}"

        [ws.append(x) for x in grid]
        wb.save(cls.fileName)