import os
from multiprocessing import Pool
from typing import List, Tuple, Dict

import psutil
from bs4 import BeautifulSoup

from utils import Request, DataType


class ParserEngine:
    __slots__ = ["request", "errors"]

    def __init__(self, request: Request):
        self.errors = []
        self.request = request

    def parse(self, file: str) -> Dict[str, float]:
        psutil.Process(os.getpid()).nice(psutil.REALTIME_PRIORITY_CLASS)
        soup = BeautifulSoup(file, 'lxml')
        names = [x.text + ", " + y.text for x, y in zip(
            soup.select("div[class=tooltip-info-header] > a"),
            soup.select("span[class=form-title]")
        )]
        prices = [x.text.strip().rstrip(" р.").lstrip("от ") for x in
                  soup.select("span[class=price-value]")]
        entry = {name: float(price) for name, price in zip(names, prices)}
        return entry

    def f(self, code):
        return self.parse(self.request.fetch(code))

    def process(self, entries: List[Tuple[str, int]]) -> Tuple[List[str], DataType]:
        codes = [y for x, y in entries]
        titles = [x for x, y in entries]

        with Pool(len(codes)) as pool:
            parse_res: List[Dict[str, float]] = pool.map(
                self.f,
                codes
            )

        return titles, dict(zip(titles, parse_res))
