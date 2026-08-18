import os
from multiprocessing import Pool

import psutil
from bs4 import BeautifulSoup

from ..utils import DataType, Request


class ParserEngine:
    __slots__ = ["callback", "errors", "request"]

    def __init__(self, request: Request):
        self.errors: list[Exception] = []
        self.request = request

    def parse(self, file: str) -> dict[str, float]:
        psutil.Process(os.getpid()).nice(psutil.REALTIME_PRIORITY_CLASS)
        soup = BeautifulSoup(file, 'lxml')
        # TODO(B7, docs/REFACTOR_PLAN.md phase 2): zipping three independent selections
        # truncates to the shortest and pairs names with the wrong prices when the page
        # structure drifts. strict=False preserves today's behaviour until the parser is
        # rewritten to walk one result row at a time.
        names = [x.text + ", " + y.text for x, y in zip(
            soup.select("div[class=tooltip-info-header] > a"),
            soup.select("span[class=form-title]"),
            strict=False,
        )]
        prices = [x.text.strip().rstrip(" р.").lstrip("от ") for x in
                  soup.select("span[class=price-value]")]
        return {name: float(price) for name, price in zip(names, prices, strict=False)}

    def f(self, code):
        result = {}
        for page in self.request.fetch(code):
            page_result = self.parse(page)
            result |= page_result
        return result

    def process(self, entries: list[tuple[str, int]]) -> tuple[list[str], DataType]:
        codes = [y for x, y in entries]
        titles = [x for x, y in entries]
        with Pool(len(codes)) as pool:
            parse_res: list[dict[str, float]] = pool.map(self.f, codes)
        return titles, dict(zip(titles, parse_res, strict=True))
