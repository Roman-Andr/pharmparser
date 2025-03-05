import json
import os
from multiprocessing import Pool
from typing import List, Tuple, Dict

import psutil
from bs4 import BeautifulSoup

from utils import Request, DataType
from utils import Settings


class ParserEngine:
    __slots__ = ["config", "settings", "profiles", "request", "errors"]

    def __init__(self, config):
        self.config = config
        self.errors = []
        with open(config, "r", encoding="utf-8") as f:
            loaded = json.load(f)
            self.settings = Settings(**loaded["settings"])
            self.profiles: dict[str, dict[str, str]] = loaded["profiles"]
            self.request = Request(**loaded["request"])

    def parse(self, file: str) -> Dict[str, float]:
        psutil.Process(os.getpid()).nice(psutil.HIGH_PRIORITY_CLASS)
        soup = BeautifulSoup(file, 'lxml')
        names = [x.text + ", " + y.text for x, y in zip(
            soup.select("div[class=tooltip-info-header] > a"),
            soup.select("span[class=form-title]")
        )]
        prices = [x.text.strip().rstrip(" р.").lstrip("от ") for x in
                  soup.select("div[class=tooltip-info-header] > span[class=price-value]")]
        entry = {name: float(price) for name, price in zip(names, prices)}
        return entry

    def download(self, code):
        return self.request.fetch(code)

    def process(self, entries: List[Tuple[str, int]]) -> Tuple[List[str], DataType]:
        codes = [y for x, y in entries]
        titles = [x for x, y in entries]
        with Pool(len(codes)) as pool:
            pages = pool.map(self.download, codes)
        with Pool(len(pages)) as pool:
            parse_res: List[Dict[str, float]] = pool.map(self.parse, pages)
        return titles, dict(zip(titles, parse_res))
