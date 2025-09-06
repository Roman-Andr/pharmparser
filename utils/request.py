import json
import os
import urllib.parse
from copy import deepcopy
from dataclasses import dataclass
from http.client import HTTPSConnection

import psutil


@dataclass
class Request:
    __slots__ = ["url", "headers", "data"]
    url: str
    headers: dict[str, str]
    data: dict[str, str]

    def fetch(self, target):
        psutil.Process(os.getpid()).nice(psutil.REALTIME_PRIORITY_CLASS)

        response = self.request(target, 0, 10)
        data = []
        for i in range(1, response["priceCount"] // 5000 + 2):
            data.append(self.request(target, i)["data"])
        return data

    def request(self, target, page=0, limit=5000):
        conn = HTTPSConnection("tabletka.by")
        headers = deepcopy(self.headers)
        headers["Cookie"] = headers["Cookie"].replace("lim-result=5000", f"lim-result={limit}")
        conn.request(
            "POST",
            "/ajax-request/reload-pharmacy-price/",
            body=urllib.parse.urlencode({
                "id": int(target),
                "page": str(page),
                **self.data
            }),
            headers=headers,
        )
        response = conn.getresponse().read().decode()
        conn.close()
        return json.loads(response)
