import json
import os
import urllib.parse
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

        response = self.request(target, 0)
        data = response["data"]
        if response["priceCount"] > 5000:
            for i in range(1, response["priceCount"] // 5000 + 2):
                data += self.request(target, i + 1)["data"]
        return data

    def request(self, target, page=0):
        conn = HTTPSConnection("tabletka.by")

        conn.request(
            "POST",
            "/ajax-request/reload-pharmacy-price/",
            body=urllib.parse.urlencode({
                "id": int(target),
                "page": str(page),
                **self.data
            }),
            headers=self.headers,
        )
        response = conn.getresponse().read().decode()
        conn.close()
        return json.loads(response)
