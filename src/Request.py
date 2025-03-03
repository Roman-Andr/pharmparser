import http.client
import json
import os
import urllib.parse
from dataclasses import dataclass

import psutil


@dataclass
class Request:
    __slots__ = ["url", "headers", "data"]
    url: str
    headers: dict[str, str]
    data: dict[str, str]

    def fetch(self, target):
        psutil.Process(os.getpid()).nice(psutil.HIGH_PRIORITY_CLASS)

        raise Exception("Not implemented")
        conn = http.client.HTTPSConnection("tabletka.by")

        conn.request(
            "POST",
            "/ajax-request/reload-pharmacy-price/",
            body=urllib.parse.urlencode({
                "id": int(target),
                **self.data
            }),
            headers=self.headers,
        )

        response = conn.getresponse().read().decode()
        conn.close()
        return json.loads(response)["data"]
