import json
import os
from dataclasses import dataclass

import psutil
import requests


@dataclass
class Request:
    __slots__ = ["url", "headers", "cookies", "data"]
    url: str
    headers: dict[str, str]
    cookies: dict[str, str]
    data: dict[str, str]

    def fetch(self, target):
        psutil.Process(os.getpid()).nice(psutil.HIGH_PRIORITY_CLASS)
        request = requests.post(
            self.url,
            headers=self.headers,
            cookies=self.cookies,
            data={
                "id": int(target),
                **self.data
            }
        )
        return json.loads(request.content)["data"]
