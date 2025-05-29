import asyncio
import json
import os
from dataclasses import asdict
from threading import Thread
from typing import List, Tuple, Callable

import aiohttp
from CTkMessagebox import CTkMessagebox
from customtkinter import CTk, CTkButton, CTkProgressBar

from core import ParserEngine
from excel import Spreadsheet, DataFormatter, AnalysisFormatter
from utils import Request, Settings
from .profile import Profile
from .profile_selector import ProfileSelector


class App(CTk):
    __slots__ = ["progress", "profiles", "current_profile", "engine", "selector", "add_profile_button",
                 "delete_profile_button", "settings"]
    processing: bool = False

    def __init__(self):
        super().__init__()

        self.progress: CTkProgressBar = CTkProgressBar(self)
        self.geometry(f"{1100}x{600}")
        self.title("PharmParser")

        self.profiles: list[Profile] = []
        self.current_profile: Profile = None
        CTkButton(self, text="Add", command=self.add_entry).grid(row=1, column=0, padx=30, pady=5)
        CTkButton(self, text="Delete", command=self.delete_entry).grid(row=1, column=1, padx=45, pady=5)
        CTkButton(self, text="Parse", command=self.click).grid(row=1, column=2, padx=30, pady=5)
        self.config = "config.json"
        with open(self.config, "r", encoding="utf-8") as f:
            loaded = json.load(f)
            self.settings = Settings(**loaded["settings"])
            request = Request(**loaded["request"])
            for values in loaded["profiles"].values():
                self.profiles.append(Profile(self, values))
        self.engine = ParserEngine(request)

        self.selector = ProfileSelector(self, self.profiles)
        self.selector.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        CTkButton(self, text="Add Profile", command=self.selector.add).grid(row=0, column=3, padx=10, pady=10)
        CTkButton(self, text="Delete Profile", command=self.selector.remove).grid(row=0, column=4, padx=10,
                                                                                  pady=10)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def add_entry(self):
        if self.current_profile:
            self.current_profile.add_entry()

    def delete_entry(self):
        if self.current_profile:
            self.current_profile.delete_entry()

    def click(self):
        if self.processing:
            return
        self.processing = True
        try:
            thread = Thread(target=self.start, args=(
                [(entry.get_text(), int(entry.get_url().split("/")[-1])) for entry in self.current_profile.entries],
                self.done))
            thread.start()
            self.progress.grid(row=sum([len(p.entries) for p in self.profiles]) + 1,
                               column=0,
                               columnspan=3,
                               padx=(20, 10),
                               pady=(10, 10),
                               sticky="ew")
            self.progress.configure(mode="indeterminate")
            self.progress.start()
        except Exception as e:
            self.done(False)
            CTkMessagebox(title="Error", message=f"An error occurred: {str(e)}", icon="cancel")

    def start(self, entries: List[Tuple[str, int]], done: Callable):
        titles, data = self.engine.process(entries)
        # asyncio.run(self.send_request(data))
        with open("data.json", "w", encoding="utf-8") as file:
            json.dump(data, file, ensure_ascii=False, indent=2)
        Spreadsheet(data, self.settings, [
            (DataFormatter(self.settings, data, titles, lambda p1, p2: p2 - p1), "Данные"),
            (DataFormatter(self.settings, data, titles, lambda p1, p2: (p2 - p1) / p1 * 100), "Проценты"),
            (AnalysisFormatter(self.settings, data, titles), "Анализ")
        ]).export(data)
        done()

    async def send_request(self, data):
        async with aiohttp.ClientSession() as session:
            async with session.post(
                    "http://localhost:8000/upload-prices/",
                    json=data,
                    headers={'Content-Type': 'application/json'}
            ) as response:
                response_data = await response.json()
                print("Response:", response_data)

    def done(self, status=True):
        self.processing = False
        self.progress.stop()
        self.progress.grid_forget()
        if status:
            os.startfile(os.path.abspath(os.getcwd()) + f"\\{self.settings.fileName.replace("xlsx", "xlsm")}")
        else:
            for e in self.engine.errors:
                CTkMessagebox(title="Error", message=f"An error occurred: {str(e)}\n\n{e}", icon="cancel")

    def on_closing(self):
        config = {
            "profiles": {},
            "settings": asdict(self.settings),
            "request": asdict(self.engine.request)
        }

        for i, profile in enumerate(self.profiles):
            config["profiles"][f"Profile {i + 1}"] = {entry.get_text(): entry.get_url() for entry in profile.entries}

        with open(self.config, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        self.destroy()
