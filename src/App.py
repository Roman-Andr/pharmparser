import json
import os
from dataclasses import asdict

from customtkinter import CTk, CTkButton, CTkProgressBar

from ParserEngine import ParserEngine
from Profile import Profile
from ProfileSelector import ProfileSelector


class App(CTk):
    __slots__ = ["progress", "profiles", "current_profile", "engine", "selector", "add_profile_button",
                 "delete_profile_button"]
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
        self.engine = ParserEngine("config.json")

        for values in self.engine.profiles.values():
            self.profiles.append(Profile(self, values))

        self.selector = ProfileSelector(self, self.profiles)
        self.selector.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        CTkButton(self, text="Add Profile", command=self.selector.add_profile).grid(row=0, column=3, padx=10, pady=10)
        CTkButton(self, text="Delete Profile", command=self.selector.delete_profile).grid(row=0, column=4, padx=10,
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
        self.engine.start(
            [(entry.get_text(), int(entry.get_url().split("/")[-1])) for entry in self.current_profile.entries],
            self.done)
        self.progress.grid(row=sum([len(p.entries) for p in self.profiles]) + 1,
                           column=0,
                           columnspan=3,
                           padx=(20, 10),
                           pady=(10, 10),
                           sticky="ew")
        self.progress.configure(mode="indeterminate")
        self.progress.start()

    def done(self):
        self.processing = False
        self.progress.stop()
        self.progress.grid_forget()
        os.startfile(os.path.abspath(os.getcwd()) + f"\\{self.engine.settings.fileName.replace("xlsx", "xlsm")}")

    def on_closing(self):
        config = {
            "profiles": {},
            "settings": asdict(self.engine.settings),
            "request": asdict(self.engine.request)
        }

        for i, profile in enumerate(self.profiles):
            config["profiles"][f"Profile {i + 1}"] = {entry.get_text(): entry.get_url() for entry in profile.entries}

        with open(self.engine.config, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        self.destroy()
