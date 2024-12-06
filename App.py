import json
import os
from dataclasses import asdict
from time import sleep

from customtkinter import CTk, CTkButton, CTkEntry, CTkProgressBar, CTkSegmentedButton

from ButtonInjector import run
from ParserEngine import ParserEngine


class Entry:
    __slots__ = ["text", "url", "delete_button"]

    def __init__(self, parent, text_placeholder: str = "Pharmacy Name",
                 url_placeholder: str = "https://tabletka.by/pharmacies/****", initial_text: str = "",
                 initial_url: str = ""):
        self.text = self.create_entry(parent, text_placeholder, initial_text)
        self.url = self.create_entry(parent, url_placeholder, initial_url)

    def create_entry(self, parent, placeholder: str, initial_text: str = ""):
        entry = CTkEntry(parent, placeholder_text=placeholder)
        if initial_text:
            entry.insert(0, initial_text)
        entry.bind("<Control-a>", lambda e: ["break", e.widget.select_range(0, 'end'), e.widget.icursor('end')][0])
        entry.bind("<Escape>", lambda e: e.widget.select_clear())
        return entry

    def grid(self, text_row, url_row, column, padx, pady, sticky):
        self.text.grid(row=text_row, column=column, padx=padx, pady=pady, sticky=sticky)
        self.url.grid(row=url_row, column=column + 1, padx=padx, pady=pady, sticky=sticky)

    def destroy(self):
        self.text.grid_forget()
        self.url.grid_forget()

    def get_text(self):
        return self.text.get()

    def get_url(self):
        return self.url.get()


class Profile:
    __slots__ = ["parent", "entries"]

    def __init__(self, parent, values):
        self.parent = parent
        self.entries = []
        for title, url in values.items():
            entry = Entry(parent, initial_text=title, initial_url=url)
            self.entries.append(entry)

    def hide(self):
        for entry in self.entries:
            entry.destroy()

    def display(self):
        for i, entry in enumerate(self.entries):
            entry.grid(text_row=i + 2, url_row=i + 2, column=0, padx=(5, 0), pady=(5, 5), sticky="nsew")

    def add_entry(self):
        entry = Entry(self.parent)
        self.entries.append(entry)
        self.display()

    def delete_entry(self):
        if self.entries:
            entry = self.entries.pop()
            entry.destroy()
            self.display()


class ProfileSelector(CTkSegmentedButton):
    __slots__ = ["app", "profiles"]

    def __init__(self, app, profiles, **kwargs):
        super().__init__(app, **kwargs)
        self.app = app
        self.profiles: list[Profile] = profiles
        self.configure(values=[f"Profile {i+1}" for i in range(len(self.profiles))], command=self.change_profile)
        self.set(f"Profile 1")
        self.change_profile(f"Profile 1")

    def change_profile(self, profile):
        index = int(profile.split(" ")[-1]) - 1
        for p in self.profiles:
            p.hide()
        self.app.current_profile = self.profiles[index]
        self.app.current_profile.display()

    def add_profile(self):
        new_profile_name = f"Profile {len(self.profiles) + 1}"
        self.profiles.append(Profile(self.app, {}))
        self.configure(values=[f"Profile {i+1}" for i in range(len(self.profiles))])
        self.set(new_profile_name)
        self.change_profile(new_profile_name)

    def delete_profile(self):
        if self.app.current_profile and len(self.profiles) > 1:
            index = self.profiles.index(self.app.current_profile)
            self.app.current_profile.hide()
            self.profiles.pop(index)
            self.configure(values=[f"Profile {i+1}" for i in range(len(self.profiles))])
            self.set(f"Profile {len(self.profiles)}")
            self.change_profile(f"Profile {len(self.profiles)}")


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
        CTkButton(self, text="Delete Profile", command=self.selector.delete_profile).grid(row=0, column=4, padx=10, pady=10)

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
            "data": self.engine.cookies["data"],
            "cookies": self.engine.cookies["cookies"]
        }

        for i, profile in enumerate(self.profiles):
            config["profiles"][f"Profile {i + 1}"] = {entry.get_text(): entry.get_url() for entry in profile.entries}

        with open(self.engine.config, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        self.destroy()
