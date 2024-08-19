import os

from customtkinter import CTk, CTkButton, CTkEntry, CTkProgressBar

from ParserEngine import ParserEngine


class App(CTk):
    def __init__(self):
        super().__init__()

        self.progressbar_1 = None
        self.title("PharmParser")
        self.geometry(f"{1100}x{580}")

        self.entries = []
        self.add_btn("Add", self.add_entry, 0)
        self.add_btn("Delete", self.delete_button, 1)
        self.add_btn("Parse", self.click, 2)
        config = ParserEngine.load()
        if config is not None:
            for x in config:
                self.add_entry(x, config[x])

    def add_btn(self, text, command, column):
        btn = CTkButton(self, text=text, command=command)
        btn.grid(row=0, column=column, padx=100, pady=5)

    def add_entry(self, title="", url=""):
        entry_name = CTkEntry(self, placeholder_text="Pharmacy Name")
        if title != "":
            entry_name.insert(0, title)
        entry_name.grid(row=len(self.entries) + 1, column=0, columnspan=1, padx=(5, 0), pady=(5, 5), sticky="nsew")
        entry_url = CTkEntry(self, placeholder_text="https://tabletka.by/pharmacies/****")
        if url != "":
            entry_url.insert(0, url)
        entry_url.grid(row=len(self.entries) + 1, column=1, columnspan=1, padx=(5, 0), pady=(5, 5), sticky="nsew")
        self.entries.append((entry_name, entry_url))

    def delete_button(self):
        if len(self.entries) == 0:
            return

        [x.destroy() for x in self.entries[-1]]
        self.entries.remove(self.entries[-1])

    def click(self):
        ParserEngine.start([(x.get(), int(y.get().split("/")[-1])) for x, y in self.entries], self.done)
        self.progressbar_1 = CTkProgressBar(self)
        self.progressbar_1.grid(row=len(self.entries) + 2,
                                column=0,
                                columnspan=2,
                                padx=(20, 10),
                                pady=(10, 10),
                                sticky="ew")
        self.progressbar_1.configure(mode="indeterminnate")
        self.progressbar_1.start()

    def done(self):
        self.progressbar_1.stop()
        self.progressbar_1.destroy()
        os.startfile(os.path.abspath(os.getcwd()) + f"\\{ParserEngine.fileName}")