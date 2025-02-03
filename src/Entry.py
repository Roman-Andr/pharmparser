from customtkinter import CTkEntry


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
