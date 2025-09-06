from utils import create_custom_entry


class Entry:
    __slots__ = ["text", "url", "delete_button"]

    def __init__(self, parent, text_placeholder: str = "Pharmacy Name",
                 url_placeholder: str = "https://tabletka.by/pharmacies/****", initial_text: str = "",
                 initial_url: str = ""):
        self.text = create_custom_entry(parent, text_placeholder, initial_text)
        self.url = create_custom_entry(parent, url_placeholder, initial_url)

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
