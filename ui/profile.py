from .entry import Entry


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
