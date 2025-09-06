from customtkinter import CTkEntry


def create_custom_entry(parent, placeholder: str, initial_text: str = ""):
    entry = CTkEntry(parent, placeholder_text=placeholder)
    if initial_text:
        entry.insert(0, initial_text)
    entry.bind("<Control-a>", lambda e: ["break", e.widget.select_range(0, 'end'), e.widget.icursor('end')][0])
    entry.bind("<Escape>", lambda e: e.widget.select_clear())
    return entry