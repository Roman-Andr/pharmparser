from customtkinter import CTkSegmentedButton

from Entry import Entry


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
        self.configure(values=[f"Profile {i + 1}" for i in range(len(self.profiles))], command=self.change_profile)
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
        self.configure(values=[f"Profile {i + 1}" for i in range(len(self.profiles))])
        self.set(new_profile_name)
        self.change_profile(new_profile_name)

    def delete_profile(self):
        if self.app.current_profile and len(self.profiles) > 1:
            index = self.profiles.index(self.app.current_profile)
            self.app.current_profile.hide()
            self.profiles.pop(index)
            self.configure(values=[f"Profile {i + 1}" for i in range(len(self.profiles))])
            self.set(f"Profile {len(self.profiles)}")
            self.change_profile(f"Profile {len(self.profiles)}")
