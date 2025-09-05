from customtkinter import CTkSegmentedButton

from .profile import Profile


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

    def add(self):
        new_profile_name = f"Profile {len(self.profiles) + 1}"
        self.profiles.append(Profile(self.app, {}))
        self.configure(values=[f"Profile {i + 1}" for i in range(len(self.profiles))])
        self.set(new_profile_name)
        self.change_profile(new_profile_name)

    def remove(self):
        if self.app.current_profile and len(self.profiles) > 1:
            index = self.profiles.index(self.app.current_profile)
            self.app.current_profile.hide()
            self.profiles.pop(index)
            self.configure(values=[f"Profile {i + 1}" for i in range(len(self.profiles))])
            self.set(f"Profile {len(self.profiles)}")
            self.change_profile(f"Profile {len(self.profiles)}")
