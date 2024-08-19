from multiprocessing import freeze_support

from customtkinter import set_appearance_mode, set_default_color_theme

from App import App


def main():
    set_appearance_mode("System")
    set_default_color_theme("blue")

    app = App()
    app.mainloop()


if __name__ == "__main__":
    freeze_support()
    main()