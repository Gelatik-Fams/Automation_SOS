import customtkinter as ctk

from gui.splash import SplashScreen
from gui.main_window import MainWindow


def run_app():
    ctk.set_appearance_mode('light')
    ctk.set_default_color_theme('blue')

    root = MainWindow()
    root.withdraw()

    def _show():
        root.deiconify()
        root.lift()
        root.focus_force()

    SplashScreen(root, duration=2000, on_done=_show)
    root.mainloop()
