import sys
from pathlib import Path

import customtkinter as ctk

from gui.config import get_version
from gui.widgets import C_BLUE_DARK, C_BLUE_LIGHT, C_WHITE

try:
    from PIL import Image
    _PIL_OK = True
except ImportError:
    _PIL_OK = False


def _assets_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS) / 'assets'  # type: ignore[attr-defined]
    return Path(__file__).parent.parent / 'assets'


class SplashScreen(ctk.CTkToplevel):
    """2-second branded splash shown before the main window."""

    def __init__(self, master, duration: int = 2000, on_done=None):
        super().__init__(master)
        self._on_done = on_done
        self._duration = duration

        self.overrideredirect(True)
        self.configure(fg_color=C_BLUE_DARK)
        self.resizable(False, False)

        w, h = 360, 240
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f'{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}')
        self.lift()
        self.focus_force()

        self._build_ui()
        self.after(self._duration, self._finish)

    def _build_ui(self):
        container = ctk.CTkFrame(self, fg_color='transparent')
        container.pack(expand=True)

        logo_path = _assets_dir() / 'logo.png'
        if _PIL_OK and logo_path.exists():
            img = Image.open(logo_path).resize((80, 80))
            self._logo_img = ctk.CTkImage(img, size=(80, 80))
            ctk.CTkLabel(container, image=self._logo_img, text='').pack(pady=(0, 8))

        ctk.CTkLabel(
            container,
            text='Gelatik Automation',
            font=ctk.CTkFont(family='Helvetica', size=22, weight='bold'),
            text_color=C_WHITE,
        ).pack()

        ctk.CTkLabel(
            container,
            text='SOS Dashboard Generator',
            font=ctk.CTkFont(family='Helvetica', size=13),
            text_color='#A3C2E3',
        ).pack(pady=(2, 4))

        ctk.CTkLabel(
            container,
            text=f'v{get_version()}',
            font=ctk.CTkFont(family='Helvetica', size=10),
            text_color='#6E90AE',
        ).pack(pady=(0, 14))

        self._bar = ctk.CTkProgressBar(
            container, width=240, mode='indeterminate',
            fg_color='#2E6DA4', progress_color=C_BLUE_LIGHT,
        )
        self._bar.pack()
        self._bar.start()

    def _finish(self):
        self._bar.stop()
        self.destroy()
        if self._on_done:
            self._on_done()
