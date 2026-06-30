import customtkinter as ctk

C_BLUE_DARK = '#1F4E78'
C_BLUE_MED = '#2E6DA4'
C_BLUE_LIGHT = '#A3C2E3'
C_GREY_BG = '#F4F4F4'
C_GREEN = '#2E7D32'
C_RED = '#C62828'
C_WHITE = '#FFFFFF'
C_TEXT_DARK = '#1A1A1A'
C_TEXT_MUTED = '#6B6B6B'


class SectionCard(ctk.CTkFrame):
    """Titled card container with a header bar and content area."""

    def __init__(self, master, title: str, **kwargs):
        super().__init__(master, fg_color=C_WHITE, corner_radius=8, **kwargs)
        self.configure(border_width=1, border_color=C_BLUE_LIGHT)

        header = ctk.CTkFrame(self, fg_color=C_BLUE_DARK, corner_radius=0, height=32)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)

        ctk.CTkLabel(
            header,
            text=title,
            font=ctk.CTkFont(family='Helvetica', size=12, weight='bold'),
            text_color=C_WHITE,
        ).pack(side='left', padx=10, pady=4)

        self.content = ctk.CTkFrame(self, fg_color=C_WHITE, corner_radius=0)
        self.content.pack(fill='both', expand=True, padx=8, pady=8)


class ClusterRow(ctk.CTkFrame):
    """One row showing cluster name + detected status."""

    def __init__(self, master, cluster_name: str, **kwargs):
        super().__init__(master, fg_color='transparent', **kwargs)

        dot = ctk.CTkLabel(self, text='●', font=ctk.CTkFont(size=14), text_color=C_GREEN)
        dot.pack(side='left', padx=(0, 6))

        ctk.CTkLabel(
            self,
            text=cluster_name,
            font=ctk.CTkFont(size=12),
            text_color=C_TEXT_DARK,
        ).pack(side='left')

        ctk.CTkLabel(
            self,
            text='Terdeteksi',
            font=ctk.CTkFont(size=11),
            text_color=C_TEXT_MUTED,
        ).pack(side='right', padx=4)
