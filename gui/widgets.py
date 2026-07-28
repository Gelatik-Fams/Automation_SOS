import customtkinter as ctk

# ── Color Palette ─────────────────────────────────────────────────────
C_BLUE_DARK = '#1B2A4A'    # Topbar background, primary accent
C_BLUE_MED = '#3B82F6'     # Buttons, links, interactive elements
C_BLUE_LIGHT = '#E0E7EF'   # Borders, dividers
C_GREY_BG = '#F5F7FA'      # App background (warm off-white)
C_GREEN = '#22C55E'        # Success states
C_RED = '#EF4444'          # Error states
C_WHITE = '#FFFFFF'        # Card surfaces
C_TEXT_DARK = '#1E293B'    # Primary text
C_TEXT_MUTED = '#94A3B8'   # Secondary/muted text
C_TERMINAL_BG = '#1E293B'  # Terminal background
C_TERMINAL_FG = '#E2E8F0'  # Terminal text

# Shared font family
FONT_FAMILY = 'Segoe UI'

# ── Typography Scale ──────────────────────────────────────────────────
FONT_HEADER = lambda: ctk.CTkFont(family=FONT_FAMILY, size=14, weight='bold')
FONT_BODY = lambda: ctk.CTkFont(family=FONT_FAMILY, size=12)
FONT_BODY_BOLD = lambda: ctk.CTkFont(family=FONT_FAMILY, size=12, weight='bold')
FONT_CAPTION = lambda: ctk.CTkFont(family=FONT_FAMILY, size=11)
FONT_BTN = lambda: ctk.CTkFont(family=FONT_FAMILY, size=12, weight='bold')
FONT_LOG = lambda: ctk.CTkFont(family='Consolas', size=13)


class SectionCard(ctk.CTkFrame):
    """Modern card container with icon + title and content area."""

    def __init__(self, master, title: str, icon: str = '', **kwargs):
        super().__init__(master, fg_color=C_WHITE, corner_radius=12, **kwargs)
        self.configure(border_width=1, border_color=C_BLUE_LIGHT)

        # Header row with icon + title
        header_row = ctk.CTkFrame(self, fg_color='transparent')
        header_row.pack(side='top', fill='x', padx=14, pady=(10, 0))

        display_text = f'{icon}  {title}' if icon else title
        self.header_label = ctk.CTkLabel(
            header_row,
            text=display_text,
            font=FONT_HEADER(),
            text_color=C_TEXT_DARK,
        )
        self.header_label.pack(side='left')

        # Thin separator
        separator = ctk.CTkFrame(self, fg_color=C_BLUE_LIGHT, height=1, corner_radius=0)
        separator.pack(fill='x', padx=14, pady=(6, 0))
        separator.pack_propagate(False)

        self.content = ctk.CTkFrame(self, fg_color='transparent', corner_radius=0)
        self.content.pack(fill='both', expand=True, padx=14, pady=(6, 10))


class ClusterRow(ctk.CTkFrame):
    """One row showing cluster name + detected status."""

    def __init__(self, master, cluster_name: str, **kwargs):
        super().__init__(master, fg_color='transparent', **kwargs)

        dot = ctk.CTkLabel(self, text='●', font=ctk.CTkFont(size=10), text_color=C_GREEN)
        dot.pack(side='left', padx=(0, 6))

        ctk.CTkLabel(
            self,
            text=cluster_name,
            font=FONT_BODY_BOLD(),
            text_color=C_TEXT_DARK,
        ).pack(side='left')

        ctk.CTkLabel(
            self,
            text='Terdeteksi',
            font=FONT_CAPTION(),
            text_color=C_TEXT_MUTED,
        ).pack(side='right', padx=4)
