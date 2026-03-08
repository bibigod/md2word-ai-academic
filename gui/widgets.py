"""自定义 GUI 组件"""
import customtkinter as ctk
from .constants import UI_FONT, MUTED_LIGHT, MUTED_DARK


class FormRow(ctk.CTkFrame):
    """表单行：标签 + 控件，统一间距"""
    def __init__(self, master, label, widget_factory, **kw):
        super().__init__(master, fg_color='transparent')
        self.label = ctk.CTkLabel(self, text=label, font=(UI_FONT, 13),
                                  text_color=(MUTED_LIGHT, MUTED_DARK))
        self.label.pack(side='left', padx=(0, 8))
        self.widget = widget_factory(self)
        self.widget.pack(side='left')


class SectionCard(ctk.CTkFrame):
    """带标题的卡片分组"""
    def __init__(self, master, title, **kw):
        super().__init__(master, corner_radius=12, border_width=1,
                         border_color=('#D4D4E8', '#3e3e5c'), **kw)
        title_bar = ctk.CTkFrame(self, fg_color=('#EEEDF8', '#2e2e44'),
                                  corner_radius=0, height=34)
        title_bar.pack(fill='x', padx=1, pady=(1, 0))
        ctk.CTkLabel(title_bar, text=title, font=(UI_FONT, 13, 'bold'),
                     text_color=('#4A48B0', '#A8A6E0')
                     ).pack(side='left', padx=14, pady=6)
        self.body = ctk.CTkFrame(self, fg_color='transparent')
        self.body.pack(fill='both', expand=True, padx=16, pady=(10, 14))
