"""Tab 页面构建方法 (Mixin)"""
import customtkinter as ctk
from .constants import (UI_FONT, CN_FONTS, EN_FONTS, HEADING_LEVELS,
                        HEADING_LABELS, MUTED_LIGHT, MUTED_DARK)
from .widgets import SectionCard


class TabsMixin:
    """各设置 Tab 页的构建方法，混入 Md2DocxApp"""

    def _tab_page_body(self, tab):
        # 页面边距
        sec1 = SectionCard(tab, '页面边距')
        sec1.pack(fill='x', pady=(0, 10))
        row = ctk.CTkFrame(sec1.body, fg_color='transparent')
        row.pack(fill='x')
        for label, var in [('上 (cm)', self.page_top), ('下 (cm)', self.page_bottom),
                           ('左 (cm)', self.page_left), ('右 (cm)', self.page_right)]:
            f = ctk.CTkFrame(row, fg_color='transparent')
            f.pack(side='left', padx=(0, 24))
            ctk.CTkLabel(f, text=label, font=(UI_FONT, 12),
                         text_color=(MUTED_LIGHT, MUTED_DARK)).pack(anchor='w')
            ctk.CTkEntry(f, textvariable=var, width=72, height=32).pack(anchor='w', pady=(2, 0))

        # 正文样式
        sec2 = SectionCard(tab, '正文样式')
        sec2.pack(fill='x')

        row1 = ctk.CTkFrame(sec2.body, fg_color='transparent')
        row1.pack(fill='x', pady=(0, 8))
        for label, var, vals, w in [('中文字体', self.body_font_cn, CN_FONTS, 140),
                                    ('英文字体', self.body_font_en, EN_FONTS, 160)]:
            f = ctk.CTkFrame(row1, fg_color='transparent')
            f.pack(side='left', padx=(0, 24))
            ctk.CTkLabel(f, text=label, font=(UI_FONT, 12),
                         text_color=(MUTED_LIGHT, MUTED_DARK)).pack(anchor='w')
            ctk.CTkComboBox(f, variable=var, values=vals, width=w, height=32).pack(anchor='w', pady=(2, 0))

        row2 = ctk.CTkFrame(sec2.body, fg_color='transparent')
        row2.pack(fill='x')
        for label, var in [('字号 (pt)', self.body_size), ('行距 (倍)', self.body_line_spacing),
                           ('首行缩进 (cm)', self.body_indent)]:
            f = ctk.CTkFrame(row2, fg_color='transparent')
            f.pack(side='left', padx=(0, 24))
            ctk.CTkLabel(f, text=label, font=(UI_FONT, 12),
                         text_color=(MUTED_LIGHT, MUTED_DARK)).pack(anchor='w')
            ctk.CTkEntry(f, textvariable=var, width=72, height=32).pack(anchor='w', pady=(2, 0))

    def _tab_headings(self, tab):
        sw_frame = ctk.CTkFrame(tab, fg_color='transparent')
        sw_frame.pack(fill='x', pady=(0, 10))
        ctk.CTkSwitch(sw_frame, text='  启用标题自动编号  (一、 / 1.1 / 1.1.1)',
                      variable=self.auto_numbering, font=(UI_FONT, 13)).pack(anchor='w')

        sec = SectionCard(tab, '各级标题样式')
        sec.pack(fill='both', expand=True)
        grid = sec.body

        headers = ['级别', '中文字体', '字号 (pt)', '加粗']
        for col, text in enumerate(headers):
            ctk.CTkLabel(grid, text=text, font=(UI_FONT, 12, 'bold'),
                         text_color=(MUTED_LIGHT, MUTED_DARK)).grid(
                row=0, column=col, padx=(0, 12), pady=(0, 8), sticky='w')

        for idx, key in enumerate(HEADING_LEVELS):
            r = idx + 1
            font_size = max(11, 15 - idx)
            ctk.CTkLabel(grid, text=HEADING_LABELS[idx],
                         font=(UI_FONT, font_size, 'bold')).grid(
                row=r, column=0, padx=(0, 12), pady=5, sticky='w')
            ctk.CTkComboBox(grid, variable=self.headings[key]['font'],
                            values=CN_FONTS, width=130, height=30).grid(
                row=r, column=1, padx=(0, 12), pady=5)
            ctk.CTkEntry(grid, textvariable=self.headings[key]['size'],
                         width=68, height=30).grid(
                row=r, column=2, padx=(0, 12), pady=5)
            ctk.CTkCheckBox(grid, text='', variable=self.headings[key]['bold'],
                            width=24, height=24).grid(
                row=r, column=3, padx=(8, 0), pady=5)

    def _tab_table(self, tab):
        top_row = ctk.CTkFrame(tab, fg_color='transparent')
        top_row.pack(fill='x', pady=(0, 8))
        top_row.columnconfigure(0, weight=1)
        top_row.columnconfigure(1, weight=1)

        # 表头样式
        sec1 = SectionCard(top_row, '表头样式')
        sec1.grid(row=0, column=0, sticky='nsew', padx=(0, 5))
        r1 = sec1.body
        row = ctk.CTkFrame(r1, fg_color='transparent')
        row.pack(fill='x', pady=(0, 6))
        f = ctk.CTkFrame(row, fg_color='transparent')
        f.pack(side='left', padx=(0, 16))
        ctk.CTkLabel(f, text='字体', font=(UI_FONT, 12),
                     text_color=(MUTED_LIGHT, MUTED_DARK)).pack(anchor='w')
        ctk.CTkComboBox(f, variable=self.tbl_header_font, values=CN_FONTS,
                        width=120, height=30).pack(anchor='w', pady=(2, 0))
        f2 = ctk.CTkFrame(row, fg_color='transparent')
        f2.pack(side='left', padx=(0, 12))
        ctk.CTkLabel(f2, text='字号', font=(UI_FONT, 12),
                     text_color=(MUTED_LIGHT, MUTED_DARK)).pack(anchor='w')
        ctk.CTkEntry(f2, textvariable=self.tbl_header_size, width=60, height=30).pack(anchor='w', pady=(2, 0))
        ctk.CTkCheckBox(r1, text='加粗', variable=self.tbl_header_bold,
                        font=(UI_FONT, 12)).pack(anchor='w')

        # 正文样式
        sec2 = SectionCard(top_row, '正文样式')
        sec2.grid(row=0, column=1, sticky='nsew', padx=(5, 0))
        r2 = sec2.body
        row2 = ctk.CTkFrame(r2, fg_color='transparent')
        row2.pack(fill='x')
        f3 = ctk.CTkFrame(row2, fg_color='transparent')
        f3.pack(side='left', padx=(0, 16))
        ctk.CTkLabel(f3, text='字体', font=(UI_FONT, 12),
                     text_color=(MUTED_LIGHT, MUTED_DARK)).pack(anchor='w')
        ctk.CTkComboBox(f3, variable=self.tbl_body_font, values=CN_FONTS,
                        width=120, height=30).pack(anchor='w', pady=(2, 0))
        f4 = ctk.CTkFrame(row2, fg_color='transparent')
        f4.pack(side='left')
        ctk.CTkLabel(f4, text='字号', font=(UI_FONT, 12),
                     text_color=(MUTED_LIGHT, MUTED_DARK)).pack(anchor='w')
        ctk.CTkEntry(f4, textvariable=self.tbl_body_size, width=60, height=30).pack(anchor='w', pady=(2, 0))

        # 行高 + 边框
        bot_row = ctk.CTkFrame(tab, fg_color='transparent')
        bot_row.pack(fill='x')
        bot_row.columnconfigure(0, weight=1)
        bot_row.columnconfigure(1, weight=1)

        sec3 = SectionCard(bot_row, '行高设置')
        sec3.grid(row=0, column=0, sticky='nsew', padx=(0, 5))
        rh = sec3.body
        ctk.CTkRadioButton(rh, text='Word 默认', variable=self.tbl_row_height_mode,
                           value='none', font=(UI_FONT, 12)).pack(anchor='w', pady=(0, 4))
        rh_row = ctk.CTkFrame(rh, fg_color='transparent')
        rh_row.pack(fill='x', pady=(0, 4))
        ctk.CTkRadioButton(rh_row, text='固定 (pt)', variable=self.tbl_row_height_mode,
                           value='pt', font=(UI_FONT, 12)).pack(side='left', padx=(0, 8))
        ctk.CTkRadioButton(rh_row, text='行距 (倍)', variable=self.tbl_row_height_mode,
                           value='line', font=(UI_FONT, 12)).pack(side='left')
        val_row = ctk.CTkFrame(rh, fg_color='transparent')
        val_row.pack(fill='x')
        ctk.CTkLabel(val_row, text='数值:', font=(UI_FONT, 12),
                     text_color=(MUTED_LIGHT, MUTED_DARK)).pack(side='left', padx=(0, 6))
        ctk.CTkEntry(val_row, textvariable=self.tbl_row_height_val,
                     width=72, height=30).pack(side='left')

        sec4 = SectionCard(bot_row, '边框样式')
        sec4.grid(row=0, column=1, sticky='nsew', padx=(5, 0))
        bs = sec4.body
        ctk.CTkRadioButton(bs, text='全边框网格线', variable=self.tbl_style,
                           value='grid', font=(UI_FONT, 12)).pack(anchor='w', pady=(0, 6))
        ctk.CTkRadioButton(bs, text='三线表 (学术标准)', variable=self.tbl_style,
                           value='three_line', font=(UI_FONT, 12)).pack(anchor='w')

    def _tab_formula_caption(self, tab):
        sec1 = SectionCard(tab, '公式编号格式')
        sec1.pack(fill='x', pady=(0, 10))
        ff = sec1.body
        ctk.CTkRadioButton(ff, text='纯序号:   (1),  (2),  (3)  …',
                           variable=self.formula_numbering, value='plain',
                           font=(UI_FONT, 13)).pack(anchor='w', pady=(0, 8))
        ctk.CTkRadioButton(ff, text='章节序号:   (1.1),  (1.2),  (2.1)  …  按 H1 分章',
                           variable=self.formula_numbering, value='chapter',
                           font=(UI_FONT, 13)).pack(anchor='w')

        sec_list = SectionCard(tab, '列表编号格式')
        sec_list.pack(fill='x', pady=(0, 10))
        lf = sec_list.body
        ctk.CTkRadioButton(lf, text='学术格式:   1）→（1）→ ① → a）',
                           variable=self.list_style, value='academic',
                           font=(UI_FONT, 13)).pack(anchor='w', pady=(0, 6))
        ctk.CTkRadioButton(lf, text='圈号格式:   ①②③ → ⑴⑵⑶ → a. → 1.',
                           variable=self.list_style, value='circle',
                           font=(UI_FONT, 13)).pack(anchor='w')

        sec2 = SectionCard(tab, '图表标题样式')
        sec2.pack(fill='x')
        row = ctk.CTkFrame(sec2.body, fg_color='transparent')
        row.pack(fill='x')
        f1 = ctk.CTkFrame(row, fg_color='transparent')
        f1.pack(side='left', padx=(0, 24))
        ctk.CTkLabel(f1, text='字体', font=(UI_FONT, 12),
                     text_color=(MUTED_LIGHT, MUTED_DARK)).pack(anchor='w')
        ctk.CTkComboBox(f1, variable=self.cap_font, values=CN_FONTS,
                        width=140, height=32).pack(anchor='w', pady=(2, 0))
        f2 = ctk.CTkFrame(row, fg_color='transparent')
        f2.pack(side='left')
        ctk.CTkLabel(f2, text='字号 (pt)', font=(UI_FONT, 12),
                     text_color=(MUTED_LIGHT, MUTED_DARK)).pack(anchor='w')
        ctk.CTkEntry(f2, textvariable=self.cap_size, width=72, height=32).pack(anchor='w', pady=(2, 0))

        # 页眉页脚
        sec3 = SectionCard(tab, '页眉页脚')
        sec3.pack(fill='x', pady=(10, 0))
        hf = sec3.body

        hf_row1 = ctk.CTkFrame(hf, fg_color='transparent')
        hf_row1.pack(fill='x', pady=(0, 8))
        f_ht = ctk.CTkFrame(hf_row1, fg_color='transparent')
        f_ht.pack(side='left', padx=(0, 16))
        ctk.CTkLabel(f_ht, text='页眉文字 (留空不显示)', font=(UI_FONT, 12),
                     text_color=(MUTED_LIGHT, MUTED_DARK)).pack(anchor='w')
        ctk.CTkEntry(f_ht, textvariable=self.hf_header_text, width=240, height=30,
                     placeholder_text='如：XX技术报告').pack(anchor='w', pady=(2, 0))
        f_hf = ctk.CTkFrame(hf_row1, fg_color='transparent')
        f_hf.pack(side='left', padx=(0, 16))
        ctk.CTkLabel(f_hf, text='页眉字体', font=(UI_FONT, 12),
                     text_color=(MUTED_LIGHT, MUTED_DARK)).pack(anchor='w')
        ctk.CTkComboBox(f_hf, variable=self.hf_header_font, values=CN_FONTS,
                        width=120, height=30).pack(anchor='w', pady=(2, 0))
        f_hs = ctk.CTkFrame(hf_row1, fg_color='transparent')
        f_hs.pack(side='left')
        ctk.CTkLabel(f_hs, text='字号', font=(UI_FONT, 12),
                     text_color=(MUTED_LIGHT, MUTED_DARK)).pack(anchor='w')
        ctk.CTkEntry(f_hs, textvariable=self.hf_header_size, width=60, height=30).pack(anchor='w', pady=(2, 0))

        hf_row2 = ctk.CTkFrame(hf, fg_color='transparent')
        hf_row2.pack(fill='x')
        ctk.CTkCheckBox(hf_row2, text='显示页码', variable=self.hf_page_number,
                        font=(UI_FONT, 12)).pack(side='left', padx=(0, 16))
        ctk.CTkRadioButton(hf_row2, text='底部居中', variable=self.hf_page_pos,
                           value='bottom_center', font=(UI_FONT, 12)).pack(side='left', padx=(0, 8))
        ctk.CTkRadioButton(hf_row2, text='底部靠右', variable=self.hf_page_pos,
                           value='bottom_right', font=(UI_FONT, 12)).pack(side='left', padx=(0, 16))
        f_ps = ctk.CTkFrame(hf_row2, fg_color='transparent')
        f_ps.pack(side='left')
        ctk.CTkLabel(f_ps, text='字号:', font=(UI_FONT, 12),
                     text_color=(MUTED_LIGHT, MUTED_DARK)).pack(side='left', padx=(0, 4))
        ctk.CTkEntry(f_ps, textvariable=self.hf_page_size, width=50, height=30).pack(side='left')
