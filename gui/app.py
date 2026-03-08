"""
MD → DOCX 转换器 GUI 主应用
"""
import sys
import os
import io
import contextlib
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import yaml
import customtkinter as ctk

from .constants import (UI_FONT, MUTED_LIGHT, MUTED_DARK, HEADING_LEVELS,
                        DEFAULT_SIZES, get_resource_path)

# 确保能 import 同目录的 converter 包
sys.path.insert(0, get_resource_path())
from converter import convert, DEFAULT_CONFIG, _deep_merge
from .tabs import TabsMixin
from .preview import PreviewMixin


class Md2DocxApp(PreviewMixin, TabsMixin, ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('Md2word-ai-academic | 适合中国宝宝体质的 Markdown 转 Word')
        self.geometry('1100x860')
        self.minsize(900, 750)

        self._init_vars()
        self._build_top_bar()
        self._build_file_bar()
        self._build_tabs()

        self._auto_load_config()
        self._auto_load_sample()

    def _init_vars(self):
        self.md_path = tk.StringVar()
        self.docx_path = tk.StringVar()

        self.page_top = tk.StringVar(value='2.54')
        self.page_bottom = tk.StringVar(value='2.54')
        self.page_left = tk.StringVar(value='3.17')
        self.page_right = tk.StringVar(value='3.17')

        self.body_font_cn = tk.StringVar(value='仿宋_GB2312')
        self.body_font_en = tk.StringVar(value='Times New Roman')
        self.body_size = tk.StringVar(value='12')
        self.body_line_spacing = tk.StringVar(value='1.0')
        self.body_indent = tk.StringVar(value='0.74')

        self.headings = {}
        for idx, key in enumerate(HEADING_LEVELS):
            self.headings[key] = {
                'font': tk.StringVar(value='黑体'),
                'size': tk.StringVar(value=str(DEFAULT_SIZES[idx])),
                'bold': tk.BooleanVar(value=True),
            }
        self.auto_numbering = tk.BooleanVar(value=True)

        self.tbl_header_font = tk.StringVar(value='黑体')
        self.tbl_header_size = tk.StringVar(value='11')
        self.tbl_header_bold = tk.BooleanVar(value=True)
        self.tbl_body_font = tk.StringVar(value='仿宋_GB2312')
        self.tbl_body_size = tk.StringVar(value='11')
        self.tbl_row_height_mode = tk.StringVar(value='pt')
        self.tbl_row_height_val = tk.StringVar(value='20')
        self.tbl_style = tk.StringVar(value='three_line')

        self.formula_numbering = tk.StringVar(value='chapter')

        self.cap_font = tk.StringVar(value='黑体')
        self.cap_size = tk.StringVar(value='11')

        self.list_style = tk.StringVar(value='academic')

        self.hf_header_text = tk.StringVar(value='')
        self.hf_header_font = tk.StringVar(value='仿宋_GB2312')
        self.hf_header_size = tk.StringVar(value='9')
        self.hf_page_number = tk.BooleanVar(value=True)
        self.hf_page_pos = tk.StringVar(value='bottom_center')
        self.hf_page_size = tk.StringVar(value='9')

    # ============================================================
    #  TOP BAR
    # ============================================================
    def _build_top_bar(self):
        header = ctk.CTkFrame(self, fg_color='#4F4DC8', corner_radius=0, height=56)
        header.pack(fill='x')
        header.pack_propagate(False)

        ctk.CTkLabel(header, text='适合中国宝宝体质的 Markdown 转 Word',
                     font=(UI_FONT, 16, 'bold'),
                     text_color='#FFFFFF').pack(side='left', padx=(20, 0))

        ctk.CTkButton(header, text=' ▶ 开始转换 ', width=130, height=36,
                      font=(UI_FONT, 13, 'bold'), corner_radius=8,
                      fg_color='#FFFFFF', text_color='#4A48C4',
                      hover_color='#E8E7FF',
                      command=self._on_convert).pack(side='right', padx=(10, 20), pady=10)

        hdr_btn = dict(height=32, corner_radius=8, fg_color='transparent',
                       border_width=1.5, border_color='#9997E8',
                       text_color='#FFFFFF', hover_color='#5D5BD4',
                       font=(UI_FONT, 12))
        ctk.CTkButton(header, text='恢复默认', width=80, command=self._on_reset,
                      **hdr_btn).pack(side='right', padx=4, pady=10)
        ctk.CTkButton(header, text='保存配置', width=80, command=self._on_save_config,
                      **hdr_btn).pack(side='right', padx=4, pady=10)
        ctk.CTkButton(header, text='加载配置', width=80, command=self._on_load_config,
                      **hdr_btn).pack(side='right', padx=4, pady=10)

    # ============================================================
    #  FILE BAR
    # ============================================================
    def _build_file_bar(self):
        bar = ctk.CTkFrame(self, fg_color='transparent')
        bar.pack(fill='x', padx=20, pady=(8, 2))

        ctk.CTkLabel(bar, text='输入文件', font=(UI_FONT, 12),
                     text_color=(MUTED_LIGHT, MUTED_DARK)).pack(side='left', padx=(0, 4))
        ctk.CTkEntry(bar, textvariable=self.md_path, height=30, width=320,
                     placeholder_text='.md 文件…').pack(side='left')
        ctk.CTkButton(bar, text='📂', width=36, height=30,
                      font=(UI_FONT, 13), command=self._browse_md).pack(
            side='left', padx=(4, 20))

        ctk.CTkLabel(bar, text='输出文件', font=(UI_FONT, 12),
                     text_color=(MUTED_LIGHT, MUTED_DARK)).pack(side='left', padx=(0, 4))
        ctk.CTkEntry(bar, textvariable=self.docx_path, height=30, width=320,
                     placeholder_text='.docx 路径…').pack(side='left')
        ctk.CTkButton(bar, text='📂', width=36, height=30,
                      font=(UI_FONT, 13), command=self._browse_docx).pack(
            side='left', padx=(4, 0))

    # ============================================================
    #  TABVIEW
    # ============================================================
    def _build_tabs(self):
        self.tabs = ctk.CTkTabview(self, corner_radius=12, anchor='nw',
                                   segmented_button_fg_color='#F3F4F6',
                                   segmented_button_selected_color='#5856D6',
                                   segmented_button_unselected_color='#F3F4F6',
                                   segmented_button_selected_hover_color='#4A48C4',
                                   segmented_button_unselected_hover_color='#E5E7EB',
                                   text_color='#FFFFFF',
                                   text_color_disabled='#AEAEB2')
        self.tabs.pack(fill='both', expand=True, padx=20, pady=(4, 8))

        self.tabs.add('预览')
        self.tabs.add('页面与正文')
        self.tabs.add('标题')
        self.tabs.add('表格')
        self.tabs.add('公式与标注')

        _SETTINGS = {'页面与正文', '标题', '表格', '公式与标注'}
        _PREVIEW_COLOR = '#5856D6'
        _SETTING_COLOR = '#2E86AB'
        _SETTING_HOVER = '#256E8D'
        _UNSEL = '#F3F4F6'
        _UNSEL_HOVER = '#E5E7EB'
        try:
            seg_btn = self.tabs._segmented_button
            original_cb = seg_btn._command

            def _on_tab_change(value):
                original_cb(value)
                for name, btn in seg_btn._buttons_dict.items():
                    if name == value:
                        c = _SETTING_COLOR if name in _SETTINGS else _PREVIEW_COLOR
                        btn.configure(text_color='#FFFFFF', fg_color=c,
                                      hover_color=(_SETTING_HOVER if name in _SETTINGS
                                                   else '#4A48C4'))
                    else:
                        btn.configure(text_color='#48484a', fg_color=_UNSEL,
                                      hover_color=_UNSEL_HOVER)
            seg_btn.configure(command=_on_tab_change)

            btns = list(seg_btn._buttons_dict.items())
            if len(btns) >= 2:
                btns[0][1].grid_configure(padx=(4, 20))

            for name, btn in seg_btn._buttons_dict.items():
                if name == '预览':
                    btn.configure(text_color='#FFFFFF', fg_color=_PREVIEW_COLOR)
                else:
                    btn.configure(text_color='#48484a', fg_color=_UNSEL)
        except Exception:
            pass

        self._tab_preview(self.tabs.tab('预览'))
        self._tab_page_body(self.tabs.tab('页面与正文'))
        self._tab_headings(self.tabs.tab('标题'))
        self._tab_table(self.tabs.tab('表格'))
        self._tab_formula_caption(self.tabs.tab('公式与标注'))

    # ============================================================
    #  Config <-> UI
    # ============================================================
    def _s(self, var):
        try:
            return float(var.get())
        except (ValueError, tk.TclError):
            return 0.0

    def _build_config(self):
        cfg = {
            'page': {
                'top_margin_cm': self._s(self.page_top),
                'bottom_margin_cm': self._s(self.page_bottom),
                'left_margin_cm': self._s(self.page_left),
                'right_margin_cm': self._s(self.page_right),
            },
            'body': {
                'font_cn': self.body_font_cn.get(),
                'font_en': self.body_font_en.get(),
                'font_size_pt': self._s(self.body_size),
                'line_spacing': self._s(self.body_line_spacing),
                'first_line_indent_cm': self._s(self.body_indent),
            },
            'headings': {'auto_numbering': self.auto_numbering.get()},
            'table': {
                'header_font_cn': self.tbl_header_font.get(),
                'header_font_size_pt': self._s(self.tbl_header_size),
                'header_bold': self.tbl_header_bold.get(),
                'body_font_cn': self.tbl_body_font.get(),
                'body_font_size_pt': self._s(self.tbl_body_size),
                'style': self.tbl_style.get(),
            },
            'formula': {'numbering': self.formula_numbering.get()},
            'list': {'style': self.list_style.get()},
            'caption': {
                'font_cn': self.cap_font.get(),
                'font_size_pt': self._s(self.cap_size),
            },
            'header_footer': {
                'header_text': self.hf_header_text.get(),
                'header_font_cn': self.hf_header_font.get(),
                'header_font_size_pt': self._s(self.hf_header_size),
                'page_number': self.hf_page_number.get(),
                'page_number_position': self.hf_page_pos.get(),
                'page_number_font_size_pt': self._s(self.hf_page_size),
            },
        }
        for key in HEADING_LEVELS:
            cfg['headings'][key] = {
                'font_cn': self.headings[key]['font'].get(),
                'size_pt': self._s(self.headings[key]['size']),
                'bold': self.headings[key]['bold'].get(),
            }
        mode = self.tbl_row_height_mode.get()
        if mode == 'pt':
            cfg['table']['row_height_pt'] = self._s(self.tbl_row_height_val)
        elif mode == 'line':
            cfg['table']['row_height_line'] = self._s(self.tbl_row_height_val)
        return cfg

    def _load_from_config(self, cfg):
        page = cfg.get('page', {})
        self.page_top.set(str(page.get('top_margin_cm', 2.54)))
        self.page_bottom.set(str(page.get('bottom_margin_cm', 2.54)))
        self.page_left.set(str(page.get('left_margin_cm', 3.17)))
        self.page_right.set(str(page.get('right_margin_cm', 3.17)))

        body = cfg.get('body', {})
        self.body_font_cn.set(body.get('font_cn', '仿宋_GB2312'))
        self.body_font_en.set(body.get('font_en', 'Times New Roman'))
        self.body_size.set(str(body.get('font_size_pt', 12)))
        self.body_line_spacing.set(str(body.get('line_spacing', 1.0)))
        self.body_indent.set(str(body.get('first_line_indent_cm', 0.74)))

        headings = cfg.get('headings', {})
        self.auto_numbering.set(headings.get('auto_numbering', False))
        for key in HEADING_LEVELS:
            h = headings.get(key, {})
            self.headings[key]['font'].set(h.get('font_cn', '黑体'))
            self.headings[key]['size'].set(str(h.get('size_pt', 12)))
            self.headings[key]['bold'].set(h.get('bold', True))

        tbl = cfg.get('table', {})
        self.tbl_header_font.set(tbl.get('header_font_cn', '黑体'))
        self.tbl_header_size.set(str(tbl.get('header_font_size_pt', 11)))
        self.tbl_header_bold.set(tbl.get('header_bold', True))
        self.tbl_body_font.set(tbl.get('body_font_cn', '仿宋_GB2312'))
        self.tbl_body_size.set(str(tbl.get('body_font_size_pt', 11)))
        self.tbl_style.set(tbl.get('style', 'grid'))

        if 'row_height_pt' in tbl:
            self.tbl_row_height_mode.set('pt')
            self.tbl_row_height_val.set(str(tbl['row_height_pt']))
        elif 'row_height_line' in tbl:
            self.tbl_row_height_mode.set('line')
            self.tbl_row_height_val.set(str(tbl['row_height_line']))
        else:
            self.tbl_row_height_mode.set('none')
            self.tbl_row_height_val.set('20')

        formula = cfg.get('formula', {})
        self.formula_numbering.set(formula.get('numbering', 'plain'))

        list_cfg = cfg.get('list', {})
        self.list_style.set(list_cfg.get('style', 'academic'))

        cap = cfg.get('caption', {})
        self.cap_font.set(cap.get('font_cn', '黑体'))
        self.cap_size.set(str(cap.get('font_size_pt', 11)))

        hf = cfg.get('header_footer', {})
        self.hf_header_text.set(hf.get('header_text', ''))
        self.hf_header_font.set(hf.get('header_font_cn', '仿宋_GB2312'))
        self.hf_header_size.set(str(hf.get('header_font_size_pt', 9)))
        self.hf_page_number.set(hf.get('page_number', True))
        self.hf_page_pos.set(hf.get('page_number_position', 'bottom_center'))
        self.hf_page_size.set(str(hf.get('page_number_font_size_pt', 9)))

    # ============================================================
    #  Handlers
    # ============================================================
    def _browse_md(self):
        path = filedialog.askopenfilename(
            title='选择 Markdown 文件',
            filetypes=[('Markdown', '*.md'), ('All Files', '*.*')],
        )
        if path:
            self.md_path.set(path)
            if not self.docx_path.get():
                self.docx_path.set(os.path.splitext(path)[0] + '.docx')
            self._refresh_preview()

    def _browse_docx(self):
        path = filedialog.asksaveasfilename(
            title='选择输出路径', defaultextension='.docx',
            filetypes=[('Word', '*.docx')],
        )
        if path:
            self.docx_path.set(path)

    def _on_load_config(self):
        path = filedialog.askopenfilename(
            title='加载 YAML 配置',
            filetypes=[('YAML', '*.yaml *.yml'), ('All', '*.*')],
            initialdir=get_resource_path(),
        )
        if path:
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    user_cfg = yaml.safe_load(f) or {}
                cfg = _deep_merge(dict(DEFAULT_CONFIG), user_cfg)
                self._load_from_config(cfg)
                messagebox.showinfo('成功', f'已加载配置:\n{path}')
            except Exception as e:
                messagebox.showerror('错误', f'加载配置失败:\n{e}')

    def _on_save_config(self):
        path = filedialog.asksaveasfilename(
            title='保存 YAML 配置', defaultextension='.yaml',
            filetypes=[('YAML', '*.yaml *.yml')],
            initialfile='md2docx.yaml',
            initialdir=get_resource_path(),
        )
        if path:
            try:
                cfg = self._build_config()
                with open(path, 'w', encoding='utf-8') as f:
                    yaml.dump(cfg, f, allow_unicode=True,
                              default_flow_style=False, sort_keys=False)
                messagebox.showinfo('成功', f'配置已保存:\n{path}')
            except Exception as e:
                messagebox.showerror('错误', f'保存配置失败:\n{e}')

    def _on_reset(self):
        self._load_from_config(DEFAULT_CONFIG)

    def _on_convert(self):
        md = self.md_path.get()
        docx = self.docx_path.get()
        if not md or not os.path.isfile(md):
            messagebox.showerror('错误', '请选择一个有效的 Markdown 文件。')
            return
        if not docx:
            messagebox.showerror('错误', '请指定输出文件路径。')
            return

        cfg = self._build_config()

        def _do_convert():
            buf = io.StringIO()
            try:
                with contextlib.redirect_stdout(buf):
                    convert(md, docx, cfg)
                output = buf.getvalue()
                self.after(0, self._on_convert_done, output, docx)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror('转换失败', str(e)))

        threading.Thread(target=_do_convert, daemon=True).start()

    def _on_convert_done(self, output, docx):
        messagebox.showinfo('转换完成', output if output else f'已生成: {docx}')
        self._preview_docx_file(docx)

    def _auto_load_sample(self):
        sample = os.path.join(get_resource_path(),
                              'sample.md')
        if os.path.isfile(sample) and not self.md_path.get():
            self.md_path.set(sample)
            self.docx_path.set(os.path.splitext(sample)[0] + '.docx')
            self.after(500, self._refresh_preview)

    def _auto_load_config(self):
        auto_path = os.path.join(get_resource_path(),
                                 'md2docx.yaml')
        if os.path.isfile(auto_path):
            try:
                with open(auto_path, 'r', encoding='utf-8') as f:
                    user_cfg = yaml.safe_load(f) or {}
                cfg = _deep_merge(dict(DEFAULT_CONFIG), user_cfg)
                self._load_from_config(cfg)
            except Exception:
                pass
