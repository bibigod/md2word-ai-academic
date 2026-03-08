"""预览渲染相关方法 (Mixin)"""
import os
import sys
import io
import tempfile
import threading
import contextlib
import tkinter as tk
import customtkinter as ctk
from PIL import Image, ImageTk

from .constants import UI_FONT, MUTED_LIGHT, MUTED_DARK, SECTION_TITLE_LIGHT, SECTION_TITLE_DARK

# 用于保护 stdout 重定向和渲染标志的锁
_render_lock = threading.Lock()


class PreviewMixin:
    """预览 Tab 构建 + 渲染逻辑，混入 Md2DocxApp"""

    def _tab_preview(self, tab):
        # 顶部工具栏
        top = ctk.CTkFrame(tab, fg_color='transparent')
        top.pack(fill='x', pady=(0, 6))
        ctk.CTkButton(top, text='刷新预览', width=100, height=32,
                      font=(UI_FONT, 13), command=self._refresh_preview).pack(side='left')
        self._preview_info = ctk.CTkLabel(top, text='选择 MD 文件后自动预览',
                                          font=(UI_FONT, 12),
                                          text_color=(MUTED_LIGHT, MUTED_DARK))
        self._preview_info.pack(side='left', padx=(12, 0))

        # 双栏容器
        panes = ctk.CTkFrame(tab, fg_color='transparent')
        panes.pack(fill='both', expand=True)
        panes.columnconfigure(0, weight=1)
        panes.columnconfigure(1, weight=1)
        panes.rowconfigure(1, weight=1)

        ctk.CTkLabel(panes, text='Markdown 源文件', font=(UI_FONT, 12, 'bold'),
                     text_color=(SECTION_TITLE_LIGHT, SECTION_TITLE_DARK)).grid(
            row=0, column=0, sticky='w', padx=(4, 0), pady=(0, 4))
        ctk.CTkLabel(panes, text='Word 文档预览', font=(UI_FONT, 12, 'bold'),
                     text_color=(SECTION_TITLE_LIGHT, SECTION_TITLE_DARK)).grid(
            row=0, column=1, sticky='w', padx=(12, 0), pady=(0, 4))

        # 左栏：MD 源码
        self._md_box = ctk.CTkTextbox(panes, font=('Consolas', 13), wrap='none',
                                       state='disabled', activate_scrollbars=True)
        self._md_box.grid(row=1, column=0, sticky='nsew', padx=(0, 4))

        # 右栏：Canvas 预览
        right_frame = ctk.CTkFrame(panes, fg_color=('#E8E8EC', '#38383a'), corner_radius=8)
        right_frame.grid(row=1, column=1, sticky='nsew', padx=(4, 0))
        right_frame.rowconfigure(0, weight=1)
        right_frame.columnconfigure(0, weight=1)

        self._preview_canvas = tk.Canvas(right_frame, bg='#E8E8EC', highlightthickness=0)
        self._preview_canvas.grid(row=0, column=0, sticky='nsew', padx=2, pady=2)

        self._preview_vscroll = ctk.CTkScrollbar(right_frame,
                                                  command=self._preview_canvas.yview)
        self._preview_vscroll.grid(row=0, column=1, sticky='ns')
        self._preview_hscroll = ctk.CTkScrollbar(right_frame, orientation='horizontal',
                                                  command=self._preview_canvas.xview)
        self._preview_hscroll.grid(row=1, column=0, sticky='ew')
        self._preview_canvas.configure(
            yscrollcommand=self._preview_vscroll.set,
            xscrollcommand=self._preview_hscroll.set)

        # 事件绑定
        self._preview_canvas.bind('<MouseWheel>', self._on_preview_scroll)
        self._preview_canvas.bind('<Shift-MouseWheel>', self._on_preview_hscroll)
        self._preview_canvas.bind('<Control-MouseWheel>', self._on_preview_zoom)
        self._preview_canvas.bind('<Configure>', self._on_preview_canvas_resize)
        self._preview_canvas.bind('<ButtonPress-2>', self._on_preview_drag_start)
        self._preview_canvas.bind('<B2-Motion>', self._on_preview_drag_move)

        # 状态
        self._preview_images = []
        self._preview_rendering = False
        self._preview_request = None
        self._preview_zoom = 1.0
        self._preview_cached_tk = []
        self._preview_cached_zoom = 0
        self._drag_start = (0, 0)

    # ---------- 滚动/缩放事件 ----------

    def _on_preview_scroll(self, event):
        self._preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

    def _on_preview_hscroll(self, event):
        self._preview_canvas.xview_scroll(int(-1 * (event.delta / 120)), 'units')

    def _on_preview_drag_start(self, event):
        self._preview_canvas.config(cursor='fleur')
        self._drag_start = (event.x, event.y)

    def _on_preview_drag_move(self, event):
        dx = event.x - self._drag_start[0]
        dy = event.y - self._drag_start[1]
        self._preview_canvas.xview_scroll(int(-dx), 'units')
        self._preview_canvas.yview_scroll(int(-dy), 'units')
        self._drag_start = (event.x, event.y)
        self._preview_canvas.config(cursor='')

    def _on_preview_zoom(self, event):
        canvas = self._preview_canvas
        cx = canvas.canvasx(event.x)
        cy = canvas.canvasy(event.y)

        old_zoom = self._preview_zoom
        factor = 1.08
        if event.delta > 0:
            new_zoom = min(old_zoom * factor, 4.0)
        else:
            new_zoom = max(old_zoom / factor, 0.3)

        if abs(new_zoom - old_zoom) < 0.001:
            return
        self._preview_zoom = new_zoom

        if not self._preview_images:
            return

        self._draw_pages_on_canvas()

        ratio = new_zoom / old_zoom
        new_cx = cx * ratio
        new_cy = cy * ratio

        sr = canvas.cget('scrollregion').split()
        if len(sr) == 4:
            sr_w = float(sr[2])
            sr_h = float(sr[3])
            if sr_w > 0 and sr_h > 0:
                canvas.xview_moveto(max(0, (new_cx - event.x) / sr_w))
                canvas.yview_moveto(max(0, (new_cy - event.y) / sr_h))

    def _on_preview_canvas_resize(self, event):
        if self._preview_images:
            self._preview_cached_zoom = 0
            self._draw_pages_on_canvas()

    # ---------- 绘制 ----------

    def _draw_pages_on_canvas(self):
        canvas = self._preview_canvas
        canvas_w = canvas.winfo_width()
        if canvas_w < 50:
            canvas_w = 500

        page_margin = 12
        gap = 16
        base_scale = (canvas_w - page_margin * 2) / max(
            self._preview_images[0].width, 1) if self._preview_images else 1.0
        scale = base_scale * self._preview_zoom

        need_resize = abs(scale - self._preview_cached_zoom) > 0.001

        if need_resize:
            self._preview_cached_tk = []
            self._preview_page_sizes = []
            for pil_img in self._preview_images:
                new_w = max(int(pil_img.width * scale), 1)
                new_h = max(int(pil_img.height * scale), 1)
                resized = pil_img.resize((new_w, new_h), Image.LANCZOS)
                tk_img = ImageTk.PhotoImage(resized)
                self._preview_cached_tk.append(tk_img)
                self._preview_page_sizes.append((new_w, new_h))
            self._preview_cached_zoom = scale

        max_page_w = 0
        for w, h in self._preview_page_sizes:
            max_page_w = max(max_page_w, w)
        content_w = max_page_w + page_margin * 2
        sr_w = max(canvas_w, content_w)

        canvas.delete('all')
        y_offset = page_margin

        for i, tk_img in enumerate(self._preview_cached_tk):
            new_w, new_h = self._preview_page_sizes[i]
            x0 = (sr_w - new_w) // 2

            canvas.create_rectangle(
                x0 + 3, y_offset + 3, x0 + new_w + 3, y_offset + new_h + 3,
                fill='#B0B0B0', outline='')
            canvas.create_rectangle(
                x0, y_offset, x0 + new_w, y_offset + new_h,
                fill='white', outline='#C8C8CC')
            canvas.create_image(x0, y_offset, anchor='nw', image=tk_img)

            y_offset += new_h + gap

        sr_h = y_offset + page_margin
        canvas.configure(scrollregion=(0, 0, sr_w, sr_h))

    # ---------- 预览调度 ----------

    def _refresh_preview(self):
        md = self.md_path.get()
        if not md or not os.path.isfile(md):
            self._set_textbox(self._md_box, '请先选择一个 Markdown 文件。')
            self._preview_canvas.delete('all')
            return

        try:
            with open(md, 'r', encoding='utf-8') as f:
                raw = f.read()
                lines = raw.splitlines()
        except Exception as e:
            self._set_textbox(self._md_box, f'读取文件失败: {e}')
            return

        numbered = []
        for idx, line in enumerate(lines, 1):
            numbered.append(f'{idx:>4}  {line}')
        self._set_textbox(self._md_box, '\n'.join(numbered))

        self._preview_request = ('md', md, self._build_config())
        if not self._preview_rendering:
            self._start_next_preview()

    def _preview_docx_file(self, docx_path):
        """直接预览已有的 DOCX 文件（转换完成后调用）"""
        self._preview_request = ('docx', docx_path, None)
        if not self._preview_rendering:
            self._start_next_preview()

    def _start_next_preview(self):
        with _render_lock:
            req = self._preview_request
            if not req:
                return
            self._preview_request = None
            self._preview_rendering = True
        self._preview_info.configure(text='正在生成 Word 预览...')

        if req[0] == 'docx':
            threading.Thread(target=self._do_render,
                             args=(req[1], None, None), daemon=True).start()
        else:
            threading.Thread(target=self._do_render,
                             args=(req[1], req[2], 'md'), daemon=True).start()

    def _do_render(self, path, cfg, mode):
        """后台线程：MD→DOCX→PDF→images 或 DOCX→PDF→images"""
        import fitz
        from converter import convert
        tmp_docx = None
        tmp_pdf = None
        try:
            if mode == 'md':
                tmp_f = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
                tmp_docx = tmp_f.name
                tmp_f.close()
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf):
                    convert(path, tmp_docx, cfg)
                docx_path = tmp_docx
            else:
                docx_path = path

            tmp_f2 = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
            tmp_pdf = tmp_f2.name
            tmp_f2.close()
            self._docx_to_pdf_wps(docx_path, tmp_pdf)

            doc = fitz.open(tmp_pdf)
            images = []
            for page in doc:
                mat = fitz.Matrix(1.5, 1.5)
                pix = page.get_pixmap(matrix=mat)
                img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
                images.append(img)
            doc.close()

            self.after(0, self._update_preview_images, images, len(images))
        except Exception as e:
            self.after(0, self._preview_error, str(e))
        finally:
            with _render_lock:
                self._preview_rendering = False
            for f in [tmp_docx, tmp_pdf]:
                if f and os.path.isfile(f):
                    try:
                        os.remove(f)
                    except OSError:
                        pass
            self.after(0, self._check_pending_preview)

    def _check_pending_preview(self):
        with _render_lock:
            if self._preview_request and not self._preview_rendering:
                pass  # will call _start_next_preview below
            else:
                return
        self._start_next_preview()

    # ---------- DOCX → PDF ----------

    def _docx_to_pdf_wps(self, docx_path, pdf_path):
        abs_docx = os.path.abspath(docx_path)
        abs_pdf = os.path.abspath(pdf_path)
        errors = []

        import pythoncom
        pythoncom.CoInitialize()
        try:
            return self._docx_to_pdf_inner(abs_docx, abs_pdf, errors)
        finally:
            pythoncom.CoUninitialize()

    def _docx_to_pdf_inner(self, abs_docx, abs_pdf, errors):
        # 1) WPS Office
        try:
            import win32com.client
            wps = win32com.client.Dispatch('KWPS.Application')
            wps.Visible = False
            wps.DisplayAlerts = False
            doc = wps.Documents.Open(abs_docx)
            doc.ExportAsFixedFormat(abs_pdf, 17)
            doc.Close(False)
            wps.Quit()
            return
        except Exception as e:
            errors.append(f'WPS: {e}')

        # 2) Microsoft Word
        try:
            import win32com.client
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            word.DisplayAlerts = False
            doc = word.Documents.Open(abs_docx)
            doc.SaveAs(abs_pdf, FileFormat=17)
            doc.Close(False)
            word.Quit()
            return
        except Exception as e:
            errors.append(f'Word: {e}')

        # 3) LibreOffice
        try:
            import subprocess
            import shutil
            soffice = shutil.which('soffice')
            if not soffice:
                for p in [r'C:\Program Files\LibreOffice\program\soffice.exe',
                          r'C:\Program Files (x86)\LibreOffice\program\soffice.exe']:
                    if os.path.isfile(p):
                        soffice = p
                        break
            if soffice:
                out_dir = os.path.dirname(abs_pdf)
                subprocess.run([soffice, '--headless', '--convert-to', 'pdf',
                                '--outdir', out_dir, abs_docx],
                               capture_output=True, timeout=60, check=True)
                generated = os.path.join(out_dir,
                    os.path.splitext(os.path.basename(abs_docx))[0] + '.pdf')
                if generated != abs_pdf and os.path.isfile(generated):
                    os.replace(generated, abs_pdf)
                return
            else:
                errors.append('LibreOffice: 未找到 soffice')
        except Exception as e:
            errors.append(f'LibreOffice: {e}')

        raise RuntimeError(
            '无法将 DOCX 转换为 PDF，请安装以下任一软件:\n'
            '  • WPS Office\n  • Microsoft Word\n  • LibreOffice\n\n'
            f'详细错误:\n' + '\n'.join(errors))

    # ---------- 回调 ----------

    def _update_preview_images(self, images, page_count):
        self._preview_images = images
        self._preview_cached_zoom = 0
        self._preview_cached_tk = []
        self._draw_pages_on_canvas()

        try:
            with open(self.md_path.get(), 'r', encoding='utf-8') as f:
                total_chars = len(f.read())
        except Exception:
            total_chars = 0
        self._preview_info.configure(
            text=f'Word 预览  ·  {page_count} 页  ·  {total_chars} 字符')

    def _preview_error(self, msg):
        self._preview_canvas.delete('all')
        self._preview_canvas.create_text(
            20, 30, anchor='nw', fill='#CC3333',
            font=(UI_FONT, 12),
            text=f'预览生成失败:\n{msg}')
        self._preview_info.configure(text='预览失败')

    def _set_textbox(self, textbox, text):
        textbox.configure(state='normal')
        textbox.delete('1.0', 'end')
        textbox.insert('1.0', text)
        textbox.configure(state='disabled')
