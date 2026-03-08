"""MD → DOCX 转换器入口"""
import sys
import os
import traceback
import tkinter

# 确保项目根目录在 sys.path 中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui.app import Md2DocxApp


def _report_callback_exception(exc_type, exc_value, exc_tb):
    msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_tb))
    print(msg)


if __name__ == '__main__':
    tkinter.Tk.report_callback_exception = _report_callback_exception
    app = Md2DocxApp()
    app.mainloop()
