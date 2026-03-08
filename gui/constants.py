"""GUI 常量与主题设置"""
import sys
import os
import customtkinter as ctk


def get_resource_path(*parts):
    """获取资源文件路径，兼容 PyInstaller 打包和开发模式"""
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


# ========== 主题 ==========
THEME_PATH = get_resource_path('theme_breeze.json')
ctk.set_appearance_mode('System')
if os.path.isfile(THEME_PATH):
    ctk.set_default_color_theme(THEME_PATH)
else:
    ctk.set_default_color_theme('blue')

# ========== 字体列表 ==========
CN_FONTS = ['黑体', '宋体', '仿宋', '仿宋_GB2312', '楷体', '楷体_GB2312', '微软雅黑']
EN_FONTS = ['Times New Roman', 'Arial', 'Calibri', 'Cambria']

# ========== 标题 ==========
HEADING_LEVELS = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']
HEADING_LABELS = ['H1  一级标题', 'H2  二级标题', 'H3  三级标题',
                  'H4  四级标题', 'H5  五级标题', 'H6  六级标题']
DEFAULT_SIZES = [16, 14, 13, 12, 11, 11]

# ========== 颜色 ==========
ACCENT = '#5856D6'
ACCENT_HOVER = '#4A48C4'
SECTION_TITLE_LIGHT = '#5856D6'
SECTION_TITLE_DARK = '#8E8CE8'
MUTED_LIGHT = '#48484a'
MUTED_DARK = '#b0b0b5'

# ========== UI 字体 ==========
UI_FONT = 'Microsoft YaHei UI'
