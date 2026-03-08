"""
配置加载模块
"""
import sys
import os
import copy

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import yaml


# ========== 配置加载 ==========

DEFAULT_CONFIG = {
    'page': {
        'top_margin_cm': 2.54, 'bottom_margin_cm': 2.54,
        'left_margin_cm': 3.17, 'right_margin_cm': 3.17,
    },
    'body': {
        'font_cn': '仿宋_GB2312', 'font_en': 'Times New Roman',
        'font_size_pt': 12, 'line_spacing': 1.0,
        'first_line_indent_cm': 0.74,
    },
    'headings': {
        'h1': {'font_cn': '黑体', 'size_pt': 16, 'bold': True},
        'h2': {'font_cn': '黑体', 'size_pt': 14, 'bold': True},
        'h3': {'font_cn': '黑体', 'size_pt': 13, 'bold': True},
        'h4': {'font_cn': '黑体', 'size_pt': 12, 'bold': True},
        'h5': {'font_cn': '黑体', 'size_pt': 11, 'bold': True},
        'h6': {'font_cn': '黑体', 'size_pt': 11, 'bold': True},
        'auto_numbering': True,
    },
    'table': {
        'header_font_cn': '黑体', 'header_font_size_pt': 11, 'header_bold': True,
        'body_font_cn': '仿宋_GB2312', 'body_font_size_pt': 11,
    },
    'formula': {
        'numbering': 'chapter',
    },
    'list': {
        'style': 'academic',
    },
    'caption': {'font_cn': '仿宋_GB2312', 'font_size_pt': 10},
    'header_footer': {
        'header_text': '',
        'header_font_cn': '仿宋_GB2312',
        'header_font_size_pt': 9,
        'page_number': True,
        'page_number_position': 'bottom_center',
        'page_number_font_size_pt': 9,
    },
}


def _deep_merge(base, override):
    """递归合并配置，override 覆盖 base"""
    result = dict(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


def load_config(config_path=None):
    """加载 YAML 配置，找不到则使用默认值"""
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    if config_path and os.path.isfile(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            user_cfg = yaml.safe_load(f) or {}
        cfg = _deep_merge(cfg, user_cfg)
        print(f'已加载配置: {config_path}')
    else:
        # 尝试项目根目录下的 md2docx.yaml
        if getattr(sys, 'frozen', False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
        auto_path = os.path.join(base, 'md2docx.yaml')
        if os.path.isfile(auto_path):
            with open(auto_path, 'r', encoding='utf-8') as f:
                user_cfg = yaml.safe_load(f) or {}
            cfg = _deep_merge(cfg, user_cfg)
            print(f'已加载配置: {auto_path}')
    return cfg
