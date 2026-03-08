"""
converter 包 — Markdown → Word 转换器
"""
from .config import DEFAULT_CONFIG, _deep_merge, load_config
from .convert import convert

__all__ = ['convert', 'load_config', 'DEFAULT_CONFIG', '_deep_merge']
