"""
Pages module - 包含所有页面功能

此模块包含了应用程序的所有页面功能，每个页面都有独立的文件。
"""

from . import prompt_generator
from . import batch_translation
from . import ytdlp_downloader
from . import translation_processor
from . import excel_comparison
from . import term_lookup
from . import excel_matchpro
from . import danmu
from . import excel_sreplace
from . import excel_abc
from . import excel_replace
from . import jacky
from . import grand_match

__all__ = [
    'prompt_generator',
    'batch_translation',
    'ytdlp_downloader',
    'translation_processor',
    'excel_comparison',
    'term_lookup',
    'excel_matchpro',
    'danmu',
    'excel_sreplace',
    'excel_abc',
    'excel_replace',
    'jacky',
    'grand_match'
]
