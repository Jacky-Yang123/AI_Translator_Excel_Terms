# pages 模块初始化文件
# 导出所有页面函数

from .prompt_generator import prompt_generator_page
from .translation_result import translation_result_processor_page
from .batch_translation import batch_translation_page
from .term_lookup import term_lookup_page
from .excel_replace import excel_replace_page
from .excel_sreplace import excel_sreplace_page
from .excel_comparison import excel_comparison_page
from .excel_abc import excel_ABC_page
from .danmu import danmu_page
from .ytdlp_downloader import ytdlp_downloader_app
from .excel_matchpro import excel_matchpro_page
from .grand_match import grand_match
from .jacky import jacky_page

__all__ = [
    'prompt_generator_page',
    'translation_result_processor_page',
    'batch_translation_page',
    'term_lookup_page',
    'excel_replace_page',
    'excel_sreplace_page',
    'excel_comparison_page',
    'excel_ABC_page',
    'danmu_page',
    'ytdlp_downloader_app',
    'excel_matchpro_page',
    'grand_match',
    'jacky_page',
]
