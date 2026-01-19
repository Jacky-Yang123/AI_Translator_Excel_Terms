import os
import sys
import streamlit as st
import warnings
warnings.filterwarnings('ignore')

from pages.ytdlp_downloader import ytdlp_downloader_app
from pages.batch_translation import batch_translation_page
from pages.prompt_generator import prompt_generator_page
from pages.excel_replace import excel_replace_page
from pages.jacky import jacky_page
from pages.grand_match import grand_match
from pages.translation_processor import translation_result_processor_page
from pages.excel_comparison import excel_comparison_page
from pages.term_lookup import term_lookup_page
from pages.excel_matchpro import excel_matchpro_page
from pages.danmu import danmu_page
from pages.excel_sreplace import excel_sreplace_page
from pages.excel_abc import excel_ABC_page


def main():
    st.set_page_config(
        page_title="API_AI_Excel翻译分析工具_Jacky",
        page_icon="🎮",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    st.sidebar.title("🎮 多API Excel智能翻译工具")
    st.sidebar.markdown("---")
    
    page = st.sidebar.radio(
        "选择功能页面",
        [
            "📝 提示词生成器",
            "📊 翻译结果处理",
            "🔄 批量翻译工具",
            "术语查找",
            "excel查找替换",
            "excel高级替换",
            "Jacky的主页",
            "🔍 Excel表格对比",
            "🔍 ExcelABC操作",
            "🔍 抓弹幕（只支持nikoniko)",
            "blbl视频弹幕评论下载",
            "文件夹单向匹配程序",
            "模板一键匹配"
        ],
        index=0
    )
    
    st.sidebar.markdown("---")
    st.sidebar.markdown("""
    ### 📖 使用说明
    
    **提示词生成器：**
    - 上传待翻译文本
    - 加载术语库和性格库
    - 生成翻译提示词
    - 复制给AI进行翻译
    
    **翻译结果处理：**
    - 上传原始Excel文件
    - 粘贴AI翻译结果
    - 自动匹配合并
    - 下载完整结果
    
    **批量翻译工具：**
    - 配置API密钥
    - 上传文件和术语库
    - 自动批量翻译
    - 支持重试机制
    
    ### ⚙️ 版本信息
    版本: v2.0 合并版
    作者: Jacky_9S
    """)
    
    if page == "📝 提示词生成器":
        prompt_generator_page()
    elif page == "📊 翻译结果处理":
        translation_result_processor_page()
    elif page == "🔄 批量翻译工具":
        batch_translation_page()
    elif page == '术语查找':
        term_lookup_page()
    elif page == "excel查找替换":
        excel_replace_page()
    elif page == "excel高级替换":
        excel_sreplace_page()
    elif page == "Jacky的主页":
        jacky_page()
    elif page == "🔍 Excel表格对比":
        excel_comparison_page()
    elif page == "🔍 ExcelABC操作":
        excel_ABC_page()
    elif page == "🔍 抓弹幕（只支持nikoniko)":
        danmu_page()
    elif page == "blbl视频弹幕评论下载":
        ytdlp_downloader_app()
    elif page == "文件夹单向匹配程序":
        excel_matchpro_page()
    elif page == "模板一键匹配":
        grand_match()


if __name__ == "__main__":
    try:
        import jieba
    except ImportError:
        print("jieba 库未安装，正在尝试安装...")
        os.system(f"{sys.executable} -m pip install jieba")
        import jieba
    main()
