# app.py - 主入口文件（模块化版本）
# 
# 这是应用程序的主入口点，负责页面配置和路由
# 所有功能模块已拆分到独立文件中：
#   - utils.py: 共享工具函数
#   - translator.py: 翻译器类
#   - api_config.py: API配置
#   - pages/: 各页面模块

import streamlit as st

# 检查并安装必要的依赖
try:
    import jieba
except ImportError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "jieba"])
    import jieba

# 导入页面模块
from pages.prompt_generator import prompt_generator_page
from pages.translation_result import translation_result_processor_page
from pages.batch_translation import batch_translation_page
from pages.term_lookup import term_lookup_page
from pages.excel_replace import excel_replace_page
from pages.excel_sreplace import excel_sreplace_page
from pages.excel_comparison import excel_comparison_page
from pages.excel_abc import excel_ABC_page
from pages.danmu import danmu_page
from pages.ytdlp_downloader import ytdlp_downloader_app
from pages.excel_matchpro import excel_matchpro_page
from pages.grand_match import grand_match
from pages.jacky import jacky_page
from pages.format_factory import format_factory_page


def main():
    """主函数 - 设置页面配置和路由"""
    
    st.set_page_config(
        page_title="API_AI_Excel翻译分析工具_Jacky",
        page_icon="🎮",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # 黑白渐变 ins 简约风格 CSS
    st.markdown("""
    <style>
    /* ========== 全局样式 ========== */
    * {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft YaHei', 'PingFang SC', 'Hiragino Sans GB', 'Helvetica Neue', Arial, sans-serif;
    }
    
    /* 主背景 - 黑色渐变 */
    .stApp {
        background: linear-gradient(135deg, #0a0a0a 0%, #1a1a1a 50%, #0f0f0f 100%);
        background-attachment: fixed;
    }
    
    /* ========== 侧边栏样式 ========== */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0d0d0d 0%, #151515 50%, #0d0d0d 100%);
        border-right: 1px solid rgba(255, 255, 255, 0.08);
    }
    
    [data-testid="stSidebar"] .stMarkdown {
        color: #e0e0e0;
    }
    
    /* 侧边栏标题 */
    [data-testid="stSidebar"] h1 {
        background: linear-gradient(90deg, #ffffff 0%, #888888 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-weight: 600;
        letter-spacing: -0.5px;
    }
    
    /* 侧边栏分割线 */
    [data-testid="stSidebar"] hr {
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent);
        margin: 1rem 0;
    }
    
    /* Radio 按钮组样式 */
    [data-testid="stSidebar"] .stRadio > div {
        background: transparent;
    }
    
    [data-testid="stSidebar"] .stRadio label {
        color: #b0b0b0 !important;
        font-weight: 400;
        padding: 0.6rem 1rem;
        border-radius: 8px;
        transition: all 0.2s ease;
        border: 1px solid transparent;
    }
    
    [data-testid="stSidebar"] .stRadio label:hover {
        background: rgba(255, 255, 255, 0.05);
        color: #ffffff !important;
        border-color: rgba(255, 255, 255, 0.1);
    }
    
    [data-testid="stSidebar"] .stRadio label[data-checked="true"],
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label[aria-checked="true"] {
        background: linear-gradient(135deg, rgba(255,255,255,0.1) 0%, rgba(255,255,255,0.05) 100%);
        color: #ffffff !important;
        border-color: rgba(255, 255, 255, 0.2);
    }
    
    /* ========== 主内容区域 ========== */
    .main .block-container {
        padding: 2rem 3rem;
    }
    
    /* 标题样式 */
    h1, h2, h3 {
        background: linear-gradient(90deg, #ffffff 0%, #a0a0a0 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-weight: 600;
        letter-spacing: -0.5px;
    }
    
    h1 {
        font-size: 2.2rem !important;
        margin-bottom: 1.5rem !important;
    }
    
    h2 {
        font-size: 1.6rem !important;
        margin-top: 2rem !important;
    }
    
    h3 {
        font-size: 1.2rem !important;
    }
    
    /* ========== 卡片式容器 ========== */
    .stExpander {
        background: linear-gradient(145deg, rgba(30,30,30,0.8) 0%, rgba(20,20,20,0.9) 100%);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 12px;
        overflow: hidden;
        backdrop-filter: blur(10px);
    }
    
    .stExpander:hover {
        border-color: rgba(255, 255, 255, 0.15);
    }
    
    /* ========== 按钮样式 ========== */
    .stButton > button {
        background: linear-gradient(135deg, #2a2a2a 0%, #1a1a1a 100%);
        color: #ffffff;
        border: 1px solid rgba(255, 255, 255, 0.15);
        border-radius: 8px;
        padding: 0.6rem 1.5rem;
        font-weight: 500;
        letter-spacing: 0.3px;
        transition: all 0.3s ease;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.3);
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, #3a3a3a 0%, #2a2a2a 100%);
        border-color: rgba(255, 255, 255, 0.3);
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.4);
        transform: translateY(-1px);
    }
    
    .stButton > button:active {
        transform: translateY(0);
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.3);
    }
    
    /* 主要按钮 (Primary) */
    .stButton > button[kind="primary"],
    .stDownloadButton > button {
        background: linear-gradient(135deg, #ffffff 0%, #e0e0e0 100%);
        color: #0a0a0a;
        border: none;
        font-weight: 600;
    }
    
    .stButton > button[kind="primary"]:hover,
    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #f0f0f0 0%, #d0d0d0 100%);
        box-shadow: 0 4px 20px rgba(255, 255, 255, 0.2);
    }
    
    /* ========== 输入框样式 ========== */
    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea,
    .stNumberInput > div > div > input,
    .stSelectbox > div > div {
        background: rgba(20, 20, 20, 0.8);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 8px;
        color: #e0e0e0;
        font-size: 0.95rem;
        transition: all 0.2s ease;
    }
    
    .stTextInput > div > div > input:focus,
    .stTextArea > div > div > textarea:focus,
    .stNumberInput > div > div > input:focus {
        border-color: rgba(255, 255, 255, 0.3);
        box-shadow: 0 0 0 2px rgba(255, 255, 255, 0.1);
        outline: none;
    }
    
    /* ========== 文件上传器 ========== */
    .stFileUploader {
        background: rgba(20, 20, 20, 0.6);
        border: 2px dashed rgba(255, 255, 255, 0.15);
        border-radius: 12px;
        padding: 1rem;
        transition: all 0.3s ease;
    }
    
    .stFileUploader:hover {
        border-color: rgba(255, 255, 255, 0.3);
        background: rgba(30, 30, 30, 0.6);
    }
    
    /* ========== 进度条 ========== */
    .stProgress > div > div > div {
        background: linear-gradient(90deg, #ffffff 0%, #888888 100%);
        border-radius: 10px;
    }
    
    .stProgress > div > div {
        background: rgba(255, 255, 255, 0.1);
        border-radius: 10px;
    }
    
    /* ========== 数据表格 ========== */
    .stDataFrame {
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 10px;
        overflow: hidden;
    }
    
    .stDataFrame [data-testid="stDataFrameResizable"] {
        background: rgba(15, 15, 15, 0.9);
    }
    
    /* ========== Tab 样式 ========== */
    .stTabs [data-baseweb="tab-list"] {
        background: transparent;
        gap: 0.5rem;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: rgba(30, 30, 30, 0.6);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 8px;
        color: #b0b0b0;
        padding: 0.5rem 1rem;
        font-weight: 500;
        transition: all 0.2s ease;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        background: rgba(40, 40, 40, 0.8);
        color: #ffffff;
        border-color: rgba(255, 255, 255, 0.15);
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, rgba(255,255,255,0.15) 0%, rgba(255,255,255,0.05) 100%) !important;
        color: #ffffff !important;
        border-color: rgba(255, 255, 255, 0.2) !important;
    }
    
    .stTabs [data-baseweb="tab-highlight"] {
        display: none;
    }
    
    .stTabs [data-baseweb="tab-border"] {
        display: none;
    }
    
    /* ========== 警告和提示框 ========== */
    .stAlert {
        background: rgba(25, 25, 25, 0.9);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 10px;
        backdrop-filter: blur(10px);
    }
    
    .stSuccess {
        border-left: 4px solid #4ade80;
    }
    
    .stInfo {
        border-left: 4px solid #60a5fa;
    }
    
    .stWarning {
        border-left: 4px solid #fbbf24;
    }
    
    .stError {
        border-left: 4px solid #f87171;
    }
    
    /* ========== 滚动条样式 ========== */
    ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
    }
    
    ::-webkit-scrollbar-track {
        background: rgba(0, 0, 0, 0.2);
        border-radius: 4px;
    }
    
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(180deg, #3a3a3a 0%, #2a2a2a 100%);
        border-radius: 4px;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(180deg, #4a4a4a 0%, #3a3a3a 100%);
    }
    
    /* ========== 指标卡片 ========== */
    [data-testid="stMetric"] {
        background: linear-gradient(145deg, rgba(30,30,30,0.8) 0%, rgba(20,20,20,0.9) 100%);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 12px;
        padding: 1rem;
        backdrop-filter: blur(10px);
    }
    
    [data-testid="stMetricLabel"] {
        color: #888888 !important;
        font-size: 0.9rem !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    [data-testid="stMetricValue"] {
        background: linear-gradient(90deg, #ffffff 0%, #c0c0c0 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-weight: 700 !important;
    }
    
    /* ========== 代码块 ========== */
    .stCodeBlock {
        background: rgba(10, 10, 10, 0.9) !important;
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 10px;
    }
    
    /* ========== 分割线 ========== */
    hr {
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(255,255,255,0.15), transparent);
        margin: 2rem 0;
    }
    
    /* ========== Checkbox 和 Radio ========== */
    .stCheckbox > label,
    .stRadio > label {
        color: #d0d0d0 !important;
        font-weight: 400;
    }
    
    /* ========== Slider ========== */
    .stSlider [data-baseweb="slider"] {
        background: rgba(255, 255, 255, 0.1);
    }
    
    .stSlider [data-testid="stThumbValue"] {
        color: #ffffff;
        background: rgba(30, 30, 30, 0.9);
        border: 1px solid rgba(255, 255, 255, 0.2);
        border-radius: 6px;
        padding: 2px 8px;
    }
    
    /* ========== 链接样式 ========== */
    a {
        color: #ffffff !important;
        text-decoration: none;
        border-bottom: 1px solid rgba(255, 255, 255, 0.3);
        transition: all 0.2s ease;
    }
    
    a:hover {
        border-bottom-color: #ffffff;
    }
    
    /* ========== 多选框 ========== */
    .stMultiSelect [data-baseweb="tag"] {
        background: linear-gradient(135deg, rgba(255,255,255,0.15) 0%, rgba(255,255,255,0.08) 100%);
        border: 1px solid rgba(255, 255, 255, 0.2);
        border-radius: 6px;
        color: #ffffff;
    }
    
    /* ========== 日期选择器 ========== */
    .stDateInput > div > div {
        background: rgba(20, 20, 20, 0.8);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 8px;
    }
    
    /* ========== 工具提示 ========== */
    [data-baseweb="tooltip"] {
        background: rgba(20, 20, 20, 0.95) !important;
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 8px;
        backdrop-filter: blur(10px);
    }
    
    /* ========== 选中效果 ========== */
    ::selection {
        background: rgba(255, 255, 255, 0.2);
        color: #ffffff;
    }
    
    /* ========== 动画效果 ========== */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    .main .block-container > div {
        animation: fadeIn 0.4s ease-out;
    }
    </style>
    """, unsafe_allow_html=True)

    # 侧边栏页面选择
    st.sidebar.title("🎮 多API Excel智能翻译工具")
    st.sidebar.markdown("---\n")

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
            "模板一键匹配",
            "🏭 格式工厂"
        ],
        index=0
    )

    st.sidebar.markdown("---\n")
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
    """)

    # 根据选择显示不同页面
    if page == "📝 提示词生成器":
        prompt_generator_page()
    elif page == "📊 翻译结果处理":
        translation_result_processor_page()
    elif page == "🔄 批量翻译工具":
        batch_translation_page()
    elif page == "术语查找":
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
    elif page == "🏭 格式工厂":
        format_factory_page()


if __name__ == "__main__":
    main()
