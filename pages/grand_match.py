import streamlit as st


def grand_match():
    """
    模板一键匹配页面
    
    注意：此功能需要外部模块 model_GRAND_match.model_grand_match
    如果该模块不存在，将显示提示信息
    """
    st.title("🔗 模板一键匹配工具")
    st.markdown("### 批量匹配Excel文件中的模板数据")
    
    try:
        import model_GRAND_match.model_grand_match
        model_GRAND_match.model_grand_match.grand_match()
    except ImportError:
        st.error("❌ 缺少必要的模块")
        st.warning("""
        **model_GRAND_match** 模块未找到
        
        此功能需要安装额外的模块才能使用。请确保：
        1. 已安装 model_GRAND_match 模块
        2. 模块路径正确
        3. 所有依赖项都已安装
        
        如需使用此功能，请联系开发者获取相关模块。
        """)
        
        st.info("💡 提示：您可以使用其他可用的功能，如批量翻译工具、Excel查找替换等")
