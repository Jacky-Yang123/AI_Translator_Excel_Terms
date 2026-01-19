import pandas as pd
import streamlit as st


def term_lookup_page():
    st.title("术语查找")
    st.markdown("### 在术语库中查找术语")
    
    uploaded_file = st.file_uploader(
        "📄 上传术语库文件 (Excel)",
        type=['xlsx', 'xls'],
        key="term_lookup_file"
    )
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
            st.success(f"✅ 成功读取术语库，共 {len(df)} 条记录")
            
            with st.expander("📊 术语库预览"):
                st.dataframe(df.head(10))
            
            cols = df.columns.tolist()
            search_col = st.selectbox(
                "选择搜索列:",
                options=cols,
                index=0,
                key="term_lookup_col"
            )
            
            search_term = st.text_input(
                "输入搜索术语:",
                placeholder="请输入要查找的术语...",
                key="term_lookup_search"
            )
            
            if st.button("🔍 查找", key="term_lookup_btn", use_container_width=True):
                if search_term:
                    results = df[df[search_col].astype(str).str.contains(search_term, case=False, na=False)]
                    
                    if not results.empty:
                        st.success(f"✅ 找到 {len(results)} 条匹配记录")
                        st.dataframe(results)
                    else:
                        st.warning("⚠️ 未找到匹配的术语")
        except Exception as e:
            st.error(f"❌ 文件读取失败: {e}")
