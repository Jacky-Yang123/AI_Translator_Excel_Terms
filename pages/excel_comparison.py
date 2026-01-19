import pandas as pd
import streamlit as st


def excel_comparison_page():
    st.title("🔍 Excel表格对比工具")
    st.markdown("### 对比两个Excel文件的差异")
    
    col1, col2 = st.columns(2)
    
    with col1:
        uploaded_file1 = st.file_uploader(
            "📄 上传第一个Excel文件",
            type=['xlsx', 'xls'],
            key="comparison_file1"
        )
        
        if uploaded_file1 is not None:
            try:
                df1 = pd.read_excel(uploaded_file1)
                df1.columns = df1.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                st.success(f"✅ 成功读取文件1，共 {len(df1)} 行数据")
                
                with st.expander("📊 文件1预览"):
                    st.dataframe(df1.head(10))
            except Exception as e:
                st.error(f"❌ 文件读取失败: {e}")
    
    with col2:
        uploaded_file2 = st.file_uploader(
            "📄 上传第二个Excel文件",
            type=['xlsx', 'xls'],
            key="comparison_file2"
        )
        
        if uploaded_file2 is not None:
            try:
                df2 = pd.read_excel(uploaded_file2)
                df2.columns = df2.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                st.success(f"✅ 成功读取文件2，共 {len(df2)} 行数据")
                
                with st.expander("📊 文件2预览"):
                    st.dataframe(df2.head(10))
            except Exception as e:
                st.error(f"❌ 文件读取失败: {e}")
    
    if uploaded_file1 is not None and uploaded_file2 is not None:
        st.header("🔄 对比设置")
        
        cols1 = df1.columns.tolist()
        cols2 = df2.columns.tolist()
        
        key_col1 = st.selectbox(
            "选择文件1的关键列:",
            options=cols1,
            index=0,
            key="comparison_key_col1"
        )
        
        key_col2 = st.selectbox(
            "选择文件2的关键列:",
            options=cols2,
            index=0,
            key="comparison_key_col2"
        )
        
        if st.button("🚀 开始对比", key="compare_btn", use_container_width=True):
            try:
                merged = pd.merge(df1, df2, left_on=key_col1, right_on=key_col2, how='outer', indicator=True, suffixes=('_file1', '_file2'))
                
                only_in_file1 = merged[merged['_merge'] == 'left_only']
                only_in_file2 = merged[merged['_merge'] == 'right_only']
                in_both = merged[merged['_merge'] == 'both']
                
                st.success("✅ 对比完成！")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("仅在文件1中", len(only_in_file1))
                
                with col2:
                    st.metric("仅在文件2中", len(only_in_file2))
                
                with col3:
                    st.metric("两个文件都有", len(in_both))
                
                if not only_in_file1.empty:
                    st.subheader("📋 仅在文件1中的记录")
                    st.dataframe(only_in_file1.head(20))
                
                if not only_in_file2.empty:
                    st.subheader("📋 仅在文件2中的记录")
                    st.dataframe(only_in_file2.head(20))
                
                if not in_both.empty:
                    st.subheader("📋 两个文件都有的记录")
                    st.dataframe(in_both.head(20))
                
            except Exception as e:
                st.error(f"❌ 对比失败: {e}")
