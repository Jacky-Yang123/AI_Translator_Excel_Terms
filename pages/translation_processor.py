import pandas as pd
import streamlit as st


def translation_result_processor_page():
    st.title("📊 翻译结果处理工具")
    st.markdown("### 将AI翻译结果与原始Excel文件合并")
    
    st.header("📁 上传文件")
    
    col1, col2 = st.columns(2)
    
    with col1:
        uploaded_original = st.file_uploader(
            "📄 上传原始Excel文件",
            type=['xlsx', 'xls'],
            key="processor_original_file"
        )
        
        if uploaded_original is not None:
            try:
                df_original = pd.read_excel(uploaded_original)
                df_original.columns = df_original.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                st.success(f"✅ 成功读取原始文件，共 {len(df_original)} 行数据")
                
                with st.expander("📊 原始文件预览"):
                    st.dataframe(df_original.head(10))
            except Exception as e:
                st.error(f"❌ 文件读取失败: {e}")
    
    with col2:
        st.markdown("### 📝 粘贴AI翻译结果")
        st.markdown("请将AI返回的翻译结果粘贴到下方文本框中")
        
        translation_result = st.text_area(
            "翻译结果",
            height=300,
            placeholder="请粘贴AI翻译结果...",
            key="processor_translation_result"
        )
    
    if uploaded_original is not None and translation_result:
        st.header("🔄 处理翻译结果")
        
        try:
            lines = translation_result.strip().split('\n')
            translations = []
            
            for line in lines:
                line = line.strip()
                if line and '|' in line:
                    parts = [p.strip() for p in line.split('|')]
                    if len(parts) >= 2:
                        translations.append({
                            '原文': parts[0],
                            '翻译': parts[1]
                        })
            
            if translations:
                st.success(f"✅ 成功解析 {len(translations)} 条翻译结果")
                
                df_translations = pd.DataFrame(translations)
                
                with st.expander("📊 翻译结果预览"):
                    st.dataframe(df_translations.head(10))
                
                st.header("📊 合并结果")
                
                cols = df_original.columns.tolist()
                target_col = st.selectbox(
                    "选择要添加翻译结果的列:",
                    options=cols,
                    index=0,
                    key="processor_target_col"
                )
                
                if st.button("🚀 合并翻译结果", key="merge_btn", use_container_width=True):
                    df_merged = df_original.copy()
                    df_merged['翻译结果'] = ''
                    
                    for idx, row in df_merged.iterrows():
                        original_text = str(row[target_col]).strip()
                        
                        for trans in translations:
                            if trans['原文'] == original_text:
                                df_merged.at[idx, '翻译结果'] = trans['翻译']
                                break
                    
                    st.success("✅ 合并完成！")
                    
                    with st.expander("📊 合并结果预览"):
                        st.dataframe(df_merged.head(20))
                    
                    output = pd.ExcelWriter('translation_result_merged.xlsx', engine='openpyxl')
                    df_merged.to_excel(output, index=False)
                    output.close()
                    
                    with open('translation_result_merged.xlsx', 'rb') as f:
                        st.download_button(
                            label="💾 下载合并后的Excel文件",
                            data=f.read(),
                            file_name="translation_result_merged.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
            else:
                st.warning("⚠️ 未找到有效的翻译结果")
        except Exception as e:
            st.error(f"❌ 处理失败: {e}")
