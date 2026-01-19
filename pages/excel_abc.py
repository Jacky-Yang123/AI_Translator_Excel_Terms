import os
from io import BytesIO
import zipfile

import pandas as pd
import streamlit as st


def excel_ABC_page():
    st.title("📊 Excel批量处理工具")
    
    if 'excel_files' not in st.session_state:
        st.session_state.excel_files = []
    if 'dataframes' not in st.session_state:
        st.session_state.dataframes = {}
    
    def load_excel_file(file_path):
        try:
            df = pd.read_excel(file_path)
            return df
        except Exception as e:
            st.error(f"读取文件失败 {file_path}: {str(e)}")
            return None
    
    def check_condition(value, keywords, match_mode):
        if not keywords:
            return True
        
        value_str = str(value).lower()
        keywords_list = [kw.strip().lower() for kw in keywords if kw.strip()]
        
        if not keywords_list:
            return True
        
        if match_mode == "同时包含所有关键词":
            return all(kw in value_str for kw in keywords_list)
        else:
            return any(kw in value_str for kw in keywords_list)
    
    def process_dataframe(df, col1, col2, keywords, match_mode, operation, params):
        df_copy = df.copy()
        modified_count = 0
        
        for idx, row in df_copy.iterrows():
            if check_condition(row[col1], keywords, match_mode):
                if operation == "删除值":
                    target_value = params.get('target_value', '')
                    if target_value:
                        cell_value = str(row[col2])
                        if target_value in cell_value:
                            df_copy.at[idx, col2] = cell_value.replace(target_value, '')
                            modified_count += 1
                        
                elif operation == "替换值":
                    old_value = params.get('old_value', '')
                    new_value = params.get('new_value', '')
                    if old_value:
                        cell_value = str(row[col2])
                        if old_value in cell_value:
                            df_copy.at[idx, col2] = cell_value.replace(old_value, new_value)
                            modified_count += 1
                        
                elif operation == "修改中间值":
                    value_a = params.get('value_a', '')
                    value_c = params.get('value_c', '')
                    new_value = params.get('new_value', '')
                    
                    cell_value = str(row[col2])
                    if value_a and value_c and value_a in cell_value and value_c in cell_value:
                        pos_a = cell_value.find(value_a)
                        pos_c = cell_value.find(value_c, pos_a + len(value_a))
                        
                        if pos_c > pos_a:
                            before = cell_value[:pos_a + len(value_a)]
                            after = cell_value[pos_c:]
                            df_copy.at[idx, col2] = before + new_value + after
                            modified_count += 1
        
        return df_copy, modified_count
    
    st.header("1️⃣ 上传文件夹")
    uploaded_files = st.file_uploader(
        "选择Excel文件（可多选）", 
        type=['xlsx', 'xls'], 
        accept_multiple_files=True,
        key="excel_uploader"
    )
    
    if uploaded_files:
        st.session_state.excel_files = []
        st.session_state.dataframes = {}
        
        for uploaded_file in uploaded_files:
            temp_path = f"temp_{uploaded_file.name}"
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            df = load_excel_file(temp_path)
            if df is not None:
                st.session_state.excel_files.append(uploaded_file.name)
                st.session_state.dataframes[uploaded_file.name] = df
            
            os.remove(temp_path)
        
        st.success(f"已加载 {len(st.session_state.excel_files)} 个Excel文件")
        
        with st.expander("查看已加载的文件"):
            for file_name in st.session_state.excel_files:
                st.write(f"- {file_name}")
    
    if st.session_state.excel_files:
        st.header("2️⃣ 配置处理规则")
        
        sample_file = st.session_state.excel_files[0]
        sample_df = st.session_state.dataframes[sample_file]
        columns = list(sample_df.columns)
        
        col_left, col_right = st.columns(2)
        
        with col_left:
            st.subheader("选择列")
            col1 = st.selectbox("第一列（条件列）", columns, key="col1")
            col2 = st.selectbox("第二列（操作列）", columns, key="col2")
        
        with col_right:
            st.subheader("条件设置")
            keywords_input = st.text_area(
                "关键词（每行一个，留空则处理所有行）",
                height=100,
                placeholder="输入关键词\n可输入多个\n每行一个",
                key="keywords_input"
            )
            keywords = [kw.strip() for kw in keywords_input.split('\n') if kw.strip()]
            
            match_mode = st.radio(
                "匹配模式",
                ["同时包含所有关键词", "包含任意一个关键词"],
                disabled=len(keywords) == 0,
                key="match_mode"
            )
        
        st.divider()
        
        st.subheader("选择操作")
        operation = st.radio(
            "操作类型",
            ["删除值", "替换值", "修改中间值"],
            key="operation"
        )
        
        params = {}
        
        if operation == "删除值":
            st.info("💡 删除第二列文本中包含的指定内容")
            params['target_value'] = st.text_input(
                "要删除的内容", 
                key="delete_value", 
                placeholder="例如：删除'帅'，则'我好帅'变为'我好'"
            )
            
        elif operation == "替换值":
            st.info("💡 将第二列文本中的某个内容替换为新内容")
            col_a, col_b = st.columns(2)
            with col_a:
                params['old_value'] = st.text_input(
                    "要替换的内容", 
                    key="old_value",
                    placeholder="例如：帅"
                )
            with col_b:
                params['new_value'] = st.text_input(
                    "替换为", 
                    key="new_value",
                    placeholder="例如：丑"
                )
                
        elif operation == "修改中间值":
            st.info("💡 修改夹在两个值之间的内容")
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                params['value_a'] = st.text_input("起始值 A", key="value_a")
            with col_b:
                params['value_c'] = st.text_input("结束值 C", key="value_c")
            with col_c:
                params['new_value'] = st.text_input("新的中间值", key="middle_new_value")
        
        st.divider()
        
        st.header("3️⃣ 预览和执行")
        
        col_preview, col_execute = st.columns(2)
        
        with col_preview:
            if st.button("🔍 预览效果（使用第一个文件）", type="secondary", use_container_width=True):
                preview_df = st.session_state.dataframes[sample_file].copy()
                processed_df, count = process_dataframe(
                    preview_df, col1, col2, keywords, match_mode, operation, params
                )
                
                st.success(f"预览完成！共修改 {count} 行数据")
                
                col_before, col_after = st.columns(2)
                with col_before:
                    st.write("**处理前**")
                    st.dataframe(preview_df[[col1, col2]].head(20), use_container_width=True)
                with col_after:
                    st.write("**处理后**")
                    st.dataframe(processed_df[[col1, col2]].head(20), use_container_width=True)
        
        with col_execute:
            if st.button("✅ 批量处理所有文件", type="primary", use_container_width=True):
                with st.spinner("正在处理..."):
                    results = {}
                    total_modified = 0
                    
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    for idx, file_name in enumerate(st.session_state.excel_files):
                        status_text.text(f"正在处理: {file_name}")
                        df = st.session_state.dataframes[file_name]
                        processed_df, count = process_dataframe(
                            df, col1, col2, keywords, match_mode, operation, params
                        )
                        results[file_name] = processed_df
                        total_modified += count
                        progress_bar.progress((idx + 1) / len(st.session_state.excel_files))
                    
                    status_text.empty()
                    progress_bar.empty()
                    
                    st.success(f"✅ 处理完成！共修改 {total_modified} 行数据")
                    
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for file_name, df in results.items():
                            excel_buffer = BytesIO()
                            df.to_excel(excel_buffer, index=False)
                            zip_file.writestr(f"processed_{file_name}", excel_buffer.getvalue())
                    
                    st.download_button(
                        label="📥 下载处理后的文件（ZIP）",
                        data=zip_buffer.getvalue(),
                        file_name="processed_excel_files.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
    
    else:
        st.info("👆 请先上传Excel文件")
    
    with st.expander("📖 使用说明"):
        st.markdown("""
        ### 功能说明
        
        1. **上传文件**：选择多个Excel文件（支持.xlsx和.xls格式）
        
        2. **选择列**：
           - 第一列：条件列，用于判断是否符合处理条件
           - 第二列：操作列，对符合条件的行进行操作
        
        3. **设置条件**：
           - 输入关键词（每行一个）
           - 留空表示处理所有行
           - 选择匹配模式：同时包含所有关键词 或 包含任意一个关键词
        
        4. **选择操作**：
           - **删除值**：删除第二列文本中包含的指定内容（例如：删除"帅"，"我好帅"变为"我好"）
           - **替换值**：将第二列文本中的某个内容替换为新内容（例如："帅"替换为"丑"，"我好帅"变为"我好丑"）
           - **修改中间值**：修改夹在A和C之间的B值（例如：A="<"，C=">"，将"<旧值>"改为"<新值>"）
        
        5. **预览和执行**：
           - 先预览第一个文件的处理效果
           - 确认无误后批量处理所有文件
           - 下载处理后的文件压缩包
        
        ### 使用示例
        
        **示例1：删除指定内容**
        - 条件：第一列包含"产品"
        - 操作：删除第二列中的"旧版"
        - 结果："旧版产品说明" → "产品说明"
        
        **示例2：替换内容**
        - 条件：第一列包含"评价"
        - 操作：将"帅"替换为"丑"
        - 结果："这个人好帅" → "这个人好丑"
        
        **示例3：修改中间值**
        - 条件：第一列包含"标签"
        - 操作：A="【"，C="】"，新值="已处理"
        - 结果："【待处理】任务" → "【已处理】任务"
        """)
