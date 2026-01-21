# pages/excel_matchpro.py - 文件夹单向匹配程序

import os
import concurrent.futures
from pathlib import Path
from datetime import datetime
from difflib import SequenceMatcher
from io import BytesIO

import pandas as pd
import streamlit as st


def similar(a, b):
    """计算两个字符串的相似度"""
    return SequenceMatcher(None, str(a), str(b)).ratio()


def load_single_file(file_path):
    """加载单个文件（Excel或CSV）"""
    try:
        if file_path.suffix.lower() in ['.xlsx', '.xls', '.xlsm']:
            excel_file = pd.read_excel(file_path, sheet_name=None)
            results = {}
            for sheet_name, df in excel_file.items():
                if not df.empty:
                    key = f"{file_path.name} - {sheet_name}"
                    results[key] = {
                        'dataframe': df,
                        'file_path': file_path,
                        'sheet_name': sheet_name,
                        'file_type': 'excel'
                    }
            return results
        elif file_path.suffix.lower() == '.csv':
            encodings = ['utf-8', 'gbk', 'gb2312', 'latin1']
            df = None

            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue

            if df is not None and not df.empty:
                return {
                    file_path.name: {
                        'dataframe': df,
                        'file_path': file_path,
                        'sheet_name': 'CSV',
                        'file_type': 'csv'
                    }
                }
    except Exception as e:
        st.warning(f"无法读取文件 {file_path}: {str(e)}")

    return {}


def load_all_files_parallel(folder_path, max_workers=4):
    """并行加载文件夹中的所有Excel和CSV文件"""
    all_files = {}
    folder_path = Path(folder_path)

    file_paths = []
    for pattern in ['*.xlsx', '*.xls', '*.xlsm', '*.csv']:
        file_paths.extend(folder_path.rglob(pattern))

    if not file_paths:
        return all_files

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_path = {executor.submit(load_single_file, path): path for path in file_paths}

        for future in concurrent.futures.as_completed(future_to_path):
            try:
                result = future.result()
                all_files.update(result)
            except Exception as e:
                path = future_to_path[future]
                st.warning(f"处理文件 {path} 时出错: {str(e)}")

    return all_files


def find_matching_text(search_text, files_dict, source_col, target_col, match_strategy, similarity_threshold):
    """在文件字典中查找匹配的文本"""
    if pd.isna(search_text) or search_text == '':
        return None, None, 0

    search_text = str(search_text).strip()

    best_match = None
    best_similarity = 0
    best_source = None

    for file_info in files_dict.values():
        df = file_info['dataframe']

        if source_col in df.columns and target_col in df.columns:
            if match_strategy == "精确匹配":
                matches = df[df[source_col].astype(str).str.strip() == search_text]
                if not matches.empty:
                    return matches[target_col].iloc[0], matches[source_col].iloc[0], 1.0

            elif match_strategy == "模糊匹配":
                for _, row in df.iterrows():
                    source_text = str(row[source_col]).strip()
                    sim = similar(search_text, source_text)

                    if sim > best_similarity and sim >= similarity_threshold:
                        best_similarity = sim
                        best_match = row[target_col]
                        best_source = source_text

            elif match_strategy == "包含匹配":
                for _, row in df.iterrows():
                    source_text = str(row[source_col]).strip()
                    if search_text in source_text or source_text in search_text:
                        return row[target_col], source_text, 0.9

    return best_match, best_source, best_similarity


def excel_matchpro_page():
    """文件夹单向匹配程序"""
    st.title("📁 文件夹单向匹配程序")
    st.markdown("### 在源文件夹中查找匹配的翻译，应用到目标文件")

    # 配置区域
    st.header("⚙️ 配置")

    col1, col2 = st.columns(2)

    with col1:
        source_folder = st.text_input(
            "源文件夹路径（包含翻译参考）",
            placeholder="例如: C:/翻译参考",
            key="source_folder_input"
        )

    with col2:
        target_folder = st.text_input(
            "目标文件夹路径（待匹配）",
            placeholder="例如: C:/待翻译文件",
            key="target_folder_input"
        )

    col1, col2, col3 = st.columns(3)

    with col1:
        source_col = st.text_input(
            "源文件原文列名",
            value="中文",
            key="source_col_input"
        )

    with col2:
        target_col = st.text_input(
            "源文件译文列名",
            value="英文",
            key="target_col_input"
        )

    with col3:
        dest_text_col = st.text_input(
            "目标文件待匹配列名",
            value="中文",
            key="dest_text_col_input"
        )

    col1, col2 = st.columns(2)

    with col1:
        match_strategy = st.selectbox(
            "匹配策略",
            options=["精确匹配", "模糊匹配", "包含匹配"],
            index=0,
            key="match_strategy_select"
        )

    with col2:
        similarity_threshold = st.slider(
            "相似度阈值（模糊匹配时使用）",
            min_value=0.5,
            max_value=1.0,
            value=0.8,
            step=0.05,
            key="similarity_threshold_slider"
        )

    output_col_name = st.text_input(
        "输出列名",
        value="匹配译文",
        key="output_col_name_input"
    )

    # 执行匹配
    if st.button("🚀 开始匹配", type="primary", use_container_width=True):
        if not source_folder or not target_folder:
            st.error("❌ 请输入源文件夹和目标文件夹路径")
            return

        if not Path(source_folder).exists():
            st.error("❌ 源文件夹不存在")
            return

        if not Path(target_folder).exists():
            st.error("❌ 目标文件夹不存在")
            return

        # 加载源文件
        with st.spinner("正在加载源文件..."):
            source_files = load_all_files_parallel(source_folder)

        if not source_files:
            st.error("❌ 源文件夹中没有找到Excel或CSV文件")
            return

        st.success(f"✅ 加载了 {len(source_files)} 个源文件/工作表")

        # 加载目标文件
        with st.spinner("正在加载目标文件..."):
            target_files = load_all_files_parallel(target_folder)

        if not target_files:
            st.error("❌ 目标文件夹中没有找到Excel或CSV文件")
            return

        st.success(f"✅ 加载了 {len(target_files)} 个目标文件/工作表")

        # 执行匹配
        results = []
        total_matched = 0
        total_processed = 0

        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, (file_key, file_info) in enumerate(target_files.items()):
            df = file_info['dataframe'].copy()

            if dest_text_col not in df.columns:
                st.warning(f"⚠️ 文件 {file_key} 中没有列 '{dest_text_col}'，跳过")
                continue

            matched_col = []
            source_matched_col = []
            similarity_col = []

            for _, row in df.iterrows():
                text = row[dest_text_col]
                match_result, source_text, similarity = find_matching_text(
                    text, source_files, source_col, target_col, match_strategy, similarity_threshold
                )

                matched_col.append(match_result)
                source_matched_col.append(source_text)
                similarity_col.append(similarity)

                total_processed += 1
                if match_result is not None:
                    total_matched += 1

            df[output_col_name] = matched_col
            df['匹配源文'] = source_matched_col
            df['相似度'] = similarity_col

            results.append({
                'file_key': file_key,
                'file_info': file_info,
                'result_df': df
            })

            progress_bar.progress((i + 1) / len(target_files))
            status_text.text(f"处理中: {i + 1}/{len(target_files)} 个文件")

        progress_bar.empty()
        status_text.empty()

        st.success(f"✅ 匹配完成！处理 {total_processed} 条，匹配 {total_matched} 条")

        # 显示结果
        st.header("📊 匹配结果")

        for result in results:
            with st.expander(f"📄 {result['file_key']}"):
                st.dataframe(result['result_df'].head(50))

                # 下载单个文件
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result['result_df'].to_excel(writer, index=False, sheet_name='匹配结果')

                st.download_button(
                    label=f"📥 下载 {result['file_key']}",
                    data=output.getvalue(),
                    file_name=f"matched_{result['file_key']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{result['file_key']}"
                )

        # 下载所有结果
        st.header("📥 下载所有结果")

        all_output = BytesIO()
        with pd.ExcelWriter(all_output, engine='openpyxl') as writer:
            for result in results:
                sheet_name = result['file_key'][:31]  # Excel工作表名最大31字符
                result['result_df'].to_excel(writer, index=False, sheet_name=sheet_name)

        st.download_button(
            label="📥 下载所有匹配结果",
            data=all_output.getvalue(),
            file_name=f"all_matched_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
