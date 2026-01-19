import os
import sys
import re
import json
import time
import shutil
import subprocess
import threading
from datetime import datetime
from pathlib import Path
from io import BytesIO
import zipfile

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


progress_lock = threading.Lock()


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


def load_source_files_parallel(folder_path, max_workers=4):
    """并行加载源文件夹中的Excel和CSV文件"""
    source_files = {}
    folder_path = Path(folder_path)
    
    file_paths = []
    for pattern in ['*.xlsx', '*.xls', '*.xlsm', '*.csv']:
        file_paths.extend(folder_path.glob(pattern))
    
    if not file_paths:
        return source_files
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_path = {executor.submit(load_single_file, path): path for path in file_paths}
        
        for future in concurrent.futures.as_completed(future_to_path):
            try:
                result = future.result()
                source_files.update(result)
            except Exception as e:
                path = future_to_path[future]
                st.error(f"处理文件 {path} 时出错: {str(e)}")
    
    return source_files


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
            else:
                for idx, row in df.iterrows():
                    source_text = str(row[source_col])
                    if pd.isna(source_text) or source_text == '':
                        continue
                    
                    if match_strategy == "包含匹配":
                        if search_text in source_text or source_text in search_text:
                            similarity = similar(search_text, source_text)
                            if similarity > best_similarity:
                                best_similarity = similarity
                                best_match = row[target_col]
                                best_source = source_text
                    else:
                        similarity = similar(search_text, source_text)
                        if similarity > best_similarity and similarity >= similarity_threshold:
                            best_similarity = similarity
                            best_match = row[target_col]
                            best_source = source_text
    
    if match_strategy != "精确匹配" and best_match is not None:
        return best_match, best_source, best_similarity
    
    return None, None, 0


def similar(text1, text2):
    """计算两个文本的相似度（简单实现）"""
    if text1 == text2:
        return 1.0
    
    set1 = set(text1)
    set2 = set(text2)
    
    if not set1 and not set2:
        return 1.0
    
    if not set1 or not set2:
        return 0.0
    
    intersection = len(set1 & set2)
    union = len(set1 | set2)
    
    return intersection / union if union > 0 else 0.0


def process_single_row(args):
    """处理单行数据的匹配"""
    index, row, folder1_match_col, folder1_fill_col, folder2_files, folder2_source_col, folder2_target_col, match_strategy, similarity_threshold, skip_filled = args
    
    if skip_filled and not pd.isna(row.get(folder1_fill_col, None)) and str(row[folder1_fill_col]).strip() != '':
        return index, None, None, None, 0, "跳过已填充"
    
    search_text = row[folder1_match_col]
    
    if pd.isna(search_text) or search_text == '':
        return index, None, None, None, 0, "空值"
    
    matched_text, matched_source, similarity = find_matching_text(
        search_text, folder2_files, folder2_source_col, folder2_target_col, match_strategy, similarity_threshold
    )
    
    match_status = "匹配成功" if matched_text is not None else "未匹配"
    
    return index, matched_text, matched_source, search_text, similarity, match_status


def process_single_file(args):
    """处理单个文件的匹配"""
    filename, file_info, folder1_match_col, folder1_fill_col, folder2_files, folder2_source_col, folder2_target_col, match_strategy, similarity_threshold, skip_filled, thread_id = args
    
    df = file_info['dataframe'].copy()
    file_matches = 0
    file_total = 0
    file_skipped = 0
    
    if folder1_match_col not in df.columns:
        return filename, None, {"error": f"文件 {filename} 中找不到列 '{folder1_match_col}'"}
        
    if folder1_fill_col not in df.columns:
        return filename, None, {"error": f"文件 {filename} 中找不到列 '{folder1_fill_col}'"}
    
    rows_to_process = []
    for index, row in df.iterrows():
        file_total += 1
        rows_to_process.append((index, row, folder1_match_col, folder1_fill_col, folder2_files, folder2_source_col, folder2_target_col, match_strategy, similarity_threshold, skip_filled))
    
    matched_results = {}
    match_details = []
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        future_to_index = {executor.submit(process_single_row, args): args[0] for args in rows_to_process}
        
        for future in concurrent.futures.as_completed(future_to_index):
            try:
                index, matched_text, matched_source, search_text, similarity, match_status = future.result()
                
                match_details.append({
                    'index': index,
                    'search_text': search_text,
                    'matched_text': matched_text,
                    'matched_source': matched_source,
                    'similarity': similarity,
                    'status': match_status
                })
                
                if match_status == "跳过已填充":
                    file_skipped += 1
                elif matched_text is not None:
                    matched_results[index] = matched_text
                    file_matches += 1
            except Exception as e:
                index = future_to_index[future]
                st.warning(f"处理文件 {filename} 的第 {index} 行时出错: {str(e)}")
    
    for index, matched_text in matched_results.items():
        df.at[index, folder1_fill_col] = matched_text
    
    report = {
        'total_rows': file_total,
        'matched_rows': file_matches,
        'unmatched_rows': file_total - file_matches - file_skipped,
        'skipped_rows': file_skipped,
        'match_details': match_details
    }
    
    return filename, df, report


def process_file_matching_parallel(folder1_path, folder2_path, folder1_match_col, folder1_fill_col, 
                                  folder2_source_col, folder2_target_col, match_strategy, 
                                  similarity_threshold, skip_filled, max_workers=4):
    """并行处理文件匹配"""
    
    st.info("正在加载第一个文件夹中的文件...")
    folder1_files = load_source_files_parallel(folder1_path, max_workers)
    
    if not folder1_files:
        st.error("在第一个文件夹中未找到Excel或CSV文件")
        return None, None
    
    st.info("正在加载第二个文件夹中的文件...")
    folder2_files = load_all_files_parallel(folder2_path, max_workers)
    
    if not folder2_files:
        st.error("在第二个文件夹中未找到Excel或CSV文件")
        return None, None
    
    st.success(f"在第一个文件夹中找到 {len(folder1_files)} 个文件")
    st.success(f"在第二个文件夹中找到 {len(folder2_files)} 个文件")
    
    st.info("开始匹配处理...")
    results = {}
    match_report = {
        'total_files': len(folder1_files),
        'total_rows': 0,
        'matched_rows': 0,
        'unmatched_rows': 0,
        'skipped_rows': 0,
        'file_details': {}
    }
    
    files_to_process = []
    for i, (filename, file_info) in enumerate(folder1_files.items()):
        files_to_process.append((
            filename, file_info, folder1_match_col, folder1_fill_col,
            folder2_files, folder2_source_col, folder2_target_col, 
            match_strategy, similarity_threshold, skip_filled, i % max_workers
        ))
    
    progress_bar = st.progress(0)
    processed_count = 0
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_filename = {executor.submit(process_single_file, args): args[0] for args in files_to_process}
        
        for future in concurrent.futures.as_completed(future_to_filename):
            try:
                filename, processed_df, report = future.result()
                if processed_df is not None:
                    results[filename] = processed_df
                    match_report['file_details'][filename] = report
                    match_report['total_rows'] += report['total_rows']
                    match_report['matched_rows'] += report['matched_rows']
                    match_report['unmatched_rows'] += report['unmatched_rows']
                    match_report['skipped_rows'] += report['skipped_rows']
                elif "error" in report:
                    st.error(report["error"])
            except Exception as e:
                filename = future_to_filename[future]
                st.error(f"处理文件 {filename} 时出错: {str(e)}")
            
            with progress_lock:
                processed_count += 1
                progress_bar.progress(processed_count / len(files_to_process))
    
    progress_bar.empty()
    return results, match_report


def save_processed_files(processed_files):
    """保存处理后的文件到ZIP包"""
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for filename, df in processed_files.items():
            if filename.lower().endswith('.csv'):
                csv_buffer = BytesIO()
                df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                csv_buffer.seek(0)
                zip_file.writestr(filename, csv_buffer.getvalue())
            else:
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                excel_buffer.seek(0)
                zip_file.writestr(filename, excel_buffer.getvalue())
    
    return zip_buffer.getvalue()


def excel_matchpro_page():
    st.title("⚡ Excel/CSV文件匹配工具(增强版)")
    st.markdown("""
    这个工具使用多线程技术加速处理，可以快速匹配两个文件夹中的Excel和CSV文件内容。  
    支持精确匹配、包含匹配和相似度匹配，可跳过已有翻译文本的行。
    """)
    
    st.sidebar.header("配置参数")
    
    with st.sidebar.expander("文件夹设置", expanded=True):
        folder1_path = st.text_input("第一个文件夹路径（固定格式文件）", value="./folder1")
        folder2_path = st.text_input("第二个文件夹路径（翻译文件）", value="./folder2")
        max_workers = st.slider("线程数", min_value=1, max_value=16, value=4, step=1)
    
    with st.sidebar.expander("匹配策略设置", expanded=True):
        match_strategy = st.selectbox(
            "匹配策略",
            ["精确匹配", "包含匹配", "相似度匹配"],
            help="精确匹配: 完全相同的文本; 包含匹配: 文本互相包含; 相似度匹配: 基于文本相似度"
        )
        
        similarity_threshold = st.slider(
            "相似度阈值(仅对相似度匹配有效)",
            min_value=0.1,
            max_value=1.0,
            value=0.8,
            step=0.05,
            help="相似度高于此阈值的文本将被视为匹配"
        )
        
        skip_filled = st.checkbox(
            "跳过已有翻译文本的行",
            value=True,
            help="如果目标列已有内容，则跳过该行不进行匹配"
        )
    
    with st.sidebar.expander("列映射设置", expanded=True):
        st.markdown("**第一个文件夹列设置**")
        folder1_match_col = st.text_input("匹配列名（用于查找的列）", value="中文文本")
        folder1_fill_col = st.text_input("填充列名（要填入翻译的列）", value="英文文本")
        
        st.markdown("**第二个文件夹列设置**")
        folder2_source_col = st.text_input("原文列名", value="原文")
        folder2_target_col = st.text_input("翻译列名", value="翻译结果")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        if st.button("开始处理", type="primary", use_container_width=True):
            if not folder1_path or not folder2_path:
                st.error("请填写两个文件夹的路径")
                return
                
            if not os.path.exists(folder1_path):
                st.error(f"第一个文件夹路径不存在: {folder1_path}")
                return
                
            if not os.path.exists(folder2_path):
                st.error(f"第二个文件夹路径不存在: {folder2_path}")
                return
            
            start_time = time.time()
            with st.spinner("正在处理文件匹配..."):
                processed_files, match_report = process_file_matching_parallel(
                    folder1_path, folder2_path, 
                    folder1_match_col, folder1_fill_col,
                    folder2_source_col, folder2_target_col,
                    match_strategy, similarity_threshold, skip_filled, max_workers
                )
            
            end_time = time.time()
            
            if processed_files is not None:
                st.success(f"处理完成！耗时: {end_time - start_time:.2f} 秒")
                
                st.subheader("匹配报告")
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("总文件数", match_report['total_files'])
                    st.metric("总行数", match_report['total_rows'])
                
                with col2:
                    st.metric("匹配行数", match_report['matched_rows'])
                    match_rate = (match_report['matched_rows'] / (match_report['total_rows'] - match_report['skipped_rows']) * 100) if (match_report['total_rows'] - match_report['skipped_rows']) > 0 else 0
                    st.metric("匹配率", f"{match_rate:.1f}%")
                
                with col3:
                    st.metric("未匹配行数", match_report['unmatched_rows'])
                    st.metric("跳过行数", match_report['skipped_rows'])
                
                with col4:
                    st.metric("处理速度", f"{match_report['total_rows'] / (end_time - start_time):.1f} 行/秒")
                
                st.subheader("文件详情")
                for filename, details in match_report['file_details'].items():
                    with st.expander(f"📄 {filename}", expanded=False):
                        col1, col2, col3, col4 = st.columns(4)
                        col1.metric("总行数", details['total_rows'])
                        col2.metric("匹配", details['matched_rows'])
                        col3.metric("未匹配", details['unmatched_rows'])
                        col4.metric("跳过", details['skipped_rows'])
                
                st.subheader("下载结果")
                zip_data = save_processed_files(processed_files)
                st.download_button(
                    label="📥 下载处理后的文件（ZIP）",
                    data=zip_data,
                    file_name="processed_files.zip",
                    mime="application/zip",
                    use_container_width=True
                )


def find_yt_dlp():
    """查找yt-dlp可执行文件"""
    if sys.platform.startswith('win'):
        candidates = ["yt-dlp.exe", "yt-dlp"]
    else:
        candidates = ["./yt-dlp", "yt-dlp"]
    
    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate
    return "yt-dlp"


def normalize_url(url):
    """将非标准的Niconico URL格式转换为标准格式"""
    return url.replace("www.video.nicovideo.jp", "www.nicovideo.jp")


def extract_watch_id(url):
    """从Niconico视频URL中提取watch ID (sm/nm号)"""
    match = re.search(r'(sm|nm)\d+', url)
    if match:
        return match.group(0)
    return "unknown_id"


def extract_bilibili_id(url):
    """从Bilibili视频URL中提取video ID (BV号或av号)"""
    bv_match = re.search(r'BV[a-zA-Z0-9]+', url)
    if bv_match:
        return bv_match.group(0)
    
    av_match = re.search(r'av(\d+)', url)
    if av_match:
        return f"av{av_match.group(1)}"
    
    return "unknown_id"


def run_yt_dlp_to_get_json(url, output_filename_base="danmaku"):
    """运行yt-dlp命令来抓取弹幕数据并保存为JSON文件"""
    yt_dlp_path = find_yt_dlp()
    
    command = [
        yt_dlp_path,
        "--skip-download",
        "--write-sub",
        "--all-subs",
        "--sub-format", "json",
        "--output", f"{output_filename_base}.%(ext)s",
        url
    ]
    
    try:
        result = subprocess.run(command, capture_output=True, text=True, check=True)
        
        json_filename = f"{output_filename_base}.comments.json"
        if os.path.exists(json_filename):
            return json_filename
        else:
            return None
            
    except subprocess.CalledProcessError as e:
        st.error(f"yt-dlp执行失败: {e.stderr}")
        return None
    except FileNotFoundError:
        st.error(f"找不到yt-dlp可执行文件。请确保yt-dlp已安装或在PATH中。")
        return None


def process_niconico_json_to_dataframe(json_path):
    """读取yt-dlp生成的JSON文件，处理Niconico弹幕数据"""
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        st.error(f"JSON文件处理失败: {e}")
        return None

    danmaku_list = []
    for comment in data:
        vpos_ms = comment.get("vposMs", 0)
        time_sec = vpos_ms / 1000
        video_time = time.strftime('%H:%M:%S', time.gmtime(time_sec))
        
        posted_at_str = comment.get("postedAt")
        try:
            posted_at = datetime.fromisoformat(posted_at_str)
            send_time = posted_at.strftime('%Y-%m-%d %H:%M:%S')
        except:
            send_time = posted_at_str
            
        commands = " ".join(comment.get("commands", []))
        
        danmaku_info = {
            "弹幕内容": comment.get("body"),
            "视频时间": video_time,
            "时间(秒)": time_sec,
            "格式/颜色": commands,
            "用户ID": comment.get("userId"),
            "发送时间": send_time,
            "编号": comment.get("no"),
        }
        danmaku_list.append(danmaku_info)
        
    if not danmaku_list:
        return None
        
    df = pd.DataFrame(danmaku_list)
    df = df[['编号', '视频时间', '时间(秒)', '弹幕内容', '格式/颜色', '用户ID', '发送时间']]
    return df


def scrape_niconico_danmaku(url):
    """抓取Niconico弹幕"""
    normalized_url = normalize_url(url)
    watch_id = extract_watch_id(normalized_url)
    
    with st.spinner(f"正在抓取Niconico视频 {watch_id} 的弹幕..."):
        json_path = run_yt_dlp_to_get_json(normalized_url, output_filename_base=watch_id)
        
        if json_path:
            df = process_niconico_json_to_dataframe(json_path)
            
            try:
                os.remove(json_path)
            except OSError:
                pass
            
            return df, watch_id
        else:
            return None, watch_id


def scrape_bilibili_danmaku(url, cookies_file=None):
    """抓取Bilibili弹幕"""
    video_id = extract_bilibili_id(url)
    
    with st.spinner(f"正在抓取Bilibili视频 {video_id} 的弹幕..."):
        yt_dlp_path = find_yt_dlp()
        
        command = [
            yt_dlp_path,
            "--skip-download",
            "--write-sub",
            "--all-subs",
            "--sub-format", "json",
            "--output", f"{video_id}.%(ext)s",
        ]
        
        if cookies_file and os.path.exists(cookies_file):
            command.extend(["--cookies", cookies_file])
        
        command.append(url)
        
        try:
            result = subprocess.run(command, capture_output=True, text=True, check=True, timeout=60)
            
            json_filename = f"{video_id}.comments.json"
            if os.path.exists(json_filename):
                try:
                    with open(json_filename, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                except (FileNotFoundError, json.JSONDecodeError) as e:
                    st.error(f"JSON文件处理失败: {e}")
                    return None, video_id
                
                danmaku_list = []
                for comment in data:
                    danmaku_info = {
                        "弹幕内容": comment.get("body", comment.get("text", "")),
                        "发送时间": comment.get("postedAt", comment.get("timestamp", "")),
                        "用户ID": comment.get("userId", comment.get("author", "")),
                    }
                    danmaku_list.append(danmaku_info)
                
                if danmaku_list:
                    df = pd.DataFrame(danmaku_list)
                    
                    try:
                        os.remove(json_filename)
                    except OSError:
                        pass
                    
                    return df, video_id
                else:
                    st.warning("未找到弹幕数据")
                    return None, video_id
            else:
                st.error(f"yt-dlp未生成预期的文件")
                return None, video_id
                
        except subprocess.CalledProcessError as e:
            error_msg = e.stderr if e.stderr else str(e)
            if "412" in error_msg or "Precondition Failed" in error_msg:
                st.error(
                    "❌ Bilibili API速率限制（HTTP 412错误）\n\n"
                    "这通常是因为：\n"
                    "1. 没有提供有效的Cookie认证\n"
                    "2. 该IP地址的请求已达到限制\n\n"
                    "**解决方案：**\n"
                    "请在左侧栏上传您的Bilibili Cookie文件，然后重试。"
                )
            else:
                st.error(f"yt-dlp执行失败: {error_msg}")
            return None, video_id
        except subprocess.TimeoutExpired:
            st.error("❌ 请求超时，请检查网络连接或稍后重试")
            return None, video_id
        except FileNotFoundError:
            st.error(f"找不到yt-dlp可执行文件")
            return None, video_id


def danmu_page():
    st.markdown("""
    # 🎬 弹幕抓取工具
    
    支持从 **Niconico** 和 **Bilibili** 抓取视频弹幕，并导出为 Excel 文件。
    """)
    
    with st.sidebar:
        st.header("⚙️ 配置1")
        platform = st.radio(
            "选择视频平台",
            options=["Niconico", "Bilibili"],
            help="选择您要抓取弹幕的视频平台",
            key="video_pla_selector"
        )
        
        st.divider()
        
        bilibili_cookies_file = None
        if platform == "Bilibili":
            st.subheader("🔐 Bilibili Cookie配置")
            st.markdown(
                """Bilibili需要Cookie认证以避免速率限制。\n\n
                **获取Cookie的方法：**
                1. 打开浏览器访问 https://www.bilibili.com
                2. 登录您的账号
                3. 按F12打开开发者工具 → Application → Cookies
                4. 复制所有Cookie内容到文本文件
                5. 上传该文件
                """
            )
            
            uploaded_file = st.file_uploader(
                "上传Cookie文件",
                type=["txt"],
                help="上传从浏览器导出的Cookie文件"
            )
            
            if uploaded_file is not None:
                cookies_content = uploaded_file.read().decode('utf-8')
                bilibili_cookies_file = "temp_cookies.txt"
                with open(bilibili_cookies_file, 'w', encoding='utf-8') as f:
                    f.write(cookies_content)
                st.success("✅ Cookie文件已加载")
            else:
                st.warning("⚠️ 未上传Cookie文件，可能导致速率限制错误")
        
        st.divider()
        
        st.markdown("""
        ### 📌 使用说明
        
        **Niconico:**
        - 输入格式: `https://www.nicovideo.jp/watch/sm500873`
        - 或: `https://www.video.nicovideo.jp/watch/sm500873` (会自动转换)
        
        **Bilibili:**
        - 输入格式: `https://www.bilibili.com/video/BV1xx411c7mD`
        - 或: `https://www.bilibili.com/video/av123456789`
        - 建议上传Cookie文件以避免速率限制
        """)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader(f"输入{platform}视频链接")
        video_url = st.text_input(
            "视频链接",
            placeholder=f"请输入{platform}视频链接...",
            label_visibility="collapsed"
        )
    
    with col2:
        st.subheader("操作")
        scrape_button = st.button(
            "🔍 开始抓取",
            use_container_width=True,
            type="primary"
        )
    
    st.divider()
    
    if scrape_button:
        if not video_url.strip():
            st.error("❌ 请输入视频链接")
        else:
            if platform == "Niconico":
                df, video_id = scrape_niconico_danmaku(video_url)
            else:
                df, video_id = scrape_bilibili_danmaku(video_url, cookies_file=bilibili_cookies_file)
            
            if df is not None and len(df) > 0:
                st.success(f"✅ 成功抓取 {len(df)} 条弹幕！")
                
                st.subheader("📊 弹幕数据预览")
                st.dataframe(df, use_container_width=True, height=400)
                
                st.subheader("💾 导出选项")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    excel_buffer = pd.ExcelWriter(
                        f"danmaku_{video_id}.xlsx",
                        engine='openpyxl'
                    )
                    df.to_excel(excel_buffer, index=False, sheet_name='弹幕数据')
                    excel_buffer.close()
                    
                    with open(f"danmaku_{video_id}.xlsx", 'rb') as f:
                        st.download_button(
                            label="📥 下载 Excel",
                            data=f.read(),
                            file_name=f"danmaku_{video_id}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    
                    try:
                        os.remove(f"danmaku_{video_id}.xlsx")
                    except OSError:
                        pass
                    
                    if bilibili_cookies_file and os.path.exists(bilibili_cookies_file):
                        try:
                            os.remove(bilibili_cookies_file)
                        except OSError:
                            pass
                
                with col2:
                    csv_buffer = df.to_csv(index=False)
                    st.download_button(
                        label="📥 下载 CSV",
                        data=csv_buffer,
                        file_name=f"danmaku_{video_id}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                
                with col3:
                    json_buffer = df.to_json(orient='records', force_ascii=False)
                    st.download_button(
                        label="📥 下载 JSON",
                        data=json_buffer,
                        file_name=f"danmaku_{video_id}.json",
                        mime="application/json",
                        use_container_width=True
                    )
                
                st.subheader("📈 统计信息")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("总弹幕数", len(df))
                
                with col2:
                    if '用户ID' in df.columns:
                        unique_users = df['用户ID'].nunique()
                        st.metric("独立用户数", unique_users)
                
                with col3:
                    if '弹幕内容' in df.columns:
                        avg_length = df['弹幕内容'].str.len().mean()
                        st.metric("平均弹幕长度", f"{avg_length:.1f} 字符")
            
            elif df is not None and len(df) == 0:
                st.warning("⚠️ 未找到弹幕数据，请检查视频链接是否正确")
            else:
                st.error("❌ 抓取失败，请检查视频链接或网络连接")
    
    st.divider()
    st.markdown("""
    ---
    **弹幕抓取工具** | 基于 Streamlit 和 yt-dlp
    
    💡 **提示:**
    - 某些视频的弹幕可能需要登录才能访问
    - 如果遇到问题，请检查您的网络连接
    - 支持的视频平台：Niconico、Bilibili
    """)


def find_excel_files(folder_path):
    """查找文件夹中的所有Excel文件"""
    excel_files = []
    folder_path = Path(folder_path)
    
    if not folder_path.exists():
        return False, "文件夹路径不存在"
    
    excel_extensions = ['.xlsx', '.xls', '.xlsm', '.xlsb']
    
    for ext in excel_extensions:
        excel_files.extend(folder_path.rglob(f'*{ext}'))
    
    return True, excel_files


def search_in_dataframe(df, col_name, target_values, keyword, case_sensitive=False, match_whole_word=False):
    """在DataFrame中搜索满足条件的行"""
    matches = []
    
    if col_name not in df.columns:
        return matches, f"列 '{col_name}' 不存在"
    
    filtered_df = df[df[col_name].astype(str).isin([str(v) for v in target_values])]
    
    if len(filtered_df) == 0:
        return matches, "未找到包含指定目标值的行"
    
    if not keyword or not keyword.strip():
        for idx, row in filtered_df.iterrows():
            matches.append({
                'row_index': idx,
                'row_data': row.to_dict(),
                'matched_column': col_name,
                'matched_value': row[col_name],
                'keyword_found': False,
                'keyword_matches': [],
                'match_count': 0
            })
        return matches, f"找到 {len(matches)} 行包含目标值，但未搜索关键词"
    
    keyword = keyword.strip()
    flags = 0 if case_sensitive else re.IGNORECASE
    
    if match_whole_word:
        pattern = r'\b' + re.escape(keyword) + r'\b'
    else:
        pattern = re.escape(keyword)
    
    keyword_matches = 0
    
    for idx, row in filtered_df.iterrows():
        row_matches = []
        
        for col_idx, (col, cell_value) in enumerate(row.items()):
            if pd.isna(cell_value):
                continue
                
            cell_str = str(cell_value)
            cell_matches = list(re.finditer(pattern, cell_str, flags))
            
            for match in cell_matches:
                row_matches.append({
                    'column': col,
                    'original_value': cell_str,
                    'match_text': match.group(),
                    'start_pos': match.start(),
                    'end_pos': match.end(),
                    'replaced_value': cell_str[:match.start()] + f"**[{match.group()}]**" + cell_str[match.end():]
                })
        
        if row_matches:
            keyword_matches += 1
            matches.append({
                'row_index': idx,
                'row_data': row.to_dict(),
                'matched_column': col_name,
                'matched_value': row[col_name],
                'keyword_found': True,
                'keyword_matches': row_matches,
                'match_count': len(row_matches)
            })
    
    return matches, f"在 {len(filtered_df)} 行目标行中找到 {keyword_matches} 行包含关键词"


def highlight_keyword(text, keyword, case_sensitive=False):
    """高亮显示关键词"""
    if not text or not keyword:
        return text
    
    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.escape(keyword)
    
    if case_sensitive:
        highlighted = re.sub(f'({pattern})', r'**\1**', text)
    else:
        highlighted = re.sub(f'({pattern})', r'**\1**', text, flags=re.IGNORECASE)
    
    return highlighted


def replace_in_excel(file_path, replacements, backup=True):
    """在Excel文件中执行替换操作"""
    try:
        file_path = Path(file_path)
        
        if backup:
            backup_path = file_path.parent / f"{file_path.stem}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}{file_path.suffix}"
            shutil.copy2(file_path, backup_path)
            st.info(f"📁 已创建备份文件: {backup_path.name}")
        
        df = pd.read_excel(file_path, engine='openpyxl')
        
        replaced_count = 0
        for replacement in replacements:
            row_idx = replacement['row_index']
            col_name = replacement['column']
            old_text = replacement['original_value']
            new_text = replacement['new_value']
            
            if row_idx < len(df) and col_name in df.columns:
                current_value = df.at[row_idx, col_name]
                if pd.isna(current_value):
                    current_value = ""
                
                current_str = str(current_value)
                
                if replacement.get('replace_all', False):
                    flags = 0 if replacement.get('case_sensitive', False) else re.IGNORECASE
                    pattern = re.escape(replacement['search_keyword'])
                    if replacement.get('match_whole_word', False):
                        pattern = r'\b' + pattern + r'\b'
                    
                    new_str = re.sub(pattern, new_text, current_str, flags=flags)
                else:
                    new_str = old_text
                
                df.at[row_idx, col_name] = new_str
                replaced_count += 1
        
        df.to_excel(file_path, index=False, engine='openpyxl')
        return True, replaced_count
    except Exception as e:
        return False, str(e)


def excel_sreplace_page():
    st.sidebar.header("🔧 搜索参数设置")
    
    folder_path = st.sidebar.text_input(
        "📁 请输入文件夹路径:",
        placeholder="例如: C:/Users/用户名/Documents/Excel文件",
        help="请输入包含Excel文件的文件夹完整路径"
    )
    
    col_name = st.sidebar.text_input(
        "📊 要搜索的列名:",
        value="角色名",
        placeholder="例如: 角色名",
        help="请输入要搜索的Excel列名称"
    )
    
    target_values_input = st.sidebar.text_input(
        "🎯 列目标值（用逗号分隔）:",
        value="班长,班长大人",
        placeholder="例如: 班长,班长大人",
        help="请输入要在指定列中查找的值，多个值用逗号分隔"
    )
    
    search_keyword = st.sidebar.text_input(
        "🔤 要查找的关键词YYY:",
        value="私",
        placeholder="请输入要在行中查找的关键词",
        help="在满足条件的行中查找此关键词"
    )
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("⚙️ 高级选项")
    
    case_sensitive = st.sidebar.checkbox(
        "区分大小写",
        value=False,
        help="勾选后搜索时区分英文大小写"
    )
    
    match_whole_word = st.sidebar.checkbox(
        "全词匹配",
        value=False,
        help="勾选后只匹配完整的词语"
    )
    
    target_values = [v.strip() for v in target_values_input.split(',') if v.strip()]
    
    if not folder_path:
        st.warning("⚠️ 请输入文件夹路径")
        return
    
    if not col_name:
        st.warning("⚠️ 请输入要搜索的列名")
        return
    
    if not target_values:
        st.warning("⚠️ 请输入至少一个目标值")
        return
    
    success, result = find_excel_files(folder_path)
    
    if not success:
        st.error(f"❌ {result}")
        return
    
    excel_files = result
    
    if not excel_files:
        st.warning("⚠️ 在指定文件夹中未找到Excel文件")
        return
    
    st.success(f"✅ 找到 {len(excel_files)} 个Excel文件")
    
    with st.expander("📁 找到的Excel文件"):
        for i, file_path in enumerate(excel_files[:10]):
            st.write(f"{i+1}. {file_path.name}")
        
        if len(excel_files) > 10:
            st.info(f"... 还有 {len(excel_files) - 10} 个文件")
    
    if st.button("🚀 开始搜索", type="primary", use_container_width=True):
        if not search_keyword.strip():
            st.warning("⚠️ 关键词YYY为空，将只显示包含目标值的行")
        
        all_matches = []
        files_with_matches = 0
        total_matches = 0
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i, file_path in enumerate(excel_files):
            progress = (i + 1) / len(excel_files)
            progress_bar.progress(progress)
            status_text.text(f"🔍 正在处理文件 {i+1}/{len(excel_files)}: {file_path.name}")
            
            try:
                df = pd.read_excel(file_path, engine='openpyxl')
                
                matches, message = search_in_dataframe(
                    df, col_name, target_values, search_keyword, 
                    case_sensitive, match_whole_word
                )
                
                if matches:
                    files_with_matches += 1
                    total_matches += len(matches)
                    
                    for match in matches:
                        match['file_path'] = str(file_path)
                        match['file_name'] = file_path.name
                        all_matches.append(match)
                
                import time
                time.sleep(0.1)
                
            except Exception as e:
                st.error(f"处理文件 {file_path.name} 时出错: {e}")
        
        progress_bar.progress(1.0)
        status_text.text(f"✅ 搜索完成！")
        
        st.session_state.search_results = all_matches
        st.session_state.search_keyword = search_keyword
        st.session_state.case_sensitive = case_sensitive
        st.session_state.match_whole_word = match_whole_word
        
        st.header("📊 搜索结果")
        
        if total_matches == 0:
            st.warning("⚠️ 未找到满足条件的行")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("扫描文件数", len(excel_files))
            with col2:
                st.metric("包含匹配的文件", files_with_matches)
            with col3:
                st.metric("总匹配行数", total_matches)
        else:
            st.success(f"✅ 在 {files_with_matches} 个文件中找到 {total_matches} 行匹配结果")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("扫描文件数", len(excel_files))
            with col2:
                st.metric("包含匹配的文件", files_with_matches)
            with col3:
                st.metric("总匹配行数", total_matches)
            
            files_group = {}
            for match in all_matches:
                file_name = match['file_name']
                if file_name not in files_group:
                    files_group[file_name] = []
                files_group[file_name].append(match)
            
            for file_name, file_matches in files_group.items():
                with st.expander(f"📄 {file_name} ({len(file_matches)} 行匹配)", expanded=True):
                    st.write(f"**文件路径:** {file_matches[0]['file_path']}")
                    
                    display_data = []
                    for i, match in enumerate(file_matches):
                        row_data = match['row_data']
                        
                        row_display = {}
                        for col_name_display, cell_value in row_data.items():
                            if pd.isna(cell_value):
                                display_value = ""
                            else:
                                cell_str = str(cell_value)
                                
                                if search_keyword and search_keyword.strip():
                                    highlighted = highlight_keyword(
                                        cell_str, search_keyword, case_sensitive
                                    )
                                    row_display[col_name_display] = highlighted
                                else:
                                    row_display[col_name_display] = cell_str
                        
                        display_data.append({
                            '序号': i + 1,
                            '匹配列值': match['matched_value'],
                            '关键词匹配数': match.get('match_count', 0) if match.get('keyword_found', False) else '无',
                            **row_display
                        })
                    
                    if display_data:
                        all_columns = set()
                        for item in display_data:
                            all_columns.update(item.keys())
                        
                        base_columns = ['序号', '匹配列值', '关键词匹配数']
                        other_columns = [col for col in all_columns if col not in base_columns]
                        display_columns = base_columns + sorted(other_columns)
                        
                        display_df = pd.DataFrame(display_data)
                        
                        for col in display_columns:
                            if col not in display_df.columns:
                                display_df[col] = ""
                        
                        display_df = display_df[display_columns]
                        
                        st.dataframe(display_df, use_container_width=True)
                    
                    if search_keyword and search_keyword.strip():
                        st.subheader("🔍 匹配详情")
                        for i, match in enumerate(file_matches):
                            if match.get('keyword_found', False) and match.get('keyword_matches'):
                                with st.expander(f"匹配详情 - 行 {i+1}"):
                                    st.write(f"**文件:** {file_name}")
                                    st.write(f"**行索引:** {match['row_index'] + 2}")
                                    st.write(f"**在列 '{col_name}' 中找到值:** {match['matched_value']}")
                                    st.write(f"**关键词匹配位置:**")
                                    
                                    for kw_match in match['keyword_matches']:
                                        col_name_kw = kw_match['column']
                                        cell_value = kw_match['original_value']
                                        match_text = kw_match['match_text']
                                        start_pos = kw_match['start_pos']
                                        end_pos = kw_match['end_pos']
                                        
                                        context_start = max(0, start_pos - 20)
                                        context_end = min(len(cell_value), end_pos + 20)
                                        context = cell_value[context_start:context_end]
                                        
                                        if context_start > 0:
                                            context = "..." + context
                                        if context_end < len(cell_value):
                                            context = context + "..."
                                        
                                        highlighted_context = highlight_keyword(
                                            context, search_keyword, case_sensitive
                                        )
                                        
                                        st.write(f"**列 '{col_name_kw}':** {highlighted_context}")
            
            if all_matches:
                st.header("💾 下载搜索结果")
                
                download_data = []
                for match in all_matches:
                    row_data = {
                        '文件路径': match['file_path'],
                        '文件名称': match['file_name'],
                        '行索引': match['row_index'] + 2,
                        '匹配列': match['matched_column'],
                        '匹配列值': match['matched_value'],
                        '是否找到关键词': '是' if match.get('keyword_found', False) else '否',
                        '关键词匹配数': match.get('match_count', 0)
                    }
                    
                    for col_name_dl, cell_value in match['row_data'].items():
                        row_data[col_name_dl] = cell_value if not pd.isna(cell_value) else ""
                    
                    download_data.append(row_data)
                
                download_df = pd.DataFrame(download_data)
                
                csv_data = download_df.to_csv(index=False).encode('utf-8-sig')
                
                st.download_button(
                    label="📥 下载搜索结果(CSV)",
                    data=csv_data,
                    file_name=f"excel_search_results.csv",
                    mime="text/csv",
                    use_container_width=True
                )
    
    if st.session_state.get('search_results'):
        st.markdown("---")
        st.header("🔄 替换功能")
        
        search_results = st.session_state.search_results
        search_keyword = st.session_state.search_keyword
        case_sensitive = st.session_state.case_sensitive
        match_whole_word = st.session_state.match_whole_word
        
        col1, col2 = st.columns(2)
        
        with col1:
            replace_keyword = st.text_input(
                "🔄 替换为:",
                value="僕",
                placeholder="请输入替换后的词语",
                help="将查找到的关键词替换为此词语"
            )
        
        with col2:
            replace_mode = st.radio(
                "替换模式:",
                ["替换所有匹配", "仅替换选中项"],
                help="选择替换全部匹配项还是仅替换选中的匹配项",
                key = "tihuanA"
            )
            
            create_backup = st.checkbox(
                "创建备份文件",
                value=True,
                help="替换前自动创建备份文件"
            )
        
        st.subheader("👁️ 替换预览")
        
        all_replacements = []
        files_to_replace = set()
        
        for match in search_results:
            if match.get('keyword_found', False) and match.get('keyword_matches'):
                file_path = match['file_path']
                files_to_replace.add(file_path)
                
                for kw_match in match['keyword_matches']:
                    replacement_info = {
                        'file_path': file_path,
                        'file_name': match['file_name'],
                        'row_index': match['row_index'],
                        'column': kw_match['column'],
                        'original_value': kw_match['original_value'],
                        'search_keyword': search_keyword,
                        'replace_keyword': replace_keyword,
                        'start_pos': kw_match['start_pos'],
                        'end_pos': kw_match['end_pos'],
                        'case_sensitive': case_sensitive,
                        'match_whole_word': match_whole_word,
                        'replace_all': (replace_mode == "替换所有匹配")
                    }
                    
                    if replacement_info['replace_all']:
                        flags = 0 if case_sensitive else re.IGNORECASE
                        pattern = re.escape(search_keyword)
                        if match_whole_word:
                            pattern = r'\b' + pattern + r'\b'
                        
                        new_value = re.sub(pattern, replace_keyword, kw_match['original_value'], flags=flags)
                    else:
                        new_value = kw_match['original_value'][:kw_match['start_pos']] + replace_keyword + kw_match['original_value'][kw_match['end_pos']:]
                    
                    replacement_info['new_value'] = new_value
                    all_replacements.append(replacement_info)
        
        if all_replacements:
            st.info(f"📊 共找到 {len(all_replacements)} 处可替换内容，涉及 {len(files_to_replace)} 个文件")
            
            for file_path in files_to_replace:
                file_replacements = [r for r in all_replacements if r['file_path'] == file_path]
                
                with st.expander(f"📄 {Path(file_path).name} - {len(file_replacements)} 处替换"):
                    preview_data = []
                    
                    for i, replacement in enumerate(file_replacements[:10]):
                        original_text = replacement['original_value']
                        new_text = replacement['new_value']
                        
                        preview_data.append({
                            '序号': i + 1,
                            '行索引': replacement['row_index'] + 2,
                            '列名': replacement['column'],
                            '原文本': original_text,
                            '替换后': new_text
                        })
                    
                    if preview_data:
                        st.dataframe(pd.DataFrame(preview_data), use_container_width=True)
                    
                    if len(file_replacements) > 10:
                        st.info(f"... 还有 {len(file_replacements) - 10} 处替换")
            
            if st.button("✅ 执行替换", type="primary", use_container_width=True):
                with st.spinner("正在执行替换..."):
                    success_count = 0
                    failed_count = 0
                    
                    for file_path in files_to_replace:
                        file_replacements = [r for r in all_replacements if r['file_path'] == file_path]
                        success, count = replace_in_excel(file_path, file_replacements, create_backup)
                        
                        if success:
                            success_count += 1
                            st.success(f"✅ 成功替换 {Path(file_path).name} 中的 {count} 处内容")
                        else:
                            failed_count += 1
                            st.error(f"❌ 替换 {Path(file_path).name} 失败: {count}")
                    
                    st.markdown("---")
                    st.subheader("替换结果")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("成功替换文件数", success_count)
                    with col2:
                        st.metric("失败文件数", failed_count)
                    
                    if success_count > 0:
                        st.success("✅ 替换完成！")
                    else:
                        st.error("❌ 替换失败！")


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
