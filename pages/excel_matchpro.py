import os
import time
import threading
from pathlib import Path
from io import BytesIO
import zipfile
import concurrent.futures

import pandas as pd
import streamlit as st


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
