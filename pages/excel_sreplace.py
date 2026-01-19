import re
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st


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
