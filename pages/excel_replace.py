import pandas as pd
import streamlit as st
import re
import shutil
import openpyxl
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed


class ExcelSearchReplace:
    def __init__(self):
        self.excel_files = []
        self.search_results = {}
        self.case_sensitive = False
        self.match_whole_word = False
        
    def find_excel_files(self, folder_path):
        """查找文件夹中的所有Excel文件"""
        self.excel_files = []
        folder_path = Path(folder_path)
        
        if not folder_path.exists():
            return False, "文件夹路径不存在"
        
        excel_extensions = ['.xlsx', '.xls', '.xlsm', '.xlsb']
        
        for ext in excel_extensions:
            self.excel_files.extend(folder_path.rglob(f'*{ext}'))
        
        return True, f"找到 {len(self.excel_files)} 个Excel文件"
    
    def search_in_excel(self, search_term, case_sensitive=False, match_whole_word=False):
        """在Excel文件中搜索词语"""
        self.search_results = {}
        self.case_sensitive = case_sensitive
        self.match_whole_word = match_whole_word
        total_matches = 0
        
        for file_path in self.excel_files:
            try:
                excel_data = pd.read_excel(file_path, sheet_name=None, dtype=str)
                file_matches = []
                
                for sheet_name, df in excel_data.items():
                    sheet_matches = self._search_in_dataframe(df, search_term, sheet_name, str(file_path))
                    file_matches.extend(sheet_matches)
                
                if file_matches:
                    self.search_results[str(file_path)] = {
                        'matches': file_matches,
                        'match_count': len(file_matches)
                    }
                    total_matches += len(file_matches)
                    
            except Exception as e:
                st.error(f"读取文件 {file_path.name} 时出错: {e}")
        
        return total_matches
    
    def _search_in_dataframe(self, df, search_term, sheet_name, file_path):
        """在DataFrame中搜索词语"""
        matches = []
        
        if self.match_whole_word:
            pattern = r'\b' + re.escape(search_term) + r'\b'
        else:
            pattern = re.escape(search_term)
        
        flags = 0 if self.case_sensitive else re.IGNORECASE
        
        for row_idx, row in df.iterrows():
            for col_idx, cell_value in enumerate(row):
                if pd.isna(cell_value):
                    continue
                
                cell_str = str(cell_value)
                matches_found = list(re.finditer(pattern, cell_str, flags))
                
                for match in matches_found:
                    matches.append({
                        'file_path': file_path,
                        'sheet_name': sheet_name,
                        'row': row_idx + 2,
                        'column': df.columns[col_idx] if col_idx < len(df.columns) else f'Col{col_idx+1}',
                        'original_text': cell_str,
                        'matched_text': match.group(),
                        'start_pos': match.start(),
                        'end_pos': match.end()
                    })
        
        return matches


def multithreaded_search(search_tool, search_term, case_sensitive, match_whole_word, progress_bar, status_text):
    """多线程搜索Excel文件"""
    total_files = len(search_tool.excel_files)
    if total_files == 0:
        return 0
    
    completed = 0
    total_matches = 0
    
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {}
        
        for file_path in search_tool.excel_files:
            future = executor.submit(
                _search_single_file,
                file_path,
                search_term,
                case_sensitive,
                match_whole_word
            )
            futures[future] = file_path
        
        for future in as_completed(futures):
            file_path = futures[future]
            try:
                matches = future.result()
                if matches:
                    search_tool.search_results[str(file_path)] = {
                        'matches': matches,
                        'match_count': len(matches)
                    }
                    total_matches += len(matches)
            except Exception as e:
                st.error(f"搜索文件 {file_path.name} 时出错: {e}")
            
            completed += 1
            progress = completed / total_files
            progress_bar.progress(progress)
            status_text.text(f"正在搜索... {completed}/{total_files} 文件")
    
    return total_matches


def _search_single_file(file_path, search_term, case_sensitive, match_whole_word):
    """搜索单个文件"""
    matches = []
    
    try:
        excel_data = pd.read_excel(file_path, sheet_name=None, dtype=str)
        
        for sheet_name, df in excel_data.items():
            for row_idx, row in df.iterrows():
                for col_idx, cell_value in enumerate(row):
                    if pd.isna(cell_value):
                        continue
                    
                    cell_str = str(cell_value)
                    
                    if match_whole_word:
                        pattern = r'\b' + re.escape(search_term) + r'\b'
                    else:
                        pattern = re.escape(search_term)
                    
                    flags = 0 if case_sensitive else re.IGNORECASE
                    matches_found = list(re.finditer(pattern, cell_str, flags))
                    
                    for match in matches_found:
                        matches.append({
                            'file_path': str(file_path),
                            'sheet_name': sheet_name,
                            'row': row_idx + 2,
                            'column': df.columns[col_idx] if col_idx < len(df.columns) else f'Col{col_idx+1}',
                            'original_text': cell_str,
                            'matched_text': match.group()
                        })
    except Exception as e:
        pass
    
    return matches


def get_row_data_as_list(file_path, sheet_name, row_num):
    """获取指定行的数据"""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
        if row_num - 2 < len(df):
            row_data = df.iloc[row_num - 2].tolist()
            return [str(x) if not pd.isna(x) else "" for x in row_data]
    except:
        pass
    return []


def open_folder(file_path):
    """打开文件所在文件夹"""
    folder_path = Path(file_path).parent
    if folder_path.exists():
        import subprocess
        import platform
        if platform.system() == "Windows":
            os.startfile(str(folder_path))
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", str(folder_path)])
        else:
            subprocess.Popen(["xdg-open", str(folder_path)])


def open_file(file_path):
    """打开文件"""
    try:
        import subprocess
        import platform
        if platform.system() == "Windows":
            os.startfile(str(file_path))
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", str(file_path)])
        else:
            subprocess.Popen(["xdg-open", str(file_path)])
    except Exception as e:
        st.error(f"无法打开文件: {e}")


def excel_replace_page():
    st.title("🔍 Excel文件批量搜索替换工具")
    st.markdown("### 批量搜索和替换文件夹中所有Excel文件的内容")
    
    if 'search_tool' not in st.session_state:
        st.session_state.search_tool = ExcelSearchReplace()
    
    if 'folder_path' not in st.session_state:
        st.session_state.folder_path = ""
    if 'search_term' not in st.session_state:
        st.session_state.search_term = ""
    if 'replace_term' not in st.session_state:
        st.session_state.replace_term = ""
    if 'case_sensitive' not in st.session_state:
        st.session_state.case_sensitive = False
    if 'match_whole_word' not in st.session_state:
        st.session_state.match_whole_word = False
    
    search_tool = st.session_state.search_tool
    
    st.sidebar.header("📁 文件夹设置")
    folder_path = st.sidebar.text_input(
        "请输入文件夹路径:",
        value=st.session_state.folder_path,
        placeholder="例如: C:/Users/用户名/Documents/Excel文件",
        help="请输入包含Excel文件的文件夹完整路径"
    )
    
    if folder_path and folder_path != st.session_state.folder_path:
        st.session_state.folder_path = folder_path
        success, message = search_tool.find_excel_files(folder_path)
        if success:
            st.sidebar.success(message)
        else:
            st.sidebar.error(message)
    
    if search_tool.excel_files:
        st.sidebar.subheader("📊 找到的Excel文件")
        for i, file_path in enumerate(search_tool.excel_files[:10]):
            st.sidebar.write(f"{i+1}. {file_path.name}")
        
        if len(search_tool.excel_files) > 10:
            st.sidebar.info(f"... 还有 {len(search_tool.excel_files) - 10} 个文件")
    
    st.header("🔍 搜索设置")
    
    col1, col2 = st.columns(2)
    
    with col1:
        search_term = st.text_input(
            "搜索词语:",
            value=st.session_state.search_term,
            placeholder="请输入要搜索的词语",
            help="支持正则表达式语法"
        )
        st.session_state.search_term = search_term
    
    with col2:
        st.subheader("⚙️ 搜索选项")
        case_sensitive = st.checkbox(
            "大小写敏感",
            value=st.session_state.case_sensitive,
            help="勾选后区分大小写"
        )
        st.session_state.case_sensitive = case_sensitive
        
        match_whole_word = st.checkbox(
            "全词匹配",
            value=st.session_state.match_whole_word,
            help="勾选后只匹配完整词语"
        )
        st.session_state.match_whole_word = match_whole_word
    
    if st.button("🚀 开始搜索", key="search_btn", use_container_width=True):
        if not folder_path:
            st.error("❌ 请输入文件夹路径")
            return
        
        if not search_term:
            st.error("❌ 请输入搜索词语")
            return
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        with st.spinner("正在搜索Excel文件..."):
            total_matches = multithreaded_search(
                search_tool,
                search_term,
                case_sensitive,
                match_whole_word,
                progress_bar,
                status_text
            )
        
        progress_bar.empty()
        status_text.empty()
        
        if total_matches > 0:
            st.success(f"✅ 搜索完成！共找到 {total_matches} 个匹配项")
        else:
            st.warning("⚠️ 未找到匹配项")
    
    if search_tool.search_results:
        st.header("📊 搜索结果预览")
        
        total_files = len(search_tool.search_results)
        total_matches = sum(data['match_count'] for data in search_tool.search_results.values())
        
        st.info(f"**统计信息:** 在 {total_files} 个文件中找到 {total_matches} 个匹配项")
        
        selected_file = st.selectbox(
            "选择文件查看详情:",
            options=list(search_tool.search_results.keys()),
            format_func=lambda x: f"{Path(x).name} ({search_tool.search_results[x]['match_count']} 处)"
        )
        
        if selected_file:
            file_data = search_tool.search_results[selected_file]
            matches = file_data['matches']
            
            st.subheader(f"📄 文件: {Path(selected_file).name}")
            st.code(selected_file, language=None)
            
            col_btn1, col_btn2 = st.columns([1, 1])
            
            with col_btn1:
                if st.button("📂 打开文件夹", key=f"open_folder_{selected_file}"):
                    open_folder(selected_file)
            
            with col_btn2:
                if st.button("📊 打开Excel", key=f"open_excel_{selected_file}"):
                    open_file(selected_file)
            
            st.write(f"**匹配数量:** {len(matches)} 处")
            
            display_rows = []
            for i, match in enumerate(matches[:50]):
                row_data = get_row_data_as_list(selected_file, match['sheet_name'], match['row'])
                
                row_dict = {
                    "位置": f"{match['sheet_name']} | 行{match['row']} 列{match['column']}"
                }
                
                if row_data:
                    for col_idx, cell_value in enumerate(row_data, start=1):
                        row_dict[f"列{col_idx}"] = cell_value
                
                display_rows.append(row_dict)
            
            if display_rows:
                st.markdown("### 📋 匹配详情")
                df_display = pd.DataFrame(display_rows)
                st.dataframe(df_display, use_container_width=True)
            
            if len(matches) > 50:
                st.info(f"仅显示前 50 个匹配项，共有 {len(matches)} 个匹配项")
