# pages/excel_comparison.py - Excel 表格对比页面

import re
import difflib
from datetime import datetime

import pandas as pd
import streamlit as st


def calculate_similarity(str1, str2):
    """计算两个字符串的相似度（0-1）"""
    if not str1 and not str2:
        return 1.0
    if not str1 or not str2:
        return 0.0
    return difflib.SequenceMatcher(None, str1, str2).ratio()


def preprocess_dataframe_simple(df, ignore_case=True, ignore_whitespace=True):
    """简化的DataFrame预处理函数"""
    df_clean = df.copy()
    df_clean = df_clean.fillna('')

    for col in df_clean.columns:
        df_clean[col] = df_clean[col].astype(str)
        if ignore_case:
            df_clean[col] = df_clean[col].str.lower()
        if ignore_whitespace:
            df_clean[col] = df_clean[col].str.strip()
            df_clean[col] = df_clean[col].str.replace(r'\s+', ' ', regex=True)

    return df_clean


def compare_rows_simple(row_a, row_b, columns, compare_mode="精确匹配", sensitivity=5):
    """简化的行比较函数"""
    changes = []

    for col in columns:
        val_a = str(row_a[col]) if pd.notna(row_a[col]) else ""
        val_b = str(row_b[col]) if pd.notna(row_b[col]) else ""

        if val_a == "" and val_b == "":
            continue

        change_type = "未变化"

        if compare_mode == "精确匹配":
            if val_a != val_b:
                change_type = "修改"
        elif compare_mode == "模糊匹配":
            similarity = calculate_similarity(val_a, val_b)
            threshold = sensitivity / 10.0
            if similarity < threshold:
                change_type = "修改"
        elif compare_mode == "仅比较文本内容":
            text_a = re.sub(r'[^a-zA-Z\u4e00-\u9fa5]', '', val_a)
            text_b = re.sub(r'[^a-zA-Z\u4e00-\u9fa5]', '', val_b)
            if text_a != text_b:
                change_type = "修改"

        if change_type == "修改":
            changes.append({
                'column': col,
                'value_a': val_a,
                'value_b': val_b,
                'change_type': change_type,
                'similarity': calculate_similarity(val_a, val_b) if compare_mode == "模糊匹配" else None
            })

    return changes


def compare_dataframes_simple(df_a, df_b, key_column=None, compare_mode="精确匹配",
                              sensitivity=5, ignore_case=True, ignore_whitespace=True,
                              include_additions=True, include_deletions=True):
    """简化的DataFrame比较函数"""
    results = {
        'added_rows': [],
        'deleted_rows': [],
        'modified_rows': [],
        'modified_cells': [],
        'summary': {
            'total_rows_a': len(df_a),
            'total_rows_b': len(df_b),
            'added_count': 0,
            'deleted_count': 0,
            'modified_count': 0,
            'similarity_score': 0
        }
    }

    df_a_clean = preprocess_dataframe_simple(df_a, ignore_case, ignore_whitespace)
    df_b_clean = preprocess_dataframe_simple(df_b, ignore_case, ignore_whitespace)

    if key_column and key_column in df_a.columns and key_column in df_b.columns:
        a_keys = df_a[key_column].astype(str).tolist()
        b_keys = df_b[key_column].astype(str).tolist()

        if include_additions:
            for i, key in enumerate(b_keys):
                if key not in a_keys:
                    results['added_rows'].append({
                        'key': key,
                        'row_index_b': i,
                        'row_data': df_b.iloc[i].to_dict()
                    })

        if include_deletions:
            for i, key in enumerate(a_keys):
                if key not in b_keys:
                    results['deleted_rows'].append({
                        'key': key,
                        'row_index_a': i,
                        'row_data': df_a.iloc[i].to_dict()
                    })

        common_keys = set(a_keys) & set(b_keys)
        for key in common_keys:
            idx_a = a_keys.index(key)
            idx_b = b_keys.index(key)

            row_a = df_a_clean.iloc[idx_a]
            row_b = df_b_clean.iloc[idx_b]

            changes = compare_rows_simple(row_a, row_b, df_a.columns.tolist(), compare_mode, sensitivity)

            if changes:
                results['modified_rows'].append({
                    'key': key,
                    'row_index_a': idx_a,
                    'row_index_b': idx_b,
                    'row_data_a': df_a.iloc[idx_a].to_dict(),
                    'row_data_b': df_b.iloc[idx_b].to_dict(),
                    'changes': changes,
                    'change_count': len(changes)
                })

                for change in changes:
                    results['modified_cells'].append({
                        'key': key,
                        'row_index_a': idx_a,
                        'row_index_b': idx_b,
                        'column': change['column'],
                        'value_a': change['value_a'],
                        'value_b': change['value_b'],
                        'change_type': change['change_type']
                    })
    else:
        max_rows = min(len(df_a), len(df_b))

        for i in range(max_rows):
            row_a = df_a_clean.iloc[i]
            row_b = df_b_clean.iloc[i]

            changes = compare_rows_simple(row_a, row_b, df_a.columns.tolist(), compare_mode, sensitivity)

            if changes:
                results['modified_rows'].append({
                    'key': f"行{i + 1}",
                    'row_index_a': i,
                    'row_index_b': i,
                    'row_data_a': df_a.iloc[i].to_dict(),
                    'row_data_b': df_b.iloc[i].to_dict(),
                    'changes': changes,
                    'change_count': len(changes)
                })

                for change in changes:
                    results['modified_cells'].append({
                        'key': f"行{i + 1}",
                        'row_index_a': i,
                        'row_index_b': i,
                        'column': change['column'],
                        'value_a': change['value_a'],
                        'value_b': change['value_b'],
                        'change_type': change['change_type']
                    })

        if include_additions and len(df_b) > len(df_a):
            for i in range(len(df_a), len(df_b)):
                results['added_rows'].append({
                    'key': f"新增行{i + 1}",
                    'row_index_b': i,
                    'row_data': df_b.iloc[i].to_dict()
                })

        if include_deletions and len(df_a) > len(df_b):
            for i in range(len(df_b), len(df_a)):
                results['deleted_rows'].append({
                    'key': f"删除行{i + 1}",
                    'row_index_a': i,
                    'row_data': df_a.iloc[i].to_dict()
                })

    results['summary']['added_count'] = len(results['added_rows'])
    results['summary']['deleted_count'] = len(results['deleted_rows'])
    results['summary']['modified_count'] = len(results['modified_rows'])

    total_cells = results['summary']['total_rows_a'] * len(df_a.columns) if len(df_a.columns) > 0 else 0
    if total_cells > 0:
        changed_cells = len(results['modified_cells'])
        similarity = 1 - (changed_cells / total_cells)
        results['summary']['similarity_score'] = round(similarity * 100, 2)

    return results


def display_comparison_results_simple(results, highlight_changes=True, show_unchanged=False):
    """简化的比较结果显示函数"""
    st.markdown("---")
    st.header("📊 比较结果")

    summary = results['summary']
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("文件A行数", summary['total_rows_a'])
    with col2:
        st.metric("文件B行数", summary['total_rows_b'])
    with col3:
        st.metric("修改行数", summary['modified_count'])
    with col4:
        st.metric("相似度", f"{summary['similarity_score']}%")

    if results['added_rows']:
        st.subheader("🆕 新增行")
        added_data = []
        for row in results['added_rows']:
            row_data = {'关键值': row['key'], '行号(B)': row['row_index_b'] + 1}
            row_data.update(row['row_data'])
            added_data.append(row_data)
        st.dataframe(pd.DataFrame(added_data), use_container_width=True)

    if results['deleted_rows']:
        st.subheader("🗑️ 删除行")
        deleted_data = []
        for row in results['deleted_rows']:
            row_data = {'关键值': row['key'], '行号(A)': row['row_index_a'] + 1}
            row_data.update(row['row_data'])
            deleted_data.append(row_data)
        st.dataframe(pd.DataFrame(deleted_data), use_container_width=True)

    if results['modified_rows']:
        st.subheader("✏️ 修改的行")

        for mod_row in results['modified_rows']:
            with st.expander(f"🔍 {mod_row['key']} - {mod_row['change_count']} 处修改", expanded=True):
                comparison_data = []
                all_columns = set(mod_row['row_data_a'].keys()) | set(mod_row['row_data_b'].keys())

                for col in sorted(all_columns):
                    val_a = mod_row['row_data_a'].get(col, '')
                    val_b = mod_row['row_data_b'].get(col, '')

                    is_changed = any(change['column'] == col for change in mod_row['changes'])

                    if is_changed or show_unchanged:
                        comparison_data.append({
                            '列名': col,
                            '文件A值': val_a,
                            '文件B值': val_b,
                            '状态': '❌ 已修改' if is_changed else '✅ 未修改'
                        })

                st.dataframe(pd.DataFrame(comparison_data), use_container_width=True)

    if results['modified_cells']:
        st.subheader("💾 下载比较结果")

        download_data = []
        for cell in results['modified_cells']:
            download_data.append({
                '关键值': cell['key'],
                '行号(A)': cell.get('row_index_a', '') + 1,
                '行号(B)': cell.get('row_index_b', '') + 1,
                '列名': cell['column'],
                '文件A值': cell['value_a'],
                '文件B值': cell['value_b'],
                '修改类型': cell['change_type']
            })

        if download_data:
            download_df = pd.DataFrame(download_data)
            csv_data = download_df.to_csv(index=False).encode('utf-8-sig')

            st.download_button(
                label="📥 下载差异报告(CSV)",
                data=csv_data,
                file_name=f"excel_comparison_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True
            )


def excel_comparison_page():
    st.title("🔍 Excel表格对比工具")
    st.markdown("### 比较两个相似Excel表格，找出差异和改动")

    st.info("💡 此功能适用于比较两个版本相似的Excel文件，找出被修改的内容")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📄 原始表格 (版本A)")
        file_a = st.file_uploader(
            "上传原始Excel文件",
            type=['xlsx', 'xls'],
            key="comparison_file_a"
        )

        if file_a is not None:
            try:
                df_a = pd.read_excel(file_a)
                st.success(f"✅ 成功读取文件A: {len(df_a)} 行, {len(df_a.columns)} 列")

                with st.expander("📊 文件A预览"):
                    st.dataframe(df_a.head(10))
            except Exception as e:
                st.error(f"❌ 读取文件A失败: {e}")

    with col2:
        st.subheader("📄 修改后表格 (版本B)")
        file_b = st.file_uploader(
            "上传修改后的Excel文件",
            type=['xlsx', 'xls'],
            key="comparison_file_b"
        )

        if file_b is not None:
            try:
                df_b = pd.read_excel(file_b)
                st.success(f"✅ 成功读取文件B: {len(df_b)} 行, {len(df_b.columns)} 列")

                with st.expander("📊 文件B预览"):
                    st.dataframe(df_b.head(10))
            except Exception as e:
                st.error(f"❌ 读取文件B失败: {e}")

    st.markdown("---")
    st.subheader("⚙️ 比较设置")

    col1, col2, col3 = st.columns(3)

    with col1:
        key_column = st.text_input(
            "关键列名（用于行匹配）:",
            placeholder="例如: ID、序号等",
            help="用于匹配两个表格中对应行的列名，留空则按行号匹配"
        )

    with col2:
        compare_mode = st.selectbox(
            "比较模式:",
            options=["精确匹配", "模糊匹配", "仅比较文本内容"],
            index=0
        )

    with col3:
        sensitivity = st.slider(
            "差异敏感度:",
            min_value=1,
            max_value=10,
            value=5
        )

    with st.expander("🔧 高级选项"):
        col1, col2 = st.columns(2)

        with col1:
            ignore_case = st.checkbox("忽略大小写", value=True)
            ignore_whitespace = st.checkbox("忽略空白字符", value=True)
            show_unchanged = st.checkbox("显示未更改的行", value=False)

        with col2:
            highlight_changes = st.checkbox("高亮显示更改", value=True)
            include_additions = st.checkbox("检测新增行", value=True)
            include_deletions = st.checkbox("检测删除行", value=True)

    if st.button("🚀 开始比较", type="primary", use_container_width=True):
        if file_a is None or file_b is None:
            st.error("❌ 请先上传两个Excel文件")
            return

        try:
            df_a = pd.read_excel(file_a)
            df_b = pd.read_excel(file_b)

            with st.spinner("🔍 正在比较两个表格..."):
                comparison_results = compare_dataframes_simple(
                    df_a, df_b, key_column, compare_mode, sensitivity,
                    ignore_case, ignore_whitespace, include_additions, include_deletions
                )

            display_comparison_results_simple(
                comparison_results, highlight_changes, show_unchanged
            )

        except Exception as e:
            st.error(f"❌ 比较过程中出错: {e}")
            import traceback
            st.error(traceback.format_exc())
