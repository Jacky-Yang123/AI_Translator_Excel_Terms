# pages/excel_abc.py - Excel ABC 操作页面

import os
import re
from pathlib import Path
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
import openpyxl


def excel_ABC_page():
    """Excel 批量操作页面"""
    st.title("🔧 ExcelABC操作工具")
    st.markdown("### 批量处理Excel文件：删除、替换、修改等操作")

    # 文件上传
    st.header("📁 文件上传")
    uploaded_file = st.file_uploader(
        "上传Excel文件",
        type=['xlsx', 'xls'],
        key="excel_abc_uploader"
    )

    if uploaded_file is None:
        st.info("请上传一个Excel文件开始操作")
        return

    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"✅ 成功读取文件: {len(df)} 行, {len(df.columns)} 列")

        with st.expander("📊 文件预览"):
            st.dataframe(df.head(20))
    except Exception as e:
        st.error(f"❌ 读取文件失败: {e}")
        return

    st.header("⚙️ 操作选择")

    operation = st.selectbox(
        "选择操作类型:",
        options=[
            "删除包含特定内容的行",
            "删除空行",
            "替换单元格内容",
            "删除特定列",
            "添加新列",
            "条件筛选",
            "去重"
        ]
    )

    result_df = df.copy()

    if operation == "删除包含特定内容的行":
        col = st.selectbox("选择列:", options=df.columns.tolist())
        content = st.text_input("包含的内容:")

        if st.button("执行删除"):
            if content:
                before_count = len(result_df)
                result_df = result_df[~result_df[col].astype(str).str.contains(content, na=False)]
                after_count = len(result_df)
                st.success(f"✅ 删除了 {before_count - after_count} 行")

    elif operation == "删除空行":
        if st.button("执行删除空行"):
            before_count = len(result_df)
            result_df = result_df.dropna(how='all')
            after_count = len(result_df)
            st.success(f"✅ 删除了 {before_count - after_count} 个空行")

    elif operation == "替换单元格内容":
        col = st.selectbox("选择列:", options=df.columns.tolist())
        search = st.text_input("要查找的内容:")
        replace = st.text_input("替换为:")

        if st.button("执行替换"):
            if search:
                result_df[col] = result_df[col].astype(str).str.replace(search, replace, regex=False)
                st.success("✅ 替换完成")

    elif operation == "删除特定列":
        cols_to_delete = st.multiselect("选择要删除的列:", options=df.columns.tolist())

        if st.button("执行删除列"):
            if cols_to_delete:
                result_df = result_df.drop(columns=cols_to_delete)
                st.success(f"✅ 删除了 {len(cols_to_delete)} 列")

    elif operation == "添加新列":
        new_col_name = st.text_input("新列名称:")
        default_value = st.text_input("默认值:", value="")

        if st.button("添加新列"):
            if new_col_name:
                result_df[new_col_name] = default_value
                st.success(f"✅ 添加了新列: {new_col_name}")

    elif operation == "条件筛选":
        col = st.selectbox("选择筛选列:", options=df.columns.tolist())
        condition = st.selectbox("条件:", options=["等于", "不等于", "包含", "不包含", "大于", "小于"])
        value = st.text_input("值:")

        if st.button("执行筛选"):
            if value:
                if condition == "等于":
                    result_df = result_df[result_df[col].astype(str) == value]
                elif condition == "不等于":
                    result_df = result_df[result_df[col].astype(str) != value]
                elif condition == "包含":
                    result_df = result_df[result_df[col].astype(str).str.contains(value, na=False)]
                elif condition == "不包含":
                    result_df = result_df[~result_df[col].astype(str).str.contains(value, na=False)]
                elif condition == "大于":
                    result_df = result_df[pd.to_numeric(result_df[col], errors='coerce') > float(value)]
                elif condition == "小于":
                    result_df = result_df[pd.to_numeric(result_df[col], errors='coerce') < float(value)]

                st.success(f"✅ 筛选后剩余 {len(result_df)} 行")

    elif operation == "去重":
        cols_for_dedup = st.multiselect("选择用于去重的列（留空则全列）:", options=df.columns.tolist())

        if st.button("执行去重"):
            before_count = len(result_df)
            if cols_for_dedup:
                result_df = result_df.drop_duplicates(subset=cols_for_dedup)
            else:
                result_df = result_df.drop_duplicates()
            after_count = len(result_df)
            st.success(f"✅ 删除了 {before_count - after_count} 个重复行")

    # 结果预览和下载
    st.header("📊 结果预览")
    st.dataframe(result_df.head(20))
    st.write(f"结果行数: {len(result_df)}")

    # 下载结果
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False, sheet_name='结果')

    st.download_button(
        label="📥 下载处理结果",
        data=output.getvalue(),
        file_name=f"processed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
