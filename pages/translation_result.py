# pages/translation_result.py - 翻译结果处理页面

import re
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st

from api_config import get_preset_languages


def parse_ai_translation_result(text):
    try:
        text = text.strip()
        lines = text.split('\n')
        translations = {}

        # 查找表格开始位置
        table_start = -1
        header_found = False

        for i, line in enumerate(lines):
            line = line.strip()
            if not line or not '|' in line:
                continue

            # 检查是否是表头行
            if ('原文' in line or '中文' in line) and ('翻译' in line or 'Translation' in line or '英文' in line or '日文' in line):
                header_found = True
                # 检查下一行是否是分隔线
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if '|' in next_line and ('---' in next_line or '===' in next_line or '--' in next_line):
                        table_start = i + 2  # 数据从分隔线后开始
                    else:
                        table_start = i + 1  # 没有分隔线，数据从表头后开始
                else:
                    table_start = i + 1
                break

        # 如果没找到标准表头，尝试查找任何包含 | 的行作为数据起点
        if not header_found:
            for i, line in enumerate(lines):
                if '|' in line and not ('---' in line or '===' in line):
                    # 尝试解析这一行，看是否有两列数据
                    parts = [part.strip() for part in line.split('|') if part.strip()]
                    if len(parts) >= 2:
                        table_start = i
                        break

        if table_start == -1:
            st.warning("未找到表格结构，尝试使用备用解析方法...")
            return parse_fallback_format(text)

        # 解析表格数据
        success_count = 0
        for i in range(table_start, len(lines)):
            line = lines[i].strip()

            # 跳过空行和分隔线
            if not line or not '|' in line:
                continue
            if '---' in line or '===' in line:
                continue

            # 分割行，移除空白部分
            parts = [part.strip() for part in line.split('|')]
            # 移除首尾的空字符串（来自行首尾的|）
            parts = [p for p in parts if p]

            if len(parts) >= 2:
                original_text = parts[0]
                translation_text = parts[1]

                # 清理Markdown格式符号
                original_text = re.sub(r'\*\*|\*|`|#', '', original_text).strip()
                translation_text = re.sub(r'\*\*|\*|`|#', '', translation_text).strip()

                # 只添加非空的有效翻译
                if original_text and translation_text:
                    if original_text not in translations:
                        translations[original_text] = translation_text
                        success_count += 1

        if success_count > 0:
            st.success(f"✅ 成功解析 {success_count} 条翻译")
        else:
            st.warning("表格解析成功但未找到有效数据，尝试备用方法...")
            return parse_fallback_format(text)

        return translations

    except Exception as e:
        st.error(f"解析AI翻译结果时出错: {e}")
        st.warning("尝试使用备用解析方法...")
        return parse_fallback_format(text)


def parse_fallback_format(text):
    try:
        translations = {}
        lines = text.strip().split('\n')
        success_count = 0

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 跳过分隔线和表头
            if '---' in line or '===' in line:
                continue
            if '原文' in line or 'Translation' in line:
                continue

            # 尝试多种分隔符
            if '|' in line:
                # 处理表格格式
                parts = [p.strip() for p in line.split('|')]
                # 移除空字符串
                parts = [p for p in parts if p]

                if len(parts) >= 2:
                    original = parts[0]
                    translation = parts[1]

                    # 清理文本
                    original = re.sub(r'\*\*|\*|`|#', '', original).strip()
                    translation = re.sub(r'\*\*|\*|`|#', '', translation).strip()

                    if original and translation:
                        if original not in translations:
                            translations[original] = translation
                            success_count += 1
            elif '\t' in line:
                # 处理制表符分隔
                parts = line.split('\t')
                if len(parts) >= 2:
                    original = parts[0].strip()
                    translation = parts[1].strip()

                    original = re.sub(r'\*\*|\*|`|#', '', original).strip()
                    translation = re.sub(r'\*\*|\*|`|#', '', translation).strip()

                    if original and translation:
                        if original not in translations:
                            translations[original] = translation
                            success_count += 1

        if success_count > 0:
            st.info(f"📝 备用方法成功解析 {success_count} 条翻译")
        else:
            st.error("❌ 备用方法也未能解析出有效数据")

        return translations
    except Exception as e:
        st.error(f"备用解析方法也失败: {e}")
        return {}


def merge_translations_with_excel(original_df, text_col, translations, target_language):
    try:
        result_df = original_df.copy()
        result_df[f'{target_language}翻译结果'] = ''

        matched_count = 0
        unmatched_texts = []

        for index, row in result_df.iterrows():
            original_text = str(row[text_col]) if not pd.isna(row[text_col]) else ''
            if original_text and original_text in translations:
                result_df.at[index, f'{target_language}翻译结果'] = translations[original_text]
                matched_count += 1
            elif original_text:
                unmatched_texts.append(original_text)

        return result_df, matched_count, unmatched_texts
    except Exception as e:
        st.error(f"合并翻译结果时出错: {e}")
        return original_df, 0, []


def translation_result_processor_page():
    st.title("📊 翻译结果处理工具")
    st.markdown("### 将AI翻译结果合并到原始Excel文件中")

    st.header("🎯 基本设置")

    col1, col2 = st.columns([1, 2])

    with col1:
        language_option = st.selectbox(
            "🌍 选择翻译结果语言",
            options=get_preset_languages()[:-1],  # 排除 "自定义"
            index=0,
            key="result_language_option"
        )

    with col2:
        target_language = language_option
        st.info(f"🎯 当前处理语言: {target_language}")

    col1, col2 = st.columns(2)

    with col1:
        st.header("📁 原始Excel文件")

        uploaded_file = st.file_uploader(
            "📄 上传原始Excel文件",
            type=['xlsx', 'xls'],
            key="result_original_file_uploader"
        )

        df_original = None
        text_col = None

        if uploaded_file is not None:
            try:
                df_original = pd.read_excel(uploaded_file)
                df_original.columns = df_original.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                st.session_state.result_df_original = df_original
                st.success(f"✅ 成功读取文件，共 {len(df_original)} 行数据")

                with st.expander("📊 文件预览"):
                    st.dataframe(df_original.head(10))

                cols = df_original.columns.tolist()
                text_col = st.selectbox(
                    "📝 选择原文文本列（用于匹配）",
                    options=cols,
                    index=0,
                    key="result_text_col_select"
                )

                st.session_state.result_text_col = text_col

            except Exception as e:
                st.error(f"❌ 文件读取失败: {e}")
        else:
            if 'result_df_original' in st.session_state:
                df_original = st.session_state.result_df_original
                text_col = st.session_state.get('result_text_col')

                if df_original is not None:
                    st.info(f"✅ 已加载文件：{len(df_original)} 行数据")
                    if text_col:
                        st.write(f"📝 匹配列: {text_col}")

    with col2:
        st.header("📋 AI翻译结果")

        ai_result = st.text_area(
            "粘贴AI翻译结果（表格格式）",
            height=300,
            placeholder="请粘贴AI翻译结果...\n\n支持的格式:\n1. Markdown表格格式:\n| 原文 | 翻译 |\n|---|---|\n| 你好 | Hello |\n\n2. 制表符分隔格式:\n你好\tHello",
            key="result_ai_text"
        )

    if st.button("🔄 解析并合并", key="merge_btn", use_container_width=True):
        if df_original is None or text_col is None:
            st.error("❌ 请先上传原始Excel文件并选择匹配列。")
            return

        if not ai_result or ai_result.strip() == "":
            st.error("❌ 请粘贴AI翻译结果。")
            return

        with st.spinner("正在解析AI翻译结果..."):
            translations = parse_ai_translation_result(ai_result)

        if not translations:
            st.error("❌ 未能解析出任何翻译结果。")
            return

        st.info(f"📊 解析出 {len(translations)} 条翻译")

        with st.spinner("正在合并翻译结果..."):
            result_df, matched_count, unmatched_texts = merge_translations_with_excel(
                df_original, text_col, translations, target_language
            )

        st.session_state.result_merged_df = result_df

        st.success(f"✅ 合并完成！成功匹配 {matched_count} 条，共 {len(df_original)} 条原文")

        if unmatched_texts:
            with st.expander(f"⚠️ 未匹配的原文 ({len(unmatched_texts)} 条)"):
                for i, text in enumerate(unmatched_texts[:20]):
                    st.write(f"{i+1}. {text[:100]}...")
                if len(unmatched_texts) > 20:
                    st.write(f"... 还有 {len(unmatched_texts) - 20} 条")

    if 'result_merged_df' in st.session_state:
        st.header("📊 合并结果预览")

        st.dataframe(st.session_state.result_merged_df.head(20))

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            st.session_state.result_merged_df.to_excel(writer, index=False, sheet_name='合并结果')

        st.download_button(
            label=f"📥 下载合并结果 ({target_language})",
            data=output.getvalue(),
            file_name=f"merged_translation_{target_language}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
