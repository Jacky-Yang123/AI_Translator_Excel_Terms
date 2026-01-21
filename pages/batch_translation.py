# pages/batch_translation.py - 批量翻译工具页面

import os
from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from translator import MultiAPIExcelTranslator
from api_config import (
    get_api_providers,
    get_preset_languages,
    get_preset_options,
    get_default_custom_requirements
)


def batch_translation_page():
    st.title("批量翻译工具 - 多语言优化版")
    text_col = None
    role_col = None
    source_col = None
    role_name_col = None
    personality_col = None

    st.markdown("### 支持同时翻译多种语言，独立的术语库和上下文管理")

    # 初始化会话状态
    if 'translator' not in st.session_state:
        st.session_state.translator = None
    if 'current_file' not in st.session_state:
        st.session_state.current_file = None

    # 侧边栏 - API配置
    st.sidebar.header("🔑 API配置")

    providers = get_api_providers()
    selected_provider = st.sidebar.selectbox(
        "选择API提供商",
        options=list(providers.keys()),
        index=0
    )

    if selected_provider == "自定义API":
        api_url = st.sidebar.text_input(
            "API URL",
            value=providers[selected_provider]["url"],
            placeholder="输入自定义API URL"
        )
        model = st.sidebar.text_input(
            "模型名称",
            value="custom-model",
            placeholder="输入模型名称"
        )
    else:
        api_url = providers[selected_provider]["url"]
        model = st.sidebar.selectbox(
            "选择模型",
            options=providers[selected_provider]["models"],
            index=0
        )

    api_key = st.sidebar.text_input(
        "API Key",
        type="password",
        placeholder="输入您的API Key"
    )

    # 翻译设置
    st.sidebar.header("⚙️ 翻译设置")

    context_size = st.sidebar.slider(
        "上下文大小",
        min_value=0,
        max_value=20,
        value=10,
        help="保留多少条历史翻译作为上下文参考"
    )

    max_retries = st.sidebar.slider(
        "最大重试次数",
        min_value=1,
        max_value=20,
        value=10
    )

    # 目标语言选择
    st.sidebar.header("🌍 目标语言")

    preset_languages = get_preset_languages()[:-1]  # 排除"自定义"
    selected_languages = st.sidebar.multiselect(
        "选择目标语言",
        options=preset_languages,
        default=["英文"]
    )

    custom_language = st.sidebar.text_input(
        "自定义语言（可选）",
        placeholder="输入其他语言"
    )

    if custom_language and custom_language not in selected_languages:
        selected_languages.append(custom_language)

    # 主界面
    col1, col2 = st.columns(2)

    with col1:
        st.header("📁 待翻译文件")

        uploaded_file = st.file_uploader(
            "上传Excel文件",
            type=['xlsx', 'xls'],
            key="batch_file_uploader"
        )

        if uploaded_file is not None:
            try:
                df = pd.read_excel(uploaded_file)
                df.columns = df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                st.session_state.batch_df = df
                st.success(f"✅ 成功读取文件: {len(df)} 行")

                with st.expander("📊 文件预览"):
                    st.dataframe(df.head(10))

                cols = df.columns.tolist()

                text_col = st.selectbox(
                    "选择待翻译文本列",
                    options=cols,
                    index=0,
                    key="batch_text_col"
                )

                role_col = st.selectbox(
                    "选择角色列（可选）",
                    options=["无"] + cols,
                    index=0,
                    key="batch_role_col"
                )
                role_col = role_col if role_col != "无" else None

            except Exception as e:
                st.error(f"❌ 读取文件失败: {e}")

    with col2:
        st.header("📚 术语库和性格库")

        # 术语库
        st.subheader("📖 术语库")
        term_file = st.file_uploader(
            "上传术语库文件",
            type=['xlsx', 'xls'],
            key="batch_term_file"
        )

        if term_file is not None:
            try:
                term_df = pd.read_excel(term_file)
                term_df.columns = term_df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                st.session_state.batch_term_df = term_df
                st.success(f"✅ 读取术语库: {len(term_df)} 条")

                term_cols = term_df.columns.tolist()
                source_col = st.selectbox(
                    "原文列",
                    options=term_cols,
                    index=0,
                    key="batch_term_source"
                )

                # 多语言术语库支持
                st.write("选择各语言的译文列：")
                term_target_cols = {}
                for lang in selected_languages:
                    col_options = ["无"] + term_cols
                    selected = st.selectbox(
                        f"{lang}译文列",
                        options=col_options,
                        index=0,
                        key=f"batch_term_target_{lang}"
                    )
                    if selected != "无":
                        term_target_cols[lang] = selected

            except Exception as e:
                st.error(f"❌ 读取术语库失败: {e}")

        # 角色性格库
        st.subheader("👤 角色性格库")
        role_file = st.file_uploader(
            "上传角色性格库文件",
            type=['xlsx', 'xls'],
            key="batch_role_file"
        )

        if role_file is not None:
            try:
                role_df = pd.read_excel(role_file)
                role_df.columns = role_df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                st.session_state.batch_role_df = role_df
                st.success(f"✅ 读取角色库: {len(role_df)} 条")

                role_cols = role_df.columns.tolist()
                role_name_col = st.selectbox(
                    "角色名称列",
                    options=role_cols,
                    index=0,
                    key="batch_role_name_col"
                )

                personality_col = st.selectbox(
                    "性格描述列",
                    options=role_cols,
                    index=min(1, len(role_cols)-1),
                    key="batch_personality_col"
                )

            except Exception as e:
                st.error(f"❌ 读取角色库失败: {e}")

    # 翻译要求
    st.header("📝 翻译要求")

    presets = get_preset_options()
    selected_presets = st.multiselect(
        "选择预设要求",
        options=list(presets.keys()),
        default=[]
    )

    custom_requirements = st.text_area(
        "自定义翻译要求",
        value=get_default_custom_requirements(),
        height=100
    )

    all_requirements = []
    if selected_presets:
        all_requirements.extend(selected_presets)
    if custom_requirements:
        all_requirements.append(custom_requirements)
    final_requirements = "；".join(all_requirements) if all_requirements else ""

    # 保存路径
    st.header("💾 保存设置")

    save_directory = st.text_input(
        "保存目录",
        value=os.path.join(os.path.expanduser("~"), "Downloads", "翻译结果"),
        key="batch_save_dir"
    )

    # 开始翻译按钮
    if st.button("🚀 开始翻译", type="primary", use_container_width=True):
        # 验证
        if not api_key:
            st.error("❌ 请输入API Key")
            return

        if 'batch_df' not in st.session_state:
            st.error("❌ 请上传待翻译文件")
            return

        if not text_col:
            st.error("❌ 请选择待翻译文本列")
            return

        if not selected_languages:
            st.error("❌ 请选择至少一种目标语言")
            return

        # 创建翻译器
        translator = MultiAPIExcelTranslator(
            api_key=api_key,
            api_provider=selected_provider,
            api_url=api_url,
            model=model,
            context_size=context_size,
            max_retries=max_retries
        )

        # 加载术语库
        if 'batch_term_df' in st.session_state and source_col:
            term_df = st.session_state.batch_term_df
            if term_target_cols:
                translator.load_term_base_multilang(term_df, source_col, term_target_cols)
            else:
                # 使用第一个非原文列作为目标列
                target_col = [c for c in term_df.columns if c != source_col][0] if len(term_df.columns) > 1 else source_col
                translator.load_term_base(term_df, source_col, target_col)

        # 加载角色性格库
        if 'batch_role_df' in st.session_state and role_name_col and personality_col:
            role_df = st.session_state.batch_role_df
            translator.load_role_personality(role_df, role_name_col, personality_col)

        # 设置目标语言
        language_column_names = {lang: f"{lang}翻译结果" for lang in selected_languages}
        translator.set_target_languages(selected_languages, language_column_names)

        # 开始翻译
        df = st.session_state.batch_df.copy()
        total_rows = len(df)

        # 添加结果列
        for lang in selected_languages:
            df[f"{lang}翻译结果"] = ""

        progress_bar = st.progress(0)
        status_text = st.empty()
        current_translation = st.empty()

        try:
            for idx, row in df.iterrows():
                text = row[text_col]
                role = row[role_col] if role_col and role_col in df.columns else None

                if pd.isna(text) or str(text).strip() == "":
                    continue

                for lang in selected_languages:
                    status_text.text(f"翻译中: 第 {idx + 1}/{total_rows} 行 - {lang}")
                    current_translation.text(f"原文: {str(text)[:100]}...")

                    translated = translator.translate_text(
                        text, lang, final_requirements, role
                    )

                    df.at[idx, f"{lang}翻译结果"] = translated

                progress_bar.progress((idx + 1) / total_rows)

            progress_bar.empty()
            status_text.empty()
            current_translation.empty()

            st.success(f"✅ 翻译完成！共 {total_rows} 行")

            # 保存结果
            os.makedirs(save_directory, exist_ok=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_filename = f"translated_{timestamp}.xlsx"
            output_path = os.path.join(save_directory, output_filename)

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='翻译结果')

            st.success(f"📁 文件已保存至: {output_path}")

            # 下载按钮
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='翻译结果')

            st.download_button(
                label="📥 下载翻译结果",
                data=output.getvalue(),
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            # 结果预览
            with st.expander("📊 翻译结果预览"):
                st.dataframe(df.head(20))

        except Exception as e:
            st.error(f"❌ 翻译过程中出错: {e}")
            import traceback
            st.error(traceback.format_exc())
