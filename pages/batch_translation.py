import os
import time
import io
import pandas as pd
import streamlit as st
from datetime import datetime
from translator import MultiAPIExcelTranslator, get_api_providers, render_role_matching_interface


def batch_translation_page():
    st.title("批量翻译工具 - 多语言优化版")
    text_col = None
    role_col = None
    source_col = None
    role_name_col = None
    personality_col = None
    
    st.markdown("### 支持同时翻译多种语言，独立的术语库和上下文管理")

    if 'translator' not in st.session_state:
        st.session_state.translator = None
    if 'current_file' not in st.session_state:
        st.session_state.current_file = None
    if 'term_base_df' not in st.session_state:
        st.session_state.term_base_df = None
    if 'role_personality_df' not in st.session_state:
        st.session_state.role_personality_df = None
    if 'role_matching_confirmed' not in st.session_state:
        st.session_state.role_matching_confirmed = False
    if 'translation_progress' not in st.session_state:
        st.session_state.translation_progress = None
    if 'language_configs' not in st.session_state:
        st.session_state.language_configs = {}
    if 'term_language_mapping' not in st.session_state:
        st.session_state.term_language_mapping = {}

    with st.sidebar:
        st.header("⚙️ API配置")

        api_providers = get_api_providers()
        api_provider = st.selectbox(
            "🌍 API提供商",
            options=list(api_providers.keys()),
            index=0,
            key="batch_api_provider"
        )

        api_key = st.text_input(
            "🔑 API密钥",
            type="password",
            key="batch_api_key"
        )

        if api_provider == "自定义API":
            api_url = st.text_input(
                "🔗 API URL",
                value="https://tb.api.mkeai.com/v1/chat/completions",
                key="batch_api_url"
            )
        else:
            api_url = api_providers[api_provider]["url"]

        model = st.text_input(
            "🤖 模型名称",
            value="deepseek-chat",
            key="batch_model"
        )

        st.markdown("---")
        st.header("🌐 多语言配置")
        
        available_languages = ["英文", "日文", "韩文", "法文", "德文", "西班牙文", "俄文", "阿拉伯文", "葡萄牙文", "意大利文"]
        
        selected_languages = st.multiselect(
            "🎯 选择目标语言（可多选）",
            options=available_languages,
            default=["英文"],
            help="可以同时选择多种语言进行翻译",
            key="selected_languages"
        )
        
        if not selected_languages:
            st.warning("⚠️ 请至少选择一种目标语言")
        
        st.subheader("📝 自定义结果列名")
        language_column_names = {}
        
        for lang in selected_languages:
            default_name = f"{lang}翻译结果"
            col_name = st.text_input(
                f"{lang} 结果列名",
                value=default_name,
                key=f"col_name_{lang}",
                help=f"设置{lang}翻译结果在Excel中的列名"
            )
            language_column_names[lang] = col_name
        
        st.session_state.language_configs = {
            'languages': selected_languages,
            'column_names': language_column_names
        }

        st.markdown("---")
        st.header("💾 自动保存设置")
        
        auto_save_interval = st.number_input(
            "自动保存间隔（每N行）",
            min_value=10,
            max_value=500,
            value=50,
            step=10,
            help="每翻译N行后自动保存一次"
        )
        
        save_directory = st.text_input(
            "保存目录",
            value="./translation_saves",
            help="自动保存文件的目录路径"
        )

        st.markdown("---")
        st.header("🎭 角色匹配设置")

        enable_fuzzy = st.checkbox(
            "启用模糊角色匹配",
            value=True,
            help="自动识别文档中的角色名变体（如空格、错别字等）",
            key="enable_fuzzy_match"
        )

        if enable_fuzzy:
            fuzzy_threshold = st.slider(
                "匹配相似度阈值",
                min_value=0.5,
                max_value=1.0,
                value=0.6,
                step=0.05,
                help="相似度越高越严格，0.6为推荐值",
                key="fuzzy_threshold"
            )
        else:
            fuzzy_threshold = 1.0

        st.markdown("---")
        context_size = st.slider(
            "📚 上下文记录数量",
            min_value=1,
            max_value=20,
            value=10,
            help="每种语言独立维护的上下文数量",
            key="batch_context_size"
        )

        max_retries = st.number_input(
            "🔄 最大重试次数",
            min_value=1,
            max_value=10000,
            value=10,
            key="batch_max_retries"
        )

    col1, col2 = st.columns([1, 1])

    with col1:
        st.header("📁 文件上传")

        saved_files = []
        if os.path.exists(save_directory):
            saved_files = [f for f in os.listdir(save_directory) if f.endswith('_progress.xlsx')]
        
        resume_mode = st.checkbox(
            "🔄 从上次进度继续翻译",
            value=False,
            help="从之前保存的进度文件继续翻译"
        )
        
        if resume_mode and saved_files:
            st.info("📋 找到以下进度文件：")
            selected_progress_file = st.selectbox(
                "选择要继续的进度文件",
                options=saved_files,
                format_func=lambda x: f"{x} ({datetime.fromtimestamp(os.path.getmtime(os.path.join(save_directory, x))).strftime('%Y-%m-%d %H:%M:%S')})"
            )
            
            if st.button("📂 加载进度文件"):
                try:
                    progress_path = os.path.join(save_directory, selected_progress_file)
                    df = pd.read_excel(progress_path)
                    df.columns = df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                    st.session_state.current_file = df
                    
                    progress_info = []
                    for lang, col_name in st.session_state.language_configs['column_names'].items():
                        if col_name in df.columns:
                            translated_count = df[col_name].notna().sum()
                            total_count = len(df)
                            progress_info.append(f"{lang}: {translated_count}/{total_count}")
                    
                    st.success(f"✅ 成功加载进度文件！")
                    if progress_info:
                        st.info(f"📊 翻译进度: {', '.join(progress_info)}")
                    
                    with st.expander("📊 文件预览"):
                        st.dataframe(df.head(10))
                except Exception as e:
                    st.error(f"❌ 加载进度文件失败: {e}")
        else:
            uploaded_file = st.file_uploader(
                "📄 上传翻译文件 (Excel)",
                type=['xlsx', 'xls', 'csv'],
                key="batch_file_uploader"
            )

            if uploaded_file is not None:
                try:
                    if uploaded_file.name.endswith('.csv'):
                        df = pd.read_csv(uploaded_file)
                    else:
                        df = pd.read_excel(uploaded_file)
                    df.columns = df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                    st.session_state.current_file = df
                    st.success(f"✅ 成功读取文件，共 {len(df)} 行数据")

                    with st.expander("📊 文件预览"):
                        st.dataframe(df.head(10))

                except Exception as e:
                    st.error(f"❌ 文件读取失败: {e}")
        
        if st.session_state.current_file is not None:
            df = st.session_state.current_file
            cols = df.columns.tolist()
            text_col = st.selectbox(
                "📝 选择文本列",
                options=cols,
                index=0,
                key="batch_text_col"
            )

            role_col = st.selectbox(
                "👥 选择角色列 (可选)",
                options=["无"] + cols,
                index=0,
                key="batch_role_col"
            )
            if role_col == "无":
                role_col = None

    with col2:
        st.header("📚 术语库和性格库")

        uploaded_term_base = st.file_uploader(
            "📚 上传多语言术语库 (Excel)",
            type=['xlsx', 'xls'],
            key="batch_term_base_uploader",
            help="术语库应包含原文列和多个目标语言列"
        )

        if uploaded_term_base is not None:
            try:
                term_df = pd.read_excel(uploaded_term_base)
                term_df.columns = term_df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                st.session_state.term_base_df = term_df
                st.success(f"✅ 成功读取术语库，共 {len(term_df)} 条术语")

                with st.expander("📋 配置术语库列映射", expanded=True):
                    term_cols = term_df.columns.tolist()
                    
                    source_col = st.selectbox(
                        "📤 选择原文列（中文）",
                        options=term_cols,
                        index=0,
                        key="batch_source_col"
                    )
                    
                    st.markdown("---")
                    st.subheader("🌐 为每种语言选择对应的术语列")
                    st.info("💡 提示：为每种目标语言选择术语库中对应的翻译列")
                    
                    term_language_mapping = {}
                    for lang in selected_languages:
                        st.markdown(f"**{lang} 术语列：**")
                        target_col = st.selectbox(
                            f"选择 {lang} 对应的术语列",
                            options=["不使用术语库"] + term_cols,
                            index=0,
                            key=f"term_col_{lang}",
                            help=f"选择术语库中 {lang} 翻译对应的列"
                        )
                        
                        if target_col != "不使用术语库":
                            term_language_mapping[lang] = target_col
                            
                            sample_terms = term_df[[source_col, target_col]].head(3).dropna()
                            if not sample_terms.empty:
                                st.caption(f"示例：")
                                for _, row in sample_terms.iterrows():
                                    st.caption(f"  • {row[source_col]} → {row[target_col]}")
                    
                    st.session_state.term_language_mapping = term_language_mapping
                    
                    if term_language_mapping:
                        st.success(f"✅ 已配置 {len(term_language_mapping)} 种语言的术语映射")
                    else:
                        st.warning("⚠️ 未配置任何语言的术语映射")

            except Exception as e:
                st.error(f"❌ 术语库读取失败: {e}")

        uploaded_role_personality = st.file_uploader(
            "📋 上传角色性格库文件 (Excel)",
            type=['xlsx', 'xls'],
            key="batch_role_personality_uploader"
        )

        if uploaded_role_personality is not None:
            try:
                role_personality_df = pd.read_excel(uploaded_role_personality)
                role_personality_df.columns = role_personality_df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                st.session_state.role_personality_df = role_personality_df
                st.success(f"✅ 成功读取角色性格库，共 {len(role_personality_df)} 条记录")

                role_personality_cols = role_personality_df.columns.tolist()
                role_name_col = st.selectbox(
                    "👥 选择角色名称列",
                    options=role_personality_cols,
                    index=0,
                    key="batch_role_name_col"
                )

                personality_col = st.selectbox(
                    "💬 选择性格描述列",
                    options=role_personality_cols,
                    index=min(1, len(role_personality_cols) - 1) if len(role_personality_cols) > 1 else 0,
                    key="batch_personality_col"
                )

            except Exception as e:
                st.error(f"❌ 角色性格库读取失败: {e}")

    st.header("🎯 翻译要求设置")

    custom_requirements = st.text_area(
        "💬 自定义翻译要求（适用于所有语言）",
        value="角色对话自然流畅；专业术语统一；保持原文风格；本地化适配；保持上下文一致性；根据角色调整语气;请注意使用语体，且所有角色除了微型机和班长，其他都为女生用语，不要用男性用语，现在角色们都十分熟悉彼此了，不需要使用太正式尊重的语体了例如日语的话不需要ですます型了。",
        height=100,
        key="batch_custom_requirements"
    )

    if st.button("🔧 初始化翻译器", type="secondary", use_container_width=True):
        if not api_key:
            st.error("❌ 请先输入API密钥")
        elif not selected_languages:
            st.error("❌ 请至少选择一种目标语言")
        else:
            try:
                translator = MultiAPIExcelTranslator(
                    api_key, api_provider, api_url, model,
                    context_size, max_retries
                )
                translator.enable_fuzzy_match = enable_fuzzy
                translator.fuzzy_threshold = fuzzy_threshold
                translator.set_target_languages(selected_languages, language_column_names)

                st.session_state.translator = translator
                st.session_state.role_matching_confirmed = False

                st.info(f"🌍 使用 {api_provider} API")
                st.info(f"🤖 使用模型: {model}")
                st.info(f"🎯 目标语言: {', '.join(selected_languages)}")
                
                with st.expander("📋 查看列名配置"):
                    for lang, col_name in language_column_names.items():
                        st.write(f"• {lang} → `{col_name}`")

                if st.session_state.term_base_df is not None and st.session_state.term_language_mapping:
                    if translator.load_term_base_multilang(
                        st.session_state.term_base_df, 
                        source_col, 
                        st.session_state.term_language_mapping
                    ):
                        st.success("✅ 多语言术语库加载成功")
                elif st.session_state.term_base_df is not None:
                    st.warning("⚠️ 术语库已上传但未配置语言映射")

                if st.session_state.role_personality_df is not None:
                    if translator.load_role_personality(
                            st.session_state.role_personality_df,
                            role_name_col,
                            personality_col
                    ):
                        st.success("✅ 角色性格库加载成功")

                st.success("✅ 翻译器初始化完成！")

            except Exception as e:
                st.error(f"❌ 初始化失败: {e}")
                import traceback
                st.error(traceback.format_exc())

    if (st.session_state.translator is not None and
            st.session_state.current_file is not None and
            role_col is not None and
            enable_fuzzy and
            not st.session_state.role_matching_confirmed):

        st.markdown("---")
        confirmed = render_role_matching_interface(
            st.session_state.translator,
            st.session_state.current_file,
            role_col
        )
        if confirmed:
            st.session_state.role_matching_confirmed = True
            st.rerun()

    st.markdown("---")
    translation_ready = (
            st.session_state.translator is not None and
            st.session_state.current_file is not None and
            selected_languages and
            (not enable_fuzzy or not role_col or st.session_state.role_matching_confirmed)
    )

    if not translation_ready and enable_fuzzy and role_col:
        st.info("💡 请先完成角色匹配确认后再开始翻译")

    if st.button(
            "🎯 开始多语言翻译",
            type="primary",
            use_container_width=True,
            disabled=not translation_ready,
            key="batch_start_translation"
    ):
        try:
            translator = st.session_state.translator
            df = st.session_state.current_file.copy()
            languages = st.session_state.language_configs['languages']
            column_names = st.session_state.language_configs['column_names']

            os.makedirs(save_directory, exist_ok=True)

            with st.expander("📋 查看翻译配置", expanded=True):
                st.write("**语言配置：**")
                for lang, col_name in column_names.items():
                    term_status = "✅ 已配置术语库" if lang in translator.term_base_dict and translator.term_base_dict[lang] else "⚠️ 未配置术语库"
                    st.write(f"• {lang} → `{col_name}` ({term_status})")
                
                if translator.role_mapping:
                    st.write("**角色映射：**")
                    for orig, mapped in translator.role_mapping.items():
                        st.write(f"• `{orig}` → `{mapped}`")

            for lang in languages:
                col_name = column_names[lang]
                if col_name not in df.columns:
                    df[col_name] = ''
            
            start_index = 0
            if resume_mode:
                min_translated_index = len(df)
                for lang in languages:
                    col_name = column_names[lang]
                    if col_name in df.columns:
                        last_index = -1
                        for idx in range(len(df)):
                            if not pd.isna(df.at[idx, col_name]) and str(df.at[idx, col_name]).strip() != '':
                                if not str(df.at[idx, col_name]).startswith('[翻译失败'):
                                    last_index = idx
                        min_translated_index = min(min_translated_index, last_index + 1)
                
                start_index = min_translated_index
                if start_index > 0:
                    st.info(f"🔄 继续翻译：跳过前 {start_index} 行（已翻译）")

            progress_bar = st.progress(0)
            status_text = st.empty()
            
            stats = {lang: {'success': 0, 'error': 0} for lang in languages}
            
            total_rows = len(df)
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            progress_filename = f"translation_progress_multilang_{timestamp}.xlsx"
            progress_path = os.path.join(save_directory, progress_filename)

            try:
                for index in range(start_index, total_rows):
                    row = df.iloc[index]
                    progress = (index + 1) / total_rows
                    progress_bar.progress(progress)
                    
                    stats_str = " | ".join([f"{lang}: ✓{stats[lang]['success']} ✗{stats[lang]['error']}" for lang in languages])
                    status_text.text(f"📝 正在翻译第 {index + 1}/{total_rows} 行... | {stats_str}")

                    text = str(row[text_col])
                    role = row[role_col] if role_col and role_col in row else None

                    if pd.isna(text) == "" or str(text).strip() == "" or text == "nan":
                        print("为空")
                        continue
                    
                    for lang in languages:
                        col_name = column_names[lang]
                        
                        existing_translation = df.at[index, col_name]
                        if not pd.isna(existing_translation) and str(existing_translation).strip() != '' and not str(existing_translation).startswith('[翻译失败'):
                            stats[lang]['success'] += 1
                            continue

                        try:
                            translated_text = translator.translate_text(
                                text, lang, custom_requirements, role
                            )
                            
                            df.at[index, col_name] = translated_text
                            stats[lang]['success'] += 1
                            
                        except Exception as e:
                            error_msg = str(e)
                            st.warning(f"⚠️ [{lang}] 第 {index + 1} 行翻译失败: {error_msg}")
                            df.at[index, col_name] = f"[翻译失败: {error_msg}]"
                            stats[lang]['error'] += 1
                        
                        time.sleep(0.15)

                    if (index + 1) % auto_save_interval == 0:
                        try:
                            with pd.ExcelWriter(progress_path, engine='openpyxl') as writer:
                                df.to_excel(writer, index=False, sheet_name='翻译进度')
                            st.info(f"💾 已自动保存进度: {index + 1}/{total_rows} 行")
                        except Exception as save_error:
                            st.warning(f"⚠️ 自动保存失败: {save_error}")

                final_filename = f"translation_final_multilang_{timestamp}.xlsx"
                final_path = os.path.join(save_directory, final_filename)
                
                with pd.ExcelWriter(final_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='翻译结果')
                
                progress_bar.progress(1.0)
                
                st.success("✅ 多语言翻译完成！")
                
                stats_cols = st.columns(len(languages))
                for idx, lang in enumerate(languages):
                    with stats_cols[idx]:
                        st.metric(
                            f"{lang}",
                            f"✓ {stats[lang]['success']}",
                            f"✗ {stats[lang]['error']}" if stats[lang]['error'] > 0 else None,
                            delta_color="inverse"
                        )

                st.subheader("📊 翻译结果预览")
                
                display_cols = [text_col]
                if role_col:
                    display_cols.append(role_col)
                display_cols.extend([column_names[lang] for lang in languages])
                
                st.dataframe(df[display_cols].head(20))

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='翻译结果')

                st.download_button(
                    label="💾 下载多语言翻译结果",
                    data=output.getvalue(),
                    file_name=f"translated_multilang_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.success(f"📁 文件已自动保存至: {final_path}")

            except KeyboardInterrupt:
                st.warning("⚠️ 翻译被中断，正在保存当前进度...")
                try:
                    interrupt_filename = f"translation_interrupted_{timestamp}.xlsx"
                    interrupt_path = os.path.join(save_directory, interrupt_filename)
                    with pd.ExcelWriter(interrupt_path, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='翻译进度')
                    st.info(f"💾 进度已保存至: {interrupt_path}")
                except Exception as save_error:
                    st.error(f"❌ 保存进度失败: {save_error}")

        except Exception as e:
            st.error(f"❌ 翻译过程中出现错误: {e}")
            import traceback
            st.error(traceback.format_exc())
            
            try:
                error_filename = f"translation_error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                error_path = os.path.join(save_directory, error_filename)
                with pd.ExcelWriter(error_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='翻译进度')
                st.info(f"💾 错误前的进度已保存至: {error_path}")
            except Exception as save_error:
                st.error(f"❌ 保存进度失败: {save_error}")
