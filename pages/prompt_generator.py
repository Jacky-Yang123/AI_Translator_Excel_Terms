import pandas as pd
import streamlit as st
import re
import shutil
import openpyxl
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from translator import MultiAPIExcelTranslator, get_api_providers
from pages.helpers import get_preset_options, get_preset_languages, get_default_custom_requirements


def prompt_generator_page():
    st.title("📝 单语种翻译提示词生成器")
    st.markdown("### 根据待翻译文本、术语库和角色性格信息，生成针对单一目标语言的翻译提示词。")
    st.markdown("**注意：** 本页面仅用于生成提示词文本，不进行实际的API翻译调用。")
    
    if 'prompt_translator' not in st.session_state:
        st.session_state.prompt_translator = MultiAPIExcelTranslator(
            api_key="", 
            api_provider="DeepSeek", 
            api_url=get_api_providers()["DeepSeek"]["url"], 
            model="deepseek-chat"
        )
    
    translator = st.session_state.prompt_translator
    
    if 'term_base_loaded' not in st.session_state:
        st.session_state.term_base_loaded = False
    if 'role_personality_loaded' not in st.session_state:
        st.session_state.role_personality_loaded = False
    
    st.header("🎯 基本设置")
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        language_option = st.selectbox(
            "🌍 选择目标语言",
            options=get_preset_languages(),
            index=0,
            key="prompt_language_option"
        )
    
    with col2:
        if language_option == "自定义":
            custom_language = st.text_input(
                "✏️ 输入自定义语言",
                value=st.session_state.get('prompt_custom_language', ''),
                placeholder="例如：俄文、葡萄牙文、阿拉伯文等",
                key="prompt_custom_language_input"
            )
            st.session_state.prompt_custom_language = custom_language
            target_language = custom_language
        else:
            target_language = language_option
            st.session_state.prompt_custom_language = ""
    
    if target_language:
        translator.set_target_language(target_language)
        st.info(f"🎯 当前目标语言: {target_language}")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.header("📁 待翻译文本")
        
        uploaded_file = st.file_uploader(
            "📄 上传待翻译文本文件 (Excel)",
            type=['xlsx', 'xls'],
            key="prompt_file_uploader"
        )
        
        df_text = None
        text_col = None
        role_col = None
        personality_col = None
        
        if uploaded_file is not None:
            try:
                df_text = pd.read_excel(uploaded_file)
                df_text.columns = df_text.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                st.session_state.prompt_df_text = df_text
                st.success(f"✅ 成功读取文件，共 {len(df_text)} 行数据")
                
                with st.expander("📊 文件预览"):
                    st.dataframe(df_text.head(10))
                
                cols = df_text.columns.tolist()
                text_col = st.selectbox(
                    "📝 选择文本列",
                    options=cols,
                    index=0,
                    key="prompt_text_col_select"
                )
                
                role_col = st.selectbox(
                    "👥 选择说话人/角色列 (可选)",
                    options=["无"] + cols,
                    index=0,
                    key="prompt_role_col_select"
                )
                role_col = role_col if role_col != "无" else None
                
                personality_col = st.selectbox(
                    "💬 选择性格描述列 (可选)",
                    options=["无"] + cols,
                    index=0,
                    key="prompt_personality_col_select"
                )
                personality_col = personality_col if personality_col != "无" else None
                
                st.session_state.prompt_text_col = text_col
                st.session_state.prompt_role_col = role_col
                st.session_state.prompt_personality_col = personality_col
                
            except Exception as e:
                st.error(f"❌ 文件读取失败: {e}")
        else:
            if 'prompt_df_text' in st.session_state:
                df_text = st.session_state.prompt_df_text
                text_col = st.session_state.get('prompt_text_col')
                role_col = st.session_state.get('prompt_role_col')
                personality_col = st.session_state.get('prompt_personality_col')
                
                if df_text is not None:
                    st.info(f"✅ 已加载文件：{len(df_text)} 行数据")
                    if text_col:
                        st.write(f"📝 文本列: {text_col}")
                    if role_col:
                        st.write(f"👥 角色列: {role_col}")
                    if personality_col:
                        st.write(f"💬 性格列: {personality_col}")
        
        st.subheader("✂️ 分批次设置")
        batch_size = st.number_input(
            "每批次行数",
            min_value=1,
            max_value=200,
            value=50,
            step=10,
            key="prompt_batch_size"
        )
    
    with col2:
        st.header("📚 术语库和性格库")
        
        st.subheader("📚 术语库功能")
        uploaded_term_base = st.file_uploader(
            "📚 上传术语库文件 (Excel)",
            type=['xlsx', 'xls'],
            key="prompt_term_base_uploader"
        )
        
        if uploaded_term_base is not None:
            try:
                df = pd.read_excel(uploaded_term_base)
                df.columns = df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                st.session_state.prompt_term_base_df = df
                st.success(f"✅ 成功读取术语库，共 {len(df)} 条记录")
                
                with st.expander("📊 术语库预览"):
                    st.dataframe(df.head(10))
                
                cols = df.columns.tolist()
                
                source_col = st.selectbox(
                    "📝 选择中文源文列",
                    options=cols,
                    index=0,
                    key="prompt_term_source_col"
                )
                
                target_col = st.selectbox(
                    "📤 选择翻译列",
                    options=cols,
                    index=min(1, len(cols)-1) if len(cols) > 1 else 0,
                    key="prompt_term_target_col"
                )
                
                if st.button("📥 加载术语库", key="prompt_load_term_base"):
                    if translator.load_term_base(df, source_col, target_col):
                        st.session_state.term_base_loaded = True
                        st.success("✅ 术语库加载成功")
                        st.rerun()
                
            except Exception as e:
                st.error(f"❌ 处理术语库文件失败: {e}")
        
        if st.session_state.get('term_base_loaded', False):
            st.info(f"✅ 术语库已加载: {len(translator.term_base_list)} 条术语")
        
        st.divider()
        
        st.subheader("👤 角色性格库功能")
        uploaded_role = st.file_uploader(
            "📋 上传角色性格库文件 (Excel)",
            type=['xlsx', 'xls'],
            key="prompt_role_personality_uploader"
        )
        
        if uploaded_role is not None:
            try:
                df = pd.read_excel(uploaded_role)
                df.columns = df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                st.session_state.prompt_role_personality_df = df
                st.success(f"✅ 成功读取角色性格库，共 {len(df)} 条记录")
                
                with st.expander("📊 角色性格库预览"):
                    st.dataframe(df.head(10))
                
                cols = df.columns.tolist()
                role_col = st.selectbox(
                    "👥 选择角色名称列",
                    options=cols,
                    index=0,
                    key="prompt_role_name_col"
                )
                
                personality_col = st.selectbox(
                    "💬 选择性格描述列",
                    options=cols,
                    index=min(1, len(cols)-1) if len(cols) > 1 else 0,
                    key="prompt_personality_desc_col"
                )
                
                if st.button("📥 加载角色性格库", key="prompt_load_role_personality"):
                    if translator.load_role_personality(df, role_col, personality_col):
                        st.session_state.role_personality_loaded = True
                        st.success("✅ 角色性格库加载成功")
                        st.rerun()
                
            except Exception as e:
                st.error(f"❌ 处理角色性格库文件失败: {e}")
        
        if st.session_state.get('role_personality_loaded', False):
            st.info(f"✅ 角色性格库已加载: {len(translator.role_personality_dict)} 条角色")
    
    st.divider()
    
    st.header("🎯 翻译要求设置")
    
    st.subheader("🏷️ 预设选项")
    presets = get_preset_options()
    preset_options = st.multiselect(
        "选择预设翻译要求（可多选）",
        options=list(presets.keys()),
        default=st.session_state.get('prompt_preset_options', []),
        key="prompt_preset_multiselect"
    )
    st.session_state.prompt_preset_options = preset_options
    
    custom_requirements = st.text_area(
        "💬 自定义翻译要求",
        value=st.session_state.get('prompt_custom_requirements', get_default_custom_requirements()),
        placeholder=f"例如：游戏UI简约风格、角色对话自然流畅、专业术语统一、{target_language}本地化适配等",
        height=100,
        key="prompt_custom_requirements_text"
    )
    st.session_state.prompt_custom_requirements = custom_requirements
    
    all_requirements = []
    if preset_options:
        all_requirements.extend(preset_options)
    if custom_requirements:
        all_requirements.append(custom_requirements)
    
    final_requirements = "；".join(all_requirements) if all_requirements else ""
    
    if st.button("🚀 生成提示词", key="generate_prompt_btn", use_container_width=True):
        if df_text is None or text_col is None:
            st.error("❌ 请先上传待翻译文本文件并选择文本列。")
            return
        
        text_col = st.session_state.get('prompt_text_col')
        role_col = st.session_state.get('prompt_role_col')
        personality_col = st.session_state.get('prompt_personality_col')
        
        if not text_col:
            st.error("❌ 请选择文本列。")
            return
        
        if not target_language or target_language.strip() == "":
            st.error("❌ 请先选择或输入目标语言。")
            return
        
        term_base_loaded = st.session_state.get('term_base_loaded', False)
        role_personality_loaded = st.session_state.get('role_personality_loaded', False)
        
        if term_base_loaded:
            st.info(f"✅ 术语库已加载: {len(translator.term_base_list)} 条术语")
        else:
            st.warning("⚠️ 未加载术语库，提示词中将不包含术语匹配信息")
        
        if role_personality_loaded:
            st.info(f"✅ 角色性格库已加载: {len(translator.role_personality_dict)} 条角色")
        else:
            st.warning("⚠️ 未加载角色性格库，提示词中将不包含角色性格信息")
        
        fixed_requirements = f"""
## 翻译要求：(固定)
你是一名专业的二次元游戏本地化翻译专家，擅长将中文二次元游戏文案翻译为{target_language}。请将用户输入的中文游戏文本，以表格形式输出对应的{target_language}翻译。表格应包含两列：原文（中文）、{target_language}翻译。"""
        
        language_specific_requirements = translator.get_language_specific_requirements(target_language)
        
        other_requirements = f"""
## 其他要求：(用户输入)
{final_requirements if final_requirements else "无"}
"""
        
        important_notes = f"""
## 重要说明！：(固定)
• 请只返回翻译后的文本结果，以表格形式输出中文原文，{target_language}翻译
• 不要添加任何解释或备注
• 术语库中的特定词汇翻译需要严格采用相同的翻译
• 请根据角色性格特点调整翻译风格和语气。
• 本次翻译目标语言为：{target_language}
"""
        
        num_batches = (len(df_text) + batch_size - 1) // batch_size
        all_prompts = []
        
        for i in range(num_batches):
            start_index = i * batch_size
            end_index = min((i + 1) * batch_size, len(df_text))
            batch_df = df_text.iloc[start_index:end_index].copy()
            
            translator.reset_context()
            
            text_list = batch_df[text_col].tolist()
            
            if role_col and role_col != "无" and role_col in batch_df.columns:
                role_list = batch_df[role_col].tolist()
            else:
                role_list = [None] * len(batch_df)
            
            if personality_col and personality_col != "无" and personality_col in batch_df.columns:
                personality_list = batch_df[personality_col].tolist()
            else:
                personality_list = [None] * len(batch_df)
            
            all_text_in_batch = " ".join([str(t) for t in text_list if not pd.isna(t)])
            
            term_base_prompt = ""
            if term_base_loaded:
                term_base_prompt = translator.build_term_base_prompt(all_text_in_batch)
            else:
                term_base_prompt = "\n\n### 术语库匹配：\n无术语库加载，跳过术语匹配。"
            
            text_and_personality_prompt = "### 待翻译文本及说话人性格格式：(用户输入)\n"
            text_and_personality_prompt += "说话人\t原文\t说话人性格\n"
            
            for j in range(len(batch_df)):
                role = role_list[j]
                text = text_list[j]
                personality = personality_list[j]
                
                if personality_col and not pd.isna(personality):
                    personality_desc = str(personality).strip()
                elif role_col and role and role_personality_loaded:
                    personality_desc = translator.find_role_personality(role)
                    personality_desc = personality_desc if personality_desc else "无"
                else:
                    personality_desc = "无"
                
                role_name = str(role).strip() if role and not pd.isna(role) else "无"
                text_content = str(text).strip() if not pd.isna(text) else ""
                
                text_and_personality_prompt += f"{role_name}\t{text_content}\t{personality_desc}\n"
            
            full_prompt = f"""
# 批次 {i+1}/{num_batches} - 目标语言: {target_language}

{fixed_requirements}

{language_specific_requirements}

{other_requirements}

{term_base_prompt}

{text_and_personality_prompt}

{important_notes}
"""
            all_prompts.append(full_prompt.strip())
        
        st.subheader("✅ 生成结果")
        
        st.session_state.all_prompts = all_prompts
        st.session_state.num_batches = num_batches
        st.session_state.current_batch_index = 0
        st.session_state.target_language = target_language
        
        st.success(f"✅ 提示词生成成功，共 {num_batches} 个批次，目标语言: {target_language}。")
    
    if st.session_state.get('all_prompts'):
        all_prompts = st.session_state.all_prompts
        num_batches = st.session_state.num_batches
        current_batch_index = st.session_state.current_batch_index
        target_language = st.session_state.get('target_language', '英文')
        
        st.subheader(f"批次 {current_batch_index + 1}/{num_batches} 提示词 - 目标语言: {target_language}")
        
        current_prompt = all_prompts[current_batch_index]
        
        st.code(current_prompt, language=None)
        st.info("👆 请使用上方代码块右下角的复制按钮进行一键复制。")
        
        col_prev, col_info, col_next = st.columns([1, 2, 1])
        
        with col_prev:
            if st.button("上一批次", disabled=(current_batch_index == 0), key="prompt_prev_batch"):
                st.session_state.current_batch_index -= 1
                st.rerun()
                
        with col_info:
            st.markdown(f"<p style='text-align: center;'>当前批次: {current_batch_index + 1} / {num_batches} | 目标语言: {target_language}</p>", unsafe_allow_html=True)
            
        with col_next:
            if st.button("下一批次", disabled=(current_batch_index == num_batches - 1), key="prompt_next_batch"):
                st.session_state.current_batch_index += 1
                st.rerun()
                
        st.markdown("---")
        
        final_output = f"# 翻译提示词 - 目标语言: {target_language}\n\n" + ("-"*80) + "\n\n".join(all_prompts)
        
        st.download_button(
            label=f"📥 下载所有提示词 (.txt)",
            data=final_output,
            file_name=f"translation_prompts_{target_language}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            mime="text/plain",
            use_container_width=True
        )
