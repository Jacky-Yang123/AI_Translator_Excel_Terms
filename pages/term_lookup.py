# pages/term_lookup.py - 术语查找页面

import pandas as pd
import streamlit as st

from translator import MultiAPIExcelTranslator
from api_config import get_api_providers, get_preset_languages


def term_lookup_page():
    st.title("🔍 术语和角色查找工具")
    st.markdown("### 在已加载的术语库和角色性格库中搜索")

    if 'lookup_translator' not in st.session_state:
        st.session_state.lookup_translator = MultiAPIExcelTranslator(
            api_key="",
            api_provider="DeepSeek",
            api_url=get_api_providers()["DeepSeek"]["url"],
            model="deepseek-chat"
        )

    translator = st.session_state.lookup_translator

    col1, col2 = st.columns(2)

    with col1:
        st.header("📚 术语库")

        uploaded_term_base = st.file_uploader(
            "📚 上传术语库文件 (Excel)",
            type=['xlsx', 'xls'],
            key="lookup_term_base_uploader"
        )

        if uploaded_term_base is not None:
            try:
                df = pd.read_excel(uploaded_term_base)
                df.columns = df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                st.success(f"✅ 成功读取术语库，共 {len(df)} 条记录")

                with st.expander("📊 术语库预览"):
                    st.dataframe(df.head(10))

                cols = df.columns.tolist()

                source_col = st.selectbox(
                    "📝 选择中文源文列",
                    options=cols,
                    index=0,
                    key="lookup_term_source_col"
                )

                target_col = st.selectbox(
                    "📤 选择翻译列",
                    options=cols,
                    index=min(1, len(cols)-1) if len(cols) > 1 else 0,
                    key="lookup_term_target_col"
                )

                if st.button("📥 加载术语库", key="lookup_load_term_base"):
                    if translator.load_term_base(df, source_col, target_col):
                        st.session_state.lookup_term_loaded = True
                        st.success("✅ 术语库加载成功")
                        st.rerun()

            except Exception as e:
                st.error(f"❌ 处理术语库文件失败: {e}")

        if st.session_state.get('lookup_term_loaded', False):
            st.info(f"✅ 术语库已加载: {len(translator.term_base_list)} 条术语")

    with col2:
        st.header("👤 角色性格库")

        uploaded_role = st.file_uploader(
            "📋 上传角色性格库文件 (Excel)",
            type=['xlsx', 'xls'],
            key="lookup_role_uploader"
        )

        if uploaded_role is not None:
            try:
                df = pd.read_excel(uploaded_role)
                df.columns = df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                st.success(f"✅ 成功读取角色性格库，共 {len(df)} 条记录")

                with st.expander("📊 角色性格库预览"):
                    st.dataframe(df.head(10))

                cols = df.columns.tolist()

                role_col = st.selectbox(
                    "👥 选择角色名称列",
                    options=cols,
                    index=0,
                    key="lookup_role_name_col"
                )

                personality_col = st.selectbox(
                    "💬 选择性格描述列",
                    options=cols,
                    index=min(1, len(cols)-1) if len(cols) > 1 else 0,
                    key="lookup_personality_col"
                )

                if st.button("📥 加载角色性格库", key="lookup_load_role"):
                    if translator.load_role_personality(df, role_col, personality_col):
                        st.session_state.lookup_role_loaded = True
                        st.success("✅ 角色性格库加载成功")
                        st.rerun()

            except Exception as e:
                st.error(f"❌ 处理角色性格库文件失败: {e}")

        if st.session_state.get('lookup_role_loaded', False):
            st.info(f"✅ 角色性格库已加载: {len(translator.role_personality_dict)} 个角色")

    st.divider()

    st.header("🔎 查找功能")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📚 术语搜索")
        term_search = st.text_input("输入要搜索的术语:", key="term_search_input")

        if st.button("🔍 搜索术语", key="search_term_btn"):
            if not term_search:
                st.warning("请输入要搜索的术语")
            elif not translator.term_base_list:
                st.warning("请先加载术语库")
            else:
                found = []
                for term in translator.term_base_list:
                    if term_search.lower() in term['source'].lower() or term_search.lower() in term['target'].lower():
                        found.append(term)

                if found:
                    st.success(f"找到 {len(found)} 条匹配术语")
                    for i, term in enumerate(found[:20]):
                        st.write(f"{i+1}. {term['source']} → {term['target']}")
                    if len(found) > 20:
                        st.info(f"... 还有 {len(found) - 20} 条")
                else:
                    st.info("未找到匹配的术语")

    with col2:
        st.subheader("👤 角色搜索")
        role_search = st.text_input("输入要搜索的角色名:", key="role_search_input")

        if st.button("🔍 搜索角色", key="search_role_btn"):
            if not role_search:
                st.warning("请输入要搜索的角色名")
            elif not translator.role_personality_dict:
                st.warning("请先加载角色性格库")
            else:
                found = {}
                for role, personality in translator.role_personality_dict.items():
                    if role_search.lower() in role.lower():
                        found[role] = personality

                if found:
                    st.success(f"找到 {len(found)} 个匹配角色")
                    for role, personality in list(found.items())[:10]:
                        st.write(f"**{role}**: {personality[:100]}..." if len(personality) > 100 else f"**{role}**: {personality}")
                    if len(found) > 10:
                        st.info(f"... 还有 {len(found) - 10} 个")
                else:
                    st.info("未找到匹配的角色")

    st.divider()

    st.header("📊 统计信息")

    col1, col2 = st.columns(2)

    with col1:
        if translator.term_base_list:
            st.metric("术语库条目", len(translator.term_base_list))
        else:
            st.metric("术语库条目", 0)

    with col2:
        if translator.role_personality_dict:
            st.metric("角色数量", len(translator.role_personality_dict))
        else:
            st.metric("角色数量", 0)
