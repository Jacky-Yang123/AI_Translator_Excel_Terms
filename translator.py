# translator.py - 核心翻译器类

import re
import time
import difflib
import requests
import pandas as pd
import streamlit as st
import jieba


class MultiAPIExcelTranslator:
    def __init__(self, api_key, api_provider, api_url, model, context_size=10, max_retries=10):
        self.api_key = api_key
        self.api_provider = api_provider
        self.api_url = api_url
        self.model = model
        self.headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        self.context_history = {}  # 按语言分别存储上下文
        self.term_dict = {}
        self.role_column = None
        self.context_size = context_size
        self.max_retries = max_retries

        # 改为按语言存储术语库
        self.term_base_dict = {}  # {语言: [{source: xxx, target: xxx}]}
        self.term_base_list = []  # 单语言术语库列表

        self.role_personality_dict = {}
        self.current_text_terms = {}
        self.current_role_personality = None
        self.target_languages = ["英文"]
        self.language_column_names = {"英文": "英文翻译结果"}
        self.target_language = "英文"

        # 新增：角色映射表
        self.role_mapping = {}
        self.enable_fuzzy_match = False
        self.fuzzy_threshold = 0.6

        self.init_chinese_tokenizer()

    def init_chinese_tokenizer(self):
        try:
            self.chinese_tokenizer = jieba
            st.success("✅ 中文分词器初始化成功")
        except Exception as e:
            st.warning(f"⚠️ 中文分词器初始化失败: {e}")
            self.chinese_tokenizer = None

    def tokenize_chinese_text(self, text):
        if not text or pd.isna(text):
            return []

        text = str(text).strip()
        if not text:
            return []

        try:
            if self.chinese_tokenizer:
                words = self.chinese_tokenizer.cut(text)
                return [word for word in words if word.strip() and re.search(r'[\w\u4e00-\u9fa5]', word)]
            else:
                return [char for char in text if char.strip() and re.search(r'[\w\u4e00-\u9fa5]', char)]
        except Exception as e:
            st.warning(f"⚠️ 中文分词失败: {e}")
            return [char for char in text if char.strip()]

    def clean_role_name(self, role_name):
        """清理角色名称：去除额外标记、空格等"""
        if not role_name or pd.isna(role_name):
            return ""

        role_name = str(role_name).strip()
        role_name = re.sub(r'\|.*$', '', role_name)
        role_name = re.sub(r'[\s\u3000]+', '', role_name)
        role_name = role_name.strip()

        return role_name

    def fuzzy_match_role(self, role_name, threshold=None):
        """模糊匹配角色名称"""
        if not role_name or not self.role_personality_dict:
            return None, 0

        if threshold is None:
            threshold = self.fuzzy_threshold

        cleaned_role = self.clean_role_name(role_name)
        if not cleaned_role:
            return None, 0

        if role_name in self.role_mapping:
            return self.role_mapping[role_name], 1.0

        best_match = None
        best_score = 0

        for official_role in self.role_personality_dict.keys():
            cleaned_official = self.clean_role_name(official_role)

            if cleaned_role == cleaned_official:
                return official_role, 1.0

            score = difflib.SequenceMatcher(None, cleaned_role, cleaned_official).ratio()

            if cleaned_role in cleaned_official or cleaned_official in cleaned_role:
                score = max(score, 0.8)

            if score > best_score:
                best_score = score
                best_match = official_role

        if best_score >= threshold:
            return best_match, best_score

        return None, best_score

    def analyze_role_matches(self, df, role_col):
        """分析数据中的所有角色名称"""
        if not role_col or role_col not in df.columns:
            return {}

        unique_roles = df[role_col].dropna().unique()
        fuzzy_matches = {}

        for role in unique_roles:
            role_str = str(role).strip()
            if not role_str:
                continue

            if role_str in self.role_mapping:
                continue

            if role_str in self.role_personality_dict:
                continue

            matched_role, score = self.fuzzy_match_role(role_str)

            if matched_role:
                if role_str not in fuzzy_matches:
                    fuzzy_matches[role_str] = []
                fuzzy_matches[role_str].append((matched_role, score))

                if score < 1.0:
                    for official_role in self.role_personality_dict.keys():
                        if official_role == matched_role:
                            continue
                        alt_score = difflib.SequenceMatcher(
                            None,
                            self.clean_role_name(role_str),
                            self.clean_role_name(official_role)
                        ).ratio()

                        if alt_score >= self.fuzzy_threshold * 0.8:
                            fuzzy_matches[role_str].append((official_role, alt_score))

                fuzzy_matches[role_str].sort(key=lambda x: x[1], reverse=True)

        return fuzzy_matches

    def find_role_personality(self, role_name):
        """查找角色性格描述"""
        if not role_name or not self.role_personality_dict:
            return None

        role_name = str(role_name).strip()
        if not role_name:
            return None

        if role_name in self.role_mapping:
            mapped_role = self.role_mapping[role_name]
            return self.role_personality_dict.get(mapped_role)

        if role_name in self.role_personality_dict:
            return self.role_personality_dict[role_name]

        if self.enable_fuzzy_match:
            matched_role, score = self.fuzzy_match_role(role_name)
            if matched_role:
                return self.role_personality_dict[matched_role]

        return None

    def add_to_context(self, original, translation, role=None, language="英文"):
        """为指定语言添加上下文（确保语言独立）"""
        if language not in self.context_history:
            self.context_history[language] = []

        self.context_history[language].append((original, translation, role))
        if len(self.context_history[language]) > self.context_size:
            self.context_history[language].pop(0)

    def build_context_prompt(self, language="英文"):
        """为指定语言构建上下文提示（确保语言独立）"""
        if language not in self.context_history or not self.context_history[language]:
            return ""

        context_str = f"\n\n### 重要上下文参考（{language}翻译）：\n"
        for i, (orig, trans, role) in enumerate(self.context_history[language], 1):
            role_info = f" [{role}]" if role else ""
            context_str += f"前文{i}{role_info}:\n原文: {orig}\n{language}译文: {trans}\n\n"
        return context_str

    def find_matched_terms(self, text, language):
        """为指定语言查找匹配的术语"""
        if language not in self.term_base_dict or not self.term_base_dict[language]:
            return {}

        words = self.tokenize_chinese_text(text)
        matched_terms = {}

        # 词级别匹配
        for word in words:
            for term_entry in self.term_base_dict[language]:
                if term_entry['source'] == word:
                    if word not in matched_terms:
                        matched_terms[word] = []
                    if term_entry['target'] not in matched_terms[word]:
                        matched_terms[word].append(term_entry['target'])

        # 短语级别匹配
        for term_entry in self.term_base_dict[language]:
            term = term_entry['source']
            if term in text:
                if term not in matched_terms:
                    matched_terms[term] = []
                if term_entry['target'] not in matched_terms[term]:
                    matched_terms[term].append(term_entry['target'])

        return matched_terms

    def build_term_base_prompt(self, text, language="英文"):
        """为指定语言构建术语库提示"""
        # 尝试使用多语言术语库
        matched_terms = self.find_matched_terms(text, language)

        # 如果没有多语言术语库，使用单语言术语库
        if not matched_terms and self.term_base_list:
            words = self.tokenize_chinese_text(text)
            matched_terms = {}

            for word in words:
                for term_entry in self.term_base_list:
                    if term_entry['source'] == word:
                        if word not in matched_terms:
                            matched_terms[word] = []
                        if term_entry['target'] not in matched_terms[word]:
                            matched_terms[word].append(term_entry['target'])

            for term_entry in self.term_base_list:
                term = term_entry['source']
                if term in text:
                    if term not in matched_terms:
                        matched_terms[term] = []
                    if term_entry['target'] not in matched_terms[term]:
                        matched_terms[term].append(term_entry['target'])

        if not matched_terms:
            return ""

        term_base_str = f"\n\n### 术语库匹配：\n"

        for orig, trans_list in matched_terms.items():
            if len(trans_list) == 1:
                term_base_str += f"- 「{orig}」 → {language}译名：「{trans_list[0]}」\n"
            else:
                term_base_str += f"- 「{orig}」 → {language}译名候选：{' / '.join([f'「{t}」' for t in trans_list])} （根据上下文选择最合适的）\n"

        return term_base_str

    def build_role_personality_prompt(self, role_name):
        if not role_name:
            return ""

        personality = self.find_role_personality(role_name)
        self.current_role_personality = personality

        if not personality:
            return ""

        mapped_role = self.role_mapping.get(str(role_name).strip())
        if mapped_role and mapped_role != str(role_name).strip():
            role_personality_str = f"\n\n### 角色性格描述：\n角色「{role_name}」(映射为「{mapped_role}」)的性格特点：{personality}\n"
        else:
            role_personality_str = f"\n\n### 角色性格描述：\n角色「{role_name}」的性格特点：{personality}\n"

        return role_personality_str

    def set_target_languages(self, languages, column_names):
        """设置目标语言列表和对应的列名"""
        self.target_languages = languages
        self.language_column_names = column_names
        # 为每种语言初始化独立的上下文历史
        for lang in languages:
            if lang not in self.context_history:
                self.context_history[lang] = []

    def set_target_language(self, language):
        self.target_language = language

    def get_language_specific_requirements(self, language):
        language_requirements = {
            "英文": """
英文翻译要求：
- 保持自然流畅，符合英语母语者的表达习惯
- 游戏UI文本要简洁明了，避免冗长
- 角色对话要符合人物性格，使用恰当的语气
- 专有名词和术语要保持一致性
- 文化特定表达要进行适当的本地化处理
""",
            "日文": """
日文翻译要求：
- 注意敬体和常体的使用，根据角色关系和场景选择合适的语体
- 游戏UI文本要简洁明了，符合日语表达习惯
- 角色对话要符合人物性格，使用恰当的语气和语尾
- 专有名词和术语要保持一致性
- 文化特定表达要进行适当的本地化处理
""",
            "韩文": """
韩文翻译要求：
- 注意尊敬语和非尊敬语的使用，根据角色关系和场景选择合适的语体
- 游戏UI文本要简洁明了，符合韩语表达习惯
- 角色对话要符合人物性格，使用恰当的语气
- 专有名词和术语要保持一致性
- 文化特定表达要进行适当的本地化处理
"""
        }

        if language in language_requirements:
            return language_requirements[language]
        else:
            return f"""
{language}翻译要求：
- 注意正式和非正式语体的使用，根据角色关系和场景选择合适的语体
- 游戏UI文本要简洁明了，符合{language}表达习惯
- 角色对话要符合人物性格，使用恰当的语气
- 专有名词和术语要保持一致性
- 文化特定表达要进行适当的本地化处理
"""

    def is_translation_error(self, response_text, original_text):
        if not response_text or response_text.strip() == "":
            return True
        if len(response_text) < len(original_text) * 0.1:
            return True
        return False

    def translate_text_with_retry(self, text, target_language, custom_requirements="", role=None):
        if not text or pd.isna(text) or str(text).strip() == "":
            return text

        last_exception = None

        for attempt in range(self.max_retries):
            try:
                translated_text = self._translate_single_attempt(text, target_language, custom_requirements, role)

                if not self.is_translation_error(translated_text, text):
                    return translated_text
                else:
                    st.warning(f"⚠️ [{target_language}] 第 {attempt + 1} 次翻译结果异常，准备重试...")

            except requests.exceptions.RequestException as e:
                last_exception = e
                st.warning(f"⚠️ [{target_language}] 网络错误 (第 {attempt + 1} 次尝试): {e}")

            except requests.exceptions.Timeout as e:
                last_exception = e
                st.warning(f"⚠️ [{target_language}] 请求超时 (第 {attempt + 1} 次尝试): {e}")

            except requests.exceptions.ConnectionError as e:
                last_exception = e
                st.warning(f"⚠️ [{target_language}] 连接错误 (第 {attempt + 1} 次尝试): {e}")

            except Exception as e:
                last_exception = e
                st.warning(f"⚠️ [{target_language}] API错误 (第 {attempt + 1} 次尝试): {e}")

            if attempt < self.max_retries - 1:
                wait_time = min(2 ** attempt, 60)
                time.sleep(wait_time)

        st.error(f"❌ [{target_language}] 翻译失败，已达到最大重试次数 {self.max_retries}")
        if last_exception:
            st.error(f"最后错误: {last_exception}")

        return text

    def _translate_single_attempt(self, text, target_language, custom_requirements="", role=None):
        # 为当前语言构建独立的上下文和术语提示
        context_prompt = self.build_context_prompt(target_language)
        term_base_prompt = self.build_term_base_prompt(text, target_language)
        role_personality_prompt = self.build_role_personality_prompt(role) if role else ""
        language_requirements = self.get_language_specific_requirements(target_language)

        role_prompt = ""
        if role and not pd.isna(role) and str(role).strip() != "":
            role_prompt = f"\n当前文本的说话人: {role}\n"

        prompt = f"""
请将以下文本翻译成{target_language}。

## 角色信息：
{role_prompt}{role_personality_prompt}

## {target_language}翻译规范（优先级最低）：
{language_requirements}

## 用户自定义要求（优先级第一高）：
{custom_requirements}

{context_prompt}

{term_base_prompt}

## 待翻译文本：
{text}

## 重要说明（优先级第二高）：
1. 请只返回{target_language}翻译结果，不要添加任何解释或备注
2. 术语库中的特定词汇翻译，如果是人名或者固定特殊名称需要严格采用相同的翻译,但注意如果是一些普通的词汇则看句子翻译不必一定按照术语库来
3. 如果一个术语有多个候选译名，请根据上下文选择最合适的
4. 请根据角色性格描述调整翻译风格和语气
5. 参考上下文中的{target_language}译文风格，保持翻译一致性，但是有的时候上下文角色可能不是一个人或者只是UI翻译，你还是需要参考原文判断是否参考上下文

{target_language}翻译结果：
"""

        data = {
            "model": self.model,
            "messages": [
                {
                    "role": "system",
                    "content": f"你是一名专业的{target_language}翻译专家，擅长游戏本地化、UI界面翻译和角色文案翻译。你正在进行中文到{target_language}的翻译工作。请确保术语一致性和风格统一，并根据角色特点调整翻译风格。"
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            "temperature": 0.2,
            "max_tokens": 4000
        }
        print(str(prompt))
        response = requests.post(self.api_url, headers=self.headers, json=data, timeout=60)
        response.raise_for_status()
        result = response.json()
        translated_text = result["choices"][0]["message"]["content"].strip()

        translated_text = self.clean_translation(translated_text)
        # 将翻译结果添加到该语言的独立上下文中
        self.add_to_context(text, translated_text, role, target_language)

        return translated_text

    def translate_text(self, text, target_language, custom_requirements="", role=None):
        return self.translate_text_with_retry(text, target_language, custom_requirements, role)

    def clean_translation(self, text):
        if text.startswith('"') and text.endswith('"'):
            text = text[1:-1]
        elif text.startswith("'") and text.endswith("'"):
            text = text[1:-1]

        prefixes = [
            "monior：", "monior:", "角色：", "角色:",
            "翻译：", "翻译:", "译文：", "译文:",
            "结果：", "结果:", "翻译结果：", "翻译结果:",
            "英文翻译结果：", "日文翻译结果：", "韩文翻译结果：",
            "英文：", "日文：", "韩文：", "法文：", "德文：",
        ]
        for prefix in prefixes:
            if text.startswith(prefix):
                text = text[len(prefix):].strip()

        return text

    def reset_context(self):
        self.context_history = {}
        self.term_dict = {}
        self.term_base_dict = {}
        self.role_column = None
        self.current_text_terms = {}
        self.current_role_personality = None
        self.role_mapping = {}

    def load_term_base(self, df, source_col, target_col):
        """加载术语库 - 支持重复术语"""
        try:
            # 构建术语库列表，支持重复术语
            self.term_base_list = []
            missing_count = 0

            for _, row in df.iterrows():
                source = row[source_col]
                target = row[target_col]

                if pd.isna(source) or pd.isna(target):
                    missing_count += 1
                    continue

                source = str(source).strip()
                target = str(target).strip()

                if source and target:
                    # 不再检查重复，直接添加到列表
                    self.term_base_list.append({
                        'source': source,
                        'target': target
                    })

            st.success(f"✅ 成功加载术语: {len(self.term_base_list)} 条")
            if missing_count > 0:
                st.warning(f"⚠️ 跳过 {missing_count} 条不完整的记录")

            return True
        except Exception as e:
            st.error(f"❌❌ 加载术语库失败: {e}")
            return False

    def load_term_base_multilang(self, df, source_col, target_cols_dict):
        """
        加载多语言术语库
        df: 术语库DataFrame
        source_col: 原文列名
        target_cols_dict: {语言: 列名} 例如 {"英文": "English", "日文": "Japanese"}
        """
        try:
            self.term_base_dict = {}

            for language, target_col in target_cols_dict.items():
                if target_col not in df.columns:
                    st.warning(f"⚠️ 术语库中未找到 {language} 对应的列: {target_col}")
                    continue

                self.term_base_dict[language] = []
                missing_count = 0

                for _, row in df.iterrows():
                    source = row[source_col]
                    target = row[target_col]

                    if pd.isna(source) or pd.isna(target):
                        missing_count += 1
                        continue

                    source = str(source).strip()
                    target = str(target).strip()

                    if source and target:
                        # 将原文添加到分词词典
                        try:
                            self.chinese_tokenizer.add_word(source)
                        except:
                            pass

                        self.term_base_dict[language].append({
                            'source': source,
                            'target': target
                        })

                st.success(f"✅ {language} 术语加载成功: {len(self.term_base_dict[language])} 条")
                if missing_count > 0:
                    st.info(f"   跳过 {missing_count} 条不完整的 {language} 术语")

            # 显示术语库统计
            total_terms = sum(len(terms) for terms in self.term_base_dict.values())
            st.success(f"📊 总计加载术语: {total_terms} 条，覆盖 {len(self.term_base_dict)} 种语言")

            # 显示术语示例
            with st.expander("📋 查看术语库示例"):
                for language, terms in self.term_base_dict.items():
                    if terms:
                        st.write(f"**{language}术语示例：**")
                        for i, term in enumerate(terms[:5]):
                            st.write(f"  {i+1}. {term['source']} → {term['target']}")
                        if len(terms) > 5:
                            st.write(f"  ... 还有 {len(terms)-5} 条")

            return True

        except Exception as e:
            st.error(f"❌ 加载多语言术语库失败: {e}")
            import traceback
            st.error(traceback.format_exc())
            return False

    def load_role_personality(self, df, role_col, personality_col):
        try:
            self.role_personality_dict = {}
            missing_count = 0

            for _, row in df.iterrows():
                role = row[role_col]
                personality = row[personality_col]

                if pd.isna(role) or pd.isna(personality):
                    missing_count += 1
                    continue

                role = str(role).strip()
                personality = str(personality).strip()

                if role and personality:
                    self.role_personality_dict[role] = personality

            st.success(f"✅ 成功加载角色性格: {len(self.role_personality_dict)} 条")
            if missing_count > 0:
                st.warning(f"⚠️ 跳过 {missing_count} 条不完整的记录")

            self.analyze_role_personality()

            return True
        except Exception as e:
            st.error(f"❌ 加载角色性格库失败: {e}")
            return False

    def analyze_role_personality(self):
        if not self.role_personality_dict:
            return

        st.write(f"📊 角色性格库统计: {len(self.role_personality_dict)} 个角色")

        st.write("📋 部分角色性格预览:")
        count = 0
        for role, personality in list(self.role_personality_dict.items())[:5]:
            st.write(f"  - {role}: {personality[:50]}..." if len(personality) > 50 else f"  - {role}: {personality}")
            count += 1
            if count >= 5:
                break
