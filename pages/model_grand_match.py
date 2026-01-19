import streamlit as st
import pandas as pd
import re
import jieba
import time
import json
import os
import random
from openai import OpenAI
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

# ==========================================
# 0. 配置管理系统
# ==========================================

CONFIG_FILE = "l10n_config_v13.json"

def load_config():
    default_config = {
        "api_base": "https://api.deepseek.com",
        "api_key": "",
        "model_name": "deepseek-chat",
        "max_threads": 5,
        "max_retries": 3,
        "timeout": 30,
        "custom_instruction": "",
        "min_group_size": 3,
        "col_src_name": "Chinese_PRC",
        "col_tgt1_name": "English",
        "col_tgt2_name": "Japanese",
        "col_main_text": ""
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                default_config.update(saved)
        except: pass
    return default_config

def save_config(data):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        st.sidebar.success("✅ 配置已保存！")
    except Exception as e:
        st.sidebar.error(f"保存失败: {e}")

# 加载配置到全局
APP_CONFIG = load_config()

# ==========================================
# 1. 核心逻辑函数 (Helper Functions)
# ==========================================

def abstract_sentence_with_index(text):
    if not isinstance(text, str): return str(text), [], []
    variables, nums = [], []
    var_counter, num_counter = 0, 0
    
    def replace_var(match):
        nonlocal var_counter
        var_counter += 1
        variables.append({"type": "var", "index": var_counter, "content": match.group(2)})
        return f"{{VAR{var_counter}}}"
    
    text_masked = re.sub(r'([“"【\[](.*?)[”"】\]])', replace_var, text)
    
    def replace_num(match):
        nonlocal num_counter
        num_counter += 1
        nums.append(match.group(1))
        return f"{{NUM{num_counter}}}"
    
    skeleton = re.sub(r'(\d+)', replace_num, text_masked)
    return skeleton, [v['content'] for v in variables], nums

def build_glossary_dict(df, src_col, tgt1_col, tgt2_col):
    lookup = {}
    if df is None: return lookup
    df = df.fillna("")
    for _, row in df.iterrows():
        key = str(row[src_col]).strip()
        if not key: continue
        t1, t2 = str(row[tgt1_col]).strip(), str(row[tgt2_col]).strip()
        if key not in lookup:
            lookup[key] = {"t1": t1, "t2": t2}
        else:
            if not lookup[key]["t1"] and t1: lookup[key]["t1"] = t1
            if not lookup[key]["t2"] and t2: lookup[key]["t2"] = t2
    return lookup

def generate_hints_list(text, lookup_dict):
    if not text: return []
    clean_text = re.sub(r'\{(VAR|NUM)\d+\}', ' ', str(text))
    words = jieba.lcut(clean_text)
    found = []
    seen = set()
    for w in words:
        w = w.strip()
        if w in lookup_dict and w not in seen:
            res = lookup_dict[w]
            if res['t1'] or res['t2']:
                found.append({"term": w, "en": res['t1'], "jp": res['t2']})
                seen.add(w)
    return found

# ==========================================
# 2. AI 处理逻辑
# ==========================================

print_lock = threading.Lock()

def call_ai_api_with_retry(client, model, prompt, sys_prompt, max_retries=3, timeout=30):
    with print_lock:
        print(f"\n🚀 [SENDING] >>>\n{prompt[:150]}...")
    delay = 1
    for attempt in range(max_retries + 1):
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[{"role": "system", "content": sys_prompt}, {"role": "user", "content": prompt}],
                temperature=0.1,
                timeout=timeout 
            )
            result = resp.choices[0].message.content.strip().strip('"')
            with print_lock:
                print(f"✅ [SUCCESS] <<< {result}")
            return result
        except Exception as e:
            if attempt == max_retries:
                with print_lock:
                    print(f"❌ [FAILED] Final give up. Error: {e}")
                return None
            with print_lock:
                print(f"⚠️ [RETRY] ({attempt+1}/{max_retries}) Error: {e}. Waiting {delay}s...")
            time.sleep(delay)
            delay *= 2
            delay += random.uniform(0, 1)

def translate_skeleton_task(client, model, skeleton, hints, target_lang, user_instruction, max_retries, timeout):
    glossary_text = ""
    if hints:
        glossary_text = "Refer to this glossary (Strictly):\n"
        for h in hints:
            val = h['en'] if target_lang == "English" else h['jp']
            if val: glossary_text += f"- {h['term']} -> {val}\n"
    sys_prompt = "You are a professional game localization expert."
    extra = f"\nAdditional Instructions:\n{user_instruction}\n" if user_instruction else ""
    user_prompt = f"""
Task: Translate the UI pattern to {target_lang}.
Source: "{skeleton}"
Constraints:
1. Keep {{VARx}}, {{NUMx}} intact.
2. Reorder tags if needed.
3. Output ONLY translation.
{glossary_text}
{extra}
"""
    res = call_ai_api_with_retry(client, model, user_prompt, sys_prompt, max_retries, timeout)
    if res: return res
    return skeleton

def translate_variable_task(client, model, var_text, hints, target_lang, user_instruction, max_retries, timeout):
    glossary_text = ""
    if hints:
        glossary_text = "Glossary Hints:\n"
        for h in hints:
            val = h['en'] if target_lang == "English" else h['jp']
            if val: glossary_text += f"- {h['term']} -> {val}\n"
    sys_prompt = "You are a game translator. Translate the specific term."
    extra = f"\nInstructions: {user_instruction}" if user_instruction else ""
    user_prompt = f"""
Translate this Game Term to {target_lang}.
Term: "{var_text}"
Context: Inside a UI sentence.
Constraints: concise, accurate, no extra explanation.
{glossary_text}
{extra}
"""
    res = call_ai_api_with_retry(client, model, user_prompt, sys_prompt, max_retries, timeout)
    if res: return res
    return var_text

def batch_process_scope(do_skeletons, do_vars, skeletons, all_vars, glossary_lookup, api_key, base_url, model, workers, instruction, max_retries, timeout):
    clean_url = base_url.rstrip("/").replace("/chat/completions", "").replace("/v1", "")
    if not clean_url.endswith("/v1"): clean_url += "/v1"
    
    try: client = OpenAI(api_key=api_key, base_url=clean_url)
    except: return {}, {}

    skel_results = {} 
    var_results = {}  
    vars_to_translate = []
    if do_vars:
        for v in all_vars:
            if v not in glossary_lookup: vars_to_translate.append(v)
            else: var_results[v] = {'t1': glossary_lookup[v]['t1'], 't2': glossary_lookup[v]['t2']}
    
    task_count_skel = len(skeletons) * 2 if do_skeletons else 0
    task_count_var = len(vars_to_translate) * 2 if do_vars else 0
    total_tasks = task_count_skel + task_count_var
    if total_tasks == 0: return {}, {}

    prog_bar = st.progress(0)
    status = st.empty()
    completed = 0
    print(f"🚀 [START] Tasks: {task_count_skel/2} Skels, {task_count_var/2} Vars. Threads: {workers}")

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {}
        if do_skeletons:
            for sk in skeletons:
                hints = generate_hints_list(sk, glossary_lookup)
                f1 = executor.submit(translate_skeleton_task, client, model, sk, hints, "English", instruction, max_retries, timeout)
                futures[f1] = ("skel", sk, "t1")
                f2 = executor.submit(translate_skeleton_task, client, model, sk, hints, "Japanese", instruction, max_retries, timeout)
                futures[f2] = ("skel", sk, "t2")
        if do_vars:
            for v in vars_to_translate:
                hints = generate_hints_list(v, glossary_lookup)
                f3 = executor.submit(translate_variable_task, client, model, v, hints, "English", instruction, max_retries, timeout)
                futures[f3] = ("var", v, "t1")
                f4 = executor.submit(translate_variable_task, client, model, v, hints, "Japanese", instruction, max_retries, timeout)
                futures[f4] = ("var", v, "t2")
        
        for f in as_completed(futures):
            ctype, key, lang = futures[f]
            try:
                res = f.result()
                if ctype == "skel":
                    if key not in skel_results: skel_results[key] = {}
                    skel_results[key][lang] = res
                else:
                    if key not in var_results: var_results[key] = {}
                    var_results[key][lang] = res
            except Exception as e: print(f"❌ Thread Error: {e}")
            completed += 1
            prog_bar.progress(completed / total_tasks)
            status.text(f"Processing... {completed}/{total_tasks}")
    status.text("✅ 处理完成！")
    return skel_results, var_results

def get_index(opt, val, default=0):
    try: return opt.index(val)
    except: return default

# ==========================================
# 3. 主函数 (UI 逻辑封装在这里)
# ==========================================

def grand_match():
    st.title("🎮 L10n AI 全局同步版 V13")

    # Session State 初始化
    if 'user_inputs' not in st.session_state: st.session_state['user_inputs'] = {}
    if 'selection_state' not in st.session_state: st.session_state['selection_state'] = pd.DataFrame()
    if 'ai_var_cache' not in st.session_state: st.session_state['ai_var_cache'] = {} 
    
    # --- Sidebar ---
    with st.sidebar:
        st.header("1. 基础配置")
        source_file = st.file_uploader("源文 Excel", type=["xlsx", "xls"], key="src_uploader")
        glossary_file = st.file_uploader("术语表 Excel", type=["xlsx", "xls"], key="glossary_uploader")
        
        glossary_lookup = {}
        g_src, g_tgt1, g_tgt2 = None, None, None

        if glossary_file:
            df_g = pd.read_excel(glossary_file)
            g_cols = df_g.columns.tolist()
            g_src = st.selectbox("原文列", g_cols, index=get_index(g_cols, APP_CONFIG["col_src_name"], 0))
            g_tgt1 = st.selectbox("译文1 (En)", g_cols, index=get_index(g_cols, APP_CONFIG["col_tgt1_name"], 1))
            g_tgt2 = st.selectbox("译文2 (Jp)", g_cols, index=get_index(g_cols, APP_CONFIG["col_tgt2_name"], 2))
            glossary_lookup = build_glossary_dict(df_g, g_src, g_tgt1, g_tgt2)
            st.success(f"已加载 {len(glossary_lookup)} 条术语")

        st.divider()
        st.header("2. AI & 网络配置")
        api_base = st.text_input("Base URL", value=APP_CONFIG["api_base"])
        api_key = st.text_input("API Key", value=APP_CONFIG["api_key"], type="password")
        model_name = st.text_input("模型名称", value=APP_CONFIG["model_name"])
        
        c1, c2, c3 = st.columns(3)
        with c1: max_retries = st.number_input("重试", value=APP_CONFIG["max_retries"], min_value=0)
        with c2: max_threads = st.number_input("并发", value=APP_CONFIG["max_threads"], min_value=1)
        with c3: timeout_sec = st.number_input("超时(s)", value=APP_CONFIG["timeout"], min_value=5)

        st.divider()
        st.header("3. 翻译控制")
        custom_inst = st.text_area("提示词", value=APP_CONFIG["custom_instruction"], height=80)
        min_group = st.number_input("最小阈值", value=APP_CONFIG["min_group_size"])
        
        if st.button("💾 保存配置"):
            new_cfg = APP_CONFIG.copy()
            new_cfg.update({
                "api_base": api_base, "api_key": api_key, "model_name": model_name,
                "max_threads": max_threads, "max_retries": max_retries, "timeout": timeout_sec,
                "custom_instruction": custom_inst, "min_group_size": min_group,
                "col_src_name": g_src if glossary_file else "",
                "col_tgt1_name": g_tgt1 if glossary_file else "",
                "col_tgt2_name": g_tgt2 if glossary_file else ""
            })
            if 'current_col' in st.session_state: new_cfg['col_main_text'] = st.session_state['current_col']
            save_config(new_cfg)

    # --- Main Interface ---
    if source_file:
        if 'raw_df_v13' not in st.session_state or st.sidebar.button("重新加载文件"):
            st.session_state['raw_df_v13'] = pd.read_excel(source_file)
            st.session_state['processed_v13'] = False
            st.session_state['user_inputs'] = {}
            st.session_state['ai_var_cache'] = {}
        
        df = st.session_state['raw_df_v13']
        
        # --- Step 1 Analysis ---
        if not st.session_state.get('processed_v13'):
            cols = df.columns.tolist()
            col_text = st.selectbox("选择文本列", cols, index=get_index(cols, APP_CONFIG["col_main_text"], 0))
            st.session_state['current_col'] = col_text
            if st.button("开始分析"):
                with st.spinner("Analyzing..."):
                    res = df[col_text].apply(abstract_sentence_with_index)
                    df['__Skeleton__'] = [r[0] for r in res]
                    df['__Vars__'] = [r[1] for r in res]
                    df['__Nums__'] = [r[2] for r in res]
                    counts = df['__Skeleton__'].value_counts()
                    valid = counts[counts >= min_group].index.tolist()
                    st.session_state['processed_df_v13'] = df
                    st.session_state['valid_skels_v13'] = valid
                    st.session_state['processed_v13'] = True
                    init_data = [{"应用": True, "行数": counts[s], "状态": "待翻译", "句型骨架": s} for s in valid]
                    st.session_state['selection_state'] = pd.DataFrame(init_data)
                    st.rerun()

        # --- Step 2 Translation & Review ---
        if st.session_state.get('processed_v13'):
            df_proc = st.session_state['processed_df_v13']
            skeletons = st.session_state['valid_skels_v13']
            
            st.divider()
            st.header("3. AI 翻译控制台")
            col_chk1, col_chk2, col_btn = st.columns([1, 1, 2])
            with col_chk1: do_skeletons = st.checkbox("翻译骨架", value=True)
            with col_chk2: do_vars = st.checkbox("翻译变量", value=True)
            with col_btn:
                if st.button("⚡ 开始 AI 翻译", type="primary"):
                    if not api_key: st.error("请配置 API Key")
                    elif not (do_skeletons or do_vars): st.warning("请至少勾选一项")
                    else:
                        all_unique_vars = set()
                        if do_vars:
                            for sk in skeletons:
                                sub = df_proc[df_proc['__Skeleton__'] == sk]
                                for vl in sub['__Vars__']: all_unique_vars.update(vl)
                        skel_res, var_res = batch_process_scope(do_skeletons, do_vars, skeletons, list(all_unique_vars), glossary_lookup, api_key, api_base, model_name, max_threads, custom_inst, max_retries, timeout_sec)
                        
                        if do_skeletons:
                            for sk, res in skel_res.items():
                                if sk not in st.session_state['user_inputs']: st.session_state['user_inputs'][sk] = {'t1_tmpl': "", 't2_tmpl': "", 'var_map': pd.DataFrame()}
                                st.session_state['user_inputs'][sk]['t1_tmpl'] = res.get('t1', '')
                                st.session_state['user_inputs'][sk]['t2_tmpl'] = res.get('t2', '')
                                st.session_state[f"t1_{sk}"] = res.get('t1', '')
                                st.session_state[f"t2_{sk}"] = res.get('t2', '')
                        if do_vars: st.session_state['ai_var_cache'].update(var_res)
                        st.session_state['selection_state']['状态'] = "🤖 部分已填"
                        st.success("任务完成！")
                        time.sleep(1)
                        st.rerun()

            # --- Review Interface ---
            st.divider()
            st.subheader("4. 审查与微调")
            sel_sk = st.selectbox("选择句型", skeletons, format_func=lambda x: f"[{len(df_proc[df_proc['__Skeleton__']==x])}] {x}")
            curr = st.session_state['user_inputs'].get(sel_sk, {})
            v_t1 = curr.get('t1_tmpl', sel_sk)
            v_t2 = curr.get('t2_tmpl', sel_sk)
            
            sk_hints = generate_hints_list(sel_sk, glossary_lookup)
            if sk_hints:
                h_str = " | ".join([f"{h['term']}:{h['en']}/{h['jp']}" for h in sk_hints])
                st.caption(f"💡 骨架术语: {h_str}")

            c1, c2 = st.columns(2)
            with c1: new_t1 = st.text_input("Target 1 (En)", value=v_t1, key=f"t1_{sel_sk}")
            with c2: new_t2 = st.text_input("Target 2 (Jp)", value=v_t2, key=f"t2_{sel_sk}")
            
            # --- 变量编辑器 ---
            sub = df_proc[df_proc['__Skeleton__'] == sel_sk]
            grp_vars = sorted(list(set([v for vl in sub['__Vars__'] for v in vl])))
            
            if grp_vars:
                rows = []
                for v in grp_vars:
                    def_t1, def_t2 = "", ""
                    if v in st.session_state['ai_var_cache']:
                        cached = st.session_state['ai_var_cache'][v]
                        def_t1, def_t2 = cached.get('t1', ''), cached.get('t2', '')
                    elif v in glossary_lookup:
                         def_t1, def_t2 = glossary_lookup[v]['t1'], glossary_lookup[v]['t2']
                    rows.append({"原文变量": v, "Target 1": def_t1, "Target 2": def_t2})
                
                df_vars_display = pd.DataFrame(rows)
                st.caption("变量翻译:")
                edited = st.data_editor(df_vars_display, key=f"ed_{sel_sk}", hide_index=True, use_container_width=True)
                
                st.session_state['user_inputs'][sel_sk] = {
                    't1_tmpl': new_t1, 't2_tmpl': new_t2, 'var_map': edited
                }
                
                # 全局同步按钮
                if st.button("🌍 全局应用变量翻译 (同步给所有句型)"):
                    current_map = edited.set_index("原文变量")
                    updated_count = 0
                    for v_txt, row in current_map.iterrows():
                        t1, t2 = row['Target 1'], row['Target 2']
                        if t1 or t2: st.session_state['ai_var_cache'][v_txt] = {'t1': t1, 't2': t2}
                    for k, data in st.session_state['user_inputs'].items():
                        if k == sel_sk: continue 
                        target_df = data.get('var_map')
                        if target_df is not None and not target_df.empty:
                            modified = False
                            for idx, row in target_df.iterrows():
                                vt = row['原文变量']
                                if vt in current_map.index:
                                    new_v1 = current_map.loc[vt, 'Target 1']
                                    new_v2 = current_map.loc[vt, 'Target 2']
                                    if (new_v1 or new_v2) and (target_df.at[idx, 'Target 1'] != new_v1 or target_df.at[idx, 'Target 2'] != new_v2):
                                        target_df.at[idx, 'Target 1'] = new_v1
                                        target_df.at[idx, 'Target 2'] = new_v2
                                        modified = True
                            if modified:
                                st.session_state['user_inputs'][k]['var_map'] = target_df
                                updated_count += 1
                    st.success(f"✅ 同步完成！已更新 {updated_count} 个句型。")

            # --- Generate ---
            st.divider()
            st.subheader("5. 导出")
            final_sel = st.data_editor(st.session_state['selection_state'], column_config={"应用":st.column_config.CheckboxColumn(default=True), "状态":st.column_config.TextColumn(disabled=True), "句型骨架":st.column_config.TextColumn(disabled=True, width="large")}, disabled=["状态","句型骨架","行数"], hide_index=True, use_container_width=True, height=200)
            st.session_state['selection_state'] = final_sel
            
            if st.button("🚀 生成最终 Excel", type="primary"):
                app_rows = final_sel[final_sel['应用']==True]
                active = set(app_rows['句型骨架'].tolist())
                res_df = df_proc.copy()
                res_df['Target1_Result'] = ""
                res_df['Target2_Result'] = ""
                cfg = {}
                for sk, val in st.session_state['user_inputs'].items():
                    vm = {}
                    if 'var_map' in val and not val['var_map'].empty:
                        for _, r in val['var_map'].iterrows(): vm[r['原文变量']] = {'t1': r['Target 1'], 't2': r['Target 2']}
                    cfg[sk] = {'t1': val['t1_tmpl'], 't2': val['t2_tmpl'], 'vars': vm}
                
                pb = st.progress(0)
                for i, row in res_df.iterrows():
                    sk = row['__Skeleton__']
                    if sk in active and sk in cfg:
                        c = cfg[sk]
                        vs, ns = row['__Vars__'], row['__Nums__']
                        r1 = c['t1']
                        for k, n in enumerate(ns): r1 = r1.replace(f"{{NUM{k+1}}}", str(n))
                        for k, v in enumerate(vs):
                            vt = c['vars'].get(v, {}).get('t1', v) 
                            r1 = r1.replace(f"{{VAR{k+1}}}", str(vt))
                        res_df.at[i, 'Target1_Result'] = r1
                        r2 = c['t2']
                        for k, n in enumerate(ns): r2 = r2.replace(f"{{NUM{k+1}}}", str(n))
                        for k, v in enumerate(vs):
                            vt = c['vars'].get(v, {}).get('t2', v)
                            r2 = r2.replace(f"{{VAR{k+1}}}", str(vt))
                        res_df.at[i, 'Target2_Result'] = r2
                    if i%100==0: pb.progress(i/len(res_df))
                pb.progress(1.0)
                csv = res_df[[c for c in res_df.columns if not c.startswith('__')]].to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 下载结果", csv, "Localized_V13.csv", "text/csv")
