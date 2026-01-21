# pages/profile_analysis.py - 文献综述分析工具
import streamlit as st
import os
import glob
import json
import requests
from docx import Document
import openpyxl
import tempfile
import shutil
from pathlib import Path
import fitz  # PyMuPDF
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
from datetime import datetime
import pandas as pd
import io

# --- 1. 核心逻辑函数 (移植自原 123.py) ---

API_CONFIGS = {
    "DeepSeek": {
        "url": "https://api.deepseek.com/v1/chat/completions",
        "models": ["deepseek-chat", "deepseek-coder", "deepseek-reasoner"],
        "default_model": "deepseek-chat",
        "headers": {"Content-Type": "application/json"}
    },
    "OpenAI": {
        "url": "https://api.openai.com/v1/chat/completions",
        "models": ["gpt-4", "gpt-4-turbo", "gpt-3.5-turbo", "gpt-3.5-turbo-16k"],
        "default_model": "gpt-3.5-turbo",
        "headers": {"Content-Type": "application/json"}
    },
    "New API (DeepSeek-v3)": {
        "url": "https://tb.api.mkeai.com/v1/chat/completions",
        "models": ["deepseek-v3", "deepseek-v2", "deepseek-v1"],
        "default_model": "deepseek-v3",
        "headers": {"Content-Type": "application/json"}
    }
}

CONFIG_FILE = "api_config_profile.json" # 避免与主配置冲突

def load_config():
    """加载保存的配置"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            pass
    return {}

def save_config(api_provider, api_key, custom_url, model, custom_models, custom_prompt, max_workers, max_retries):
    """保存配置到文件"""
    try:
        config = {
            "api_provider": api_provider,
            "api_key": api_key,
            "custom_url": custom_url,
            "model": model,
            "custom_models": custom_models,
            "custom_prompt": custom_prompt,
            "max_workers": max_workers,
            "max_retries": max_retries,
            "last_used": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        return False

def read_word_file(file_path):
    """读取Word文档内容"""
    try:
        doc = Document(file_path)
        full_text = []
        for paragraph in doc.paragraphs:
            full_text.append(paragraph.text)
        return '\n'.join(full_text)
    except Exception as e:
        return f"读取Word文件时出错: {e}"

def read_pdf_file(file_path):
    """读取PDF文档内容 - 使用PyMuPDF"""
    try:
        doc = fitz.open(file_path)
        full_text = []
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            text = page.get_text()
            full_text.append(text)
        doc.close()
        return '\n'.join(full_text)
    except Exception as e:
        return f"读取PDF文件时出错: {e}"

def get_file_content(file_path):
    """根据文件类型读取内容"""
    file_ext = Path(file_path).suffix.lower()
    if file_ext == '.docx':
        return read_word_file(file_path)
    elif file_ext == '.pdf':
        return read_pdf_file(file_path)
    else:
        return f"不支持的文件类型: {file_ext}"

def call_ai_api(content, api_provider, api_key, custom_url=None, model=None, custom_prompt=""):
    """调用AI API分析文档内容"""
    if api_provider == "Custom" and custom_url:
        config = {
            "url": custom_url,
            "model": model or "gpt-3.5-turbo",
            "headers": {"Content-Type": "application/json"}
        }
    else:
        config = API_CONFIGS.get(api_provider, API_CONFIGS["DeepSeek"]).copy()
        if model:
            config["model"] = model
        else:
            config["model"] = config.get("default_model", config["models"][0])

    url = config["url"]
    model_name = config["model"]
    headers = config["headers"].copy()
    headers["Authorization"] = f"Bearer {api_key}"

    base_prompt = """请分析以下学术文档，严格按照以下格式返回JSON数据：
{
    "title": "文章题目",
    "authors": "作者信息", 
    "research_content": "研究内容（200-300字，详细描述研究的问题、方法、理论框架等）",
    "research_results": "研究结果（200-300字，详细描述主要发现、结论、贡献等）"
}"""

    if custom_prompt.strip():
        full_prompt = f"{base_prompt}\n\n额外要求（用户自定义）：{custom_prompt}\n\n文档内容："
    else:
        full_prompt = f"{base_prompt}\n\n文档内容："

    prompt = f"{full_prompt}\n{content[:12000]}"

    data = {
        "model": model_name,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3
    }

    response = requests.post(url, headers=headers, json=data, timeout=120)
    response.raise_for_status()
    result = response.json()
    return result["choices"][0]["message"]["content"]

def parse_api_response(response_text):
    """解析API返回的JSON数据"""
    try:
        start_idx = response_text.find('{')
        end_idx = response_text.rfind('}') + 1
        if start_idx != -1 and end_idx != 0:
            json_str = response_text[start_idx:end_idx]
            return json.loads(json_str)
        else:
            return {"error": "未找到有效的JSON数据"}
    except Exception as e:
        return {"error": f"解析API响应失败: {e}"}

def process_single_file(file_path, api_provider, api_key, custom_url, model, custom_prompt, max_retries):
    """处理单个文件"""
    filename = os.path.basename(file_path)
    retry_count = 0
    
    while retry_count <= max_retries:
        if st.session_state.get('stop_processing', False):
             return {
                "status": "stopped",
                "filename": filename,
                "error": "用户手动终止",
                "retry_count": retry_count
            }

        try:
            content = get_file_content(file_path)
            if content.startswith("读取文件时出错") or content.startswith("不支持的文件类型"):
                return {
                    "status": "failed",
                    "filename": filename,
                    "error": content,
                    "retry_count": retry_count
                }

            api_response = call_ai_api(content, api_provider, api_key, custom_url, model, custom_prompt)
            
            parsed_data = parse_api_response(api_response)
            if "error" in parsed_data:
                raise Exception(parsed_data["error"])
            
            parsed_data["filename"] = filename
            parsed_data["retry_count"] = retry_count
            return {
                "status": "success",
                "data": parsed_data,
                "filename": filename,
                "retry_count": retry_count
            }
            
        except Exception as e:
            retry_count += 1
            if retry_count <= max_retries:
                 if st.session_state.get('stop_processing', False):
                     return {
                        "status": "stopped",
                        "filename": filename,
                        "error": "用户手动终止",
                        "retry_count": retry_count - 1
                    }
                 time.sleep(2 ** retry_count) # Backoff
            else:
                return {
                    "status": "failed",
                    "filename": filename,
                    "error": str(e),
                    "retry_count": retry_count - 1
                }

# --- 2. 页面主函数 ---

def profile_analysis_page():
    st.title("📑 文献综述分析工具")
    st.caption("📚 支持多线程处理 | 实时预览 | 自动重试 | 结果导出")
    
    if "pa_logs" not in st.session_state:
        st.session_state.pa_logs = []
    if "pa_results" not in st.session_state:
        st.session_state.pa_results = []
    
    saved_config = load_config()
    
    # --- 左侧配置栏 ---
    with st.sidebar:
        st.header("1. API 配置")
        api_provider = st.selectbox(
            "API提供商",
            ["DeepSeek", "OpenAI", "New API (DeepSeek-v3)", "Custom"],
            index=["DeepSeek", "OpenAI", "New API (DeepSeek-v3)", "Custom"].index(saved_config.get("api_provider", "DeepSeek")),
            key="pa_api_provider"
        )
        
        api_key = st.text_input("API 密钥", type="password", value=saved_config.get("api_key", ""), key="pa_api_key")
        
        custom_url = ""
        custom_models = ""
        if api_provider == "Custom":
            custom_url = st.text_input("自定义API地址", value=saved_config.get("custom_url", ""), key="pa_custom_url")
            model_options = ["gpt-3.5-turbo"] # Default
            current_model = saved_config.get("model", "gpt-3.5-turbo")
            allow_custom_model = True
        else:
            config = API_CONFIGS.get(api_provider)
            model_options = config["models"]
            current_model = saved_config.get("model", config["default_model"])
            allow_custom_model = True 
            
        model = st.selectbox(
            "模型选择", 
            model_options + [current_model] if current_model not in model_options else model_options,
            index=model_options.index(current_model) if current_model in model_options else 0,
            key="pa_model"
        )
        if allow_custom_model and api_provider != "Custom": # Allow manual input override logic if needed, simplify for now
             pass 

        st.divider()
        st.header("2. 性能配置")
        max_workers = st.slider("并发线程数", 1, 10, saved_config.get("max_workers", 3), key="pa_workers")
        max_retries = st.number_input("最大重试次数", 0, 10, saved_config.get("max_retries", 3), key="pa_retries")
        
        st.divider()
        st.header("3. 自定义要求")
        custom_prompt = st.text_area("额外分析指令", value=saved_config.get("custom_prompt", ""), placeholder="例如：重点关注实验数据...", key="pa_prompt")
        
        if st.button("💾 保存配置"):
            save_config(api_provider, api_key, custom_url, model, "", custom_prompt, max_workers, max_retries)
            st.success("配置已保存")
            
    # --- 主操作区 ---
    
    st.subheader("📁 文件上传")
    input_type = st.radio("处理方式", ["单个文件上传", "文件夹批量处理"], horizontal=True, key="pa_input_type")
    
    files_to_process = []
    temp_dir = None
    
    if input_type == "单个文件上传":
        uploaded_files = st.file_uploader("选择文档 (PDF/Word)", type=["docx", "pdf"], accept_multiple_files=True, key="pa_uploader")
        if uploaded_files:
             # Create temp dir and save files
             temp_dir = tempfile.mkdtemp()
             for uploaded_file in uploaded_files:
                 path = os.path.join(temp_dir, uploaded_file.name)
                 with open(path, "wb") as f:
                     f.write(uploaded_file.getbuffer())
                 files_to_process.append(path)
                 
    else:
        folder_path = st.text_input("本地文件夹路径", key="pa_folder_path")
        if folder_path and os.path.exists(folder_path):
             word_files = glob.glob(os.path.join(folder_path, "*.docx"))
             pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))
             files_to_process = word_files + pdf_files
             if not files_to_process:
                 st.warning("该文件夹下没有找到支持的文档 (docx/pdf)")
        elif folder_path:
             st.error("文件夹不存在")
    
    col_start, col_stop = st.columns([1, 1])
    with col_start:
        start_btn = st.button("🚀 开始分析", type="primary", use_container_width=True, disabled=not files_to_process)
    with col_stop:
        stop_btn = st.button("⏸️ 停止", type="secondary", use_container_width=True)
        if stop_btn:
            st.session_state.stop_processing = True

    # --- 执行逻辑 ---
    if start_btn and files_to_process:
        if not api_key:
            st.error("请先配置 API Key")
            return
            
        st.session_state.stop_processing = False
        st.session_state.pa_results = []
        st.session_state.pa_logs = []
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        log_container = st.empty()
        
        total_files = len(files_to_process)
        completed = 0
        success = 0
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_file = {executor.submit(
                process_single_file, 
                f, api_provider, api_key, custom_url, model, custom_prompt, max_retries
            ): f for f in files_to_process}
            
            for future in concurrent.futures.as_completed(future_to_file):
                if st.session_state.get('stop_processing', False):
                    st.warning("⚠️ 处理已停止")
                    executor.shutdown(wait=False)
                    break
                
                f_path = future_to_file[future]
                f_name = os.path.basename(f_path)
                
                try:
                    result = future.result()
                    
                    if result["status"] == "success":
                        st.session_state.pa_results.append(result["data"])
                        success += 1
                        log_msg = f"✅ 成功: {f_name}"
                        if result["retry_count"] > 0: log_msg += f" (重试{result['retry_count']}次)"
                    elif result["status"] == "stopped":
                         log_msg = f"⏸️ 停止: {f_name}"
                    else:
                        log_msg = f"❌ 失败: {f_name} - {result['error']}"
                        
                    st.session_state.pa_logs.append(log_msg)
                    
                except Exception as e:
                     st.session_state.pa_logs.append(f"❌ 异常: {f_name} - {str(e)}")
                
                completed += 1
                progress_bar.progress(completed / total_files)
                status_text.text(f"进度: {completed}/{total_files} | 成功: {success}")
                log_container.code("\n".join(st.session_state.pa_logs[-10:])) # Show last 10 logs

        # Clean up temp dir if created
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            
        if completed == total_files:
            st.success("🎉 所有任务处理完成！")
            
    # --- 结果展示 & 下载 ---
    if st.session_state.pa_results:
        st.divider()
        st.subheader("📊 分析结果")
        
        df_results = pd.DataFrame(st.session_state.pa_results)
        
        # 显示预览 (只显示关键列)
        display_cols = ["filename", "title", "authors", "research_results"]
        # Ensure cols exist
        display_cols = [c for c in display_cols if c in df_results.columns]
        
        st.dataframe(df_results[display_cols], use_container_width=True)
        
        # Download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_results.to_excel(writer, index=False)
        
        st.download_button(
            "📥 下载完整 Excel 报告",
            data=output.getvalue(),
            file_name=f"literature_review_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
