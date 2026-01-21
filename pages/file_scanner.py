# pages/file_scanner.py - 文件扫描仪
import streamlit as st
import os
import pandas as pd
from openai import OpenAI
import io
import concurrent.futures

# --- 1. 核心工具函数 ---

CODE_EXTENSIONS = {
    '.py', '.js', '.txt','.jsx', '.ts', '.tsx', '.java', '.c', '.cpp', 
    '.h', '.cs', '.go', '.rs', '.php', '.rb', '.swift', '.kt', 
    '.html', '.css', '.sql', '.sh', '.bat', '.vue', '.lua', '.json', '.yaml', '.yml'
}

def get_all_code_files(root_path):
    """递归获取目录下所有代码文件路径"""
    code_files = []
    for root, dirs, files in os.walk(root_path):
        if any(ignore in root for ignore in ['.git', '__pycache__', 'node_modules', '.idea', '.vscode', 'venv', 'dist', 'build']):
            continue
        for file in files:
            _, ext = os.path.splitext(file)
            if ext.lower() in CODE_EXTENSIONS:
                full_path = os.path.join(root, file)
                code_files.append(full_path)
    return code_files

def analyze_code_with_llm(client_params, file_path, code_content):
    """并发分析单个文件"""
    base_url, api_key, model = client_params
    client = OpenAI(base_url=base_url, api_key=api_key)
    
    file_name = os.path.basename(file_path)
    # 截断防止 Token 溢出
    if len(code_content) > 15000: 
        code_content = code_content[:15000] + "\n...(代码过长已截断)..."

    prompt = f"""
    分析代码文件: {file_name}
    路径: {file_path}
    
    请输出简短的纯文本摘要（不要Markdown格式），包含：
    1. summary: 一句话概括文件作用。
    2. functions: 核心函数/类及其功能列表。
    
    代码内容:
    {code_content}
    """
    
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "你是一个代码审计专家。请用中文简练回答。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"分析出错: {str(e)}"

def build_context_from_df(df):
    """从 DataFrame 重建上下文知识库字符串"""
    context_str = "以下是项目中所有文件的分析摘要（基于历史扫描记录）：\n\n"
    # 确保列名存在，防止用户上传错误的 Excel
    required_cols = ['文件名', '分析详情']
    if not all(col in df.columns for col in required_cols):
        return None
        
    for index, row in df.iterrows():
        path_info = row['路径'] if '路径' in df.columns else "未知路径"
        context_str += f"=== 文件名: {row['文件名']} ===\n路径: {path_info}\n功能摘要: {row['分析详情']}\n\n"
    return context_str

# --- 2. 页面主函数 ---

def file_scanner_page():
    
    if "scanner_messages" not in st.session_state:
        st.session_state.scanner_messages = [] 
    if "project_context" not in st.session_state:
        st.session_state.project_context = "" 
    if "current_source" not in st.session_state:
        st.session_state.current_source = "未加载" # 记录当前数据来源
    
    st.title("🤖 AI 代码全能助手 (文件扫描仪)")
    
    # --- 配置区域 ---
    st.markdown("### 1. 配置")
    
    col1, col2 = st.columns(2)
    
    with col1:
        base_url = st.text_input("Base URL", value="https://api.openai.com/v1", key="scanner_base_url")
        api_key = st.text_input("API Key", type="password", key="scanner_api_key")
        model_name = st.text_input("Model Name", value="gpt-4o-mini", key="scanner_model")
        
    with col2:
        # 使用 Tabs 切换两种模式
        tab_scan, tab_load = st.tabs(["🚀 新建扫描", "📂 读取 Excel"])
        
        # --- 模式 A: 新建扫描 ---
        with tab_scan:
            st.caption("扫描本地文件夹生成新报告")
            target_folder = st.text_input("项目路径", placeholder="C:\\Projects\\MyCode", key="scanner_target_folder")
            max_workers = st.slider("并发线程", 1, 10, 5, key="scanner_workers")
            btn_scan = st.button("开始扫描", type="primary", key="btn_scan_start")

        # --- 模式 B: 读取 Excel ---
        with tab_load:
            st.caption("上传之前的分析报告 (.xlsx) 直接对话")
            uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"], key="scanner_upload")
            
        if st.button("🗑️ 清空对话历史", key="btn_clear_history"):
            st.session_state.scanner_messages = []
            st.rerun()
            
        st.info(f"当前状态: {st.session_state.current_source}")

    st.divider()

    # --- 逻辑处理 ---
    
    # 逻辑 A: 处理扫描
    if btn_scan:
        if not target_folder or not os.path.exists(target_folder) or not api_key:
            st.error("请检查路径和 API Key！")
        else:
            files = get_all_code_files(target_folder)
            if not files:
                st.warning("未找到代码文件。")
            else:
                progress_container = st.empty()
                progress_bar = progress_container.progress(0)
                status_text = st.empty()
                temp_results = []
                
                client_params = (base_url, api_key, model_name)
                
                # 并发执行
                with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                    future_to_file = {executor.submit(analyze_code_with_llm, client_params, f, open(f, 'r', encoding='utf-8', errors='ignore').read()): f for f in files}
                    
                    completed = 0
                    for future in concurrent.futures.as_completed(future_to_file):
                        file_path = future_to_file[future]
                        file_name = os.path.basename(file_path)
                        try:
                            res = future.result()
                        except:
                            res = "Error"
                        
                        temp_results.append({"文件名": file_name, "路径": file_path, "分析详情": res})
                        completed += 1
                        progress_bar.progress(completed / len(files), text=f"分析中: {file_name}")

                progress_container.empty()
                
                # 存入 DataFrame 并构建上下文
                df_res = pd.DataFrame(temp_results)
                st.session_state.project_context = build_context_from_df(df_res)
                st.session_state.scanner_messages = [] # 新项目清空历史
                st.session_state.current_source = f"新扫描 ({len(files)} 文件)"
                
                # 保存到 session 以便下载
                st.session_state.last_scan_df = df_res
                st.success("✅ 扫描完成！")

    # 逻辑 B: 处理 Excel 上传
    if uploaded_file is not None:
        # 只有当上传的文件变了，或者当前没有上下文时才处理
        if st.session_state.current_source != f"Excel: {uploaded_file.name}":
            try:
                df_load = pd.read_excel(uploaded_file)
                context = build_context_from_df(df_load)
                
                if context:
                    st.session_state.project_context = context
                    st.session_state.scanner_messages = [] # 加载新 Excel 清空历史
                    st.session_state.current_source = f"Excel: {uploaded_file.name}"
                    st.session_state.last_scan_df = df_load # 方便查看
                    st.success(f"📂 成功加载记录！包含了 {len(df_load)} 个文件的分析。")
                else:
                    st.error("Excel 格式不正确，缺少‘文件名’或‘分析详情’列。")
            except Exception as e:
                st.error(f"读取 Excel 失败: {e}")

    # --- 结果展示区 (折叠) ---
    if "last_scan_df" in st.session_state:
        with st.expander("📊 查看当前加载的数据详情 / 下载 Excel", expanded=False):
            st.dataframe(st.session_state.last_scan_df, use_container_width=True)
            
            # 提供下载（方便如果是扫描生成的，可以下载下来下次用）
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                st.session_state.last_scan_df.to_excel(writer, index=False)
            
            st.download_button("📥 下载此记录 (.xlsx)", buffer.getvalue(), "code_analysis.xlsx")

    # --- 对话区 (核心功能) ---
    
    if st.session_state.project_context:
        st.divider()
        st.subheader("💬 项目知识库对话")
        
        # 1. 回显历史记录
        for msg in st.session_state.scanner_messages:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        # 2. 处理新输入
        if prompt := st.chat_input("关于这个项目，你想问什么？", key="scanner_chat_input"):
            st.chat_message("user").markdown(prompt)
            st.session_state.scanner_messages.append({"role": "user", "content": prompt})
            
            # 构建 Prompt
            system_prompt = f"""
            你是一个高级技术专家。你已经阅读了该项目的代码分析报告。
            
            【已有知识库】
            {st.session_state.project_context}
            
            【用户指令】
            请基于知识库回答用户的问题。如果涉及上下文历史（比如用户说“它在哪”），请结合上文理解。
            """
            
            try:
                # 检查必要的参数
                if not api_key:
                     st.error("请先配置 API Key")
                else:
                    client = OpenAI(base_url=base_url, api_key=api_key)
                    
                    with st.chat_message("assistant"):
                        stream = client.chat.completions.create(
                            model=model_name,
                            messages=[
                                {"role": "system", "content": system_prompt},
                                *st.session_state.scanner_messages[-6:] # 保留最近 6 轮对话历史
                            ],
                            stream=True,
                            temperature=0.3
                        )
                        response = st.write_stream(stream)
                    
                    st.session_state.scanner_messages.append({"role": "assistant", "content": response})
                
            except Exception as e:
                st.error(f"对话发生错误: {e}")

    elif not btn_scan and not uploaded_file:
        st.info("👈 请配置 API 并选择 [扫描新项目] 或 [读取 Excel] 开始。")
