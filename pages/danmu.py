import os
import sys
import re
import json
import subprocess
from datetime import datetime

import pandas as pd
import streamlit as st


def find_yt_dlp():
    """查找yt-dlp可执行文件"""
    if sys.platform.startswith('win'):
        candidates = ["yt-dlp.exe", "yt-dlp"]
    else:
        candidates = ["./yt-dlp", "yt-dlp"]
    
    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate
    return "yt-dlp"


def normalize_url(url):
    """将非标准的Niconico URL格式转换为标准格式"""
    return url.replace("www.video.nicovideo.jp", "www.nicovideo.jp")


def extract_watch_id(url):
    """从Niconico视频URL中提取watch ID (sm/nm号)"""
    match = re.search(r'(sm|nm)\d+', url)
    if match:
        return match.group(0)
    return "unknown_id"


def extract_bilibili_id(url):
    """从Bilibili视频URL中提取video ID (BV号或av号)"""
    bv_match = re.search(r'BV[a-zA-Z0-9]+', url)
    if bv_match:
        return bv_match.group(0)
    
    av_match = re.search(r'av(\d+)', url)
    if av_match:
        return f"av{av_match.group(1)}"
    
    return "unknown_id"


def run_yt_dlp_to_get_json(url, output_filename_base="danmaku"):
    """运行yt-dlp命令来抓取弹幕数据并保存为JSON文件"""
    yt_dlp_path = find_yt_dlp()
    
    command = [
        yt_dlp_path,
        "--skip-download",
        "--write-sub",
        "--all-subs",
        "--sub-format", "json",
        "--output", f"{output_filename_base}.%(ext)s",
        url
    ]
    
    try:
        result = subprocess.run(command, capture_output=True, text=True, check=True)
        
        json_filename = f"{output_filename_base}.comments.json"
        if os.path.exists(json_filename):
            return json_filename
        else:
            return None
            
    except subprocess.CalledProcessError as e:
        st.error(f"yt-dlp执行失败: {e.stderr}")
        return None
    except FileNotFoundError:
        st.error(f"找不到yt-dlp可执行文件。请确保yt-dlp已安装或在PATH中。")
        return None


def process_niconico_json_to_dataframe(json_path):
    """读取yt-dlp生成的JSON文件，处理Niconico弹幕数据"""
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        st.error(f"JSON文件处理失败: {e}")
        return None

    danmaku_list = []
    for comment in data:
        vpos_ms = comment.get("vposMs", 0)
        time_sec = vpos_ms / 1000
        video_time = time.strftime('%H:%M:%S', time.gmtime(time_sec))
        
        posted_at_str = comment.get("postedAt")
        try:
            posted_at = datetime.fromisoformat(posted_at_str)
            send_time = posted_at.strftime('%Y-%m-%d %H:%M:%S')
        except:
            send_time = posted_at_str
            
        commands = " ".join(comment.get("commands", []))
        
        danmaku_info = {
            "弹幕内容": comment.get("body"),
            "视频时间": video_time,
            "时间(秒)": time_sec,
            "格式/颜色": commands,
            "用户ID": comment.get("userId"),
            "发送时间": send_time,
            "编号": comment.get("no"),
        }
        danmaku_list.append(danmaku_info)
        
    if not danmaku_list:
        return None
        
    df = pd.DataFrame(danmaku_list)
    df = df[['编号', '视频时间', '时间(秒)', '弹幕内容', '格式/颜色', '用户ID', '发送时间']]
    return df


def scrape_niconico_danmaku(url):
    """抓取Niconico弹幕"""
    normalized_url = normalize_url(url)
    watch_id = extract_watch_id(normalized_url)
    
    with st.spinner(f"正在抓取Niconico视频 {watch_id} 的弹幕..."):
        json_path = run_yt_dlp_to_get_json(normalized_url, output_filename_base=watch_id)
        
        if json_path:
            df = process_niconico_json_to_dataframe(json_path)
            
            try:
                os.remove(json_path)
            except OSError:
                pass
            
            return df, watch_id
        else:
            return None, watch_id


def scrape_bilibili_danmaku(url, cookies_file=None):
    """抓取Bilibili弹幕"""
    video_id = extract_bilibili_id(url)
    
    with st.spinner(f"正在抓取Bilibili视频 {video_id} 的弹幕..."):
        yt_dlp_path = find_yt_dlp()
        
        command = [
            yt_dlp_path,
            "--skip-download",
            "--write-sub",
            "--all-subs",
            "--sub-format", "json",
            "--output", f"{video_id}.%(ext)s",
        ]
        
        if cookies_file and os.path.exists(cookies_file):
            command.extend(["--cookies", cookies_file])
        
        command.append(url)
        
        try:
            result = subprocess.run(command, capture_output=True, text=True, check=True, timeout=60)
            
            json_filename = f"{video_id}.comments.json"
            if os.path.exists(json_filename):
                try:
                    with open(json_filename, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                except (FileNotFoundError, json.JSONDecodeError) as e:
                    st.error(f"JSON文件处理失败: {e}")
                    return None, video_id
                
                danmaku_list = []
                for comment in data:
                    danmaku_info = {
                        "弹幕内容": comment.get("body", comment.get("text", "")),
                        "发送时间": comment.get("postedAt", comment.get("timestamp", "")),
                        "用户ID": comment.get("userId", comment.get("author", "")),
                    }
                    danmaku_list.append(danmaku_info)
                
                if danmaku_list:
                    df = pd.DataFrame(danmaku_list)
                    
                    try:
                        os.remove(json_filename)
                    except OSError:
                        pass
                    
                    return df, video_id
                else:
                    st.warning("未找到弹幕数据")
                    return None, video_id
            else:
                st.error(f"yt-dlp未生成预期的文件")
                return None, video_id
                
        except subprocess.CalledProcessError as e:
            error_msg = e.stderr if e.stderr else str(e)
            if "412" in error_msg or "Precondition Failed" in error_msg:
                st.error(
                    "❌ Bilibili API速率限制（HTTP 412错误）\n\n"
                    "这通常是因为：\n"
                    "1. 没有提供有效的Cookie认证\n"
                    "2. 该IP地址的请求已达到限制\n\n"
                    "**解决方案：**\n"
                    "请在左侧栏上传您的Bilibili Cookie文件，然后重试。"
                )
            else:
                st.error(f"yt-dlp执行失败: {error_msg}")
            return None, video_id
        except subprocess.TimeoutExpired:
            st.error("❌ 请求超时，请检查网络连接或稍后重试")
            return None, video_id
        except FileNotFoundError:
            st.error(f"找不到yt-dlp可执行文件")
            return None, video_id


def danmu_page():
    st.markdown("""
    # 🎬 弹幕抓取工具
    
    支持从 **Niconico** 和 **Bilibili** 抓取视频弹幕，并导出为 Excel 文件。
    """)
    
    with st.sidebar:
        st.header("⚙️ 配置1")
        platform = st.radio(
            "选择视频平台",
            options=["Niconico", "Bilibili"],
            help="选择您要抓取弹幕的视频平台",
            key="video_pla_selector"
        )
        
        st.divider()
        
        bilibili_cookies_file = None
        if platform == "Bilibili":
            st.subheader("🔐 Bilibili Cookie配置")
            st.markdown(
                """Bilibili需要Cookie认证以避免速率限制。\n\n
                **获取Cookie的方法：**
                1. 打开浏览器访问 https://www.bilibili.com
                2. 登录您的账号
                3. 按F12打开开发者工具 → Application → Cookies
                4. 复制所有Cookie内容到文本文件
                5. 上传该文件
                """
            )
            
            uploaded_file = st.file_uploader(
                "上传Cookie文件",
                type=["txt"],
                help="上传从浏览器导出的Cookie文件"
            )
            
            if uploaded_file is not None:
                cookies_content = uploaded_file.read().decode('utf-8')
                bilibili_cookies_file = "temp_cookies.txt"
                with open(bilibili_cookies_file, 'w', encoding='utf-8') as f:
                    f.write(cookies_content)
                st.success("✅ Cookie文件已加载")
            else:
                st.warning("⚠️ 未上传Cookie文件，可能导致速率限制错误")
        
        st.divider()
        
        st.markdown("""
        ### 📌 使用说明
        
        **Niconico:**
        - 输入格式: `https://www.nicovideo.jp/watch/sm500873`
        - 或: `https://www.video.nicovideo.jp/watch/sm500873` (会自动转换)
        
        **Bilibili:**
        - 输入格式: `https://www.bilibili.com/video/BV1xx411c7mD`
        - 或: `https://www.bilibili.com/video/av123456789`
        - 建议上传Cookie文件以避免速率限制
        """)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader(f"输入{platform}视频链接")
        video_url = st.text_input(
            "视频链接",
            placeholder=f"请输入{platform}视频链接...",
            label_visibility="collapsed"
        )
    
    with col2:
        st.subheader("操作")
        scrape_button = st.button(
            "🔍 开始抓取",
            use_container_width=True,
            type="primary"
        )
    
    st.divider()
    
    if scrape_button:
        if not video_url.strip():
            st.error("❌ 请输入视频链接")
        else:
            if platform == "Niconico":
                df, video_id = scrape_niconico_danmaku(video_url)
            else:
                df, video_id = scrape_bilibili_danmaku(video_url, cookies_file=bilibili_cookies_file)
            
            if df is not None and len(df) > 0:
                st.success(f"✅ 成功抓取 {len(df)} 条弹幕！")
                
                st.subheader("📊 弹幕数据预览")
                st.dataframe(df, use_container_width=True, height=400)
                
                st.subheader("💾 导出选项")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    excel_buffer = pd.ExcelWriter(
                        f"danmaku_{video_id}.xlsx",
                        engine='openpyxl'
                    )
                    df.to_excel(excel_buffer, index=False, sheet_name='弹幕数据')
                    excel_buffer.close()
                    
                    with open(f"danmaku_{video_id}.xlsx", 'rb') as f:
                        st.download_button(
                            label="📥 下载 Excel",
                            data=f.read(),
                            file_name=f"danmaku_{video_id}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    
                    try:
                        os.remove(f"danmaku_{video_id}.xlsx")
                    except OSError:
                        pass
                    
                    if bilibili_cookies_file and os.path.exists(bilibili_cookies_file):
                        try:
                            os.remove(bilibili_cookies_file)
                        except OSError:
                            pass
                
                with col2:
                    csv_buffer = df.to_csv(index=False)
                    st.download_button(
                        label="📥 下载 CSV",
                        data=csv_buffer,
                        file_name=f"danmaku_{video_id}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                
                with col3:
                    json_buffer = df.to_json(orient='records', force_ascii=False)
                    st.download_button(
                        label="📥 下载 JSON",
                        data=json_buffer,
                        file_name=f"danmaku_{video_id}.json",
                        mime="application/json",
                        use_container_width=True
                    )
                
                st.subheader("📈 统计信息")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("总弹幕数", len(df))
                
                with col2:
                    if '用户ID' in df.columns:
                        unique_users = df['用户ID'].nunique()
                        st.metric("独立用户数", unique_users)
                
                with col3:
                    if '弹幕内容' in df.columns:
                        avg_length = df['弹幕内容'].str.len().mean()
                        st.metric("平均弹幕长度", f"{avg_length:.1f} 字符")
            
            elif df is not None and len(df) == 0:
                st.warning("⚠️ 未找到弹幕数据，请检查视频链接是否正确")
            else:
                st.error("❌ 抓取失败，请检查视频链接或网络连接")
    
    st.divider()
    st.markdown("""
    ### 📖 功能说明
    
    **Niconico弹幕抓取：**
    - 支持所有Niconico视频
    - 自动提取弹幕时间戳、用户ID等信息
    - 支持导出Excel、CSV、JSON格式
    
    **Bilibili弹幕抓取：**
    - 支持BV号和av号视频
    - 推荐使用Cookie认证以避免速率限制
    - 支持导出Excel、CSV、JSON格式
    
    **注意事项：**
    - 请确保已安装yt-dlp工具
    - 抓取大量弹幕可能需要较长时间
    - 请遵守各平台的使用条款
    """)
