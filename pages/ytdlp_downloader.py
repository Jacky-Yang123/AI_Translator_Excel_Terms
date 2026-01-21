# pages/ytdlp_downloader.py - yt-dlp 视频下载器页面

import os
import json
import time
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

try:
    from wordcloud import WordCloud
    HAS_WORDCLOUD = True
except ImportError:
    HAS_WORDCLOUD = False

try:
    import matplotlib.pyplot as plt
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False

from utils import Utils


def ytdlp_downloader_app():
    """yt-dlp 视频下载器应用"""
    st.title("🎬 视频弹幕评论下载器")
    st.markdown("### 使用yt-dlp下载视频、弹幕和评论")

    # 配置文件路径
    config_file = os.path.join(os.path.expanduser("~"), ".ytdlp_downloader_config.json")
    config = Utils.load_config(config_file)

    # 侧边栏配置
    with st.sidebar:
        st.header("⚙️ 设置")

        save_path = st.text_input(
            "保存路径",
            value=config.get("save_path", os.path.join(os.path.expanduser("~"), "Downloads", "Yt-DLP-Data")),
            key="save_path_input"
        )

        proxy = st.text_input(
            "代理设置 (可选)",
            value=config.get("proxy", ""),
            placeholder="例如: http://127.0.0.1:7890",
            key="proxy_input"
        )

        naming_tmpl = st.text_input(
            "文件命名模板",
            value=config.get("naming_tmpl", "%(title)s"),
            key="naming_tmpl_input"
        )

        if st.button("💾 保存设置"):
            config["save_path"] = save_path
            config["proxy"] = proxy
            config["naming_tmpl"] = naming_tmpl
            Utils.save_config(config_file, config)
            st.success("✅ 设置已保存")

        st.divider()

        if st.button("📂 打开保存文件夹"):
            if os.path.exists(save_path):
                Utils.open_folder(save_path)
            else:
                os.makedirs(save_path, exist_ok=True)
                Utils.open_folder(save_path)

    # 主界面
    st.header("📥 下载设置")

    video_url = st.text_input(
        "视频链接",
        placeholder="请输入Bilibili/Niconico视频链接...",
        key="video_url_input"
    )

    col1, col2 = st.columns(2)

    with col1:
        platform = st.selectbox(
            "平台",
            options=["Bilibili", "Niconico", "YouTube", "其他"],
            key="platform_select"
        )

    with col2:
        download_type = st.multiselect(
            "下载内容",
            options=["视频", "弹幕", "评论", "字幕"],
            default=["弹幕"],
            key="download_type_select"
        )

    # Cookie设置
    with st.expander("🔐 Cookie设置（登录后内容需要）"):
        cookie_upload = st.file_uploader(
            "上传Cookie文件",
            type=['txt'],
            key="cookie_uploader"
        )

        cookie_string = st.text_area(
            "或粘贴Cookie字符串",
            placeholder="SESSDATA=xxx; bili_jct=xxx; ...",
            key="cookie_string_input"
        )

    # 下载按钮
    if st.button("🚀 开始下载", type="primary", use_container_width=True):
        if not video_url:
            st.error("❌ 请输入视频链接")
            return

        # 创建保存目录
        os.makedirs(save_path, exist_ok=True)

        # 准备Cookie文件
        cookies_file = None
        if cookie_upload:
            temp_cookie = tempfile.NamedTemporaryFile(delete=False, suffix='.txt')
            temp_cookie.write(cookie_upload.read())
            temp_cookie.close()
            cookies_file = temp_cookie.name
        elif cookie_string:
            cookies_file = Utils.create_netscape_cookie_file(cookie_string)

        try:
            import yt_dlp

            progress_bar = st.progress(0)
            status_text = st.empty()

            output_template = os.path.join(save_path, f"{naming_tmpl}.%(ext)s")

            ydl_opts = {
                'outtmpl': output_template,
                'quiet': True,
                'no_warnings': True,
            }

            if proxy:
                ydl_opts['proxy'] = proxy

            if cookies_file:
                ydl_opts['cookiefile'] = cookies_file

            # 根据下载类型设置选项
            if "视频" not in download_type:
                ydl_opts['skip_download'] = True

            if "弹幕" in download_type or "字幕" in download_type:
                ydl_opts['writesubtitles'] = True
                ydl_opts['subtitlesformat'] = 'xml'

            if "评论" in download_type:
                ydl_opts['getcomments'] = True
                ydl_opts['writeinfojson'] = True

            status_text.text("正在获取视频信息...")
            progress_bar.progress(20)

            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(video_url, download=True)
                title = info.get('title', 'video')

            progress_bar.progress(80)
            status_text.text("正在处理下载的文件...")

            results = []

            # 处理弹幕
            if "弹幕" in download_type:
                xml_files = list(Path(save_path).glob(f"*{title}*.xml"))
                for xml_file in xml_files:
                    excel_path = xml_file.with_suffix('.xlsx')
                    success, count = Utils.process_xml_to_excel(str(xml_file), str(excel_path))
                    if success:
                        results.append(f"弹幕: {count} 条")

            # 处理评论
            if "评论" in download_type:
                json_files = list(Path(save_path).glob(f"*{title}*.info.json"))
                for json_file in json_files:
                    excel_path = json_file.with_suffix('.comments.xlsx')
                    success, count = Utils.process_json_to_excel(str(json_file), str(excel_path))
                    if success:
                        results.append(f"评论: {count} 条")

            progress_bar.progress(100)
            status_text.empty()
            progress_bar.empty()

            if results:
                st.success(f"✅ 下载完成！视频: {title}")
                for result in results:
                    st.info(result)
            else:
                st.success(f"✅ 下载完成！视频: {title}")

        except ImportError:
            st.error("❌ 请安装yt-dlp: pip install yt-dlp")
        except Exception as e:
            st.error(f"❌ 下载失败: {e}")

    # 词云生成
    st.header("☁️ 词云生成")

    uploaded_excel = st.file_uploader(
        "上传弹幕/评论Excel文件",
        type=['xlsx', 'xls'],
        key="wordcloud_uploader"
    )

    if uploaded_excel:
        try:
            df = pd.read_excel(uploaded_excel)
            st.success(f"✅ 读取成功: {len(df)} 条数据")

            with st.expander("📊 数据预览"):
                st.dataframe(df.head(20))

            text_col = st.selectbox(
                "选择文本列",
                options=df.columns.tolist(),
                key="text_col_select"
            )

            if st.button("☁️ 生成词云", use_container_width=True):
                if not HAS_WORDCLOUD:
                    st.error("❌ 请安装wordcloud: pip install wordcloud")
                    return

                if not HAS_MATPLOTLIB:
                    st.error("❌ 请安装matplotlib: pip install matplotlib")
                    return

                text_list = df[text_col].dropna().astype(str).tolist()
                wc = Utils.generate_wordcloud_img(text_list)

                if wc:
                    fig, ax = plt.subplots(figsize=(10, 5))
                    ax.imshow(wc, interpolation='bilinear')
                    ax.axis('off')
                    st.pyplot(fig)

                    # 保存词云
                    img_buffer = BytesIO()
                    wc.to_image().save(img_buffer, format='PNG')

                    st.download_button(
                        label="📥 下载词云图片",
                        data=img_buffer.getvalue(),
                        file_name=f"wordcloud_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                        mime="image/png",
                        use_container_width=True
                    )
                else:
                    st.warning("⚠️ 无法生成词云，请检查数据")

        except Exception as e:
            st.error(f"❌ 处理失败: {e}")
