# pages/danmu.py - 弹幕抓取页面

import re
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

from utils import Utils


def scrape_niconico_danmaku(video_url, cookies_file=None):
    """抓取Niconico弹幕"""
    try:
        import yt_dlp

        temp_dir = tempfile.mkdtemp()
        output_template = os.path.join(temp_dir, "%(title)s.%(ext)s")

        ydl_opts = {
            'outtmpl': output_template,
            'writesubtitles': True,
            'subtitlesformat': 'xml',
            'skip_download': True,
            'quiet': True,
        }

        if cookies_file:
            ydl_opts['cookiefile'] = cookies_file

        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(video_url, download=True)
            title = info.get('title', 'video')

        # 查找生成的XML文件
        xml_files = list(Path(temp_dir).glob("*.xml"))
        if xml_files:
            xml_path = xml_files[0]
            excel_path = xml_path.with_suffix('.xlsx')

            success, count = Utils.process_xml_to_excel(str(xml_path), str(excel_path))

            if success:
                with open(excel_path, 'rb') as f:
                    return True, f.read(), title, count
            else:
                return False, None, title, 0
        else:
            return False, None, info.get('title', 'video'), 0

    except Exception as e:
        return False, None, str(e), 0


def scrape_bilibili_danmaku(video_url, cookies_file=None):
    """抓取Bilibili弹幕"""
    try:
        import yt_dlp

        temp_dir = tempfile.mkdtemp()
        output_template = os.path.join(temp_dir, "%(title)s.%(ext)s")

        ydl_opts = {
            'outtmpl': output_template,
            'writesubtitles': True,
            'subtitlesformat': 'xml',
            'skip_download': True,
            'quiet': True,
        }

        if cookies_file:
            ydl_opts['cookiefile'] = cookies_file

        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(video_url, download=True)
            title = info.get('title', 'video')

        # 查找生成的XML文件
        xml_files = list(Path(temp_dir).glob("*.xml"))
        if xml_files:
            xml_path = xml_files[0]
            excel_path = xml_path.with_suffix('.xlsx')

            success, count = Utils.process_xml_to_excel(str(xml_path), str(excel_path))

            if success:
                with open(excel_path, 'rb') as f:
                    return True, f.read(), title, count
            else:
                return False, None, title, 0
        else:
            return False, None, info.get('title', 'video'), 0

    except Exception as e:
        return False, None, str(e), 0


def danmu_page():
    """弹幕抓取页面"""
    st.markdown("""
    # 🎬 弹幕抓取工具

    支持从 **Niconico** 和 **Bilibili** 抓取视频弹幕，并导出为 Excel 文件。
    """)

    # 侧边栏配置
    with st.sidebar:
        st.header("⚙️ 配置")
        platform = st.radio(
            "选择视频平台",
            options=["Niconico", "Bilibili"],
            help="选择您要抓取弹幕的视频平台",
            key="video_platform_selector"
        )

        st.divider()

        # Bilibili Cookie配置
        bilibili_cookies_file = None
        if platform == "Bilibili":
            st.subheader("🔐 Bilibili Cookie配置")
            st.info("部分视频需要登录才能查看，请上传Cookie文件")

            cookie_upload = st.file_uploader(
                "上传Cookie文件 (txt格式)",
                type=['txt'],
                key="bilibili_cookie_uploader"
            )

            if cookie_upload:
                # 保存上传的cookie文件
                temp_cookie = tempfile.NamedTemporaryFile(delete=False, suffix='.txt')
                temp_cookie.write(cookie_upload.read())
                temp_cookie.close()
                bilibili_cookies_file = temp_cookie.name
                st.success("✅ Cookie文件已上传")

            cookie_string = st.text_area(
                "或者粘贴Cookie字符串",
                placeholder="SESSDATA=xxx; bili_jct=xxx; ...",
                key="bilibili_cookie_string"
            )

            if cookie_string and not bilibili_cookies_file:
                # 将cookie字符串转换为Netscape格式
                bilibili_cookies_file = Utils.create_netscape_cookie_file(cookie_string)
                if bilibili_cookies_file:
                    st.success("✅ Cookie字符串已处理")

    # 主内容区域
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

    if scrape_button:
        if not video_url:
            st.error("❌ 请输入视频链接")
            return

        with st.spinner(f"正在从 {platform} 抓取弹幕..."):
            if platform == "Niconico":
                success, data, title, count = scrape_niconico_danmaku(video_url)
            else:
                success, data, title, count = scrape_bilibili_danmaku(video_url, bilibili_cookies_file)

            if success and data:
                st.success(f"✅ 成功抓取 {count} 条弹幕！")
                st.info(f"视频标题: {title}")

                # 显示预览
                try:
                    preview_df = pd.read_excel(BytesIO(data))
                    with st.expander("📊 弹幕预览", expanded=True):
                        st.dataframe(preview_df.head(50))
                except:
                    pass

                # 下载按钮
                st.download_button(
                    label="📥 下载弹幕Excel",
                    data=data,
                    file_name=f"{title}_danmaku_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.error(f"❌ 抓取失败: {title}")

    # 使用说明
    with st.expander("📖 使用说明"):
        st.markdown("""
        ## 使用说明

        ### Niconico
        1. 复制Niconico视频链接
        2. 粘贴到输入框
        3. 点击"开始抓取"

        ### Bilibili
        1. 复制Bilibili视频链接
        2. 如果需要登录才能查看的视频，请上传Cookie文件或粘贴Cookie字符串
        3. 点击"开始抓取"

        ### Cookie获取方法
        1. 在浏览器登录Bilibili
        2. 使用浏览器扩展导出Cookie（推荐使用Get Cookies.txt）
        3. 或者从浏览器开发者工具中复制Cookie字符串

        ### 注意事项
        - 需要安装yt-dlp库
        - 部分视频可能有地区限制
        - 弹幕数量可能因视频而异
        """)
