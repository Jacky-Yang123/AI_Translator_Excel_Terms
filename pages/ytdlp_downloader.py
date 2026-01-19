import os
import glob
import io
import pandas as pd
import streamlit as st
import yt_dlp
from utils import Utils, HAS_WORDCLOUD


def ytdlp_downloader_app():
    """
    Streamlit 主函数组件。
    调用此函数即可在任何页面渲染下载器。
    """
    
    if 'ytdlp_queue' not in st.session_state: 
        st.session_state.ytdlp_queue = []
    if 'ytdlp_history' not in st.session_state: 
        st.session_state.ytdlp_history = []
    if 'current_meta' not in st.session_state: 
        st.session_state.current_meta = None
    if 'available_formats' not in st.session_state: 
        st.session_state.available_formats = []
    
    CONFIG_FILE = "ytdlp_config.json"
    config = Utils.load_config(CONFIG_FILE)

    st.title("📺 YT-DLP 全能媒体终端")
    if not HAS_WORDCLOUD:
        st.warning("⚠️ 检测到未安装 wordcloud 库，词云功能将不可用，但下载功能正常。")

    with st.sidebar:
        st.header("⚙️ 设置")
        new_path = st.text_input("📂 保存路径", value=config['save_path'])
        if new_path != config['save_path']:
            config['save_path'] = new_path
            Utils.save_config(CONFIG_FILE, config)

        st.divider()
        st.subheader("🍪 Cookie (VIP/评论)")
        raw_cookie = st.text_area("粘贴 Cookie (SESSDATA=...)", height=100, help="F12抓取B站请求头中的Cookie")
        
        temp_cookie_path = None
        if raw_cookie and "SESSDATA" in raw_cookie:
            temp_cookie_path = Utils.create_netscape_cookie_file(raw_cookie)
            if temp_cookie_path: 
                st.success("✅ Cookie 已激活")
        
        st.divider()
        config['proxy'] = st.text_input("代理 (Proxy)", value=config['proxy'])
        if st.button("📂 打开文件夹"): 
            if os.path.exists(config['save_path']): 
                Utils.open_folder(config['save_path'])

    tab_dl, tab_review = st.tabs(["⬇️ 下载与解析", "👁️ 资产管理与词云"])

    with tab_dl:
        col1, col2 = st.columns([4,1])
        with col1: 
            url = st.text_input("视频链接", key="url_input")
        with col2: 
            btn_analyze = st.button("🔍 解析", use_container_width=True, type="primary")

        def get_opts():
            opts = {'quiet': True, 'proxy': config['proxy'] or None, 'no_warnings': True, 'extractor_args': {'bilibili': {'comment_sort': 'time'}}}
            if temp_cookie_path: 
                opts['cookiefile'] = temp_cookie_path
            return opts

        if btn_analyze and url:
            with st.spinner("正在解析流..."):
                try:
                    with yt_dlp.YoutubeDL(get_opts()) as ydl:
                        meta = ydl.extract_info(url, download=False)
                        st.session_state.current_meta = meta
                        formats = meta.get('formats', [])
                        heights = sorted(list(set([f.get('height') for f in formats if f.get('height')])), reverse=True)
                        st.session_state.available_formats = [f"{h}p" for h in heights]
                except Exception as e: 
                    st.error(f"解析错误: {e}")

        if st.session_state.current_meta:
            meta = st.session_state.current_meta
            st.divider()
            c1, c2 = st.columns([1, 2])
            with c1: 
                if meta.get('thumbnail'): 
                    st.image(meta['thumbnail'], use_container_width=True)
            with c2:
                st.subheader(meta.get('title'))
                quality = st.selectbox("画质选择", ["✨ 最佳 (MP4)"] + st.session_state.available_formats + ["🎵 纯音频"])
                
                cd1, cd2 = st.columns(2)
                with cd1: 
                    get_danmaku = st.checkbox("导出弹幕 Excel", value=True)
                with cd2: 
                    get_comments = st.checkbox("导出评论 Excel", value=True)
                
                limit_cmt = 100
                if get_comments: 
                    limit_cmt = st.slider("评论抓取量", 10, 5000, 500, step=50)

                if st.button("➕ 加入队列", type="primary"):
                    st.session_state.ytdlp_queue.append({
                        "url": meta['webpage_url'], "title": meta['title'], "quality": quality,
                        "danmaku": get_danmaku, "comments": get_comments, "limit_cmt": limit_cmt
                    })
                    st.success("已加入下载队列")

        if st.session_state.ytdlp_queue:
            st.divider()
            if st.button(f"🚀 开始下载 ({len(st.session_state.ytdlp_queue)} 个任务)", type="primary", use_container_width=True):
                prog = st.progress(0)
                for idx, task in enumerate(st.session_state.ytdlp_queue):
                    opts = get_opts()
                    opts.update({'outtmpl': os.path.join(config['save_path'], f"{task['title']}.%(ext)s"), 'ignoreerrors': True, 'merge_output_format': 'mp4', 'writeinfojson': True})
                    
                    if "纯音频" in task['quality']: 
                        opts['format'] = 'bestaudio/best'
                    elif "最佳" in task['quality']: 
                        opts['format'] = 'bestvideo+bestaudio/best'
                    else: 
                        opts['format'] = f"bestvideo[height={task['quality'].replace('p','')}]" + "+bestaudio/best"

                    if task['danmaku']: 
                        opts.update({'writesubtitles': True, 'allsubtitles': True})
                    if task['comments']: 
                        opts.update({'getcomments': True, 'max_comments': task['limit_cmt']})

                    try:
                        with yt_dlp.YoutubeDL(opts) as ydl:
                            ydl.download([task['url']])
                            base = os.path.join(config['save_path'], task['title'])
                            
                            if task['danmaku']:
                                xmls = glob.glob(f"{base}*.xml")
                                if xmls: 
                                    Utils.process_xml_to_excel(xmls[0], f"{base}_弹幕.xlsx")
                                    try: 
                                        os.remove(xmls[0])
                                    except: 
                                        pass
                            
                            if task['comments']:
                                json_f = f"{base}.info.json"
                                if os.path.exists(json_f):
                                    Utils.process_json_to_excel(json_f, f"{base}_评论.xlsx")
                                    try: 
                                        os.remove(json_f)
                                    except: 
                                        pass

                            st.session_state.ytdlp_history.append({"title": task['title'], "video_path": f"{base}.mp4", "base_name": base})
                    except Exception as e: 
                        st.error(f"任务失败: {e}")
                    prog.progress((idx+1)/len(st.session_state.ytdlp_queue))
                
                st.session_state.ytdlp_queue = []
                st.success("全部任务完成！")

    with tab_review:
        if not st.session_state.ytdlp_history: 
            st.info("暂无历史记录")
        
        for item in reversed(st.session_state.ytdlp_history):
            with st.expander(f"🎥 {item['title']}", expanded=True):
                c_vid, c_data = st.columns([1, 1.5])
                with c_vid:
                    if os.path.exists(item['video_path']): 
                        st.video(item['video_path'])
                    else: 
                        st.warning("文件未找到")
                
                with c_data:
                    dm_path = f"{item['base_name']}_弹幕.xlsx"
                    cm_path = f"{item['base_name']}_评论.xlsx"
                    
                    t1, t2 = st.tabs(["📊 数据", "☁️ 词云"])
                    with t1:
                        if os.path.exists(dm_path): 
                            st.dataframe(pd.read_excel(dm_path), height=150)
                        if os.path.exists(cm_path): 
                            st.dataframe(pd.read_excel(cm_path), height=150)
                    
                    with t2:
                        if not HAS_WORDCLOUD:
                            st.error("词云库缺失，请安装 wordcloud")
                        else:
                            wc1, wc2 = st.columns(2)
                            with wc1:
                                if os.path.exists(dm_path) and st.button("弹幕词云", key=f"d_{item['title']}"):
                                    wc = Utils.generate_wordcloud_img(pd.read_excel(dm_path)['内容'].tolist())
                                    if wc: 
                                        st.image(wc.to_array(), use_container_width=True)
                                        buf = io.BytesIO()
                                        wc.to_image().save(buf, format='PNG')
                                        st.download_button("下载", buf.getvalue(), "dm_wc.png", "image/png", key=f"dd_{item['title']}")
                            with wc2:
                                if os.path.exists(cm_path) and st.button("评论词云", key=f"c_{item['title']}"):
                                    wc = Utils.generate_wordcloud_img(pd.read_excel(cm_path)['内容'].tolist())
                                    if wc: 
                                        st.image(wc.to_array(), use_container_width=True)
                                        buf = io.BytesIO()
                                        wc.to_image().save(buf, format='PNG')
                                        st.download_button("下载", buf.getvalue(), "cm_wc.png", "image/png", key=f"dc_{item['title']}")
