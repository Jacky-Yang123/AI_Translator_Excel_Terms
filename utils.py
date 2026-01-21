# utils.py - 共享工具函数和类

import os
import sys
import re
import json
import time
import shutil
import platform
import subprocess
import tempfile
import difflib
import threading
from pathlib import Path
from datetime import datetime
from io import BytesIO
from difflib import SequenceMatcher
from threading import Lock
from concurrent.futures import ThreadPoolExecutor, as_completed
import concurrent.futures

import pandas as pd
import streamlit as st
import openpyxl
import jieba

try:
    from wordcloud import WordCloud
    HAS_WORDCLOUD = True
except ImportError:
    HAS_WORDCLOUD = False

import xml.etree.ElementTree as ET


# --- 核心工具函数类 ---
class Utils:
    @staticmethod
    def load_config(config_file):
        if os.path.exists(config_file):
            with open(config_file, 'r') as f:
                return json.load(f)
        return {
            "save_path": os.path.join(os.path.expanduser("~"), "Downloads", "Yt-DLP-Data"),
            "proxy": "",
            "naming_tmpl": "%(title)s"
        }

    @staticmethod
    def save_config(config_file, config):
        with open(config_file, 'w') as f:
            json.dump(config, f)

    @staticmethod
    def open_folder(path):
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])

    @staticmethod
    def create_netscape_cookie_file(raw_cookie_str):
        if not raw_cookie_str or "=" not in raw_cookie_str:
            return None
        try:
            fd, path = tempfile.mkstemp(suffix='.txt', text=True)
            with os.fdopen(fd, 'w') as f:
                f.write("# Netscape HTTP Cookie File\n\n")
                for item in raw_cookie_str.split(';'):
                    if '=' in item:
                        key, value = item.strip().split('=', 1)
                        f.write(f".bilibili.com\tTRUE\t/\tFALSE\t253402300799\t{key}\t{value}\n")
            return path
        except:
            return None

    @staticmethod
    def get_chinese_font():
        system = platform.system()
        if system == "Windows":
            fonts = ["simhei.ttf", "msyh.ttc", "simsun.ttc"]
            for f in fonts:
                path = os.path.join("C:\\Windows\\Fonts", f)
                if os.path.exists(path):
                    return path
        elif system == "Darwin":
            return "/System/Library/Fonts/PingFang.ttc"
        return None

    @staticmethod
    def generate_wordcloud_img(text_list):
        if not HAS_WORDCLOUD:
            return None
        if not text_list:
            return None

        full_text = " ".join([str(t) for t in text_list if str(t)])
        cut_text = " ".join(jieba.cut(full_text))

        font_path = Utils.get_chinese_font()
        params = {
            'background_color': 'white', 'width': 800, 'height': 400,
            'max_words': 200, 'colormap': 'viridis',
            'stopwords': {'的', '了', '是', '在', '也', '就', '不', '都', '吗', '啊', '吧', '我', '这'}
        }
        if font_path:
            params['font_path'] = font_path

        wc = WordCloud(**params).generate(cut_text)
        return wc

    @staticmethod
    def process_xml_to_excel(xml_path, excel_path):
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            data = []
            for d in root.findall('d'):
                content = d.text
                p_attr = d.get('p')
                if p_attr and content:
                    attrs = p_attr.split(',')
                    if len(attrs) >= 7:
                        data.append({
                            "时间": f"{int(float(attrs[0])//60):02d}:{int(float(attrs[0])%60):02d}",
                            "秒数": round(float(attrs[0]), 2),
                            "内容": content,
                            "用户Hash": attrs[6],
                            "日期": datetime.fromtimestamp(int(attrs[4])).strftime('%Y-%m-%d')
                        })
            if data:
                pd.DataFrame(data).sort_values(by="秒数").to_excel(excel_path, index=False)
                return True, len(data)
            return False, 0
        except:
            return False, 0

    @staticmethod
    def process_json_to_excel(json_path, excel_path):
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                info = json.load(f)
            comments = info.get('comments', [])
            data = []
            for c in comments:
                data.append({
                    "用户": c.get('author'),
                    "内容": c.get('text'),
                    "点赞": c.get('like_count'),
                    "时间": datetime.fromtimestamp(c.get('timestamp', 0)).strftime('%Y-%m-%d %H:%M') if c.get('timestamp') else "-"
                })
            if data:
                pd.DataFrame(data).to_excel(excel_path, index=False)
                return True, len(data)
            return False, 0
        except:
            return False, 0


# --- 通用辅助函数 ---
def similar(a, b):
    """计算两个字符串的相似度"""
    return SequenceMatcher(None, str(a), str(b)).ratio()


def calculate_similarity(str1, str2):
    """计算两个字符串的相似度（0-1）"""
    if not str1 and not str2:
        return 1.0
    if not str1 or not str2:
        return 0.0
    return difflib.SequenceMatcher(None, str1, str2).ratio()


def open_folder(file_path):
    """打开文件所在的文件夹"""
    try:
        folder_path = os.path.dirname(os.path.abspath(file_path))
        system = platform.system()

        if system == "Windows":
            os.startfile(folder_path)
        elif system == "Darwin":
            subprocess.run(["open", folder_path])
        else:
            subprocess.run(["xdg-open", folder_path])

        st.success(f"✅ 已打开文件夹: {folder_path}")
    except Exception as e:
        st.error(f"❌ 打开文件夹失败: {str(e)}")


def open_file(file_path):
    """打开Excel文件"""
    try:
        system = platform.system()

        if system == "Windows":
            os.startfile(file_path)
        elif system == "Darwin":
            subprocess.run(["open", file_path])
        else:
            subprocess.run(["xdg-open", file_path])

        st.success(f"✅ 已打开文件: {Path(file_path).name}")
    except Exception as e:
        st.error(f"❌ 打开文件失败: {str(e)}")


def find_excel_files(folder_path):
    """查找文件夹中的所有Excel文件"""
    excel_files = []
    folder_path = Path(folder_path)

    if not folder_path.exists():
        return False, "文件夹路径不存在"

    excel_extensions = ['*.xlsx', '*.xls', '*.xlsm', '*.xlsb']

    for ext in excel_extensions:
        excel_files.extend(folder_path.rglob(ext))

    return True, excel_files


def highlight_keyword(text, keyword, case_sensitive=False):
    """高亮显示关键词"""
    if not text or not keyword:
        return text

    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.escape(keyword)

    if case_sensitive:
        highlighted = re.sub(f'({pattern})', r'**\1**', text)
    else:
        highlighted = re.sub(f'({pattern})', r'**\1**', text, flags=re.IGNORECASE)

    return highlighted


# 全局锁用于线程安全的进度更新
progress_lock = Lock()
