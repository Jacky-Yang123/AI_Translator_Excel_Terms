import pandas as pd
import streamlit as st
from datetime import datetime
from translator import MultiAPIExcelTranslator, get_api_providers


def get_preset_options():
    presets = {
        "游戏UI简约风格": "游戏UI简约风格",
        "角色对话自然流畅": "角色对话自然流畅", 
        "专业术语统一": "专业术语统一",
        "保持原文风格": "保持原文风格",
        "本地化适配": "本地化适配",
        "保持上下文一致性": "保持上下文一致性",
        "根据角色调整语气": "根据角色调整语气"
    }
    return presets

def get_preset_languages():
    return ["英文", "日文", "韩文", "法文", "德文", "西班牙文", "自定义"]

def get_default_custom_requirements():
    return "角色对话自然流畅；专业术语统一；保持原文风格；本地化适配；保持上下文一致性；根据角色调整语气；请注意使用语体，且所有角色除了微型机和炽长，其他都为女生用语，不要用男性用语，现在角色们都十分熟悉彼此了，不需要使用太正式尊重的语体了例如日语的话不需要ですます型了。"
