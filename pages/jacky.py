import pandas as pd
import streamlit as st
from datetime import datetime


def jacky_page():
    st.header("作者主页")
    col1,col2,col3 = st.columns(3)
    
    with col1:
        if st.button("📖 打开作者主页", use_container_width=True):
            st.markdown("[作者主页](https://jackyjay.cn)")
        if st.button("🔍 打开百度", use_container_width=True):
            st.markdown("[百度](https://www.baidu.com)")
    
    with col2:
        if st.button("📚 打开GitHub", use_container_width=True):
            st.markdown("[GitHub](https://github.com)")
        if st.button("💬 打开Stack Overflow", use_container_width=True):
            st.markdown("[Stack Overflow](https://stackoverflow.com)")
    
    with col3:
        if st.button("📊 打开Streamlit文档", use_container_width=True):
            st.markdown("[Streamlit文档](https://docs.streamlit.io)")
        if st.button("🐼 打开Pandas文档", use_container_width=True):
            st.markdown("[Pandas文档](https://pandas.pydata.org/docs)")


def grand_match():
    import model_GRAND_match.model_grand_match
    model_GRAND_match.model_grand_match.grand_match()
