# API_AI_Excel翻译分析工具_Jacky

## 📖 项目简介

这是一个基于Python3.12.10以及Streamlit开发的多功能AI翻译与Excel处理工具集，集成了多种翻译API、Excel数据处理、媒体下载等功能，专为本地化翻译工作者和数据分析师设计。

## ✨ 主要功能

### 1. AI翻译工具
- **批量翻译工具**：支持多种API提供商（DeepSeek、OpenAI、自定义API），可同时翻译多种语言
- **提示词生成器**：生成优化的翻译提示词，支持术语库和角色性格信息
- **翻译结果处理**：将AI翻译结果与原始Excel文件自动合并

### 2. Excel数据处理
- **Excel查找替换**：在多个Excel文件中批量查找和替换内容，支持大小写敏感/不敏感模式
- **Excel高级替换**：基于条件的精准查找和替换功能
- **Excel表格对比**：对比两个Excel文件的差异，识别新增、删除和共同内容
- **ExcelABC操作**：批量处理Excel文件，支持删除、替换、修改中间值等操作
- **术语查找**：在术语库中快速查找特定术语
- **模板一键匹配**：快速匹配Excel模板
- **文件夹单向匹配程序**：批量处理文件夹中的Excel文件

### 3. 媒体处理工具
- **B站视频弹幕评论下载**：支持下载B站视频、弹幕和评论
- **Niconico弹幕抓取**：抓取Niconico网站的弹幕数据

### 4. 其他功能
- **Jacky的主页**：作者相关资源链接

## 📦 安装要求

### 依赖库
```
streamlit
pandas
jieba
openai
requests
yt-dlp
matplotlib
openpyxl
wordcloud
```

### 安装方法
1. 克隆或下载项目
2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
3. 运行应用：
   ```bash
   streamlit run app.py
   ```

## 🚀 使用指南

### 1. 批量翻译工具
- 在左侧选择"🔄 批量翻译工具"
- 配置API密钥和模型
- 选择目标语言（可多选）
- 上传待翻译的Excel文件
- 设置术语库和角色性格信息
- 开始翻译

### 2. 提示词生成器
- 在左侧选择"📝 提示词生成器"
- 选择目标语言
- 上传待翻译文本
- 加载术语库和角色性格库
- 生成翻译提示词

### 3. Excel表格对比
- 在左侧选择"🔍 Excel表格对比"
- 上传两个Excel文件
- 选择关键列
- 点击对比按钮查看结果

### 4. 媒体下载功能
- 在左侧选择"blbl视频弹幕评论下载"
- 输入视频链接
- 解析视频信息
- 选择下载格式和质量
- 下载视频、弹幕或评论

## 🔧 配置说明

### API配置
- 支持DeepSeek、OpenAI等API提供商
- 可自定义API URL和模型名称
- API密钥采用密码框输入，确保安全

### 媒体下载配置
- 支持设置保存路径
- 可配置代理
- 支持B站Cookie认证（用于VIP视频和评论下载）

## 📁 项目结构

```
ExcelTranslator_Pro/
├── app.py                  # 主应用入口（路由）
├── translator.py           # 翻译核心模块
├── utils.py                # 工具函数
├── api_config.py           # API配置
├── requirements.txt        # 项目依赖
├── model_GRAND_match/      # 模板匹配模型
│   └── model_grand_match.py
└── pages/                  # 功能页面模块
    ├── __init__.py              # 模块初始化
    ├── batch_translation.py     # 批量翻译
    ├── prompt_generator.py      # 提示词生成器
    ├── translation_result.py    # 翻译结果处理
    ├── ytdlp_downloader.py      # 媒体下载
    ├── danmu.py                 # 弹幕抓取
    ├── excel_replace.py         # Excel查找替换
    ├── excel_sreplace.py        # Excel高级替换
    ├── excel_comparison.py      # Excel表格对比
    ├── excel_abc.py             # ExcelABC操作
    ├── term_lookup.py           # 术语查找
    ├── grand_match.py           # 模板一键匹配
    ├── excel_matchpro.py        # 文件夹单向匹配
    └── jacky.py                 # 作者主页
```

## 📝 使用注意事项

1. **API密钥安全**：API密钥仅在当前会话中使用，不会被保存到文件
2. **文件格式**：支持Excel文件格式（.xlsx, .xls, .xlsm, .xlsb）
3. **性能优化**：处理大量数据时建议分批处理
4. **网络要求**：使用在线翻译API和媒体下载功能需要稳定的网络连接
5. **Cookie使用**：使用B站VIP功能时需要正确配置Cookie

## 🔄 更新日志

### v2.1 模块化版本 (2026-01-21)
- 将单体 app.py（6400+行）拆分为17个独立模块
- 新增 api_config.py 配置模块
- 修复大小写不敏感替换的bug
- 优化代码结构和可维护性

### v2.0 合并版
- 整合了所有功能模块
- 优化了用户界面
- 增强了多语言支持
- 添加了新的Excel处理功能

## 🤝 贡献

欢迎提交Issue和Pull Request！

## 📧 联系方式

作者：Jacky_9S
版本：v2.1 模块化版本

---

**使用说明**：
1. 运行`streamlit run app.py`启动应用
2. 在左侧导航栏选择需要的功能
3. 根据页面提示进行操作
4. 查看结果并下载处理后的文件
