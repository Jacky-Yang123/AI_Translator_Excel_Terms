# pages/format_factory.py - 格式工厂页面
# 多功能格式转换和媒体编辑工具

import os
import io
import re
import tempfile
import subprocess
from pathlib import Path
from datetime import datetime

import streamlit as st

# 图片处理
try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# 音频处理
try:
    from pydub import AudioSegment
    from pydub.effects import normalize, speedup
    HAS_PYDUB = True
except ImportError:
    HAS_PYDUB = False

# PDF处理
try:
    from pdf2image import convert_from_bytes
    HAS_PDF2IMAGE = True
except ImportError:
    HAS_PDF2IMAGE = False

try:
    import img2pdf
    HAS_IMG2PDF = True
except ImportError:
    HAS_IMG2PDF = False

# Excel/CSV处理
import pandas as pd


def check_ffmpeg():
    """检查FFmpeg是否可用"""
    try:
        result = subprocess.run(
            ['ffmpeg', '-version'],
            capture_output=True,
            text=True,
            timeout=5
        )
        return result.returncode == 0
    except (subprocess.TimeoutExpired, FileNotFoundError):
        # 尝试查找项目目录中的ffmpeg
        project_dir = Path(__file__).parent.parent
        ffmpeg_path = project_dir / "ffmpeg.exe"
        if ffmpeg_path.exists():
            return True
        return False


def get_ffmpeg_path():
    """获取FFmpeg路径"""
    # 先尝试系统PATH
    try:
        result = subprocess.run(['ffmpeg', '-version'], capture_output=True, timeout=5)
        if result.returncode == 0:
            return 'ffmpeg'
    except:
        pass
    
    # 尝试项目目录
    project_dir = Path(__file__).parent.parent
    ffmpeg_path = project_dir / "ffmpeg.exe"
    if ffmpeg_path.exists():
        return str(ffmpeg_path)
    
    return None


def run_ffmpeg(args, input_file=None, output_file=None):
    """运行FFmpeg命令"""
    ffmpeg = get_ffmpeg_path()
    if not ffmpeg:
        return False, "FFmpeg未找到"
    
    cmd = [ffmpeg, '-y']  # -y 覆盖输出文件
    
    if input_file:
        cmd.extend(['-i', input_file])
    
    cmd.extend(args)
    
    if output_file:
        cmd.append(output_file)
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300  # 5分钟超时
        )
        if result.returncode == 0:
            return True, "成功"
        else:
            return False, result.stderr
    except subprocess.TimeoutExpired:
        return False, "处理超时"
    except Exception as e:
        return False, str(e)


# ==================== 图片处理函数 ====================

def convert_image(input_bytes, input_format, output_format, quality=85):
    """转换图片格式"""
    if not HAS_PIL:
        return None, "请安装Pillow: pip install Pillow"
    
    try:
        img = Image.open(io.BytesIO(input_bytes))
        
        # 处理RGBA到RGB转换（用于不支持透明度的格式）
        if output_format.upper() in ['JPEG', 'JPG', 'BMP'] and img.mode in ['RGBA', 'P']:
            # 创建白色背景
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
            img = background
        
        output = io.BytesIO()
        
        # 根据格式保存
        if output_format.upper() in ['JPEG', 'JPG']:
            img.save(output, format='JPEG', quality=quality, optimize=True)
        elif output_format.upper() == 'PNG':
            img.save(output, format='PNG', optimize=True)
        elif output_format.upper() == 'WEBP':
            img.save(output, format='WEBP', quality=quality)
        elif output_format.upper() == 'GIF':
            img.save(output, format='GIF')
        elif output_format.upper() == 'BMP':
            img.save(output, format='BMP')
        elif output_format.upper() == 'ICO':
            img.save(output, format='ICO')
        elif output_format.upper() == 'TIFF':
            img.save(output, format='TIFF')
        else:
            img.save(output, format=output_format.upper())
        
        output.seek(0)
        return output.getvalue(), None
    except Exception as e:
        return None, str(e)


def compress_image(input_bytes, quality=50, max_size=None):
    """压缩图片"""
    if not HAS_PIL:
        return None, "请安装Pillow: pip install Pillow"
    
    try:
        img = Image.open(io.BytesIO(input_bytes))
        
        # 调整大小
        if max_size:
            img.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
        
        # 转换为RGB
        if img.mode in ['RGBA', 'P']:
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
            img = background
        
        output = io.BytesIO()
        img.save(output, format='JPEG', quality=quality, optimize=True)
        output.seek(0)
        return output.getvalue(), None
    except Exception as e:
        return None, str(e)


# ==================== 音频处理函数 ====================

def convert_audio_format(input_bytes, input_format, output_format):
    """转换音频格式"""
    if not HAS_PYDUB:
        return None, "请安装pydub: pip install pydub"
    
    try:
        audio = AudioSegment.from_file(io.BytesIO(input_bytes), format=input_format.lower())
        output = io.BytesIO()
        audio.export(output, format=output_format.lower())
        output.seek(0)
        return output.getvalue(), None
    except Exception as e:
        return None, str(e)


def trim_audio(input_bytes, input_format, start_ms, end_ms):
    """裁剪音频"""
    if not HAS_PYDUB:
        return None, "请安装pydub: pip install pydub"
    
    try:
        audio = AudioSegment.from_file(io.BytesIO(input_bytes), format=input_format.lower())
        trimmed = audio[start_ms:end_ms]
        output = io.BytesIO()
        trimmed.export(output, format=input_format.lower())
        output.seek(0)
        return output.getvalue(), None
    except Exception as e:
        return None, str(e)


def adjust_audio_speed(input_bytes, input_format, speed_factor, preserve_pitch=True):
    """调整音频速度"""
    if not HAS_PYDUB:
        return None, "请安装pydub: pip install pydub"
    
    try:
        audio = AudioSegment.from_file(io.BytesIO(input_bytes), format=input_format.lower())
        
        if preserve_pitch:
            # 使用pydub的speedup（只支持加速且保留音调）
            if speed_factor > 1:
                # 加速
                audio = speedup(audio, playback_speed=speed_factor)
            else:
                # 减速 - 通过改变帧率实现（会改变音调）
                new_frame_rate = int(audio.frame_rate * speed_factor)
                audio = audio._spawn(audio.raw_data, overrides={'frame_rate': new_frame_rate})
                audio = audio.set_frame_rate(44100)
        else:
            # 直接改变帧率（会改变音调）
            new_frame_rate = int(audio.frame_rate * speed_factor)
            audio = audio._spawn(audio.raw_data, overrides={'frame_rate': new_frame_rate})
            audio = audio.set_frame_rate(44100)
        
        output = io.BytesIO()
        audio.export(output, format=input_format.lower())
        output.seek(0)
        return output.getvalue(), None
    except Exception as e:
        return None, str(e)


def adjust_audio_volume(input_bytes, input_format, volume_db):
    """调整音频音量"""
    if not HAS_PYDUB:
        return None, "请安装pydub: pip install pydub"
    
    try:
        audio = AudioSegment.from_file(io.BytesIO(input_bytes), format=input_format.lower())
        audio = audio + volume_db  # 增加或减少分贝
        output = io.BytesIO()
        audio.export(output, format=input_format.lower())
        output.seek(0)
        return output.getvalue(), None
    except Exception as e:
        return None, str(e)


def fade_audio(input_bytes, input_format, fade_in_ms=0, fade_out_ms=0):
    """音频淡入淡出"""
    if not HAS_PYDUB:
        return None, "请安装pydub: pip install pydub"
    
    try:
        audio = AudioSegment.from_file(io.BytesIO(input_bytes), format=input_format.lower())
        
        if fade_in_ms > 0:
            audio = audio.fade_in(fade_in_ms)
        if fade_out_ms > 0:
            audio = audio.fade_out(fade_out_ms)
        
        output = io.BytesIO()
        audio.export(output, format=input_format.lower())
        output.seek(0)
        return output.getvalue(), None
    except Exception as e:
        return None, str(e)


def reverse_audio(input_bytes, input_format):
    """音频倒放"""
    if not HAS_PYDUB:
        return None, "请安装pydub: pip install pydub"
    
    try:
        audio = AudioSegment.from_file(io.BytesIO(input_bytes), format=input_format.lower())
        reversed_audio = audio.reverse()
        output = io.BytesIO()
        reversed_audio.export(output, format=input_format.lower())
        output.seek(0)
        return output.getvalue(), None
    except Exception as e:
        return None, str(e)


def merge_audios(audio_files, output_format):
    """合并多个音频"""
    if not HAS_PYDUB:
        return None, "请安装pydub: pip install pydub"
    
    try:
        combined = AudioSegment.empty()
        
        for audio_file in audio_files:
            audio = AudioSegment.from_file(io.BytesIO(audio_file['bytes']), 
                                           format=audio_file['format'].lower())
            combined += audio
        
        output = io.BytesIO()
        combined.export(output, format=output_format.lower())
        output.seek(0)
        return output.getvalue(), None
    except Exception as e:
        return None, str(e)


# ==================== 视频处理函数 ====================

def convert_video_format(input_path, output_path, output_format):
    """转换视频格式"""
    args = ['-c:v', 'libx264', '-c:a', 'aac']
    return run_ffmpeg(args, input_path, output_path)


def trim_video(input_path, output_path, start_time, end_time):
    """裁剪视频"""
    args = ['-ss', start_time, '-to', end_time, '-c', 'copy']
    return run_ffmpeg(args, input_path, output_path)


def change_video_speed(input_path, output_path, speed_factor):
    """改变视频速度"""
    # 视频滤镜
    video_filter = f"setpts={1/speed_factor}*PTS"
    audio_filter = f"atempo={speed_factor}" if speed_factor <= 2 else f"atempo=2,atempo={speed_factor/2}"
    
    args = ['-filter:v', video_filter, '-filter:a', audio_filter]
    return run_ffmpeg(args, input_path, output_path)


def reverse_video(input_path, output_path):
    """视频倒放"""
    args = ['-vf', 'reverse', '-af', 'areverse']
    return run_ffmpeg(args, input_path, output_path)


def extract_audio_from_video(input_path, output_path):
    """从视频提取音频"""
    args = ['-vn', '-acodec', 'libmp3lame', '-q:a', '2']
    return run_ffmpeg(args, input_path, output_path)


def mute_video(input_path, output_path):
    """静音视频"""
    args = ['-c:v', 'copy', '-an']
    return run_ffmpeg(args, input_path, output_path)


def compress_video(input_path, output_path, crf=28):
    """压缩视频"""
    args = ['-c:v', 'libx264', '-crf', str(crf), '-preset', 'medium', '-c:a', 'aac', '-b:a', '128k']
    return run_ffmpeg(args, input_path, output_path)


# ==================== 文档处理函数 ====================

def excel_to_csv(input_bytes, sheet_name=None):
    """Excel转CSV"""
    try:
        df = pd.read_excel(io.BytesIO(input_bytes), sheet_name=sheet_name)
        output = io.StringIO()
        df.to_csv(output, index=False, encoding='utf-8-sig')
        return output.getvalue().encode('utf-8-sig'), None
    except Exception as e:
        return None, str(e)


def csv_to_excel(input_bytes):
    """CSV转Excel"""
    try:
        # 尝试不同编码
        encodings = ['utf-8', 'utf-8-sig', 'gbk', 'gb2312', 'latin1']
        df = None
        
        for encoding in encodings:
            try:
                df = pd.read_csv(io.BytesIO(input_bytes), encoding=encoding)
                break
            except:
                continue
        
        if df is None:
            return None, "无法读取CSV文件"
        
        output = io.BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        return output.getvalue(), None
    except Exception as e:
        return None, str(e)


def pdf_to_images(input_bytes):
    """PDF转图片"""
    if not HAS_PDF2IMAGE:
        return None, "请安装pdf2image: pip install pdf2image（还需要安装poppler）"
    
    try:
        images = convert_from_bytes(input_bytes)
        results = []
        
        for i, img in enumerate(images):
            output = io.BytesIO()
            img.save(output, format='PNG')
            output.seek(0)
            results.append({
                'name': f'page_{i+1}.png',
                'data': output.getvalue()
            })
        
        return results, None
    except Exception as e:
        return None, str(e)


def images_to_pdf(image_files):
    """图片转PDF"""
    if not HAS_IMG2PDF:
        return None, "请安装img2pdf: pip install img2pdf"
    
    try:
        # 收集所有图片数据
        img_bytes_list = []
        
        for img_file in image_files:
            # 确保是JPEG或PNG格式
            img = Image.open(io.BytesIO(img_file))
            
            # 转换为RGB
            if img.mode in ['RGBA', 'P']:
                background = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                if img.mode == 'RGBA':
                    background.paste(img, mask=img.split()[-1])
                else:
                    background.paste(img)
                img = background
            
            output = io.BytesIO()
            img.save(output, format='JPEG', quality=95)
            output.seek(0)
            img_bytes_list.append(output.getvalue())
        
        pdf_bytes = img2pdf.convert(img_bytes_list)
        return pdf_bytes, None
    except Exception as e:
        return None, str(e)


# ==================== 主页面 ====================

def format_factory_page():
    """格式工厂主页面"""
    st.title("🏭 格式工厂")
    st.markdown("### 多功能格式转换与媒体编辑工具")
    
    # 检查FFmpeg
    ffmpeg_available = check_ffmpeg()
    
    if not ffmpeg_available:
        st.warning("⚠️ 未检测到FFmpeg。视频处理功能将不可用。请安装FFmpeg或将ffmpeg.exe放置在项目目录。")
    
    # 功能选择
    tab1, tab2, tab3, tab4 = st.tabs([
        "📷 图片处理", 
        "🎬 视频处理", 
        "🎵 音频处理",
        "📄 文档转换"
    ])
    
    # ==================== 图片处理标签页 ====================
    with tab1:
        st.header("📷 图片格式转换与压缩")
        
        if not HAS_PIL:
            st.error("❌ 图片处理需要Pillow库。请运行: pip install Pillow")
            return
        
        img_operation = st.radio(
            "选择操作",
            ["格式转换", "图片压缩", "批量转换"],
            horizontal=True,
            key="img_operation"
        )
        
        if img_operation == "格式转换":
            col1, col2 = st.columns(2)
            
            with col1:
                uploaded_img = st.file_uploader(
                    "上传图片",
                    type=['png', 'jpg', 'jpeg', 'webp', 'bmp', 'gif', 'ico', 'tiff'],
                    key="img_convert_uploader"
                )
            
            with col2:
                output_format = st.selectbox(
                    "输出格式",
                    ["PNG", "JPEG", "WEBP", "BMP", "GIF", "ICO", "TIFF"],
                    key="img_output_format"
                )
                
                if output_format in ['JPEG', 'WEBP']:
                    quality = st.slider("质量", 1, 100, 85, key="img_quality")
                else:
                    quality = 85
            
            if uploaded_img and st.button("🔄 转换", key="img_convert_btn"):
                with st.spinner("正在转换..."):
                    result, error = convert_image(
                        uploaded_img.getvalue(),
                        uploaded_img.name.split('.')[-1],
                        output_format,
                        quality
                    )
                
                if result:
                    st.success("✅ 转换成功！")
                    
                    # 显示预览
                    st.image(result, caption="转换后的图片", use_container_width=True)
                    
                    # 下载按钮
                    output_name = Path(uploaded_img.name).stem + f".{output_format.lower()}"
                    st.download_button(
                        "📥 下载",
                        data=result,
                        file_name=output_name,
                        mime=f"image/{output_format.lower()}",
                        key="img_download"
                    )
                else:
                    st.error(f"❌ 转换失败: {error}")
        
        elif img_operation == "图片压缩":
            col1, col2 = st.columns(2)
            
            with col1:
                uploaded_img = st.file_uploader(
                    "上传图片",
                    type=['png', 'jpg', 'jpeg', 'webp', 'bmp'],
                    key="img_compress_uploader"
                )
            
            with col2:
                quality = st.slider("压缩质量", 1, 100, 50, key="compress_quality")
                max_size = st.number_input("最大尺寸（像素）", 100, 10000, 1920, key="max_size")
            
            if uploaded_img and st.button("🗜️ 压缩", key="img_compress_btn"):
                original_size = len(uploaded_img.getvalue())
                
                with st.spinner("正在压缩..."):
                    result, error = compress_image(
                        uploaded_img.getvalue(),
                        quality,
                        max_size
                    )
                
                if result:
                    compressed_size = len(result)
                    reduction = (1 - compressed_size / original_size) * 100
                    
                    st.success(f"✅ 压缩成功！大小减少 {reduction:.1f}%")
                    st.info(f"原始大小: {original_size/1024:.1f} KB → 压缩后: {compressed_size/1024:.1f} KB")
                    
                    output_name = Path(uploaded_img.name).stem + "_compressed.jpg"
                    st.download_button(
                        "📥 下载",
                        data=result,
                        file_name=output_name,
                        mime="image/jpeg",
                        key="compress_download"
                    )
                else:
                    st.error(f"❌ 压缩失败: {error}")
        
        elif img_operation == "批量转换":
            uploaded_imgs = st.file_uploader(
                "上传多个图片",
                type=['png', 'jpg', 'jpeg', 'webp', 'bmp', 'gif'],
                accept_multiple_files=True,
                key="img_batch_uploader"
            )
            
            output_format = st.selectbox(
                "输出格式",
                ["PNG", "JPEG", "WEBP"],
                key="batch_output_format"
            )
            
            if uploaded_imgs and st.button("🔄 批量转换", key="batch_convert_btn"):
                results = []
                progress = st.progress(0)
                
                for i, img_file in enumerate(uploaded_imgs):
                    result, error = convert_image(
                        img_file.getvalue(),
                        img_file.name.split('.')[-1],
                        output_format,
                        85
                    )
                    
                    if result:
                        output_name = Path(img_file.name).stem + f".{output_format.lower()}"
                        results.append((output_name, result))
                    
                    progress.progress((i + 1) / len(uploaded_imgs))
                
                st.success(f"✅ 成功转换 {len(results)}/{len(uploaded_imgs)} 个文件")
                
                for name, data in results:
                    st.download_button(
                        f"📥 {name}",
                        data=data,
                        file_name=name,
                        mime=f"image/{output_format.lower()}",
                        key=f"batch_dl_{name}"
                    )
    
    # ==================== 视频处理标签页 ====================
    with tab2:
        st.header("🎬 视频处理")
        
        if not ffmpeg_available:
            st.error("❌ 视频处理需要FFmpeg。请安装FFmpeg或将ffmpeg.exe放置在项目目录。")
            return
        
        video_operation = st.radio(
            "选择操作",
            ["格式转换", "视频剪辑", "速度调整", "倒放", "提取音频", "静音", "压缩"],
            horizontal=True,
            key="video_operation"
        )
        
        uploaded_video = st.file_uploader(
            "上传视频",
            type=['mp4', 'avi', 'mkv', 'mov', 'webm', 'flv', 'wmv'],
            key="video_uploader"
        )
        
        if uploaded_video:
            # 预览原始视频
            st.markdown("##### 📹 原始视频预览")
            st.video(uploaded_video)

            # 保存到临时文件
            with tempfile.NamedTemporaryFile(delete=False, suffix=f'.{uploaded_video.name.split(".")[-1]}') as tmp:
                tmp.write(uploaded_video.getvalue())
                input_path = tmp.name
            
            if video_operation == "格式转换":
                output_format = st.selectbox(
                    "输出格式",
                    ["mp4", "avi", "mkv", "mov", "webm"],
                    key="video_output_format"
                )
                
                if st.button("🔄 转换", key="video_convert_btn"):
                    output_path = input_path.rsplit('.', 1)[0] + f'_converted.{output_format}'
                    
                    with st.spinner("正在转换视频...（可能需要几分钟）"):
                        success, msg = convert_video_format(input_path, output_path, output_format)
                    
                    if success and os.path.exists(output_path):
                        st.success("✅ 转换成功！")
                        
                        st.markdown("##### 🎬 处理后预览")
                        st.video(output_path)

                        with open(output_path, 'rb') as f:
                            st.download_button(
                                "📥 下载",
                                data=f.read(),
                                file_name=Path(uploaded_video.name).stem + f'.{output_format}',
                                mime=f"video/{output_format}",
                                key="video_dl"
                            )
                        
                        os.remove(output_path)
                    else:
                        st.error(f"❌ 转换失败: {msg}")
            
            elif video_operation == "视频剪辑":
                col1, col2 = st.columns(2)
                with col1:
                    start_time = st.text_input("开始时间 (HH:MM:SS)", "00:00:00", key="video_start")
                with col2:
                    end_time = st.text_input("结束时间 (HH:MM:SS)", "00:00:10", key="video_end")
                
                if st.button("✂️ 剪辑", key="video_trim_btn"):
                    output_path = input_path.rsplit('.', 1)[0] + '_trimmed.mp4'
                    
                    with st.spinner("正在剪辑..."):
                        success, msg = trim_video(input_path, output_path, start_time, end_time)
                    
                    if success and os.path.exists(output_path):
                        st.success("✅ 剪辑成功！")
                        
                        st.markdown("##### 🎬 处理后预览")
                        st.video(output_path)

                        with open(output_path, 'rb') as f:
                            st.download_button(
                                "📥 下载",
                                data=f.read(),
                                file_name=Path(uploaded_video.name).stem + '_trimmed.mp4',
                                mime="video/mp4",
                                key="trim_dl"
                            )
                        
                        os.remove(output_path)
                    else:
                        st.error(f"❌ 剪辑失败: {msg}")
            
            elif video_operation == "速度调整":
                speed = st.slider("速度倍率", 0.25, 4.0, 1.0, 0.25, key="video_speed")
                
                if st.button("⚡ 调整速度", key="video_speed_btn"):
                    output_path = input_path.rsplit('.', 1)[0] + f'_speed{speed}x.mp4'
                    
                    with st.spinner("正在处理..."):
                        success, msg = change_video_speed(input_path, output_path, speed)
                    
                    if success and os.path.exists(output_path):
                        st.success("✅ 处理成功！")
                        
                        st.markdown("##### 🎬 处理后预览")
                        st.video(output_path)

                        with open(output_path, 'rb') as f:
                            st.download_button(
                                "📥 下载",
                                data=f.read(),
                                file_name=Path(uploaded_video.name).stem + f'_speed{speed}x.mp4',
                                mime="video/mp4",
                                key="speed_dl"
                            )
                        
                        os.remove(output_path)
                    else:
                        st.error(f"❌ 处理失败: {msg}")
            
            elif video_operation == "倒放":
                st.warning("⚠️ 视频倒放可能需要较长时间处理")
                
                if st.button("⏪ 倒放", key="video_reverse_btn"):
                    output_path = input_path.rsplit('.', 1)[0] + '_reversed.mp4'
                    
                    with st.spinner("正在处理...（可能需要几分钟）"):
                        success, msg = reverse_video(input_path, output_path)
                    
                    if success and os.path.exists(output_path):
                        st.success("✅ 处理成功！")
                        
                        st.markdown("##### 🎬 处理后预览")
                        st.video(output_path)

                        with open(output_path, 'rb') as f:
                            st.download_button(
                                "📥 下载",
                                data=f.read(),
                                file_name=Path(uploaded_video.name).stem + '_reversed.mp4',
                                mime="video/mp4",
                                key="reverse_dl"
                            )
                        
                        os.remove(output_path)
                    else:
                        st.error(f"❌ 处理失败: {msg}")
            
            elif video_operation == "提取音频":
                if st.button("🎵 提取音频", key="extract_audio_btn"):
                    output_path = input_path.rsplit('.', 1)[0] + '.mp3'
                    
                    with st.spinner("正在提取音频..."):
                        success, msg = extract_audio_from_video(input_path, output_path)
                    
                    if success and os.path.exists(output_path):
                        st.success("✅ 提取成功！")
                        
                        st.markdown("##### 🎵 提取音频预览")
                        st.audio(output_path)

                        with open(output_path, 'rb') as f:
                            st.download_button(
                                "📥 下载MP3",
                                data=f.read(),
                                file_name=Path(uploaded_video.name).stem + '.mp3',
                                mime="audio/mp3",
                                key="extract_dl"
                            )
                        
                        os.remove(output_path)
                    else:
                        st.error(f"❌ 提取失败: {msg}")
            
            elif video_operation == "静音":
                if st.button("🔇 移除音频", key="mute_btn"):
                    output_path = input_path.rsplit('.', 1)[0] + '_muted.mp4'
                    
                    with st.spinner("正在处理..."):
                        success, msg = mute_video(input_path, output_path)
                    
                    if success and os.path.exists(output_path):
                        st.success("✅ 处理成功！")
                        
                        st.markdown("##### 🎬 处理后预览")
                        st.video(output_path)

                        with open(output_path, 'rb') as f:
                            st.download_button(
                                "📥 下载",
                                data=f.read(),
                                file_name=Path(uploaded_video.name).stem + '_muted.mp4',
                                mime="video/mp4",
                                key="mute_dl"
                            )
                        
                        os.remove(output_path)
                    else:
                        st.error(f"❌ 处理失败: {msg}")
            
            elif video_operation == "压缩":
                crf = st.slider("压缩级别 (CRF)", 18, 40, 28, help="数值越大压缩越强，质量越低")
                
                if st.button("🗜️ 压缩", key="compress_video_btn"):
                    output_path = input_path.rsplit('.', 1)[0] + '_compressed.mp4'
                    
                    with st.spinner("正在压缩...（可能需要几分钟）"):
                        success, msg = compress_video(input_path, output_path, crf)
                    
                    if success and os.path.exists(output_path):
                        original_size = os.path.getsize(input_path)
                        compressed_size = os.path.getsize(output_path)
                        reduction = (1 - compressed_size / original_size) * 100
                        
                        st.success(f"✅ 压缩成功！大小减少 {reduction:.1f}%")
                        
                        st.markdown("##### 🎬 处理后预览")
                        st.video(output_path)

                        with open(output_path, 'rb') as f:
                            st.download_button(
                                "📥 下载",
                                data=f.read(),
                                file_name=Path(uploaded_video.name).stem + '_compressed.mp4',
                                mime="video/mp4",
                                key="compress_video_dl"
                            )
                        
                        os.remove(output_path)
                    else:
                        st.error(f"❌ 压缩失败: {msg}")
            
            # 清理临时文件
            if os.path.exists(input_path):
                os.remove(input_path)
    
    # ==================== 音频处理标签页 ====================
    with tab3:
        st.header("🎵 音频处理")
        
        if not HAS_PYDUB:
            st.error("❌ 音频处理需要pydub库。请运行: pip install pydub")
            st.info("💡 pydub还需要FFmpeg支持")
            return
        
        audio_operation = st.radio(
            "选择操作",
            ["格式转换", "音频剪辑", "淡入淡出", "速度调整", "音量调整", "倒放", "合并音频"],
            horizontal=True,
            key="audio_operation"
        )
        
        if audio_operation != "合并音频":
            uploaded_audio = st.file_uploader(
                "上传音频",
                type=['mp3', 'wav', 'flac', 'aac', 'ogg', 'm4a', 'wma'],
                key="audio_uploader"
            )
            
            if uploaded_audio:
                # 预览原始音频
                st.markdown("##### 🎵 原始音频预览")
                st.audio(uploaded_audio)

                input_format = uploaded_audio.name.split('.')[-1].lower()
                
                # 显示音频信息
                try:
                    audio = AudioSegment.from_file(io.BytesIO(uploaded_audio.getvalue()), format=input_format)
                    duration_sec = len(audio) / 1000
                    st.info(f"📊 时长: {duration_sec:.2f} 秒 | 采样率: {audio.frame_rate} Hz | 声道: {audio.channels}")
                except:
                    pass
                
                if audio_operation == "格式转换":
                    output_format = st.selectbox(
                        "输出格式",
                        ["mp3", "wav", "flac", "ogg", "aac"],
                        key="audio_output_format"
                    )
                    
                    if st.button("🔄 转换", key="audio_convert_btn"):
                        with st.spinner("正在转换..."):
                            result, error = convert_audio_format(
                                uploaded_audio.getvalue(),
                                input_format,
                                output_format
                            )
                        
                        if result:
                            st.success("✅ 转换成功！")
                            
                            st.markdown("##### 🎧 处理后预览")
                            st.audio(result, format=f'audio/{output_format}')

                            output_name = Path(uploaded_audio.name).stem + f'.{output_format}'
                            st.download_button(
                                "📥 下载",
                                data=result,
                                file_name=output_name,
                                mime=f"audio/{output_format}",
                                key="audio_convert_dl"
                            )
                        else:
                            st.error(f"❌ 转换失败: {error}")
                
                elif audio_operation == "音频剪辑":
                    col1, col2 = st.columns(2)
                    with col1:
                        start_sec = st.number_input("开始时间（秒）", 0.0, float(duration_sec), 0.0, key="audio_start")
                    with col2:
                        end_sec = st.number_input("结束时间（秒）", 0.0, float(duration_sec), float(duration_sec), key="audio_end")
                    
                    if st.button("✂️ 剪辑", key="audio_trim_btn"):
                        with st.spinner("正在剪辑..."):
                            result, error = trim_audio(
                                uploaded_audio.getvalue(),
                                input_format,
                                int(start_sec * 1000),
                                int(end_sec * 1000)
                            )
                        
                        if result:
                            st.success("✅ 剪辑成功！")
                            
                            st.markdown("##### 🎧 处理后预览")
                            st.audio(result, format=f'audio/{input_format}')

                            output_name = Path(uploaded_audio.name).stem + f'_trimmed.{input_format}'
                            st.download_button(
                                "📥 下载",
                                data=result,
                                file_name=output_name,
                                mime=f"audio/{input_format}",
                                key="audio_trim_dl"
                            )
                        else:
                            st.error(f"❌ 剪辑失败: {error}")
                
                elif audio_operation == "淡入淡出":
                    col1, col2 = st.columns(2)
                    with col1:
                        fade_in = st.number_input("淡入时长（秒）", 0.0, 30.0, 2.0, key="fade_in")
                    with col2:
                        fade_out = st.number_input("淡出时长（秒）", 0.0, 30.0, 2.0, key="fade_out")
                    
                    if st.button("🎚️ 应用淡入淡出", key="fade_btn"):
                        with st.spinner("正在处理..."):
                            result, error = fade_audio(
                                uploaded_audio.getvalue(),
                                input_format,
                                int(fade_in * 1000),
                                int(fade_out * 1000)
                            )
                        
                        if result:
                            st.success("✅ 处理成功！")
                            
                            st.markdown("##### 🎧 处理后预览")
                            st.audio(result, format=f'audio/{input_format}')

                            output_name = Path(uploaded_audio.name).stem + f'_faded.{input_format}'
                            st.download_button(
                                "📥 下载",
                                data=result,
                                file_name=output_name,
                                mime=f"audio/{input_format}",
                                key="fade_dl"
                            )
                        else:
                            st.error(f"❌ 处理失败: {error}")
                
                elif audio_operation == "速度调整":
                    speed = st.slider("速度倍率", 0.5, 2.0, 1.0, 0.1, key="audio_speed")
                    preserve_pitch = st.checkbox("保留音调", value=True, key="preserve_pitch")
                    
                    if st.button("⚡ 调整速度", key="audio_speed_btn"):
                        with st.spinner("正在处理..."):
                            result, error = adjust_audio_speed(
                                uploaded_audio.getvalue(),
                                input_format,
                                speed,
                                preserve_pitch
                            )
                        
                        if result:
                            st.success("✅ 处理成功！")
                            
                            st.markdown("##### 🎧 处理后预览")
                            st.audio(result, format=f'audio/{input_format}')

                            output_name = Path(uploaded_audio.name).stem + f'_speed{speed}x.{input_format}'
                            st.download_button(
                                "📥 下载",
                                data=result,
                                file_name=output_name,
                                mime=f"audio/{input_format}",
                                key="audio_speed_dl"
                            )
                        else:
                            st.error(f"❌ 处理失败: {error}")
                
                elif audio_operation == "音量调整":
                    volume_db = st.slider("音量调整 (dB)", -20, 20, 0, key="volume_db")
                    st.info(f"{'增大' if volume_db > 0 else '减小' if volume_db < 0 else '保持'} {abs(volume_db)} 分贝")
                    
                    if st.button("🔊 调整音量", key="volume_btn"):
                        with st.spinner("正在处理..."):
                            result, error = adjust_audio_volume(
                                uploaded_audio.getvalue(),
                                input_format,
                                volume_db
                            )
                        
                        if result:
                            st.success("✅ 处理成功！")
                            
                            st.markdown("##### 🎧 处理后预览")
                            st.audio(result, format=f'audio/{input_format}')

                            output_name = Path(uploaded_audio.name).stem + f'_vol{volume_db}db.{input_format}'
                            st.download_button(
                                "📥 下载",
                                data=result,
                                file_name=output_name,
                                mime=f"audio/{input_format}",
                                key="volume_dl"
                            )
                        else:
                            st.error(f"❌ 处理失败: {error}")
                
                elif audio_operation == "倒放":
                    if st.button("⏪ 倒放", key="audio_reverse_btn"):
                        with st.spinner("正在处理..."):
                            result, error = reverse_audio(
                                uploaded_audio.getvalue(),
                                input_format
                            )
                        
                        if result:
                            st.success("✅ 处理成功！")
                            
                            st.markdown("##### 🎧 处理后预览")
                            st.audio(result, format=f'audio/{input_format}')

                            output_name = Path(uploaded_audio.name).stem + f'_reversed.{input_format}'
                            st.download_button(
                                "📥 下载",
                                data=result,
                                file_name=output_name,
                                mime=f"audio/{input_format}",
                                key="audio_reverse_dl"
                            )
                        else:
                            st.error(f"❌ 处理失败: {error}")
        
        else:  # 合并音频
            uploaded_audios = st.file_uploader(
                "上传多个音频文件",
                type=['mp3', 'wav', 'flac', 'ogg'],
                accept_multiple_files=True,
                key="merge_audio_uploader"
            )
            
            output_format = st.selectbox(
                "输出格式",
                ["mp3", "wav", "flac", "ogg"],
                key="merge_output_format"
            )
            
            if uploaded_audios and len(uploaded_audios) >= 2:
                st.info(f"已选择 {len(uploaded_audios)} 个文件")
                
                if st.button("🔗 合并", key="merge_btn"):
                    audio_files = []
                    for f in uploaded_audios:
                        audio_files.append({
                            'bytes': f.getvalue(),
                            'format': f.name.split('.')[-1]
                        })
                    
                    with st.spinner("正在合并..."):
                        result, error = merge_audios(audio_files, output_format)
                    
                    if result:
                        st.success("✅ 合并成功！")
                        
                        st.markdown("##### 🎧 处理后预览")
                        st.audio(result, format=f'audio/{output_format}')

                        st.download_button(
                            "📥 下载",
                            data=result,
                            file_name=f"merged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{output_format}",
                            mime=f"audio/{output_format}",
                            key="merge_dl"
                        )
                    else:
                        st.error(f"❌ 合并失败: {error}")
            elif uploaded_audios:
                st.warning("请至少上传2个音频文件")
    
    # ==================== 文档转换标签页 ====================
    with tab4:
        st.header("📄 文档格式转换")
        
        doc_operation = st.radio(
            "选择操作",
            ["Excel → CSV", "CSV → Excel", "PDF → 图片", "图片 → PDF"],
            horizontal=True,
            key="doc_operation"
        )
        
        if doc_operation == "Excel → CSV":
            uploaded_file = st.file_uploader(
                "上传Excel文件",
                type=['xlsx', 'xls'],
                key="excel_to_csv_uploader"
            )
            
            if uploaded_file:
                # 读取工作表名称
                try:
                    xl = pd.ExcelFile(io.BytesIO(uploaded_file.getvalue()))
                    sheet_names = xl.sheet_names
                    
                    if len(sheet_names) > 1:
                        selected_sheet = st.selectbox("选择工作表", sheet_names, key="sheet_select")
                    else:
                        selected_sheet = sheet_names[0]
                except:
                    selected_sheet = None
                
                if st.button("🔄 转换", key="excel_csv_btn"):
                    with st.spinner("正在转换..."):
                        result, error = excel_to_csv(uploaded_file.getvalue(), selected_sheet)
                    
                    if result:
                        st.success("✅ 转换成功！")
                        
                        output_name = Path(uploaded_file.name).stem + '.csv'
                        st.download_button(
                            "📥 下载CSV",
                            data=result,
                            file_name=output_name,
                            mime="text/csv",
                            key="excel_csv_dl"
                        )
                    else:
                        st.error(f"❌ 转换失败: {error}")
        
        elif doc_operation == "CSV → Excel":
            uploaded_file = st.file_uploader(
                "上传CSV文件",
                type=['csv'],
                key="csv_to_excel_uploader"
            )
            
            if uploaded_file and st.button("🔄 转换", key="csv_excel_btn"):
                with st.spinner("正在转换..."):
                    result, error = csv_to_excel(uploaded_file.getvalue())
                
                if result:
                    st.success("✅ 转换成功！")
                    
                    output_name = Path(uploaded_file.name).stem + '.xlsx'
                    st.download_button(
                        "📥 下载Excel",
                        data=result,
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="csv_excel_dl"
                    )
                else:
                    st.error(f"❌ 转换失败: {error}")
        
        elif doc_operation == "PDF → 图片":
            if not HAS_PDF2IMAGE:
                st.error("❌ 需要安装pdf2image: pip install pdf2image")
                st.info("💡 还需要安装poppler: https://github.com/osber/poppler-windows/releases")
                return
            
            uploaded_file = st.file_uploader(
                "上传PDF文件",
                type=['pdf'],
                key="pdf_to_img_uploader"
            )
            
            if uploaded_file and st.button("🔄 转换", key="pdf_img_btn"):
                with st.spinner("正在转换..."):
                    result, error = pdf_to_images(uploaded_file.getvalue())
                
                if result:
                    st.success(f"✅ 成功转换 {len(result)} 页！")
                    
                    for page in result:
                        col1, col2 = st.columns([3, 1])
                        with col1:
                            st.image(page['data'], caption=page['name'], use_container_width=True)
                        with col2:
                            st.download_button(
                                f"📥 {page['name']}",
                                data=page['data'],
                                file_name=page['name'],
                                mime="image/png",
                                key=f"pdf_img_dl_{page['name']}"
                            )
                else:
                    st.error(f"❌ 转换失败: {error}")
        
        elif doc_operation == "图片 → PDF":
            if not HAS_IMG2PDF:
                st.error("❌ 需要安装img2pdf: pip install img2pdf")
                return
            
            uploaded_files = st.file_uploader(
                "上传图片文件（可多选）",
                type=['png', 'jpg', 'jpeg'],
                accept_multiple_files=True,
                key="img_to_pdf_uploader"
            )
            
            if uploaded_files and st.button("🔄 转换", key="img_pdf_btn"):
                with st.spinner("正在转换..."):
                    image_bytes = [f.getvalue() for f in uploaded_files]
                    result, error = images_to_pdf(image_bytes)
                
                if result:
                    st.success("✅ 转换成功！")
                    
                    st.download_button(
                        "📥 下载PDF",
                        data=result,
                        file_name=f"images_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                        mime="application/pdf",
                        key="img_pdf_dl"
                    )
                else:
                    st.error(f"❌ 转换失败: {error}")
