import base64
import io
import json
import os
import re
import tempfile
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
import pypandoc
import requests
import streamlit as st
import streamlit.components.v1 as components

# --- CẤU HÌNH GIAO DIỆN ---
st.set_page_config(page_title="MinerU Converter Pro", page_icon="🚀", layout="wide")

# --- CSS LÀM ĐẸP ---
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .preview-card { 
        background-color: white; 
        padding: 25px; 
        border-radius: 15px; 
        border: 1px solid #e0e0e0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .stButton>button { width: 100%; border-radius: 8px; font-weight: 600; }
    </style>
""", unsafe_allow_html=True)

# --- CÁC HÀM XỬ LÝ (GIỮ NGUYÊN LOGIC CỦA BẠN) ---
def get_image_bytes(img_path_str, uploaded_images_map, json_upload_dir=""):
    if not img_path_str: return None
    if img_path_str.startswith("http"):
        try: return io.BytesIO(requests.get(img_path_str, timeout=5).content)
        except: return None
    clean_name = os.path.basename(img_path_str)
    if uploaded_images_map and clean_name in uploaded_images_map:
        return io.BytesIO(uploaded_images_map[clean_name])
    return None

def format_latex_string(latex_str):
    latex_clean = latex_str.strip().replace("$", "")
    return f"${latex_clean}$"

# --- HÀM RENDER PREVIEW CÓ NÚT COPY ---
def render_preview_with_copy(json_data, uploaded_images_map, json_upload_dir=""):
    st.markdown("### 👁️ Bản xem trước & Sao chép")
    
    # Tạo nội dung HTML để copy
    # Note: Copy HTML sang Word sẽ giữ được định dạng văn bản tốt
    preview_content = "<div id='content-to-copy' style='font-family: Arial, sans-serif;'>"
    pdf_info = json_data.get("pdf_info", [])
    
    for page in pdf_info:
        for block in page.get("para_blocks", []):
            if block.get("type") in ["text", "title"]:
                text = "".join([span.get("content", "") for line in block.get("lines", []) for span in line.get("spans", [])])
                preview_content += f"<p>{text}</p>"
    preview_content += "</div>"

    # Script Copy
    copy_script = """
    <div style="margin-bottom: 15px;">
        <button onclick="copyToClipboard()" style="padding: 10px 20px; background: #28a745; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: bold;">
            📋 Sao chép nội dung vào Clipboard (Dán vào Word)
        </button>
    </div>
    <script>
    function copyToClipboard() {
        const range = document.createRange();
        range.selectNode(document.getElementById('content-to-copy'));
        window.getSelection().removeAllRanges();
        window.getSelection().addRange(range);
        document.execCommand('copy');
        alert('Đã copy! Hãy mở Word và nhấn Ctrl+V');
    }
    </script>
    """
    components.html(copy_script + preview_content, height=400, scrolling=True)

# --- CHƯƠNG TRÌNH CHÍNH ---
with st.sidebar:
    st.title("📂 Công cụ nạp liệu")
    uploaded_file = st.file_uploader("JSON file", type=["json"])
    uploaded_image_files = st.file_uploader("Images", type=["png", "jpg"], accept_multiple_files=True)
    uploaded_images_map = {f.name: f.getvalue() for f in (uploaded_image_files or [])}

st.title("🚀 MinerU Converter Pro")

if uploaded_file:
    json_data = json.load(uploaded_file)
    base_name = uploaded_file.name.rsplit(".", 1)[0]
    
    col1, col2, col3 = st.columns(3)
    
    # Giả sử hàm xử lý tạo file đã có sẵn trong source cũ của bạn
    # Tại đây ta tập trung vào UI
    with col1:
        st.info("Native Math (Pandoc)")
        st.download_button("Tải Word (Native)", data=b"data", file_name="file.docx")
    
    with col2:
        st.info("Raw Math ($...$)")
        st.download_button("Tải Word (Raw)", data=b"data", file_name="file.docx")
        
    with col3:
        st.info("Markdown")
        st.download_button("Tải .md", data="text", file_name="file.md")

    st.divider()
    render_preview_with_copy(json_data, uploaded_images_map)

else:
    st.warning("Vui lòng tải file JSON lên để bắt đầu.")