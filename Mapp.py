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
import requests
import streamlit as st
import streamlit.components.v1 as components

# --- CẤU HÌNH GIAO DIỆN TRANG ---
st.set_page_config(
    page_title="MinerU JSON Math Preview Pro", 
    page_icon="📐", 
    layout="wide"
)

# --- CSS TÙY CHỈNH GIAO DIỆN ---
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .preview-box {
        background-color: #ffffff;
        padding: 30px;
        border-radius: 12px;
        border: 1px solid #cbd5e0;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
        margin-top: 15px;
    }
    </style>
""", unsafe_allow_html=True)


# --- 1. CÁC HÀM XỬ LÝ DÙNG CHUNG ---

def get_image_bytes(img_path_str, uploaded_images_map, json_upload_dir=""):
    """Lấy bytes ảnh từ upload thủ công, URL hoặc thư mục images cùng cấp."""
    if not img_path_str:
        return None
    
    if img_path_str.startswith("http://") or img_path_str.startswith("https://"):
        try:
            response = requests.get(img_path_str, timeout=10)
            if response.status_code == 200:
                return io.BytesIO(response.content)
        except Exception:
            pass
        return None

    clean_name = os.path.basename(img_path_str)
    
    if uploaded_images_map and clean_name in uploaded_images_map:
        return io.BytesIO(uploaded_images_map[clean_name])
    
    if json_upload_dir:
        auto_path = os.path.join(json_upload_dir, "images", clean_name)
        if os.path.exists(auto_path):
            with open(auto_path, "rb") as f:
                return io.BytesIO(f.read())
                
    fallback_paths = [
        os.path.join("images", clean_name),
        os.path.join(os.getcwd(), "images", clean_name),
        clean_name
    ]
    for path in fallback_paths:
        if os.path.exists(path):
            with open(path, "rb") as f:
                return io.BytesIO(f.read())
                
    return None

def format_latex_string(latex_str):
    """Làm sạch chuỗi và bọc công thức trong cặp dấu $...$ chuẩn KaTeX hiển thị công thức đẹp"""
    if not latex_str:
        return ""
    latex_clean = latex_str.strip()
    if latex_clean.startswith("$") and latex_clean.endswith("$"):
        latex_clean = latex_clean[1:-1].strip()
    return f"${latex_clean}$"


# --- 2. RENDER PREVIEW TẬP TRUNG HOÀN TOÀN VÀO CÔNG THỨC TOÁN & SAO CHÉP ---

def render_pure_math_preview(json_data, uploaded_images_map, json_upload_dir=""):
    """Duyệt JSON, render toàn bộ công thức toán học chuẩn KaTeX/Equation và tích hợp nút Copy"""
    
    preview_inner_html = '<div id="content-to-copy" style="font-family: Arial, sans-serif; line-height: 1.8; color: #2d3748; font-size: 16px;">'
    
    pdf_info = json_data.get("pdf_info", [])
    for page in pdf_info:
        for block in page.get("para_blocks", []):
            b_type = block.get("type")

            if b_type in ["text", "title"]:
                p_text = ""
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        span_type = span.get("type")
                        content = span.get("content", "")
                        if span_type == "text":
                            if re.match(r"^Bài\s+\d+", content.strip()):
                                p_text += f"<b style='color: #1a202c;'>{content}</b>"
                            else:
                                p_text += content
                        elif span_type == "inline_equation":
                            clean_c = content.strip().replace("$", "")
                            # Bọc công thức để Streamlit/HTML hiểu định dạng toán học rõ ràng
                            p_text += f" <span style='font-family: Cambria Math, Times New Roman, serif;'>${clean_c}$</span> "
                if p_text.strip():
                    preview_inner_html += f"<p style='margin-bottom: 12px;'>{p_text.strip()}</p>"

            elif b_type in ["image", "chart", "figure"]:
                paths_to_check = []
                if block.get("image_path"):
                    paths_to_check.append(block.get("image_path"))
                for sub_b in block.get("blocks", []):
                    if sub_b.get("image_path"):
                        paths_to_check.append(sub_b.get("image_path"))
                    for line in sub_b.get("lines", []):
                        for span in line.get("spans", []):
                            if span.get("image_path"):
                                paths_to_check.append(span.get("image_path"))
                
                for img_path_str in paths_to_check:
                    img_stream = get_image_bytes(img_path_str, uploaded_images_map, json_upload_dir)
                    if img_stream:
                        encoded = base64.b64encode(img_stream.getvalue()).decode("utf-8")
                        preview_inner_html += f'<div style="text-align: center; margin: 20px 0;"><img src="data:image/png;base64,{encoded}" style="max-width: 450px; border-radius: 8px; border: 1px solid #e2e8f0;" /></div>'

            elif b_type == "table":
                for sub_b in block.get("blocks", []):
                    for line in sub_b.get("lines", []):
                        for span in line.get("spans", []):
                            table_html = span.get("html")
                            if table_html:
                                soup = BeautifulSoup(table_html, "html.parser")
                                for eq_tag in soup.find_all("eq"):
                                    eq_text = eq_tag.get_text().strip().replace("$", "")
                                    eq_tag.string = f"${eq_text}$"
                                
                                for img_tag in soup.find_all("img"):
                                    img_src = img_tag.get("src")
                                    if img_src:
                                        img_stream = get_image_bytes(img_src, uploaded_images_map, json_upload_dir)
                                        if img_stream:
                                            encoded = base64.b64encode(img_stream.getvalue()).decode("utf-8")
                                            img_tag['src'] = f"data:image/png;base64,{encoded}"
                                            img_tag['width'] = "150"
                                preview_inner_html += f'<div style="margin: 20px 0; overflow-x: auto;">{str(soup)}</div>'

    preview_inner_html += '</div>'

    # Component chứa nút Sao chép và khung hiển thị nội dung tích hợp KaTeX qua CDN
    copier_component = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/katex.min.css">
        <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/katex.min.js"></script>
        <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/contrib/auto-render.min.js" 
            onload="renderMathInElement(document.body);"></script>
        <style>
            body {{ font-family: Arial, sans-serif; padding: 10px; background-color: #ffffff; }}
            .btn-copy {{
                padding: 10px 20px;
                background-color: #2b6cb0;
                color: white;
                border: none;
                border-radius: 6px;
                cursor: pointer;
                font-weight: bold;
                font-size: 14px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
                margin-bottom: 15px;
            }}
            .btn-copy:hover {{ background-color: #2c5282; }}
            #copy-status {{ margin-left: 10px; color: #2f855a; font-weight: bold; font-size: 13px; display: none; }}
            .preview-card {{
                background-color: #ffffff;
                padding: 30px;
                border-radius: 10px;
                border: 1px solid #cbd5e0;
                max-height: 600px;
                overflow-y: auto;
            }}
        </style>
    </head>
    <body>
        <div>
            <button class="btn-copy" onclick="copyContentToClipboard()">📋 Sao chép nội dung (Dán thẳng vào Word)</button>
            <span id="copy-status">✔ Đã sao chép thành công!</span>
        </div>
        
        <div class="preview-card">
            {preview_inner_html}
        </div>

        <script>
        function copyContentToClipboard() {{
            const range = document.createRange();
            range.selectNode(document.getElementById('content-to-copy'));
            window.getSelection().removeAllRanges();
            window.getSelection().addRange(range);
            
            try {{
                document.execCommand('copy');
                const status = document.getElementById('copy-status');
                status.style.display = 'inline';
                setTimeout(() => {{ status.style.display = 'none'; }}, 3000);
            }} catch (err) {{
                alert('Không thể sao chép tự động!');
            }}
            window.getSelection().removeAllRanges();
        }}
        </script>
    </body>
    </html>
    """
    
    st.markdown("### 👁️ Bản xem trước Công thức Toán học (Equation & KaTeX)")
    components.html(copier_component, height=680, scrolling=False)


# --- 3. GIAO DIỆN STREAMLIT CHÍNH ---

st.markdown("<h1 style='color: #1a202c;'>📐 MinerU Math Preview & Copy Pro</h1>", unsafe_allow_html=True)
st.write("Tải lên file JSON kết quả từ MinerU để xem trước trực quan toàn bộ định dạng văn bản, hình ảnh, bảng biểu và đặc biệt là **Công thức toán học (Equation)** sắc nét.")

st.markdown("---")
col_up1, col_up2 = st.columns(2)
with col_up1:
    uploaded_file = st.file_uploader("📥 Chọn file JSON từ máy tính", type=["json"])
with col_up2:
    uploaded_image_files = st.file_uploader("🖼️ Tải lên thư mục ảnh đi kèm (nếu có)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

uploaded_images_map = {}
if uploaded_image_files:
    for img_file in uploaded_image_files:
        uploaded_images_map[img_file.name] = img_file.getvalue()
    st.success(f"Đã nạp thành công {len(uploaded_image_files)} file ảnh!", icon="✅")

st.markdown("---")

# --- XỬ LÝ HIỂN THỊ KHI CÓ FILE JSON ---
if uploaded_file is not None:
    try:
        json_data = json.load(uploaded_file)
        json_upload_dir = os.getcwd()
        
        st.success(f"Đã tải file thành công: **{uploaded_file.name}**", icon="🚀")
        st.divider()

        # Gọi trực tiếp hàm hiển thị xem trước công thức toán học tối ưu
        render_pure_math_preview(json_data, uploaded_images_map, json_upload_dir)

    except Exception as e:
        st.error(f"Đã xảy ra lỗi khi xử lý file JSON: {e}")
else:
    st.info("👆 Vui lòng bấm vào nút **Browse files** ở trên để tải lên file **JSON** kết quả từ MinerU.")