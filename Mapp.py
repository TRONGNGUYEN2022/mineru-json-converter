import base64
import io
import json
import os
import re
from bs4 import BeautifulSoup
import requests
import streamlit as st
import streamlit.components.v1 as components

# --- CẤU HÌNH GIAO DIỆN TRANG ---
st.set_page_config(
    page_title="MinerU Math Equation Preview Pro", 
    page_icon="📐", 
    layout="wide"
)

# --- CSS TÙY CHỈNH GIAO DIỆN ---
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
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

def clean_and_wrap_latex(latex_str):
    """Làm sạch chuỗi LaTeX và bọc trong cặp dấu $...$ chuẩn KaTeX hiển thị công thức"""
    if not latex_str:
        return ""
    clean_str = latex_str.strip()
    if clean_str.startswith("$") and clean_str.endswith("$"):
        clean_str = clean_str[1:-1].strip()
    return f"${clean_str}$"


# --- 2. RENDER PREVIEW TẬP TRUNG HOÀN TOÀN VÀO CÔNG THỨC TOÁN & LƯU WORD ---

def render_pure_math_preview(json_data, uploaded_images_map, json_upload_dir="", file_name="document"):
    """Duyệt JSON, render công thức KaTeX và thêm nút Lưu nhanh ra Word / Sao chép"""
    
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
                            latex_formatted = clean_and_wrap_latex(content)
                            p_text += f" {latex_formatted} "
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
                                    eq_text = eq_tag.get_text()
                                    eq_tag.string = clean_and_wrap_latex(eq_text)
                                
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

    # Tích hợp thư viện html-docx-js qua CDN để người dùng bấm nút là tải xuống file .docx trực tiếp từ Preview
    copier_component = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/katex.min.css">
        <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/katex.min.js"></script>
        <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/contrib/auto-render.min.js" 
            onload="renderMathInElement(document.body, {{
                delimiters: [
                    {{left: '$$', right: '$$', display: true}},
                    {{left: '$', right: '$', display: false}}
                ],
                throwOnError: false
            }});"></script>
        <!-- Thư viện chuyển đổi HTML sang Word (.docx) trực tiếp trên trình duyệt -->
        <script src="https://cdn.jsdelivr.net/npm/html-docx-js@0.3.1/dist/html-docx.js"></script>
        <style>
            body {{ font-family: Arial, sans-serif; padding: 10px; background-color: #ffffff; }}
            .btn-action {{
                padding: 10px 20px;
                color: white;
                border: none;
                border-radius: 6px;
                cursor: pointer;
                font-weight: bold;
                font-size: 14px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
                margin-right: 10px;
                margin-bottom: 15px;
            }}
            .btn-copy {{ background-color: #2b6cb0; }}
            .btn-copy:hover {{ background-color: #2c5282; }}
            .btn-word {{ background-color: #2f855a; }}
            .btn-word:hover {{ background-color: #276749; }}
            #status-msg {{ margin-left: 10px; color: #2f855a; font-weight: bold; font-size: 13px; display: none; }}
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
            <button class="btn-action btn-copy" onclick="copyContentToClipboard()">📋 Sao chép nhanh (Dán vào Word)</button>
            <button class="btn-action btn-word" onclick="saveAsWordDocx()">💾 Lưu thành file Word (.docx)</button>
            <span id="status-msg">✔ Thao tác thành công!</span>
        </div>
        
        <div class="preview-card" id="preview-box">
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
                showStatus("Đã sao chép vào bộ nhớ tạm! Mở Word và nhấn Ctrl+V");
            }} catch (err) {{
                alert('Không thể sao chép tự động!');
            }}
            window.getSelection().removeAllRanges();
        }}

        function saveAsWordDocx() {{
            const contentHTML = document.getElementById('preview-box').innerHTML;
            const converted = htmlDocx.asBlob('<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>' + contentHTML + '</body></html>');
            
            const link = document.createElement('a');
            link.href = URL.createObjectURL(converted);
            link.download = "{file_name}_preview.docx";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            showStatus("Đã tải xuống file Word thành công!");
        }}

        function showStatus(msg) {{
            const status = document.getElementById('status-msg');
            status.innerText = "✔ " + msg;
            status.style.display = 'inline';
            setTimeout(() => {{ status.style.display = 'none'; }}, 4000);
        }}
        </script>
    </body>
    </html>
    """
    
    st.markdown("### 👁️ Bản xem trước Công thức Toán học & Tùy chọn lưu nhanh")
    components.html(copier_component, height=700, scrolling=False)


# --- 3. GIAO DIỆN STREAMLIT CHÍNH ---

st.markdown("<h1 style='color: #1a202c;'>📐 MinerU Math Equation Preview Pro</h1>", unsafe_allow_html=True)
st.write("Tải lên file JSON kết quả từ MinerU để xem trước văn bản, hình ảnh, bảng biểu, công thức toán học sắc nét và hỗ trợ lưu file Word ngay lập tức.")

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
        base_name = uploaded_file.name.rsplit(".", 1)[0]
        
        st.success(f"Đã tải file thành công: **{uploaded_file.name}**", icon="🚀")
        st.divider()

        render_pure_math_preview(json_data, uploaded_images_map, json_upload_dir, file_name=base_name)

    except Exception as e:
        st.error(f"Đã xảy ra lỗi khi xử lý file JSON: {e}")
else:
    st.info("👆 Vui lòng bấm vào nút **Browse files** ở trên để tải lên file **JSON** kết quả từ MinerU.")