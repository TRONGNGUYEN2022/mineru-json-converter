import io
import json
import re
import tempfile
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
import pypandoc
import requests
import streamlit as st

# --- 1. CÁC HÀM XỬ LÝ DÙNG CHUNG ---


def download_image_stream(url):
    """Tải ảnh từ URL về bộ nhớ dưới dạng BytesIO stream."""
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return io.BytesIO(response.content)
    except Exception:
        pass
    return None


def format_latex_string(latex_str):
    """Làm sạch chuỗi và bọc công thức trong cặp dấu $...$"""
    if not latex_str:
        return ""
    latex_clean = latex_str.strip()
    if latex_clean.startswith("$") and latex_clean.endswith("$"):
        latex_clean = latex_clean[1:-1].strip()
    return f"${latex_clean}$"


# --- 2. XỬ LÝ CHUYỂN JSON SANG MARKDOWN ---


def process_html_table_md(table_html):
    """Chuyển đổi bảng HTML sang dạng bảng Markdown."""
    soup = BeautifulSoup(table_html, "html.parser")
    rows = soup.find_all("tr")
    if not rows:
        return ""

    md_table = []
    headers_parsed = False

    for tr in rows:
        cells = tr.find_all(["td", "th"])
        row_content = []
        for cell in cells:
            cell_text = ""
            for node in cell.children:
                if node.name == "eq":
                    cell_text += f" {format_latex_string(node.get_text())} "
                elif node.name == "img":
                    img_src = node.get("src")
                    if img_src:
                        cell_text += f" ![Image]({img_src}) "
                elif node.name is None:
                    cell_text += str(node).strip()
            row_content.append(cell_text.strip().replace("\n", " "))

        md_row = "| " + " | ".join(row_content) + " |"
        md_table.append(md_row)

        if not headers_parsed:
            separator = "| " + " | ".join(["---"] * len(row_content)) + " |"
            md_table.append(separator)
            headers_parsed = True

    return "\n".join(md_table) + "\n\n"


def convert_json_to_markdown(json_data):
    """Quét dữ liệu JSON MinerU và xuất ra chuỗi Markdown."""
    md_lines = []
    pdf_info = json_data.get("pdf_info", [])

    for page in pdf_info:
        para_blocks = page.get("para_blocks", [])

        for block in para_blocks:
            b_type = block.get("type")

            if b_type in ["text", "title"]:
                p_text = ""
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        span_type = span.get("type")
                        content = span.get("content", "")

                        if span_type == "text":
                            if re.match(r"^Bài\s+\d+", content.strip()):
                                p_text += f"**{content}**"
                            else:
                                p_text += content
                        elif span_type == "inline_equation":
                            p_text += f" {format_latex_string(content)} "

                if p_text.strip():
                    md_lines.append(p_text.strip() + "\n\n")

            elif b_type in ["image", "chart"]:
                for sub_b in block.get("blocks", []):
                    for line in sub_b.get("lines", []):
                        for span in line.get("spans", []):
                            img_path = span.get("image_path")
                            if img_path:
                                md_lines.append(f"![Hình ảnh]({img_path})\n\n")

            elif b_type == "table":
                for sub_b in block.get("blocks", []):
                    for line in sub_b.get("lines", []):
                        for span in line.get("spans", []):
                            table_html = span.get("html")
                            if table_html:
                                md_table_str = process_html_table_md(
                                    table_html
                                )
                                md_lines.append(md_table_str)

    return "".join(md_lines)


# --- 3. DẠNG 1: PANDOC (WORD NATIVE EQUATION) ---


def convert_md_to_docx_via_pandoc(md_text):
    """Chuyển Markdown -> Word Native Equation bằng Pandoc."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_file:
        output_docx_path = tmp_file.name

    pypandoc.convert_text(
        source=md_text,
        format="markdown",
        to="docx",
        outputfile=output_docx_path,
        extra_args=["--mathjax"],
    )

    with open(output_docx_path, "rb") as f:
        return f.read()


# --- 4. DẠNG 2: RAW WORD (GIỮ NGUYÊN $...$ + CHÈN HÌNH ĐẦY ĐỦ) ---


def convert_json_to_docx_raw_bytes(json_data):
    """Tạo Word giữ nguyên chuỗi $...$ chưa render, tự động chèn đầy đủ hình
    ảnh."""
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    pdf_info = json_data.get("pdf_info", [])
    for page in pdf_info:
        for block in page.get("para_blocks", []):
            b_type = block.get("type")

            # 1. Khối Văn bản & Tiêu đề
            if b_type in ["text", "title"]:
                p = doc.add_paragraph()
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        span_type = span.get("type")
                        content = span.get("content", "")

                        if span_type == "text":
                            if re.match(r"^Bài\s+\d+", content.strip()):
                                run = p.add_run(content)
                                run.bold = True
                            else:
                                p.add_run(content)
                        elif span_type == "inline_equation":
                            run = p.add_run(
                                f" {format_latex_string(content)} "
                            )
                            run.font.name = "Cambria Math"

            # 2. Khối Hình ảnh & Biểu đồ
            elif b_type in ["image", "chart"]:
                for sub_b in block.get("blocks", []):
                    for line in sub_b.get("lines", []):
                        for span in line.get("spans", []):
                            img_path = span.get("image_path")
                            if img_path:
                                img_stream = download_image_stream(img_path)
                                if img_stream:
                                    p = doc.add_paragraph()
                                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                    p.add_run().add_picture(
                                        img_stream, width=Inches(3.5)
                                    )

            # 3. Khối Bảng (Đáp án/Hướng dẫn chấm)
            elif b_type == "table":
                for sub_b in block.get("blocks", []):
                    for line in sub_b.get("lines", []):
                        for span in line.get("spans", []):
                            table_html = span.get("html")
                            if table_html:
                                soup = BeautifulSoup(table_html, "html.parser")
                                rows = soup.find_all("tr")
                                if rows:
                                    doc.add_paragraph()  # Cách dòng
                                    num_cols = max(
                                        len(r.find_all(["td", "th"]))
                                        for r in rows
                                    )
                                    w_tbl = doc.add_table(
                                        rows=len(rows), cols=num_cols
                                    )
                                    w_tbl.style = "Table Grid"

                                    for r_idx, tr in enumerate(rows):
                                        for c_idx, cell in enumerate(
                                            tr.find_all(["td", "th"])
                                        ):
                                            if c_idx >= num_cols:
                                                break
                                            cp = w_tbl.cell(
                                                r_idx, c_idx
                                            ).paragraphs[0]

                                            for node in cell.children:
                                                if node.name == "eq":
                                                    r = cp.add_run(
                                                        f" {format_latex_string(node.get_text())} "
                                                    )
                                                    r.font.name = "Cambria Math"
                                                elif node.name == "img":
                                                    img_src = node.get("src")
                                                    if img_src:
                                                        img_stream = download_image_stream(
                                                            img_src
                                                        )
                                                        if img_stream:
                                                            cp.add_run().add_picture(
                                                                img_stream,
                                                                width=Inches(
                                                                    2.5
                                                                ),
                                                            )
                                                elif node.name is None:
                                                    cp.add_run(
                                                        str(node).strip()
                                                    )

    docx_io = io.BytesIO()
    doc.save(docx_io)
    docx_io.seek(0)
    return docx_io.getvalue()


# --- 5. GIAO DIỆN STREAMLIT ---

st.set_page_config(
    page_title="MinerU JSON Converter", page_icon="🚀", layout="wide"
)

st.title("🚀 Chuyển đổi MinerU JSON Đa định dạng")
st.write(
    "Tải lên file JSON để xuất ra file Word dạng **Công thức chuẩn (Pandoc)**, **Công thức thô (`$...$`)**, hoặc **Markdown**."
)

uploaded_file = st.file_uploader("Chọn file JSON từ máy tính", type=["json"])

if uploaded_file is not None:
    try:
        json_data = json.load(uploaded_file)
        base_name = uploaded_file.name.rsplit(".", 1)[0]
        st.success(f"Đã nạp file thành công: **{uploaded_file.name}**", icon="✅")

        # 1. Tạo chuỗi Markdown chung
        md_content = convert_json_to_markdown(json_data)

        # 2. Giao diện 3 cột tải file
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown("### 1. Word (Native Math)")
            st.caption("Công thức toán được Pandoc render chuẩn Word Equation")
            with st.spinner("Pandoc đang biên dịch..."):
                docx_pandoc_bytes = convert_md_to_docx_via_pandoc(md_content)

            st.download_button(
                label="📥 Tải Word (Native Equation)",
                data=docx_pandoc_bytes,
                file_name=f"{base_name}_NativeMath.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
            )

        with col2:
            st.markdown("### 2. Word (Giữ $...$)")
            st.caption(
                "Đầy đủ ảnh, giữ nguyên dạng `$latex$` thích hợp dùng MathType"
            )
            docx_raw_bytes = convert_json_to_docx_raw_bytes(json_data)

            st.download_button(
                label="📥 Tải Word (Dạng $...$ thô)",
                data=docx_raw_bytes,
                file_name=f"{base_name}_RawMath.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        with col3:
            st.markdown("### 3. File Markdown")
            st.caption("File văn bản thuần dạng .md dùng cho Notion / Obsidian")

            st.download_button(
                label="📥 Tải File Markdown (.md)",
                data=md_content,
                file_name=f"{base_name}.md",
                mime="text/markdown",
                use_container_width=True,
            )

        st.divider()

        # 3. Xem trước nội dung
        st.subheader("👁️ Xem trước nội dung (Markdown Preview)")
        st.markdown(md_content)

    except Exception as e:
        st.error(f"Đã xảy ra lỗi khi xử lý: {e}")