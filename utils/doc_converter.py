"""
doc_converter.py - Smart Document Converter (Image & PDF to Word .docx)
Powered by Google Gemini Vision & python-docx.

Converts scanned images, photos, and PDF files into fully editable, beautifully styled 
Microsoft Word (.docx) documents with rich formatting (headings, tables, lists, styles, spacing).
Includes bulletproof download utilities to guarantee .docx extension across all browsers.
"""

import io
import os
import re
import time
import unicodedata
from typing import Optional, Tuple, List, Dict, Any
from PIL import Image
import google.generativeai as genai
from google.api_core.exceptions import ResourceExhausted, GoogleAPIError

import docx
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

from pypdf import PdfReader, PdfWriter
import pypdfium2 as pdfium
import pdfplumber


# ---------------------------------------------------------------------------
# Bulletproof Filename & Export Helpers
# ---------------------------------------------------------------------------
def sanitize_filename(name: str, fallback: str = "Tai_Lieu_Chuyen_Doi.docx") -> str:
    """
    Sanitizes any filename to ensure it is clean, ASCII-compatible, 
    and ALWAYS ends with a valid .docx extension.
    Prevents Chrome/Edge from falling back to UUID or omitting extensions.
    """
    if not name or not name.strip():
        return fallback

    # Strip existing extension
    clean = name.strip()
    if clean.lower().endswith(".docx"):
        clean = clean[:-5]
    elif "." in clean:
        clean = clean.rsplit(".", 1)[0]

    # Convert Vietnamese specific character 'đ'/'Đ' before NFKD decomposition
    clean = clean.replace("đ", "d").replace("Đ", "D")

    # Decompose unicode and keep ASCII
    clean = unicodedata.normalize("NFKD", clean).encode("ascii", "ignore").decode("ascii")

    # Replace invalid filename characters, spaces, and punctuation with underscores
    clean = re.sub(r"[^a-zA-Z0-9_\-]", "_", clean)
    clean = re.sub(r"_+", "_", clean).strip("_")

    if not clean:
        clean = "Tai_Lieu_Chuyen_Doi"

    return f"{clean}.docx"


def _write_with_fallback(file_bytes: bytes, target_path: str) -> str:
    """
    Attempts to write file_bytes to target_path.
    If the file is locked (PermissionError), automatically appends a timestamp
    to the filename and retries, guaranteeing the save always succeeds.
    Returns the actual path that was written to.
    """
    try:
        with open(target_path, "wb") as f:
            f.write(file_bytes)
        return target_path
    except PermissionError:
        # File is locked (e.g. open in Word) – generate a unique fallback name
        import datetime
        base, ext = os.path.splitext(target_path)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_path = f"{base}_{timestamp}{ext}"
        # If even the timestamped path somehow exists and is locked, add counter
        counter = 0
        while os.path.exists(fallback_path):
            counter += 1
            fallback_path = f"{base}_{timestamp}_{counter}{ext}"
        with open(fallback_path, "wb") as f:
            f.write(file_bytes)
        return fallback_path


def save_file_locally(file_bytes: bytes, filename: str, folder_name: str = "exports") -> str:
    """
    Saves the generated docx file to local disk (exports/) and static web folder (static/)
    to ensure 100% reliable direct download with correct .docx filename on Windows/Chrome.
    If the target file is locked (e.g. open in Word), automatically saves with a
    timestamped filename to avoid PermissionError.
    """
    clean_name = sanitize_filename(filename)
    
    # 1. Save to exports directory for persistent local storage
    os.makedirs(folder_name, exist_ok=True)
    target_path = os.path.join(folder_name, clean_name)
    actual_path = _write_with_fallback(file_bytes, target_path)

    # 2. Save to static directory for direct HTTP static serving via Streamlit
    try:
        os.makedirs("static", exist_ok=True)
        # Use the actual saved filename (may include timestamp) for consistency
        actual_name = os.path.basename(actual_path)
        static_path = os.path.join("static", actual_name)
        _write_with_fallback(file_bytes, static_path)
    except Exception:
        pass

    return os.path.abspath(actual_path)


# ---------------------------------------------------------------------------
# Specialized Prompts for Faithful Layout & Formatting Preservation
# ---------------------------------------------------------------------------
PROMPT_TEMPLATES = {
    "exam": """Bạn là chuyên gia số hóa đề thi và tài liệu giáo dục hàng đầu tại Việt Nam.
Nhiệm vụ của bạn là nhận diện và TÁI TẠO NGUYÊN VẸN 100% CẤU TRÚC, BẢNG BIỂU, KHOẢNG CÁCH VÀ ĐỊNH DẠNG từ hình ảnh/PDF thành định dạng Markdown chuẩn mực cao để xuất sang Microsoft Word (.docx).

QUY TẮC BẢO TOÀN BỐ CỤC ĐẶC THÙ CHO ĐỀ THI & PHIẾU BÀI TẬP:

1. PHẦN ĐẦU TRANG (TIÊU ĐỀ 2 CỘT):
   Nếu đầu tài liệu chia 2 bên (bên trái là Tên Trường/Phòng GD, bên phải là Quốc hiệu hoặc Thời gian thi), BẮT BUỘC dùng Markdown Table 2 cột:
   | [CƠ QUAN / TRƯỜNG HỌC] | [CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM] |
   | [MÃ ĐỀ / KHỐI LỚP] | [Độc lập - Tự do - Hạnh phúc] |

2. TÊN ĐỀ THI & THÔNG TIN CHUNG:
   # ĐỀ KIỂM TRA ĐỊNH KỲ ...
   ## MÔN: ... • THỜI GIAN: ...

3. KHUNG THÔNG TIN HỌC SINH / ĐIỂM SỐ / LỜI PHÊ:
   Nếu đề thi có khung Họ và tên, Lớp, Điểm số, hãy định dạng thành Markdown Table hoặc hộp thông tin rõ ràng:
   | Họ và tên: ................................................ | Điểm số: .................... |
   | Lớp: ..................................................... | Lời nhận xét: ............... |

4. PHẦN TRẮC NGHIỆM:
   - Mục lớn: ### I. PHẦN TRẮC NGHIỆM (điểm số nếu có)
   - Tên câu hỏi: **Câu 1:** [Nội dung câu hỏi]
   - 4 phương án trả lời A, B, C, D trình bày rõ ràng:
     A. [Nội dung A]       B. [Nội dung B]       C. [Nội dung C]       D. [Nội dung D]
     (Nếu phương án dài, mỗi phương án trên 1 dòng riêng:
     A. [Nội dung A]
     B. [Nội dung B]...)

5. PHẦN TỰ LUẬN & BÀI TẬP TÍNH TOÁN:
   - Mục lớn: ### II. PHẦN TỰ LUẬN (điểm số nếu có)
   - Tên bài: **Bài 1:** [Đề bài]
   - Nếu có bảng số liệu, BẮT BUỘC chuyển thành Markdown Table đầy đủ mọi cột và hàng:
     | Cột 1 | Cột 2 | Cột 3 |
     |---|---|---|
     | Dữ liệu | Dữ liệu | Dữ liệu |
   - Nếu bài tập có phần chừa dòng làm bài cho học sinh, ghi rõ `*Bài làm:*` và thêm các dòng chấm chấm:
     ....................................................................................................
     ....................................................................................................

6. CÔNG THỨC TOÁN HỌC & KÝ HIỆU:
   - Giữ nguyên số mũ (x², cm²), phân số (1/2, 3/4), dấu nhân (x hoặc ×), dấu chia (: hoặc ÷), biểu thức tính toán.

7. CHỈ TRẢ VỀ DUY NHẤT nội dung Markdown, không kèm bất kỳ lời chào hỏi hay giải thích nào.""",

    "general": """Bạn là chuyên gia số hóa và phục chế tài liệu văn bản.
Nhiệm vụ của bạn là nhận diện và TÁI TẠO NGUYÊN VẸN CẤU TRÚC, TIÊU ĐỀ, BẢNG BIỂU, ĐOẠN VĂN từ tài liệu hình ảnh/PDF sang Markdown chuẩn để đưa vào Microsoft Word (.docx).

QUY TẮC BẢO TOÀN BỐ CỤC:
1. ĐỀ MỤC & PHÂN CẤP:
   - # Tiêu đề chính của tài liệu
   - ## Tiêu đề mục lớn (I, II, III...)
   - ### Tiêu đề mục nhỏ (1, 2, 3... hoặc A, B, C...)
   - #### Tiểu mục chi tiết (a, b, c...)
2. BẢNG BIỂU:
   Mọi bảng biểu trong tài liệu BẮT BUỘC phải chuyển thành Markdown Table hoàn chỉnh:
   | Tiêu đề 1 | Tiêu đề 2 | Tiêu đề 3 |
   |---|---|---|
   | Dữ liệu | Dữ liệu | Dữ liệu |
   Không được bỏ sót cột hoặc dòng nào.
3. ĐOẠN VĂN & KHOẢNG CÁCH:
   - Tách đoạn rõ ràng bằng dòng trống giữa các đoạn văn.
   - Giữ nguyên định dạng **in đậm** cho từ khóa, *in nghiêng* cho chú thích/trích dẫn.
4. DANH SÁCH: Dùng `- ` cho gạch đầu dòng và `1. `, `2. ` cho danh sách đánh số.
5. CHỈ TRẢ VỀ DUY NHẤT nội dung Markdown chuẩn.""",

    "lesson_plan": """Bạn là chuyên gia giáo dục am hiểu cấu trúc Giáo án / Kế hoạch bài dạy (KHBD) theo quy định của Bộ GD&ĐT Việt Nam.
Nhiệm vụ của bạn là số hóa giáo án từ hình ảnh/PDF thành văn bản Markdown chuẩn để xuất ra file Word (.docx).

QUY TẮC CẤU TRÚC:
1. THÔNG TIN CHUNG:
   # KẾ HOẠCH BÀI DẠY (Tên bài học, Môn học, Khối lớp, Tuần, Tiết)
   ## I. YÊU CẦU CẦN ĐẠT (Năng lực đặc thù, Năng lực chung, Phẩm chất chủ yếu)
   ## II. ĐỒ DÙNG DẠY HỌC (Thiết bị dạy học của GV và HS)
   ## III. CÁC HOẠT ĐỘNG DẠY HỌC CHỦ YẾU
2. BẢNG HOẠT ĐỘNG 2 CỘT:
   Nếu giáo án có bảng tiến trình bài dạy, hãy chuyển thành Markdown Table 2 cột chuẩn:
   | Hoạt động của giáo viên | Hoạt động của học sinh |
   |---|---|
   | [Nội dung hướng dẫn của GV] | [Nội dung thực hiện của HS] |
3. ## IV. ĐIỀU CHỈNH SAU BÀI DẠY
4. CHỈ TRẢ VỀ DUY NHẤT nội dung Markdown, không kèm lời dẫn.""",

    "official_doc": """Bạn là chuyên viên văn thư lưu trữ am hiểu thể thức văn bản hành chính Việt Nam (Nghị định 30/2020/NĐ-CP).
Nhiệm vụ của bạn là chuyển đổi công văn, thông tư, quyết định, biên bản từ ảnh/PDF sang Markdown chuẩn để tái tạo vào file Word (.docx).

QUY TẮC THỂ THỨC:
1. ĐẦU VĂN BẢN (2 CỘT):
   | [CƠ QUAN BAN HÀNH] | [CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM] |
   | Số: .../... | [Độc lập - Tự do - Hạnh phúc] |
2. TÊN LOẠI VĂN BẢN & TRÍCH YẾU:
   # [QUYẾT ĐỊNH / THÔNG BÁO / KẾ HOẠCH]
   **Về việc: [Trích yếu nội dung]**
3. CẤU TRÚC ĐIỀU KHOẢN: Giữ nguyên Điều 1, Điều 2, Khoản 1, 2, Điểm a, b...
4. BẢNG PHÂN CÔNG / DANH SÁCH: Chuyển thành Markdown Table đầy đủ.
5. CHÂN VĂN BẢN (NƠI NHẬN & CHỮ KÝ):
   | *Nơi nhận:* | [CHỨC VỤ NGƯỜI KÝ] |
   | - Như Điều ...<br>- Lưu: VT | *(Ký, ghi rõ họ tên)* |
6. CHỈ TRẢ VỀ DUY NHẤT nội dung Markdown.""",

    "table_data": """Bạn là chuyên gia trích xuất bảng biểu số liệu từ tài liệu scan.
Nhiệm vụ của bạn là trích xuất toàn bộ bảng số liệu, bảng điểm, danh sách học sinh từ ảnh/PDF thành Markdown Table hoàn chỉnh.

QUY TẮC:
1. Nhận diện mọi cột và hàng của bảng. Không bỏ sót bất kỳ dòng nào.
2. Mọi dữ liệu phải khớp chính xác hàng và cột:
   | STT | Họ và tên | Điểm KT | Nhận xét |
   |---|---|---|---|
   | 1 | Nguyễn Văn A | 9.0 | Hoàn thành tốt |
3. CHỈ TRẢ VỀ DUY NHẤT bảng Markdown, không kèm bất kỳ lời dẫn nào."""
}


# ---------------------------------------------------------------------------
# Core Word Document Generator from Markdown (python-docx)
# ---------------------------------------------------------------------------
class MarkdownToDocxBuilder:
    """Builds highly professional, styled Word (.docx) documents with faithful layout preservation."""

    def __init__(
        self,
        font_name: str = "Times New Roman",
        font_size_pt: int = 13,
        primary_color_hex: str = "1E3A8A",  # Deep Navy
        secondary_color_hex: str = "0284C7",  # Blue
        table_theme: str = "education_standard",  # "education_standard", "modern_navy", "minimal_classic"
        page_margins_cm: Tuple[float, float, float, float] = (2.0, 2.0, 2.5, 2.0),  # Top, Bottom, Left, Right
    ):
        self.font_name = font_name
        self.font_size_pt = font_size_pt
        self.primary_color_hex = primary_color_hex
        self.secondary_color_hex = secondary_color_hex
        self.table_theme = table_theme
        self.margins_cm = page_margins_cm

    def create_docx(self, markdown_text: str, document_title: Optional[str] = None) -> bytes:
        """Converts Markdown text into a formatted .docx file with high layout fidelity."""
        doc = docx.Document()

        # Page Setup: A4 with standard Vietnamese administrative margins
        section = doc.sections[0]
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(self.margins_cm[0])
        section.bottom_margin = Cm(self.margins_cm[1])
        section.left_margin = Cm(self.margins_cm[2])
        section.right_margin = Cm(self.margins_cm[3])

        # Configure Header & Footer
        self._setup_header_footer(section, document_title)

        # Parse and process lines
        lines = markdown_text.splitlines()
        i = 0
        total_lines = len(lines)

        while i < total_lines:
            line = lines[i].strip()

            # Empty lines
            if not line:
                i += 1
                continue

            # Markdown Table detection
            if line.startswith("|") and line.endswith("|"):
                table_lines = []
                while i < total_lines and lines[i].strip().startswith("|") and lines[i].strip().endswith("|"):
                    table_lines.append(lines[i].strip())
                    i += 1
                self._add_markdown_table(doc, table_lines)
                continue

            # Heading 1 (# ...)
            if line.startswith("# "):
                text = line[2:].strip()
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(14)
                p.paragraph_format.space_after = Pt(6)
                p.paragraph_format.line_spacing = 1.2
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if any(kw in text.upper() for kw in ["CỘNG HÒA", "ĐỀ THI", "ĐỀ KIỂM TRA", "KẾ HOẠCH", "BẢNG", "QUYẾT ĐỊNH"]) else WD_ALIGN_PARAGRAPH.LEFT
                run = p.add_run(text)
                run.font.name = self.font_name
                run.font.size = Pt(self.font_size_pt + 3)
                run.font.bold = True
                run.font.color.rgb = RGBColor(30, 58, 138)  # Deep Navy
                i += 1
                continue

            # Heading 2 (## ...)
            if line.startswith("## "):
                text = line[3:].strip()
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(10)
                p.paragraph_format.space_after = Pt(4)
                p.paragraph_format.line_spacing = 1.15
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if any(kw in text.upper() for kw in ["MÔN:", "THỜI GIAN", "LỚP:"]) else WD_ALIGN_PARAGRAPH.LEFT
                run = p.add_run(text)
                run.font.name = self.font_name
                run.font.size = Pt(self.font_size_pt + 1.5)
                run.font.bold = True
                run.font.color.rgb = RGBColor(15, 23, 42)  # Slate 900
                i += 1
                continue

            # Heading 3 (### ...)
            if line.startswith("### "):
                text = line[4:].strip()
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.15
                run = p.add_run(text)
                run.font.name = self.font_name
                run.font.size = Pt(self.font_size_pt)
                run.font.bold = True
                run.font.color.rgb = RGBColor(37, 99, 235)  # Blue 600
                i += 1
                continue

            # Heading 4 (#### ...)
            if line.startswith("#### "):
                text = line[5:].strip()
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(2)
                run = p.add_run(text)
                run.font.name = self.font_name
                run.font.size = Pt(self.font_size_pt)
                run.font.bold = True
                run.font.italic = True
                i += 1
                continue

            # Horizontal Rule (--- or ***)
            if re.match(r"^(\-{3,}|\*{3,}|_{3,})$", line):
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)
                p_run = p.add_run("―" * 45)
                p_run.font.name = self.font_name
                p_run.font.size = Pt(10)
                p_run.font.color.rgb = RGBColor(148, 163, 184)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                i += 1
                continue

            # Dotted Answer Writing Lines (e.g. .....................)
            if re.match(r"^\.{10,}$", line) or "........" in line:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.35
                p_run = p.add_run(line)
                p_run.font.name = self.font_name
                p_run.font.size = Pt(self.font_size_pt - 1)
                p_run.font.color.rgb = RGBColor(148, 163, 184)
                i += 1
                continue

            # Blockquote (> ...)
            if line.startswith(">"):
                quote_text = re.sub(r"^>\s*", "", line).strip()
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.4)
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(4)
                self._format_text_runs(p, quote_text, is_italic=True, text_color=RGBColor(71, 85, 105))
                i += 1
                continue

            # Bullet List (- , * , + )
            if re.match(r"^[-*+]\s+", line):
                item_text = re.sub(r"^[-*+]\s+", "", line).strip()
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.25)
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15
                bullet_run = p.add_run("•  ")
                bullet_run.font.name = self.font_name
                bullet_run.font.size = Pt(self.font_size_pt)
                bullet_run.font.bold = True
                bullet_run.font.color.rgb = RGBColor(30, 58, 138)
                self._format_text_runs(p, item_text)
                i += 1
                continue

            # Numbered List (1. , 2. ...)
            num_match = re.match(r"^(\d+[\.\)])\s+(.*)", line)
            if num_match:
                prefix = num_match.group(1)
                item_text = num_match.group(2).strip()
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.3)
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.15
                num_run = p.add_run(f"{prefix}  ")
                num_run.font.name = self.font_name
                num_run.font.size = Pt(self.font_size_pt)
                num_run.font.bold = True
                self._format_text_runs(p, item_text)
                i += 1
                continue

            # Question Header detection (e.g. Câu 1:, Bài 1:)
            is_question = bool(re.match(r"^(Câu|Bài)\s+\d+[\s\:\.]", line, re.IGNORECASE))

            # Regular Paragraph
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6 if is_question else 2)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.15

            # Center alignment for official phrases like "Độc lập - Tự do - Hạnh phúc"
            if any(key_term in line for key_term in ["Độc lập - Tự do - Hạnh phúc", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"]):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            self._format_text_runs(p, line)
            i += 1

        # Save to memory buffer
        out_stream = io.BytesIO()
        doc.save(out_stream)
        return out_stream.getvalue()

    def _setup_header_footer(self, section, document_title: Optional[str]):
        """Sets up running header and footer with page number."""
        footer = section.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        f_p.paragraph_format.space_before = Pt(6)

        doc_name = document_title or "Tài liệu số hóa từ Google AI"
        run_title = f_p.add_run(f"{doc_name}   |   ")
        run_title.font.name = self.font_name
        run_title.font.size = Pt(9)
        run_title.font.color.rgb = RGBColor(148, 163, 184)

        run_page = f_p.add_run("Trang ")
        run_page.font.name = self.font_name
        run_page.font.size = Pt(9)
        run_page.font.color.rgb = RGBColor(100, 116, 139)

        # Word Page Field
        self._add_page_number_field(run_page)

    def _add_page_number_field(self, run):
        """Adds standard Word XML PAGE number field."""
        fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
        instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
        fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
        fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
        run._r.append(fldChar3)

    def _add_markdown_table(self, doc: docx.Document, table_lines: List[str]):
        """Converts Markdown table rows into a beautifully formatted Word table."""
        parsed_rows: List[List[str]] = []
        for line in table_lines:
            raw_cells = line.strip().strip("|").split("|")
            clean_cells = [c.strip() for c in raw_cells]
            # Skip separator line like |---|---|
            if all(re.match(r"^:?-+:?$", c) for c in clean_cells if c):
                continue
            parsed_rows.append(clean_cells)

        if not parsed_rows:
            return

        # Normalize column count
        col_count = max(len(r) for r in parsed_rows)
        if col_count == 0:
            return

        for r in parsed_rows:
            while len(r) < col_count:
                r.append("")

        # -------------------------------------------------------------
        # Check if this is a 2-column Header Table (School name & National motto)
        # -------------------------------------------------------------
        is_header_table = False
        if col_count == 2 and len(parsed_rows) <= 3:
            first_row_joined = " ".join(parsed_rows[0]).upper()
            if any(k in first_row_joined for k in ["CỘNG HÒA", "TRƯỜNG", "PHÒNG GD", "SỞ GD", "UBND", "MÃ ĐỀ", "HỌ VÀ TÊN"]):
                is_header_table = True

        table = doc.add_table(rows=len(parsed_rows), cols=col_count)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        # Apply Table Borders XML
        if is_header_table:
            self._set_borderless_table(table)
        else:
            self._set_table_borders(table)

        # Calculate optimal proportional column widths based on maximum characters
        col_max_lens = [max(len(r[c]) for r in parsed_rows) for c in range(col_count)]
        total_len = max(sum(col_max_lens), 1)
        total_table_width_cm = 16.5  # A4 width (21cm) - Left margin (2.5cm) - Right margin (2.0cm)

        if is_header_table:
            col_widths_cm = [7.5, 9.0]
        else:
            col_widths_cm = []
            for clen in col_max_lens:
                # Proportional share with min width 1.2cm
                share = (clen / total_len) * total_table_width_cm
                col_widths_cm.append(max(1.2, round(share, 2)))
            # Rebalance sum to 16.5 cm
            current_sum = sum(col_widths_cm)
            if current_sum > 0:
                col_widths_cm = [round(w * (total_table_width_cm / current_sum), 2) for w in col_widths_cm]

        # Populate and style cells
        for row_idx, row_data in enumerate(parsed_rows):
            row = table.rows[row_idx]

            # cantSplit: Prevent rows from splitting across page breaks
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(r'<w:cantSplit %s/>' % nsdecls('w')))

            is_header = (row_idx == 0 and not is_header_table)

            # Repeat header on every page
            if is_header:
                trPr.append(parse_xml(r'<w:tblHeader %s/>' % nsdecls('w')))

            for col_idx, cell_value in enumerate(row_data):
                cell = row.cells[col_idx]
                cell.width = Cm(col_widths_cm[col_idx])
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                tcPr = cell._tc.get_or_add_tcPr()

                # Padding
                top_pad = "80" if is_header_table else "120"
                side_pad = "100" if is_header_table else "160"
                tcMar = parse_xml(
                    r'<w:tcMar %s>'
                    rf'  <w:top w:w="{top_pad}" w:type="dxa"/>'
                    rf'  <w:bottom w:w="{top_pad}" w:type="dxa"/>'
                    rf'  <w:left w:w="{side_pad}" w:type="dxa"/>'
                    rf'  <w:right w:w="{side_pad}" w:type="dxa"/>'
                    r'</w:tcMar>' % nsdecls('w')
                )
                tcPr.append(tcMar)

                # Shading / Background color based on theme
                if not is_header_table:
                    if is_header:
                        if self.table_theme == "modern_navy":
                            shd = parse_xml(r'<w:shd %s w:fill="1E3A8A"/>' % nsdecls('w'))
                        elif self.table_theme == "minimal_classic":
                            shd = parse_xml(r'<w:shd %s w:fill="FFFFFF"/>' % nsdecls('w'))
                        else:  # education_standard
                            shd = parse_xml(r'<w:shd %s w:fill="F1F5F9"/>' % nsdecls('w'))
                        tcPr.append(shd)
                    elif row_idx % 2 == 1:
                        # Subtle zebra striping for odd rows
                        shd = parse_xml(r'<w:shd %s w:fill="F8FAFC"/>' % nsdecls('w'))
                        tcPr.append(shd)

                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.1

                if is_header_table:
                    # Alignment in header table: left column centered or left, right column centered
                    if col_idx == 0:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    self._format_text_runs(
                        p,
                        cell_value,
                        is_bold=("CỘNG HÒA" in cell_value.upper() or "TRƯỜNG" in cell_value.upper() or "PHÒNG" in cell_value.upper()),
                        is_italic=("Độc lập" in cell_value),
                        custom_size_pt=self.font_size_pt - 0.5,
                    )
                else:
                    # Header styling for normal data tables
                    if is_header:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        header_color = RGBColor(255, 255, 255) if self.table_theme == "modern_navy" else RGBColor(15, 23, 42)
                        self._format_text_runs(
                            p,
                            cell_value,
                            is_bold=True,
                            text_color=header_color,
                            custom_size_pt=self.font_size_pt - 1,
                        )
                    else:
                        # Center align for numbers, short codes
                        if re.match(r"^(\d+|[T|H|C]|[A-D]|\d+[\.,]\d+|[\+\-])$", cell_value.strip()):
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        else:
                            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                        self._format_text_runs(
                            p,
                            cell_value,
                            custom_size_pt=self.font_size_pt - 1,
                        )

        # Space after table
        space_p = doc.add_paragraph()
        space_p.paragraph_format.space_before = Pt(4)
        space_p.paragraph_format.space_after = Pt(4)

    def _set_borderless_table(self, table):
        """Sets borders to none for 2-column header tables."""
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            r'<w:tblBorders %s>'
            r'  <w:top w:val="none"/>'
            r'  <w:bottom w:val="none"/>'
            r'  <w:left w:val="none"/>'
            r'  <w:right w:val="none"/>'
            r'  <w:insideH w:val="none"/>'
            r'  <w:insideV w:val="none"/>'
            r'</w:tblBorders>' % nsdecls('w')
        )
        tblPr.append(borders)

    def _set_table_borders(self, table):
        """Sets subtle, elegant table borders."""
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            r'<w:tblBorders %s>'
            r'  <w:top w:val="single" w:sz="6" w:space="0" w:color="94A3B8"/>'
            r'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="94A3B8"/>'
            r'  <w:left w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
            r'  <w:right w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
            r'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            r'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            r'</w:tblBorders>' % nsdecls('w')
        )
        tblPr.append(borders)

    def _format_text_runs(
        self,
        paragraph,
        text: str,
        is_bold: bool = False,
        is_italic: bool = False,
        text_color: Optional[RGBColor] = None,
        custom_size_pt: Optional[float] = None,
    ):
        """Parses inline markdown like **bold**, *italic*, and `code` into styled docx runs."""
        # Simple regex tokenizer for bold, italic, code
        pattern = re.compile(r"(\*\*.*?\*\*|\*.*?\*|`.*?`|[^\*`]+)")
        tokens = pattern.findall(text)

        for token in tokens:
            if not token:
                continue

            run = paragraph.add_run()
            run.font.name = self.font_name
            run.font.size = Pt(custom_size_pt or self.font_size_pt)

            if text_color:
                run.font.color.rgb = text_color
            else:
                run.font.color.rgb = RGBColor(15, 23, 42)  # Slate 900

            run_text = token
            run_bold = is_bold
            run_italic = is_italic

            if token.startswith("**") and token.endswith("**") and len(token) >= 4:
                run_text = token[2:-2]
                run_bold = True
            elif token.startswith("*") and token.endswith("*") and len(token) >= 2:
                run_text = token[1:-1]
                run_italic = True
            elif token.startswith("`") and token.endswith("`") and len(token) >= 2:
                run_text = token[1:-1]
                run.font.name = "Consolas"
                run.font.color.rgb = RGBColor(225, 29, 72)  # Pink-red code

            run.text = run_text
            run.bold = run_bold
            run.italic = run_italic


# ---------------------------------------------------------------------------
# Google Gemini Multimodal Vision Document Extractor
# ---------------------------------------------------------------------------
class GoogleVisionDocExtractor:
    """Uses Google Gemini API multimodal vision to extract and reconstruct document structure."""

    def __init__(self, api_key: str, model_name: str = "gemini-3.6-flash"):
        self.api_key = api_key.strip()
        # Strictly lock model to gemini-3.6-flash to avoid 404 errors on deprecated models
        if not model_name or "1.5" in model_name or "2.5" in model_name:
            self.model_name = "gemini-3.6-flash"
        else:
            self.model_name = model_name
        genai.configure(api_key=self.api_key)

    @staticmethod
    def inspect_pdf(pdf_bytes: bytes) -> Dict[str, Any]:
        """Returns metadata and page count of a PDF file."""
        try:
            reader = PdfReader(io.BytesIO(pdf_bytes))
            num_pages = len(reader.pages)
            meta = reader.metadata or {}
            title = meta.get("/Title") or ""
            return {
                "num_pages": num_pages,
                "title": title,
                "is_encrypted": reader.is_encrypted,
            }
        except Exception as e:
            return {"num_pages": 1, "title": "", "error": str(e)}

    @staticmethod
    def render_pdf_page_thumbnail(pdf_bytes: bytes, page_idx: int = 0) -> Optional[Image.Image]:
        """Renders a PDF page to PIL Image for instant UI preview using pypdfium2."""
        try:
            pdf = pdfium.PdfDocument(pdf_bytes)
            if page_idx < 0 or page_idx >= len(pdf):
                page_idx = 0
            page = pdf[page_idx]
            # scale=1.5 gives crisp quality thumbnail
            pil_img = page.render(scale=1.5).to_pil()
            return pil_img
        except Exception:
            return None

    def slice_pdf_pages(self, pdf_bytes: bytes, start_page: int, end_page: int) -> bytes:
        """Extracts a range of pages from PDF into a new PDF byte buffer (1-indexed)."""
        reader = PdfReader(io.BytesIO(pdf_bytes))
        writer = PdfWriter()
        total = len(reader.pages)

        start_idx = max(0, start_page - 1)
        end_idx = min(total, end_page)

        for idx in range(start_idx, end_idx):
            writer.add_page(reader.pages[idx])

        out = io.BytesIO()
        writer.write(out)
        return out.getvalue()

    def extract_digital_text_hint(self, pdf_bytes: bytes, max_pages: int = 5) -> str:
        """Extracts digital text layer from PDF using pdfplumber to provide ground-truth text context."""
        hints = []
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for idx, page in enumerate(pdf.pages[:max_pages]):
                    t = page.extract_text()
                    if t and t.strip():
                        hints.append(f"--- Trang {idx + 1} ---\n{t.strip()}")
        except Exception:
            pass
        return "\n\n".join(hints)

    def process_document(
        self,
        file_bytes: bytes,
        mime_type: str,
        document_type: str = "general",
        custom_instructions: str = "",
        max_retries: int = 3,
    ) -> str:
        """
        Sends the image or PDF bytes to Google Gemini multimodal vision.
        Returns clean, formatted Markdown representing the document.
        """
        system_prompt = PROMPT_TEMPLATES.get(document_type, PROMPT_TEMPLATES["general"])
        if custom_instructions and custom_instructions.strip():
            system_prompt += f"\n\nYÊU CẦU BỔ SUNG TỪ NGƯỜI DÙNG:\n{custom_instructions.strip()}"

        # If it's a PDF, check if digital text can be extracted to supplement OCR accuracy
        if mime_type == "application/pdf":
            digital_hint = self.extract_digital_text_hint(file_bytes)
            if digital_hint:
                system_prompt += (
                    "\n\nLỚP VĂN BẢN GỐC ĐÍNH KÈM TỪ FILE PDF (HÃY ĐỐI CHIẾU ĐỂ ĐẢM BẢO ĐÚNG TỪNG TỪ TIẾNG VIỆT VÀ SỐ LIỆU):\n"
                    + digital_hint
                )

        model = genai.GenerativeModel(self.model_name)

        # Prepare payload part
        media_part = {
            "mime_type": mime_type,
            "data": file_bytes,
        }

        contents = [
            media_part,
            system_prompt,
        ]

        # Call with retry logic for 429/ResourceExhausted
        for attempt in range(max_retries):
            try:
                response = model.generate_content(
                    contents,
                    generation_config=genai.types.GenerationConfig(
                        temperature=0.1,  # Low temperature for highest fidelity OCR & layout
                    ),
                )

                if response and response.text:
                    raw_text = response.text.strip()
                    # Strip any ```markdown code fences if returned by the model
                    if raw_text.startswith("```markdown"):
                        raw_text = raw_text[len("```markdown"):].strip()
                    elif raw_text.startswith("```"):
                        raw_text = raw_text[len("```"):].strip()
                    if raw_text.endswith("```"):
                        raw_text = raw_text[:-3].strip()
                    return raw_text

                raise ValueError("Không nhận được nội dung phản hồi từ Google Gemini.")

            except ResourceExhausted:
                if attempt < max_retries - 1:
                    wait_time = (attempt + 1) * 3
                    time.sleep(wait_time)
                else:
                    raise RuntimeError("Hạn mức Google API (Rate Limit) tạm thời vượt ngưỡng. Vui lòng thử lại sau giây lát.")
            except Exception as e:
                err_msg = str(e)
                if "404" in err_msg or "not found" in err_msg.lower():
                    # Fallback to gemini-3.6-flash if user selected an obsolete model
                    try:
                        self.model_name = "gemini-3.6-flash"
                        fb_model = genai.GenerativeModel("gemini-3.6-flash")
                        res_fb = fb_model.generate_content(contents)
                        if res_fb and res_fb.text:
                            text_fb = res_fb.text.strip()
                            if text_fb.startswith("```markdown"):
                                text_fb = text_fb[len("```markdown"):].strip()
                            if text_fb.endswith("```"):
                                text_fb = text_fb[:-3].strip()
                            return text_fb
                    except Exception:
                        pass
                raise e

        raise RuntimeError("Không thể xử lý tài liệu sau các lần thử kết nối.")
