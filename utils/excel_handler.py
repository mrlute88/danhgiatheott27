"""
excel_handler.py - Robust Excel processing using openpyxl & pandas.
Preserves original Excel formatting, fonts, sheet structures, and formulas.
Supports automated pedagogical metadata extraction and intelligent TT27 column detection.
"""

import io
import re
from copy import copy
from typing import Dict, List, Optional, Tuple, Any
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

try:
    import xlrd
except ImportError:
    xlrd = None

try:
    import lxml.html
except ImportError:
    lxml = None

from utils.tt27_knowledge import SUBJECTS, GRADES, PERIODS



def detect_header_row(ws) -> int:
    """
    Scans the first 15 rows of a worksheet to locate the actual table header.
    Returns 1-based row index.
    """
    keyword_patterns = [
        r"họ\s*và\s*tên",
        r"họ\s*tên",
        r"tên\s*học\s*sinh",
        r"\bstt\b",
        r"mã\s*hs",
        r"điểm",
        r"mức\s*(đạt|hoàn\s*thành)",
        r"nhận\s*xét",
        r"ghi\s*chú",
        r"\bxl\b|\bxl\s*(gk|ck|hk|cn|[12])\b",
        r"\bkt\b|\bkt\s*(gk|ck|hk|cn|[12])\b",
        r"\bktđk\b|\bđgk\b",
    ]

    best_row = 1
    max_matches = 0

    for r in range(1, min(16, ws.max_row + 1)):
        row_values = [str(cell.value or "").strip().lower() for cell in ws[r]]
        matches = 0
        for val in row_values:
            if not val:
                continue
            for pat in keyword_patterns:
                if re.search(pat, val):
                    matches += 1
                    break
        if matches > max_matches:
            max_matches = matches
            best_row = r

    # If at least 2 keywords matched, use that row; otherwise default to row 1
    return best_row if max_matches >= 2 else 1


def extract_pedagogical_metadata(
    wb: openpyxl.Workbook,
    sheet_name: Optional[str] = None,
    file_name: str = ""
) -> Dict[str, str]:
    """
    Scans the uploaded workbook (banner title, class title, sheet name, column headers, and filename)
    to automatically detect default pedagogical settings according to Circular 27:
    - Subject (Môn học)
    - Grade (Khối lớp)
    - Period (Thời điểm đánh giá: Giữa Học kỳ 1, Cuối Học kỳ 1, Giữa Học kỳ 2, Cuối Học kỳ 2)
    """
    target_sheet = sheet_name if sheet_name and sheet_name in wb.sheetnames else wb.sheetnames[0]
    ws = wb[target_sheet]

    extracted_texts: List[str] = []
    if file_name:
        extracted_texts.append(str(file_name))
    if target_sheet:
        extracted_texts.append(str(target_sheet))

    # Scan the first 16 rows of the sheet across first 25 columns
    for r in range(1, min(16, ws.max_row + 1)):
        for c in range(1, min(25, ws.max_column + 1)):
            v = ws.cell(r, c).value
            if v is not None:
                s = str(v).strip()
                if s:
                    extracted_texts.append(s)

    combined = " ".join(extracted_texts).lower()

    # 1. Subject detection
    detected_subject = None
    subject_rules = [
        ("Tiếng Việt", [r"môn\s*tiếng\s*việt", r"tiếng\s*việt", r"tieng\s*viet", r"\btv\b"]),
        ("Toán", [r"môn\s*toán", r"\btoán\b", r"\btoan\b"]),
        ("Tiếng Anh", [r"tiếng\s*anh", r"tieng\s*anh", r"ngoại\s*ngữ", r"tiếng\s*nước\s*ngoài", r"english"]),
        ("Tự nhiên và Xã hội (Lớp 1, 2, 3)", [r"tự\s*nhiên\s*(?:và|&)?\s*xã\s*hội", r"tn\s*&\s*xh", r"tnxh"]),
        ("Khoa học (Lớp 4, 5)", [r"khoa\s*học", r"khoa\s*hoc"]),
        ("Lịch sử và Địa lí (Lớp 4, 5)", [r"lịch\s*sử\s*(?:và|&)?\s*địa\s*l[íy]", r"ls\s*&\s*đl", r"lịch\s*sử", r"địa\s*l[íy]"]),
        ("Tin học và Công nghệ (Lớp 3, 4, 5)", [r"tin\s*học\s*(?:và|&)?\s*công\s*nghệ", r"tin\s*học", r"công\s*nghệ", r"th\s*&\s*cn"]),
        ("Đạo đức", [r"đạo\s*đức", r"dao\s*duc"]),
        ("Âm nhạc", [r"âm\s*nhạc", r"am\s*nhac"]),
        ("Mĩ thuật", [r"m[ĩỹ]\s*thuật", r"m[iy]\s*thuat"]),
        ("Giáo dục thể chất", [r"giáo\s*dục\s*thể\s*chất", r"gdtc", r"thể\s*dục"]),
        ("Hoạt động trải nghiệm", [r"hoạt\s*động\s*trải\s*nghiệm", r"hđtn"]),
    ]
    for subj_name, patterns in subject_rules:
        if any(re.search(p, combined) for p in patterns):
            if subj_name in SUBJECTS:
                detected_subject = subj_name
                break

    # 2. Grade detection (1-5)
    detected_grade = None
    grade_match = re.search(r"(?:khối|lớp|khoi|lop)[_\s]*([1-5])\b", combined)
    if not grade_match:
        grade_match = re.search(r"\b([1-5])[a-z]\d?\b", combined)
    if not grade_match:
        grade_match = re.search(r"(?<![a-z0-9])k[_\s]*([1-5])\b", combined)
    if grade_match:
        g_val = f"Lớp {grade_match.group(1)}"
        if g_val in GRADES:
            detected_grade = g_val

    # 3. Period detection (Giữa kỳ vs Cuối kỳ)
    detected_period = None
    # Check Giua Ky 1 (e.g. GIỮA KỲ 1, GK1, XL GK1...)
    if re.search(r"giữa\s*(?:học\s*)?kỳ\s*1\b|\bgk1\b|\bxl\s*gk1\b|\bkt\s*gk1\b|giữa\s*kỳ\s*i\b", combined):
        detected_period = "Giữa Học kỳ 1"
    # Check Cuoi Ky 1 (e.g. CUỐI KỲ 1, CK1, KT CK1, XL CK1...)
    elif re.search(r"cuối\s*(?:học\s*)?kỳ\s*1\b|\bck1\b|\bkt\s*ck1\b|\bxl\s*ck1\b|cuối\s*kỳ\s*i\b", combined):
        detected_period = "Cuối Học kỳ 1"
    # Check Giua Ky 2 (e.g. GIỮA KỲ 2, GK2, XL GK2...)
    elif re.search(r"giữa\s*(?:học\s*)?kỳ\s*2\b|\bgk2\b|\bxl\s*gk2\b|\bkt\s*gk2\b|giữa\s*kỳ\s*ii\b", combined):
        detected_period = "Giữa Học kỳ 2"
    # Check Cuoi Ky 2 / Ca Nam (e.g. CUỐI KỲ 2, CK2, KT CK2, XL CK2, CUỐI NĂM, CẢ NĂM, CN...)
    elif re.search(r"cuối\s*(?:học\s*)?kỳ\s*2\b|\bck2\b|\bkt\s*ck2\b|\bxl\s*ck2\b|cuối\s*kỳ\s*ii\b|cuối\s*năm|cả\s*năm|\bcn\b|\bxl\s*cn\b|\bkt\s*cn\b", combined):
        detected_period = "Cuối Học kỳ 2 (Cả năm)"
    elif re.search(r"học\s*kỳ\s*1\b|\bhk1\b", combined):
        detected_period = "Cuối Học kỳ 1"
    elif re.search(r"học\s*kỳ\s*2\b|\bhk2\b", combined):
        detected_period = "Cuối Học kỳ 2 (Cả năm)"

    return {
        "subject": detected_subject or ("Tiếng Việt" if "Tiếng Việt" in SUBJECTS else SUBJECTS[0]),
        "grade": detected_grade or ("Lớp 1" if "Lớp 1" in GRADES else GRADES[0]),
        "period": detected_period or ("Giữa Học kỳ 1" if "Giữa Học kỳ 1" in PERIODS else PERIODS[0]),
    }


def ensure_openpyxl_workbook(file_bytes: bytes, file_name: str = "") -> openpyxl.Workbook:
    """
    Universally loads Excel file bytes into a standard openpyxl.Workbook.
    Supports:
    1. Modern .xlsx (OpenXML / ZIP)
    2. Legacy binary .xls (BIFF8 via xlrd with full merged cells support)
    3. HTML tables disguised as .xls (common in VNedu, SMAS, CSDL Ngành exports)
    4. XML Spreadsheet 2003 (.xml / .xls)
    5. Fallback via pandas read_excel / read_html
    """
    if hasattr(file_bytes, "getvalue"):
        file_bytes = file_bytes.getvalue()
    elif hasattr(file_bytes, "read"):
        file_bytes = file_bytes.read()

    if not file_bytes:
        raise ValueError("Dữ liệu file Excel rỗng (0 bytes).")

    # 1. If it's a modern OpenXML (.xlsx / zip)
    if file_bytes.startswith(b"PK\x03\x04"):
        try:
            return openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        except Exception:
            try:
                return openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            except Exception:
                pass

    # 2. If it's XML Spreadsheet 2003
    if b"urn:schemas-microsoft-com:office:spreadsheet" in file_bytes:
        try:
            root = ET.fromstring(file_bytes)
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            for ws_elem in root.findall(".//{urn:schemas-microsoft-com:office:spreadsheet}Worksheet"):
                name = ws_elem.attrib.get("{urn:schemas-microsoft-com:office:spreadsheet}Name", "Sheet1")
                ws = wb.create_sheet(title=name)
                curr_row = 1
                for row_elem in ws_elem.findall(".//{urn:schemas-microsoft-com:office:spreadsheet}Row"):
                    row_idx_attr = row_elem.attrib.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
                    if row_idx_attr and row_idx_attr.isdigit():
                        curr_row = int(row_idx_attr)
                    curr_col = 1
                    for cell_elem in row_elem.findall(".//{urn:schemas-microsoft-com:office:spreadsheet}Cell"):
                        col_idx_attr = cell_elem.attrib.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
                        if col_idx_attr and col_idx_attr.isdigit():
                            curr_col = int(col_idx_attr)
                        data_elem = cell_elem.find(".//{urn:schemas-microsoft-com:office:spreadsheet}Data")
                        if data_elem is not None and data_elem.text is not None:
                            val = data_elem.text.strip()
                            ws.cell(row=curr_row, column=curr_col, value=val)
                        curr_col += 1
                    curr_row += 1
            if wb.sheetnames:
                return wb
        except Exception:
            pass

    # 3. If it looks like HTML / XML with <table> (VNedu, SMAS, CSDL Ngành exports)
    stripped_start = file_bytes[:500].strip().lower()
    is_html = (
        stripped_start.startswith(b"<html")
        or stripped_start.startswith(b"<!doctype")
        or stripped_start.startswith(b"<?xml")
        or b"<table" in stripped_start
        or b"<tr" in stripped_start
    )
    if is_html:
        for enc in ["utf-8", "utf-8-sig", "windows-1258", "cp1252", "latin-1"]:
            try:
                html_text = file_bytes.decode(enc)
                if lxml is not None:
                    doc = lxml.html.fromstring(html_text)
                    tables = doc.xpath("//table")
                    if tables:
                        wb = openpyxl.Workbook()
                        wb.remove(wb.active)
                        for t_idx, tbl in enumerate(tables):
                            sheet_title = f"Sheet{t_idx + 1}"
                            ws = wb.create_sheet(title=sheet_title)
                            occupied = set()
                            for r_idx, tr in enumerate(tbl.xpath(".//tr")):
                                row_num = r_idx + 1
                                col_idx = 1
                                for cell_node in tr.xpath("./th|./td"):
                                    while (row_num, col_idx) in occupied:
                                        col_idx += 1

                                    val = cell_node.text_content().strip()
                                    try:
                                        colspan = int(cell_node.get("colspan", 1))
                                    except Exception:
                                        colspan = 1
                                    try:
                                        rowspan = int(cell_node.get("rowspan", 1))
                                    except Exception:
                                        rowspan = 1

                                    if val:
                                        clean_val = val
                                        if re.match(r"^-?\d+$", clean_val):
                                            try:
                                                clean_val = int(clean_val)
                                            except Exception:
                                                pass
                                        elif re.match(r"^-?\d+\.\d+$", clean_val):
                                            try:
                                                clean_val = float(clean_val)
                                            except Exception:
                                                pass
                                        ws.cell(row=row_num, column=col_idx, value=clean_val)

                                    for dr in range(rowspan):
                                        for dc in range(colspan):
                                            occupied.add((row_num + dr, col_idx + dc))

                                    if colspan > 1 or rowspan > 1:
                                        try:
                                            ws.merge_cells(
                                                start_row=row_num,
                                                end_row=row_num + rowspan - 1,
                                                start_column=col_idx,
                                                end_column=col_idx + colspan - 1,
                                            )
                                        except Exception:
                                            pass

                                    col_idx += colspan
                        if wb.sheetnames and wb.active.max_row > 0:
                            return wb

                # If lxml wasn't available or didn't find tables, try pandas read_html
                dfs = pd.read_html(io.BytesIO(file_bytes), encoding=enc)
                if dfs:
                    wb = openpyxl.Workbook()
                    wb.remove(wb.active)
                    for idx, tbl_df in enumerate(dfs):
                        ws = wb.create_sheet(title=f"Sheet{idx + 1}")
                        ws.append([str(c) for c in tbl_df.columns])
                        for row in tbl_df.itertuples(index=False):
                            ws.append(["" if pd.isna(v) else v for v in row])
                    if wb.sheetnames and wb.active.max_row > 0:
                        return wb
            except Exception:
                continue

    # 4. Try legacy BIFF8 .xls via xlrd
    if xlrd is not None:
        try:
            try:
                x_wb = xlrd.open_workbook(file_contents=file_bytes, formatting_info=True)
            except Exception:
                x_wb = xlrd.open_workbook(file_contents=file_bytes, formatting_info=False)

            wb = openpyxl.Workbook()
            wb.remove(wb.active)

            horiz_align_map = {0: 'general', 1: 'left', 2: 'center', 3: 'right', 4: 'fill', 5: 'justify', 6: 'center_continuous'}
            vert_align_map = {0: 'top', 1: 'center', 2: 'bottom', 3: 'justify'}
            has_formatting = bool(hasattr(x_wb, "xf_list") and x_wb.xf_list and hasattr(x_wb, "font_list"))

            for sheet_idx in range(x_wb.nsheets):
                xs = x_wb.sheet_by_index(sheet_idx)
                ws = wb.create_sheet(title=xs.name or f"Sheet{sheet_idx+1}")

                # Copy column dimensions (widths & hidden status)
                if hasattr(xs, "colinfo_map") and xs.colinfo_map:
                    for c, cinfo in xs.colinfo_map.items():
                        col_letter = get_column_letter(c + 1)
                        if cinfo.width > 0:
                            ws.column_dimensions[col_letter].width = round(cinfo.width / 256.0, 2)
                        if getattr(cinfo, "hidden", 0):
                            ws.column_dimensions[col_letter].hidden = True

                # Copy row dimensions (heights & hidden status)
                if hasattr(xs, "rowinfo_map") and xs.rowinfo_map:
                    for r, rinfo in xs.rowinfo_map.items():
                        if rinfo.height > 0:
                            ws.row_dimensions[r + 1].height = round(rinfo.height / 20.0, 2)
                        if getattr(rinfo, "hidden", 0):
                            ws.row_dimensions[r + 1].hidden = True

                for r in range(xs.nrows):
                    for c in range(xs.ncols):
                        cell = xs.cell(r, c)
                        val = cell.value
                        if cell.ctype == xlrd.XL_CELL_DATE:
                            try:
                                val = xlrd.xldate_as_datetime(val, x_wb.datemode)
                            except Exception:
                                pass
                        elif cell.ctype == xlrd.XL_CELL_NUMBER:
                            if isinstance(val, float) and val.is_integer():
                                val = int(val)
                        elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                            val = bool(val)
                        elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                            val = None

                        target_cell = ws.cell(row=r + 1, column=c + 1)
                        if val is not None and val != "":
                            target_cell.value = val

                        if has_formatting:
                            try:
                                xf = x_wb.xf_list[cell.xf_index]
                                font = x_wb.font_list[xf.font_index]
                                target_cell.font = Font(
                                    name=font.name or "Arial",
                                    size=font.height / 20.0 if font.height else 10,
                                    bold=bool(font.bold),
                                    italic=bool(font.italic),
                                )
                                h_align = horiz_align_map.get(xf.alignment.hor_align, "left")
                                v_align = vert_align_map.get(xf.alignment.vert_align, "center")
                                target_cell.alignment = Alignment(
                                    horizontal=h_align if h_align != "general" else None,
                                    vertical=v_align,
                                    wrap_text=bool(xf.alignment.text_wrapped)
                                )
                                b = xf.border
                                if b.left_line_style or b.right_line_style or b.top_line_style or b.bottom_line_style:
                                    target_cell.border = Border(
                                        left=Side(style="thin" if b.left_line_style else None),
                                        right=Side(style="thin" if b.right_line_style else None),
                                        top=Side(style="thin" if b.top_line_style else None),
                                        bottom=Side(style="thin" if b.bottom_line_style else None),
                                    )
                            except Exception:
                                pass

                if hasattr(xs, "merged_cells") and xs.merged_cells:
                    for rlo, rhi, clo, chi in xs.merged_cells:
                        try:
                            ws.merge_cells(
                                start_row=rlo + 1,
                                end_row=rhi,
                                start_column=clo + 1,
                                end_column=chi,
                            )
                        except Exception:
                            pass

            if wb.sheetnames:
                return wb
        except Exception:
            pass

    # 5. Last-ditch attempt: use pandas read_excel
    try:
        excel_file = io.BytesIO(file_bytes)
        dfs = pd.read_excel(excel_file, sheet_name=None)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for sname, s_df in dfs.items():
            ws = wb.create_sheet(title=str(sname))
            ws.append([str(c) for c in s_df.columns])
            for row in s_df.itertuples(index=False):
                ws.append(["" if pd.isna(v) else v for v in row])
        if wb.sheetnames:
            return wb
    except Exception:
        pass

    raise ValueError(
        f"Không thể đọc file Excel '{file_name or 'này'}'. Định dạng không được hỗ trợ hoặc file bị hỏng/có mật khẩu bảo vệ."
    )


def load_workbook_and_df(
    file_bytes: bytes,
    sheet_name: Optional[str] = None,
    *args,
    **kwargs
) -> Tuple[openpyxl.Workbook, pd.DataFrame, List[str], int]:
    """
    Loads openpyxl Workbook and generates corresponding Pandas DataFrame with detected header.
    Universally supports modern .xlsx, legacy binary .xls (BIFF8), and HTML/XML tables disguised as .xls.
    Intelligently handles split/merged Name columns (Họ đệm + Tên) into a clean Full Name,
    and records exact 1-based Excel row indices for flawless writing back.
    Returns (workbook, dataframe, sheet_names, header_row_idx_1based)
    """
    file_name = kwargs.get("file_name", "") or (args[0] if args else "")
    wb = ensure_openpyxl_workbook(file_bytes, file_name=file_name)
    sheet_names = wb.sheetnames

    target_sheet_name = sheet_name if sheet_name in sheet_names else sheet_names[0]
    ws = wb[target_sheet_name]

    header_row_idx = detect_header_row(ws)

    # Read into pandas from openpyxl workbook via in-memory xlsx buffer to guarantee 100% compatibility
    xlsx_buffer = io.BytesIO()
    wb.save(xlsx_buffer)
    xlsx_buffer.seek(0)
    df = pd.read_excel(xlsx_buffer, sheet_name=target_sheet_name, header=header_row_idx - 1, engine="openpyxl")

    # Track merged ranges in openpyxl on header row (e.g. C7:D7 merged for 'Họ và tên')
    merged_col_map = {}
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= header_row_idx <= rng.max_row and rng.min_col < rng.max_col:
            # 0-indexed column indices
            merged_col_map[rng.min_col - 1] = rng.max_col - 1

    # Check if 'Họ và tên' / 'Họ đệm' is followed by a 'Tên' column or merged sub-column
    cols_to_drop = []
    col_names = [str(c) for c in df.columns]

    for i, c in enumerate(col_names):
        c_lower = c.lower().strip()
        if re.search(r"họ\s*và\s*tên|họ\s*tên|họ\s*đệm|họ\s*(?:và\s*)?chữ\s*lót", c_lower):
            next_idx = i + 1
            should_combine = False
            if i in merged_col_map and merged_col_map[i] == next_idx:
                should_combine = True
            elif next_idx < len(df.columns):
                next_header = str(df.columns[next_idx]).lower().strip()
                if ("unnamed:" in next_header or next_header == "tên" or not next_header) and not df.iloc[:, next_idx].dropna().empty:
                    should_combine = True

            if should_combine and next_idx < len(df.columns):
                name_s = df.iloc[:, i].fillna("").astype(str).str.strip()
                ten_s = df.iloc[:, next_idx].fillna("").astype(str).str.strip()
                combined = []
                for n_val, t_val in zip(name_s, ten_s):
                    if n_val and t_val and t_val.lower() != "nan" and t_val != n_val:
                        combined.append(f"{n_val} {t_val}")
                    elif n_val and n_val.lower() != "nan":
                        combined.append(n_val)
                    elif t_val and t_val.lower() != "nan":
                        combined.append(t_val)
                    else:
                        combined.append("")
                df.iloc[:, i] = combined
                cols_to_drop.append(df.columns[next_idx])
                # Ensure column is cleanly named "Họ và tên"
                if "họ" in df.columns[i].lower():
                    df = df.rename(columns={df.columns[i]: "Họ và tên"})
                break

    if cols_to_drop:
        df = df.drop(columns=cols_to_drop)

    # Clean up dataframe columns: drop columns that are completely unnamed AND completely empty
    empty_unnamed = [
        c for c in df.columns
        if ("Unnamed:" in str(c) or not str(c).strip()) and df[c].dropna().empty
    ]
    if empty_unnamed:
        df = df.drop(columns=empty_unnamed)

    # Convert any remaining unnamed columns with content to friendly names
    cleaned_cols = []
    for i, col in enumerate(df.columns):
        c_str = str(col).strip()
        if "Unnamed:" in c_str or not c_str:
            c_str = f"Cột_{get_column_letter(i + 1)}"
        cleaned_cols.append(c_str)
    df.columns = cleaned_cols

    # Calculate exact 1-based original Excel row numbers for each dataframe row
    start_data_row = header_row_idx + 1
    excel_rows = [start_data_row + idx for idx in range(len(df))]
    df["_excel_row"] = excel_rows

    # Drop completely empty rows
    data_cols = [c for c in df.columns if c != "_excel_row"]
    df = df.dropna(subset=data_cols, how="all")

    # Filter out footer/signature rows (e.g. 'Tổng số: ...', 'Người lập bảng', 'Hiệu trưởng')
    footer_keywords = [
        r"^tổng\s*số",
        r"^người\s*lập",
        r"^hiệu\s*trưởng",
        r"^giáo\s*viên\s*(chủ\s*nhiệm|bộ\s*môn)",
        r"^ban\s*giám\s*hiệu",
        r"^xác\s*nhận"
    ]
    # Identify primary column for filtering
    first_col = df.columns[0]
    name_col = "Họ và tên" if "Họ và tên" in df.columns else (df.columns[1] if len(df.columns) > 1 else first_col)

    valid_mask = []
    for _, row in df.iterrows():
        val1 = str(row.get(first_col, "")).strip().lower()
        val_name = str(row.get(name_col, "")).strip().lower()
        is_footer = any(re.search(pat, val1) or re.search(pat, val_name) for pat in footer_keywords)
        # Also ensure student name is not empty
        has_content = bool(val_name and val_name != "nan") or bool(val1 and val1 != "nan")
        valid_mask.append(has_content and not is_footer)

    if any(valid_mask):
        df = df[valid_mask].copy()

    df = df.reset_index(drop=True)

    return wb, df, sheet_names, header_row_idx


def auto_detect_columns(columns: List[str]) -> Dict[str, Optional[str]]:
    """
    Intelligently maps standard Vietnamese educational column names.
    Supports both:
    - Mid-term files (Giữa kỳ: XL GK1, XL GK2, Nhận xét...)
    - End-term files (Cuối kỳ: KT CK1 [Điểm số], XL CK1 [Mức xếp loại], Nhận xét...)
    """
    mapping: Dict[str, Optional[str]] = {
        "name": None,
        "score": None,      # Cột điểm kiểm tra định kỳ (KT CK1, KT GK1, Điểm KT, Điểm số...)
        "level": None,      # Cột mức xếp loại / hoàn thành (XL GK1, XL CK1, Mức đạt...)
        "note": None,
        "comment": None,
    }

    # Exclude internal helper columns
    clean_cols = [c for c in columns if c != "_excel_row"]
    cols_lower = [str(c).strip().lower() for c in clean_cols]

    # 1. Name column patterns
    name_patterns = [
        r"họ\s*và\s*tên",
        r"họ\s*tên",
        r"tên\s*học\s*sinh",
        r"họ\s*đệm\s*và\s*tên",
        r"họ\s*(?:và\s*)?chữ\s*lót",
        r"họ\s*đệm",
        r"\btên\b"
    ]
    for pat in name_patterns:
        for idx, cl in enumerate(cols_lower):
            if re.search(pat, cl):
                mapping["name"] = clean_cols[idx]
                break
        if mapping["name"]:
            break

    # 2. Numerical exam score column patterns (KT CK1, KT GK1, Điểm KT, Điểm số, etc.)
    score_patterns = [
        r"kt\s*(ck|gk|hk|cn|[12])\b",
        r"điểm\s*(số|ktđk|kt|thi|đgk|hk|chính\s*thức)?",
        r"\bktđk\b",
        r"\bđgk\b",
        r"\bđiểm\b",
        r"\bkt\b"
    ]
    for pat in score_patterns:
        for idx, cl in enumerate(cols_lower):
            if cl != str(mapping["name"]).lower() and re.search(pat, cl):
                mapping["score"] = clean_cols[idx]
                break
        if mapping["score"]:
            break

    # 3. Level column patterns (XL GK1, XL CK1, Mức đạt, Xếp loại, etc.)
    level_patterns = [
        r"xl\s*(ck|gk|hk|cn|[12])\b",
        r"mức\s*(đạt|hoàn\s*thành|đánh\s*giá|xếp\s*loại)?",
        r"xếp\s*loại",
        r"kết\s*quả",
        r"\bxl\b",
        r"\bmức\b",
        r"\bmđ\b"
    ]
    for pat in level_patterns:
        for idx, cl in enumerate(cols_lower):
            if (
                cl != str(mapping["name"]).lower()
                and clean_cols[idx] != mapping["score"]
                and re.search(pat, cl)
            ):
                mapping["level"] = clean_cols[idx]
                break
        if mapping["level"]:
            break

    # Cross-fallbacks for compatibility:
    # If file only has XL GK1 (Giữa kỳ), score column can point to XL GK1
    if not mapping["score"] and mapping["level"]:
        mapping["score"] = mapping["level"]
    # If file only has KT CK1 / Điểm số (Cuối kỳ), level column can point to score
    if not mapping["level"] and mapping["score"]:
        mapping["level"] = mapping["score"]

    # 4. Note column patterns
    note_patterns = [
        r"ghi\s*chú",
        r"lưu\s*ý",
        r"nhận\s*xét\s*thường\s*xuyên",
        r"đặc\s*điểm",
        r"\bnote\b"
    ]
    for pat in note_patterns:
        for idx, cl in enumerate(cols_lower):
            if re.search(pat, cl):
                mapping["note"] = clean_cols[idx]
                break
        if mapping["note"]:
            break

    # 5. Comment column patterns
    comment_patterns = [
        r"nhận\s*xét\s*(tt27|thông\s*tư\s*27|tiến\s*bộ|đánh\s*giá)?",
        r"lời\s*phê",
        r"đánh\s*giá\s*thường\s*xuyên",
        r"\bnhận\s*xét\b",
        r"\bcomment\b"
    ]
    for pat in comment_patterns:
        for idx, cl in enumerate(cols_lower):
            if clean_cols[idx] != mapping["note"] and re.search(pat, cl):
                mapping["comment"] = clean_cols[idx]
                break
        if mapping["comment"]:
            break

    return mapping


def update_excel_workbook(
    wb: openpyxl.Workbook,
    sheet_name: str,
    header_row_idx: int,
    target_column_mode: str,  # "existing" or "new"
    existing_col_idx: Optional[int] = None,  # 1-based col index
    new_col_name: str = "Nhận xét Thông tư 27",
    row_comments: Optional[List[Tuple[int, str]]] = None,  # list of (row_index_or_excel_row, comment_text)
    existing_col_name: Optional[str] = None
) -> bytes:
    """
    Applies comments into openpyxl worksheet preserving original styles.
    Supports either relative DataFrame row indices or absolute 1-based Excel row indices.
    Returns bytes of updated Excel workbook.
    """
    if row_comments is None:
        row_comments = []

    ws = wb[sheet_name]

    # Resolve existing column index if name was provided
    if target_column_mode == "existing" and existing_col_idx is None and existing_col_name:
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=header_row_idx, column=c).value or "").strip()
            if val.lower() == existing_col_name.strip().lower():
                existing_col_idx = c
                break

    # Determine target column index
    if target_column_mode == "new" or existing_col_idx is None:
        target_col_idx = ws.max_column + 1
        # Style the new header matching the adjacent header cell
        header_cell = ws.cell(row=header_row_idx, column=target_col_idx, value=new_col_name)
        prev_header = ws.cell(row=header_row_idx, column=max(1, target_col_idx - 1))

        # Clone styles
        if prev_header.font:
            header_cell.font = Font(
                name=prev_header.font.name or "Segoe UI",
                size=prev_header.font.size or 11,
                bold=True,
                color=prev_header.font.color or "FFFFFF"
            )
        else:
            header_cell.font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")

        if prev_header.fill and prev_header.fill.fill_type:
            try:
                header_cell.fill = copy(prev_header.fill)
            except Exception:
                header_cell.fill = PatternFill(start_color="1B3B6F", end_color="1B3B6F", fill_type="solid")
        else:
            header_cell.fill = PatternFill(start_color="1B3B6F", end_color="1B3B6F", fill_type="solid")

        header_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if prev_header.border:
            try:
                header_cell.border = copy(prev_header.border)
            except Exception:
                pass

        col_letter = get_column_letter(target_col_idx)
        ws.column_dimensions[col_letter].width = 45
    else:
        target_col_idx = existing_col_idx
        col_letter = get_column_letter(target_col_idx)
        # Only set default width if column width is not already set in original template
        if ws.column_dimensions[col_letter].width is None or ws.column_dimensions[col_letter].width == 0:
            ws.column_dimensions[col_letter].width = 40

    # Determine data rows:
    start_data_row = header_row_idx + 1

    for row_key, comment in row_comments:
        # Check if row_key is an absolute 1-based Excel row number or 0-indexed df index
        if row_key >= start_data_row:
            ws_row = row_key
        else:
            ws_row = start_data_row + row_key

        if ws_row > ws.max_row:
            continue

        clean_comment = str(comment).strip()
        # If comment contains dict or json representation by accident, extract clean text:
        if clean_comment.startswith("{") and ("Nhận xét" in clean_comment or "comment" in clean_comment):
            try:
                import ast
                parsed = ast.literal_eval(clean_comment)
                if isinstance(parsed, dict):
                    clean_comment = (
                        parsed.get("Nhận xét Thông tư 27 (Có thể chỉnh sửa)")
                        or parsed.get("Nhận xét Thông tư 27")
                        or parsed.get("comment")
                        or clean_comment
                    )
            except Exception:
                pass

        cell = ws.cell(row=ws_row, column=target_col_idx, value=str(clean_comment).strip())
        # Format comment cell: preserve existing font and border, enable wrap text
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        # Do not force arbitrary row heights if already set in original template
        if ws.row_dimensions[ws_row].height is None:
            ws.row_dimensions[ws_row].height = 28

        # Copy border and font from adjacent cell only if not already present on cell
        ref_cell = ws.cell(row=ws_row, column=max(1, target_col_idx - 1))
        if not cell.border and ref_cell.border:
            try:
                cell.border = copy(ref_cell.border)
            except Exception:
                pass
        if not cell.font and ref_cell.font:
            try:
                cell.font = copy(ref_cell.font)
            except Exception:
                pass

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
