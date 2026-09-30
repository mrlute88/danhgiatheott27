"""
excel_handler.py - Robust Excel processing using openpyxl & pandas.
Preserves original Excel formatting, fonts, sheet structures, and formulas.
"""

import io
import re
from copy import copy
from typing import Dict, List, Optional, Tuple, Any
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd


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


def load_workbook_and_df(file_bytes: bytes, sheet_name: Optional[str] = None) -> Tuple[openpyxl.Workbook, pd.DataFrame, List[str], int]:
    """
    Loads openpyxl Workbook and generates corresponding Pandas DataFrame with detected header.
    Returns (workbook, dataframe, sheet_names, header_row_idx_1based)
    """
    buffer = io.BytesIO(file_bytes)
    wb = openpyxl.load_workbook(buffer, data_only=False)
    sheet_names = wb.sheetnames

    target_sheet_name = sheet_name if sheet_name in sheet_names else sheet_names[0]
    ws = wb[target_sheet_name]

    header_row_idx = detect_header_row(ws)

    # Read into pandas starting from header_row_idx (header is 0-indexed in pandas)
    buffer.seek(0)
    df = pd.read_excel(buffer, sheet_name=target_sheet_name, header=header_row_idx - 1)

    # Clean up dataframe columns: only drop columns that are completely unnamed AND completely empty
    cols_to_drop = [c for c in df.columns if ("Unnamed:" in str(c) or not str(c).strip()) and df[c].dropna().empty]
    if cols_to_drop:
        df = df.drop(columns=cols_to_drop)

    # Convert remaining unnamed columns to friendly names
    cleaned_cols = []
    for i, col in enumerate(df.columns):
        c_str = str(col).strip()
        if "Unnamed:" in c_str or not c_str:
            c_str = f"Cột_{get_column_letter(i + 1)}"
        cleaned_cols.append(c_str)
    df.columns = cleaned_cols

    # Drop completely empty rows
    df = df.dropna(how="all")

    return wb, df, sheet_names, header_row_idx


def auto_detect_columns(columns: List[str]) -> Dict[str, Optional[str]]:
    """
    Intelligently maps standard Vietnamese educational column names.
    """
    mapping = {
        "name": None,
        "score": None,
        "note": None,
        "comment": None,
    }

    cols_lower = [str(c).strip().lower() for c in columns]

    # Name column patterns
    name_patterns = [r"họ\s*và\s*tên", r"họ\s*tên", r"tên\s*học\s*sinh", r"họ\s*đệm\s*và\s*tên", r"\btên\b"]
    for pat in name_patterns:
        for idx, cl in enumerate(cols_lower):
            if re.search(pat, cl):
                mapping["name"] = columns[idx]
                break
        if mapping["name"]:
            break

    # Score/level column patterns
    score_patterns = [
        r"điểm\s*(số|ktđk|thi|đgk|hk|chính\s*thức)?",
        r"mức\s*(đạt|hoàn\s*thành|đánh\s*giá)",
        r"kết\s*quả",
        r"xếp\s*loại",
        r"\bđiểm\b",
        r"\bmức\b"
    ]
    for pat in score_patterns:
        for idx, cl in enumerate(cols_lower):
            if cl != str(mapping["name"]).lower() and re.search(pat, cl):
                mapping["score"] = columns[idx]
                break
        if mapping["score"]:
            break

    # Note column patterns
    note_patterns = [r"ghi\s*chú", r"lưu\s*ý", r"nhận\s*xét\s*thường\s*xuyên", r"đặc\s*điểm", r"\bnote\b"]
    for pat in note_patterns:
        for idx, cl in enumerate(cols_lower):
            if re.search(pat, cl):
                mapping["note"] = columns[idx]
                break
        if mapping["note"]:
            break

    # Comment column patterns
    comment_patterns = [
        r"nhận\s*xét\s*(tt27|thông\s*tư\s*27|tiến\s*bộ|đánh\s*giá)?",
        r"lời\s*phê",
        r"đánh\s*giá\s*thường\s*xuyên",
        r"\bnhận\s*xét\b",
        r"\bcomment\b"
    ]
    for pat in comment_patterns:
        for idx, cl in enumerate(cols_lower):
            # Don't pick note column as comment if separate
            if columns[idx] != mapping["note"] and re.search(pat, cl):
                mapping["comment"] = columns[idx]
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
    row_comments: Optional[List[Tuple[int, str]]] = None,  # list of (df_row_index, comment_text)
    existing_col_name: Optional[str] = None
) -> bytes:
    """
    Applies comments into openpyxl worksheet preserving original styles.
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
        if ws.column_dimensions[col_letter].width < 35:
            ws.column_dimensions[col_letter].width = 45

    # Determine data rows:
    # row_comments is a list of (i, comment) where i is 0-indexed relative to DataFrame rows.
    # The first data row in worksheet is header_row_idx + 1.
    start_data_row = header_row_idx + 1

    for df_idx, comment in row_comments:
        ws_row = start_data_row + df_idx
        if ws_row > ws.max_row:
            break
        
        cell = ws.cell(row=ws_row, column=target_col_idx, value=str(comment).strip())
        # Format comment cell
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        # Ensure row height is sufficient for 2 lines
        if ws.row_dimensions[ws_row].height is None or ws.row_dimensions[ws_row].height < 28:
            ws.row_dimensions[ws_row].height = 28

        # Copy border and font size from adjacent cell
        ref_cell = ws.cell(row=ws_row, column=max(1, target_col_idx - 1))
        if ref_cell.border:
            try:
                cell.border = copy(ref_cell.border)
            except Exception:
                pass
        if ref_cell.font:
            try:
                cell.font = Font(
                    name=ref_cell.font.name or "Segoe UI",
                    size=ref_cell.font.size or 10,
                    color="1A202C"
                )
            except Exception:
                pass

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
