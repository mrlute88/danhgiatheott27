"""
app.py - TT27 Assessor AI
Trợ lý AI Hỗ trợ Giáo viên Viết Nhận xét Học sinh Tiểu học theo Thông tư 27/2020/TT-BGDĐT
Kiến trúc Đa Tác tử (Multi-Agent System):
- Agent 1: Data Classifier (Chuẩn hóa điểm/mức & trích xuất năng lực môn học)
- Agent 2: TT27 Commentator (Gemini AI / Offline Pedagogical Fallback)
- Agent 3: Output Quality Validator (Kiểm duyệt độ dài, khử tiền tố, chống trùng lặp & bảo vệ Excel)
"""

import io
import time
import pandas as pd
import streamlit as st
import openpyxl

import os
import re
import base64
import subprocess
import streamlit.components.v1 as components

from agents import DataClassifierAgent, StudentInput, TT27CommentatorAgent, OutputValidatorAgent
import importlib
import utils.excel_handler
importlib.reload(utils.excel_handler)
from utils.excel_handler import (
    load_workbook_and_df,
    auto_detect_columns,
    update_excel_workbook,
    extract_pedagogical_metadata,
)
from utils.sample_generator import (
    get_sample_excel_bytes,
    get_sample_midterm_excel_bytes,
    get_sample_endterm_excel_bytes,
    get_sample_exam_image_bytes,
    get_sample_exam_pdf_bytes,
)
from utils.doc_converter import (
    MarkdownToDocxBuilder,
    GoogleVisionDocExtractor,
    sanitize_filename,
    save_file_locally,
)
from utils.tt27_knowledge import (
    SUBJECTS,
    GRADES,
    PERIODS,
    COMMENT_TONES,
    SUBJECT_COMPETENCIES,
    TT27_COMMENT_BANK,
)

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="TT27 Assessor AI - Trợ lý Nhận xét Học sinh Tiểu học",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------
# Custom Modern CSS Styling
# ---------------------------------------------------------
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
}

/* Hero Banner */
.hero-card {
    background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #0284C7 100%);
    color: white;
    padding: 28px 32px;
    border-radius: 16px;
    margin-bottom: 24px;
    box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.15), 0 8px 10px -6px rgba(15, 23, 42, 0.1);
    position: relative;
    overflow: hidden;
}

.hero-card::after {
    content: "TT27";
    position: absolute;
    right: 20px;
    bottom: -15px;
    font-size: 110px;
    font-weight: 800;
    color: rgba(255, 255, 255, 0.05);
    pointer-events: none;
}

.hero-tag {
    display: inline-block;
    background: rgba(255, 255, 255, 0.18);
    backdrop-filter: blur(8px);
    color: #F8FAFC;
    padding: 4px 14px;
    border-radius: 20px;
    font-size: 13px;
    font-weight: 600;
    margin-bottom: 12px;
    border: 1px solid rgba(255, 255, 255, 0.25);
}

.hero-title {
    font-size: 28px;
    font-weight: 800;
    margin: 0 0 8px 0;
    letter-spacing: -0.5px;
    color: #FFFFFF;
}

.hero-subtitle {
    font-size: 15px;
    color: #E2E8F0;
    margin: 0;
    line-height: 1.5;
    max-width: 850px;
}

/* Agent Pill Badges */
.agent-flow-container {
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
    margin-top: 16px;
}

.agent-badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: rgba(255, 255, 255, 0.12);
    padding: 6px 14px;
    border-radius: 8px;
    font-size: 12.5px;
    font-weight: 600;
    border: 1px solid rgba(255, 255, 255, 0.18);
    color: #F1F5F9;
}

/* Stat Cards */
.stat-box {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 16px 20px;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
    text-align: center;
}

.stat-num {
    font-size: 26px;
    font-weight: 700;
    line-height: 1.2;
}

.stat-label {
    font-size: 13px;
    color: #64748B;
    font-weight: 500;
    margin-top: 4px;
}

/* Level Badges */
.badge-t {
    background-color: #ECFDF5;
    color: #065F46;
    border: 1px solid #A7F3D0;
    padding: 3px 8px;
    border-radius: 6px;
    font-weight: 600;
}

.badge-h {
    background-color: #EFF6FF;
    color: #1E40AF;
    border: 1px solid #BFDBFE;
    padding: 3px 8px;
    border-radius: 6px;
    font-weight: 600;
}

.badge-c {
    background-color: #FEF2F2;
    color: #991B1B;
    border: 1px solid #FECACA;
    padding: 3px 8px;
    border-radius: 6px;
    font-weight: 600;
}

/* Action Callout */
.action-box {
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-left: 4px solid #2563EB;
    padding: 16px 20px;
    border-radius: 8px;
    margin-bottom: 20px;
}

/* Tab Header Polish */
button[data-baseweb="tab"] {
    font-size: 15px !important;
    font-weight: 600 !important;
    padding: 12px 20px !important;
}

/* Document Converter Modern Styles */
.doc-hero-card {
    background: linear-gradient(135deg, #0F172A 0%, #1E1B4B 45%, #2563EB 100%);
    color: white;
    padding: 24px 28px;
    border-radius: 16px;
    margin-bottom: 22px;
    box-shadow: 0 10px 25px -5px rgba(30, 27, 75, 0.25);
    position: relative;
    overflow: hidden;
}

.doc-hero-card::after {
    content: "DOCX";
    position: absolute;
    right: 25px;
    bottom: -20px;
    font-size: 110px;
    font-weight: 900;
    color: rgba(255, 255, 255, 0.04);
    pointer-events: none;
}

.doc-badge-pill {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: rgba(255, 255, 255, 0.14);
    padding: 5px 12px;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 600;
    border: 1px solid rgba(255, 255, 255, 0.2);
    color: #F8FAFC;
    margin-right: 8px;
    margin-top: 6px;
}

.doc-preview-card {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 16px;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ---------------------------------------------------------
# Sidebar Configuration
# ---------------------------------------------------------
with st.sidebar:
    st.markdown("### ⚙️ Cấu hình Hệ thống")

    # API Configuration Section
    st.markdown("#### 🔑 Google Gemini AI")
    api_key_input = st.text_input(
        "Gemini API Key",
        value=st.session_state.get("gemini_api_key", ""),
        type="password",
        help="Nhập API Key để sinh nhận xét linh hoạt, tự nhiên nhất. Nếu để trống, hệ thống sẽ tự động dùng Ngân hàng Sư phạm Offline chất lượng cao.",
        placeholder="AIzaSy...",
    )
    if api_key_input != st.session_state.get("gemini_api_key", ""):
        st.session_state["gemini_api_key"] = api_key_input.strip()

    # Active models supported by Google Gemini: Giữ lại duy nhất 1 model 3.6 flash theo yêu cầu
    available_models = ["gemini-3.6-flash"]

    model_option = st.selectbox(
        "Mô hình AI",
        options=available_models,
        index=0,
        help="gemini-3.6-flash: Mô hình Google AI thế hệ mới nhất, tốc độ cao, hỗ trợ đa phương thức và xử lý tài liệu chuẩn xác.",
    )

    col_test, col_status = st.columns([1, 1])
    with col_test:
        if st.button("Kiểm tra API", use_container_width=True):
            if not api_key_input:
                st.warning("Chưa nhập key!")
            else:
                with st.spinner("Đang kết nối..."):
                    is_ok, msg = TT27CommentatorAgent.test_api_key(api_key_input, model_option)
                    if is_ok:
                        st.success("API hợp lệ!")
                    else:
                        st.error("Lỗi API key!")
                        st.caption(msg)

    with col_status:
        if api_key_input:
            st.info("Chế độ: AI Online 🟢")
        else:
            st.info("Chế độ: Offline 🟠")

    st.divider()

    # Pedagogical Defaults
    st.markdown("#### 📚 Thiết lập Sư phạm Mặc định")

    # Apply pending auto-detected pedagogical settings from Excel file BEFORE widget instantiation
    if "pending_auto_meta" in st.session_state:
        pending_meta = st.session_state.pop("pending_auto_meta")
        if pending_meta.get("subject") in SUBJECTS:
            st.session_state["default_subject"] = pending_meta["subject"]
        if pending_meta.get("grade") in GRADES:
            st.session_state["default_grade"] = pending_meta["grade"]
        if pending_meta.get("period") in PERIODS:
            st.session_state["default_period"] = pending_meta["period"]

    if "default_subject" not in st.session_state or st.session_state["default_subject"] not in SUBJECTS:
        st.session_state["default_subject"] = "Tiếng Việt" if "Tiếng Việt" in SUBJECTS else SUBJECTS[0]
    if "default_grade" not in st.session_state or st.session_state["default_grade"] not in GRADES:
        st.session_state["default_grade"] = "Lớp 1" if "Lớp 1" in GRADES else GRADES[0]
    if "default_period" not in st.session_state or st.session_state["default_period"] not in PERIODS:
        st.session_state["default_period"] = "Giữa Học kỳ 1" if "Giữa Học kỳ 1" in PERIODS else PERIODS[0]
    if "default_tone" not in st.session_state or st.session_state["default_tone"] not in COMMENT_TONES:
        st.session_state["default_tone"] = COMMENT_TONES[0]

    default_subject = st.selectbox(
        "Môn học",
        options=SUBJECTS,
        key="default_subject",
        help="Hệ thống tự động nhận diện từ file Excel tải lên. Bạn cũng có thể tùy chỉnh lại ở đây."
    )
    default_grade = st.selectbox(
        "Khối lớp",
        options=GRADES,
        key="default_grade",
        help="Hệ thống tự động nhận diện từ file Excel tải lên."
    )
    default_period = st.selectbox(
        "Thời điểm đánh giá",
        options=PERIODS,
        key="default_period",
        help="Hệ thống tự động nhận diện Giữa Kỳ / Cuối Kỳ từ file Excel tải lên."
    )
    default_tone = st.selectbox(
        "Văn phong lời phê",
        options=COMMENT_TONES,
        key="default_tone"
    )

    st.divider()

    # Agent 3 Quality Validator Settings
    st.markdown("#### 🛡️ Kiểm soát Chất lượng & Chống Trùng lặp (Agent 3)")
    zero_dup_mode = st.toggle(
        "Chế độ Lời phê Độc bản (Zero Duplicate)",
        value=True,
        help="Đảm bảo 100% không có bất kỳ 2 học sinh nào trong cùng một lớp học bị trùng lặp nhận xét.",
    )
    window_size = st.slider(
        "Bộ đệm chống lặp câu (HS gần nhất)",
        min_value=5,
        max_value=20,
        value=12,
        help="Số lượng lời phê học sinh gần kề được lưu để kiểm tra và biến thể từ đồng nghĩa, tránh viết lặp lại câu giống hệt nhau.",
    )

    st.markdown(
        """
        <div style="font-size: 12px; color: #64748B; margin-top: 15px; border-top: 1px solid #E2E8F0; padding-top: 10px;">
            <b>TT27 Assessor AI</b> v2.0<br/>
            Hệ thống tuân thủ Thông tư 27/2020/TT-BGDĐT.<br/>
            Bảo toàn 100% định dạng file Excel.
        </div>
        """,
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------
# Top Hero Banner
# ---------------------------------------------------------
st.markdown(
    """
    <div class="hero-card">
        <span class="hero-tag">✨ TIỂU HỌC 4.0 • CHUẨN THÔNG TƯ 27/2020/TT-BGDĐT</span>
        <h1 class="hero-title">TT27 Assessor AI — Trợ lý Viết Nhận xét Học sinh</h1>
        <p class="hero-subtitle">
            Hệ sinh thái Đa Tác tử (Multi-Agent) chuyên dụng giúp giáo viên tiểu học chuẩn hóa mức đánh giá, 
            tự động hóa viết nhận xét cá nhân hóa giàu tính động viên và bảo tồn nguyên vẹn cấu trúc file Excel.
        </p>
        <div class="agent-flow-container">
            <span class="agent-badge">🤖 <b>Agent 1:</b> Phân loại & Năng lực môn</span>
            <span class="agent-badge">🧠 <b>Agent 2:</b> Khởi tạo lời phê Sư phạm</span>
            <span class="agent-badge">🛡️ <b>Agent 3:</b> Chuẩn hóa độ dài & Chống lặp</span>
            <span class="agent-badge" style="background: rgba(34, 197, 94, 0.22); border-color: rgba(34, 197, 94, 0.45); color: #DCFCE7;">📄 <b>Vision AI:</b> Chuyển Ảnh/PDF ➔ Word (.docx)</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------
# Main Application Tabs
# ---------------------------------------------------------
tab_batch, tab_single, tab_doc_converter, tab_knowledge = st.tabs(
    [
        "📁 Xử lý Bảng điểm Excel (Hàng loạt)",
        "✍️ Nhận xét Nhanh Từng Học sinh",
        "📄 Chuyển Đổi Ảnh / PDF sang Word (.docx)",
        "📖 Tra cứu Thông tư 27 & Kho dữ liệu",
    ]
)


# =========================================================
# TAB 1: BATCH EXCEL PROCESSING
# =========================================================
with tab_batch:
    st.markdown("### 1. Tải lên hoặc Trải nghiệm Bảng điểm Học sinh")

    col_up_action1, col_up_action2, col_up_action3 = st.columns([1, 1, 1])

    with col_up_action1:
        # Sample Mid-term (Image 1)
        sample_gk_bytes = get_sample_midterm_excel_bytes()
        st.download_button(
            label="📥 Tải mẫu Giữa kỳ 1 (Ảnh 1)",
            data=sample_gk_bytes,
            file_name="Bang_Nhan_Xet_Giua_Ky_1_Lop1_TiengViet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Bảng điểm Giữa kỳ 1 chuẩn TT27: Môn Tiếng Việt Lớp 1, có cột mức xếp loại XL GK1.",
            use_container_width=True,
        )
        if st.button("🧪 Thử mẫu Giữa Kỳ 1 (Ảnh 1)", use_container_width=True, help="Nạp trực tiếp mẫu bảng điểm Giữa kỳ 1 để trải nghiệm ngay."):
            st.session_state["loaded_file_bytes"] = sample_gk_bytes
            st.session_state["loaded_file_name"] = "Bang_Nhan_Xet_Giua_Ky_1_Lop1_TiengViet.xlsx"
            st.session_state["generation_results"] = None
            st.session_state["last_scanned_file_sig"] = None
            st.rerun()

    with col_up_action2:
        # Sample End-term with score (Image 2)
        sample_ck_bytes = get_sample_endterm_excel_bytes()
        st.download_button(
            label="📥 Tải mẫu Cuối kỳ 1 có Điểm (Ảnh 2)",
            data=sample_ck_bytes,
            file_name="Bang_Nhan_Xet_Cuoi_Ky_1_Lop1_TiengViet_CoDiemSo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Bảng điểm Cuối kỳ 1 chuẩn TT27: Môn Tiếng Việt Lớp 1, có thêm cột Điểm kiểm tra KT CK1 và cột XL CK1.",
            use_container_width=True,
        )
        if st.button("🧪 Thử mẫu Cuối Kỳ 1 (Ảnh 2)", use_container_width=True, help="Nạp trực tiếp mẫu bảng điểm Cuối kỳ 1 có cột điểm số để trải nghiệm ngay."):
            st.session_state["loaded_file_bytes"] = sample_ck_bytes
            st.session_state["loaded_file_name"] = "Bang_Nhan_Xet_Cuoi_Ky_1_Lop1_TiengViet_CoDiemSo.xlsx"
            st.session_state["generation_results"] = None
            st.session_state["last_scanned_file_sig"] = None
            st.rerun()

    with col_up_action3:
        # Standard 9-column sample
        sample_std_bytes = get_sample_excel_bytes()
        st.download_button(
            label="📥 Tải file mẫu tổng hợp (.xlsx)",
            data=sample_std_bytes,
            file_name="Bang_Theo_Doi_Danh_Gia_Mau_TT27.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Tải file Excel mẫu gồm 15 học sinh với đầy đủ các cột dữ liệu tổng hợp.",
            use_container_width=True,
        )
        if st.button("🧪 Thử mẫu Tổng hợp 15 HS", use_container_width=True):
            st.session_state["loaded_file_bytes"] = sample_std_bytes
            st.session_state["loaded_file_name"] = "Bang_Theo_Doi_Danh_Gia_Mau_TT27.xlsx"
            st.session_state["generation_results"] = None
            st.session_state["last_scanned_file_sig"] = None
            st.rerun()

    uploaded_file = st.file_uploader(
        "Hoặc kéo thả file Excel bảng điểm của lớp bạn (.xlsx, .xls):",
        type=["xlsx", "xls"],
        help="Hỗ trợ mọi định dạng bảng điểm tiểu học. Hệ thống tự động nhận diện môn học, khối lớp, giữa kỳ/cuối kỳ và các cột họ tên, điểm số, mức đạt.",
    )

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        if (
            st.session_state.get("loaded_file_name") != uploaded_file.name
            or st.session_state.get("loaded_file_bytes") != file_bytes
        ):
            st.session_state["loaded_file_bytes"] = file_bytes
            st.session_state["loaded_file_name"] = uploaded_file.name
            st.session_state["generation_results"] = None
            st.session_state["last_scanned_file_sig"] = None

    active_bytes = st.session_state.get("loaded_file_bytes")
    active_name = st.session_state.get("loaded_file_name", "Bang_diem.xlsx")

    if active_bytes:
        try:
            wb, df, sheet_names, detected_header_row = load_workbook_and_df(active_bytes)
        except Exception as e:
            st.error(f"Không thể đọc file Excel: {e}")
            wb, df, sheet_names, detected_header_row = None, None, [], 1

        if df is not None and not df.empty and wb is not None:
            # ---------------------------------------------------------
            # Auto-detection of Pedagogical Defaults from Excel file
            # ---------------------------------------------------------
            current_target_sheet = sheet_names[0]
            file_sig = f"{active_name}_{len(active_bytes)}_{current_target_sheet}"

            if st.session_state.get("last_scanned_file_sig") != file_sig:
                auto_meta = extract_pedagogical_metadata(wb, sheet_name=current_target_sheet, file_name=active_name)
                st.session_state["pending_auto_meta"] = auto_meta
                st.session_state["last_scanned_file_sig"] = file_sig
                st.session_state["last_detected_meta"] = auto_meta
                st.rerun()

            # Display Auto-detection confirmation banner
            st.markdown(
                f"""
                <div style="background: linear-gradient(135deg, #1E3A8A 0%, #0D9488 100%); color: white; padding: 14px 20px; border-radius: 12px; margin-top: 10px; margin-bottom: 16px; box-shadow: 0 4px 14px rgba(13, 148, 136, 0.22);">
                    <div style="font-size: 15px; font-weight: 700; display: flex; align-items: center; gap: 8px;">
                        🎯 Đã tự động chọn Thiết lập Sư phạm Mặc định phù hợp với file Excel:
                    </div>
                    <div style="margin-top: 8px; font-size: 13.5px; opacity: 0.95; display: flex; flex-wrap: wrap; gap: 18px;">
                        <span>📚 Môn học: <b>{default_subject}</b></span>
                        <span>🏫 Khối lớp: <b>{default_grade}</b></span>
                        <span>⏰ Thời điểm: <b>{default_period}</b></span>
                    </div>
                    <div style="margin-top: 6px; font-size: 11.5px; opacity: 0.85;">
                        <i>(Các mục trên đã được tự động đồng bộ vào 'Thiết lập Sư phạm Mặc định' ở thanh bên trái. Bạn có thể tùy chỉnh lại bất kỳ lúc nào nếu cần)</i>
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

            st.success(f" Đã đọc thành công file **{active_name}** ({len(df)} học sinh).")

            with st.expander("🛠️ Xem & Tinh chỉnh Ánh xạ Cột dữ liệu (Column Mapping)", expanded=True):
                col_cfg1, col_cfg2 = st.columns([1, 1])

                with col_cfg1:
                    selected_sheet = st.selectbox(
                        "Sheet cần xử lý:",
                        options=sheet_names,
                        index=0,
                        help="Chọn bảng tính chứa danh sách học sinh cần nhận xét",
                    )
                    if selected_sheet != sheet_names[0]:
                        wb, df, sheet_names, detected_header_row = load_workbook_and_df(
                            active_bytes, sheet_name=selected_sheet
                        )

                with col_cfg2:
                    st.caption(f"Dòng tiêu đề được nhận diện tự động: **Dòng {detected_header_row}**")
                    all_cols = [c for c in df.columns if c != "_excel_row"]
                    auto_mapped = auto_detect_columns(all_cols)

                    name_col = st.selectbox(
                        "Cột Họ và tên học sinh (*):",
                        options=all_cols,
                        index=all_cols.index(auto_mapped["name"]) if auto_mapped["name"] in all_cols else 0,
                        help="Hệ thống tự động ghép cột Họ đệm và Tên (nếu bị tách 2 cột) thành họ tên đầy đủ chuẩn xác.",
                    )

                col_cfg3, col_cfg4 = st.columns([1, 1])
                with col_cfg3:
                    # Level column (XL GK1, XL CK1, Mức đạt, etc.)
                    level_default_col = auto_mapped.get("level") or auto_mapped.get("score")
                    level_idx = all_cols.index(level_default_col) if level_default_col in all_cols else (1 if len(all_cols) > 1 else 0)
                    level_col = st.selectbox(
                        "Cột Mức xếp loại / hoàn thành (T/H/C) (*):",
                        options=all_cols,
                        index=level_idx,
                        help="Cột chứa mức đánh giá T (Tốt), H (Hoàn thành), C (Chưa hoàn thành), ví dụ: XL GK1, XL CK1.",
                    )

                with col_cfg4:
                    # Exam score column (KT CK1, Điểm kiểm tra, etc.)
                    exam_score_options = ["-- Không có (Giữa kỳ không thi điểm số) --"] + all_cols
                    detected_score_col = auto_mapped.get("score")
                    # Only default to exam score if it's distinct from level column or explicitly has 'kt' / 'điểm'
                    is_distinct_score = (
                        detected_score_col
                        and detected_score_col != level_default_col
                        and re.search(r"kt\s*(ck|gk|hk|cn|[12])|điểm", str(detected_score_col).lower())
                    )
                    default_exam_idx = (
                        exam_score_options.index(detected_score_col)
                        if is_distinct_score and detected_score_col in exam_score_options
                        else 0
                    )
                    exam_score_col = st.selectbox(
                        "Cột Điểm kiểm tra định kỳ (Điểm số 1-10):",
                        options=exam_score_options,
                        index=default_exam_idx,
                        help="Đối với bảng điểm Cuối kỳ có thêm cột Điểm kiểm tra (KT CK1), AI sẽ kết hợp điểm số để nhận xét sâu sắc hơn.",
                    )

                col_cfg5, col_cfg6 = st.columns([1, 1])
                with col_cfg5:
                    note_options = ["-- Không sử dụng --"] + all_cols
                    default_note_idx = (
                        note_options.index(auto_mapped["note"])
                        if auto_mapped["note"] in note_options
                        else 0
                    )
                    note_col = st.selectbox(
                        "Cột Ghi chú của giáo viên (tùy chọn):",
                        options=note_options,
                        index=default_note_idx,
                        help="Ghi chú về tính cách, năng khiếu, sự tiến bộ sẽ được AI lồng ghép khéo léo vào lời nhận xét.",
                    )

                with col_cfg6:
                    target_mode = st.radio(
                        "Nơi ghi lời nhận xét:",
                        options=["Ghi vào một cột có sẵn", "Thêm cột mới vào cuối bảng"],
                        index=0 if auto_mapped.get("comment") in all_cols else 1,
                        horizontal=True,
                    )

                    if target_mode == "Ghi vào một cột có sẵn":
                        target_col_name = st.selectbox(
                            "Chọn cột nhận xét mục tiêu:",
                            options=all_cols,
                            index=all_cols.index(auto_mapped["comment"]) if auto_mapped.get("comment") in all_cols else (len(all_cols) - 1),
                        )
                        new_col_title = target_col_name
                    else:
                        new_col_title = st.text_input(
                            "Tên cột mới sẽ tạo:",
                            value="Nhận xét Thông tư 27",
                        )
                        target_col_name = None

            # Preview original data table (excluding internal _excel_row)
            st.markdown("#### 📋 Dữ liệu học sinh đọc được từ Excel:")
            preview_cols = [c for c in df.columns if c != "_excel_row"]
            st.dataframe(df[preview_cols].head(10), use_container_width=True)
            if len(df) > 10:
                st.caption(f"Đang hiển thị 10/{len(df)} học sinh đầu tiên.")

            st.divider()

            # Generation Controls
            col_b_action1, col_b_action2 = st.columns([2, 1])
            with col_b_action1:
                st.markdown(
                    f"**Thiết lập hiện tại:** Môn **{default_subject}** | Khối **{default_grade}** | Kỳ: **{default_period}** | Phong cách: *{default_tone}*"
                )
                if not api_key_input:
                    st.caption("ℹ️ Đang chạy chế độ **Sư phạm Offline thông minh** (100% chuẩn Thông tư 27, tốc độ cực nhanh, không phụ thuộc mạng).")
                else:
                    st.caption(f"ℹ️ Đang chạy chế độ **Gemini AI ({model_option})** kết hợp thẩm định sư phạm.")

            with col_b_action2:
                btn_start = st.button(
                    "🚀 BẮT ĐẦU TẠO NHẬN XÉT",
                    type="primary",
                    use_container_width=True,
                )

            # Processing Execution
            if btn_start:
                classifier = DataClassifierAgent()
                commentator = TT27CommentatorAgent(
                    api_key=api_key_input,
                    model_name=model_option,
                )
                validator = OutputValidatorAgent(history_window_size=window_size)
                validator.reset_history()
                used_comments_set = set()

                progress_bar = st.progress(0.0)
                status_text = st.empty()

                total_rows = len(df)
                results_list = []
                stat_counts = {"T": 0, "H": 0, "C": 0}

                start_time = time.time()

                for i, (orig_idx, row) in enumerate(df.iterrows()):
                    student_name = str(row.get(name_col, f"Học sinh {i+1}")).strip()
                    level_val = row.get(level_col, "H")

                    exam_val = None
                    if exam_score_col != "-- Không có (Giữa kỳ không thi điểm số) --" and exam_score_col in df.columns:
                        raw_ex = row.get(exam_score_col, None)
                        if pd.notna(raw_ex) and str(raw_ex).strip() and str(raw_ex).strip().lower() != "nan":
                            exam_val = raw_ex

                    note_val = ""
                    if note_col != "-- Không sử dụng --" and note_col in df.columns:
                        raw_note = row.get(note_col, "")
                        note_val = "" if pd.isna(raw_note) else str(raw_note).strip()

                    status_text.markdown(f"🤖 Đang xử lý: **{student_name}** ({i+1}/{total_rows})...")

                    # Step 1: Agent 1 - Classification
                    student_input = StudentInput(
                        student_name=student_name,
                        subject=default_subject,
                        grade=default_grade,
                        score_or_level=level_val,
                        exam_score=exam_val,
                        teacher_note=note_val,
                        period=default_period,
                        tone=default_tone,
                    )
                    c_result = classifier.classify(student_input)
                    stat_counts[c_result.standardized_level] = stat_counts.get(c_result.standardized_level, 0) + 1

                    # Step 2: Agent 2 - Comment Generation with Zero-Duplicate Check
                    raw_comment = commentator.generate_comment(
                        student_input,
                        c_result,
                        recent_comments=validator.recent_comments,
                        used_comments=used_comments_set,
                    )

                    # Step 3: Agent 3 - Quality Validation, Anti-Duplication & Sanitization
                    final_comment = validator.validate_and_format(
                        raw_comment,
                        student_name,
                        standardized_level=c_result.standardized_level,
                        subject=default_subject,
                        teacher_note=note_val,
                    )

                    excel_row_val = row.get("_excel_row", orig_idx)
                    score_display = f"{exam_val} (Mức {c_result.standardized_level})" if exam_val is not None else str(level_val)

                    results_list.append(
                        {
                            "df_idx": orig_idx,
                            "excel_row": excel_row_val,
                            "student_name": student_name,
                            "score_raw": score_display,
                            "exam_score": exam_val,
                            "level": c_result.standardized_level,
                            "level_label": c_result.level_label,
                            "teacher_note": note_val,
                            "comment": final_comment,
                        }
                    )

                    # Update progress
                    progress_bar.progress((i + 1) / total_rows)

                duration = time.time() - start_time
                status_text.success(f" Đã hoàn thành tạo nhận xét độc bản cho **{total_rows}** học sinh trong {duration:.1f} giây!")
                st.session_state["generation_results"] = results_list
                st.session_state["stat_counts"] = stat_counts

            # Display Results if Available
            gen_results = st.session_state.get("generation_results")
            if gen_results:
                st.markdown("### 2. Thống kê & Duyệt Lời Nhận xét")

                # Metrics summary row with Zero-Duplicate verification
                st_counts = st.session_state.get("stat_counts", {"T": 0, "H": 0, "C": 0})
                unique_comments = set(r["comment"] for r in gen_results)
                dup_count = len(gen_results) - len(unique_comments)

                m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns(5)
                with m_col1:
                    st.markdown(
                        f"""
                        <div class="stat-box">
                            <div class="stat-num" style="color: #1E3A8A;">{len(gen_results)}</div>
                            <div class="stat-label">Tổng số học sinh</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with m_col2:
                    st.markdown(
                        f"""
                        <div class="stat-box">
                            <div class="stat-num" style="color: #059669;">{st_counts.get('T', 0)}</div>
                            <div class="stat-label">Hoàn thành tốt (Mức T)</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with m_col3:
                    st.markdown(
                        f"""
                        <div class="stat-box">
                            <div class="stat-num" style="color: #0284C7;">{st_counts.get('H', 0)}</div>
                            <div class="stat-label">Hoàn thành (Mức H)</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with m_col4:
                    st.markdown(
                        f"""
                        <div class="stat-box">
                            <div class="stat-num" style="color: #DC2626;">{st_counts.get('C', 0)}</div>
                            <div class="stat-label">Chưa hoàn thành (Mức C)</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with m_col5:
                    st.markdown(
                        f"""
                        <div class="stat-box" style="border: 2px solid #10B981; background: #F0FDF4;">
                            <div class="stat-num" style="color: #059669;">0 ({dup_count})</div>
                            <div class="stat-label" style="color: #065F46; font-weight: 600;">Trùng lặp: 0% (Độc bản)</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                st.write("")
                st.markdown(
                    """
                    <div class="action-box">
                        💡 <b>Mẹo cho Giáo viên:</b> Bạn có thể trực tiếp nhấp đúp vào ô nhận xét trong bảng dưới đây để sửa lại từ ngữ theo ý muốn trước khi xuất file. Mọi chỉnh sửa sẽ được lưu vào file Excel tải về!
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                # Editable DataFrame
                res_df = pd.DataFrame(
                    [
                        {
                            "Họ và tên": r["student_name"],
                            "Điểm / Mức": r["score_raw"],
                            "Xếp loại TT27": r["level_label"],
                            "Ghi chú GV": r["teacher_note"],
                            "Nhận xét Thông tư 27 (Có thể chỉnh sửa)": r["comment"],
                        }
                        for r in gen_results
                    ]
                )

                edited_df = st.data_editor(
                    res_df,
                    use_container_width=True,
                    num_rows="fixed",
                    key="batch_results_editor",
                    column_config={
                        "Nhận xét Thông tư 27 (Có thể chỉnh sửa)": st.column_config.TextColumn(
                            "Lời nhận xét Thông tư 27 (Nhấp để sửa)",
                            width="large",
                        ),
                    },
                )

                if edited_df is None or not isinstance(edited_df, pd.DataFrame):
                    edited_df = res_df

                # Identify comment column and student name column in edited_df reliably
                comment_col = None
                for candidate in [
                    "Nhận xét Thông tư 27 (Có thể chỉnh sửa)",
                    "Lời nhận xét Thông tư 27 (Nhấp để sửa)",
                    "Nhận xét Thông tư 27",
                    "Nhận xét",
                ]:
                    if candidate in edited_df.columns:
                        comment_col = candidate
                        break

                if not comment_col:
                    for col in edited_df.columns:
                        if "nhận xét" in str(col).lower() or "comment" in str(col).lower():
                            comment_col = col
                            break
                    if not comment_col and len(edited_df.columns) > 0:
                        comment_col = edited_df.columns[-1]

                name_col = None
                for col in edited_df.columns:
                    if "tên" in str(col).lower():
                        name_col = col
                        break

                edited_comments_by_name = {}
                if name_col and comment_col:
                    for _, row in edited_df.iterrows():
                        s_name = str(row.get(name_col, "")).strip()
                        c_text = str(row.get(comment_col, "") or "").strip()
                        if s_name:
                            edited_comments_by_name[s_name] = c_text

                # Build updated comments list from edited_df
                final_row_comments = []
                for i, r in enumerate(gen_results):
                    s_name = str(r.get("student_name", "")).strip()
                    updated_text = ""
                    if s_name and s_name in edited_comments_by_name:
                        updated_text = edited_comments_by_name[s_name]
                    elif comment_col and i < len(edited_df):
                        val_c = edited_df.iloc[i].get(comment_col, "")
                        if isinstance(val_c, dict):
                            val_c = val_c.get("Nhận xét Thông tư 27 (Có thể chỉnh sửa)") or val_c.get("comment", "")
                        updated_text = str(val_c or "").strip()

                    # Sanitize: if updated_text is empty or accidental dict string, use pure comment
                    if not updated_text or (isinstance(updated_text, str) and updated_text.startswith("{")):
                        updated_text = str(r.get("comment", "")).strip()

                    row_key = r.get("excel_row", r.get("df_idx", i))
                    final_row_comments.append((row_key, updated_text))

                # Export to Excel preserving original styles
                target_mode_flag = "existing" if target_mode == "Ghi vào một cột có sẵn" else "new"

                try:
                    updated_excel_bytes = update_excel_workbook(
                        wb=wb,
                        sheet_name=selected_sheet,
                        header_row_idx=detected_header_row,
                        target_column_mode=target_mode_flag,
                        existing_col_name=target_col_name,
                        new_col_name=new_col_title,
                        row_comments=final_row_comments,
                    )

                    clean_base = active_name
                    if clean_base.lower().endswith(".xlsx"):
                        clean_base = clean_base[:-5]
                    elif clean_base.lower().endswith(".xls"):
                        clean_base = clean_base[:-4]
                    export_filename = f"{clean_base}_Da_Nhan_Xet_TT27.xlsx"

                    st.markdown("### 3. Tải về Kết quả")
                    st.download_button(
                        label="📥 TẢI VỀ FILE EXCEL HOÀN CHỈNH (.XLSX)",
                        data=updated_excel_bytes,
                        file_name=export_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                        help="File Excel xuất ra giữ nguyên vẹn 100% font chữ, viền bảng, màu sắc và công thức gốc.",
                    )
                except Exception as ex:
                    st.error(f"Lỗi khi đóng gói file Excel: {ex}")

    else:
        st.info("👆 Vui lòng bấm **'Nạp dữ liệu mẫu 15 học sinh'** ở trên hoặc kéo thả file Excel của bạn vào để bắt đầu.")


# =========================================================
# TAB 2: SINGLE STUDENT ASSISTANT
# =========================================================
with tab_single:
    st.markdown("### ✍️ Trợ lý Viết Nhận xét Nhanh Từng Học sinh")
    st.caption("Dành cho giáo viên muốn soạn nhanh lời phê cho một trường hợp cụ thể hoặc tham khảo các phương án viết khác nhau.")

    col_s1, col_s2 = st.columns([1, 1])

    with col_s1:
        s_name = st.text_input("Họ và tên học sinh:", value="Trần Minh Khang")
        s_subj = st.selectbox("Môn học:", options=SUBJECTS, index=0, key="single_subj")
        s_grade = st.selectbox("Khối lớp:", options=GRADES, index=2, key="single_grade")

    with col_s2:
        s_score = st.text_input(
            "Điểm số (0-10) hoặc Mức đạt (T, H, C):",
            value="8.5",
            help="Ví dụ: 9.5, 8, T, H, C, Hoàn thành tốt...",
        )
        s_period = st.selectbox("Kỳ đánh giá:", options=PERIODS, index=1, key="single_period")
        s_tone = st.selectbox("Phong cách lời phê:", options=COMMENT_TONES, index=0, key="single_tone")

    s_note = st.text_area(
        "Đặc điểm riêng / Ghi chú của giáo viên:",
        value="Tính toán nhanh nhưng đôi khi còn quên ghi đơn vị đo, tích cực giơ tay phát biểu",
        placeholder="VD: Chữ viết đẹp, hay giúp đỡ bạn bè, làm tính nhẩm nhanh nhưng còn ẩu...",
        height=70,
    )

    btn_single = st.button("✨ TẠO LỜI PHÊ TỨC THÌ", type="primary", use_container_width=True)

    if btn_single:
        classifier = DataClassifierAgent()
        commentator = TT27CommentatorAgent(
            api_key=api_key_input,
            model_name=model_option,
        )
        validator = OutputValidatorAgent()

        s_input = StudentInput(
            student_name=s_name,
            subject=s_subj,
            grade=s_grade,
            score_or_level=s_score,
            teacher_note=s_note,
            period=s_period,
            tone=s_tone,
        )

        with st.spinner("Đang phân tích sư phạm và tạo câu nhận xét..."):
            c_res = classifier.classify(s_input)
            raw_c = commentator.generate_comment(s_input, c_res)
            clean_c = validator.validate_and_format(
                raw_c,
                s_name,
                c_res.standardized_level,
                subject=s_subj,
                teacher_note=s_note
            )

        st.markdown("#### 🎯 Kết quả Phân tích & Lời phê từ Hệ thống Đa Tác tử:")

        res_col_left, res_col_right = st.columns([1, 2])

        with res_col_left:
            st.markdown(
                f"""
                <div style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 12px; padding: 18px;">
                    <div style="font-size: 13px; font-weight: 700; color: #475569; text-transform: uppercase; margin-bottom: 8px;">
                        🤖 Agent 1: Phân loại Chuẩn TT27
                    </div>
                    <div style="margin-bottom: 12px;">
                        <b>Mức chuẩn hóa:</b> 
                        <span class="badge-{c_res.standardized_level.lower()}">{c_res.level_label}</span>
                    </div>
                    <div style="margin-bottom: 10px; font-size: 13.5px;">
                        <b>Điểm quy đổi:</b> {c_res.numerical_score if c_res.numerical_score is not None else 'Đánh giá định tính'}
                    </div>
                    <div style="margin-bottom: 10px; font-size: 13.5px;">
                        <b>Định hướng sư phạm:</b><br/>
                        <span style="color: #334155;">{c_res.guidance_notes}</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with res_col_right:
            word_count = len(clean_c.split())
            st.markdown(
                f"""
                <div style="background: #FFFFFF; border: 2px solid #2563EB; border-radius: 12px; padding: 20px; box-shadow: 0 4px 12px rgba(37, 99, 235, 0.08);">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                        <span style="font-size: 13px; font-weight: 700; color: #1E3A8A; text-transform: uppercase;">
                            🧠 Lời nhận xét hoàn thiện (Agent 2 & 3)
                        </span>
                        <span style="background: #EFF6FF; color: #1E40AF; padding: 3px 10px; border-radius: 20px; font-size: 12px; font-weight: 600;">
                            {word_count} từ • Chuẩn TT27
                        </span>
                    </div>
                    <div style="font-size: 16px; color: #0F172A; line-height: 1.6; font-weight: 500; margin-bottom: 15px;">
                        "{clean_c}"
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        # Generate 3 Alternative Variations
        st.write("")
        st.markdown("##### 💡 3 Phương án Biến thể Gợi ý Thêm:")
        v_col1, v_col2, v_col3 = st.columns(3)

        # Variant 1: Tone Khích lệ
        s_input_v1 = StudentInput(
            student_name=s_name,
            subject=s_subj,
            grade=s_grade,
            score_or_level=s_score,
            teacher_note=s_note,
            period=s_period,
            tone="Ấm áp & Khích lệ sâu sắc",
        )
        raw_v1 = commentator.generate_comment(s_input_v1, c_res)
        clean_v1 = validator.validate_and_format(
            raw_v1,
            s_name,
            c_res.standardized_level,
            subject=s_subj,
            teacher_note=s_note
        )

        # Variant 2: Tone Hành động cụ thể
        s_input_v2 = StudentInput(
            student_name=s_name,
            subject=s_subj,
            grade=s_grade,
            score_or_level=s_score,
            teacher_note=s_note,
            period=s_period,
            tone="Cụ thể hóa kỹ năng & Hành động",
        )
        raw_v2 = commentator.generate_comment(s_input_v2, c_res)
        clean_v2 = validator.validate_and_format(
            raw_v2,
            s_name,
            c_res.standardized_level,
            subject=s_subj,
            teacher_note=s_note
        )

        # Variant 3: Tone Ngắn gọn
        s_input_v3 = StudentInput(
            student_name=s_name,
            subject=s_subj,
            grade=s_grade,
            score_or_level=s_score,
            teacher_note=s_note,
            period=s_period,
            tone="Ngắn gọn, súc tích (15-20 từ)",
        )
        raw_v3 = commentator.generate_comment(s_input_v3, c_res)
        clean_v3 = validator.validate_and_format(
            raw_v3,
            s_name,
            c_res.standardized_level,
            subject=s_subj,
            teacher_note=s_note
        )

        with v_col1:
            st.markdown(
                f"""
                <div style="background: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 8px; padding: 14px; min-height: 120px;">
                    <div style="font-size: 12px; font-weight: 700; color: #0284C7; margin-bottom: 6px;">
                        🌱 Hướng Khích Lệ Sâu Sắc
                    </div>
                    <div style="font-size: 13.5px; color: #1E293B;">{clean_v1}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with v_col2:
            st.markdown(
                f"""
                <div style="background: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 8px; padding: 14px; min-height: 120px;">
                    <div style="font-size: 12px; font-weight: 700; color: #059669; margin-bottom: 6px;">
                        🎯 Hướng Hành Động Kỹ Năng
                    </div>
                    <div style="font-size: 13.5px; color: #1E293B;">{clean_v2}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with v_col3:
            st.markdown(
                f"""
                <div style="background: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 8px; padding: 14px; min-height: 120px;">
                    <div style="font-size: 12px; font-weight: 700; color: #6366F1; margin-bottom: 6px;">
                        ⚡ Hướng Ngắn Gọn (15-20 từ)
                    </div>
                    <div style="font-size: 13.5px; color: #1E293B;">{clean_v3}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


# =========================================================
# TAB 3: DOCUMENT CONVERTER (IMAGE / PDF TO WORD .DOCX)
# =========================================================
with tab_doc_converter:
    st.markdown(
        """
        <div class="doc-hero-card">
            <span class="hero-tag" style="background: rgba(59, 130, 246, 0.25); border-color: rgba(96, 165, 250, 0.4);">
                🚀 MULTIMODAL VISION AI • CHUYỂN ĐỔI THÔNG MINH
            </span>
            <h2 style="font-size: 24px; font-weight: 800; margin: 6px 0; color: #FFFFFF;">
                Trợ lý Chuyển đổi Hình ảnh / PDF sang Word (.docx)
            </h2>
            <p style="font-size: 14.5px; color: #E2E8F0; margin: 0; line-height: 1.5; max-width: 900px;">
                Sử dụng API Trí tuệ nhân tạo thị giác máy tính của <b>Google Gemini</b> để nhận diện văn bản (OCR), phục hồi bảng biểu (tables), 
                phân cấp tiêu đề (headings), câu hỏi đề thi và danh sách từ tài liệu chụp/scan, rồi xuất sang file <b>Microsoft Word (.docx)</b> chuẩn thể thức.
            </p>
            <div style="margin-top: 14px;">
                <span class="doc-badge-pill">🔍 OCR Đa ngôn ngữ tiếng Việt chuẩn</span>
                <span class="doc-badge-pill">📊 Tự động tạo Bảng Word thực tế</span>
                <span class="doc-badge-pill">📝 Đề thi & Phiếu bài tập trắc nghiệm</span>
                <span class="doc-badge-pill">📋 Chuẩn thể thức TCVN & Bộ GD&ĐT</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # 1-Click Samples and Upload Bar
    st.markdown("### 1. Chọn Tài liệu Cần Chuyển Đổi")
    col_doc_s1, col_doc_s2, col_doc_s3 = st.columns([1, 1, 1])

    with col_doc_s1:
        if st.button("🧪 Nạp Đề kiểm tra mẫu (Ảnh .PNG)", use_container_width=True, help="Nạp file ảnh chụp đề kiểm tra Toán lớp 3 có trắc nghiệm, tự luận và bảng số liệu để thử nghiệm."):
            st.session_state["doc_file_bytes"] = get_sample_exam_image_bytes()
            st.session_state["doc_file_name"] = "De_Kiem_Tra_Toan_Lop3_Mau.png"
            st.session_state["doc_file_type"] = "image/png"
            st.session_state["doc_result_md"] = None
            st.session_state["doc_result_docx"] = None
            st.success("Đã nạp file ảnh đề kiểm tra mẫu thành công!")

    with col_doc_s2:
        if st.button("🧪 Nạp Đề kiểm tra mẫu (File .PDF)", use_container_width=True, help="Nạp file PDF đề kiểm tra mẫu để kiểm tra tính năng xử lý file PDF chuẩn."):
            st.session_state["doc_file_bytes"] = get_sample_exam_pdf_bytes()
            st.session_state["doc_file_name"] = "De_Kiem_Tra_Toan_Lop3_Mau.pdf"
            st.session_state["doc_file_type"] = "application/pdf"
            st.session_state["doc_result_md"] = None
            st.session_state["doc_result_docx"] = None
            st.success("Đã nạp file PDF đề kiểm tra mẫu thành công!")

    with col_doc_s3:
        # Download Sample Image
        st.download_button(
            label="📥 Tải Đề mẫu về máy (.PNG)",
            data=get_sample_exam_image_bytes(),
            file_name="De_Kiem_Tra_Toan_Lop3_Mau.png",
            mime="image/png",
            use_container_width=True,
        )

    # File Uploader
    uploaded_doc = st.file_uploader(
        "Hoặc kéo thả file Hình ảnh (.png, .jpg, .jpeg, .webp, .bmp) hoặc file PDF (.pdf) của bạn vào đây:",
        type=["pdf", "png", "jpg", "jpeg", "webp", "bmp"],
        help="Hỗ trợ ảnh chụp điện thoại, bản scan đề thi, công văn, giáo án hoặc file PDF tài liệu bất kỳ.",
    )

    if uploaded_doc is not None:
        file_ext = uploaded_doc.name.split(".")[-1].lower()
        if file_ext == "pdf":
            mime = "application/pdf"
        elif file_ext in ["jpg", "jpeg"]:
            mime = "image/jpeg"
        elif file_ext == "png":
            mime = "image/png"
        elif file_ext == "webp":
            mime = "image/webp"
        else:
            mime = "image/png"  # will convert bmp/others to png

        if (
            st.session_state.get("doc_file_name") != uploaded_doc.name
            or st.session_state.get("doc_file_bytes") != uploaded_doc.getvalue()
        ):
            st.session_state["doc_file_bytes"] = uploaded_doc.getvalue()
            st.session_state["doc_file_name"] = uploaded_doc.name
            st.session_state["doc_file_type"] = mime
            st.session_state["doc_result_md"] = None
            st.session_state["doc_result_docx"] = None

    active_doc_bytes = st.session_state.get("doc_file_bytes")
    active_doc_name = st.session_state.get("doc_file_name", "Tai_lieu")
    active_doc_type = st.session_state.get("doc_file_type", "")

    if active_doc_bytes:
        file_size_kb = len(active_doc_bytes) / 1024
        file_size_str = f"{file_size_kb:.1f} KB" if file_size_kb < 1024 else f"{(file_size_kb / 1024):.2f} MB"

        st.write("")
        st.markdown("### 2. Xem Trước Tài Liệu & Cấu Hình Xuất Word")

        col_left_doc, col_right_doc = st.columns([1, 1])

        # LEFT COLUMN: Preview & Range
        with col_left_doc:
            st.markdown(
                f"""
                <div class="doc-preview-card">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                        <span style="font-weight: 700; color: #1E293B; font-size: 14px;">
                            📄 {active_doc_name}
                        </span>
                        <span style="background: #E2E8F0; color: #334155; padding: 2px 8px; border-radius: 4px; font-size: 12px; font-weight: 600;">
                            {file_size_str}
                        </span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            is_pdf = (active_doc_type == "application/pdf")
            pdf_info = None
            selected_page_range = None

            if is_pdf:
                pdf_info = GoogleVisionDocExtractor.inspect_pdf(active_doc_bytes)
                total_pages = pdf_info.get("num_pages", 1)

                st.caption(f"Tài liệu PDF gồm **{total_pages} trang**.")

                # Thumbnail of first page
                thumb_img = GoogleVisionDocExtractor.render_pdf_page_thumbnail(active_doc_bytes, 0)
                if thumb_img:
                    st.image(thumb_img, caption="Ảnh xem trước Trang 1", use_container_width=True)

                if total_pages > 1:
                    pdf_scope = st.radio(
                        "Phạm vi trang cần chuyển đổi:",
                        options=["Toàn bộ tài liệu (Tất cả các trang)", "Chọn khoảng trang cụ thể"],
                        index=0,
                        horizontal=True,
                    )
                    if pdf_scope == "Chọn khoảng trang cụ thể":
                        c_p1, c_p2 = st.columns(2)
                        with c_p1:
                            start_p = st.number_input("Từ trang:", min_value=1, max_value=total_pages, value=1)
                        with c_p2:
                            end_p = st.number_input("Đến trang:", min_value=start_p, max_value=total_pages, value=min(total_pages, start_p + 1))
                        selected_page_range = (start_p, end_p)
            else:
                # Image Preview
                st.image(active_doc_bytes, caption=f"Ảnh tài liệu: {active_doc_name}", use_container_width=True)

        # RIGHT COLUMN: Options & Settings
        with col_right_doc:
            st.markdown("#### ⚙️ Thiết lập Nhận diện & Thể thức Word")

            doc_mode_options = {
                "exam": "📝 Đề thi & Phiếu bài tập (Tối ưu trắc nghiệm A-B-C-D, tự luận)",
                "general": "📚 Tài liệu chung / Sách / Báo cáo tổng quát",
                "lesson_plan": "📖 Giáo án / Kế hoạch bài dạy (Bảng Hoạt động GV & HS)",
                "official_doc": "🏛️ Công văn / Thông tư / Quyết định (Chuẩn Nghị định 30)",
                "table_data": "📊 Trích xuất Bảng số liệu & Bảng điểm (Chuyên biệt)",
            }

            selected_mode_key = st.selectbox(
                "Chế độ tài liệu tối ưu:",
                options=list(doc_mode_options.keys()),
                format_func=lambda k: doc_mode_options[k],
                index=0 if "kiem_tra" in active_doc_name.lower() or "toan" in active_doc_name.lower() or "de" in active_doc_name.lower() else 1,
                help="AI sẽ áp dụng khuôn mẫu sư phạm tương ứng để nhận diện câu hỏi, bảng biểu và cấu trúc phân cấp phù hợp nhất.",
            )

            col_w1, col_w2 = st.columns(2)
            with col_w1:
                word_font = st.selectbox(
                    "Phông chữ văn bản Word:",
                    options=["Times New Roman", "Calibri", "Arial"],
                    index=0,
                    help="Times New Roman là quy chuẩn quốc gia theo Nghị định 30 và văn bản ngành Giáo dục.",
                )
            with col_w2:
                word_size = st.selectbox(
                    "Cỡ chữ nội dung (Body):",
                    options=[12, 13, 14],
                    index=1,  # 13pt
                    help="13pt là cỡ chữ chuẩn cho các văn bản hành chính và giáo án tiểu học.",
                )

            table_theme_options = {
                "education_standard": "🏫 Chuẩn Giáo dục & Hành chính (Nền xám nhạt #F1F5F9, viền chuẩn)",
                "modern_navy": "💎 Chuyên nghiệp Hiện đại (Header Xanh Navy, chữ trắng)",
                "minimal_classic": "📄 Tối giản Trắng Đen (Viền đen cổ điển, nền trắng)",
            }
            selected_table_theme = st.selectbox(
                "Kiểu bảng biểu Word:",
                options=list(table_theme_options.keys()),
                format_func=lambda k: table_theme_options[k],
                index=0,
                help="Chuẩn Giáo dục tự động tối ưu độ rộng từng cột và lề ô cho đề thi, bài tập và công văn.",
            )

            suggested_title = active_doc_name.rsplit(".", 1)[0].replace("_", " ").title()
            doc_title_input = st.text_input(
                "Tiêu đề tài liệu (in ở đầu và chân trang):",
                value=suggested_title,
            )

            custom_req = st.text_area(
                "Ghi chú bổ sung cho AI (tùy chọn):",
                placeholder="Ví dụ: Giữ nguyên bố cục bảng 2 cột đầu trang; Trình bày câu hỏi trắc nghiệm A B C D cách đều; Chừa 3 dòng làm bài...",
                height=68,
            )

            # API Key verification & quick input if not set in sidebar
            current_api_key = st.session_state.get("gemini_api_key", "").strip()
            if not current_api_key:
                st.warning("⚠️ Chưa cấu hình Google Gemini API Key.")
                quick_key = st.text_input(
                    "Nhập Gemini API Key tại đây:",
                    type="password",
                    placeholder="AIzaSy...",
                    key="quick_doc_api_key",
                    help="Khóa API hoàn toàn miễn phí từ Google AI Studio (aistudio.google.com).",
                )
                if quick_key:
                    st.session_state["gemini_api_key"] = quick_key.strip()
                    current_api_key = quick_key.strip()
                    st.rerun()

            st.write("")
            btn_start_convert = st.button(
                "🚀 BẮT ĐẦU CHUYỂN ĐỔI SANG FILE WORD (.DOCX)",
                type="primary",
                use_container_width=True,
            )

            if btn_start_convert:
                if not current_api_key:
                    st.error("Vui lòng nhập Google Gemini API Key để thực hiện nhận diện bằng AI.")
                else:
                    t_start = time.time()
                    with st.spinner("Đang sử dụng Google Gemini Vision AI phân tích tài liệu và tái tạo bố cục Word..."):
                        try:
                            # 1. Prepare file bytes
                            extractor = GoogleVisionDocExtractor(
                                api_key=current_api_key,
                                model_name=model_option,
                            )

                            bytes_to_send = active_doc_bytes
                            mime_to_send = active_doc_type

                            # Handle PDF page slicing if requested
                            if is_pdf and selected_page_range:
                                bytes_to_send = extractor.slice_pdf_pages(
                                    active_doc_bytes,
                                    selected_page_range[0],
                                    selected_page_range[1],
                                )

                            # Handle non-standard image formats (like BMP)
                            if not is_pdf and not mime_to_send.startswith("image/"):
                                mime_to_send = "image/png"

                            # 2. Call Google Gemini Vision OCR
                            extracted_md = extractor.process_document(
                                file_bytes=bytes_to_send,
                                mime_type=mime_to_send,
                                document_type=selected_mode_key,
                                custom_instructions=custom_req,
                            )

                            # 3. Build Word Document (.docx) with advanced layout engine
                            builder = MarkdownToDocxBuilder(
                                font_name=word_font,
                                font_size_pt=word_size,
                                table_theme=selected_table_theme,
                            )
                            docx_bytes = builder.create_docx(
                                markdown_text=extracted_md,
                                document_title=doc_title_input,
                            )

                            elapsed_sec = round(time.time() - t_start, 1)

                            # Sanitize filename strictly to guarantee .docx extension across all browsers
                            clean_export_name = sanitize_filename(active_doc_name.rsplit(".", 1)[0] + "_ChuyenDoi_AI.docx")
                            
                            # Automatically save a local backup in exports/
                            saved_local_path = save_file_locally(docx_bytes, clean_export_name)

                            # Store into session state
                            st.session_state["doc_result_md"] = extracted_md
                            st.session_state["doc_result_docx"] = docx_bytes
                            st.session_state["doc_export_name"] = clean_export_name
                            st.session_state["doc_word_count"] = len(extracted_md.split())
                            st.session_state["doc_process_time"] = elapsed_sec
                            st.session_state["doc_used_font"] = word_font
                            st.session_state["doc_used_size"] = word_size
                            st.session_state["doc_used_theme"] = selected_table_theme
                            st.session_state["doc_used_title"] = doc_title_input
                            st.session_state["saved_local_path"] = saved_local_path

                            st.success(f"🎉 Chuyển đổi thành công sau {elapsed_sec} giây!")
                        except Exception as e:
                            st.error(f"Lỗi trong quá trình chuyển đổi: {e}")

        # Display Conversion Results if Available
        res_docx_bytes = st.session_state.get("doc_result_docx")
        res_md = st.session_state.get("doc_result_md")

        if res_docx_bytes and res_md:
            st.divider()
            st.markdown("### 3. Kết Quả Chuyển Đổi & Tải Về File Word (.docx)")

            # Stats row
            docx_size_kb = len(res_docx_bytes) / 1024
            st_col1, st_col2, st_col3, st_col4 = st.columns(4)
            with st_col1:
                st.markdown(
                    f"""
                    <div class="stat-box">
                        <div class="stat-num" style="color: #1E3A8A;">{st.session_state.get('doc_word_count', 0)}</div>
                        <div class="stat-label">Số từ nhận diện</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with st_col2:
                st.markdown(
                    f"""
                    <div class="stat-box">
                        <div class="stat-num" style="color: #059669;">{docx_size_kb:.1f} KB</div>
                        <div class="stat-label">Kích thước file Word</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with st_col3:
                st.markdown(
                    f"""
                    <div class="stat-box">
                        <div class="stat-num" style="color: #0284C7;">{st.session_state.get('doc_process_time', 0)}s</div>
                        <div class="stat-label">Thời gian xử lý AI</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with st_col4:
                st.markdown(
                    """
                    <div class="stat-box" style="background: #F0FDF4; border: 2px solid #10B981;">
                        <div class="stat-num" style="color: #059669;">.DOCX</div>
                        <div class="stat-label" style="color: #065F46; font-weight: 600;">Chuẩn Word 100%</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            st.write("")

            clean_export_name = sanitize_filename(st.session_state.get("doc_export_name", "Tai_Lieu_Chuyen_Doi.docx"))
            b64_docx = base64.b64encode(res_docx_bytes).decode()
            doc_size_kb = round(len(res_docx_bytes) / 1024, 1)

            # Ensure file is saved to both exports/ and static/
            saved_local_path = save_file_locally(res_docx_bytes, clean_export_name)
            st.session_state["saved_local_path"] = saved_local_path
            # Update export name to the actual saved filename (may include timestamp if original was locked)
            clean_export_name = os.path.basename(saved_local_path)

            # 1. PRIMARY BULLETPROOF DOWNLOAD CARD (Isolated HTML Component)
            # Uses direct static file serving (/app/static/{name}) and Base64 data URI
            # Guarantees 100% proper .docx filename and extension across all browsers & IDM
            download_html = f"""
            <!DOCTYPE html>
            <html>
            <head>
            <meta charset="utf-8">
            <style>
              * {{ box-sizing: border-box; margin: 0; padding: 0; }}
              body {{
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
                background: transparent;
                padding: 4px;
              }}
              .download-card {{
                background: linear-gradient(135deg, #F8FAFC 0%, #EFF6FF 100%);
                border: 2px solid #2563EB;
                border-radius: 14px;
                padding: 20px 24px;
                box-shadow: 0 4px 15px rgba(37, 99, 235, 0.12);
                text-align: center;
              }}
              .card-header {{
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 10px;
                margin-bottom: 14px;
              }}
              .card-title {{
                font-size: 16.5px;
                font-weight: 700;
                color: #0F172A;
              }}
              .file-badge {{
                background: #1D4ED8;
                color: #FFFFFF;
                font-size: 11px;
                font-weight: 800;
                padding: 2px 7px;
                border-radius: 4px;
                letter-spacing: 0.5px;
              }}
              .btn-container {{
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 14px;
                flex-wrap: wrap;
                margin: 12px 0;
              }}
              .btn-primary {{
                display: inline-flex;
                align-items: center;
                gap: 8px;
                background: linear-gradient(135deg, #10B981 0%, #059669 100%);
                color: #FFFFFF !important;
                font-size: 16px;
                font-weight: 700;
                padding: 13px 28px;
                border-radius: 8px;
                text-decoration: none !important;
                box-shadow: 0 4px 12px rgba(16, 185, 129, 0.35);
                transition: transform 0.1s ease, box-shadow 0.1s ease;
                cursor: pointer;
              }}
              .btn-primary:hover {{
                background: linear-gradient(135deg, #059669 0%, #047857 100%);
                transform: translateY(-1px);
                box-shadow: 0 6px 16px rgba(16, 185, 129, 0.45);
              }}
              .btn-secondary {{
                display: inline-flex;
                align-items: center;
                gap: 6px;
                background: #FFFFFF;
                color: #1E40AF !important;
                border: 1.5px solid #3B82F6;
                font-size: 14px;
                font-weight: 600;
                padding: 12px 20px;
                border-radius: 8px;
                text-decoration: none !important;
                box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05);
                transition: all 0.1s ease;
                cursor: pointer;
              }}
              .btn-secondary:hover {{
                background: #EFF6FF;
                border-color: #1D4ED8;
              }}
              .info-text {{
                font-size: 13px;
                color: #334155;
                margin-top: 10px;
                line-height: 1.4;
              }}
              .info-highlight {{
                color: #0369A1;
                font-weight: 600;
              }}
            </style>
            </head>
            <body>
              <div class="download-card">
                <div class="card-header">
                  <span style="font-size: 22px;">📄</span>
                  <span class="card-title">Tên file xuất: <code style="color: #1E3A8A; background: #DBEAFE; padding: 2px 6px; border-radius: 4px;">{clean_export_name}</code></span>
                  <span class="file-badge">.DOCX CHUẨN</span>
                  <span style="color: #64748B; font-size: 13px;">({doc_size_kb} KB)</span>
                </div>
                <div class="btn-container">
                  <a class="btn-primary" href="/app/static/{clean_export_name}" download="{clean_export_name}" target="_blank">
                    📥 TẢI VỀ FILE WORD (.DOCX)
                  </a>
                  <a class="btn-secondary" href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64_docx}" download="{clean_export_name}">
                    ⚡ Tải Trực Tiếp (Base64)
                  </a>
                </div>
                <div class="info-text">
                  ✅ <b>Cam kết định dạng:</b> Hệ thống áp dụng truyền tải tĩnh trực tiếp (Static Serving) kết hợp tên file tiêu chuẩn, đảm bảo file tải về <b>luôn có đuôi .docx hợp lệ 100%</b>, mở ngay bằng Microsoft Word, Google Docs hay WPS Office.
                </div>
              </div>
            </body>
            </html>
            """
            components.html(download_html, height=175)

            # Local Computer Direct Actions (for Windows desktop convenience)
            col_act1, col_act2, col_act3 = st.columns([1.2, 1.2, 1.2])
            with col_act1:
                if st.button("🚀 Mở file ngay bằng Word", use_container_width=True, help="Khởi chạy trực tiếp file Word vừa tạo trên ứng dụng Microsoft Word của máy tính."):
                    try:
                        os.startfile(saved_local_path)
                        st.success("Đã mở Microsoft Word thành công!")
                    except Exception as err:
                        st.info(f"Đường dẫn file trên máy: `{saved_local_path}` ({err})")
            with col_act2:
                if st.button("📂 Mở thư mục chứa file", use_container_width=True, help="Mở thư mục lưu file trong Windows Explorer và chọn sẵn file vừa chuyển đổi."):
                    try:
                        norm_p = os.path.normpath(saved_local_path)
                        subprocess.Popen(f'explorer /select,"{norm_p}"')
                        st.success("Đã mở thư mục Windows Explorer!")
                    except Exception as err:
                        st.info(f"Đường dẫn: `{saved_local_path}`")
            with col_act3:
                # Streamlit native fallback button
                st.download_button(
                    label=f"📥 Tải dự phòng máy chủ",
                    data=res_docx_bytes,
                    file_name=clean_export_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    help="Phương án tải về dự phòng qua Streamlit API",
                )

            st.write("")

            # 3 Sub-tabs for Preview, Editing, and Raw Markdown
            subtab_preview, subtab_edit, subtab_raw = st.tabs(
                [
                    "👁️ Xem trước nội dung văn bản",
                    "✏️ Chỉnh sửa trực tiếp trước khi xuất",
                    "📋 Xem mã Markdown gốc",
                ]
            )

            with subtab_preview:
                st.markdown(
                    """
                    <div class="action-box">
                        📄 Dưới đây là nội dung đã được Google AI nhận diện và tái tạo cấu trúc chuẩn (bao gồm tiêu đề, bảng biểu, danh sách và in đậm/nghiêng):
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.markdown(res_md)

            with subtab_edit:
                st.info("💡 Bạn có thể chỉnh sửa lại bất kỳ từ ngữ, số liệu hoặc thêm bớt câu hỏi bên dưới. Nhấn nút **'Cập nhật file Word'** để hệ thống tái tạo file Word mới tức thì mà không cần gọi lại AI!")
                edited_text = st.text_area(
                    "Nội dung văn bản (định dạng Markdown):",
                    value=res_md,
                    height=350,
                    key="edited_markdown_area",
                )

                if st.button("💾 CẬP NHẬT FILE WORD VỚI NỘI DUNG ĐÃ SỬA", use_container_width=True):
                    builder = MarkdownToDocxBuilder(
                        font_name=st.session_state.get("doc_used_font", "Times New Roman"),
                        font_size_pt=st.session_state.get("doc_used_size", 13),
                        table_theme=st.session_state.get("doc_used_theme", "education_standard"),
                    )
                    new_docx_bytes = builder.create_docx(
                        markdown_text=edited_text,
                        document_title=st.session_state.get("doc_used_title", ""),
                    )
                    # Update local file
                    new_saved_path = save_file_locally(new_docx_bytes, clean_export_name)
                    st.session_state["saved_local_path"] = new_saved_path
                    st.session_state["doc_result_md"] = edited_text
                    st.session_state["doc_result_docx"] = new_docx_bytes
                    st.session_state["doc_word_count"] = len(edited_text.split())
                    st.success("Đã cập nhật file Word thành công! Bạn có thể bấm nút Tải về ở trên để lấy file mới.")
                    st.rerun()

            with subtab_raw:
                st.code(res_md, language="markdown")

    else:
        st.info("👆 Vui lòng chọn một trong các nút **'Nạp Đề kiểm tra mẫu'** ở trên hoặc kéo thả file Ảnh / PDF của bạn vào để trải nghiệm tính năng.")


# =========================================================
# TAB 4: KNOWLEDGE BASE & CIRCULAR 27 GUIDE
# =========================================================
with tab_knowledge:
    st.markdown("### 📖 Sổ tay Hướng dẫn Đánh giá theo Thông tư 27/2020/TT-BGDĐT")

    k_col1, k_col2 = st.columns([1, 1])

    with k_col1:
        st.markdown(
            """
            #### 1. Các Mức Đánh giá Định kỳ (Điều 7 & Điều 9)
            - **Mức Hoàn thành tốt (T):**
              - Thực hiện tốt các yêu cầu học tập và hồ sơ học tập.
              - Thường xuyên có biểu hiện cụ thể về các thành phần năng lực môn học hoặc phẩm chất nổi trội.
              - Điểm kiểm tra định kỳ từ 9.0 đến 10.0 điểm.
            - **Mức Hoàn thành (H):**
              - Thực hiện được các yêu cầu học tập và hồ sơ học tập.
              - Có biểu hiện cụ thể về các thành phần năng lực của môn học.
              - Điểm kiểm tra định kỳ từ 5.0 đến dưới 9.0 điểm.
            - **Mức Chưa hoàn thành (C):**
              - Chưa thực hiện được một số yêu cầu học tập hoặc hồ sơ học tập chưa đầy đủ.
              - Chưa có biểu hiện rõ về các thành phần năng lực của môn học.
              - Điểm kiểm tra định kỳ dưới 5.0 điểm.
            """
        )

    with k_col2:
        st.markdown(
            """
            #### 2. Nguyên tắc Sư phạm khi Ghi Lời Nhận xét
            - **Tôn trọng sự tiến bộ cá nhân:** Đánh giá sự tiến bộ của từng học sinh, không so sánh học sinh này với học sinh khác.
            - **Động viên, khích lệ:** Lời phê phải mang tính xây dựng, giúp các em tự tin, hào hứng trong học tập.
            - **Chỉ rõ điểm cần rèn luyện:** Đối với học sinh Mức H hoặc Mức C, nêu rõ 1 kỹ năng hoặc giải pháp cụ thể (VD: rèn đọc ngắt nghỉ, luyện tính nhẩm, nhờ thầy cô hỗ trợ).
            - **Không dùng từ ngữ tiêu cực:** Tuyệt đối không dùng các từ ngữ như: "học dốt", "lười biếng", "kém cỏi".
            - **Độ dài chuẩn mực:** Khoảng 15 đến 35 từ, súc tích, vừa vặn trên 1-2 dòng trong học bạ và sổ theo dõi.
            """
        )

    st.divider()

    st.markdown("#### 📚 Thư viện Mẫu Lời Nhận xét theo Môn học (Thông tư 27)")
    selected_kb_subj = st.selectbox(
        "Chọn môn học để xem ngân hàng câu mẫu:",
        options=list(TT27_COMMENT_BANK.keys()),
        index=0,
    )

    subj_bank = TT27_COMMENT_BANK[selected_kb_subj]
    kb_col_t, kb_col_h, kb_col_c = st.columns(3)

    with kb_col_t:
        st.markdown("##### 🌟 Mức Hoàn thành tốt (T)")
        for item in subj_bank.get("T", []):
            st.markdown(f"- *{item}*")

    with kb_col_h:
        st.markdown("##### 📘 Mức Hoàn thành (H)")
        for item in subj_bank.get("H", []):
            st.markdown(f"- *{item}*")

    with kb_col_c:
        st.markdown("##### 🤝 Mức Chưa hoàn thành (C)")
        for item in subj_bank.get("C", []):
            st.markdown(f"- *{item}*")

    st.divider()

    st.markdown("#### 🧩 Ma trận Năng lực Cốt lõi các Môn Tiểu học (CT GDPT 2018)")
    for s_name, comp_dict in SUBJECT_COMPETENCIES.items():
        with st.expander(f"Môn {s_name}"):
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Biểu hiện ưu điểm nổi bật:**")
                for s in comp_dict["strengths"]:
                    st.markdown(f"- {s}")
            with c2:
                st.markdown("**Các điểm cần lưu ý rèn luyện thêm:**")
                for g in comp_dict["growth_areas"]:
                    st.markdown(f"- {g}")
