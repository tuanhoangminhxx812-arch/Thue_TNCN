"""
File: app.py
Mục đích: Ứng dụng Web Streamlit kiểm dò dữ liệu Thuế TNCN đa người dùng,
          xử lý hoàn toàn trong bộ nhớ RAM (io.BytesIO) đảm bảo an toàn & bảo mật.
          Hỗ trợ:
          1. Nạp từ 1 thư mục chứa đủ 4 file dữ liệu (không cần chọn từng file riêng lẻ).
          2. Định dạng tất cả số liệu là số nguyên không có số lẻ, phân cách hàng nghìn bằng dấu chấm.
          3. Căn phải cột số, căn trái cột chữ, căn giữa mã/trạng thái.
          4. Báo cáo Excel 4 sheet tải về ngay lập tức.
"""

import io
import os
import sys
import base64
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st

# Đảm bảo console UTF-8 và nhận diện module scripts
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parent))

from scripts.data_loader import DataLoader, identify_input_files
from scripts.reconciler import TaxReconciler
from scripts.report_builder import ExcelReportBuilder

# ==============================================================================
# CẤU HÌNH TRANG WEB STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="Kiểm Dò Thuế TNCN - Công Ty Điện Lực",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS cho giao diện sang trọng, chuyên nghiệp
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F4E78 0%, #2F5597 100%);
        padding: 24px;
        border-radius: 12px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    }
    .main-header h1 {
        color: #ffffff;
        margin: 0;
        font-size: 26px;
        font-weight: 700;
    }
    .main-header p {
        color: #E2EFDA;
        margin: 6px 0 0 0;
        font-size: 14px;
    }
    .kpi-card {
        background: #ffffff;
        border: 1px solid #E0E0E0;
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.04);
        margin-bottom: 12px;
        border-left: 5px solid #1F4E78;
    }
    .kpi-title {
        color: #6c757d;
        font-size: 12px;
        text-transform: uppercase;
        font-weight: 600;
        margin-bottom: 4px;
    }
    .kpi-value {
        color: #1F4E78;
        font-size: 22px;
        font-weight: 700;
        margin-bottom: 4px;
    }
    .kpi-sub {
        font-size: 12px;
        color: #495057;
    }
    .badge-matched {
        display: inline-block;
        background-color: #E2EFDA;
        color: #276A3C;
        padding: 4px 12px;
        border-radius: 16px;
        font-weight: 700;
        font-size: 12px;
        border: 1px solid #A9D18E;
    }
    .badge-diff {
        display: inline-block;
        background-color: #FCE4D6;
        color: #C00000;
        padding: 4px 12px;
        border-radius: 16px;
        font-weight: 700;
        font-size: 12px;
        border: 1px solid #F4B084;
    }
    .file-badge-ok {
        background-color: #e8f5e9;
        border-left: 4px solid #2e7d32;
        padding: 8px 12px;
        margin-bottom: 6px;
        border-radius: 4px;
        font-size: 13px;
    }
    .file-badge-missing {
        background-color: #ffebee;
        border-left: 4px solid #c62828;
        padding: 8px 12px;
        margin-bottom: 6px;
        border-radius: 4px;
        font-size: 13px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 48px;
        border-radius: 6px 6px 0 0;
        padding: 8px 16px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)


def format_vnd(val):
    """
    Định dạng số nguyên không có số lẻ, phân cách hàng nghìn bằng dấu chấm (chuẩn Việt Nam).
    Ví dụ: 15394909454 -> 15.394.909.454; -25000 -> -25.000; 0 -> 0; None -> '-'
    """
    if pd.isna(val) or val is None or val == "" or str(val).strip() == "-":
        return "-"
    try:
        f_val = float(str(val).replace(".", "").replace(",", ".")) if isinstance(val, str) else float(val)
        int_val = int(round(f_val))
        return f"{int_val:,}".replace(",", ".")
    except (ValueError, TypeError):
        return str(val)


def display_styled_dataframe(
    df: pd.DataFrame,
    money_cols: list = None,
    center_cols: list = None,
    diff_col_name: str = "chenh_lech",
    status_col_name: str = "kiem_do"
):
    """
    Chuẩn hóa và hiển thị bảng với Streamlit:
    - Định dạng số nguyên không số lẻ, phân cách hàng nghìn bằng dấu chấm.
    - CĂN PHẢI 100% TOÀN BỘ CỘT TIỀN / SỐ (thông qua column_config alignment='right').
    - CĂN GIỮA MÃ, STT, TRẠNG THÁI (thông qua column_config alignment='center').
    - CĂN TRÁI VĂN BẢN, HỌ TÊN, PHÒNG BAN.
    - Tô màu nền xanh khi khớp (chênh lệch = 0), đỏ khi lệch.
    """
    df_disp = df.copy()

    for col in df_disp.columns:
        if money_cols and col in money_cols:
            df_disp[col] = df_disp[col].apply(format_vnd).astype(str)
        else:
            df_disp[col] = df_disp[col].apply(lambda x: "" if pd.isna(x) else str(x))

    # Cấu hình column_config để Glide Data Grid canvas căn lề chính xác
    col_config = {}
    if money_cols:
        for c in money_cols:
            if c in df_disp.columns:
                col_config[c] = {"alignment": "right"}
    if center_cols:
        for c in center_cols:
            if c in df_disp.columns:
                col_config[c] = {"alignment": "center"}

    def apply_style(row):
        styles = [''] * len(row)
        is_matched = False

        if status_col_name in row.index:
            st_val = str(row[status_col_name]).strip()
            if st_val == "KHỚP":
                is_matched = True
            elif st_val in ["LỆCH", "CẢNH BÁO"]:
                is_matched = False

        if not is_matched and diff_col_name in row.index:
            raw_diff = str(row[diff_col_name]).replace(".", "").replace(",", ".").strip()
            try:
                num = float(raw_diff)
                if abs(num) <= 1000:
                    is_matched = True
            except Exception:
                pass

        for i, col in enumerate(row.index):
            if col in [diff_col_name, status_col_name, "lech_luong_vs_tokhai"]:
                if is_matched:
                    styles[i] = 'background-color: #E2EFDA; color: #276A3C; font-weight: bold;'
                else:
                    styles[i] = 'background-color: #FCE4D6; color: #C00000; font-weight: bold;'

        return styles

    styled_df = df_disp.style.apply(apply_style, axis=1)

    st.dataframe(
        styled_df,
        column_config=col_config,
        use_container_width=True,
        hide_index=True
    )


# ==============================================================================
# SIDEBAR: THIẾT LẬP & NẠP 1 THƯ MỤC CHỨA CẢ 4 FILE
# ==============================================================================
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/accounting.png", width=64)
    st.title("Thiết Lập Dữ Liệu")
    st.caption("⚡ Dành cho Công ty Điện lực")
    st.markdown("---")

    report_period = st.text_input(
        "📅 Kỳ báo cáo:",
        value="Tháng 09/2026",
        help="Ví dụ: Tháng 09/2026 hoặc Quý 3/2026"
    )

    company_name = st.text_input(
        "🏢 Tên đơn vị / Công ty:",
        value="TỔNG CÔNG TY ĐIỆN LỰC TP. HỒ CHÍ MINH - CÔNG TY ĐIỆN LỰC CHỢ LỚN"
    )

    st.markdown("---")
    st.markdown("### 📁 Nạp Dữ Liệu Từ 1 Thư Mục")
    st.caption("Chỉ cần chỉ định 1 thư mục chứa đủ 4 file dữ liệu. Chương trình sẽ tự động nhận diện từng file.")

    # Khởi tạo session_state cho folder
    if "folder_path" not in st.session_state:
        st.session_state["folder_path"] = "02_THAM_KHAO" if Path("02_THAM_KHAO").exists() else "01_DAU_VAO"

    folder_path_input = st.text_input(
        "Đường dẫn thư mục chứa 4 file:",
        value=st.session_state["folder_path"],
        help="Nhập đường dẫn thư mục trên máy tính (ví dụ: 02_THAM_KHAO hoặc D:/DATA/Thue_TNCN/02_THAM_KHAO)"
    )

    c_b1, c_b2 = st.columns(2)
    with c_b1:
        if st.button("📁 02_THAM_KHAO", use_container_width=True, help="Dữ liệu thực tế EVN"):
            st.session_state["folder_path"] = "02_THAM_KHAO"
            st.rerun()
    with c_b2:
        if st.button("📁 01_DAU_VAO", use_container_width=True, help="Dữ liệu mẫu kiểm thử"):
            st.session_state["folder_path"] = "01_DAU_VAO"
            st.rerun()

    st.markdown("---")
    st.caption("📂 **Hoặc kéo thả toàn bộ các file từ thư mục vào đây:**")
    uploaded_files = st.file_uploader(
        "Kéo thả tất cả các file cùng lúc (.xlsx):",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="Chọn toàn bộ 4 file trong thư mục của bạn rồi kéo thả vào đây."
    )


# ==============================================================================
# HEADER CHÍNH
# ==============================================================================
st.markdown(f"""
<div class="main-header">
    <h1>HỆ THỐNG KIỂM DÒ DỮ LIỆU THUẾ TNCN</h1>
    <p>🏢 {company_name} | 📅 Kỳ kiểm dò: <b>{report_period}</b> | 🔒 Xử lý In-Memory An toàn & Bảo mật</p>
</div>
""", unsafe_allow_html=True)


# ==============================================================================
# TỰ ĐỘNG QUÉT VÀ NHẬN DIỆN 4 FILE TRONG THƯ MỤC
# ==============================================================================
files_ready = False
salary_src = None
pr025_src = None
pr027_src = None
gl_src = None

# Ưu tiên các file được upload nếu người dùng kéo thả
if uploaded_files and len(uploaded_files) > 0:
    source_label = f"Tải lên {len(uploaded_files)} file từ trình duyệt"
    identified = identify_input_files(uploaded_files)
else:
    chosen_folder = folder_path_input.strip()
    source_label = f"Thư mục `{chosen_folder}`"
    folder_p = Path(chosen_folder)
    if folder_p.exists() and folder_p.is_dir():
        identified = identify_input_files(folder_p)
    else:
        identified = {"salary": None, "pr025": None, "pr027": None, "gl_3335": None}

salary_src = identified["salary"]
pr025_src = identified["pr025"]
pr027_src = identified["pr027"]
gl_src = identified["gl_3335"]

all_found = all([salary_src, pr025_src, pr027_src, gl_src])
files_ready = all_found


# ==============================================================================
# HIỂN THỊ TRẠNG THÁI NHẬN DIỆN 4 FILE
# ==============================================================================
with st.sidebar:
    st.markdown("#### 📋 Trạng Thái 4 File Cần Lấy:")

    def render_file_status(title, file_obj, hint):
        if file_obj:
            fname = getattr(file_obj, "name", str(Path(file_obj).name if isinstance(file_obj, (str, Path)) else file_obj))
            st.markdown(f'<div class="file-badge-ok">✅ <b>{title}:</b><br><code>{fname}</code></div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="file-badge-missing">❌ <b>{title}:</b><br><small>Chưa tìm thấy (tên file chứa <i>{hint}</i>)</small></div>', unsafe_allow_html=True)

    render_file_status("1. Bảng lương CBCNV", salary_src, "luong, ttt, salary...")
    render_file_status("2. Tờ khai thuế PR025", pr025_src, "025, pr025, 05-kk...")
    render_file_status("3. Bảng kê PR027/028", pr027_src, "027, pr027, nghi_huu...")
    render_file_status("4. Sổ cái TK 3335", gl_src, "3335, so_cai...")

    st.markdown("---")
    st.markdown("#### 📤 Thư Mục Xuất Báo Cáo:")
    if st.button("📂 Mở Thư Mục 03_DAU_RA", key="btn_open_out_sidebar", use_container_width=True):
        if hasattr(os, "startfile"):
            try:
                out_p = Path("03_DAU_RA").resolve()
                out_p.mkdir(parents=True, exist_ok=True)
                os.startfile(str(out_p))
            except Exception:
                pass
        else:
            st.info("💡 Tính năng mở thư mục trực tiếp áp dụng khi chạy trên máy tính Windows. Trên Streamlit Cloud, anh tải file qua nút 'Tải File (.xlsx)' nhé!")


# Nếu chưa đủ 4 file, hiển thị hướng dẫn chi tiết
if not files_ready:
    st.warning(f"⚠️ **Chưa nhận diện đủ 4 file dữ liệu trong {source_label}.**")

    missing = []
    if not salary_src: missing.append("1. Bảng lương tổng CBCNV (tên file chứa: `luong`, `pr_ttt`, `salary`...)")
    if not pr025_src: missing.append("2. Tờ khai thuế PR025 (tên file chứa: `025`, `pr025`, `05_kk`...)")
    if not pr027_src: missing.append("3. Tờ khai thuế PR027/PR028 (tên file chứa: `027`, `pr027`, `nghi_huu`...)")
    if not gl_src: missing.append("4. Sổ cái TK 3335 (tên file chứa: `3335`, `so_cai`...)")

    st.info("💡 **Vui lòng kiểm tra lại:**\n- " + "\n- ".join(missing))
    st.markdown("""
    > **Gợi ý nhanh:**
    > - Bấm nút **`📁 02_THAM_KHAO`** ở cột bên trái để nạp trực tiếp bộ dữ liệu thực tế của Điện lực Chợ Lớn.
    > - Bấm nút **`📁 01_DAU_VAO`** để chạy bộ dữ liệu kiểm thử mẫu.
    """)
    st.stop()


# ==============================================================================
# THỰC THI KIỂM DÒ KHI ĐÃ CÓ ĐỦ 4 FILE
# ==============================================================================
loader = DataLoader("config/mapping_rules.json")
reconciler = TaxReconciler(tolerance=1000.0)

try:
    with st.spinner("Đang nạp và phân tích dữ liệu 4 file..."):
        df_salary, summary_salary = loader.load_salary_data(salary_src)
        df_pr025, pr025_indicators = loader.load_pr025_data(pr025_src)
        df_pr027, pr027_summary = loader.load_pr027_data(pr027_src)
        df_gl, gl_summary = loader.load_gl_3335_data(gl_src)

        results = reconciler.reconcile_all(
            df_salary=df_salary,
            pr025_indicators=pr025_indicators,
            pr027_summary=pr027_summary,
            gl_3335_summary=gl_summary,
            report_period=report_period
        )

except ValueError as ve:
    st.error(f"❌ **Lỗi cấu trúc dữ liệu:** {ve}")
    st.stop()
except Exception as e:
    st.error(f"❌ **Đã xảy ra lỗi khi xử lý dữ liệu:** {e}")
    st.exception(e)
    st.stop()

# ==============================================================================
# XUẤT FILE BÁO CÁO VÀO THƯ MỤC 03_DAU_RA & CHUẨN BỊ TẢI VỀ
# ==============================================================================
output_dir = Path("03_DAU_RA")
output_dir.mkdir(parents=True, exist_ok=True)
clean_period = report_period.replace(' ', '_').replace('/', '_').replace('á', 'a').replace('Á', 'A')
excel_filename = f"Bao_Cao_Kiem_Do_Thue_TNCN_{clean_period}.xlsx"
output_file_path = output_dir / excel_filename

builder = ExcelReportBuilder(company_name=company_name, report_period=report_period)
builder.export_to_file(results, output_file_path)

# Đọc raw bytes và mã hóa base64 để hỗ trợ tải file trực tiếp với tên .xlsx chuẩn 100%
with open(output_file_path, "rb") as f:
    excel_data = f.read()

b64_excel = base64.b64encode(excel_data).decode("utf-8")

kpis = results["summary_kpis"]


# ==============================================================================
# KHU VỰC THÔNG BÁO TỔNG QUAN & NÚT THAO TÁC FILE XUẤT
# ==============================================================================
top_col1, top_col2, top_col3, top_col4 = st.columns([2.0, 1.0, 1.0, 1.2])

with top_col1:
    if results["all_passed"]:
        st.success("✅ **KẾT QUẢ KIỂM DÒ: TOÀN BỘ SỐ LIỆU ĐÃ KHỚP CHÍNH XÁC 100%!** (Chênh lệch = 0 VND)")
    else:
        st.error("⚠️ **CẢNH BÁO: PHÁT HIỆN SỐ LIỆU CHÊNH LỆCH!** Vui lòng kiểm tra chi tiết các bảng dưới đây.")

with top_col2:
    if st.button("📊 Mở Bằng Excel", use_container_width=True, help="Mở file báo cáo trực tiếp bằng Microsoft Excel trên máy tính"):
        if hasattr(os, "startfile"):
            try:
                os.startfile(str(output_file_path.resolve()))
            except Exception as e:
                st.error(f"Không thể mở file: {e}")
        else:
            st.info("Mở trực tiếp bằng Excel khả dụng khi chạy trên máy tính Windows.")

with top_col3:
    if st.button("📂 Mở 03_DAU_RA", use_container_width=True, help="Mở thư mục 03_DAU_RA trực tiếp trong Windows Explorer"):
        if hasattr(os, "startfile"):
            try:
                os.startfile(str(output_dir.resolve()))
            except Exception as e:
                st.error(f"Không thể mở thư mục: {e}")
        else:
            st.info("Mở thư mục áp dụng khi chạy trên máy tính Windows.")

with top_col4:
    # Nút tải trực tiếp qua HTML5 download attribute (chuẩn file .xlsx 100%, không bao giờ bị đổi thành UUID)
    download_html = f"""
    <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" 
       download="{excel_filename}" 
       style="display: flex; align-items: center; justify-content: center; width: 100%; height: 38px; 
              background-color: #1F4E78; color: white !important; text-decoration: none; border-radius: 8px; 
              font-weight: 600; font-size: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
       📥 Tải File (.xlsx)
    </a>
    """
    st.markdown(download_html, unsafe_allow_html=True)

st.info(f"📁 **File báo cáo đã được lưu sẵn trong thư mục máy tính:** `{output_file_path.resolve()}`")

# 4 THẺ KPI CHÍNH (Định dạng số nguyên, phân cách bằng dấu chấm)
c1, c2, c3, c4 = st.columns(4)

with c1:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Tổng Nhân Sự / Cư Trú</div>
        <div class="kpi-value">{format_vnd(kpis['total_employees'])} <span style="font-size:14px;color:#666">người</span></div>
        <div class="kpi-sub">Cư trú: <b>{format_vnd(kpis['residents_count'])}</b> | Có thuế: <b>{format_vnd(kpis['taxable_employees_count'])}</b></div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    badge_24 = '<span class="badge-matched">KHỚP (0 VND)</span>' if kpis['status_tnct'] == "KHỚP" else f'<span class="badge-diff">LỆCH: {format_vnd(kpis["diff_tnct_resident"])} VND</span>'
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">TNCT Cư Trú (Chỉ tiêu [24])</div>
        <div class="kpi-value">{format_vnd(kpis['taxable_income_resident'])}</div>
        <div class="kpi-sub">PR025: {format_vnd(kpis['pr025_tnct_resident'])} | {badge_24}</div>
    </div>
    """, unsafe_allow_html=True)

with c3:
    badge_27 = '<span class="badge-matched">KHỚP (0 VND)</span>' if kpis['status_with_tax'] == "KHỚP" else f'<span class="badge-diff">LỆCH: {format_vnd(kpis["diff_tnct_with_tax"])} VND</span>'
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">TNCT Nhóm Có Thuế ([27])</div>
        <div class="kpi-value">{format_vnd(kpis['taxable_income_with_tax'])}</div>
        <div class="kpi-sub">PR025: {format_vnd(kpis['pr025_tnct_with_tax'])} | {badge_27}</div>
    </div>
    """, unsafe_allow_html=True)

with c4:
    badge_tax = '<span class="badge-matched">KHỚP (0 VND)</span>' if kpis['status_tax'] == "KHỚP" else f'<span class="badge-diff">LỆCH: {format_vnd(kpis["diff_salary_vs_return"])} VND</span>'
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Thuế Phải Nộp (3 Bên)</div>
        <div class="kpi-value">{format_vnd(kpis['salary_total_tax'])}</div>
        <div class="kpi-sub">Tờ khai: {format_vnd(kpis['tax_return_total'])} | Có 3335: {format_vnd(kpis['gl_credit_turnover'])} | {badge_tax}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ==============================================================================
# KHU VỰC CÁC TABS HIỂN THỊ CHI TIẾT
# ==============================================================================
tab_overview, tab_sheet00, tab_sheet31, tab_sheet41, tab_sheet51 = st.tabs([
    "📋 Tổng Quan & Đối Chiếu Trọng Yếu",
    "📑 Sheet 00 - Tổng Kết Thu Nhập",
    "📊 Sheet 3.1 - TN Chịu Thuế ([24])",
    "👥 Sheet 4.1 - CNV Có Thuế ([27])",
    "💰 Sheet 5.1 - Thuế Khấu Trừ (3 Bên)"
])


# --------------------------------------------------------------------------
# TAB 1: TỔNG QUAN
# --------------------------------------------------------------------------
with tab_overview:
    st.subheader("Bảng Đối Chiếu Tổng Hợp Các Điểm Kiểm Dò Cốt Lõi")

    summary_rows = [
        {
            "Nội Dung Kiểm Dò": "1. Thu nhập chịu thuế cá nhân cư trú (Chỉ tiêu [24])",
            "Số Bảng Lương (VND)": kpis["taxable_income_resident"],
            "Số Tờ Khai Thuế (VND)": kpis["pr025_tnct_resident"],
            "Số Sổ Cái TK 3335 (VND)": "-",
            "chenh_lech": kpis["diff_tnct_resident"],
            "kiem_do": kpis["status_tnct"],
            "Giải Thích Nghiệp Vụ": "Tổng TN - Trừ Hưu trí - Phụ cấp miễn thuế (Chỉ cá nhân cư trú)"
        },
        {
            "Nội Dung Kiểm Dò": "2. TNCT cá nhân cư trú có phát sinh thuế (Chỉ tiêu [27])",
            "Số Bảng Lương (VND)": kpis["taxable_income_with_tax"],
            "Số Tờ Khai Thuế (VND)": kpis["pr025_tnct_with_tax"],
            "Số Sổ Cái TK 3335 (VND)": "-",
            "chenh_lech": kpis["diff_tnct_with_tax"],
            "kiem_do": kpis["status_with_tax"],
            "Giải Thích Nghiệp Vụ": "Tổng TN nhóm có thuế - Phụ cấp miễn thuế + Thu nhập 10%"
        },
        {
            "Nội Dung Kiểm Dò": "3. Tổng số thuế TNCN phải nộp (Lũy tiến + 10%)",
            "Số Bảng Lương (VND)": kpis["salary_total_tax"],
            "Số Tờ Khai Thuế (VND)": kpis["tax_return_total"],
            "Số Sổ Cái TK 3335 (VND)": kpis["gl_credit_turnover"],
            "chenh_lech": kpis["diff_salary_vs_return"],
            "kiem_do": kpis["status_tax"],
            "Giải Thích Nghiệp Vụ": "Đối chiếu 3 bên: Bảng lương vs Tờ khai thuế vs Phát sinh Có TK 3335"
        }
    ]
    df_overview = pd.DataFrame(summary_rows)

    money_cols_ov = ["Số Bảng Lương (VND)", "Số Tờ Khai Thuế (VND)", "Số Sổ Cái TK 3335 (VND)", "chenh_lech"]
    center_cols_ov = ["kiem_do"]

    display_styled_dataframe(
        df_overview,
        money_cols=money_cols_ov,
        center_cols=center_cols_ov,
        diff_col_name="chenh_lech",
        status_col_name="kiem_do"
    )

    st.markdown("""
    > 💡 **Quy tắc kiểm tra kiểm dò:**
    > - **Số liệu:** Toàn bộ là số nguyên, không có số lẻ, phân cách hàng nghìn bằng dấu chấm.
    > - **Ô màu xanh lá:** Chênh lệch = 0 VND (Số liệu đối chiếu hoàn toàn trùng khớp).
    > - **Ô màu cam đỏ:** Chênh lệch khác 0 VND (Phát sinh sai lệch giữa các nguồn, cần rà soát lại).
    """)


# --------------------------------------------------------------------------
# TAB 2: SHEET 00 - TỔNG KẾT
# --------------------------------------------------------------------------
with tab_sheet00:
    st.subheader("Bảng Tổng Hợp Thu Nhập Theo Phòng Ban")
    df_00_sum = results["sheet_00"]["summary"].copy()

    col_rename_00 = {
        "phong_ban": "Phòng Ban",
        "so_nhan_su": "Số Nhân Sự",
        "luong_thang": "Lương Tháng",
        "luong_quy": "Lương Quý/Bổ Sung",
        "thu_nhap_khac": "Thu Nhập Khác",
        "tong_thu_nhap": "TỔNG THU NHẬP",
        "phu_cap_khong_thue": "Phụ Cấp Không Thuế",
        "tru_huu_tri": "Trừ Hưu Trí",
        "thue_tncn_luy_tien": "Thuế Lũy Tiến",
        "thue_tncn_10": "Thuế 10%",
        "tong_thue_tncn": "TỔNG THUẾ TNCN"
    }
    df_00_sum_disp = df_00_sum.rename(columns=col_rename_00)
    money_cols_00_sum = [
        "Lương Tháng", "Lương Quý/Bổ Sung", "Thu Nhập Khác", "TỔNG THU NHẬP",
        "Phụ Cấp Không Thuế", "Trừ Hưu Trí", "Thuế Lũy Tiến", "Thuế 10%", "TỔNG THUẾ TNCN"
    ]
    center_cols_00_sum = ["Số Nhân Sự"]

    display_styled_dataframe(
        df_00_sum_disp,
        money_cols=money_cols_00_sum,
        center_cols=center_cols_00_sum
    )

    with st.expander("🔍 Xem Danh Sách Chi Tiết Toàn Bộ Nhân Viên"):
        df_00_det = results["sheet_00"]["detail"].copy()
        col_rename_00_det = {
            "stt": "STT",
            "employee_id": "Mã NV",
            "full_name": "Họ và Tên",
            "department": "Phòng Ban",
            "resident_status": "Cư Trú",
            "salary_monthly": "Lương Tháng",
            "salary_quarterly": "Lương Quý/Bổ Sung",
            "other_income": "Thu Nhập Khác",
            "total_income": "Tổng Thu Nhập",
            "tax_exempt_allowance": "Phụ Cấp Miễn Thuế",
            "insurance_deduction": "Trừ Hưu Trí / BH",
            "taxable_income": "Thu Nhập Chịu Thuế",
            "progressive_tax": "Thuế Lũy Tiến",
            "tax_10_percent": "Thuế 10%",
            "total_calculated_tax": "Tổng Thuế TNCN"
        }
        df_00_det_disp = df_00_det.rename(columns=col_rename_00_det)
        money_cols_00_det = [
            "Lương Tháng", "Lương Quý/Bổ Sung", "Thu Nhập Khác", "Tổng Thu Nhập",
            "Phụ Cấp Miễn Thuế", "Trừ Hưu Trí / BH", "Thu Nhập Chịu Thuế",
            "Thuế Lũy Tiến", "Thuế 10%", "Tổng Thuế TNCN"
        ]
        center_cols_00_det = ["STT", "Mã NV", "Cư Trú"]

        display_styled_dataframe(
            df_00_det_disp,
            money_cols=money_cols_00_det,
            center_cols=center_cols_00_det
        )


# --------------------------------------------------------------------------
# TAB 3: SHEET 3.1 - TN CHỊU THUẾ
# --------------------------------------------------------------------------
with tab_sheet31:
    st.subheader("Bảng Đối Chiếu Thu Nhập Chịu Thuế Cá Nhân Cư Trú (Chỉ Tiêu [24] PR025)")
    df_31_sum = results["sheet_31"]["summary"].copy()

    col_rename_31 = {
        "chi_tieu": "Chỉ Tiêu Kiểm Dò",
        "ma_chi_tieu": "Mã Chỉ Tiêu",
        "so_bang_luong": "Số Bảng Lương (VND)",
        "so_to_khai_pr025": "Số Tờ Khai PR025 (VND)",
        "chenh_lech": "Chênh Lệch (VND)",
        "kiem_do": "Kiểm Dò",
        "ghi_chu": "Ghi Chú"
    }
    df_31_sum_disp = df_31_sum.rename(columns=col_rename_31)
    money_cols_31_sum = ["Số Bảng Lương (VND)", "Số Tờ Khai PR025 (VND)", "Chênh Lệch (VND)"]
    center_cols_31_sum = ["Mã Chỉ Tiêu", "Kiểm Dò"]

    display_styled_dataframe(
        df_31_sum_disp,
        money_cols=money_cols_31_sum,
        center_cols=center_cols_31_sum,
        diff_col_name="Chênh Lệch (VND)",
        status_col_name="Kiểm Dò"
    )

    with st.expander("🔍 Xem Chi Tiết Tính Thu Nhập Chịu Thuế Từng Cá Nhân Cư Trú"):
        df_31_det = results["sheet_31"]["detail"].copy()
        col_rename_31_det = {
            "stt": "STT",
            "employee_id": "Mã NV",
            "full_name": "Họ và Tên",
            "department": "Phòng Ban",
            "total_income": "Tổng Thu Nhập",
            "insurance_deduction": "Trừ Hưu Trí / BH",
            "tax_exempt_allowance": "Phụ Cấp Không Thuế",
            "taxable_income": "Thu Nhập Chịu Thuế"
        }
        df_31_det_disp = df_31_det.rename(columns=col_rename_31_det)
        money_cols_31_det = ["Tổng Thu Nhập", "Trừ Hưu Trí / BH", "Phụ Cấp Không Thuế", "Thu Nhập Chịu Thuế"]
        center_cols_31_det = ["STT", "Mã NV"]

        display_styled_dataframe(
            df_31_det_disp,
            money_cols=money_cols_31_det,
            center_cols=center_cols_31_det
        )


# --------------------------------------------------------------------------
# TAB 4: SHEET 4.1 - CNV CÓ THUẾ
# --------------------------------------------------------------------------
with tab_sheet41:
    st.subheader("Bảng Đối Chiếu TNCT Nhóm Cá Nhân Cư Trú Có Thuế (Chỉ Tiêu [27] PR025)")
    df_41_sum = results["sheet_41"]["summary"].copy()

    col_rename_41 = {
        "chi_tieu": "Chỉ Tiêu",
        "ma_chi_tieu": "Mã Chỉ Tiêu",
        "so_bang_luong": "Số Bảng Lương (VND)",
        "so_to_khai_pr025": "Số Tờ Khai PR025 (VND)",
        "chenh_lech": "Chênh Lệch (VND)",
        "kiem_do": "Kiểm Dò",
        "ghi_chu": "Ghi Chú"
    }
    df_41_sum_disp = df_41_sum.rename(columns=col_rename_41)
    money_cols_41_sum = ["Số Bảng Lương (VND)", "Số Tờ Khai PR025 (VND)", "Chênh Lệch (VND)"]
    center_cols_41_sum = ["Mã Chỉ Tiêu", "Kiểm Dò"]

    display_styled_dataframe(
        df_41_sum_disp,
        money_cols=money_cols_41_sum,
        center_cols=center_cols_41_sum,
        diff_col_name="Chênh Lệch (VND)",
        status_col_name="Kiểm Dò"
    )

    with st.expander("🔍 Xem Danh Sách Chi Tiết Nhân Sự Thuộc Nhóm Có Thuế"):
        df_41_det = results["sheet_41"]["detail"].copy()
        col_rename_41_det = {
            "stt": "STT",
            "employee_id": "Mã NV",
            "full_name": "Họ và Tên",
            "department": "Phòng Ban",
            "total_income": "Tổng Thu Nhập",
            "tax_exempt_allowance": "Phụ Cấp Không Thuế",
            "income_10_percent": "Thu Nhập 10%",
            "taxable_income": "Thu Nhập Chịu Thuế",
            "progressive_tax": "Thuế Lũy Tiến",
            "tax_10_percent": "Thuế 10%",
            "total_calculated_tax": "Tổng Thuế Phải Nộp"
        }
        df_41_det_disp = df_41_det.rename(columns=col_rename_41_det)
        money_cols_41_det = [
            "Tổng Thu Nhập", "Phụ Cấp Không Thuế", "Thu Nhập 10%",
            "Thu Nhập Chịu Thuế", "Thuế Lũy Tiến", "Thuế 10%", "Tổng Thuế Phải Nộp"
        ]
        center_cols_41_det = ["STT", "Mã NV"]

        display_styled_dataframe(
            df_41_det_disp,
            money_cols=money_cols_41_det,
            center_cols=center_cols_41_det
        )


# --------------------------------------------------------------------------
# TAB 5: SHEET 5.1 - THUẾ KHẤU TRỪ
# --------------------------------------------------------------------------
with tab_sheet51:
    st.subheader("Bảng Đối Chiếu Số Thuế Khấu Trừ 3 Bên: Bảng Lương - Tờ Khai - Sổ Cái TK 3335")
    df_51_sum = results["sheet_51"]["summary"].copy()

    col_rename_51 = {
        "hang_muc": "Hạng Mục Đối Chiếu",
        "bang_luong": "Số Bảng Lương (VND)",
        "to_khai_thue": "Số Tờ Khai Thuế (VND)",
        "so_cai_tk_3335": "Số Sổ Cái TK 3335 (VND)",
        "lech_luong_vs_tokhai": "Chênh Lệch (VND)",
        "kiem_do": "Kiểm Dò",
        "ghi_chu": "Ghi Chú Nghiệp Vụ"
    }
    df_51_sum_disp = df_51_sum.rename(columns=col_rename_51)
    money_cols_51_sum = [
        "Số Bảng Lương (VND)", "Số Tờ Khai Thuế (VND)", "Số Sổ Cái TK 3335 (VND)", "Chênh Lệch (VND)"
    ]
    center_cols_51_sum = ["Kiểm Dò"]

    display_styled_dataframe(
        df_51_sum_disp,
        money_cols=money_cols_51_sum,
        center_cols=center_cols_51_sum,
        diff_col_name="Chênh Lệch (VND)",
        status_col_name="Kiểm Dò"
    )

    st.subheader("Phân Bổ Số Thuế TNCN Theo Phòng Ban / Bộ Phận")
    df_51_dept = results["sheet_51"]["detail_dept"].copy()
    col_rename_dept = {
        "phong_ban": "Phòng Ban",
        "so_nhan_su": "Số Nhân Sự",
        "so_nguoi_co_thue": "Số Người Có Thuế",
        "thue_luy_tien": "Thuế Lũy Tiến (VND)",
        "thue_10_percent": "Thuế 10% (VND)",
        "tong_thue": "TỔNG THUẾ TNCN (VND)"
    }
    df_51_dept_disp = df_51_dept.rename(columns=col_rename_dept)
    money_cols_dept = ["Thuế Lũy Tiến (VND)", "Thuế 10% (VND)", "TỔNG THUẾ TNCN (VND)"]
    center_cols_dept = ["Số Nhân Sự", "Số Người Có Thuế"]

    display_styled_dataframe(
        df_51_dept_disp,
        money_cols=money_cols_dept,
        center_cols=center_cols_dept
    )

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #888; font-size: 13px;'>"
    "Ứng dụng Kiểm dò Dữ liệu Thuế TNCN © 2026 | Dành cho Công ty Điện lực | Xử lý In-Memory An Toàn & Bảo Mật"
    "</div>",
    unsafe_allow_html=True
)
