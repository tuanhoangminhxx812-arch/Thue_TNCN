"""
Module: report_builder.py
Mục đích: Định dạng và xuất file Excel nhiều sheet bằng openpyxl cho Công ty Điện lực,
          đảm bảo định dạng chuẩn kế toán:
          - TẤT CẢ SỐ LÀ DẠNG SỐ (numeric int/float), không lưu dạng chuỗi text.
          - CĂN PHẢI TOÀN BỘ CỘT SỐ, CĂN TRÁI TOÀN BỘ CỘT CHỮ, CĂN GIỮA MÃ/STT/TRẠNG THÁI.
          - Định dạng phân cách hàng nghìn chuẩn kế toán (#,##0;(#,##0);"-").
          - Viền ô toàn bộ bảng, màu header chuyên nghiệp.
          - Tô màu cảnh báo kiểm dò (Xanh khi Khớp, Đỏ khi Lệch).
          - Auto-fit độ rộng cột tự động.
"""

import io
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Union, Optional
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ==============================================================================
# ĐỊNH DẠNG MÀU SẮC & STYLES CHUẨN KẾ TOÁN DOANH NGHIỆP
# ==============================================================================

# Màu Header chính: Xanh Navy đậm sang trọng
COLOR_HEADER_NAVY = "1F4E78"
COLOR_HEADER_SUB = "2F5597"
COLOR_WHITE = "FFFFFF"

# Màu Cảnh báo Kiểm dò
COLOR_MATCH_BG = "E2EFDA"     # Xanh lục nhạt
COLOR_MATCH_FG = "276A3C"     # Chữ xanh lục đậm
COLOR_DIFF_BG = "FCE4D6"      # Cam đỏ nhạt
COLOR_DIFF_FG = "C00000"      # Chữ đỏ đậm

# Màu dòng Tổng cộng
COLOR_TOTAL_BG = "D9E1F2"     # Xanh lam nhạt

# Đường kẻ viền
THIN_BORDER_SIDE = Side(border_style="thin", color="BFBFBF")
MEDIUM_BORDER_SIDE = Side(border_style="medium", color="1F4E78")
DOUBLE_BOTTOM_SIDE = Side(border_style="double", color="1F4E78")

BORDER_ALL_THIN = Border(
    left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE,
    top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE
)
BORDER_TOTAL = Border(
    left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE,
    top=THIN_BORDER_SIDE, bottom=DOUBLE_BOTTOM_SIDE
)

# Định dạng số chuẩn kế toán (phân cách hàng nghìn, số âm trong ngoặc, số 0 hiển thị dấu gạch ngang)
FORMAT_CURRENCY = '#,##0;(#,##0);"-"'
FORMAT_NUMBER = '#,##0'


def format_cell_value(
    cell: openpyxl.cell.cell.Cell,
    value: Any,
    is_numeric: bool = False,
    is_center: bool = False,
    is_bold: bool = False,
    is_total: bool = False
):
    """
    Hàm chuẩn hóa ghi dữ liệu và định dạng ô:
    - Nếu là số: ép kiểu về float/int chuẩn, căn lề PHẢI, định dạng #,##0.
    - Nếu là chữ: căn lề TRÁI.
    - Nếu là mã/STT/trạng thái: căn lề GIỮA.
    """
    if is_numeric or (isinstance(value, (int, float)) and not is_center):
        if value is None or pd.isna(value) or value == "-" or value == "":
            num_val = 0
        else:
            try:
                num_val = int(round(float(value)))
            except (ValueError, TypeError):
                num_val = 0
        cell.value = num_val
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = FORMAT_CURRENCY
    else:
        # Giá trị dạng chữ hoặc định danh
        cell.value = "" if pd.isna(value) else str(value)
        if is_center:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

    if is_bold or is_total:
        cell.font = Font(name="Calibri", size=11, bold=True)
    if is_total:
        cell.border = BORDER_TOTAL
        cell.fill = PatternFill(start_color=COLOR_TOTAL_BG, end_color=COLOR_TOTAL_BG, fill_type="solid")
    else:
        cell.border = BORDER_ALL_THIN


class ExcelReportBuilder:
    """Xây dựng và định dạng file Excel báo cáo kiểm dò Thuế TNCN cho Công ty Điện lực."""

    def __init__(
        self,
        company_name: str = "TỔNG CÔNG TY ĐIỆN LỰC TP. HỒ CHÍ MINH - CÔNG TY ĐIỆN LỰC CHỢ LỚN",
        report_period: str = ""
    ):
        self.company_name = company_name
        self.report_period = report_period

    def build_workbook(self, reconciliation_data: Dict[str, Any]) -> openpyxl.Workbook:
        """Tạo đối tượng Workbook với 4 sheet hoàn chỉnh và định dạng chuyên nghiệp."""
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        # 1. Sheet 00_Tong_Ket
        ws_00 = wb.create_sheet(title="00_Tong_Ket")
        self._render_sheet_00(ws_00, reconciliation_data)

        # 2. Sheet 3.1_TN_Chiu_Thue
        ws_31 = wb.create_sheet(title="3.1_TN_Chiu_Thue")
        self._render_sheet_31(ws_31, reconciliation_data)

        # 3. Sheet 4.1_CNV_Co_Thue
        ws_41 = wb.create_sheet(title="4.1_CNV_Co_Thue")
        self._render_sheet_41(ws_41, reconciliation_data)

        # 4. Sheet 5.1_Thue_Khau_Tru
        ws_51 = wb.create_sheet(title="5.1_Thue_Khau_Tru")
        self._render_sheet_51(ws_51, reconciliation_data)

        return wb

    def export_to_file(self, reconciliation_data: Dict[str, Any], output_path: Union[str, Path]) -> str:
        """Xuất báo cáo ra đường dẫn file cục bộ (.xlsx)."""
        wb = self.build_workbook(reconciliation_data)
        out_p = Path(output_path)
        out_p.parent.mkdir(parents=True, exist_ok=True)
        wb.save(str(out_p))
        return str(out_p)

    def export_to_bytes(self, reconciliation_data: Dict[str, Any]) -> io.BytesIO:
        """Xuất báo cáo ra luồng bộ nhớ io.BytesIO cho ứng dụng Streamlit web."""
        wb = self.build_workbook(reconciliation_data)
        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)
        return bio

    def _add_title_block(self, ws: openpyxl.worksheet.worksheet.Worksheet, title: str, subtitle: str = "") -> int:
        """Thêm khối tiêu đề trang trọng đầu trang tính."""
        ws.cell(row=1, column=1, value=self.company_name)
        ws.cell(row=1, column=1).font = Font(name="Calibri", size=11, bold=True, color="1F4E78")

        ws.cell(row=2, column=1, value=title.upper())
        ws.cell(row=2, column=1).font = Font(name="Calibri", size=14, bold=True, color=COLOR_HEADER_NAVY)

        period_text = f"Kỳ báo cáo: {self.report_period or 'Toàn kỳ'} | {subtitle}"
        ws.cell(row=3, column=1, value=period_text)
        ws.cell(row=3, column=1).font = Font(name="Calibri", size=10, italic=True, color="333333")

        ws.cell(row=4, column=1, value=f"Ngày xuất báo cáo: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        ws.cell(row=4, column=1).font = Font(name="Calibri", size=9, italic=True, color="7F7F7F")

        return 6

    def _auto_fit_columns(self, ws: openpyxl.worksheet.worksheet.Worksheet, max_col: int = None, min_width: int = 12):
        """Tự động co giãn độ rộng cột dựa trên độ dài nội dung."""
        for col_idx in range(1, (max_col or ws.max_column) + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for row in range(5, ws.max_row + 1):
                cell_val = ws.cell(row=row, column=col_idx).value
                if cell_val is not None:
                    if isinstance(cell_val, (int, float)):
                        val_str = f"{cell_val:,.0f}"
                    else:
                        val_str = str(cell_val)
                    max_len = max(max_len, len(val_str))
            
            adjusted_width = max(max_len + 4, min_width)
            ws.column_dimensions[col_letter].width = min(adjusted_width, 45)

    def _style_header_row(self, ws: openpyxl.worksheet.worksheet.Worksheet, row_idx: int, cols_count: int, bg_color: str = COLOR_HEADER_NAVY):
        """Áp dụng màu nền, font chữ trắng in đậm và căn giữa cho Header."""
        fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        font = Font(name="Calibri", size=11, bold=True, color=COLOR_WHITE)
        align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for col_idx in range(1, cols_count + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = fill
            cell.font = font
            cell.alignment = align
            cell.border = BORDER_ALL_THIN
        ws.row_dimensions[row_idx].height = 28

    def _render_sheet_00(self, ws: openpyxl.worksheet.worksheet.Worksheet, data: Dict[str, Any]):
        """Render sheet 00_Tong_Ket: Tổng hợp các nguồn thu nhập."""
        start_row = self._add_title_block(
            ws,
            title="BÁO CÁO TỔNG HỢP CÁC NGUỒN THU NHẬP CBCNV",
            subtitle="Phân tích tổng thu nhập: Lương tháng, Lương quý/bổ sung và Các khoản khác"
        )

        sheet_data = data["sheet_00"]
        df_summary = sheet_data["summary"]
        df_detail = sheet_data["detail"]

        # -------------------------------------------------------------
        # BẢNG 1: TỔNG HỢP THEO PHÒNG BAN
        # -------------------------------------------------------------
        ws.cell(row=start_row, column=1, value="I. BẢNG TỔNG HỢP THU NHẬP THEO PHÒNG BAN / BỘ PHẬN").font = Font(name="Calibri", size=11, bold=True, color=COLOR_HEADER_NAVY)
        start_row += 1

        headers_summary = [
            "STT", "Phòng Ban / Bộ Phận", "Số Nhân Sự", "Lương Tháng (VND)",
            "Lương Quý / Bổ Sung (VND)", "Thu Nhập Khác (VND)", "TỔNG THU NHẬP (VND)",
            "Phụ Cấp Không Thuế (VND)", "Trừ Hưu Trí / BH (VND)",
            "Thuế Lũy Tiến (VND)", "Thuế 10% (VND)", "TỔNG THUẾ TNCN (VND)"
        ]

        self._style_header_row(ws, start_row, len(headers_summary))
        for col_idx, h in enumerate(headers_summary, 1):
            ws.cell(row=start_row, column=col_idx, value=h)
        start_row += 1

        for r_idx, row in df_summary.iterrows():
            curr_row = start_row + r_idx
            is_total = "TỔNG CỘNG" in str(row["phong_ban"]).upper()

            # STT (Căn giữa)
            format_cell_value(ws.cell(row=curr_row, column=1), r_idx + 1 if not is_total else "", is_center=True, is_total=is_total)
            # Phòng ban (Chữ - Căn trái)
            format_cell_value(ws.cell(row=curr_row, column=2), row["phong_ban"], is_center=False, is_bold=is_total, is_total=is_total)
            # Số nhân sự (Số - Căn phải)
            format_cell_value(ws.cell(row=curr_row, column=3), row["so_nhan_su"], is_numeric=True, is_total=is_total)
            # Các cột số tiền (Số - Căn phải, định dạng #,##0)
            format_cell_value(ws.cell(row=curr_row, column=4), row["luong_thang"], is_numeric=True, is_total=is_total)
            format_cell_value(ws.cell(row=curr_row, column=5), row["luong_quy"], is_numeric=True, is_total=is_total)
            format_cell_value(ws.cell(row=curr_row, column=6), row["thu_nhap_khac"], is_numeric=True, is_total=is_total)
            format_cell_value(ws.cell(row=curr_row, column=7), row["tong_thu_nhap"], is_numeric=True, is_total=is_total)
            format_cell_value(ws.cell(row=curr_row, column=8), row["phu_cap_khong_thue"], is_numeric=True, is_total=is_total)
            format_cell_value(ws.cell(row=curr_row, column=9), row["tru_huu_tri"], is_numeric=True, is_total=is_total)
            format_cell_value(ws.cell(row=curr_row, column=10), row["thue_tncn_luy_tien"], is_numeric=True, is_total=is_total)
            format_cell_value(ws.cell(row=curr_row, column=11), row["thue_tncn_10"], is_numeric=True, is_total=is_total)
            format_cell_value(ws.cell(row=curr_row, column=12), row["tong_thue_tncn"], is_numeric=True, is_total=is_total)

        # -------------------------------------------------------------
        # BẢNG 2: DANH SÁCH CHI TIẾT TỪNG NHÂN VIÊN
        # -------------------------------------------------------------
        next_row = start_row + len(df_summary) + 2
        ws.cell(row=next_row, column=1, value="II. DANH SÁCH CHI TIẾT THU NHẬP VÀ THUẾ TỪNG NHÂN VIÊN").font = Font(name="Calibri", size=11, bold=True, color=COLOR_HEADER_NAVY)
        next_row += 1

        headers_detail = [
            "STT", "Mã NV", "Họ Và Tên", "Phòng Ban", "Cư Trú",
            "Lương Tháng (VND)", "Lương Quý (VND)", "Thu Nhập Khác (VND)",
            "TỔNG THU NHẬP (VND)", "Phụ Cấp Miễn Thuế (VND)", "Trừ Hưu Trí (VND)",
            "Thuế Lũy Tiến (VND)", "Thuế 10% (VND)", "Tổng Thuế Phải Nộp (VND)"
        ]

        self._style_header_row(ws, next_row, len(headers_detail), bg_color=COLOR_HEADER_SUB)
        for col_idx, h in enumerate(headers_detail, 1):
            ws.cell(row=next_row, column=col_idx, value=h)
        next_row += 1

        for r_idx, row in df_detail.iterrows():
            curr_row = next_row + r_idx
            format_cell_value(ws.cell(row=curr_row, column=1), row["stt"], is_center=True)
            format_cell_value(ws.cell(row=curr_row, column=2), row["employee_id"], is_center=True)
            format_cell_value(ws.cell(row=curr_row, column=3), row["full_name"], is_center=False)
            format_cell_value(ws.cell(row=curr_row, column=4), row["department"], is_center=False)
            format_cell_value(ws.cell(row=curr_row, column=5), row["resident_status"], is_center=True)
            format_cell_value(ws.cell(row=curr_row, column=6), row["salary_monthly"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=7), row["salary_quarterly"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=8), row["other_income"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=9), row["total_income"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=10), row["tax_exempt_allowance"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=11), row["insurance_deduction"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=12), row["progressive_tax"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=13), row["tax_10_percent"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=14), row["total_calculated_tax"], is_numeric=True)

        self._auto_fit_columns(ws, len(headers_detail))

    def _render_sheet_31(self, ws: openpyxl.worksheet.worksheet.Worksheet, data: Dict[str, Any]):
        """Render sheet 3.1_TN_Chiu_Thue: Đối chiếu Thu nhập chịu thuế cá nhân cư trú với PR025."""
        start_row = self._add_title_block(
            ws,
            title="ĐỐI CHIẾU THU NHẬP CHỊU THUẾ CÁ NHÂN CƯ TRÚ (CHỈ TIÊU [24] / [27] PR025)",
            subtitle="Công thức: TN Chịu Thuế = Tổng Thu Nhập - Trừ Hưu Trí - Phụ Cấp Không Thuế"
        )

        sheet_data = data["sheet_31"]
        df_summary = sheet_data["summary"]
        df_detail = sheet_data["detail"]

        # -------------------------------------------------------------
        # BẢNG 1: BẢNG TỔNG HỢP KIỂM DÒ
        # -------------------------------------------------------------
        ws.cell(row=start_row, column=1, value="I. KẾT QUẢ ĐỐI CHIẾU VÀ KIỂM DÒ VỚI TỜ KHAI PR025").font = Font(name="Calibri", size=11, bold=True, color=COLOR_HEADER_NAVY)
        start_row += 1

        headers_sum = [
            "Chỉ Tiêu Kiểm Dò", "Mã Chỉ Tiêu", "Số Theo Bảng Lương (VND)",
            "Số Theo Tờ Khai PR025 (VND)", "Chênh Lệch (VND)", "Kiểm Dò", "Ghi Chú"
        ]
        self._style_header_row(ws, start_row, len(headers_sum))
        for col_idx, h in enumerate(headers_sum, 1):
            ws.cell(row=start_row, column=col_idx, value=h)
        start_row += 1

        for r_idx, row in df_summary.iterrows():
            curr_row = start_row + r_idx
            is_bold = "TỔNG" in str(row["chi_tieu"]).upper()

            # Cột 1: Chỉ tiêu (Chữ - Căn trái)
            format_cell_value(ws.cell(row=curr_row, column=1), row["chi_tieu"], is_center=False, is_bold=is_bold)
            # Cột 2: Mã chỉ tiêu (Căn giữa)
            format_cell_value(ws.cell(row=curr_row, column=2), row["ma_chi_tieu"], is_center=True, is_bold=is_bold)
            # Cột 3: Bảng lương (Số - Căn phải)
            format_cell_value(ws.cell(row=curr_row, column=3), row["so_bang_luong"], is_numeric=True, is_bold=is_bold)
            # Cột 4: Tờ khai PR025 (Số - Căn phải)
            format_cell_value(ws.cell(row=curr_row, column=4), row["so_to_khai_pr025"], is_numeric=True, is_bold=is_bold)
            # Cột 5: Chênh lệch (Số - Căn phải)
            format_cell_value(ws.cell(row=curr_row, column=5), row["chenh_lech"], is_numeric=True, is_bold=is_bold)
            # Cột 6: Trạng thái Kiểm dò (Căn giữa)
            format_cell_value(ws.cell(row=curr_row, column=6), row["kiem_do"], is_center=True, is_bold=True)
            # Cột 7: Ghi chú (Chữ - Căn trái)
            format_cell_value(ws.cell(row=curr_row, column=7), row["ghi_chu"], is_center=False)

            # Định dạng màu sắc cảnh báo cột KIỂM DÒ (Cột 6) và CHÊNH LỆCH (Cột 5)
            status_val = str(row["kiem_do"]).strip()
            diff_val = row["chenh_lech"]
            cell_diff = ws.cell(row=curr_row, column=5)
            cell_status = ws.cell(row=curr_row, column=6)

            if status_val == "KHỚP" or (isinstance(diff_val, (int, float)) and abs(diff_val) <= 1000):
                cell_status.fill = PatternFill(start_color=COLOR_MATCH_BG, end_color=COLOR_MATCH_BG, fill_type="solid")
                cell_status.font = Font(name="Calibri", size=11, bold=True, color=COLOR_MATCH_FG)
                cell_diff.fill = PatternFill(start_color=COLOR_MATCH_BG, end_color=COLOR_MATCH_BG, fill_type="solid")
            elif status_val == "LỆCH" or (isinstance(diff_val, (int, float)) and abs(diff_val) > 1000):
                cell_status.fill = PatternFill(start_color=COLOR_DIFF_BG, end_color=COLOR_DIFF_BG, fill_type="solid")
                cell_status.font = Font(name="Calibri", size=11, bold=True, color=COLOR_DIFF_FG)
                cell_diff.fill = PatternFill(start_color=COLOR_DIFF_BG, end_color=COLOR_DIFF_BG, fill_type="solid")
                cell_diff.font = Font(name="Calibri", size=11, bold=True, color=COLOR_DIFF_FG)

        # -------------------------------------------------------------
        # BẢNG 2: CHI TIẾT TỪNG NHÂN SỰ CƯ TRÚ
        # -------------------------------------------------------------
        next_row = start_row + len(df_summary) + 2
        ws.cell(row=next_row, column=1, value="II. CHI TIẾT TÍNH THU NHẬP CHỊU THUẾ TỪNG CÁ NHÂN CƯ TRÚ").font = Font(name="Calibri", size=11, bold=True, color=COLOR_HEADER_NAVY)
        next_row += 1

        headers_detail = [
            "STT", "Mã NV", "Họ Và Tên", "Phòng Ban",
            "Tổng Thu Nhập (VND)", "Trừ Hưu Trí / BH (VND)",
            "Phụ Cấp Không Thuế (VND)", "THU NHẬP CHỊU THUẾ (VND)"
        ]
        self._style_header_row(ws, next_row, len(headers_detail), bg_color=COLOR_HEADER_SUB)
        for col_idx, h in enumerate(headers_detail, 1):
            ws.cell(row=next_row, column=col_idx, value=h)
        next_row += 1

        for r_idx, row in df_detail.iterrows():
            curr_row = next_row + r_idx
            format_cell_value(ws.cell(row=curr_row, column=1), row["stt"], is_center=True)
            format_cell_value(ws.cell(row=curr_row, column=2), row["employee_id"], is_center=True)
            format_cell_value(ws.cell(row=curr_row, column=3), row["full_name"], is_center=False)
            format_cell_value(ws.cell(row=curr_row, column=4), row["department"], is_center=False)
            format_cell_value(ws.cell(row=curr_row, column=5), row["total_income"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=6), row["insurance_deduction"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=7), row["tax_exempt_allowance"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=8), row["taxable_income"], is_numeric=True)

        self._auto_fit_columns(ws, max(len(headers_sum), len(headers_detail)))

    def _render_sheet_41(self, ws: openpyxl.worksheet.worksheet.Worksheet, data: Dict[str, Any]):
        """Render sheet 4.1_CNV_Co_Thue: Đối chiếu cá nhân cư trú có phát sinh thuế với PR025."""
        start_row = self._add_title_block(
            ws,
            title="ĐỐI CHIẾU TNCT CỦA NHÓM CNV CÓ PHÁT SINH THUẾ (CHỈ TIÊU [27] / [30] PR025)",
            subtitle="Công thức: TNCT Có Thuế = Tổng TN Nhóm Có Thuế - Phụ Cấp Miễn Thuế + Thu Nhập 10%"
        )

        sheet_data = data["sheet_41"]
        df_summary = sheet_data["summary"]
        df_detail = sheet_data["detail"]

        # -------------------------------------------------------------
        # BẢNG 1: BẢNG TỔNG HỢP KIỂM DÒ
        # -------------------------------------------------------------
        ws.cell(row=start_row, column=1, value="I. TỔNG HỢP KIỂM DÒ VỚI CHỈ TIÊU TRÊN TỜ KHAI PR025").font = Font(name="Calibri", size=11, bold=True, color=COLOR_HEADER_NAVY)
        start_row += 1

        headers_sum = [
            "Chỉ Tiêu", "Mã Chỉ Tiêu", "Số Theo Bảng Lương (VND)",
            "Số Theo Tờ Khai PR025 (VND)", "Chênh Lệch (VND)", "Kiểm Dò", "Ghi Chú"
        ]
        self._style_header_row(ws, start_row, len(headers_sum))
        for col_idx, h in enumerate(headers_sum, 1):
            ws.cell(row=start_row, column=col_idx, value=h)
        start_row += 1

        for r_idx, row in df_summary.iterrows():
            curr_row = start_row + r_idx
            is_bold = "TỔNG" in str(row["chi_tieu"]).upper()

            format_cell_value(ws.cell(row=curr_row, column=1), row["chi_tieu"], is_center=False, is_bold=is_bold)
            format_cell_value(ws.cell(row=curr_row, column=2), row["ma_chi_tieu"], is_center=True, is_bold=is_bold)
            format_cell_value(ws.cell(row=curr_row, column=3), row["so_bang_luong"], is_numeric=True, is_bold=is_bold)
            format_cell_value(ws.cell(row=curr_row, column=4), row["so_to_khai_pr025"], is_numeric=True, is_bold=is_bold)
            format_cell_value(ws.cell(row=curr_row, column=5), row["chenh_lech"], is_numeric=True, is_bold=is_bold)
            format_cell_value(ws.cell(row=curr_row, column=6), row["kiem_do"], is_center=True, is_bold=True)
            format_cell_value(ws.cell(row=curr_row, column=7), row["ghi_chu"], is_center=False)

            status_val = str(row["kiem_do"]).strip()
            diff_val = row["chenh_lech"]
            cell_diff = ws.cell(row=curr_row, column=5)
            cell_status = ws.cell(row=curr_row, column=6)

            if status_val == "KHỚP" or (isinstance(diff_val, (int, float)) and abs(diff_val) <= 1000):
                cell_status.fill = PatternFill(start_color=COLOR_MATCH_BG, end_color=COLOR_MATCH_BG, fill_type="solid")
                cell_status.font = Font(name="Calibri", size=11, bold=True, color=COLOR_MATCH_FG)
                cell_diff.fill = PatternFill(start_color=COLOR_MATCH_BG, end_color=COLOR_MATCH_BG, fill_type="solid")
            elif status_val == "LỆCH" or (isinstance(diff_val, (int, float)) and abs(diff_val) > 1000):
                cell_status.fill = PatternFill(start_color=COLOR_DIFF_BG, end_color=COLOR_DIFF_BG, fill_type="solid")
                cell_status.font = Font(name="Calibri", size=11, bold=True, color=COLOR_DIFF_FG)
                cell_diff.fill = PatternFill(start_color=COLOR_DIFF_BG, end_color=COLOR_DIFF_BG, fill_type="solid")
                cell_diff.font = Font(name="Calibri", size=11, bold=True, color=COLOR_DIFF_FG)

        # -------------------------------------------------------------
        # BẢNG 2: CHI TIẾT NHÓM CÓ THUẾ
        # -------------------------------------------------------------
        next_row = start_row + len(df_summary) + 2
        ws.cell(row=next_row, column=1, value="II. DANH SÁCH CHI TIẾT NHÂN SỰ CÓ PHÁT SINH THUẾ TNCN").font = Font(name="Calibri", size=11, bold=True, color=COLOR_HEADER_NAVY)
        next_row += 1

        headers_detail = [
            "STT", "Mã NV", "Họ Và Tên", "Phòng Ban",
            "Tổng Thu Nhập (VND)", "Phụ Cấp Miễn Thuế (VND)", "Thu Nhập 10% (VND)",
            "Thu Nhập Chịu Thuế (VND)", "Thuế Lũy Tiến (VND)", "Thuế 10% (VND)", "TỔNG THUẾ (VND)"
        ]
        self._style_header_row(ws, next_row, len(headers_detail), bg_color=COLOR_HEADER_SUB)
        for col_idx, h in enumerate(headers_detail, 1):
            ws.cell(row=next_row, column=col_idx, value=h)
        next_row += 1

        for r_idx, row in df_detail.iterrows():
            curr_row = next_row + r_idx
            format_cell_value(ws.cell(row=curr_row, column=1), row["stt"], is_center=True)
            format_cell_value(ws.cell(row=curr_row, column=2), row["employee_id"], is_center=True)
            format_cell_value(ws.cell(row=curr_row, column=3), row["full_name"], is_center=False)
            format_cell_value(ws.cell(row=curr_row, column=4), row["department"], is_center=False)
            format_cell_value(ws.cell(row=curr_row, column=5), row["total_income"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=6), row["tax_exempt_allowance"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=7), row["income_10_percent"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=8), row["taxable_income"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=9), row["progressive_tax"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=10), row["tax_10_percent"], is_numeric=True)
            format_cell_value(ws.cell(row=curr_row, column=11), row["total_calculated_tax"], is_numeric=True)

        self._auto_fit_columns(ws, max(len(headers_sum), len(headers_detail)))

    def _render_sheet_51(self, ws: openpyxl.worksheet.worksheet.Worksheet, data: Dict[str, Any]):
        """Render sheet 5.1_Thue_Khau_Tru: Đối chiếu 3 bên (Bảng lương, Tờ khai PR027/PR028, Sổ cái TK 3335)."""
        start_row = self._add_title_block(
            ws,
            title="ĐỐI CHIẾU SỐ THUẾ KHẤU TRỪ 3 BÊN: BẢNG LƯƠNG - TỜ KHAI - SỔ CÁI TK 3335",
            subtitle="Kiểm dò số thuế phát sinh phải nộp: Bảng lương vs Tờ khai thuế vs Phát sinh Có TK 3335"
        )

        sheet_data = data["sheet_51"]
        df_summary = sheet_data["summary"]
        df_dept = sheet_data["detail_dept"]

        # -------------------------------------------------------------
        # BẢNG 1: BẢNG TỔNG HỢP ĐỐI CHIẾU 3 BÊN
        # -------------------------------------------------------------
        ws.cell(row=start_row, column=1, value="I. BẢNG ĐỐI CHIẾU SỐ THUẾ PHẢI NỘP GIỮA CÁC NGUỒN").font = Font(name="Calibri", size=11, bold=True, color=COLOR_HEADER_NAVY)
        start_row += 1

        headers_sum = [
            "Hạng Mục Đối Chiếu", "Bảng Lương (VND)", "Tờ Khai Thuế (VND)",
            "Sổ Cái TK 3335 (VND)", "Chênh Lệch (VND)", "Kiểm Dò", "Ghi Chú"
        ]
        self._style_header_row(ws, start_row, len(headers_sum))
        for col_idx, h in enumerate(headers_sum, 1):
            ws.cell(row=start_row, column=col_idx, value=h)
        start_row += 1

        for r_idx, row in df_summary.iterrows():
            curr_row = start_row + r_idx
            is_bold = "TỔNG" in str(row["hang_muc"]).upper()

            # Cột 1: Hạng mục (Chữ - Căn trái)
            format_cell_value(ws.cell(row=curr_row, column=1), row["hang_muc"], is_center=False, is_bold=is_bold)
            # Cột 2: Bảng lương (Số - Căn phải)
            format_cell_value(ws.cell(row=curr_row, column=2), row["bang_luong"], is_numeric=True, is_bold=is_bold)
            # Cột 3: Tờ khai thuế (Số - Căn phải)
            format_cell_value(ws.cell(row=curr_row, column=3), row["to_khai_thue"], is_numeric=True, is_bold=is_bold)
            # Cột 4: Sổ cái TK 3335 (Số - Căn phải)
            format_cell_value(ws.cell(row=curr_row, column=4), row["so_cai_tk_3335"], is_numeric=True, is_bold=is_bold)
            # Cột 5: Chênh lệch (Số - Căn phải)
            format_cell_value(ws.cell(row=curr_row, column=5), row["lech_luong_vs_tokhai"], is_numeric=True, is_bold=is_bold)
            # Cột 6: Trạng thái (Căn giữa)
            format_cell_value(ws.cell(row=curr_row, column=6), row["kiem_do"], is_center=True, is_bold=True)
            # Cột 7: Ghi chú (Chữ - Căn trái)
            format_cell_value(ws.cell(row=curr_row, column=7), row["ghi_chu"], is_center=False)

            status_val = str(row["kiem_do"]).strip()
            diff_val = row["lech_luong_vs_tokhai"]
            cell_diff = ws.cell(row=curr_row, column=5)
            cell_status = ws.cell(row=curr_row, column=6)

            if status_val == "KHỚP" or (isinstance(diff_val, (int, float)) and abs(diff_val) <= 1000):
                cell_status.fill = PatternFill(start_color=COLOR_MATCH_BG, end_color=COLOR_MATCH_BG, fill_type="solid")
                cell_status.font = Font(name="Calibri", size=11, bold=True, color=COLOR_MATCH_FG)
                cell_diff.fill = PatternFill(start_color=COLOR_MATCH_BG, end_color=COLOR_MATCH_BG, fill_type="solid")
            elif status_val == "LỆCH" or status_val == "CẢNH BÁO" or (isinstance(diff_val, (int, float)) and abs(diff_val) > 1000):
                cell_status.fill = PatternFill(start_color=COLOR_DIFF_BG, end_color=COLOR_DIFF_BG, fill_type="solid")
                cell_status.font = Font(name="Calibri", size=11, bold=True, color=COLOR_DIFF_FG)
                cell_diff.fill = PatternFill(start_color=COLOR_DIFF_BG, end_color=COLOR_DIFF_BG, fill_type="solid")
                cell_diff.font = Font(name="Calibri", size=11, bold=True, color=COLOR_DIFF_FG)

        # -------------------------------------------------------------
        # BẢNG 2: PHÂN BỔ THUẾ THEO PHÒNG BAN
        # -------------------------------------------------------------
        next_row = start_row + len(df_summary) + 2
        ws.cell(row=next_row, column=1, value="II. PHÂN BỔ SỐ THUẾ TNCN THEO PHÒNG BAN / BỘ PHẬN").font = Font(name="Calibri", size=11, bold=True, color=COLOR_HEADER_NAVY)
        next_row += 1

        headers_dept = [
            "STT", "Phòng Ban", "Số Nhân Sự", "Số Người Có Thuế",
            "Thuế Lũy Tiến (VND)", "Thuế 10% (VND)", "TỔNG THUẾ TNCN (VND)"
        ]
        self._style_header_row(ws, next_row, len(headers_dept), bg_color=COLOR_HEADER_SUB)
        for col_idx, h in enumerate(headers_dept, 1):
            ws.cell(row=next_row, column=col_idx, value=h)
        next_row += 1

        for r_idx, row in df_dept.iterrows():
            curr_row = next_row + r_idx
            is_total = "TỔNG CỘNG" in str(row["phong_ban"]).upper()

            format_cell_value(ws.cell(row=curr_row, column=1), r_idx + 1 if not is_total else "", is_center=True, is_total=is_total)
            format_cell_value(ws.cell(row=curr_row, column=2), row["phong_ban"], is_center=False, is_bold=is_total, is_total=is_total)
            format_cell_value(ws.cell(row=curr_row, column=3), row["so_nhan_su"], is_numeric=True, is_total=is_total)
            format_cell_value(ws.cell(row=curr_row, column=4), row["so_nguoi_co_thue"], is_numeric=True, is_total=is_total)
            format_cell_value(ws.cell(row=curr_row, column=5), row["thue_luy_tien"], is_numeric=True, is_total=is_total)
            format_cell_value(ws.cell(row=curr_row, column=6), row["thue_10_percent"], is_numeric=True, is_total=is_total)
            format_cell_value(ws.cell(row=curr_row, column=7), row["tong_thue"], is_numeric=True, is_total=is_total)

        self._auto_fit_columns(ws, max(len(headers_sum), len(headers_dept)))
