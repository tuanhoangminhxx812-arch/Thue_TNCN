"""
Module: annual_manager.py
Mục đích: Quản lý, đóng gói và lưu trữ dữ liệu Thuế TNCN cả năm (12 tháng).
          - Cho phép nạp và lưu trữ từng tháng đã tổng hợp thành các gói dữ liệu (Precomputed Cache).
          - Gọi lại dữ liệu các tháng đã đóng gói siêu nhanh (< 0.05s) không bị giật lag,
            không cần phải đọc và tổng hợp lại 4 file Excel mỗi lần chuyển tháng.
          - Chỉ khi bấm nút "Cập nhật & Tổng hợp lại" thì mới nạp lại từ 4 file gốc.
          - Xây dựng Bảng và Sheet Tổng Hợp Cả Năm (12 tháng + Tổng cộng cả năm) riêng biệt
            để báo cáo Lãnh đạo.
"""

import os
import json
import pickle
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List, Optional
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from scripts.report_builder import (
    COLOR_HEADER_NAVY, COLOR_HEADER_SUB, COLOR_WHITE,
    COLOR_MATCH_BG, COLOR_MATCH_FG, COLOR_DIFF_BG, COLOR_DIFF_FG,
    COLOR_TOTAL_BG, THIN_BORDER_SIDE, MEDIUM_BORDER_SIDE, DOUBLE_BOTTOM_SIDE,
    BORDER_ALL_THIN, BORDER_TOTAL, FORMAT_CURRENCY, format_cell_value
)


class AnnualTaxManager:
    """Quản lý đóng gói dữ liệu thuế TNCN nhiều kỳ trong năm và xuất báo cáo tổng hợp năm."""

    def __init__(self, store_dir: str = "data_store"):
        self.store_dir = Path(store_dir)
        self.store_dir.mkdir(parents=True, exist_ok=True)
        self.monthly_dir = self.store_dir / "monthly"
        self.monthly_dir.mkdir(parents=True, exist_ok=True)
        self.catalog_file = self.store_dir / "catalog.json"
        self._init_catalog()

    def _init_catalog(self):
        """Khởi tạo file chỉ mục nếu chưa tồn tại."""
        if not self.catalog_file.exists():
            with open(self.catalog_file, "w", encoding="utf-8") as f:
                json.dump({}, f, ensure_ascii=False, indent=2)

    def _load_catalog(self) -> Dict[str, Any]:
        """Đọc danh mục các kỳ đã đóng gói."""
        try:
            with open(self.catalog_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _save_catalog(self, catalog: Dict[str, Any]):
        """Lưu danh mục các kỳ vào catalog.json."""
        with open(self.catalog_file, "w", encoding="utf-8") as f:
            json.dump(catalog, f, ensure_ascii=False, indent=2)

    def normalize_period_key(self, period_str: str) -> str:
        """
        Chuẩn hóa tên kỳ thành mã khóa duy nhất, ví dụ:
        'Tháng 09/2026' -> 'Thang_09_2026'
        'Tháng 1' -> 'Thang_01'
        """
        import re
        s = period_str.strip()
        # Tìm số tháng và số năm nếu có
        m_month = re.search(r'(?:tháng|thang|t|m)\s*0?(\d+)', s, re.IGNORECASE)
        m_year = re.search(r'(?:năm|nam|y|\/|-)\s*(\d{4})', s, re.IGNORECASE)

        month_num = int(m_month.group(1)) if m_month else 9
        year_num = int(m_year.group(1)) if m_year else datetime.now().year

        return f"Thang_{month_num:02d}_{year_num}"

    def is_period_stored(self, period_key: str) -> bool:
        """Kiểm tra kỳ này đã được đóng gói và lưu trữ chưa."""
        catalog = self._load_catalog()
        pkl_path = self.monthly_dir / f"{period_key}.pkl"
        return (period_key in catalog) and pkl_path.exists()

    def save_package(
        self,
        period_name: str,
        results_data: Dict[str, Any],
        source_folder: str = ""
    ) -> str:
        """
        Đóng gói toàn bộ kết quả kiểm dò của 1 kỳ vào kho lưu trữ (Cache).
        """
        period_key = self.normalize_period_key(period_name)
        pkl_path = self.monthly_dir / f"{period_key}.pkl"

        # 1. Lưu toàn bộ kết quả chi tiết (bao gồm DataFrames) vào file nhị phân pickle
        with open(pkl_path, "wb") as f:
            pickle.dump(results_data, f)

        # 2. Cập nhật vào catalog tóm tắt (để load danh sách và bảng cả năm siêu nhanh)
        catalog = self._load_catalog()
        kpis = results_data.get("summary_kpis", {})

        catalog[period_key] = {
            "period_name": period_name,
            "period_key": period_key,
            "saved_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "source_folder": source_folder,
            "all_passed": results_data.get("all_passed", False),
            "kpis": {
                "total_employees": int(kpis.get("total_employees", 0)),
                "residents_count": int(kpis.get("residents_count", 0)),
                "taxable_employees_count": int(kpis.get("taxable_employees_count", 0)),
                "total_income": int(kpis.get("total_income", 0)),
                "taxable_income_resident": int(kpis.get("taxable_income_resident", 0)),
                "pr025_tnct_resident": int(kpis.get("pr025_tnct_resident", 0)),
                "diff_tnct_resident": int(kpis.get("diff_tnct_resident", 0)),
                "status_tnct": str(kpis.get("status_tnct", "-")),
                "taxable_income_with_tax": int(kpis.get("taxable_income_with_tax", 0)),
                "pr025_tnct_with_tax": int(kpis.get("pr025_tnct_with_tax", 0)),
                "diff_tnct_with_tax": int(kpis.get("diff_tnct_with_tax", 0)),
                "status_with_tax": str(kpis.get("status_with_tax", "-")),
                "salary_total_tax": int(kpis.get("salary_total_tax", 0)),
                "tax_return_total": int(kpis.get("tax_return_total", 0)),
                "gl_credit_turnover": int(kpis.get("gl_credit_turnover", 0)),
                "diff_salary_vs_return": int(kpis.get("diff_salary_vs_return", 0)),
                "status_tax": str(kpis.get("status_tax", "-"))
            }
        }
        self._save_catalog(catalog)
        return period_key

    def load_package(self, period_key: str) -> Optional[Dict[str, Any]]:
        """
        Nạp lại kết quả đã đóng gói của 1 kỳ trong mili-giây mà không cần đọc lại 4 file Excel.
        """
        pkl_path = self.monthly_dir / f"{period_key}.pkl"
        if not pkl_path.exists():
            return None
        try:
            with open(pkl_path, "rb") as f:
                return pickle.load(f)
        except Exception:
            return None

    def get_stored_periods_list(self) -> List[Dict[str, Any]]:
        """Lấy danh sách các kỳ đã đóng gói kèm thông tin trạng thái."""
        catalog = self._load_catalog()
        # Sắp xếp theo khóa thời gian
        sorted_keys = sorted(catalog.keys())
        return [catalog[k] for k in sorted_keys]

    def build_annual_summary_df(self, year: int = None) -> pd.DataFrame:
        """
        Tổng hợp toàn bộ 12 tháng của năm thành một DataFrame duy nhất để báo cáo lãnh đạo.
        """
        catalog = self._load_catalog()
        rows = []

        # Lọc các kỳ theo năm hoặc lấy toàn bộ các kỳ đã lưu
        sorted_keys = sorted(catalog.keys())
        for k in sorted_keys:
            entry = catalog[k]
            kpis = entry.get("kpis", {})

            row = {
                "ky_bao_cao": entry.get("period_name", k),
                "so_lao_dong": kpis.get("total_employees", 0),
                "so_co_thue": kpis.get("taxable_employees_count", 0),
                "tong_thu_nhap": kpis.get("total_income", 0),
                "tnct_cu_tru": kpis.get("taxable_income_resident", 0),
                "pr025_tnct_cu_tru": kpis.get("pr025_tnct_resident", 0),
                "lech_tnct_cu_tru": kpis.get("diff_tnct_resident", 0),
                "tnct_nhom_co_thue": kpis.get("taxable_income_with_tax", 0),
                "pr025_tnct_nhom_co_thue": kpis.get("pr025_tnct_with_tax", 0),
                "lech_tnct_co_thue": kpis.get("diff_tnct_with_tax", 0),
                "thue_bang_luong": kpis.get("salary_total_tax", 0),
                "thue_to_khai": kpis.get("tax_return_total", 0),
                "thue_so_cai_3335": kpis.get("gl_credit_turnover", 0),
                "lech_thue": kpis.get("diff_salary_vs_return", 0),
                "kiem_do": "KHỚP" if entry.get("all_passed", False) else "LỆCH",
                "ngay_cap_nhat": entry.get("saved_at", "")
            }
            rows.append(row)

        if not rows:
            return pd.DataFrame()

        df_annual = pd.DataFrame(rows)

        # Thêm dòng TỔNG CỘNG CẢ NĂM
        total_row = {
            "ky_bao_cao": "TỔNG CỘNG CẢ NĂM",
            "so_lao_dong": int(round(df_annual["so_lao_dong"].mean())) if len(df_annual) > 0 else 0, # Số LĐ bình quân
            "so_co_thue": int(round(df_annual["so_co_thue"].mean())) if len(df_annual) > 0 else 0,
            "tong_thu_nhap": int(df_annual["tong_thu_nhap"].sum()),
            "tnct_cu_tru": int(df_annual["tnct_cu_tru"].sum()),
            "pr025_tnct_cu_tru": int(df_annual["pr025_tnct_cu_tru"].sum()),
            "lech_tnct_cu_tru": int(df_annual["lech_tnct_cu_tru"].sum()),
            "tnct_nhom_co_thue": int(df_annual["tnct_nhom_co_thue"].sum()),
            "pr025_tnct_nhom_co_thue": int(df_annual["pr025_tnct_nhom_co_thue"].sum()),
            "lech_tnct_co_thue": int(df_annual["lech_tnct_co_thue"].sum()),
            "thue_bang_luong": int(df_annual["thue_bang_luong"].sum()),
            "thue_to_khai": int(df_annual["thue_to_khai"].sum()),
            "thue_so_cai_3335": int(df_annual["thue_so_cai_3335"].sum()),
            "lech_thue": int(df_annual["lech_thue"].sum()),
            "kiem_do": "KHỚP" if (df_annual["kiem_do"] == "KHỚP").all() else "LỆCH",
            "ngay_cap_nhat": datetime.now().strftime("%d/%m/%Y")
        }

        df_final = pd.concat([df_annual, pd.DataFrame([total_row])], ignore_index=True)
        return df_final

    def export_annual_report_excel(
        self,
        company_name: str,
        year_label: str = "Năm 2026",
        output_path: str = "03_DAU_RA/Bao_Cao_Tong_Hop_Thue_TNCN_Ca_Nam.xlsx"
    ) -> str:
        """
        Xuất file Excel báo cáo tổng hợp cả năm dành riêng cho Lãnh đạo với định dạng chuẩn tài chính.
        """
        df_annual = self.build_annual_summary_df()
        out_p = Path(output_path)
        out_p.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Bao_Cao_Ca_Nam"
        ws.views.sheetView[0].showGridLines = True

        # 1. Khối Tiêu Đề Báo Cáo Lãnh Đạo
        ws.cell(row=1, column=1, value=company_name).font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
        ws.cell(row=2, column=1, value=f"BÁO CÁO TỔNG HỢP & ĐỐI CHIẾU THUẾ THU NHẬP CÁ NHÂN {year_label.upper()}").font = Font(name="Calibri", size=14, bold=True, color=COLOR_HEADER_NAVY)
        ws.cell(row=3, column=1, value=f"Dành cho Ban Lãnh Đạo | Ngày lập: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}").font = Font(name="Calibri", size=10, italic=True, color="555555")

        # 2. Dòng Tiêu Đề Bảng
        headers = [
            ("STT", 6, True),
            ("Kỳ / Tháng", 16, False),
            ("Số Lao Động\n(Bình quân)", 14, True),
            ("Số Người\nCó Thuế", 14, True),
            ("Tổng Thu Nhập\n(Gross VND)", 20, False),
            ("TNCT Cư Trú\nBảng Lương [24]", 20, False),
            ("TNCT Cư Trú\nTờ Khai PR025", 20, False),
            ("Chênh Lệch\nTNCT [24]", 18, False),
            ("TNCT Nhóm Có Thuế\nBảng Lương [27]", 20, False),
            ("TNCT Nhóm Có Thuế\nTờ Khai PR025", 20, False),
            ("Chênh Lệch\nTNCT [27]", 18, False),
            ("Tổng Thuế\nBảng Lương", 18, False),
            ("Tổng Thuế\nTờ Khai", 18, False),
            ("Thuế Hạch Toán\nCó TK 3335", 18, False),
            ("Chênh Lệch\nSố Thuế", 18, False),
            ("Kiểm Dò\nTổng Thể", 14, True),
            ("Ngày Cập Nhật", 16, True)
        ]

        header_row = 5
        ws.row_dimensions[header_row].height = 34

        for col_idx, (h_name, width, is_center) in enumerate(headers, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=h_name)
            cell.font = Font(name="Calibri", size=10, bold=True, color=COLOR_WHITE)
            cell.fill = PatternFill(start_color=COLOR_HEADER_NAVY, end_color=COLOR_HEADER_NAVY, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = BORDER_ALL_THIN
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        # 3. Ghi Dữ Liệu Các Tháng
        current_row = 6
        num_data_rows = len(df_annual)

        for idx, row in df_annual.iterrows():
            is_total_row = (idx == num_data_rows - 1)
            ws.row_dimensions[current_row].height = 22

            # STT
            stt_val = "" if is_total_row else str(idx + 1)
            format_cell_value(ws.cell(row=current_row, column=1), stt_val, is_center=True, is_total=is_total_row)

            # Kỳ báo cáo
            format_cell_value(ws.cell(row=current_row, column=2), row["ky_bao_cao"], is_bold=is_total_row, is_total=is_total_row)

            # Nhân sự
            format_cell_value(ws.cell(row=current_row, column=3), row["so_lao_dong"], is_numeric=True, is_total=is_total_row)
            format_cell_value(ws.cell(row=current_row, column=4), row["so_co_thue"], is_numeric=True, is_total=is_total_row)

            # Tổng thu nhập
            format_cell_value(ws.cell(row=current_row, column=5), row["tong_thu_nhap"], is_numeric=True, is_bold=True, is_total=is_total_row)

            # TNCT [24]
            format_cell_value(ws.cell(row=current_row, column=6), row["tnct_cu_tru"], is_numeric=True, is_total=is_total_row)
            format_cell_value(ws.cell(row=current_row, column=7), row["pr025_tnct_cu_tru"], is_numeric=True, is_total=is_total_row)
            cell_diff_24 = ws.cell(row=current_row, column=8)
            format_cell_value(cell_diff_24, row["lech_tnct_cu_tru"], is_numeric=True, is_bold=True, is_total=is_total_row)
            if not is_total_row:
                if abs(float(row["lech_tnct_cu_tru"])) <= 1000:
                    cell_diff_24.fill = PatternFill(start_color=COLOR_MATCH_BG, end_color=COLOR_MATCH_BG, fill_type="solid")
                    cell_diff_24.font = Font(name="Calibri", size=10, bold=True, color=COLOR_MATCH_FG)
                else:
                    cell_diff_24.fill = PatternFill(start_color=COLOR_DIFF_BG, end_color=COLOR_DIFF_BG, fill_type="solid")
                    cell_diff_24.font = Font(name="Calibri", size=10, bold=True, color=COLOR_DIFF_FG)

            # TNCT [27]
            format_cell_value(ws.cell(row=current_row, column=9), row["tnct_nhom_co_thue"], is_numeric=True, is_total=is_total_row)
            format_cell_value(ws.cell(row=current_row, column=10), row["pr025_tnct_nhom_co_thue"], is_numeric=True, is_total=is_total_row)
            cell_diff_27 = ws.cell(row=current_row, column=11)
            format_cell_value(cell_diff_27, row["lech_tnct_co_thue"], is_numeric=True, is_bold=True, is_total=is_total_row)
            if not is_total_row:
                if abs(float(row["lech_tnct_co_thue"])) <= 1000:
                    cell_diff_27.fill = PatternFill(start_color=COLOR_MATCH_BG, end_color=COLOR_MATCH_BG, fill_type="solid")
                    cell_diff_27.font = Font(name="Calibri", size=10, bold=True, color=COLOR_MATCH_FG)
                else:
                    cell_diff_27.fill = PatternFill(start_color=COLOR_DIFF_BG, end_color=COLOR_DIFF_BG, fill_type="solid")
                    cell_diff_27.font = Font(name="Calibri", size=10, bold=True, color=COLOR_DIFF_FG)

            # Số thuế 3 bên
            format_cell_value(ws.cell(row=current_row, column=12), row["thue_bang_luong"], is_numeric=True, is_bold=True, is_total=is_total_row)
            format_cell_value(ws.cell(row=current_row, column=13), row["thue_to_khai"], is_numeric=True, is_total=is_total_row)
            format_cell_value(ws.cell(row=current_row, column=14), row["thue_so_cai_3335"], is_numeric=True, is_total=is_total_row)

            cell_diff_tax = ws.cell(row=current_row, column=15)
            format_cell_value(cell_diff_tax, row["lech_thue"], is_numeric=True, is_bold=True, is_total=is_total_row)
            if not is_total_row:
                if abs(float(row["lech_thue"])) <= 1000:
                    cell_diff_tax.fill = PatternFill(start_color=COLOR_MATCH_BG, end_color=COLOR_MATCH_BG, fill_type="solid")
                    cell_diff_tax.font = Font(name="Calibri", size=10, bold=True, color=COLOR_MATCH_FG)
                else:
                    cell_diff_tax.fill = PatternFill(start_color=COLOR_DIFF_BG, end_color=COLOR_DIFF_BG, fill_type="solid")
                    cell_diff_tax.font = Font(name="Calibri", size=10, bold=True, color=COLOR_DIFF_FG)

            # Kiểm dò tổng thể
            cell_status = ws.cell(row=current_row, column=16)
            format_cell_value(cell_status, row["kiem_do"], is_center=True, is_bold=True, is_total=is_total_row)
            if row["kiem_do"] == "KHỚP":
                cell_status.fill = PatternFill(start_color=COLOR_MATCH_BG, end_color=COLOR_MATCH_BG, fill_type="solid")
                cell_status.font = Font(name="Calibri", size=10, bold=True, color=COLOR_MATCH_FG)
            else:
                cell_status.fill = PatternFill(start_color=COLOR_DIFF_BG, end_color=COLOR_DIFF_BG, fill_type="solid")
                cell_status.font = Font(name="Calibri", size=10, bold=True, color=COLOR_DIFF_FG)

            # Ngày cập nhật
            format_cell_value(ws.cell(row=current_row, column=17), row["ngay_cap_nhat"], is_center=True, is_total=is_total_row)

            current_row += 1

        # 4. Khối Ký Duyệt Báo Cáo Cuối Trang
        sign_row = current_row + 2
        ws.cell(row=sign_row, column=2, value="NGƯỜI LẬP BIỂU").font = Font(name="Calibri", size=11, bold=True)
        ws.cell(row=sign_row, column=2).alignment = Alignment(horizontal="center")

        ws.cell(row=sign_row, column=9, value="KẾ TOÁN TRƯỞNG").font = Font(name="Calibri", size=11, bold=True)
        ws.cell(row=sign_row, column=9).alignment = Alignment(horizontal="center")

        ws.cell(row=sign_row, column=15, value="GIÁM ĐỐC / LÃNH ĐẠO ĐƠN VỊ").font = Font(name="Calibri", size=11, bold=True)
        ws.cell(row=sign_row, column=15).alignment = Alignment(horizontal="center")

        sign_sub_row = sign_row + 1
        ws.cell(row=sign_sub_row, column=2, value="(Ký, ghi rõ họ tên)").font = Font(name="Calibri", size=9, italic=True)
        ws.cell(row=sign_sub_row, column=2).alignment = Alignment(horizontal="center")

        ws.cell(row=sign_sub_row, column=9, value="(Ký, ghi rõ họ tên)").font = Font(name="Calibri", size=9, italic=True)
        ws.cell(row=sign_sub_row, column=9).alignment = Alignment(horizontal="center")

        ws.cell(row=sign_sub_row, column=15, value="(Ký, đóng dấu)").font = Font(name="Calibri", size=9, italic=True)
        ws.cell(row=sign_sub_row, column=15).alignment = Alignment(horizontal="center")

        wb.save(str(out_p))
        return str(out_p)

    def populate_demo_months(self, base_results: Dict[str, Any], year: int = 2026):
        """
        Tạo nhanh dữ liệu mẫu các tháng còn lại trong năm (Tháng 01 -> Tháng 08)
        dựa trên số liệu thực tế để xem thử Báo cáo Cả Năm cho Lãnh đạo.
        Khi người dùng nạp dữ liệu thật của các tháng, dữ liệu thật sẽ tự động ghi đè.
        """
        import copy
        for m in range(1, 10):
            if m == 9:
                continue
            period_name = f"Tháng {m:02d}/{year}"
            factor = 1.0 + ((m - 9) * 0.005)
            mod_results = copy.deepcopy(base_results)
            kpis = mod_results.get("summary_kpis", {})

            for k in [
                "total_income", "taxable_income_resident", "pr025_tnct_resident",
                "taxable_income_with_tax", "pr025_tnct_with_tax",
                "salary_total_tax", "tax_return_total", "gl_credit_turnover"
            ]:
                if k in kpis:
                    kpis[k] = int(round(float(kpis[k]) * factor))

            # Các tháng trước khớp chuẩn
            kpis["diff_tnct_resident"] = 0
            kpis["status_tnct"] = "KHỚP"
            kpis["diff_tnct_with_tax"] = 0
            kpis["status_with_tax"] = "KHỚP"
            kpis["diff_salary_vs_return"] = 0
            kpis["status_tax"] = "KHỚP"
            mod_results["all_passed"] = True

            self.save_package(period_name, mod_results, source_folder="Mẫu tự động")
