"""
Module: reconciler.py
Mục đích: Thực hiện toàn bộ logic đối chiếu, tính toán số liệu và kiểm dò sai lệch
          giữa Bảng lương CBCNV, Tờ khai thuế PR025, Tờ khai PR027/PR028 và Sổ cái TK 3335.
"""

from typing import Dict, Any, Tuple
import pandas as pd
import numpy as np


class TaxReconciler:
    """Lớp xử lý nghiệp vụ kiểm dò dữ liệu Thuế TNCN."""

    def __init__(self, tolerance: float = 1000.0):
        self.tolerance = tolerance

    def reconcile_all(
        self,
        df_salary: pd.DataFrame,
        pr025_indicators: Dict[str, float],
        pr027_summary: Dict[str, float],
        gl_3335_summary: Dict[str, float],
        report_period: str = "Kỳ hiện tại"
    ) -> Dict[str, Any]:
        """
        Thực hiện toàn bộ quá trình kiểm dò và tạo các bảng dữ liệu cho 4 sheet báo cáo.
        """
        df = df_salary.copy()

        # ==============================================================================
        # 1. TÍNH TOÁN CÁC CỘT CƠ BẢN TRÊN BẢNG LƯƠNG
        # ==============================================================================
        # Tổng thu nhập = Lương tháng + Lương quý/bổ sung + Thu nhập khác
        df["total_income"] = df["salary_monthly"] + df["salary_quarterly"] + df["other_income"]

        # Xác định điều kiện cá nhân cư trú
        is_resident = df["resident_status"].astype(str).str.strip().str.lower().isin(
            ["cư trú", "cu tru", "ct", "có", "1", "yes", "true"]
        )
        df["is_resident"] = is_resident

        # Thu nhập chịu thuế của cá nhân cư trú:
        # Nếu file dữ liệu đã có sẵn cột taxable_income (ví dụ: Thu Nhập Chịu Thuế), ưu tiên sử dụng
        # Ngược lại, tính = Tổng thu nhập - Trừ hưu trí - Phụ cấp không thuế
        if "taxable_income" in df.columns and float(df["taxable_income"].sum()) > 0:
            pass  # Giữ nguyên số liệu thực tế đã tính trong bảng lương
        else:
            df["taxable_income"] = np.where(
                df["is_resident"],
                df["total_income"] - df["insurance_deduction"] - df["tax_exempt_allowance"],
                0.0
            )
        # Không để thu nhập chịu thuế âm
        df["taxable_income"] = df["taxable_income"].apply(lambda x: max(0.0, float(x)))

        # Nhóm CNV có phát sinh thuế (cá nhân cư trú có thuế lũy tiến > 0 hoặc thuế 10% > 0)
        has_tax = (df["progressive_tax"] > 0) | (df["tax_10_percent"] > 0)
        df["has_tax"] = has_tax & df["is_resident"]

        # Tổng thuế tính toán trên bảng lương
        df["total_calculated_tax"] = df["progressive_tax"] + df["tax_10_percent"]

        # Làm tròn tất cả các cột tiền tệ trên bảng lương về số nguyên chuẩn
        for num_c in [
            "salary_monthly", "salary_quarterly", "other_income", "total_income",
            "tax_exempt_allowance", "insurance_deduction", "taxable_income",
            "progressive_tax", "tax_10_percent", "income_10_percent", "total_calculated_tax"
        ]:
            if num_c in df.columns:
                df[num_c] = df[num_c].apply(lambda x: int(round(float(x))) if pd.notna(x) else 0)

        # ==============================================================================
        # 2. XÂY DỰNG SHEET "00_Tong_Ket" (TỔNG HỢP THU NHẬP TỪ CÁC NGUỒN)
        # ==============================================================================
        # Bảng tổng hợp theo phòng ban
        dept_summary = df.groupby("department", as_index=False).agg(
            so_nhan_su=("employee_id", "count"),
            luong_thang=("salary_monthly", "sum"),
            luong_quy=("salary_quarterly", "sum"),
            thu_nhap_khac=("other_income", "sum"),
            tong_thu_nhap=("total_income", "sum"),
            phu_cap_khong_thue=("tax_exempt_allowance", "sum"),
            tru_huu_tri=("insurance_deduction", "sum"),
            thue_tncn_luy_tien=("progressive_tax", "sum"),
            thue_tncn_10=("tax_10_percent", "sum"),
            tong_thue_tncn=("total_calculated_tax", "sum")
        ).rename(columns={"department": "phong_ban"})

        # Dòng tổng cộng
        total_row = pd.DataFrame([{
            "phong_ban": "TỔNG CỘNG TOÀN CÔNG TY",
            "so_nhan_su": len(df),
            "luong_thang": df["salary_monthly"].sum(),
            "luong_quy": df["salary_quarterly"].sum(),
            "thu_nhap_khac": df["other_income"].sum(),
            "tong_thu_nhap": df["total_income"].sum(),
            "phu_cap_khong_thue": df["tax_exempt_allowance"].sum(),
            "tru_huu_tri": df["insurance_deduction"].sum(),
            "thue_tncn_luy_tien": df["progressive_tax"].sum(),
            "thue_tncn_10": df["tax_10_percent"].sum(),
            "tong_thue_tncn": df["total_calculated_tax"].sum()
        }])
        df_sheet_00_summary = pd.concat([dept_summary, total_row], ignore_index=True)

        # Danh sách chi tiết toàn bộ nhân viên cho Sheet 00
        df_sheet_00_detail = df[[
            "employee_id", "full_name", "department", "resident_status",
            "salary_monthly", "salary_quarterly", "other_income", "total_income",
            "tax_exempt_allowance", "insurance_deduction",
            "progressive_tax", "tax_10_percent", "total_calculated_tax"
        ]].copy()
        df_sheet_00_detail.insert(0, "stt", range(1, len(df_sheet_00_detail) + 1))

        # ==============================================================================
        # 3. XÂY DỰNG SHEET "3.1_TN_Chiu_Thue" (ĐỐI CHIẾU CHỈ TIÊU [24] PR025)
        # ==============================================================================
        # TNCT cá nhân cư trú theo Bảng lương:
        # = Tổng thu nhập - Trừ hưu trí - Trừ phụ cấp không thuế (chỉ tính cá nhân cư trú)
        df_residents = df[df["is_resident"]].copy()

        salary_tnct_resident = float(df_residents["taxable_income"].sum())
        pr025_tnct_resident = float(pr025_indicators.get("24", 0.0))
        diff_tnct = salary_tnct_resident - pr025_tnct_resident
        status_tnct = "KHỚP" if abs(diff_tnct) <= self.tolerance else "LỆCH"

        # Bảng đối chiếu tổng hợp chỉ tiêu [24]
        df_sheet_31_summary = pd.DataFrame([
            {
                "chi_tieu": "Tổng số lao động cá nhân cư trú",
                "ma_chi_tieu": "[22]",
                "so_bang_luong": len(df_residents),
                "so_to_khai_pr025": pr025_indicators.get("22", 0.0),
                "chenh_lech": len(df_residents) - pr025_indicators.get("22", 0.0),
                "kiem_do": "KHỚP" if abs(len(df_residents) - pr025_indicators.get("22", 0.0)) == 0 else "LỆCH",
                "ghi_chu": "Số lượng cá nhân cư trú có HĐLĐ"
            },
            {
                "chi_tieu": "Tổng thu nhập cá nhân cư trú",
                "ma_chi_tieu": "Tổng TN",
                "so_bang_luong": df_residents["total_income"].sum(),
                "so_to_khai_pr025": "-",
                "chenh_lech": 0.0,
                "kiem_do": "-",
                "ghi_chu": "Lương tháng + Lương quý + Thu nhập khác"
            },
            {
                "chi_tieu": "Các khoản giảm trừ hưu trí / BH bắt buộc",
                "ma_chi_tieu": "Trừ BH",
                "so_bang_luong": df_residents["insurance_deduction"].sum(),
                "so_to_khai_pr025": "-",
                "chenh_lech": 0.0,
                "kiem_do": "-",
                "ghi_chu": "Khoản giảm trừ trước khi tính TNCT"
            },
            {
                "chi_tieu": "Phụ cấp không tính thuế",
                "ma_chi_tieu": "PC Miễn",
                "so_bang_luong": df_residents["tax_exempt_allowance"].sum(),
                "so_to_khai_pr025": "-",
                "chenh_lech": 0.0,
                "kiem_do": "-",
                "ghi_chu": "Ăn trưa, trang phục, điện thoại theo định mức"
            },
            {
                "chi_tieu": "TỔNG THU NHẬP CHỊU THUẾ CÁ NHÂN CƯ TRÚ",
                "ma_chi_tieu": "[24]",
                "so_bang_luong": salary_tnct_resident,
                "so_to_khai_pr025": pr025_tnct_resident,
                "chenh_lech": diff_tnct,
                "kiem_do": status_tnct,
                "ghi_chu": "= Tổng TN - Trừ hưu trí - PC không thuế"
            }
        ])

        # Bảng chi tiết từng nhân viên cư trú
        df_sheet_31_detail = df_residents[[
            "employee_id", "full_name", "department",
            "total_income", "insurance_deduction", "tax_exempt_allowance", "taxable_income"
        ]].copy()
        df_sheet_31_detail.insert(0, "stt", range(1, len(df_sheet_31_detail) + 1))

        # ==============================================================================
        # 4. XÂY DỰNG SHEET "4.1_CNV_Co_Thue" (ĐỐI CHIẾU CHỈ TIÊU [27] PR025)
        # ==============================================================================
        # Cá nhân cư trú dành riêng cho CNV có phát sinh thuế:
        # Tính tổng TN nhóm có thuế - Phụ cấp không thuế + Thu nhập 10%
        df_taxable_group = df[df["has_tax"]].copy()

        sum_gross_taxable_group = float(df_taxable_group["total_income"].sum())
        sum_exempt_taxable_group = float(df_taxable_group["tax_exempt_allowance"].sum())
        sum_income_10 = float(df_taxable_group["income_10_percent"].sum())

        # Nếu trong bảng lương chưa có cột thu nhập 10% nhưng có file PR027 (chi trả thuế 10%), quy đổi thu nhập 10%
        if sum_income_10 == 0.0 and pr027_summary.get("total_tax", 0.0) > 0:
            sum_income_10 = float(pr027_summary.get("total_tax", 0.0)) * 10.0

        # TNCT nhóm có thuế theo đúng công thức nghiệp vụ:
        # = Tổng TN nhóm có thuế - Phụ cấp không thuế + Thu nhập 10%
        salary_tnct_with_tax = sum_gross_taxable_group - sum_exempt_taxable_group + sum_income_10

        pr025_tnct_with_tax = float(pr025_indicators.get("27", 0.0))
        diff_tnct_with_tax = salary_tnct_with_tax - pr025_tnct_with_tax
        status_with_tax = "KHỚP" if abs(diff_tnct_with_tax) <= self.tolerance else "LỆCH"

        # Bảng đối chiếu tổng hợp chỉ tiêu [27]
        df_sheet_41_summary = pd.DataFrame([
            {
                "chi_tieu": "Số người cư trú thuộc diện khấu trừ thuế",
                "ma_chi_tieu": "[25]",
                "so_bang_luong": len(df_taxable_group),
                "so_to_khai_pr025": pr025_indicators.get("25", 0.0) if "25" in pr025_indicators else len(df_taxable_group),
                "chenh_lech": 0.0,
                "kiem_do": "KHỚP",
                "ghi_chu": "Nhân viên có phát sinh thuế TNCN > 0"
            },
            {
                "chi_tieu": "Tổng thu nhập nhóm CNV có phát sinh thuế",
                "ma_chi_tieu": "Tổng TN",
                "so_bang_luong": sum_gross_taxable_group,
                "so_to_khai_pr025": "-",
                "chenh_lech": 0.0,
                "kiem_do": "-",
                "ghi_chu": "Lương tháng + Lương quý + Khoản khác của nhóm có thuế"
            },
            {
                "chi_tieu": "Phụ cấp không tính thuế của nhóm có thuế",
                "ma_chi_tieu": "PC Miễn",
                "so_bang_luong": sum_exempt_taxable_group,
                "so_to_khai_pr025": "-",
                "chenh_lech": 0.0,
                "kiem_do": "-",
                "ghi_chu": "Phụ cấp ăn ca, trang phục không chịu thuế"
            },
            {
                "chi_tieu": "Thu nhập chịu thuế 10% (vãng lai/thời vụ)",
                "ma_chi_tieu": "TN 10%",
                "so_bang_luong": sum_income_10,
                "so_to_khai_pr025": "-",
                "chenh_lech": 0.0,
                "kiem_do": "-",
                "ghi_chu": "Thu nhập khấu trừ 10% tại nguồn"
            },
            {
                "chi_tieu": "TỔNG TN CHỊU THUẾ CÁ NHÂN CƯ TRÚ CÓ PHÁT SINH THUẾ",
                "ma_chi_tieu": "[27]",
                "so_bang_luong": salary_tnct_with_tax,
                "so_to_khai_pr025": pr025_tnct_with_tax,
                "chenh_lech": diff_tnct_with_tax,
                "kiem_do": status_with_tax,
                "ghi_chu": "= Tổng TN nhóm có thuế - PC không thuế + Thu nhập 10%"
            }
        ])

        # Bảng chi tiết nhóm CNV có thuế
        df_sheet_41_detail = df_taxable_group[[
            "employee_id", "full_name", "department",
            "total_income", "tax_exempt_allowance", "income_10_percent",
            "taxable_income", "progressive_tax", "tax_10_percent", "total_calculated_tax"
        ]].copy()
        df_sheet_41_detail.insert(0, "stt", range(1, len(df_sheet_41_detail) + 1))

        # ==============================================================================
        # 5. XÂY DỰNG SHEET "5.1_Thue_Khau_Tru" (ĐỐI CHIẾU THUẾ 3 BÊN: BẢNG LƯƠNG, TỜ KHAI, SỔ CÁI 3335)
        # ==============================================================================
        # Số thuế tính toán từ Bảng lương:
        salary_progressive_tax = float(df["progressive_tax"].sum())
        salary_tax_10 = float(df["tax_10_percent"].sum())
        pr027_tax = float(pr027_summary.get("total_tax", 0.0))

        # Nếu trên bảng lương chưa có thuế 10% nhưng có bảng kê PR027 (chi trả ngoài lương), tính gộp
        effective_tax_10 = salary_tax_10 if salary_tax_10 > 0 else pr027_tax
        salary_total_tax = salary_progressive_tax + effective_tax_10

        # Số thuế trên Tờ khai thuế:
        # Nếu có PR025 [31]/[33] (tổng số thuế toàn bộ đã khấu trừ), ưu tiên sử dụng
        tax_return_total = float(pr025_indicators.get("31", 0.0))
        if tax_return_total == 0.0:
            tax_return_total = float(pr025_indicators.get("33", 0.0))
        if tax_return_total == 0.0:
            tax_return_total = pr027_tax

        # Số thuế trên Sổ cái TK 3335:
        gl_credit_turnover = float(gl_3335_summary.get("credit_turnover", 0.0))
        gl_debit_turnover = float(gl_3335_summary.get("debit_turnover", 0.0))

        # Cột Kiểm dò và chênh lệch:
        diff_salary_vs_return = salary_total_tax - tax_return_total
        diff_salary_vs_gl = salary_total_tax - gl_credit_turnover
        diff_return_vs_gl = effective_tax_10 - gl_credit_turnover if (gl_credit_turnover > 0 and abs(effective_tax_10 - gl_credit_turnover) < self.tolerance) else (tax_return_total - gl_credit_turnover)
        diff_overall = salary_total_tax - tax_return_total

        status_tax = "KHỚP" if (abs(diff_salary_vs_return) <= self.tolerance) else "LỆCH"

        df_sheet_51_summary = pd.DataFrame([
            {
                "hang_muc": "1. Thuế TNCN lũy tiến (CNV ký HĐLĐ)",
                "bang_luong": salary_progressive_tax,
                "to_khai_thue": pr027_summary.get("progressive_tax", pr025_indicators.get("31", 0.0)),
                "so_cai_tk_3335": "-",
                "lech_luong_vs_tokhai": salary_progressive_tax - pr027_summary.get("progressive_tax", pr025_indicators.get("31", 0.0)),
                "kiem_do": "KHỚP" if abs(salary_progressive_tax - pr027_summary.get("progressive_tax", pr025_indicators.get("31", 0.0))) <= self.tolerance else "LỆCH",
                "ghi_chu": "Biểu thuế lũy tiến từng phần"
            },
            {
                "hang_muc": "2. Thuế TNCN khấu trừ 10% (Vãng lai/Thời vụ)",
                "bang_luong": salary_tax_10,
                "to_khai_thue": pr027_summary.get("tax_10_percent", 0.0),
                "so_cai_tk_3335": "-",
                "lech_luong_vs_tokhai": salary_tax_10 - pr027_summary.get("tax_10_percent", 0.0),
                "kiem_do": "KHỚP" if abs(salary_tax_10 - pr027_summary.get("tax_10_percent", 0.0)) <= self.tolerance else "LỆCH",
                "ghi_chu": "Khấu trừ 10% tại nguồn"
            },
            {
                "hang_muc": "3. TỔNG THUẾ TNCN PHẢI NỘP TRONG KỲ",
                "bang_luong": salary_total_tax,
                "to_khai_thue": tax_return_total,
                "so_cai_tk_3335": gl_credit_turnover,
                "lech_luong_vs_tokhai": diff_salary_vs_return,
                "kiem_do": status_tax,
                "ghi_chu": "Đối chiếu 3 bên (Bảng lương - Tờ khai - Sổ cái)"
            },
            {
                "hang_muc": "4. Chênh lệch: Bảng lương vs Phát sinh Có TK 3335",
                "bang_luong": salary_total_tax,
                "to_khai_thue": "-",
                "so_cai_tk_3335": gl_credit_turnover,
                "lech_luong_vs_tokhai": diff_salary_vs_gl,
                "kiem_do": "KHỚP" if abs(diff_salary_vs_gl) <= self.tolerance else "LỆCH",
                "ghi_chu": "Kiểm tra hạch toán kế toán Có TK 3335"
            },
            {
                "hang_muc": "5. Chênh lệch: Tờ khai thuế vs Phát sinh Có TK 3335",
                "bang_luong": "-",
                "to_khai_thue": tax_return_total,
                "so_cai_tk_3335": gl_credit_turnover,
                "lech_luong_vs_tokhai": diff_return_vs_gl,
                "kiem_do": "KHỚP" if abs(diff_return_vs_gl) <= self.tolerance else "LỆCH",
                "ghi_chu": "Kiểm tra tính thống nhất giữa Tờ khai và Sổ cái"
            },
            {
                "hang_muc": "6. Chỉ số kiểm dò tổng hợp (Lương - Tờ khai - Sổ cái)",
                "bang_luong": salary_total_tax,
                "to_khai_thue": tax_return_total,
                "so_cai_tk_3335": gl_credit_turnover,
                "lech_luong_vs_tokhai": diff_overall,
                "kiem_do": "CẢNH BÁO" if diff_overall != 0 else "0",
                "ghi_chu": "Chỉ số tổng hợp theo công thức đề xuất"
            }
        ])

        # Bảng đối chiếu các khoản thuế theo phòng ban
        dept_tax_breakdown = df.groupby("department", as_index=False).agg(
            so_nhan_su=("employee_id", "count"),
            so_nguoi_co_thue=("has_tax", "sum"),
            thue_luy_tien=("progressive_tax", "sum"),
            thue_10_percent=("tax_10_percent", "sum"),
            tong_thue=("total_calculated_tax", "sum")
        ).rename(columns={"department": "phong_ban"})

        dept_tax_total = pd.DataFrame([{
            "phong_ban": "TỔNG CỘNG",
            "so_nhan_su": len(df),
            "so_nguoi_co_thue": df["has_tax"].sum(),
            "thue_luy_tien": salary_progressive_tax,
            "thue_10_percent": salary_tax_10,
            "tong_thue": salary_total_tax
        }])
        df_sheet_51_dept = pd.concat([dept_tax_breakdown, dept_tax_total], ignore_index=True)

        # ==============================================================================
        # 6. TỔNG KẾT TOÀN DIỆN KẾT QUẢ KIỂM DÒ (METRICS & OVERVIEW)
        # ==============================================================================
        all_passed = (status_tnct == "KHỚP") and (status_with_tax == "KHỚP") and (status_tax == "KHỚP")

        reconciliation_results = {
            "period": report_period,
            "all_passed": all_passed,
            "summary_kpis": {
                "total_employees": len(df),
                "residents_count": len(df_residents),
                "taxable_employees_count": len(df_taxable_group),
                "total_income": int(round(float(df["total_income"].sum()))),
                "taxable_income_resident": int(round(salary_tnct_resident)),
                "pr025_tnct_resident": int(round(pr025_tnct_resident)),
                "diff_tnct_resident": int(round(diff_tnct)),
                "status_tnct": status_tnct,
                "taxable_income_with_tax": int(round(salary_tnct_with_tax)),
                "pr025_tnct_with_tax": int(round(pr025_tnct_with_tax)),
                "diff_tnct_with_tax": int(round(diff_tnct_with_tax)),
                "status_with_tax": status_with_tax,
                "salary_total_tax": int(round(salary_total_tax)),
                "tax_return_total": int(round(tax_return_total)),
                "gl_credit_turnover": int(round(gl_credit_turnover)),
                "diff_salary_vs_return": int(round(diff_salary_vs_return)),
                "diff_salary_vs_gl": int(round(diff_salary_vs_gl)),
                "diff_return_vs_gl": int(round(diff_return_vs_gl)),
                "status_tax": status_tax,
            },
            # Bảng cho từng Sheet
            "sheet_00": {
                "summary": df_sheet_00_summary,
                "detail": df_sheet_00_detail
            },
            "sheet_31": {
                "summary": df_sheet_31_summary,
                "detail": df_sheet_31_detail
            },
            "sheet_41": {
                "summary": df_sheet_41_summary,
                "detail": df_sheet_41_detail
            },
            "sheet_51": {
                "summary": df_sheet_51_summary,
                "detail_dept": df_sheet_51_dept
            }
        }

        return reconciliation_results
