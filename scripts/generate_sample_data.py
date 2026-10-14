"""
Script: generate_sample_data.py
Mục đích: Tự động tạo bộ 4 file Excel mẫu chuẩn nghiệp vụ kế toán Việt Nam vào thư mục 01_DAU_VAO:
          1. Bang_Luong_Tong_CBCNV.xlsx
          2. To_Khai_Thue_PR025.xlsx
          3. To_Khai_Thue_PR027.xlsx
          4. So_Cai_TK_3335.xlsx
"""

import sys
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Đảm bảo in tiếng Việt trên console Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def generate_all_samples(output_dir: str = "01_DAU_VAO"):
    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    # ==============================================================================
    # 1. BẢNG LƯƠNG TỔNG CBCNV (15 Nhân sự mẫu)
    # ==============================================================================
    employees_data = [
        {"Mã NV": "NV001", "Họ và tên": "Nguyễn Văn An", "Phòng ban": "Ban Giám Đốc", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 60000000, "Lương quý": 15000000, "Thu nhập khác": 5000000, "Phụ cấp không thuế": 730000,
         "Trừ hưu trí": 3500000, "Thuế TNCN lũy tiến": 12850000, "Thu nhập 10%": 0, "Thuế 10%": 0},

        {"Mã NV": "NV002", "Họ và tên": "Trần Thị Bích", "Phòng ban": "Kế Toán", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 25000000, "Lương quý": 6000000, "Thu nhập khác": 2000000, "Phụ cấp không thuế": 730000,
         "Trừ hưu trí": 2625000, "Thuế TNCN lũy tiến": 2350000, "Thu nhập 10%": 0, "Thuế 10%": 0},

        {"Mã NV": "NV003", "Họ và tên": "Lê Hoàng Cường", "Phòng ban": "Kỹ Thuật", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 35000000, "Lương quý": 8000000, "Thu nhập khác": 3000000, "Phụ cấp không thuế": 730000,
         "Trừ hưu trí": 3500000, "Thuế TNCN lũy tiến": 5420000, "Thu nhập 10%": 0, "Thuế 10%": 0},

        {"Mã NV": "NV004", "Họ và tên": "Phạm Minh Đức", "Phòng ban": "Kỹ Thuật", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 22000000, "Lương quý": 5000000, "Thu nhập khác": 1000000, "Phụ cấp không thuế": 730000,
         "Trừ hưu trí": 2310000, "Thuế TNCN lũy tiến": 1650000, "Thu nhập 10%": 0, "Thuế 10%": 0},

        {"Mã NV": "NV005", "Họ và tên": "Vũ Thùy Dung", "Phòng ban": "Kinh Doanh", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 20000000, "Lương quý": 12000000, "Thu nhập khác": 4000000, "Phụ cấp không thuế": 730000,
         "Trừ hưu trí": 2100000, "Thuế TNCN lũy tiến": 3850000, "Thu nhập 10%": 0, "Thuế 10%": 0},

        {"Mã NV": "NV006", "Họ và tên": "Đỗ Quang Hải", "Phòng ban": "Kinh Doanh", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 18000000, "Lương quý": 8000000, "Thu nhập khác": 2000000, "Phụ cấp không thuế": 730000,
         "Trừ hưu trí": 1890000, "Thuế TNCN lũy tiến": 1850000, "Thu nhập 10%": 0, "Thuế 10%": 0},

        {"Mã NV": "NV007", "Họ và tên": "Ngô Mai Hoa", "Phòng ban": "Nhân Sự", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 16000000, "Lương quý": 4000000, "Thu nhập khác": 1000000, "Phụ cấp không thuế": 730000,
         "Trừ hưu trí": 1680000, "Thuế TNCN lũy tiến": 750000, "Thu nhập 10%": 0, "Thuế 10%": 0},

        {"Mã NV": "NV008", "Họ và tên": "Bùi Văn Kiên", "Phòng ban": "Kỹ Thuật", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 15000000, "Lương quý": 3000000, "Thu nhập khác": 1000000, "Phụ cấp không thuế": 730000,
         "Trừ hưu trí": 1575000, "Thuế TNCN lũy tiến": 450000, "Thu nhập 10%": 0, "Thuế 10%": 0},

        {"Mã NV": "NV009", "Họ và tên": "Đinh Thị Lan", "Phòng ban": "Kế Toán", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 12000000, "Lương quý": 2000000, "Thu nhập khác": 500000, "Phụ cấp không thuế": 730000,
         "Trừ hưu trí": 1260000, "Thuế TNCN lũy tiến": 0, "Thu nhập 10%": 0, "Thuế 10%": 0},

        {"Mã NV": "NV010", "Họ và tên": "Hoàng Quốc Nam", "Phòng ban": "Hành Chính", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 11000000, "Lương quý": 1500000, "Thu nhập khác": 500000, "Phụ cấp không thuế": 730000,
         "Trừ hưu trí": 1155000, "Thuế TNCN lũy tiến": 0, "Thu nhập 10%": 0, "Thuế 10%": 0},

        {"Mã NV": "NV011", "Họ và tên": "Phan Thị Oanh", "Phòng ban": "Kinh Doanh", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 10000000, "Lương quý": 1000000, "Thu nhập khác": 500000, "Phụ cấp không thuế": 730000,
         "Trừ hưu trí": 1050000, "Thuế TNCN lũy tiến": 0, "Thu nhập 10%": 0, "Thuế 10%": 0},

        {"Mã NV": "NV012", "Họ và tên": "Trịnh Văn Phúc", "Phòng ban": "Vận Hành", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 9000000, "Lương quý": 1000000, "Thu nhập khác": 500000, "Phụ cấp không thuế": 730000,
         "Trừ hưu trí": 945000, "Thuế TNCN lũy tiến": 0, "Thu nhập 10%": 0, "Thuế 10%": 0},

        {"Mã NV": "NV013", "Họ và tên": "Võ Quỳnh Như", "Phòng ban": "Nhân Sự", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 10500000, "Lương quý": 1500000, "Thu nhập khác": 500000, "Phụ cấp không thuế": 730000,
         "Trừ hưu trí": 1102500, "Thuế TNCN lũy tiến": 0, "Thu nhập 10%": 0, "Thuế 10%": 0},

        # 2 Nhân sự thời vụ / vãng lai (khấu trừ 10%)
        {"Mã NV": "CTV01", "Họ và tên": "Nguyễn Văn Tâm", "Phòng ban": "Cộng Tác Viên", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 0, "Lương quý": 0, "Thu nhập khác": 15000000, "Phụ cấp không thuế": 0,
         "Trừ hưu trí": 0, "Thuế TNCN lũy tiến": 0, "Thu nhập 10%": 15000000, "Thuế 10%": 1500000},

        {"Mã NV": "CTV02", "Họ và tên": "Lê Thị Uyên", "Phòng ban": "Cộng Tác Viên", "Tình trạng cư trú": "Cư trú",
         "Lương tháng": 0, "Lương quý": 0, "Thu nhập khác": 10000000, "Phụ cấp không thuế": 0,
         "Trừ hưu trí": 0, "Thuế TNCN lũy tiến": 0, "Thu nhập 10%": 10000000, "Thuế 10%": 1000000},
    ]

    df_salary = pd.DataFrame(employees_data)
    salary_file_path = out_path / "Bang_Luong_Tong_CBCNV.xlsx"
    df_salary.to_excel(salary_file_path, index=False)
    print(f"-> Đã tạo file: {salary_file_path}")

    # Tính toán các chỉ tiêu thực tế để tạo file Tờ khai khớp 100%
    df_salary["tong_thu_nhap"] = df_salary["Lương tháng"] + df_salary["Lương quý"] + df_salary["Thu nhập khác"]
    df_salary["tnct"] = df_salary["tong_thu_nhap"] - df_salary["Trừ hưu trí"] - df_salary["Phụ cấp không thuế"]

    total_employees = len(df_salary)
    resident_employees = len(df_salary)  # toàn bộ đều cư trú
    sum_tnct_resident = df_salary["tnct"].sum()

    # Nhóm có thuế: Thuế lũy tiến > 0 hoặc Thuế 10% > 0
    taxable_mask = (df_salary["Thuế TNCN lũy tiến"] > 0) | (df_salary["Thuế 10%"] > 0)
    df_taxable = df_salary[taxable_mask]
    count_taxable = len(df_taxable)

    # TNCT của nhóm có thuế = Tổng TN nhóm có thuế - PC không thuế + TN 10%
    # Lưu ý đối với CTV: Tổng TN là 15tr, PC không thuế = 0, TN 10% = 15tr -> nếu theo công thức TNCT = Tổng TN - PC miễn + TN 10%
    sum_gross_taxable = df_taxable["tong_thu_nhap"].sum()
    sum_exempt_taxable = df_taxable["Phụ cấp không thuế"].sum()
    sum_income_10 = df_taxable["Thu nhập 10%"].sum()
    sum_tnct_with_tax = sum_gross_taxable - sum_exempt_taxable + sum_income_10

    total_progressive_tax = df_salary["Thuế TNCN lũy tiến"].sum()
    total_tax_10 = df_salary["Thuế 10%"].sum()
    total_tax = total_progressive_tax + total_tax_10

    # ==============================================================================
    # 2. FILE TỜ KHAI THUẾ PR025 (Mẫu 05/KK-TNCN)
    # ==============================================================================
    pr025_data = [
        {"Mã chỉ tiêu": "[21]", "Tên chỉ tiêu": "Tổng số người lao động", "Số tiền": total_employees},
        {"Mã chỉ tiêu": "[22]", "Tên chỉ tiêu": "Cá nhân cư trú có hợp đồng lao động", "Số tiền": resident_employees - 2},
        {"Mã chỉ tiêu": "[23]", "Tên chỉ tiêu": "Cá nhân không cư trú", "Số tiền": 0},
        {"Mã chỉ tiêu": "[24]", "Tên chỉ tiêu": "Tổng thu nhập chịu thuế trả cho cá nhân cư trú", "Số tiền": sum_tnct_resident},
        {"Mã chỉ tiêu": "[25]", "Tên chỉ tiêu": "Cá nhân cư trú thuộc diện khấu trừ thuế", "Số tiền": count_taxable},
        {"Mã chỉ tiêu": "[26]", "Tên chỉ tiêu": "Cá nhân không cư trú thuộc diện khấu trừ thuế", "Số tiền": 0},
        {"Mã chỉ tiêu": "[27]", "Tên chỉ tiêu": "Tổng thu nhập chịu thuế trả cho cá nhân cư trú thuộc diện khấu trừ thuế", "Số tiền": sum_tnct_with_tax},
        {"Mã chỉ tiêu": "[31]", "Tên chỉ tiêu": "Tổng số thuế TNCN đã khấu trừ của cá nhân cư trú", "Số tiền": total_tax},
        {"Mã chỉ tiêu": "[32]", "Tên chỉ tiêu": "Tổng số thuế TNCN đã khấu trừ của cá nhân không cư trú", "Số tiền": 0},
        {"Mã chỉ tiêu": "[33]", "Tên chỉ tiêu": "Tổng số thuế TNCN đã khấu trừ phát sinh trong kỳ ([31] + [32])", "Số tiền": total_tax},
    ]
    df_pr025 = pd.DataFrame(pr025_data)
    pr025_file_path = out_path / "To_Khai_Thue_PR025.xlsx"
    df_pr025.to_excel(pr025_file_path, index=False)
    print(f"-> Đã tạo file: {pr025_file_path}")

    # ==============================================================================
    # 3. FILE TỜ KHAI THUẾ PR027 (Bảng kê chi tiết khấu trừ thuế)
    # ==============================================================================
    pr027_data = [
        {"STT": 1, "Chỉ tiêu": "Thuế TNCN khấu trừ theo biểu lũy tiến (CBCNV ký HĐLĐ)", "Thuế lũy tiến": total_progressive_tax, "Thuế 10%": 0, "Thuế khấu trừ": total_progressive_tax},
        {"STT": 2, "Chỉ tiêu": "Thuế TNCN khấu trừ 10% (Lao động vãng lai, thời vụ)", "Thuế lũy tiến": 0, "Thuế 10%": total_tax_10, "Thuế khấu trừ": total_tax_10},
        {"STT": 3, "Chỉ tiêu": "TỔNG CỘNG THUẾ TNCN KHẤU TRỪ PHÁT SINH", "Thuế lũy tiến": total_progressive_tax, "Thuế 10%": total_tax_10, "Thuế khấu trừ": total_tax},
    ]
    df_pr027 = pd.DataFrame(pr027_data)
    pr027_file_path = out_path / "To_Khai_Thue_PR027.xlsx"
    df_pr027.to_excel(pr027_file_path, index=False)
    print(f"-> Đã tạo file: {pr027_file_path}")

    # ==============================================================================
    # 4. FILE SỔ CÁI / BẢNG KÊ TÀI KHOẢN 3335
    # ==============================================================================
    # Phát sinh Có = Số thuế trích từ lương trong kỳ (bằng total_tax)
    # Phát sinh Nợ = Nộp tiền thuế vào ngân sách nhà nước qua giấy nộp tiền
    gl_data = [
        {"Ngày HT": "2026-09-01", "Số chứng từ": "SDDK", "Diễn giải": "Số dư đầu kỳ TK 3335",
         "TK Nợ": "", "TK Có": "3335", "Phát sinh Nợ": 0, "Phát sinh Có": 0},
        {"Ngày HT": "2026-09-30", "Số chứng từ": "PKT-09/01", "Diễn giải": "Khấu trừ thuế TNCN từ lương CBCNV tháng 09/2026",
         "TK Nợ": "3341", "TK Có": "3335", "Phát sinh Nợ": 0, "Phát sinh Có": total_progressive_tax},
        {"Ngày HT": "2026-09-30", "Số chứng từ": "PKT-09/02", "Diễn giải": "Khấu trừ thuế TNCN 10% chi trả thù lao cộng tác viên",
         "TK Nợ": "3388", "TK Có": "3335", "Phát sinh Nợ": 0, "Phát sinh Có": total_tax_10},
        {"Ngày HT": "2026-09-20", "Số chứng từ": "UNC-09/15", "Diễn giải": "Nộp thuế TNCN kỳ trước vào NSNN qua KBNN",
         "TK Nợ": "3335", "TK Có": "1121", "Phát sinh Nợ": 25000000, "Phát sinh Có": 0},
        {"Ngày HT": "2026-09-30", "Số chứng từ": "TONG", "Diễn giải": "TỔNG PHÁT SINH TRONG KỲ",
         "TK Nợ": "", "TK Có": "", "Phát sinh Nợ": 25000000, "Phát sinh Có": total_tax},
    ]
    df_gl = pd.DataFrame(gl_data)
    gl_file_path = out_path / "So_Cai_TK_3335.xlsx"
    df_gl.to_excel(gl_file_path, index=False)
    print(f"-> Đã tạo file: {gl_file_path}")

    print("\n=== HOÀN TẤT TẠO 4 FILE MẪU KIỂM DÒ THUẾ TNCN THÀNH CÔNG ===")
    print(f"- Tổng số nhân viên: {total_employees}")
    print(f"- Tổng TNCT cá nhân cư trú [24]: {sum_tnct_resident:,.0f} VND")
    print(f"- Tổng TNCT nhóm có thuế [27]: {sum_tnct_with_tax:,.0f} VND")
    print(f"- Tổng thuế lũy tiến: {total_progressive_tax:,.0f} VND")
    print(f"- Tổng thuế 10%: {total_tax_10:,.0f} VND")
    print(f"- Tổng thuế TNCN [31] / Có TK 3335: {total_tax:,.0f} VND")


if __name__ == "__main__":
    generate_all_samples()
