"""
File: main.py
Mục đích: Chạy script kiểm dò thuế TNCN offline trên máy tính cá nhân.
          Tự động nhận diện cả 4 file từ 1 THƯ MỤC dữ liệu (mặc định theo config.yaml
          hoặc truyền qua dòng lệnh: python main.py [đường_dẫn_thư_mục]).
          Định dạng toàn bộ số nguyên không có số lẻ, phân cách hàng nghìn bằng dấu chấm.
"""

import sys
import os
from pathlib import Path
from datetime import datetime
import yaml

# Đảm bảo console Windows hỗ trợ tiếng Việt UTF-8
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# Thêm thư mục gốc vào PYTHONPATH nếu cần
sys.path.insert(0, str(Path(__file__).resolve().parent))

from scripts.data_loader import DataLoader, identify_input_files
from scripts.reconciler import TaxReconciler
from scripts.report_builder import ExcelReportBuilder


def load_config(config_path: str = "config/config.yaml") -> dict:
    """Tải cấu hình từ config.yaml."""
    cfg_file = Path(config_path)
    if not cfg_file.exists():
        print(f"[CẢNH BÁO] Không tìm thấy file {config_path}. Dùng cấu hình mặc định.")
        return {
            "company_name": "TỔNG CÔNG TY ĐIỆN LỰC TP. HỒ CHÍ MINH - CÔNG TY ĐIỆN LỰC CHỢ LỚN",
            "report_period": "Tháng 09/2026",
            "paths": {
                "input_dir": "02_THAM_KHAO",
                "output_dir": "03_DAU_RA"
            },
            "output": {
                "report_prefix": "Bao_Cao_Kiem_Do_Thue_TNCN"
            },
            "thresholds": {"warning_tolerance": 1000}
        }

    with open(cfg_file, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def format_dot_int(val) -> str:
    """Định dạng số nguyên phân cách hàng nghìn bằng dấu chấm (không số lẻ)."""
    if val is None or val == "-":
        return "-"
    try:
        int_val = int(round(float(val)))
        return f"{int_val:,}".replace(",", ".")
    except (ValueError, TypeError):
        return str(val)


def print_banner(company_name: str, period: str, folder_path: str):
    print("=" * 80)
    print(f"   HỆ THỐNG KIỂM DÒ DỮ LIỆU THUẾ TNCN - CÔNG TY ĐIỆN LỰC")
    print(f"   Đơn vị: {company_name}")
    print(f"   Kỳ báo cáo: {period} | Thư mục dữ liệu: {folder_path}")
    print(f"   Thời gian thực hiện: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("=" * 80)


def print_summary_table(results: dict):
    """In bảng tóm tắt kết quả kiểm dò ra console (số nguyên, phân cách dấu chấm)."""
    kpi = results["summary_kpis"]
    print("\n" + "-" * 80)
    print("                      BẢNG TỔNG HỢP KẾT QUẢ KIỂM DÒ")
    print("-" * 80)
    print(f"1. Tổng nhân sự:                       {kpi['total_employees']:>18} người (Cư trú: {kpi['residents_count']})")
    print(f"2. Số nhân sự phát sinh thuế TNCN:     {kpi['taxable_employees_count']:>18} người")
    print(f"3. Tổng thu nhập CBCNV (Gross):        {format_dot_int(kpi['total_income']):>18} VND")
    print("-" * 80)
    print("KẾT QUẢ ĐỐI CHIẾU CÁC CHỈ TIÊU TRỌNG YẾU (SỐ NGUYÊN, PHÂN CÁCH BẰNG DẤU CHẤM):")

    # [24]
    status_24 = "[ KHỚP ]" if kpi["status_tnct"] == "KHỚP" else "[ LỆCH! ]"
    print(f" - [24]/[27] TNCT cá nhân cư trú:")
    print(f"     + Theo Bảng lương:                {format_dot_int(kpi['taxable_income_resident']):>18} VND")
    print(f"     + Theo Tờ khai PR025:             {format_dot_int(kpi['pr025_tnct_resident']):>18} VND")
    print(f"     + Chênh lệch:                     {format_dot_int(kpi['diff_tnct_resident']):>18} VND  ==> {status_24}")

    # [27]
    status_27 = "[ KHỚP ]" if kpi["status_with_tax"] == "KHỚP" else "[ LỆCH! ]"
    print(f" - [27]/[30] TNCT nhóm có phát sinh thuế:")
    print(f"     + Theo Bảng lương:                {format_dot_int(kpi['taxable_income_with_tax']):>18} VND")
    print(f"     + Theo Tờ khai PR025:             {format_dot_int(kpi['pr025_tnct_with_tax']):>18} VND")
    print(f"     + Chênh lệch:                     {format_dot_int(kpi['diff_tnct_with_tax']):>18} VND  ==> {status_27}")

    # Thuế 3 bên
    status_tax = "[ KHỚP ]" if kpi["status_tax"] == "KHỚP" else "[ LỆCH! ]"
    print(f" - Đối chiếu số thuế TNCN phải nộp (3 bên):")
    print(f"     + Tính từ Bảng lương:             {format_dot_int(kpi['salary_total_tax']):>18} VND")
    print(f"     + Tờ khai PR025 / PR027:          {format_dot_int(kpi['tax_return_total']):>18} VND")
    print(f"     + Phát sinh Có TK 3335 (Sổ cái):  {format_dot_int(kpi['gl_credit_turnover']):>18} VND")
    print(f"     + Lệch Bảng lương vs Tờ khai:     {format_dot_int(kpi['diff_salary_vs_return']):>18} VND")
    print(f"     + Lệch Bảng lương vs Sổ cái:      {format_dot_int(kpi['diff_salary_vs_gl']):>18} VND  ==> {status_tax}")
    print("-" * 80)

    if results["all_passed"]:
        print(">>> KẾT QUẢ CHUNG: TOÀN BỘ SỐ LIỆU ĐÃ KHỚP CHÍNH XÁC 100%! <<<")
    else:
        print(">>> CẢNH BÁO: PHÁT HIỆN SỐ LIỆU CHÊNH LỆCH! HÃY KIỂM TRA FILE BÁO CÁO CHI TIẾT! <<<")
    print("-" * 80)


def main():
    config = load_config()
    company_name = config.get("company_name", "TỔNG CÔNG TY ĐIỆN LỰC TP. HỒ CHÍ MINH - CÔNG TY ĐIỆN LỰC CHỢ LỚN")
    period = config.get("report_period", "Tháng 09/2026")
    tolerance = config.get("thresholds", {}).get("warning_tolerance", 1000)

    # Lấy thư mục dữ liệu từ tham số dòng lệnh hoặc config.yaml
    if len(sys.argv) > 1:
        folder_path = sys.argv[1].strip()
    else:
        # Ưu tiên kiểm tra thư mục 02_THAM_KHAO (chứa dữ liệu EVN thực tế) nếu có đủ 4 file
        input_dir_cfg = config.get("paths", {}).get("input_dir", "02_THAM_KHAO")
        if Path(input_dir_cfg).exists():
            folder_path = input_dir_cfg
        elif Path("02_THAM_KHAO").exists():
            folder_path = "02_THAM_KHAO"
        else:
            folder_path = "01_DAU_VAO"

    print_banner(company_name, period, folder_path)

    folder_p = Path(folder_path)
    if not folder_p.exists() or not folder_p.is_dir():
        print(f"[LỖI] Thư mục '{folder_path}' không tồn tại hoặc không phải là thư mục hợp lệ.")
        return 1

    print(f"\n[1/4] Đang tự động quét và nhận diện 4 file trong thư mục '{folder_path}'...")
    identified_files = identify_input_files(folder_p)

    salary_path = identified_files["salary"]
    pr025_path = identified_files["pr025"]
    pr027_path = identified_files["pr027"]
    gl_path = identified_files["gl_3335"]

    missing = []
    if not salary_path: missing.append("Bảng lương tổng CBCNV (chứa từ khóa 'luong', 'pr_ttt'...)")
    if not pr025_path: missing.append("Tờ khai Thuế PR025 (chứa từ khóa '025'...)")
    if not pr027_path: missing.append("Tờ khai Thuế PR027 (chứa từ khóa '027'...)")
    if not gl_path: missing.append("Sổ cái TK 3335 (chứa từ khóa '3335'...)")

    print(f" - Bảng lương:   {'[V] ' + str(salary_path) if salary_path else '[X] Chưa tìm thấy'}")
    print(f" - Tờ khai PR025: {'[V] ' + str(pr025_path) if pr025_path else '[X] Chưa tìm thấy'}")
    print(f" - Tờ khai PR027: {'[V] ' + str(pr027_path) if pr027_path else '[X] Chưa tìm thấy'}")
    print(f" - Sổ cái 3335:   {'[V] ' + str(gl_path) if gl_path else '[X] Chưa tìm thấy'}")

    if missing:
        print(f"\n[LỖI] Thư mục '{folder_path}' thiếu các file sau:")
        for m in missing:
            print(f"  * {m}")
        print("Vui lòng đặt đủ 4 file Excel vào thư mục trên rồi chạy lại lệnh.")
        return 1

    print("\n[2/4] Đang nạp và chuẩn hóa dữ liệu từ các file...")
    loader = DataLoader(config.get("paths", {}).get("mapping_rules", "config/mapping_rules.json"))

    try:
        df_salary, summary_salary = loader.load_salary_data(salary_path)
        print(f" - Bảng lương: Đã đọc {len(df_salary)} nhân sự, Tổng thu nhập = {format_dot_int(summary_salary['total_income'])} VND.")

        df_pr025, pr025_indicators = loader.load_pr025_data(pr025_path)
        print(f" - Tờ khai PR025: Đã trích xuất TNCT cư trú = {format_dot_int(pr025_indicators.get('24', 0))} VND, Thuế khấu trừ = {format_dot_int(pr025_indicators.get('31', 0))} VND.")

        df_pr027, pr027_summary = loader.load_pr027_data(pr027_path)
        print(f" - Tờ khai PR027: Tổng thuế khấu trừ = {format_dot_int(pr027_summary.get('total_tax', 0))} VND.")

        df_gl, gl_summary = loader.load_gl_3335_data(gl_path)
        print(f" - Sổ cái TK 3335: Phát sinh Có = {format_dot_int(gl_summary.get('credit_turnover', 0))} VND.")

    except Exception as e:
        print(f"[LỖI XỬ LÝ DỮ LIỆU]: {e}")
        return 1

    print("\n[3/4] Đang thực hiện logic kiểm dò và đối chiếu 3 bên...")
    reconciler = TaxReconciler(tolerance=tolerance)
    results = reconciler.reconcile_all(
        df_salary=df_salary,
        pr025_indicators=pr025_indicators,
        pr027_summary=pr027_summary,
        gl_3335_summary=gl_summary,
        report_period=period
    )

    print_summary_table(results)

    print("[4/4] Đang tạo và định dạng file Excel báo cáo hoàn chỉnh (openpyxl)...")
    out_dir = Path(config.get("paths", {}).get("output_dir", "03_DAU_RA"))
    prefix = config.get("output", {}).get("report_prefix", "Bao_Cao_Kiem_Do_Thue_TNCN")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    clean_period = period.replace(" ", "_").replace("/", "_").replace("á", "a").replace("Á", "A")
    out_filename = f"{prefix}_{clean_period}_{timestamp}.xlsx"
    out_file_path = out_dir / out_filename

    builder = ExcelReportBuilder(company_name=company_name, report_period=period)
    saved_path = builder.export_to_file(results, out_file_path)

    print(f"\n[THÀNH CÔNG] Đã xuất file báo cáo Excel nhiều sheet tại:")
    print(f"  -> {saved_path}")
    print("\nQuy cách dữ liệu trong file Excel:")
    print("  * Toàn bộ số liệu là SỐ NGUYÊN (Integer), không có số lẻ thập phân.")
    print("  * Dấu phân cách hàng nghìn chuẩn kế toán.")
    print("  * Số căn lề PHẢI, chữ căn lề TRÁI.")
    return 0


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
