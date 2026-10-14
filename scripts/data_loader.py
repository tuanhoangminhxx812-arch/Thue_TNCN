"""
Module: data_loader.py
Mục đích: Đọc và chuẩn hóa dữ liệu từ file Excel (.xlsx), hỗ trợ cả đường dẫn cục bộ (file path)
          lẫn đối tượng bộ nhớ đệm (io.BytesIO / Streamlit UploadedFile).
          Xử lý an toàn các định dạng header đa tầng, loại bỏ trùng cột,
          và tương thích với cả biểu mẫu EVN thực tế lẫn mẫu doanh nghiệp thông thường.
"""

import io
import json
import logging
from pathlib import Path
from typing import Dict, Any, Union, Tuple, Optional
import pandas as pd
import numpy as np

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


def load_mapping_rules(mapping_file: Union[str, Path] = "config/mapping_rules.json") -> Dict[str, Any]:
    """Tải quy tắc ánh xạ tên cột từ file JSON."""
    path = Path(mapping_file)
    if not path.exists():
        logger.warning(f"Không tìm thấy file ánh xạ {mapping_file}, sử dụng cấu hình mặc định.")
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def find_header_row(file_source: Union[str, Path, io.BytesIO], max_rows: int = 20, expected_keywords: list = None) -> int:
    """
    Tự động tìm dòng tiêu đề (header) của bảng tính nếu file có nhiều dòng thông tin phía trên.
    """
    if expected_keywords is None:
        expected_keywords = ["msnv", "mã", "họ và tên", "tiền lương", "thu nhập", "thuế", "chỉ tiêu", "phát sinh", "ngày", "số tiền", "chứng từ"]

    try:
        if hasattr(file_source, "seek"):
            file_source.seek(0)
            df_preview = pd.read_excel(file_source, header=None, nrows=max_rows)
            file_source.seek(0)
        else:
            df_preview = pd.read_excel(file_source, header=None, nrows=max_rows)

        best_row = 0
        max_matches = 0

        for r_idx, row in df_preview.iterrows():
            row_str = " ".join([str(val).lower() for val in row.dropna().tolist()])
            matches = sum(1 for kw in expected_keywords if kw in row_str)
            if matches > max_matches:
                max_matches = matches
                best_row = r_idx

        # Nếu phát hiện dòng có từ khóa hợp lý
        if max_matches >= 2:
            return best_row
        return 0
    except Exception as e:
        logger.warning(f"Không thể tự động tìm header row, mặc định dùng dòng 0. Lỗi: {e}")
        if hasattr(file_source, "seek"):
            file_source.seek(0)
        return 0


def normalize_column_name(col_name: Any) -> str:
    """Loại bỏ khoảng trắng thừa và ký tự đặc biệt ở đầu cuối của tên cột."""
    if pd.isna(col_name):
        return ""
    col_str = str(col_name).strip()
    col_str = " ".join(col_str.split())
    return col_str


def apply_column_mapping(df: pd.DataFrame, file_type: str, mapping_rules: Dict[str, Any]) -> pd.DataFrame:
    """
    Đổi tên các cột trong DataFrame sang tên chuẩn nội bộ dựa trên mapping_rules.
    Đảm bảo không bao giờ gán 1 tên chuẩn cho nhiều cột (tránh duplicate columns).
    """
    # Khử trùng tên ban đầu nếu có
    df = df.loc[:, ~df.columns.duplicated()].copy()

    rules = mapping_rules.get(file_type, {})
    if not rules:
        return df

    rename_dict = {}
    assigned_std = set()

    # Pass 1: Khớp chính xác (exact match)
    for col in df.columns:
        norm_col = normalize_column_name(col).lower()
        if not norm_col or norm_col.startswith("unnamed"):
            continue
        for std_name, aliases in rules.items():
            if std_name in assigned_std:
                continue
            cleaned_aliases = [a.strip().lower() for a in aliases]
            if norm_col in cleaned_aliases:
                rename_dict[col] = std_name
                assigned_std.add(std_name)
                break

    # Pass 2: Khớp từ khóa (substring match) với các cột chưa được gán
    for col in df.columns:
        if col in rename_dict:
            continue
        norm_col = normalize_column_name(col).lower()
        if not norm_col or norm_col.startswith("unnamed"):
            continue
        for std_name, aliases in rules.items():
            if std_name in assigned_std:
                continue
            for alias in aliases:
                a_clean = alias.strip().lower()
                # Chỉ match substring nếu alias đủ dài (ít nhất 4 ký tự) và xuất hiện trọn vẹn
                if len(a_clean) >= 4 and a_clean in norm_col:
                    rename_dict[col] = std_name
                    assigned_std.add(std_name)
                    break
            if col in rename_dict:
                break

    df = df.rename(columns=rename_dict)
    # Khử trùng lặp cột một lần nữa
    df = df.loc[:, ~df.columns.duplicated()].copy()
    return df


def identify_input_files(sources: Union[str, Path, list]) -> Dict[str, Any]:
    """
    Tự động nhận diện 4 file đầu vào từ một thư mục hoặc từ danh sách uploaded files.
    Trả về dict:
    {
        'salary': file_or_path,
        'pr025': file_or_path,
        'pr027': file_or_path,
        'gl_3335': file_or_path
    }
    """
    file_list = []
    if isinstance(sources, (str, Path)):
        p = Path(sources)
        if p.is_dir():
            file_list = list(p.glob("*.xlsx")) + list(p.glob("*.xls"))
        elif p.exists():
            file_list = [p]
    elif isinstance(sources, list):
        file_list = sources

    identified = {
        "salary": None,
        "pr025": None,
        "pr027": None,
        "gl_3335": None
    }

    for item in file_list:
        fname = getattr(item, "name", str(item)).lower()
        # Bỏ qua file tạm mở của Excel (~$)
        if "~$" in fname:
            continue
        if "3335" in fname or "so_cai" in fname:
            identified["gl_3335"] = item
        elif "025" in fname or "pr025" in fname or "05_kk" in fname or "05-kk" in fname:
            identified["pr025"] = item
        elif "027" in fname or "028" in fname or "pr027" in fname or "pr028" in fname or "nghi_huu" in fname:
            identified["pr027"] = item
        elif any(k in fname for k in ["luong", "pr_ttt", "ttt", "salary", "tam_tinh", "cbcnv"]):
            identified["salary"] = item

    return identified


def clean_numeric_series(series: Any) -> pd.Series:
    """Làm sạch và ép kiểu dữ liệu tiền tệ/số về dạng float/int chuẩn (làm tròn số nguyên)."""
    if series is None:
        return pd.Series(dtype=float)
    if isinstance(series, pd.DataFrame):
        series = series.iloc[:, 0]
    if not isinstance(series, pd.Series):
        series = pd.Series(series)

    def parse_val(v):
        if pd.isna(v) or v is None or v == "" or v == "-":
            return 0.0
        if isinstance(v, (int, float)):
            return float(round(v))
        s = str(v).strip().replace(" ", "").replace("\xa0", "")
        if s == "" or s == "-":
            return 0.0
        # Xử lý dấu phân cách
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            parts = s.split(",")
            if len(parts) > 2 or any(len(p) == 3 for p in parts[1:]):
                s = s.replace(",", "")
            else:
                s = s.replace(",", ".")
        try:
            return float(round(float(s)))
        except ValueError:
            return 0.0

    return series.apply(parse_val)


def get_series_safe(df: pd.DataFrame, col_name: str) -> Optional[pd.Series]:
    """Lấy an toàn một pd.Series từ DataFrame kể cả khi có duplicate column."""
    if col_name not in df.columns:
        return None
    val = df[col_name]
    if isinstance(val, pd.DataFrame):
        return val.iloc[:, 0]
    return val


class DataLoader:
    """Lớp xử lý việc nạp và chuẩn hóa dữ liệu từ 4 file đầu vào."""

    def __init__(self, mapping_rules_path: Union[str, Path] = "config/mapping_rules.json"):
        self.mapping_rules = load_mapping_rules(mapping_rules_path)

    def load_salary_data(self, file_source: Union[str, Path, io.BytesIO]) -> Tuple[pd.DataFrame, Dict[str, Any]]:
        """
        Nạp file Bảng lương tổng CBCNV (Hỗ trợ cả file EVN và bảng lương tiêu chuẩn).
        """
        header_row = find_header_row(
            file_source,
            expected_keywords=["msnv", "mã", "họ và tên", "tiền lương", "lương tháng", "phụ cấp", "bảo hiểm", "thuế"]
        )
        if hasattr(file_source, "seek"):
            file_source.seek(0)

        df = pd.read_excel(file_source, header=header_row)
        df = apply_column_mapping(df, "salary", self.mapping_rules)

        # Kiểm tra cột bắt buộc tối thiểu
        required_cols = ["employee_id", "full_name"]
        missing_required = [c for c in required_cols if c not in df.columns]
        if missing_required:
            cols_preview = [str(c) for c in df.columns[:8]]
            raise ValueError(
                f"File Bảng lương thiếu cột định danh bắt buộc: {missing_required}. "
                f"Các cột tìm thấy: {cols_preview}"
            )

        # Đảm bảo lấy Series an toàn
        emp_series = get_series_safe(df, "employee_id")
        name_series = get_series_safe(df, "full_name")

        # Loại bỏ các dòng trống hoặc dòng tiêu đề lặp lại
        mask_valid = emp_series.notna() & name_series.notna()
        # Loại bỏ dòng chứa từ "STT" hoặc "MSNV" ở hàng dữ liệu
        mask_valid = mask_valid & (~emp_series.astype(str).str.lower().str.contains("msnv|mã|stt"))
        mask_valid = mask_valid & (~name_series.astype(str).str.lower().str.contains("họ và tên|tổng cộng|total"))
        df = df[mask_valid].copy()

        # Đảm bảo các cột số học tồn tại, nếu thiếu gán mặc định = 0.0
        numeric_columns = [
            "salary_monthly", "salary_quarterly", "other_income", "total_income",
            "tax_exempt_allowance", "insurance_deduction", "personal_relief",
            "dependent_relief", "progressive_tax", "income_10_percent", "tax_10_percent",
            "taxable_income"
        ]

        for col in numeric_columns:
            if col not in df.columns:
                df[col] = 0.0
            else:
                df[col] = clean_numeric_series(get_series_safe(df, col))

        # Nếu đã có cột total_income sẵn trong bảng mà > 0, dùng luôn
        # Nếu chưa có, tính total_income = salary_monthly + salary_quarterly + other_income
        calc_total = df["salary_monthly"] + df["salary_quarterly"] + df["other_income"]
        if df["total_income"].sum() == 0 and calc_total.sum() > 0:
            df["total_income"] = calc_total

        # Đảm bảo cột phân loại cư trú
        if "resident_status" not in df.columns:
            df["resident_status"] = "Cư trú"
        else:
            df["resident_status"] = df["resident_status"].fillna("Cư trú").astype(str).str.strip()

        if "department" not in df.columns:
            df["department"] = "Không xác định"

        summary = {
            "total_employees": len(df),
            "residents": len(df[df["resident_status"].str.lower().isin(["cư trú", "cu tru", "ct", "có"])]),
            "total_income": float(df["total_income"].sum()),
            "total_progressive_tax": float(df["progressive_tax"].sum()),
            "total_tax_10": float(df["tax_10_percent"].sum()),
        }

        return df, summary

    def load_pr025_data(self, file_source: Union[str, Path, io.BytesIO]) -> Tuple[pd.DataFrame, Dict[str, float]]:
        """
        Nạp file Tờ khai Thuế PR025 (Hỗ trợ cả Thông tư 80 và Thông tư cũ).
        """
        header_row = find_header_row(file_source, expected_keywords=["chỉ tiêu", "mã chỉ tiêu", "số người/số tiền", "số tiền"])
        if hasattr(file_source, "seek"):
            file_source.seek(0)

        df = pd.read_excel(file_source, header=header_row)
        df = apply_column_mapping(df, "pr025", self.mapping_rules)

        raw_indicators: Dict[str, float] = {}

        code_s = get_series_safe(df, "code")
        amt_s = get_series_safe(df, "amount")
        name_s = get_series_safe(df, "item_name")

        if code_s is not None and amt_s is not None:
            amt_clean = clean_numeric_series(amt_s)
            for idx in range(len(df)):
                code_raw = str(code_s.iloc[idx]).strip()
                amt_val = float(amt_clean.iloc[idx])
                # Trích xuất số trong dấu ngoặc vuông ví dụ [21] -> 21
                if "[" in code_raw and "]" in code_raw:
                    k = code_raw[code_raw.find("[") + 1 : code_raw.find("]")].strip()
                    raw_indicators[k] = amt_val
                elif code_raw.isdigit():
                    raw_indicators[code_raw] = amt_val

        # Chuẩn hóa các chỉ tiêu chính:
        # Trong TT 80:
        # [21]: Tổng số người lao động
        # [22]: Cá nhân cư trú có HĐLĐ
        # [27]: Tổng TNCT cá nhân cư trú (hoặc [24] ở TT cũ)
        # [30]: Tổng TNCT cá nhân cư trú thuộc diện khấu trừ thuế (hoặc [27] ở TT cũ)
        # [33]: Tổng thuế cá nhân cư trú khấu trừ (hoặc [31] ở TT cũ)
        # Kiểm tra xem mẫu tờ khai đang dùng là Thông tư 80 hay Thông tư cũ
        # Trong TT 80: luôn có chỉ tiêu [29] hoặc [30] (Tổng TNCT cá nhân trả cho cá nhân thuộc diện phải khấu trừ thuế)
        is_tt80 = ("29" in raw_indicators) or ("30" in raw_indicators)

        if is_tt80:
            tnct_resident = raw_indicators.get("27", raw_indicators.get("26", 0.0))
            tnct_with_tax = raw_indicators.get("30", raw_indicators.get("29", 0.0))
            tax_total = raw_indicators.get("33", raw_indicators.get("32", 0.0))
        else:
            # Mẫu TT cũ hoặc bảng tóm tắt cơ bản ([24] là TNCT cư trú, [27] là TNCT cư trú có thuế, [31] là thuế)
            tnct_resident = raw_indicators.get("24", 0.0)
            tnct_with_tax = raw_indicators.get("27", 0.0)
            tax_total = raw_indicators.get("31", raw_indicators.get("33", 0.0))

        indicators = {
            "21": raw_indicators.get("21", 0.0),
            "22": raw_indicators.get("22", 0.0),
            "24": tnct_resident,   # Chuẩn hóa về [24] để reconciler sử dụng
            "27": tnct_with_tax,   # Chuẩn hóa về [27] để reconciler sử dụng
            "31": tax_total,       # Chuẩn hóa về [31] để reconciler sử dụng
            "33": tax_total
        }

        return df, indicators

    def load_pr027_data(self, file_source: Union[str, Path, io.BytesIO]) -> Tuple[pd.DataFrame, Dict[str, float]]:
        """
        Nạp file Tờ khai Thuế PR027/PR028 (Bảng kê chi tiết thuế khấu trừ).
        Hỗ trợ cả file EVN dạng chi tiết nhân sự lẫn bảng tổng hợp chỉ tiêu.
        """
        header_row = find_header_row(file_source, expected_keywords=["msnv", "họ và tên", "số thuế phải nộp", "thuế", "lũy tiến", "vãng lai"])
        if hasattr(file_source, "seek"):
            file_source.seek(0)

        df = pd.read_excel(file_source, header=header_row)

        # Kiểm tra nếu dòng đầu tiên là tiêu đề con (ví dụ 'Số thuế phải nộp', 'Tiền thưởng...')
        if len(df) > 0:
            first_row_str = " ".join([str(v).lower() for v in df.iloc[0].dropna().tolist()])
            if "số thuế phải nộp" in first_row_str or "tiền thưởng" in first_row_str:
                new_cols = []
                for i, col in enumerate(df.columns):
                    val = df.iloc[0, i]
                    if pd.notna(val) and str(val).strip() != "" and str(val).strip() != str(col).strip():
                        new_cols.append(str(val).strip())
                    else:
                        new_cols.append(str(col).strip())
                df.columns = new_cols
                df = df.iloc[1:].copy()

        df = apply_column_mapping(df, "pr027", self.mapping_rules)

        summary_pr027 = {
            "progressive_tax": 0.0,
            "tax_10_percent": 0.0,
            "total_tax": 0.0
        }

        # Tìm cột thuế
        amt_col = "amount" if "amount" in df.columns else None
        if not amt_col:
            for c in df.columns:
                c_norm = str(c).lower().replace(" ", "")
                if "thuếphảinộp" in c_norm or "sốthuế" in c_norm or "thuếkhấutrừ" in c_norm:
                    amt_col = c
                    break

        if amt_col:
            df[amt_col] = clean_numeric_series(get_series_safe(df, amt_col))

            # Tìm xem có dòng "Tổng cộng" ở cuối không bằng cách quét các cột
            mask_is_total = pd.Series(False, index=df.index)
            for c in df.columns:
                if c != amt_col:
                    s_series = get_series_safe(df, c)
                    if s_series is not None:
                        mask_is_total = mask_is_total | s_series.astype(str).str.lower().str.contains("tổng cộng|tổng số|total|cộng phát sinh")

            if mask_is_total.sum() > 0:
                total_val = float(df.loc[mask_is_total, amt_col].iloc[0])
                summary_pr027["total_tax"] = total_val
            else:
                summary_pr027["total_tax"] = float(df[amt_col].sum())

            # Gán thuế lũy tiến và 10%
            summary_pr027["tax_10_percent"] = summary_pr027["total_tax"]
            summary_pr027["progressive_tax"] = 0.0

        return df, summary_pr027

    def load_gl_3335_data(self, file_source: Union[str, Path, io.BytesIO]) -> Tuple[pd.DataFrame, Dict[str, float]]:
        """
        Nạp Sổ cái / Bảng kê Tài khoản 3335.
        Trích xuất an toàn Phát sinh Có và Phát sinh Nợ.
        """
        header_row = find_header_row(file_source, expected_keywords=["chứng từ", "ngày", "số", "diễn giải", "phát sinh", "nợ", "có"])
        if hasattr(file_source, "seek"):
            file_source.seek(0)

        # Đọc trước file raw không header để có thể tìm dòng tổng cộng
        if hasattr(file_source, "seek"):
            file_source.seek(0)
            df_raw = pd.read_excel(file_source, header=None)
            file_source.seek(0)
        else:
            df_raw = pd.read_excel(file_source, header=None)

        df = pd.read_excel(file_source, header=header_row)
        df = apply_column_mapping(df, "gl_3335", self.mapping_rules)

        total_credit = 0.0
        total_debit = 0.0

        # Cách 1: Tìm dòng tổng phát sinh từ raw sheet (rất chính xác cho mẫu EVN và ERP)
        found_summary_row = False
        for idx in range(len(df_raw)):
            row_str = " ".join([str(v).lower() for v in df_raw.iloc[idx].dropna().tolist()])
            if "tổng số phát sinh trong kỳ" in row_str or "tổng phát sinh trong kỳ" in row_str or "cộng phát sinh trong kỳ" in row_str:
                row_vals = df_raw.iloc[idx]
                # Lấy các giá trị số trong dòng này
                num_vals = []
                for val in row_vals:
                    cleaned_val = clean_numeric_series(pd.Series([val])).iloc[0]
                    if cleaned_val > 0:
                        num_vals.append(cleaned_val)
                if num_vals:
                    # Giá trị phát sinh có
                    total_credit = num_vals[-1]
                    found_summary_row = True
                    break

        # Cách 2: Nếu chưa tìm thấy dòng tổng từ raw, tính từ DataFrame các dòng phát sinh chi tiết
        if not found_summary_row:
            # Lấy cột Có và Nợ
            credit_s = get_series_safe(df, "credit_amount")
            debit_s = get_series_safe(df, "debit_amount")

            if credit_s is not None:
                credit_clean = clean_numeric_series(credit_s)
                # Loại trừ các dòng số dư đầu kỳ, cuối kỳ
                desc_s = get_series_safe(df, "description")
                doc_s = get_series_safe(df, "doc_no")

                mask_exclude = pd.Series(False, index=df.index)
                for s in [desc_s, doc_s]:
                    if s is not None:
                        s_str = s.astype(str).str.lower()
                        for kw in ["số dư đầu kỳ", "số dư cuối kỳ", "tổng cộng", "tổng phát sinh", "sddk", "sdck"]:
                            mask_exclude = mask_exclude | s_str.str.contains(kw)

                total_credit = float(credit_clean[~mask_exclude].sum())

            if debit_s is not None:
                debit_clean = clean_numeric_series(debit_s)
                total_debit = float(debit_clean[~mask_exclude].sum()) if 'mask_exclude' in locals() else float(debit_clean.sum())

        summary_gl = {
            "credit_turnover": total_credit,
            "debit_turnover": total_debit,
            "net_turnover": total_credit - total_debit
        }

        return df, summary_gl
