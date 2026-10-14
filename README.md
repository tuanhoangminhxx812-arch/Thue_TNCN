# HỆ THỐNG KIỂM DÒ DỮ LIỆU THUẾ TNCN (PERSONAL INCOME TAX RECONCILIATION)

Ứng dụng Python chuyên nghiệp dành cho bộ phận Kế toán - Nhân sự - Kiểm toán nội bộ nhằm tự động hóa quy trình kiểm tra, đối chiếu và phát hiện sai lệch số liệu thuế Thu nhập cá nhân (TNCN) giữa:
1. **Bảng lương tổng CBCNV**
2. **Tờ khai Thuế TNCN PR025 (Mẫu 05/KK-TNCN)**
3. **Tờ khai / Bảng kê Thuế PR027 (hoặc PR028)**
4. **Sổ cái / Bảng kê chi tiết Tài khoản 3335 (Thuế TNCN phải nộp)**

Hệ thống hỗ trợ cả **Giao diện Web tương tác cao bằng Streamlit** (`app.py`) và **Chạy script offline tự động theo lô** (`main.py`).

---

## 1. CẤU TRÚC THƯ MỤC DỰ ÁN

```text
Thue_TNCN/
├── 01_DAU_VAO/                  # Chứa các file mẫu / file test cục bộ (.xlsx)
├── 02_THAM_KHAO/                # Chứa danh mục tham chiếu cố định (nếu có)
├── 03_DAU_RA/                   # Chứa file kết quả xuất offline theo timestamp
├── config/
│   ├── config.yaml              # Cấu hình đường dẫn, tên công ty, kỳ báo cáo, dung sai
│   └── mapping_rules.json       # Quy tắc ánh xạ và nhận diện tên cột linh hoạt
├── scripts/
│   ├── __init__.py
│   ├── data_loader.py           # Module nạp, làm sạch và chuẩn hóa dữ liệu
│   ├── reconciler.py            # Module logic đối chiếu, tính toán và kiểm dò 4 sheet
│   ├── report_builder.py        # Module tạo và định dạng file Excel (openpyxl)
│   └── generate_sample_data.py  # Script tạo bộ 4 file Excel mẫu chuẩn số liệu
├── app.py                       # Giao diện Web Streamlit cho người dùng
├── main.py                      # Script chạy kiểm dò offline trên máy tính
├── requirements.txt             # Danh sách thư viện cần thiết
└── README.md                    # Tài liệu hướng dẫn sử dụng chi tiết
```

---

## 2. YÊU CẦU CÀI ĐẶT

### 2.1. Yêu cầu môi trường
- Python 3.9 trở lên (Khuyến nghị Python 3.10, 3.11 hoặc 3.12/3.14).
- Hệ điều hành: Windows, macOS hoặc Linux.

### 2.2. Cài đặt thư viện
Mở Command Prompt / PowerShell / Terminal tại thư mục `Thue_TNCN` và chạy lệnh:

```bash
pip install -r requirements.txt
```

---

## 3. HƯỚNG DẪN SỬ DỤNG

### 3.1. Chạy Giao diện Web tương tác (Streamlit)
Phù hợp cho người dùng muốn kéo thả file, xem biểu đồ, kiểm tra trực quan trên trình duyệt:

```bash
streamlit run app.py
```

Trình duyệt sẽ tự động mở tại địa chỉ `http://localhost:8501`.

**Tính năng nổi bật trên Web:**
- **Sidebar thông minh:** Có 4 nút Upload riêng biệt cho 4 file `.xlsx` và 1 ô nhập Kỳ báo cáo (ví dụ: `Tháng 09/2026`).
- **Chế độ Demo nhanh:** Có checkbox *"Dùng bộ dữ liệu mẫu (Demo có sẵn)"* giúp kiểm thử ngay lập tức chỉ với 1 click chuột mà không cần tải file lên.
- **Xử lý In-Memory (`io.BytesIO`):** Dữ liệu được đọc và xử lý hoàn toàn trong RAM, không ghi file tạm lên đĩa cứng, bảo đảm bảo mật dữ liệu nhân sự và tránh xung đột khi nhiều người cùng truy cập web.
- **Thẻ KPI & Báo động màu sắc trực quan:**
  - **Khớp (Chênh lệch = 0 VND):** Huy hiệu màu xanh lá (`#E2EFDA`).
  - **Lệch (Chênh lệch khác 0 VND):** Huy hiệu cảnh báo màu cam đỏ (`#FCE4D6`).
- **Tabs chi tiết:** Dễ dàng chuyển đổi giữa các bảng tổng kết và danh sách chi tiết.
- **Nút tải Excel hoàn chỉnh:** Tải file báo cáo đa sheet định dạng chuẩn kế toán về máy tính cá nhân.

---

### 3.2. Chạy Script Offline cục bộ (`main.py`)
Phù hợp chạy tự động, đặt lịch định kỳ (cron job / task scheduler) hoặc kiểm tra nhanh không cần mở trình duyệt:

1. Đặt 4 file Excel dữ liệu cần kiểm tra vào thư mục `01_DAU_VAO/`:
   - `Bang_Luong_Tong_CBCNV.xlsx`
   - `To_Khai_Thue_PR025.xlsx`
   - `To_Khai_Thue_PR027.xlsx`
   - `So_Cai_TK_3335.xlsx`
2. Chạy lệnh:
   ```bash
   python main.py
   ```
3. Kết quả kiểm dò sẽ được in tóm tắt trực tiếp trên màn hình console, đồng thời hệ thống tự động xuất file Excel hoàn chỉnh vào thư mục `03_DAU_RA/Bao_Cao_Kiem_Do_Thue_TNCN_{ky}_{timestamp}.xlsx`.

---

### 3.3. Tạo lại bộ dữ liệu mẫu chuẩn (Mock Data Generator)
Nếu cần tái tạo lại bộ 4 file Excel mẫu khớp 100% để phục vụ demo hoặc đào tạo nhân sự mới:

```bash
python scripts/generate_sample_data.py
```

---

## 4. QUY CÁCH DỮ LIỆU ĐẦU VÀO (TẤT CẢ LÀ FILE .XLSX)

Hệ thống tích hợp bộ lọc từ khóa tự động tại `config/mapping_rules.json`, hỗ trợ nhận diện linh hoạt các cách đặt tên cột khác nhau:

| STT | Loại File | Các Cột / Chỉ Tiêu Nhận Diện | Ghi Chú Nghiệp Vụ |
| :---: | :--- | :--- | :--- |
| **1** | **Bảng lương tổng CBCNV** | Mã NV, Họ và tên, Phòng ban, Tình trạng cư trú, Lương tháng, Lương quý, Thu nhập khác, Phụ cấp không thuế, Trừ hưu trí/BH, Thuế lũy tiến, Thu nhập 10%, Thuế 10% | Bao gồm toàn bộ nguồn thu nhập và các khoản trích theo lương |
| **2** | **Tờ khai Thuế PR025** | Chỉ tiêu `[21]`, `[22]`, `[24]`, `[25]`, `[27]`, `[31]`, `[33]` | Hỗ trợ cả mẫu bảng dọc (Chỉ tiêu, Số tiền) và bảng ngang |
| **3** | **Tờ khai Thuế PR027/PR028** | Thuế lũy tiến, Thuế 10%, Thuế khấu trừ / Tổng thuế | Tự động loại trừ dòng tổng cộng để tránh tính trùng lặp |
| **4** | **Sổ cái TK 3335** | Ngày HT, Số chứng từ, Diễn giải, TK Nợ, TK Có, Phát sinh Nợ, Phát sinh Có | Phát sinh Có phản ánh số thuế trích nộp; Nợ phản ánh số đã nộp vào NSNN |

---

## 5. LOGIC NGHIỆP VỤ & BÁO CÁO ĐẦU RA (4 SHEET EXCEL)

File Excel kết quả được xây dựng bằng thư viện `openpyxl`, thiết kế chuẩn mực theo nguyên tắc kế toán: Header màu xanh Navy `#1F4E78`, font chữ trắng in đậm, viền ô toàn bảng, số tiền định dạng `#,##0`, các ô Chênh lệch tự động tô xanh khi Khớp và tô đỏ khi Lệch:

### Sheet 1: `00_Tong_Ket` (Tổng Hợp Các Nguồn Thu Nhập)
- **Công thức:** `Tổng thu nhập = Lương tháng + Lương quý/bổ sung + Các khoản thu nhập khác`.
- **Bảng 1:** Tổng hợp theo từng phòng ban / bộ phận và dòng Tổng cộng toàn công ty.
- **Bảng 2:** Danh sách chi tiết thu nhập, phụ cấp miễn thuế và thuế của từng nhân sự.

### Sheet 2: `3.1_TN_Chiu_Thue` (Đối Chiếu Chỉ Tiêu [24] PR025)
- **Công thức:** `TN Chịu Thuế (Cư trú) = Tổng thu nhập - Trừ hưu trí - Phụ cấp không thuế`.
- **Kiểm dò:** `Chênh lệch = TN Chịu thuế theo Bảng lương - Chỉ tiêu [24] trên PR025`.
- **Trạng thái:** Báo `KHỚP` nếu Chênh lệch = 0 VND; Báo `LỆCH` nếu Chênh lệch khác 0.
- Đi kèm danh sách chi tiết cách tính TNCT của từng cá nhân cư trú có HĐLĐ.

### Sheet 3: `4.1_CNV_Co_Thue` (Đối Chiếu Chỉ Tiêu [27] PR025)
- **Dành riêng:** Nhóm cá nhân cư trú có phát sinh thuế TNCN khấu trừ (`Thuế lũy tiến > 0` hoặc `Thuế 10% > 0`).
- **Công thức:** `TNCT Nhóm Có Thuế = Tổng TN Nhóm có thuế - Phụ cấp không thuế + Thu nhập 10%`.
- **Kiểm dò:** `Chênh lệch = TNCT Nhóm có thuế (Bảng lương) - Chỉ tiêu [27] trên PR025`.
- Đi kèm danh sách chi tiết từng nhân sự thuộc nhóm phát sinh thuế.

### Sheet 4: `5.1_Thue_Khau_Tru` (Đối Chiếu 3 Bên: Lương - Tờ Khai - Sổ Cái 3335)
- **Tổng thuế tính toán:** `Thuế CNV lũy tiến + Thuế 10%`.
- **Đối chiếu 3 chiều:**
  1. Số thuế theo Bảng lương.
  2. Số thuế theo Tờ khai Thuế (PR027 / PR025 `[31]`).
  3. Phát sinh Có Tài khoản 3335 (Số thuế trích phải nộp vào NSNN trên sổ kế toán).
- **Kiểm dò đa chiều:**
  - `Lệch Bảng lương vs Tờ khai = Thuế Bảng lương - Thuế Tờ khai`.
  - `Lệch Bảng lương vs Sổ cái = Thuế Bảng lương - Phát sinh Có TK 3335`.
  - `Lệch Tờ khai vs Sổ cái = Thuế Tờ khai - Phát sinh Có TK 3335`.
  - `Chỉ số kiểm dò = Thuế tính toán - Thuế tờ khai - Phát sinh Có TK 3335`.

---

## 6. TÙY BIẾN CẤU HÌNH

### Tùy chỉnh thông tin công ty và kỳ báo cáo (`config/config.yaml`):
```yaml
company_name: "CÔNG TY CỔ PHẦN CÔNG NGHỆ VÀ DỊCH VỤ DOANH NGHIỆP"
tax_code: "0109999999"
report_period: "Tháng 09/2026"
thresholds:
  warning_tolerance: 1000  # Cho phép dung sai làm tròn 1,000 VND
```

### Thêm từ đồng nghĩa cho tiêu đề cột (`config/mapping_rules.json`):
Nếu file Excel của doanh nghiệp sử dụng các tên cột đặc thù, bạn chỉ cần bổ sung từ khóa vào file JSON mà không cần sửa code Python.

---

## 7. BẢO MẬT & BẢN QUYỀN

- Toàn bộ dữ liệu xử lý trực tiếp trên RAM hoặc máy tính cục bộ, không gửi thông tin nhân sự hay tiền lương qua bên thứ 3.
- Bản quyền thuộc về đội ngũ phát triển. Đảm bảo an toàn, bảo mật và chính xác tuyệt đối.
