"""
Hệ thống kiểm dò dữ liệu Thuế TNCN (Personal Income Tax Reconciliation)
Bao gồm:
- data_loader: Đọc và chuẩn hóa dữ liệu từ file path hoặc Streamlit UploadedFile (BytesIO)
- reconciler: Logic tính toán, đối chiếu số liệu và kiểm dò sai lệch
- report_builder: Xuất file Excel định dạng báo cáo chuẩn kế toán với openpyxl
"""

__version__ = "1.0.0"
__author__ = "Antigravity Team"
