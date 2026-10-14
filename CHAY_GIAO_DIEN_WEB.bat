@echo off
chcp 65001 > nul
title HỆ THỐNG KIỂM DÒ THUẾ TNCN - GIAO DIỆN WEB
echo =====================================================================
echo    HỆ THỐNG KIỂM DÒ THUẾ TNCN - CÔNG TY ĐIỆN LỰC
echo    Đang khởi động giao diện Web Streamlit...
echo =====================================================================
echo.
echo Trình duyệt web sẽ tự động mở tại địa chỉ: http://localhost:8501
echo (Để tắt ứng dụng, anh chỉ cần đóng cửa sổ màu đen này lại)
echo.
streamlit run app.py
pause
