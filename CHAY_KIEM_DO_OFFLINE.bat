@echo off
chcp 65001 > nul
title HỆ THỐNG KIỂM DÒ THUẾ TNCN - CHẠY OFFLINE
echo =====================================================================
echo    HỆ THỐNG KIỂM DÒ THUẾ TNCN - CÔNG TY ĐIỆN LỰC
echo    Đang thực hiện kiểm dò dữ liệu từ thư mục 02_THAM_KHAO...
echo =====================================================================
echo.
python main.py 02_THAM_KHAO
echo.
echo =====================================================================
echo    Kiểm dò hoàn tất! Đang mở thư mục kết quả 03_DAU_RA...
echo =====================================================================
start "" "03_DAU_RA"
pause
