@echo off
chcp 65001 >nul
title Day code len GitHub - Hung Phat AI Agent
echo ========================================================
echo    DANG TIEN HANH DAY CODE MOI NHAT LEN GITHUB...
echo ========================================================
echo.
cd /d "c:\Users\Admin\Documents\Hướng dẫn Agent"
set "PATH=C:\Program Files\Git\cmd;%PATH%"

git push -f -u origin main

echo.
echo ========================================================
echo  NEU DA HIEN CHU "main -> main" LA DA THANH CONG 100%!
echo  Vercel se tu dong cap nhat lai trang web trong 1 phut.
echo ========================================================
pause
