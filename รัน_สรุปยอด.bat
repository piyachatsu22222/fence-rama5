@echo off
rem ดับเบิลคลิกไฟล์นี้เพื่อรันสรุปยอด PO — ไม่ต้องพิมพ์คำสั่งเอง
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo  สรุปยอด PO - งานรั้ว พระราม 5
echo ============================================
python "สรุปยอด_PO_รั้วพระราม5.py" %*
if errorlevel 1 (
  echo.
  echo [!] รันไม่สำเร็จ - ดูข้อความด้านบน
  echo     ถ้าไม่พบ openpyxl ให้พิมพ์:  pip install -r requirements.txt
)
echo.
pause
