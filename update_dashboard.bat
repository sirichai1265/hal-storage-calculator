@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === คำนวณไฟล์ใน INPUT ===
python storage_calc.py
echo.
echo === อัปเดต Dashboard ออนไลน์ (GitHub) ===
git add -A
git commit -m "Update storage data %date% %time%" >nul 2>&1 && echo บันทึกการเปลี่ยนแปลงแล้ว || echo ไม่มีข้อมูลใหม่
git push origin main
echo.
echo Dashboard: https://sirichai1265.github.io/hal-storage-calculator/DASHBOARD.html
echo (GitHub ใช้เวลาประมาณ 1 นาทีในการอัปเดตหน้าเว็บ)
pause
