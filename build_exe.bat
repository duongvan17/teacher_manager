@echo off
REM Build Windows .exe (one-file, no console). Yêu cầu: Python + pip install pyinstaller pillow
REM Chạy: double-click hoặc gõ build_exe.bat trong cmd.

echo === Don dep build cu ===
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist TSQ_Teacher_Manager.spec del TSQ_Teacher_Manager.spec

echo === Tao app.ico tu avt.png ===
python -c "from PIL import Image; Image.open('avt.png').save('app.ico', format='ICO', sizes=[(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)])"

echo === Build exe ===
pyinstaller --noconfirm --onefile --windowed ^
  --name "TSQ_Teacher_Manager" ^
  --icon=app.ico ^
  --collect-all customtkinter ^
  --collect-data pdfplumber ^
  --collect-data openpyxl ^
  --hidden-import=PIL ^
  --hidden-import=pandas ^
  --hidden-import=tkinter ^
  --hidden-import=tkinter.ttk ^
  --add-data "extractor.py;." ^
  --add-data "app.ico;." ^
  app.py

echo === Copy data files ===
if exist schedule.xlsx copy /y schedule.xlsx dist\
if exist "danh sach k8.xlsx" copy /y "danh sach k8.xlsx" dist\
if exist "danh sách k8.xlsx" copy /y "danh sách k8.xlsx" dist\
if exist Document xcopy /e /i /y Document dist\Document

echo.
echo ===========================================
echo  Hoan tat. File exe: dist\TSQ_Teacher_Manager.exe
echo  Copy ca thu muc dist\ cho khach.
echo ===========================================
pause
