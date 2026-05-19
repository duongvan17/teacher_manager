@echo off
REM Build Windows .exe (one-file, no console).
REM Yeu cau: pip install pyinstaller pillow customtkinter pandas openpyxl pdfplumber python-docx
REM Chay: double-click hoac go build_exe.bat trong cmd.

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
  --hidden-import=docx ^
  --collect-data docx ^
  --exclude-module matplotlib ^
  --exclude-module scipy ^
  --exclude-module pytest ^
  --exclude-module IPython ^
  --exclude-module notebook ^
  --add-data "extractor.py;." ^
  --add-data "app.ico;." ^
  app.py

echo === Copy data files ===
if exist schedule.xlsx copy /y schedule.xlsx dist\
if exist "danh sach k8.xlsx" copy /y "danh sach k8.xlsx" dist\
if exist "danh sách k8.xlsx" copy /y "danh sách k8.xlsx" dist\
if exist Document xcopy /e /i /y Document dist\Document
if exist "BÁO CÁO TUẦN 1.8.docx" copy /y "BÁO CÁO TUẦN 1.8.docx" dist\
if exist "BÁO CÁO HUẤN LUYỆN NGÀY 22.1.xlsx" copy /y "BÁO CÁO HUẤN LUYỆN NGÀY 22.1.xlsx" dist\

echo.
echo ===========================================
echo  Hoan tat. File exe: dist\TSQ_Teacher_Manager.exe
echo  Copy ca thu muc dist\ cho khach.
echo ===========================================
pause
