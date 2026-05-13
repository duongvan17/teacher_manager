import customtkinter  # This tells Python what 'customtkinter' is
import tkinter
import sys as _sys
for _stream in (_sys.stdout, _sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except Exception:
        pass
from extractor import run_extraction # Linking to your other file
import customtkinter as ctk
import pandas as pd
from tkinter import messagebox, filedialog
import os
from datetime import datetime
import pdfplumber
import re
from openpyxl.styles import Font, Alignment
from datetime import datetime, timedelta
today = datetime.now().strftime("%d/%m/%Y")
from openpyxl.styles import Font, Alignment, Border, Side
import pandas as pd
from openpyxl.styles import Alignment, Font, Border
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from openpyxl.styles import Alignment, Font, Border, Side
import sys
import glob
import customtkinter as ctk
from PIL import Image
import os
import customtkinter as ctk
from PIL import Image
import shutil
from tkinter import filedialog
import pandas as pd
from tkinter import ttk
import subprocess
import subprocess
import os
import subprocess
import platform
def load_documents_view(self, base_directory):
    # CHANGE THIS to your actual scrollable frame variable
    target_frame = self.main_scrollable_frame 

    # Clear existing UI
    for widget in target_frame.winfo_children():
        widget.destroy()

    if not os.path.exists(base_directory):
        print(f"DEBUG: Cannot find path: {base_directory}")
        return

    print(f"DEBUG: Scanning main folder: {base_directory}")

    for item_name in sorted(os.listdir(base_directory)):
        if item_name.startswith('.') or item_name.startswith('~$'):
            continue

        item_path = os.path.join(base_directory, item_name)

        if os.path.isdir(item_path):
            print(f"DEBUG: Found Sub-folder -> {item_name}")
            
            # 1. Create the Folder Frame
            folder_frame = ctk.CTkFrame(target_frame, fg_color="transparent")
            folder_frame.pack(fill="x", pady=5, padx=5)
            ctk.CTkLabel(folder_frame, text=f"📁 {item_name}", font=("Arial", 14, "bold")).pack(anchor="w", padx=5)

            # 2. Scan inside the Sub-folder
            for sub_item in sorted(os.listdir(item_path)):
                if sub_item.startswith('.') or sub_item.startswith('~$'):
                    continue
                    
                sub_item_path = os.path.join(item_path, sub_item)

                if os.path.isfile(sub_item_path):
                    print(f"DEBUG:   Found File -> {sub_item}")
                    
                    # 3. Create File Frame INSIDE Folder Frame
                    file_frame = ctk.CTkFrame(folder_frame)
                    file_frame.pack(fill="x", pady=2, padx=(30, 5)) 
                    ctk.CTkLabel(file_frame, text=f"📄 {sub_item}").pack(side="left", padx=10, pady=5)



def refresh_files(self):
    for widget in self.file_container.winfo_children():
        widget.destroy()

    files = get_all_files("Document")

    for file in files:
        self.create_file_row(file["relative"], file["path"])
def open_file(path):
    if platform.system() == "Darwin":  # macOS
        subprocess.call(["open", path])
    elif platform.system() == "Windows":
        os.startfile(path)
    else:  # Linux
        subprocess.call(["xdg-open", path])
def get_all_files(folder_path):
    all_files = []

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            full_path = os.path.join(root, file)

            # Optional: get relative path for display
            relative_path = os.path.relpath(full_path, folder_path)

            all_files.append({
                "name": file,
                "path": full_path,
                "relative": relative_path
            })

    return all_files

# Hàm quan trọng: Giúp file .exe xác định đúng thư mục đang đứng
def get_base_path():
    if getattr(sys, 'frozen', False):
        # Nếu là file .exe, lấy đường dẫn thư mục chứa file .exe
        return os.path.dirname(sys.executable)
    # Nếu đang chạy code .py trong VS Code
    return os.path.dirname(os.path.abspath(__file__))

def auto_update_schedule(schedule_file=None):
    # 1. Lấy ngày hiện tại
    today = datetime.now()
    date_str = today.strftime('%Y-%m-%d')
    vn_date = today.strftime('%d/%m/%Y')

    base_path = get_base_path()

    # 2. Lấy đường dẫn file kế hoạch tháng (theo config nếu có)
    file_path = None
    if schedule_file:
        candidate = schedule_file if os.path.isabs(schedule_file) \
            else os.path.join(base_path, schedule_file)
        if os.path.exists(candidate):
            file_path = candidate

    if file_path is None:
        # fallback: tìm file schedule*.xlsx trong thư mục
        search_pattern = os.path.join(base_path, "schedule*.*")
        files = glob.glob(search_pattern)
        if files:
            file_path = files[0]

    if not file_path:
        print(f"[auto_update_schedule] Không tìm thấy file kế hoạch tháng. "
              f"Vào tab Cài đặt để chọn lại.")
        return

    try:
        # 3. Đọc dữ liệu
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path, skiprows=3)
        else:
            df = pd.read_excel(file_path, skiprows=3)

        df.columns = [str(col).replace('\n', ' ').strip() for col in df.columns]
        slot_col = [c for c in df.columns if 'Cặp' in c and 'tiết' in c][0]
        
        # Tìm cột ngày hôm nay
        date_column = next((c for c in df.columns if c.startswith(date_str)), None)

        if not date_column:
            print(f"[auto_update_schedule] Hôm nay ({vn_date}) không có lịch dạy trong file nguồn.")
            return

        # 4. Xử lý gộp ô và lọc dữ liệu
        df['HỌ VÀ TÊN'] = df['HỌ VÀ TÊN'].ffill()
        df['MÔN HỌC'] = df['MÔN HỌC'].ffill()

        wb = Workbook()
        ws = wb.active
        
        # Định dạng style
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))
        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Tiêu đề và Header
        ws.merge_cells('A1:F1')
        ws['A1'] = f"KẾ HOẠCH GIẢNG DẠY NGÀY {today.day} THÁNG {today.month} NĂM {today.year}"
        ws['A1'].font = Font(bold=True, size=14)
        ws['A1'].alignment = center_align

        headers = ["Họ và tên", "môn học", "1 - 2", "3 - 4", "5 - 6", "7 - 8"]
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=i, value=h)
            cell.font = Font(bold=True)
            cell.border = thin_border
            cell.alignment = center_align

        # 5. Ghi dữ liệu giáo viên
        current_row = 5
        section_exact = {
            "SÁNG", "CHIỀU", "BUỔI SÁNG", "BUỔI CHIỀU",
            "TỔNG", "CỘNG", "TỔNG CỘNG", "TỔNG SỐ",
            "QUÂN SỰ", "QUỐC TẾ", "CÔNG AN",
            "QUÂN SỰ + QUỐC TẾ", "QUÂN SỰ + QUỐC TẾ:",
        }
        section_contains = ("THỐNG KÊ", "GHI CHÚ")

        def _is_section(name):
            up = str(name).strip().upper()
            if len(up) < 2:
                return True
            if up in section_exact:
                return True
            if any(k in up for k in section_contains):
                return True
            if up.endswith(":"):
                return True
            first = up.split()[0] if up.split() else ""
            if first in {"TỔNG", "CỘNG"}:
                return True
            return False

        for teacher in df['HỌ VÀ TÊN'].dropna().unique():
            if _is_section(teacher):
                continue

            teacher_df = df[df['HỌ VÀ TÊN'] == teacher]
            is_first = True

            for subject in teacher_df['MÔN HỌC'].unique():
                if _is_section(subject): continue

                sub_df = teacher_df[teacher_df['MÔN HỌC'] == subject]
                slots = {"1 - 2": "", "3 - 4": "", "5 - 6": "", "7 - 8": ""}
                
                for _, row in sub_df.iterrows():
                    s = str(row[slot_col]).strip()
                    if "7 - 9" in s: s = "7 - 8"
                    if s in slots: slots[s] = row[date_column]

                row_vals = [teacher if is_first else "", subject, slots["1 - 2"], slots["3 - 4"], slots["5 - 6"], slots["7 - 8"]]
                for idx, val in enumerate(row_vals, 1):
                    cell = ws.cell(row=current_row, column=idx, value=val)
                    cell.border = thin_border
                    cell.alignment = center_align
                is_first = False
                current_row += 1
            current_row += 1

        # Tự động căn chỉnh cột
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 15
        for col in ['C','D','E','F']: ws.column_dimensions[col].width = 18

        # 6. Lưu file cùng thư mục với EXE
        output_path = os.path.join(base_path, f"KeHoach_Ngay_{date_str}.xlsx")
        wb.save(output_path)
        print(f"[auto_update_schedule] Đã cập nhật lịch ngày {vn_date}. File: {output_path}")

    except Exception as e:
        print(f"[auto_update_schedule] Lỗi hệ thống: {e}")

# another
WEEKDAY_MAP = {
    0: "H",
    1: "B",  # Tuesday
    2: "T",  # Wednesday
    3: "N",  # Thursday
    4: "S",  # Friday
    5: "By", # Saturday
    6: "CN"  
}
TIME_WINDOWS = {
    "1-2": ("06:45", "08:15"),
    "3-4": ("08:25", "09:55"),
    "5-6": ("10:05", "11:25"),
    "7-8": ("13:45", "15:05")
}

def check_teaching_status(period_key):
    """Returns '1-2' if teaching, else 'He is out of class'"""
    now = datetime.now().time()
    clean_key = period_key.replace(" ", "") # Handles "1 - 2"
    
    if clean_key in TIME_WINDOWS:
        start_str, end_str = TIME_WINDOWS[clean_key]
        start = datetime.strptime(start_str, "%H:%M").time()
        end = datetime.strptime(end_str, "%H:%M").time()
        
        if start <= now <= end:
            return f"Teaching: {period_key}"
            
    return "He is out of class"

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")
COLORS = {
    "bg": "#F5F7F9",          # Nền chính (Xám nhẹ)
    "sidebar": "#FFFFFF",      # Sidebar (Trắng)
    "card": "#FFFFFF",         # Thẻ nội dung (Trắng)
    "accent": "#2563EB",       # Xanh dương chủ đạo
    "text": "#1E293B",         # Chữ chính (Đen xanh)
    "text_dim": "#64748B",     # Chữ phụ (Xám)
    "hover": "#F1F5F9",        # Màu khi di chuột qua
    "border": "#E2E8F0",       # Màu viền mảnh
    "success": "#10B981",      # Xanh lá (Dùng cho trạng thái sẵn sàng)
    "warning": "#F59E0B",      # Vàng cam
    "error": "#EF4444",        # Đỏ
    "purple": "#8B5CF6",       # Tím (Dự phòng cho các nút cũ)
    "orange": "#F97316",       # Cam (Dự phòng cho các nút cũ)
    "sidebar_dark": "#0F172A", # Sidebar tối (tùy chọn)
}

import json

class AppConfig:
    DEFAULTS = {
        "teacher_file": "danh sách k8.xlsx",
        "schedule_file": "schedule.xlsx",
        "document_folder": "Document",
        "user_name": "Giảng viên",
        "user_email": "",
        "auto_update_schedule": True,
        "appearance": "light",
    }
    PATH = "config.json"

    @classmethod
    def load(cls):
        try:
            with open(cls.PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            merged = dict(cls.DEFAULTS)
            merged.update(data or {})
            return merged
        except Exception:
            return dict(cls.DEFAULTS)

    @classmethod
    def save(cls, data):
        try:
            with open(cls.PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"[config] save error: {e}")
            return False


def clean_numeric_text(val):
    """Chuyển '1.0' -> '1', '2.5' giữ nguyên, NaN/None -> ''."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return f"{val:g}"
    s = str(val).strip()
    if s.lower() == "nan":
        return ""
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except (ValueError, TypeError):
        pass
    return s


class DocumentFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        
        # Configure grid
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Create Treeview
        self.tree = ttk.Treeview(self, columns=("filename"), show="headings")
        self.tree.heading("filename", text="Tên tài liệu (Double click to open)")
        self.tree.column("filename", anchor="w", width=400)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

        # Bind double click
        self.tree.bind("<Double-1>", self.open_file)

        self.refresh_list()

    def refresh_list(self):
        path = "document"
        if not os.path.exists(path):
            os.makedirs(path)
            
        for file in os.listdir(path):
            if not file.startswith('.'): # Ignore hidden mac files
                self.tree.insert("", "end", values=(file,))

    def open_file(self, event):
        selected_item = self.tree.selection()[0]
        name = self.tree.item(selected_item, "values")[0]
        full_path = os.path.join("document", name)
        
        # Mac specific command to open Word/PDF
        subprocess.call(["open", full_path])
class TeacherCard(ctk.CTkFrame):
    def __init__(self, master, name, period, detail):
        super().__init__(master)
        
        self.period = period
        
        # Teacher Name & Detail (e.g., "td+ b6,7/c2")
        self.info_label = ctk.CTkLabel(self, text=f"{name} ({detail})", font=("Arial", 13))
        self.info_label.pack(side="left", padx=10)
        
        # Status Notification Label
        self.status_label = ctk.CTkLabel(self, text="", font=("Arial", 12, "bold"))
        self.status_label.pack(side="right", padx=10)
     
        self.update_status()  
class TeacherDetailWindow(ctk.CTkToplevel):

    def __init__(self, parent, data):
        super().__init__(parent)
        # Lấy tên giảng viên làm tiêu đề
        name = str(data.get('HỌ VÀ TÊN', 'CHI TIẾT')).upper()
        self.title(f"Thông tin: {name}")
        self.geometry("550x650")
        self.attributes("-topmost", True)  # Luôn hiện trên cùng
        self.configure(fg_color="#F1F5F9")

        # Container chính
        container = ctk.CTkFrame(self, fg_color="white", corner_radius=15)
        container.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(container, text="HỒ SƠ GIẢNG VIÊN", font=("Arial", 16, "bold"), text_color="#64748B").pack(pady=(15, 0))
        ctk.CTkLabel(container, text=name, font=("Arial", 24, "bold"), text_color="#1E40AF").pack(pady=(0, 20))

        # Vùng cuộn thông tin
        info_scroll = ctk.CTkScrollableFrame(container, fg_color="transparent")
        info_scroll.pack(fill="both", expand=True, padx=10)

        # Tự động quét qua tất cả các cột dữ liệu
        for key, value in data.items():
            if "UNNAMED" in str(key).upper():
                continue
            display_val = clean_numeric_text(value)
            if not display_val:
                continue

            row_frame = ctk.CTkFrame(info_scroll, fg_color="#F8FAFC", corner_radius=8)
            row_frame.pack(fill="x", pady=3)

            ctk.CTkLabel(row_frame, text=str(key), font=("Arial", 12, "bold"),
                         text_color="#475569", width=160, anchor="w"
                         ).pack(side="left", padx=15, pady=10)

            ctk.CTkLabel(row_frame, text=display_val, font=("Arial", 13),
                         text_color="#1E293B", wraplength=280, justify="left"
                         ).pack(side="left", fill="x", expand=True, padx=5)

        ctk.CTkButton(container, text="ĐÓNG", fg_color="#1E293B", command=self.destroy).pack(pady=20)
   
class TeacherManagerPro(ctk.CTk):
    def clear_right_frame(self):
    # Ensure self.right_frame actually exists before trying to clear it
        if hasattr(self, 'right_frame'):
            for widget in self.right_frame.winfo_children():
                widget.destroy()
        else:
            print("Error: right_frame has not been initialized yet.")
    def hide_all_frames(self):
        for name in ("dashboard_frame", "mgmt_frame", "plan_frame",
                     "month_frame", "report_frame", "report_day_frame",
                     "document_frame", "settings_frame"):
            f = getattr(self, name, None)
            if f is not None:
                f.pack_forget()


    def build_tree(self, folder_path):
        def is_hidden(name):
            return name.startswith(".") or name.startswith("~$") or name.lower() == "thumbs.db"

        tree = {}

        for root, dirs, files in os.walk(folder_path):
            dirs[:] = [d for d in dirs if not is_hidden(d)]

            rel_path = os.path.relpath(root, folder_path)
            parts = rel_path.split(os.sep) if rel_path != "." else []

            current = tree
            for part in parts:
                current = current.setdefault(part, {})

            for file in files:
                if is_hidden(file):
                    continue
                current[file] = None

        return tree
    def render_documents(self):
        if not hasattr(self, "doc_tree"):
            return
        for item in self.doc_tree.get_children():
            self.doc_tree.delete(item)
        self._doc_paths.clear()

        folder = self.config_data.get("document_folder", "Document")
        self.doc_status.configure(text=f"Thư mục: {folder}")

        if not os.path.exists(folder):
            self.doc_tree.insert("", "end",
                text=f"  Không tìm thấy '{folder}'. Mở tab Cài đặt để chọn lại.")
            if hasattr(self, "doc_count"):
                self.doc_count.configure(text="")
            return

        tree = self.build_tree(folder)
        if not tree:
            self.doc_tree.insert("", "end", text="  Thư mục trống")
            if hasattr(self, "doc_count"):
                self.doc_count.configure(text="0 tài liệu")
            return

        total = self._populate_doc_tree("", tree, folder)
        if hasattr(self, "doc_count"):
            self.doc_count.configure(text=f"{total} tài liệu")

    def _populate_doc_tree(self, parent, tree, base_path):
        total = 0
        for name, content in sorted(tree.items(),
                                     key=lambda x: (not isinstance(x[1], dict),
                                                     x[0].lower())):
            full = os.path.join(base_path, name)
            if isinstance(content, dict):
                is_top = (parent == "")
                item = self.doc_tree.insert(parent, "end",
                                             text=f"  📁  {name}",
                                             open=is_top)
                total += self._populate_doc_tree(item, content, full)
            else:
                ext = os.path.splitext(name)[1].lower()
                icon = {
                    ".docx": "📝", ".doc": "📝",
                    ".xlsx": "📊", ".xls": "📊", ".csv": "📊",
                    ".pdf": "📕",
                    ".pptx": "📽", ".ppt": "📽",
                    ".txt": "📃", ".md": "📃",
                    ".png": "🖼", ".jpg": "🖼", ".jpeg": "🖼",
                }.get(ext, "📄")
                item = self.doc_tree.insert(parent, "end",
                                             text=f"  {icon}  {name}")
                self._doc_paths[item] = full
                total += 1
        return total

    def _on_doc_tree_activate(self, event=None):
        sel = self.doc_tree.selection()
        if not sel:
            return
        path = self._doc_paths.get(sel[0])
        if not path:
            return
        try:
            if os.name == "nt":
                os.startfile(path)
            else:
                subprocess.call(["open", path])
        except Exception as e:
            self.doc_status.configure(text=f"Lỗi mở file: {e}")
    def render_tree(self, parent, tree, base_path="", level=0):
        for name, content in sorted(tree.items(), key=lambda x: (not isinstance(x[1], dict), x[0].lower())):
            full_path = os.path.join(base_path, name)
            if isinstance(content, dict):
                self._render_folder_node(parent, name, content, full_path, level)
            else:
                self._render_file_node(parent, name, full_path, level)

    def _render_folder_node(self, parent, name, content, full_path, level):
        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.pack(fill="x", padx=0, pady=0, anchor="w")

        header = ctk.CTkFrame(container, fg_color=("#E2E8F0", "#1F2937"), corner_radius=6, height=36)
        header.pack(fill="x", padx=(10 + level * 22, 10), pady=2)
        header.pack_propagate(False)

        child_frame = ctk.CTkFrame(container, fg_color="transparent")
        state = {"open": False}

        arrow = ctk.CTkLabel(header, text="▸", font=("Segoe UI", 12, "bold"), width=18, cursor="hand2")
        arrow.pack(side="left", padx=(10, 0))
        icon = ctk.CTkLabel(header, text="📁", font=("Segoe UI", 13), cursor="hand2")
        icon.pack(side="left", padx=(2, 4))
        label = ctk.CTkLabel(header, text=name, font=("Segoe UI", 13, "bold"), cursor="hand2", anchor="w")
        label.pack(side="left", padx=2, fill="x", expand=True)

        def toggle(event=None):
            if state["open"]:
                child_frame.pack_forget()
                arrow.configure(text="▸")
                icon.configure(text="📁")
                state["open"] = False
            else:
                child_frame.pack(fill="x", anchor="w")
                arrow.configure(text="▾")
                icon.configure(text="📂")
                state["open"] = True

        for w in (header, arrow, icon, label):
            w.bind("<Button-1>", toggle)

        self.render_tree(child_frame, content, full_path, level + 1)

    def _render_file_node(self, parent, name, full_path, level):
        normal_color = ("#F8FAFC", "#111827")
        hover_color = ("#DBEAFE", "#1E3A8A")

        row = ctk.CTkFrame(parent, height=34, corner_radius=6, fg_color=normal_color)
        row.pack(fill="x", padx=(30 + level * 22, 10), pady=2)
        row.pack_propagate(False)

        icon = ctk.CTkLabel(row, text="📄", font=("Segoe UI", 13), cursor="hand2")
        icon.pack(side="left", padx=(10, 4))
        label = ctk.CTkLabel(row, text=name, font=("Segoe UI", 12), cursor="hand2", anchor="w")
        label.pack(side="left", padx=0, fill="x", expand=True)

        def on_click(event=None):
            self.open_document(full_path)

        def on_enter(event=None):
            row.configure(fg_color=hover_color)

        def on_leave(event=None):
            row.configure(fg_color=normal_color)

        for w in (row, icon, label):
            w.bind("<Button-1>", on_click)
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)

    def create_teacher_card(self, row):
        """Hàm phụ tạo từng dòng giảng viên"""
        card = ctk.CTkFrame(self.mgmt_scroll, fg_color="white", height=55, corner_radius=10, 
                            border_width=1, border_color="#E2E8F0")
        card.pack(fill="x", pady=2, padx=(20, 10))
        card.pack_propagate(False)

        # Hiển thị tên
        name_label = ctk.CTkLabel(card, text=row.get('HỌ VÀ TÊN', 'N/A'), font=("Arial", 14, "bold"))
        name_label.pack(side="left", padx=20)
        
        # Hiển thị cấp bậc (nếu có)
        rank = row.get('CẤP BẬC', '')
        if rank and str(rank) != "nan":
            ctk.CTkLabel(card, text=f"({rank})", font=("Arial", 12), text_color="#64748B").pack(side="left")

        # Nút bấm xem chi tiết
        # Truyền toàn bộ dữ liệu của dòng (row) vào cửa sổ mới
        btn = ctk.CTkButton(card, text="XEM CHI TIẾT", width=100, height=32, 
                            fg_color="#2563EB", hover_color="#1D4ED8",
                            command=lambda r=row.to_dict(): TeacherDetailWindow(self, r))
        btn.pack(side="right", padx=15)
   
    def show_document_frame(self):
        self.hide_all_frames()
        self.document_frame.pack(fill="both", expand=True)
        self.set_active_nav("document")
        self.render_documents()
  
    def start_live_sync(self):
        if self.mgmt_data and self.plan_path: # Ensure files are linked
            # Update the data in memory
           
            # Re-render the UI
            self.render_mgmt()
            
        # Refresh every 60 seconds to keep "Real Time"
        self.after(60000, self.start_live_sync)
    def process_military_plan_with_calendar(file_path):
        teaching_data = []
        
        # Get current date info
        now = datetime.now()
        today_num = str(now.day).zfill(2) # "31"
        today_char = WEEKDAY_MAP[now.weekday()] # "B" (for Tuesday, March 31)

        try:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if not table or len(table) < 2: continue
                    
                    # Row 0 is "Ngày" (01, 02, 03...)
                    # Row 1 is "Thứ" (T, N, S...)
                    days_row = table[0]
                    weekdays_row = table[1]
                    
                    target_col = None
                    for idx in range(len(days_row)):
                        # Check if column matches today's Date AND today's Weekday letter
                        if (days_row[idx] == today_num and 
                            weekdays_row[idx] == today_char):
                            target_col = idx
                            break
                    
                    if target_col is None: continue # Day not found on this page

                    # Process teachers in the rows below
                    for row in table[2:]:
                        # row[2] = Name, row[4] = Period (Tiết)
                        name = " ".join(str(row[2]).split()) if row[2] else None
                        period = str(row[4]).replace(" ", "") if row[4] else None
                        activity = row[target_col] # What is in today's column

                        if name and period and activity and activity.strip():
                            status = check_teaching_status(period)
                            teaching_data.append({
                                "teacher": name,
                                "period": period,
                                "detail": activity.strip(),
                                "notification": status
                            })
        except Exception as e:
            print(f"Error parsing PDF calendar: {e}")
        
        return teaching_data
    def sync_with_military_plan(self):
        # This automatically fetches today's specific assignments from the PDF[cite: 1]
        
        self.render_plan()
        # Check again every 5 minutes to see if a teacher has started a new slot
        self.after(300000, self.sync_with_military_plan)

    
    def __init__(self):
        super().__init__()
        self.title("TSQ Teacher Manager Pro")
        self.geometry("1360x860")
        self.configure(fg_color=COLORS["bg"])

        self.config_data = AppConfig.load()
        ctk.set_appearance_mode(self.config_data.get("appearance", "light"))

        self.mgmt_data = []
        self.plan_data = []

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.setup_sidebar()

        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=0, column=1, sticky="nsew", padx=18, pady=18)

        self.dashboard_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.mgmt_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.plan_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.month_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.report_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.report_day_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.document_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.settings_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")

        self.setup_dashboard_ui()
        self.setup_mgmt_ui()
        self.setup_plan_ui()
        self.setup_month_ui()
        self.setup_report_ui()
        self.setup_report_day_ui()
        self.setup_document_ui()
        self.setup_settings_ui()

        self.show_dashboard_frame()
        self.update_time()

        # Sinh file KeHoach_Ngay từ schedule (đọc theo config)
        if self.config_data.get("auto_update_schedule", True):
            try:
                auto_update_schedule(self.config_data.get("schedule_file"))
            except Exception as e:
                print(f"[init] auto_update_schedule lỗi: {e}")

        self.check_realtime_status()
        self.teacher_db = []
        self.after(100, self.auto_load_mgmt_file)
    
       
    def setup_sidebar(self):
        self.sidebar_expanded = True
        self.sidebar_w_expanded = 250
        self.sidebar_w_collapsed = 64
        self._nav_defs = [
            ("dashboard", "🏠", "Bảng điều khiển", "show_dashboard_frame"),
            ("mgmt", "👥", "Thông tin giảng viên", "show_mgmt_frame"),
            ("plan", "📅", "Kế hoạch ngày", "show_plan_frame"),
            ("month", "🗓", "Kế hoạch tháng", "show_month_frame"),
            ("report_day", "📋", "Báo cáo ngày", "show_report_day_frame"),
            ("report", "📊", "Báo cáo tuần", "show_report_frame"),
            ("document", "📂", "Môn học", "show_document_frame"),
            ("settings", "⚙", "Cài đặt", "show_settings_frame"),
        ]
        self._nav_full = {}   # key -> CTkButton (full)
        self._nav_mini = {}   # key -> CTkButton (mini)

        # === Full sidebar (expanded 250px) ===
        self.sidebar_full = ctk.CTkFrame(self, width=self.sidebar_w_expanded,
                                          corner_radius=0,
                                          fg_color=COLORS["sidebar"],
                                          border_width=0)
        self.sidebar_full.grid(row=0, column=0, sticky="nsew")
        self.sidebar_full.grid_propagate(False)
        self.sidebar_full.pack_propagate(False)

        brand = ctk.CTkFrame(self.sidebar_full, fg_color="transparent")
        brand.pack(fill="x", pady=(18, 14), padx=12)
        logo = ctk.CTkFrame(brand, fg_color=COLORS["accent"], corner_radius=8,
                            width=32, height=32)
        logo.pack(side="left")
        logo.pack_propagate(False)
        ctk.CTkLabel(logo, text="MH", font=("Arial", 13, "bold"),
                     text_color="white").pack(expand=True)
        ctk.CTkLabel(brand, text="TSQ QLGV", font=("Arial", 20, "bold"),
                     text_color=COLORS["text"]).pack(side="left", padx=10)

        ctk.CTkFrame(self.sidebar_full, height=1, fg_color=COLORS["border"]
                     ).pack(fill="x", padx=12, pady=(4, 6))

        for key, icon, label, cmd_name in self._nav_defs:
            if key == "settings":
                ctk.CTkFrame(self.sidebar_full, height=1,
                             fg_color=COLORS["border"]
                             ).pack(fill="x", padx=12, pady=(8, 6))
            btn = ctk.CTkButton(self.sidebar_full,
                                 text=f"{icon}   {label}",
                                 font=("Arial", 14),
                                 height=44,
                                 fg_color="transparent",
                                 text_color=COLORS["text"],
                                 anchor="w",
                                 hover_color=COLORS["hover"],
                                 command=getattr(self, cmd_name))
            btn.pack(pady=3, padx=10, fill="x")
            self._nav_full[key] = btn

        # Toggle full → mini
        toggle_row = ctk.CTkFrame(self.sidebar_full, fg_color="transparent")
        toggle_row.pack(side="bottom", fill="x", padx=8, pady=(6, 10))
        ctk.CTkButton(toggle_row, text="‹",
                       width=32, height=32,
                       font=("Arial", 20, "bold"),
                       fg_color="transparent",
                       text_color=COLORS["text_dim"],
                       hover_color=COLORS["hover"],
                       command=self.toggle_sidebar).pack(side="right")

        footer_full = ctk.CTkFrame(self.sidebar_full, fg_color="transparent")
        footer_full.pack(side="bottom", fill="x", padx=12, pady=(12, 0))
        self.lbl_time = ctk.CTkLabel(footer_full, text="",
                                      font=("Arial", 11),
                                      text_color=COLORS["text_dim"],
                                      justify="left", anchor="w")
        self.lbl_time.pack(fill="x")
        ctk.CTkLabel(footer_full, text="v2.0.0 · TSQCB QLGV",
                     font=("Arial", 9),
                     text_color=COLORS["text_dim"], anchor="w"
                     ).pack(fill="x", pady=(4, 0))

        # === Mini sidebar (collapsed 64px) ===
        self.sidebar_mini = ctk.CTkFrame(self, width=self.sidebar_w_collapsed,
                                          corner_radius=0,
                                          fg_color=COLORS["sidebar"],
                                          border_width=0)
        self.sidebar_mini.grid(row=0, column=0, sticky="nsew")
        self.sidebar_mini.grid_propagate(False)
        self.sidebar_mini.pack_propagate(False)

        # Logo only
        logo_mini = ctk.CTkFrame(self.sidebar_mini, fg_color=COLORS["accent"],
                                  corner_radius=8, width=32, height=32)
        logo_mini.pack(pady=(20, 16))
        logo_mini.pack_propagate(False)
        ctk.CTkLabel(logo_mini, text="MH", font=("Arial", 13, "bold"),
                     text_color="white").pack(expand=True)

        for key, icon, label, cmd_name in self._nav_defs:
            btn = ctk.CTkButton(self.sidebar_mini, text=icon,
                                 font=("Arial", 15),
                                 width=48, height=40,
                                 fg_color="transparent",
                                 text_color=COLORS["text"],
                                 hover_color=COLORS["hover"],
                                 command=getattr(self, cmd_name))
            btn.pack(pady=2, padx=8, fill="x")
            self._nav_mini[key] = btn

        # Toggle mini → full
        ctk.CTkButton(self.sidebar_mini, text="›",
                       width=32, height=32,
                       font=("Arial", 20, "bold"),
                       fg_color="transparent",
                       text_color=COLORS["text_dim"],
                       hover_color=COLORS["hover"],
                       command=self.toggle_sidebar
                       ).pack(side="bottom", pady=(6, 10))

        # Start in full mode; hide mini
        self.sidebar_mini.grid_remove()
        self.grid_columnconfigure(0, minsize=self.sidebar_w_expanded, weight=0)

    def toggle_sidebar(self):
        self.sidebar_expanded = not self.sidebar_expanded
        if self.sidebar_expanded:
            self.sidebar_mini.grid_remove()
            self.sidebar_full.grid()
            self.grid_columnconfigure(0, minsize=self.sidebar_w_expanded, weight=0)
        else:
            self.sidebar_full.grid_remove()
            self.sidebar_mini.grid()
            self.grid_columnconfigure(0, minsize=self.sidebar_w_collapsed, weight=0)

    def set_active_nav(self, key):
        for k in self._nav_full:
            full_btn = self._nav_full[k]
            mini_btn = self._nav_mini[k]
            if k == key:
                full_btn.configure(fg_color=COLORS["accent"], text_color="white",
                                    hover_color="#1D4ED8",
                                    font=("Arial", 14, "bold"))
                mini_btn.configure(fg_color=COLORS["accent"], text_color="white",
                                    hover_color="#1D4ED8")
            else:
                full_btn.configure(fg_color="transparent",
                                    text_color=COLORS["text"],
                                    hover_color=COLORS["hover"],
                                    font=("Arial", 14))
                mini_btn.configure(fg_color="transparent",
                                    text_color=COLORS["text"],
                                    hover_color=COLORS["hover"])
    def load_excel_smart(self, path, check_cols):
        try:
            raw = pd.read_excel(path, header=None)
            header_row = None
            for i, row in raw.iterrows():
                row_vals = [str(x).upper() for x in row.values]
                if any("HỌ VÀ TÊN" in str(val) for val in row_vals):
                    header_row = i
                    break
            
            if header_row is None: return None
            df = pd.read_excel(path, skiprows=header_row)
            df.columns = [str(c).strip() for c in df.columns]
            cols_joined = " ".join(df.columns).upper()
            return df.to_dict('records') if any(col.upper() in cols_joined for col in check_cols) else None
        except: return None
        
    def setup_document_ui(self):
        header = ctk.CTkFrame(self.document_frame, fg_color="transparent")
        header.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(header, text="Môn học", font=("Arial", 26, "bold"),
                     text_color=COLORS["text"]).pack(side="left")

        self.doc_count = ctk.CTkLabel(header, text="", font=("Arial", 13),
                                       text_color=COLORS["text_dim"])
        self.doc_count.pack(side="right", padx=(0, 12))

        ctk.CTkButton(header, text="Làm mới", width=100, height=36,
                      fg_color=COLORS["accent"], hover_color="#1D4ED8",
                      font=("Arial", 13, "bold"),
                      command=self.render_documents).pack(side="right")

        self.doc_status = ctk.CTkLabel(self.document_frame, text="",
                                        font=("Arial", 12),
                                        text_color=COLORS["text_dim"],
                                        anchor="w")
        self.doc_status.pack(fill="x", pady=(0, 8))

        # Card wrapper chứa treeview
        wrapper = ctk.CTkFrame(self.document_frame,
                                fg_color=COLORS["card"],
                                corner_radius=12,
                                border_width=1,
                                border_color=COLORS["border"])
        wrapper.pack(fill="both", expand=True)

        # Header strip xanh đậm như header bảng Excel
        wrap_header = ctk.CTkFrame(wrapper, fg_color=COLORS["accent"],
                                    corner_radius=0, height=46)
        wrap_header.pack(fill="x", padx=1, pady=(1, 0))
        wrap_header.pack_propagate(False)
        ctk.CTkLabel(wrap_header, text="  📁  Tài liệu môn học",
                      font=("Arial", 14, "bold"),
                      text_color="white", anchor="w"
                      ).pack(side="left", padx=14, fill="y")
        ctk.CTkLabel(wrap_header,
                      text="Bấm đôi vào tệp để mở  ",
                      font=("Arial", 11),
                      text_color="#DBEAFE"
                      ).pack(side="right", padx=14)

        container = ctk.CTkFrame(wrapper, fg_color="white",
                                  corner_radius=0, border_width=0)
        container.pack(fill="both", expand=True, padx=1, pady=(0, 1))
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        try:
            style.theme_use("default")
        except Exception:
            pass
        style.configure("Doc.Treeview",
                        rowheight=38,
                        font=("Segoe UI", 13),
                        background="white",
                        fieldbackground="white",
                        foreground=COLORS["text"],
                        borderwidth=0)
        style.configure("Doc.Treeview.Heading",
                        font=("Segoe UI", 13, "bold"))
        style.map("Doc.Treeview",
                  background=[("selected", "#DBEAFE")],
                  foreground=[("selected", COLORS["accent"])])

        self.doc_tree = ttk.Treeview(container, style="Doc.Treeview",
                                      show="tree", selectmode="browse")
        vsb = ttk.Scrollbar(container, orient="vertical",
                             command=self.doc_tree.yview)
        self.doc_tree.configure(yscrollcommand=vsb.set)
        self.doc_tree.grid(row=0, column=0, sticky="nsew",
                            padx=(2, 0), pady=2)
        vsb.grid(row=0, column=1, sticky="ns", padx=(0, 2), pady=2)

        self.doc_tree.column("#0", width=900, stretch=True)
        self.doc_tree.bind("<Double-1>", self._on_doc_tree_activate)
        self.doc_tree.bind("<Return>", self._on_doc_tree_activate)

        self._doc_paths = {}

    def setup_month_ui(self):
        header = ctk.CTkFrame(self.month_frame, fg_color="transparent")
        header.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(header, text="Kế hoạch tháng", font=("Arial", 26, "bold"),
                     text_color=COLORS["text"]).pack(side="left")

        ctk.CTkButton(header, text="Làm mới", width=100, height=32,
                      fg_color=COLORS["accent"], hover_color="#1D4ED8",
                      command=self.render_month).pack(side="right")

        ctk.CTkButton(header, text="Mở file Excel", width=120, height=32,
                      fg_color="transparent", text_color=COLORS["text"],
                      border_width=1, border_color=COLORS["border"],
                      hover_color=COLORS["hover"],
                      command=self.open_schedule_file).pack(side="right", padx=(0, 8))

        toolbar = ctk.CTkFrame(self.month_frame, fg_color=COLORS["card"],
                                corner_radius=8, border_width=1,
                                border_color=COLORS["border"])
        toolbar.pack(fill="x", pady=(0, 8))

        search_wrap = ctk.CTkFrame(toolbar, fg_color="transparent")
        search_wrap.pack(side="left", fill="y", padx=10, pady=8)
        ctk.CTkLabel(search_wrap, text="🔍", font=("Arial", 12)
                     ).pack(side="left", padx=(0, 4))
        self.month_search = ctk.CTkEntry(search_wrap, height=30, width=200,
                                          placeholder_text="Tìm giảng viên...",
                                          border_color=COLORS["border"])
        self.month_search.pack(side="left")
        self.month_search.bind("<KeyRelease>", lambda e: self.render_month())

        subj_wrap = ctk.CTkFrame(toolbar, fg_color="transparent")
        subj_wrap.pack(side="left", padx=(0, 10), pady=8)
        ctk.CTkLabel(subj_wrap, text="Môn:", font=("Arial", 11),
                     text_color=COLORS["text_dim"]
                     ).pack(side="left", padx=(0, 6))
        self.month_subject = tk.StringVar(value="Tất cả")
        self.month_subject_menu = ctk.CTkOptionMenu(subj_wrap,
                                                     variable=self.month_subject,
                                                     values=["Tất cả"],
                                                     width=110, height=30,
                                                     fg_color=COLORS["accent"],
                                                     button_color=COLORS["accent"],
                                                     button_hover_color="#1D4ED8",
                                                     command=lambda _: self.render_month())
        self.month_subject_menu.pack(side="left")

        self.hide_empty_var = tk.BooleanVar(value=True)
        ctk.CTkSwitch(toolbar, text="Ẩn tiết trống",
                      variable=self.hide_empty_var,
                      font=("Arial", 11),
                      progress_color=COLORS["accent"],
                      command=self.render_month
                      ).pack(side="left", padx=(0, 10), pady=8)

        self.month_info = ctk.CTkLabel(toolbar, text="",
                                        font=("Arial", 11),
                                        text_color=COLORS["text_dim"], anchor="e")
        self.month_info.pack(side="right", padx=10, pady=8)

        # Legend (chú thích loại tiết)
        legend = ctk.CTkFrame(self.month_frame, fg_color="transparent")
        legend.pack(fill="x", pady=(0, 6))
        for icon, label in [("⚫", "Lý thuyết"),
                            ("🔴", "Kiểm tra"),
                            ("🟢", "Thực hành")]:
            box = ctk.CTkFrame(legend, fg_color=COLORS["card"],
                                corner_radius=14,
                                border_width=1,
                                border_color=COLORS["border"])
            box.pack(side="left", padx=(0, 8))
            ctk.CTkLabel(box, text=f"  {icon} {label}  ",
                          font=("Arial", 11),
                          text_color=COLORS["text"]).pack(padx=2, pady=2)

        # Tiêu đề như trong file Excel
        excel_title = ctk.CTkFrame(self.month_frame, fg_color="white",
                                    corner_radius=8, border_width=1,
                                    border_color=COLORS["border"])
        excel_title.pack(fill="x", pady=(0, 0))
        self.lbl_excel_title = ctk.CTkLabel(excel_title,
                                             text="KẾ HOẠCH PHÂN CÔNG GIẢNG DẠY",
                                             font=("Times New Roman", 18, "bold"),
                                             text_color=COLORS["text"])
        self.lbl_excel_title.pack(pady=(12, 2))
        self.lbl_excel_subtitle = ctk.CTkLabel(excel_title, text="Tháng — Năm —",
                                                font=("Times New Roman", 13),
                                                text_color=COLORS["text_dim"])
        self.lbl_excel_subtitle.pack(pady=(0, 12))

        container = ctk.CTkFrame(self.month_frame, fg_color=COLORS["card"],
                                 border_width=1, border_color="#94A3B8",
                                 corner_radius=0)
        container.pack(fill="both", expand=True, pady=(0, 0))
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        try:
            style.theme_use("default")
        except Exception:
            pass
        # Excel-like: kẻ ô + heading đậm
        style.configure("Month.Treeview", rowheight=42,
                        font=("Times New Roman", 12),
                        background="white", fieldbackground="white",
                        foreground=COLORS["text"],
                        borderwidth=1, relief="solid")
        style.configure("Month.Treeview.Heading",
                        font=("Times New Roman", 12, "bold"),
                        background="#1E40AF", foreground="white",
                        padding=(6, 8),
                        borderwidth=1, relief="solid")
        style.map("Month.Treeview.Heading",
                  background=[("active", "#1E3A8A")])
        style.map("Month.Treeview", background=[("selected", "#FEF08A")],
                  foreground=[("selected", COLORS["text"])])
        style.layout("Month.Treeview", [
            ("Treeview.treearea", {"sticky": "nswe"})
        ])

        self.month_tree = ttk.Treeview(container, style="Month.Treeview",
                                        show="headings")
        vsb = ttk.Scrollbar(container, orient="vertical",
                             command=self.month_tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal",
                             command=self.month_tree.xview)
        self.month_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.month_tree.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.month_tree.tag_configure("alt", background="#F8FAFC")
        self.month_tree.tag_configure("group_top", background="#DBEAFE",
                                       font=("Times New Roman", 12, "bold"))
        self.month_tree.tag_configure("subj_BC", background="#FFFFFF")
        self.month_tree.tag_configure("subj_DH", background="#FFFFFF")
        self.month_tree.tag_configure("subj_KB", background="#FFFFFF")
        self.month_tree.tag_configure("subj_DN", background="#FFFFFF")

    def show_month_frame(self):
        self.hide_all_frames()
        self.month_frame.pack(fill="both", expand=True)
        self.set_active_nav("month")
        self.render_month()

    # ===================== TAB BÁO CÁO TUẦN =====================

    def setup_report_ui(self):
        header = ctk.CTkFrame(self.report_frame, fg_color="transparent")
        header.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(header, text="Báo cáo tuần",
                     font=("Arial", 26, "bold"),
                     text_color=COLORS["text"]).pack(side="left")

        ctk.CTkButton(header, text="Xuất Excel", width=120, height=36,
                       fg_color=COLORS["accent"], hover_color="#1D4ED8",
                       font=("Arial", 13, "bold"),
                       command=self.export_report_excel
                       ).pack(side="right")

        ctk.CTkButton(header, text="Làm mới", width=100, height=36,
                       fg_color="transparent", text_color=COLORS["text"],
                       border_width=1, border_color=COLORS["border"],
                       hover_color=COLORS["hover"],
                       font=("Arial", 13, "bold"),
                       command=self.render_report
                       ).pack(side="right", padx=(0, 8))

        toolbar = ctk.CTkFrame(self.report_frame, fg_color=COLORS["card"],
                                corner_radius=8, border_width=1,
                                border_color=COLORS["border"])
        toolbar.pack(fill="x", pady=(0, 10))

        # Tuần báo cáo: Thứ 5 (start) → Thứ 4 tuần sau (end)
        # weekday(): 0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri, 5=Sat, 6=Sun
        today = datetime.now()
        # Tính Thứ 5 gần nhất (về phía quá khứ hoặc hôm nay nếu hôm nay là Thứ 5)
        days_since_thu = (today.weekday() - 3) % 7
        self._report_start = today - timedelta(days=days_since_thu)

        nav = ctk.CTkFrame(toolbar, fg_color="transparent")
        nav.pack(side="left", padx=10, pady=8)

        ctk.CTkButton(nav, text="‹", width=30, height=30,
                       fg_color="transparent", text_color=COLORS["text"],
                       hover_color=COLORS["hover"],
                       command=lambda: self._shift_week(-1)
                       ).pack(side="left")
        self.lbl_report_week = ctk.CTkLabel(nav, text="",
                                              font=("Arial", 13, "bold"),
                                              text_color=COLORS["text"], width=270)
        self.lbl_report_week.pack(side="left", padx=8)
        ctk.CTkButton(nav, text="›", width=30, height=30,
                       fg_color="transparent", text_color=COLORS["text"],
                       hover_color=COLORS["hover"],
                       command=lambda: self._shift_week(1)
                       ).pack(side="left")
        ctk.CTkButton(nav, text="Tuần hiện tại",
                       width=110, height=30,
                       fg_color="transparent", text_color=COLORS["accent"],
                       border_width=1, border_color=COLORS["accent"],
                       hover_color=COLORS["hover"],
                       font=("Arial", 11, "bold"),
                       command=self._reset_week
                       ).pack(side="left", padx=(10, 0))

        # Stats tổng
        self.report_stats = ctk.CTkLabel(toolbar, text="",
                                          font=("Arial", 12),
                                          text_color=COLORS["text_dim"])
        self.report_stats.pack(side="right", padx=14, pady=8)

        # Bảng báo cáo
        container = ctk.CTkFrame(self.report_frame, fg_color=COLORS["card"],
                                  corner_radius=10, border_width=1,
                                  border_color=COLORS["border"])
        container.pack(fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        try:
            style.theme_use("default")
        except Exception:
            pass
        style.configure("Report.Treeview", rowheight=38,
                        font=("Segoe UI", 12),
                        background="white", fieldbackground="white",
                        foreground=COLORS["text"],
                        borderwidth=1, relief="solid")
        style.configure("Report.Treeview.Heading",
                        font=("Segoe UI", 12, "bold"),
                        background="#1E40AF", foreground="white",
                        padding=(6, 8))

        self.report_tree = ttk.Treeview(container, style="Report.Treeview",
                                          show="headings")
        vsb = ttk.Scrollbar(container, orient="vertical",
                             command=self.report_tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal",
                             command=self.report_tree.xview)
        self.report_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.report_tree.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.report_tree.tag_configure("total", background="#DBEAFE",
                                         font=("Segoe UI", 12, "bold"))
        self.report_tree.tag_configure("alt", background="#F8FAFC")

    def _shift_week(self, delta):
        self._report_start = self._report_start + timedelta(days=7 * delta)
        self.render_report()

    def _reset_week(self):
        today = datetime.now()
        days_since_thu = (today.weekday() - 3) % 7
        self._report_start = today - timedelta(days=days_since_thu)
        self.render_report()

    def show_report_frame(self):
        self.hide_all_frames()
        self.report_frame.pack(fill="both", expand=True)
        self.set_active_nav("report")
        # Đảm bảo có data từ render_month
        if not hasattr(self, "_last_teachers_merged"):
            self.render_month()

        # Auto-jump đến tuần đầu tiên có data nếu tuần hiện tại không có
        try:
            self._auto_jump_to_data_week()
        except Exception:
            pass

        self.render_report()

    def _auto_jump_to_data_week(self):
        """Nếu tuần hiện tại không có data, nhảy đến tuần đầu tiên có data."""
        import re as _re
        header_row = getattr(self, "_last_header_row", None)
        if not header_row:
            return
        # Lấy tất cả các ngày có trong file
        dates_in_file = []
        for h in header_row:
            m = _re.match(r"(\d{4})-(\d{2})-(\d{2})", str(h))
            if m:
                dates_in_file.append(datetime(int(m.group(1)),
                                              int(m.group(2)),
                                              int(m.group(3))))
        if not dates_in_file:
            return

        cur_start = self._report_start
        cur_end = cur_start + timedelta(days=6)
        has_overlap = any(cur_start <= d <= cur_end for d in dates_in_file)
        if has_overlap:
            return

        # Nhảy về Thu của tuần chứa ngày đầu tiên trong file
        first = min(dates_in_file)
        days_since_thu = (first.weekday() - 3) % 7
        self._report_start = first - timedelta(days=days_since_thu)

    def render_report(self):
        if not hasattr(self, "report_tree"):
            return
        for item in self.report_tree.get_children():
            self.report_tree.delete(item)

        start = self._report_start  # Thu
        end = start + timedelta(days=6)  # Wed tuần sau

        # Hiển thị nhãn tuần + cảnh báo nếu giao tháng
        cross_month = (start.month != end.month) or (start.year != end.year)
        label = f"{start.strftime('%d/%m')} (T5) → {end.strftime('%d/%m/%Y')} (T4)"
        if cross_month:
            label += "  ⚠ giao 2 tháng"
        self.lbl_report_week.configure(text=label)

        teachers = getattr(self, "_last_teachers_merged", None)
        header_row = getattr(self, "_last_header_row", None)
        if not teachers or not header_row:
            self.report_stats.configure(
                text="Chưa có dữ liệu. Vào tab Kế hoạch tháng trước.")
            return

        import re as _re

        # Xác định col index của 7 ngày trong tuần
        day_to_col = {}  # date_str → col_idx
        for i, h in enumerate(header_row):
            m = _re.match(r"(\d{4})-(\d{2})-(\d{2})", str(h))
            if m:
                date_str = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
                day_to_col[date_str] = i

        week_dates = [(start + timedelta(days=i)) for i in range(7)]
        week_keys = [d.strftime("%Y-%m-%d") for d in week_dates]
        cols_for_week = [day_to_col.get(k) for k in week_keys]

        # Setup columns: Thứ 5 → Chủ nhật → Thứ 4
        weekday_names = ["T5", "T6", "T7", "CN", "T2", "T3", "T4"]
        col_ids = ["c_tt", "c_name", "c_subj"] + [f"d{i}" for i in range(7)] + ["c_lt", "c_kt", "c_th", "c_total"]
        self.report_tree.configure(columns=col_ids)

        self.report_tree.heading("c_tt", text="TT")
        self.report_tree.column("c_tt", width=46, anchor="center", stretch=False)
        self.report_tree.heading("c_name", text="Họ và tên")
        self.report_tree.column("c_name", width=200, anchor="w", stretch=False)
        self.report_tree.heading("c_subj", text="Môn")
        self.report_tree.column("c_subj", width=110, anchor="center", stretch=False)

        for i, (wd, d) in enumerate(zip(weekday_names, week_dates)):
            cid = f"d{i}"
            self.report_tree.heading(cid, text=f"{d.strftime('%d/%m')} {wd}")
            self.report_tree.column(cid, width=110, anchor="center", stretch=False)

        for cid, lbl in [("c_lt", "LT"), ("c_kt", "KT"),
                          ("c_th", "TH"), ("c_total", "TS")]:
            self.report_tree.heading(cid, text=lbl)
            self.report_tree.column(cid, width=58, anchor="center", stretch=False)

        # Nếu tuần giao 2 tháng - báo và không render data (theo yêu cầu khách)
        if cross_month:
            self.report_stats.configure(
                text="Tuần này giao 2 tháng - bỏ qua (yêu cầu khách).")
            return

        # Aggregate: per teacher, count tiết theo loại cho từng ngày trong tuần
        sum_lt = sum_kt = sum_th = 0
        rows_inserted = 0

        for t_idx, t in enumerate(teachers):
            day_counts = {i: {"LT": 0, "KT": 0, "TH": 0} for i in range(7)}

            for slot_key, srow in t["slot_rows"].items():
                cells = srow["cells"]
                for di, col_idx in enumerate(cols_for_week):
                    if col_idx is None or col_idx >= len(cells):
                        continue
                    v = cells[col_idx]
                    if not v:
                        continue
                    # Đếm số entry trong cell (phân cách bởi " / ")
                    for entry in v.split(" / "):
                        if entry.startswith("🔴"):
                            day_counts[di]["KT"] += 1
                        elif entry.startswith("🟢"):
                            day_counts[di]["TH"] += 1
                        else:
                            day_counts[di]["LT"] += 1

            tot_lt = sum(dc["LT"] for dc in day_counts.values())
            tot_kt = sum(dc["KT"] for dc in day_counts.values())
            tot_th = sum(dc["TH"] for dc in day_counts.values())
            total = tot_lt + tot_kt + tot_th

            if total == 0:
                continue

            sum_lt += tot_lt; sum_kt += tot_kt; sum_th += tot_th

            values = [
                t["tt"] or str(rows_inserted + 1),
                t["name"].upper(),
                " / ".join(t["subjects"]),
            ]
            for di in range(7):
                dc = day_counts[di]
                total_day = dc["LT"] + dc["KT"] + dc["TH"]
                if total_day == 0:
                    values.append("")
                else:
                    parts = []
                    if dc["LT"]: parts.append(str(dc["LT"]))
                    if dc["KT"]: parts.append(f"🔴{dc['KT']}")
                    if dc["TH"]: parts.append(f"🟢{dc['TH']}")
                    values.append("·".join(parts))
            values.extend([str(tot_lt), str(tot_kt), str(tot_th), str(total)])

            tags = ["alt"] if rows_inserted % 2 == 1 else []
            self.report_tree.insert("", "end", values=values, tags=tuple(tags))
            rows_inserted += 1

        # Total row
        grand_total = sum_lt + sum_kt + sum_th
        if rows_inserted > 0:
            self.report_tree.insert("", "end",
                values=["", "TỔNG CỘNG", "", "", "", "", "", "", "", "",
                         str(sum_lt), str(sum_kt), str(sum_th), str(grand_total)],
                tags=("total",))

        self.report_stats.configure(
            text=f"{rows_inserted} GV · LT {sum_lt} · KT {sum_kt} · TH {sum_th} · Tổng {grand_total} tiết")

    def export_report_excel(self):
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

        if not hasattr(self, "_report_start"):
            return
        start = self._report_start
        end = start + timedelta(days=6)

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"BaoCao_Tuan_{start.strftime('%Y-%m-%d')}.xlsx")
        if not path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Báo cáo tuần"

            thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                           top=Side(style='thin'), bottom=Side(style='thin'))
            center = Alignment(horizontal='center', vertical='center', wrap_text=True)
            header_fill = PatternFill("solid", fgColor="1E40AF")
            total_fill = PatternFill("solid", fgColor="DBEAFE")
            header_font = Font(bold=True, color="FFFFFF", size=11)

            # Title
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=14)
            ws.cell(row=1, column=1, value=f"BÁO CÁO HUẤN LUYỆN TUẦN {start.strftime('%d/%m')} - {end.strftime('%d/%m/%Y')}")
            ws.cell(row=1, column=1).font = Font(bold=True, size=14)
            ws.cell(row=1, column=1).alignment = center

            # Headers: Thứ 5 → Thứ 4
            weekday_labels = ["T5", "T6", "T7", "CN", "T2", "T3", "T4"]
            headers = ["TT", "Họ và tên", "Môn"] + \
                      [f"{weekday_labels[i]} {(start + timedelta(days=i)).strftime('%d/%m')}"
                       for i in range(7)] + \
                      ["LT", "KT", "TH", "Tổng số"]
            for c, h in enumerate(headers, 1):
                cell = ws.cell(row=3, column=c, value=h)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = thin
                cell.alignment = center

            # Lấy data từ treeview
            data_rows = []
            total_row = None
            for item in self.report_tree.get_children():
                vals = self.report_tree.item(item)["values"]
                tags = self.report_tree.item(item)["tags"]
                if "total" in tags:
                    total_row = vals
                else:
                    data_rows.append(vals)

            r = 4
            for vals in data_rows:
                for c, v in enumerate(vals, 1):
                    cell = ws.cell(row=r, column=c, value=str(v))
                    cell.border = thin
                    cell.alignment = center
                r += 1

            if total_row:
                for c, v in enumerate(total_row, 1):
                    cell = ws.cell(row=r, column=c, value=str(v))
                    cell.border = thin
                    cell.alignment = center
                    cell.font = Font(bold=True)
                    cell.fill = total_fill

            # Column widths
            ws.column_dimensions['A'].width = 5
            ws.column_dimensions['B'].width = 26
            ws.column_dimensions['C'].width = 14
            for i in range(7):
                col = chr(ord('D') + i)
                ws.column_dimensions[col].width = 12
            for i in range(4):
                col = chr(ord('K') + i)
                ws.column_dimensions[col].width = 9

            wb.save(path)

            try:
                if os.name == "nt":
                    os.startfile(path)
                else:
                    subprocess.call(["open", path])
            except Exception:
                pass
        except Exception as e:
            print(f"[export_report] Lỗi: {e}")

    # ===================== TAB BÁO CÁO NGÀY =====================

    def setup_report_day_ui(self):
        header = ctk.CTkFrame(self.report_day_frame, fg_color="transparent")
        header.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(header, text="Báo cáo ngày",
                     font=("Arial", 26, "bold"),
                     text_color=COLORS["text"]).pack(side="left")

        ctk.CTkButton(header, text="Xuất Excel", width=120, height=36,
                       fg_color=COLORS["accent"], hover_color="#1D4ED8",
                       font=("Arial", 13, "bold"),
                       command=self.export_report_day_excel
                       ).pack(side="right")

        ctk.CTkButton(header, text="Làm mới", width=100, height=36,
                       fg_color="transparent", text_color=COLORS["text"],
                       border_width=1, border_color=COLORS["border"],
                       hover_color=COLORS["hover"],
                       font=("Arial", 13, "bold"),
                       command=self.render_report_day
                       ).pack(side="right", padx=(0, 8))

        toolbar = ctk.CTkFrame(self.report_day_frame, fg_color=COLORS["card"],
                                corner_radius=8, border_width=1,
                                border_color=COLORS["border"])
        toolbar.pack(fill="x", pady=(0, 10))

        # Chọn ngày
        self._report_day = datetime.now()

        nav = ctk.CTkFrame(toolbar, fg_color="transparent")
        nav.pack(side="left", padx=10, pady=8)

        ctk.CTkButton(nav, text="‹", width=30, height=30,
                       fg_color="transparent", text_color=COLORS["text"],
                       hover_color=COLORS["hover"],
                       command=lambda: self._shift_day(-1)
                       ).pack(side="left")
        self.lbl_report_day = ctk.CTkLabel(nav, text="",
                                             font=("Arial", 13, "bold"),
                                             text_color=COLORS["text"], width=240)
        self.lbl_report_day.pack(side="left", padx=8)
        ctk.CTkButton(nav, text="›", width=30, height=30,
                       fg_color="transparent", text_color=COLORS["text"],
                       hover_color=COLORS["hover"],
                       command=lambda: self._shift_day(1)
                       ).pack(side="left")
        ctk.CTkButton(nav, text="Hôm nay",
                       width=90, height=30,
                       fg_color="transparent", text_color=COLORS["accent"],
                       border_width=1, border_color=COLORS["accent"],
                       hover_color=COLORS["hover"],
                       font=("Arial", 11, "bold"),
                       command=self._reset_day
                       ).pack(side="left", padx=(10, 0))

        self.report_day_stats = ctk.CTkLabel(toolbar, text="",
                                              font=("Arial", 12),
                                              text_color=COLORS["text_dim"])
        self.report_day_stats.pack(side="right", padx=14, pady=8)

        # 2 cột: Sáng (tiết 1-6) | Chiều (tiết 7-9)
        body = ctk.CTkFrame(self.report_day_frame, fg_color="transparent")
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(0, weight=1, uniform="session")
        body.grid_columnconfigure(1, weight=1, uniform="session")
        body.grid_rowconfigure(0, weight=1)

        self.day_morning_card = self._make_session_card(body, "Buổi sáng (tiết 1-6)", "#F59E0B")
        self.day_morning_card["frame"].grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        self.day_afternoon_card = self._make_session_card(body, "Buổi chiều (tiết 7-9)", "#8B5CF6")
        self.day_afternoon_card["frame"].grid(row=0, column=1, sticky="nsew", padx=(5, 0))

    def _make_session_card(self, parent, title, accent):
        wrap = ctk.CTkFrame(parent, fg_color=COLORS["card"],
                             corner_radius=12, border_width=1,
                             border_color=COLORS["border"])

        hd = ctk.CTkFrame(wrap, fg_color=accent, height=46, corner_radius=0)
        hd.pack(fill="x", padx=1, pady=(1, 0))
        hd.pack_propagate(False)
        ctk.CTkLabel(hd, text=title, font=("Arial", 15, "bold"),
                      text_color="white").pack(pady=10)

        stats_box = ctk.CTkFrame(wrap, fg_color=COLORS["bg"], corner_radius=0)
        stats_box.pack(fill="x", padx=1)

        lbl_summary = ctk.CTkLabel(stats_box, text="",
                                     font=("Arial", 13, "bold"),
                                     text_color=COLORS["text"])
        lbl_summary.pack(pady=10)

        # 2 cột bên trong: GV giảng | GV không giảng
        cols = ctk.CTkFrame(wrap, fg_color="transparent")
        cols.pack(fill="both", expand=True, padx=8, pady=(8, 8))
        cols.grid_columnconfigure(0, weight=1, uniform="col")
        cols.grid_columnconfigure(1, weight=1, uniform="col")
        cols.grid_rowconfigure(0, weight=1)

        teach_wrap = ctk.CTkFrame(cols, fg_color="white",
                                    corner_radius=8, border_width=1,
                                    border_color="#10B981")
        teach_wrap.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        ctk.CTkLabel(teach_wrap, text="GV ĐI GIẢNG",
                      font=("Arial", 12, "bold"),
                      text_color="#10B981").pack(pady=(8, 4))
        teach_scroll = ctk.CTkScrollableFrame(teach_wrap, fg_color="transparent")
        teach_scroll.pack(fill="both", expand=True, padx=4, pady=(0, 6))

        not_wrap = ctk.CTkFrame(cols, fg_color="white",
                                  corner_radius=8, border_width=1,
                                  border_color="#94A3B8")
        not_wrap.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        ctk.CTkLabel(not_wrap, text="GV KHÔNG GIẢNG",
                      font=("Arial", 12, "bold"),
                      text_color="#475569").pack(pady=(8, 4))
        not_scroll = ctk.CTkScrollableFrame(not_wrap, fg_color="transparent")
        not_scroll.pack(fill="both", expand=True, padx=4, pady=(0, 6))

        return {
            "frame": wrap,
            "summary": lbl_summary,
            "teach_scroll": teach_scroll,
            "not_scroll": not_scroll,
        }

    def _shift_day(self, delta):
        self._report_day = self._report_day + timedelta(days=delta)
        self.render_report_day()

    def _reset_day(self):
        self._report_day = datetime.now()
        self.render_report_day()

    def show_report_day_frame(self):
        self.hide_all_frames()
        self.report_day_frame.pack(fill="both", expand=True)
        self.set_active_nav("report_day")
        if not hasattr(self, "_last_teachers_merged"):
            self.render_month()
        # Auto-jump đến ngày có data nếu hôm nay không có
        try:
            self._auto_jump_to_data_day()
        except Exception:
            pass
        self.render_report_day()

    def _auto_jump_to_data_day(self):
        import re as _re
        header_row = getattr(self, "_last_header_row", None)
        if not header_row:
            return
        dates_in_file = []
        for h in header_row:
            m = _re.match(r"(\d{4})-(\d{2})-(\d{2})", str(h))
            if m:
                dates_in_file.append(datetime(int(m.group(1)),
                                              int(m.group(2)),
                                              int(m.group(3))))
        if not dates_in_file:
            return
        cur = self._report_day.date()
        has = any(d.date() == cur for d in dates_in_file)
        if has:
            return
        # Lấy ngày đầu tiên trong file
        self._report_day = min(dates_in_file)

    def render_report_day(self):
        if not hasattr(self, "day_morning_card"):
            return

        d = self._report_day
        weekday_vn = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ nhật"]
        self.lbl_report_day.configure(
            text=f"{weekday_vn[d.weekday()]}, {d.strftime('%d/%m/%Y')}")

        teachers = getattr(self, "_last_teachers_merged", None)
        header_row = getattr(self, "_last_header_row", None)
        if not teachers or not header_row:
            self.report_day_stats.configure(text="Chưa có dữ liệu. Vào Kế hoạch tháng trước.")
            self._fill_session_card(self.day_morning_card, [], teachers or [], 0)
            self._fill_session_card(self.day_afternoon_card, [], teachers or [], 0)
            return

        import re as _re
        target_key = d.strftime("%Y-%m-%d")
        target_col = None
        for i, h in enumerate(header_row):
            m = _re.match(r"(\d{4})-(\d{2})-(\d{2})", str(h))
            if m and f"{m.group(1)}-{m.group(2)}-{m.group(3)}" == target_key:
                target_col = i
                break

        if target_col is None:
            self.report_day_stats.configure(
                text=f"Ngày {d.strftime('%d/%m/%Y')} không có trong file kế hoạch tháng.")
            self._fill_session_card(self.day_morning_card, [], teachers, 0)
            self._fill_session_card(self.day_afternoon_card, [], teachers, 0)
            return

        # Phân loại slot → buổi: 1-6 = sáng, 7-9 = chiều
        slot_to_session_re = _re.compile(r"(\d+)\s*-\s*(\d+)")

        morning_teachers = []  # list of (name, classes_text)
        afternoon_teachers = []
        all_teacher_names = []

        for t in teachers:
            name = t["name"].strip().upper()
            all_teacher_names.append(name)
            morning_entries = []
            afternoon_entries = []

            for slot_key, srow in t["slot_rows"].items():
                m = slot_to_session_re.match(srow["slot_text"])
                if not m:
                    continue
                slot_start = int(m.group(1))
                cell_val = srow["cells"][target_col] if target_col < len(srow["cells"]) else ""
                if not cell_val.strip():
                    continue
                if slot_start <= 6:
                    morning_entries.append((srow["slot_text"], cell_val))
                else:
                    afternoon_entries.append((srow["slot_text"], cell_val))

            if morning_entries:
                morning_teachers.append((name, morning_entries))
            if afternoon_entries:
                afternoon_teachers.append((name, afternoon_entries))

        total_teachers = len(all_teacher_names)
        self._fill_session_card(self.day_morning_card, morning_teachers,
                                  all_teacher_names, total_teachers)
        self._fill_session_card(self.day_afternoon_card, afternoon_teachers,
                                  all_teacher_names, total_teachers)

        self.report_day_stats.configure(
            text=f"Tổng {total_teachers} GV · Sáng {len(morning_teachers)} giảng / {total_teachers - len(morning_teachers)} nghỉ · "
                  f"Chiều {len(afternoon_teachers)} giảng / {total_teachers - len(afternoon_teachers)} nghỉ")

    def _fill_session_card(self, card, teach_list, all_names, total):
        for w in card["teach_scroll"].winfo_children():
            w.destroy()
        for w in card["not_scroll"].winfo_children():
            w.destroy()

        teaching_names = {name for name, _ in teach_list}
        not_teaching = [n for n in all_names if n not in teaching_names]

        card["summary"].configure(
            text=f"📊 {len(teach_list)} GV giảng  ·  {len(not_teaching)} GV nghỉ  /  Tổng {total}")

        if not teach_list:
            ctk.CTkLabel(card["teach_scroll"], text="Không có GV giảng",
                          font=("Arial", 11),
                          text_color=COLORS["text_dim"]).pack(pady=10)
        else:
            for name, entries in teach_list:
                row = ctk.CTkFrame(card["teach_scroll"],
                                    fg_color=COLORS["bg"],
                                    corner_radius=6)
                row.pack(fill="x", pady=2, padx=2)
                ctk.CTkLabel(row, text=name, font=("Arial", 12, "bold"),
                              text_color=COLORS["text"], anchor="w"
                              ).pack(side="left", padx=8, pady=4)
                detail = " · ".join(f"{slot}: {val}" for slot, val in entries)
                ctk.CTkLabel(row, text=detail,
                              font=("Arial", 10),
                              text_color=COLORS["text_dim"],
                              anchor="w").pack(side="left", padx=4, pady=4)

        if not not_teaching:
            ctk.CTkLabel(card["not_scroll"],
                          text="Tất cả GV đều có giảng",
                          font=("Arial", 11),
                          text_color=COLORS["text_dim"]).pack(pady=10)
        else:
            for name in not_teaching:
                row = ctk.CTkFrame(card["not_scroll"],
                                    fg_color=COLORS["bg"],
                                    corner_radius=6, height=30)
                row.pack(fill="x", pady=2, padx=2)
                row.pack_propagate(False)
                ctk.CTkLabel(row, text=name, font=("Arial", 12),
                              text_color=COLORS["text_dim"], anchor="w"
                              ).pack(side="left", padx=8)

    def export_report_day_excel(self):
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

        if not hasattr(self, "_report_day"):
            return
        d = self._report_day

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"BaoCao_Ngay_{d.strftime('%Y-%m-%d')}.xlsx")
        if not path:
            return

        try:
            teachers = getattr(self, "_last_teachers_merged", None)
            header_row = getattr(self, "_last_header_row", None)
            if not teachers or not header_row:
                return

            import re as _re
            target_key = d.strftime("%Y-%m-%d")
            target_col = None
            for i, h in enumerate(header_row):
                m = _re.match(r"(\d{4})-(\d{2})-(\d{2})", str(h))
                if m and f"{m.group(1)}-{m.group(2)}-{m.group(3)}" == target_key:
                    target_col = i
                    break
            if target_col is None:
                return

            slot_re = _re.compile(r"(\d+)\s*-\s*(\d+)")
            morning_teach, morning_not = [], []
            afternoon_teach, afternoon_not = [], []

            for t in teachers:
                name = t["name"].strip().upper()
                m_has = False
                a_has = False
                for slot_key, srow in t["slot_rows"].items():
                    sm = slot_re.match(srow["slot_text"])
                    if not sm:
                        continue
                    val = srow["cells"][target_col] if target_col < len(srow["cells"]) else ""
                    if not val.strip():
                        continue
                    if int(sm.group(1)) <= 6:
                        m_has = True
                    else:
                        a_has = True
                (morning_teach if m_has else morning_not).append(name)
                (afternoon_teach if a_has else afternoon_not).append(name)

            wb = Workbook()
            ws = wb.active
            ws.title = "Báo cáo ngày"

            thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                           top=Side(style='thin'), bottom=Side(style='thin'))
            center = Alignment(horizontal='center', vertical='center', wrap_text=True)
            left = Alignment(horizontal='left', vertical='center', wrap_text=True)
            morning_fill = PatternFill("solid", fgColor="F59E0B")
            afternoon_fill = PatternFill("solid", fgColor="8B5CF6")
            white_font = Font(bold=True, color="FFFFFF", size=11)

            weekday_vn = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ nhật"]

            # Title
            ws.merge_cells("A1:D1")
            ws.cell(row=1, column=1,
                     value=f"BÁO CÁO HUẤN LUYỆN NGÀY {weekday_vn[d.weekday()]}, {d.strftime('%d/%m/%Y')}")
            ws.cell(row=1, column=1).font = Font(bold=True, size=14)
            ws.cell(row=1, column=1).alignment = center

            r = 3
            # Buổi sáng
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
            cell = ws.cell(row=r, column=1,
                            value=f"BUỔI SÁNG (Tiết 1-6) · {len(morning_teach)} giảng / {len(morning_not)} nghỉ")
            cell.font = white_font
            cell.fill = morning_fill
            cell.alignment = center
            r += 1

            # Header GV giảng / GV nghỉ
            ws.cell(row=r, column=1, value="STT").font = Font(bold=True)
            ws.cell(row=r, column=2, value="GV đi giảng").font = Font(bold=True)
            ws.cell(row=r, column=3, value="STT").font = Font(bold=True)
            ws.cell(row=r, column=4, value="GV không giảng").font = Font(bold=True)
            for c in range(1, 5):
                ws.cell(row=r, column=c).border = thin
                ws.cell(row=r, column=c).alignment = center
            r += 1

            max_morning = max(len(morning_teach), len(morning_not), 1)
            for i in range(max_morning):
                ws.cell(row=r, column=1, value=i+1 if i < len(morning_teach) else "")
                ws.cell(row=r, column=2, value=morning_teach[i] if i < len(morning_teach) else "")
                ws.cell(row=r, column=3, value=i+1 if i < len(morning_not) else "")
                ws.cell(row=r, column=4, value=morning_not[i] if i < len(morning_not) else "")
                for c in range(1, 5):
                    ws.cell(row=r, column=c).border = thin
                    ws.cell(row=r, column=c).alignment = left if c in (2, 4) else center
                r += 1

            r += 1
            # Buổi chiều
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
            cell = ws.cell(row=r, column=1,
                            value=f"BUỔI CHIỀU (Tiết 7-9) · {len(afternoon_teach)} giảng / {len(afternoon_not)} nghỉ")
            cell.font = white_font
            cell.fill = afternoon_fill
            cell.alignment = center
            r += 1

            ws.cell(row=r, column=1, value="STT").font = Font(bold=True)
            ws.cell(row=r, column=2, value="GV đi giảng").font = Font(bold=True)
            ws.cell(row=r, column=3, value="STT").font = Font(bold=True)
            ws.cell(row=r, column=4, value="GV không giảng").font = Font(bold=True)
            for c in range(1, 5):
                ws.cell(row=r, column=c).border = thin
                ws.cell(row=r, column=c).alignment = center
            r += 1

            max_aft = max(len(afternoon_teach), len(afternoon_not), 1)
            for i in range(max_aft):
                ws.cell(row=r, column=1, value=i+1 if i < len(afternoon_teach) else "")
                ws.cell(row=r, column=2, value=afternoon_teach[i] if i < len(afternoon_teach) else "")
                ws.cell(row=r, column=3, value=i+1 if i < len(afternoon_not) else "")
                ws.cell(row=r, column=4, value=afternoon_not[i] if i < len(afternoon_not) else "")
                for c in range(1, 5):
                    ws.cell(row=r, column=c).border = thin
                    ws.cell(row=r, column=c).alignment = left if c in (2, 4) else center
                r += 1

            ws.column_dimensions['A'].width = 6
            ws.column_dimensions['B'].width = 32
            ws.column_dimensions['C'].width = 6
            ws.column_dimensions['D'].width = 32

            wb.save(path)
            try:
                if os.name == "nt":
                    os.startfile(path)
                else:
                    subprocess.call(["open", path])
            except Exception:
                pass
        except Exception as e:
            print(f"[export_report_day] Lỗi: {e}")

    def open_schedule_file(self):
        path = self.config_data.get("schedule_file", "schedule.xlsx")
        if not os.path.exists(path):
            self.month_info.configure(text=f"Không tìm thấy {path}")
            return
        try:
            if os.name == "nt":
                os.startfile(path)
            else:
                subprocess.call(["open", path])
        except Exception as e:
            self.month_info.configure(text=f"Lỗi mở file: {e}")

    def _classify_cell_color(self, cell):
        """Phân loại tiết theo màu font: đỏ=KT, xanh lá/dương=TH, đen=LT."""
        try:
            c = cell.font.color if cell.font else None
            rgb = None
            if c is not None:
                if hasattr(c, "type") and c.type == "rgb":
                    if isinstance(c.rgb, str):
                        rgb = c.rgb
                elif isinstance(c.rgb, str):
                    rgb = c.rgb
            if rgb and len(rgb) >= 6:
                hex_rgb = rgb[-6:].upper()
                r = int(hex_rgb[0:2], 16)
                g = int(hex_rgb[2:4], 16)
                b = int(hex_rgb[4:6], 16)
                # Đỏ trội
                if r > 150 and r > g + 40 and r > b + 40:
                    return "KT"
                # Xanh trội (xanh lá hoặc xanh dương)
                if (g > 100 and g > r + 30) or (b > 120 and b > r + 30):
                    return "TH"
        except Exception:
            pass
        return "LT"

    def _build_schedule_color_map(self, path):
        """Trả về dict {(excel_row, excel_col_1idx): 'LT'|'KT'|'TH'}."""
        color_map = {}
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=1):
                for cell in row:
                    v = cell.value
                    if v is None or str(v).strip() == "":
                        continue
                    color_map[(cell.row, cell.column)] = self._classify_cell_color(cell)
        except Exception as e:
            print(f"[color_map] {e}")
        return color_map

    def render_month(self):
        for item in self.month_tree.get_children():
            self.month_tree.delete(item)

        path = self.config_data.get("schedule_file", "schedule.xlsx")
        if not os.path.exists(path):
            self.month_info.configure(
                text=f"Không tìm thấy '{path}'. Vào Cài đặt để chọn lại.")
            self.month_tree.configure(columns=())
            return

        try:
            import re as _re

            # Đọc 3 dòng đầu để lấy tiêu đề tháng/năm
            try:
                head_raw = pd.read_excel(path, header=None, nrows=3).fillna("")
                head_text = " ".join(str(v) for row in head_raw.values for v in row)
                m_year = _re.search(r"Th[áa]ng\s+(\d{1,2}).*?N[ăa]m\s+(\d{4})",
                                     head_text, _re.IGNORECASE)
                if m_year:
                    self.lbl_excel_subtitle.configure(
                        text=f"Tháng {m_year.group(1)} Năm {m_year.group(2)}")
            except Exception:
                pass

            # Map màu (theo Excel coords)
            color_map = self._build_schedule_color_map(path)
            self._last_color_map = color_map  # save for tab Báo cáo

            raw = pd.read_excel(path, skiprows=3, header=None)
            # KHÔNG dropna+reset để giữ mapping raw_idx → excel_row (= idx + 4)
            if len(raw) < 2:
                self.month_info.configure(text="File rỗng")
                self.month_tree.configure(columns=())
                return

            header_row = [str(v) if not pd.isna(v) else "" for v in raw.iloc[0]]
            weekday_row = []
            data_start_offset = 1
            if len(raw) > 1:
                second = raw.iloc[1]
                if pd.isna(second.iloc[0]) or str(second.iloc[0]).strip() in ("", "nan"):
                    weekday_row = [str(v) if not pd.isna(v) else "" for v in second]
                    data_start_offset = 2

            # data: giữ Excel row mapping qua excel_rows_list
            data_raw = raw.iloc[data_start_offset:]
            data_raw.columns = range(len(header_row))
            data_raw = data_raw.dropna(how="all")
            # excel_row của mỗi data row = original_index + 4 (vì skiprows=3)
            excel_rows_list = [i + 4 for i in data_raw.index]
            data = data_raw.reset_index(drop=True)
            if data.empty:
                self.month_info.configure(text="Không có dữ liệu")
                return

            # Giữ bản gốc để phát hiện ranh giới GV (trước ffill)
            raw_tt = data[0].copy() if 0 in data.columns else None
            raw_name = data[1].copy() if 1 in data.columns else None
            raw_subject = data[2].copy() if 2 in data.columns else None

            search = self.month_search.get().strip().lower() if hasattr(self, "month_search") else ""
            subject_filter = self.month_subject.get() if hasattr(self, "month_subject") else "Tất cả"
            hide_empty = bool(self.hide_empty_var.get()) if hasattr(self, "hide_empty_var") else False

            today = datetime.now()
            today_key = today.strftime("%Y-%m-%d")

            weekday_map_vn = {0: "T2", 1: "T3", 2: "T4", 3: "T5", 4: "T6", 5: "T7", 6: "CN"}

            col_ids = [f"c{i}" for i in range(len(header_row))]
            self.month_tree.configure(columns=col_ids)

            def fmt_date_col(orig):
                m = _re.match(r"(\d{4})-(\d{2})-(\d{2})", str(orig))
                if not m:
                    return str(orig).replace("\n", " ").strip(), False, False
                yyyy, mm, dd = m.group(1), m.group(2), m.group(3)
                date_obj = datetime(int(yyyy), int(mm), int(dd))
                wk_label = weekday_map_vn.get(date_obj.weekday(), "")
                is_weekend = date_obj.weekday() >= 5
                is_today = (today.year == date_obj.year
                            and today.month == date_obj.month
                            and today.day == date_obj.day)
                mark = " ●" if is_today else ""
                label = f"{int(dd):02d}/{int(mm):02d} {wk_label}{mark}"
                return label, is_weekend, is_today

            for i, orig in enumerate(header_row):
                cid = col_ids[i]
                up = str(orig).upper()
                if up == "TT":
                    title = "TT"
                    self.month_tree.column(cid, width=44, minwidth=40,
                                           anchor="center", stretch=False)
                elif "HỌ VÀ TÊN" in up:
                    title = "Họ và tên"
                    self.month_tree.column(cid, width=180, minwidth=140,
                                           anchor="w", stretch=False)
                elif "MÔN" in up:
                    title = "Môn"
                    self.month_tree.column(cid, width=110, minwidth=80,
                                           anchor="center", stretch=False)
                elif _re.match(r"(\d{4})-(\d{2})-(\d{2})", str(orig)):
                    label, is_weekend, is_today = fmt_date_col(orig)
                    title = label
                    width = 110 if is_today else 100
                    self.month_tree.column(cid, width=width, minwidth=80,
                                           anchor="center", stretch=False)
                else:
                    title = str(orig).replace("\n", " ").strip() or " "
                    self.month_tree.column(cid, width=72, minwidth=60,
                                           anchor="center", stretch=False)
                self.month_tree.heading(cid, text=title)

            valid_slot_re = _re.compile(r"^\d+\s*-\s*\d+$")
            # Tên section (header phân nhóm trong file Excel) - khớp chính xác,
            # không substring để tránh loại nhầm tên người chứa các chữ này
            section_exact = {
                "SÁNG", "CHIỀU", "BUỔI SÁNG", "BUỔI CHIỀU",
                "TỔNG", "CỘNG", "TỔNG CỘNG", "TỔNG SỐ",
                "QUÂN SỰ", "QUỐC TẾ", "CÔNG AN",
                "QUÂN SỰ + QUỐC TẾ", "QUÂN SỰ + QUỐC TẾ:",
            }
            section_contains = ("THỐNG KÊ", "GHI CHÚ")

            def is_real_teacher(nm):
                up = str(nm).strip().upper()
                if len(up) < 2:
                    return False
                if up in section_exact:
                    return False
                if any(k in up for k in section_contains):
                    return False
                if up.endswith(":"):
                    return False
                # "TỔNG SỐ TIẾT", "CỘNG ..." - bắt đầu bằng từ khoá section
                first_word = up.split()[0] if up.split() else ""
                if first_word in {"TỔNG", "CỘNG"}:
                    return False
                return True

            teachers = []
            current = None
            for idx, r in data.iterrows():
                orig_name = clean_numeric_text(raw_name.iloc[idx] if raw_name is not None else "")
                orig_tt = clean_numeric_text(raw_tt.iloc[idx] if raw_tt is not None else "")
                orig_subject = clean_numeric_text(raw_subject.iloc[idx] if raw_subject is not None else "")
                slot_raw = clean_numeric_text(r[3] if len(r) > 3 else "")
                slot_norm = slot_raw.replace(" ", "")
                has_valid_slot = bool(valid_slot_re.match(slot_norm))

                if orig_name:
                    if is_real_teacher(orig_name):
                        current = {
                            "tt": orig_tt,
                            "name": orig_name,
                            "subject": orig_subject,
                            "rows": [],
                            "slots_seen": set(),
                        }
                        teachers.append(current)
                    else:
                        current = None

                if current is None or not has_valid_slot:
                    continue
                if slot_norm in current["slots_seen"] or len(current["rows"]) >= 4:
                    current = None
                    continue
                current["slots_seen"].add(slot_norm)
                current["rows"].append(r)

            teachers = [t for t in teachers if t["rows"]]

            # Gộp các block cùng tên GV: 1 GV dạy nhiều môn → 1 entry,
            # hợp nhất rows theo cặp tiết (1-2, 3-4, ...) và join data các ngày
            from collections import OrderedDict

            # Marker theo loại tiết (đọc từ màu Excel)
            TYPE_PREFIX = {"KT": "🔴 ", "TH": "🟢 ", "LT": ""}

            merged = OrderedDict()
            for t in teachers:
                key = t["name"].strip().upper()
                if key not in merged:
                    merged[key] = {
                        "tt": t["tt"],
                        "name": t["name"],
                        "subjects": [],
                        "slot_rows": OrderedDict(),
                        "type_counts": {"LT": 0, "KT": 0, "TH": 0},
                    }
                m = merged[key]
                if t["subject"] and t["subject"] not in m["subjects"]:
                    m["subjects"].append(t["subject"])
                for r in t["rows"]:
                    excel_row = excel_rows_list[r.name] if r.name < len(excel_rows_list) else None
                    slot_raw = clean_numeric_text(r[3] if len(r) > 3 else "")
                    slot_key = slot_raw.replace(" ", "") or f"slot{len(m['slot_rows'])}"
                    if slot_key not in m["slot_rows"]:
                        m["slot_rows"][slot_key] = {
                            "slot_text": slot_raw,
                            "cells": [""] * len(header_row),
                        }
                    cells = m["slot_rows"][slot_key]["cells"]
                    for i in range(len(header_row)):
                        v = clean_numeric_text(r[i] if i < len(r) else "").replace("\n", " / ")
                        if not v:
                            continue
                        if i < 4:
                            cells[i] = v
                        else:
                            ctype = "LT"
                            if excel_row is not None and color_map:
                                ctype = color_map.get((excel_row, i + 1), "LT")
                            m["type_counts"][ctype] += 1
                            v_marked = TYPE_PREFIX[ctype] + v
                            if cells[i]:
                                if v_marked not in cells[i].split(" / "):
                                    cells[i] = cells[i] + " / " + v_marked
                            else:
                                cells[i] = v_marked

            teachers_merged = list(merged.values())
            # Lưu lại cho tab Báo cáo tuần
            self._last_teachers_merged = teachers_merged
            self._last_header_row = header_row

            subjects_seen = set()
            for t in teachers_merged:
                for s in t["subjects"]:
                    if s and not s.endswith(":"):
                        subjects_seen.add(s.upper())
            if hasattr(self, "month_subject_menu"):
                menu_values = ["Tất cả"] + sorted(subjects_seen)
                try:
                    self.month_subject_menu.configure(values=menu_values)
                except Exception:
                    pass
                if subject_filter not in menu_values:
                    self.month_subject.set("Tất cả")
                    subject_filter = "Tất cả"

            total_rows = 0
            total_teachers = 0
            for t in teachers_merged:
                if search and search not in t["name"].lower():
                    continue
                if subject_filter != "Tất cả" and not any(
                        s.upper() == subject_filter.upper() for s in t["subjects"]):
                    continue

                total_teachers += 1
                subjects_str = " / ".join(t["subjects"])

                first_in_block = True
                for slot_key, srow in t["slot_rows"].items():
                    values = list(srow["cells"])

                    if hide_empty:
                        has_data = any(values[i].strip() for i in range(4, len(values)))
                        if not has_data:
                            continue

                    if not first_in_block:
                        values[0] = ""
                        values[1] = ""
                        values[2] = ""
                    else:
                        values[0] = t["tt"]
                        values[1] = t["name"].upper()
                        values[2] = subjects_str

                    tags = []
                    if first_in_block:
                        tags.append("group_top")
                    elif total_rows % 2 == 1:
                        tags.append("alt")

                    self.month_tree.insert("", "end", values=values, tags=tuple(tags))
                    total_rows += 1
                    first_in_block = False

            hint = " · Hôm nay: " + today.strftime("%d/%m")
            self.month_info.configure(
                text=f"{total_teachers} giảng viên · {total_rows} tiết{hint}")
        except Exception as e:
            self.month_info.configure(text=f"Lỗi đọc file: {e}")
    # --- TAB: QUẢN LÝ CHUNG ---
    def render_mgmt(self):
        for widget in self.mgmt_scroll.winfo_children():
            widget.destroy()

        if not self.mgmt_data:
            ctk.CTkLabel(self.mgmt_scroll, text="Chưa có dữ liệu giảng viên",
                         font=("Arial", 13), text_color=COLORS["text_dim"]).pack(pady=40)
            if hasattr(self, "mgmt_count"):
                self.mgmt_count.configure(text="")
            return

        search = self.mgmt_search.get().lower().strip()

        # Tự dò cột chức vụ
        position_keys = ("CHỨC VỤ", "CHỨC DANH", "CHỨC VỤ HIỆN TẠI", "VỊ TRÍ")
        sample = self.mgmt_data[0]
        position_col = next((k for k in position_keys if k in sample), None)

        # Thứ tự ưu tiên các chức vụ (lower-cased keyword → priority)
        position_order = [
            ("trưởng bộ môn", 0),
            ("phó trưởng bộ môn", 1),
            ("phó bộ môn", 1),
            ("trưởng khoa", 2),
            ("phó khoa", 3),
            ("giảng viên", 4),
        ]
        OTHER_PRIORITY = 5

        def get_priority(pos):
            p = (pos or "").lower().strip()
            for kw, pr in position_order:
                if kw in p:
                    return pr
            return OTHER_PRIORITY

        def normalize_position(pos):
            p = (pos or "").strip()
            if not p or p.lower() == "nan":
                return "Khác"
            return p

        # Group by position
        from collections import defaultdict
        groups = defaultdict(list)
        order_priority = {}
        for row in self.mgmt_data:
            name = str(row.get('HỌ VÀ TÊN', '')).strip()
            if not name or name.lower() == "nan":
                continue
            if search and search not in name.lower():
                continue

            raw_pos = str(row.get(position_col, "")).strip() if position_col else ""
            pos_label = normalize_position(raw_pos)
            groups[pos_label].append(row)
            order_priority[pos_label] = min(order_priority.get(pos_label,
                                                                OTHER_PRIORITY),
                                             get_priority(raw_pos))

        sorted_groups = sorted(groups.items(),
                                key=lambda x: (order_priority.get(x[0], OTHER_PRIORITY),
                                               x[0].lower()))

        total_shown = 0
        for pos_label, members in sorted_groups:
            # Group header
            header = ctk.CTkFrame(self.mgmt_scroll, fg_color="transparent")
            header.pack(fill="x", pady=(8, 4), padx=2)
            ctk.CTkLabel(header,
                          text=f"▸ {pos_label}  ({len(members)})",
                          font=("Arial", 15, "bold"),
                          text_color=COLORS["accent"], anchor="w"
                          ).pack(side="left")

            for row in members:
                name = str(row.get('HỌ VÀ TÊN', '')).strip().upper()
                card = ctk.CTkFrame(self.mgmt_scroll, fg_color=COLORS["card"],
                                     height=52, corner_radius=8,
                                     border_width=1, border_color=COLORS["border"])
                card.pack(fill="x", pady=4, padx=2)
                card.pack_propagate(False)

                ctk.CTkLabel(card, text=name, font=("Arial", 15, "bold"),
                              text_color=COLORS["text"]
                              ).pack(side="left", padx=16)

                rank = str(row.get('CẤP BẬC', '')).strip()
                if rank and rank.lower() != "nan":
                    ctk.CTkLabel(card, text=rank, font=("Arial", 13),
                                  text_color=COLORS["text_dim"]
                                  ).pack(side="left", padx=(0, 10))

                ctk.CTkButton(card, text="Chi tiết", width=90, height=32,
                               fg_color=COLORS["accent"], hover_color="#1D4ED8",
                               font=("Arial", 13, "bold"),
                               command=lambda r=row: TeacherDetailWindow(self, r)
                               ).pack(side="right", padx=12)
                total_shown += 1

        if hasattr(self, "mgmt_count"):
            total = len(self.mgmt_data)
            self.mgmt_count.configure(text=f"{total_shown}/{total} giảng viên")
    def show_mgmt_frame(self):
        self.hide_all_frames()
        self.mgmt_frame.pack(fill="both", expand=True)
        self.set_active_nav("mgmt")

    def setup_dashboard_ui(self):
        header = ctk.CTkFrame(self.dashboard_frame, fg_color="transparent")
        header.pack(fill="x", pady=(0, 14))
        ctk.CTkLabel(header, text="Bảng điều khiển", font=("Arial", 26, "bold"),
                     text_color=COLORS["text"]).pack(side="left")
        ctk.CTkLabel(header, text=datetime.now().strftime("%A, %d/%m/%Y"),
                     font=("Arial", 14), text_color=COLORS["text_dim"]
                     ).pack(side="right")

        welcome = ctk.CTkFrame(self.dashboard_frame,
                               fg_color=("#EFF6FF", "#1E3A8A"),
                               corner_radius=12, border_width=0)
        welcome.pack(fill="x", pady=(0, 14))
        welcome_pad = ctk.CTkFrame(welcome, fg_color="transparent")
        welcome_pad.pack(fill="x", padx=20, pady=16)
        ctk.CTkLabel(welcome_pad,
                     text="Xin chào 👋",
                     font=("Arial", 20, "bold"),
                     text_color=COLORS["accent"], anchor="w"
                     ).pack(fill="x")
        ctk.CTkLabel(welcome_pad,
                     text="Quản lý thông tin giảng viên, kế hoạch giảng dạy và tài liệu môn học.",
                     font=("Arial", 14),
                     text_color=COLORS["text_dim"], anchor="w"
                     ).pack(fill="x", pady=(4, 0))

        grid = ctk.CTkFrame(self.dashboard_frame, fg_color="transparent")
        grid.pack(fill="x", pady=(0, 14))
        for i in range(4):
            grid.grid_columnconfigure(i, weight=1, uniform="stat")

        self.stat_widgets = {}
        stats = [
            ("teachers", "Tổng giảng viên", "0", "#2563EB", "👥"),
            ("subjects", "Số môn học", "0", "#10B981", "📚"),
            ("today", "Tiết dạy hôm nay", "0", "#F59E0B", "⏰"),
            ("files", "Tài liệu", "0", "#8B5CF6", "📂"),
        ]
        for i, (key, title, value, color, icon) in enumerate(stats):
            card = ctk.CTkFrame(grid, fg_color=COLORS["card"], corner_radius=12,
                                border_width=1, border_color=COLORS["border"])
            card.grid(row=0, column=i, sticky="nsew", padx=6)

            top = ctk.CTkFrame(card, fg_color="transparent")
            top.pack(fill="x", padx=18, pady=(16, 4))
            ctk.CTkLabel(top, text=title, font=("Arial", 13),
                         text_color=COLORS["text_dim"], anchor="w"
                         ).pack(side="left", fill="x", expand=True)
            icon_bg = ctk.CTkFrame(top, fg_color=color, corner_radius=8,
                                   width=38, height=38)
            icon_bg.pack(side="right")
            icon_bg.pack_propagate(False)
            ctk.CTkLabel(icon_bg, text=icon, font=("Arial", 17),
                         text_color="white").pack(expand=True)

            value_lbl = ctk.CTkLabel(card, text=value,
                                     font=("Arial", 32, "bold"),
                                     text_color=COLORS["text"], anchor="w")
            value_lbl.pack(fill="x", padx=18, pady=(0, 16))
            self.stat_widgets[key] = value_lbl

        shortcuts = ctk.CTkFrame(self.dashboard_frame, fg_color=COLORS["card"],
                                 corner_radius=12, border_width=1,
                                 border_color=COLORS["border"])
        shortcuts.pack(fill="both", expand=True)
        ctk.CTkLabel(shortcuts, text="Truy cập nhanh",
                     font=("Arial", 17, "bold"),
                     text_color=COLORS["text"], anchor="w"
                     ).pack(fill="x", padx=20, pady=(16, 10))

        shortcut_grid = ctk.CTkFrame(shortcuts, fg_color="transparent")
        shortcut_grid.pack(fill="x", padx=12, pady=(0, 18))
        for i in range(4):
            shortcut_grid.grid_columnconfigure(i, weight=1, uniform="short")

        shortcut_defs = [
            ("Thông tin giảng viên", "Xem danh sách", self.show_mgmt_frame),
            ("Kế hoạch ngày", "Lịch hôm nay", self.show_plan_frame),
            ("Kế hoạch tháng", "Bảng tháng", self.show_month_frame),
            ("Môn học", "Mở thư mục", self.show_document_frame),
        ]
        for i, (title, sub, cmd) in enumerate(shortcut_defs):
            btn = ctk.CTkButton(shortcut_grid, text="",
                                fg_color="transparent",
                                hover_color=COLORS["hover"],
                                corner_radius=10, height=80,
                                border_width=1, border_color=COLORS["border"],
                                command=cmd)
            btn.grid(row=0, column=i, sticky="ew", padx=6, pady=4)

            inner = ctk.CTkFrame(btn, fg_color="transparent")
            inner.place(relx=0.5, rely=0.5, anchor="center")
            ctk.CTkLabel(inner, text=title, font=("Arial", 14, "bold"),
                         text_color=COLORS["text"]).pack()
            ctk.CTkLabel(inner, text=sub, font=("Arial", 12),
                         text_color=COLORS["text_dim"]).pack()

    def show_dashboard_frame(self):
        self.hide_all_frames()
        self.dashboard_frame.pack(fill="both", expand=True)
        self.set_active_nav("dashboard")
        self.refresh_dashboard_stats()

    def refresh_dashboard_stats(self):
        if not hasattr(self, "stat_widgets"):
            return
        teachers = sum(1 for r in self.mgmt_data if str(r.get('HỌ VÀ TÊN', '')).strip())
        subjects = set()
        for r in self.mgmt_data:
            s = str(r.get('MÔN DẠY', '') or r.get('MÔN HỌC', '') or '').strip()
            if s and s.lower() != 'nan':
                subjects.add(s)
        today_count = 0
        for r in self.plan_data:
            for slot in ("1 - 2", "3 - 4", "5 - 6", "7 - 8"):
                v = str(r.get(slot, "")).strip()
                if v and v.lower() != "nan":
                    today_count += 1
        files = 0
        doc_folder = self.config_data.get("document_folder", "Document")
        if os.path.exists(doc_folder):
            for root, dirs, fs in os.walk(doc_folder):
                dirs[:] = [d for d in dirs if not d.startswith('.')]
                for f in fs:
                    if not f.startswith('.') and not f.startswith('~$'):
                        files += 1

        self.stat_widgets["teachers"].configure(text=str(teachers))
        self.stat_widgets["subjects"].configure(text=str(len(subjects)))
        self.stat_widgets["today"].configure(text=str(today_count))
        self.stat_widgets["files"].configure(text=str(files))

    def setup_settings_ui(self):
        header = ctk.CTkFrame(self.settings_frame, fg_color="transparent")
        header.pack(fill="x", pady=(0, 14))
        ctk.CTkLabel(header, text="Cài đặt", font=("Arial", 26, "bold"),
                     text_color=COLORS["text"]).pack(side="left")
        ctk.CTkLabel(header,
                     text="Cấu hình đường dẫn file và tuỳ chọn hiển thị",
                     font=("Arial", 12),
                     text_color=COLORS["text_dim"]).pack(side="left", padx=(12, 0))

        self.settings_vars = {}

        def add_section(title):
            section = ctk.CTkFrame(self.settings_frame, fg_color=COLORS["card"],
                                    corner_radius=12, border_width=1,
                                    border_color=COLORS["border"])
            section.pack(fill="x", pady=(0, 12))
            ctk.CTkLabel(section, text=title, font=("Arial", 14, "bold"),
                         text_color=COLORS["text"], anchor="w"
                         ).pack(fill="x", padx=18, pady=(14, 8))
            return section

        def add_file_row(parent, label, key, mode="file",
                         filetypes=(("Excel", "*.xlsx *.xls"),)):
            row = ctk.CTkFrame(parent, fg_color="transparent")
            row.pack(fill="x", padx=18, pady=(0, 10))
            ctk.CTkLabel(row, text=label, font=("Arial", 12),
                         text_color=COLORS["text"], width=170, anchor="w"
                         ).pack(side="left")
            var = tk.StringVar(value=self.config_data.get(key, ""))
            self.settings_vars[key] = var
            entry = ctk.CTkEntry(row, textvariable=var, height=34,
                                 border_color=COLORS["border"])
            entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

            def pick():
                if mode == "folder":
                    p = filedialog.askdirectory(initialdir=".")
                else:
                    p = filedialog.askopenfilename(filetypes=filetypes,
                                                   initialdir=".")
                if p:
                    rel = os.path.relpath(p, os.getcwd())
                    if not rel.startswith(".."):
                        p = rel
                    var.set(p)

            ctk.CTkButton(row, text="Chọn...", width=84, height=34,
                          fg_color="transparent",
                          text_color=COLORS["text"],
                          border_width=1, border_color=COLORS["border"],
                          hover_color=COLORS["hover"], command=pick
                          ).pack(side="left")

        files_section = add_section("Đường dẫn dữ liệu")
        add_file_row(files_section, "File danh sách GV", "teacher_file",
                     mode="file", filetypes=(("Excel", "*.xlsx *.xls"),))
        add_file_row(files_section, "File kế hoạch tháng", "schedule_file",
                     mode="file", filetypes=(("Excel", "*.xlsx *.xls"),))
        add_file_row(files_section, "Thư mục tài liệu", "document_folder",
                     mode="folder")

        pref_section = add_section("Hiển thị")
        row = ctk.CTkFrame(pref_section, fg_color="transparent")
        row.pack(fill="x", padx=18, pady=(0, 10))
        ctk.CTkLabel(row, text="Chế độ giao diện", font=("Arial", 12),
                     text_color=COLORS["text"], width=170, anchor="w"
                     ).pack(side="left")
        appear_var = tk.StringVar(value=self.config_data.get("appearance", "light"))
        self.settings_vars["appearance"] = appear_var
        appear_menu = ctk.CTkOptionMenu(row, variable=appear_var,
                                        values=["light", "dark", "system"],
                                        fg_color=COLORS["accent"],
                                        button_color=COLORS["accent"],
                                        button_hover_color="#1D4ED8",
                                        width=140)
        appear_menu.pack(side="left")

        row2 = ctk.CTkFrame(pref_section, fg_color="transparent")
        row2.pack(fill="x", padx=18, pady=(0, 14))
        ctk.CTkLabel(row2, text="Tự cập nhật lịch ngày",
                     font=("Arial", 12),
                     text_color=COLORS["text"], width=170, anchor="w"
                     ).pack(side="left")
        auto_var = tk.BooleanVar(value=bool(self.config_data.get("auto_update_schedule", True)))
        self.settings_vars["auto_update_schedule"] = auto_var
        ctk.CTkSwitch(row2, text="", variable=auto_var,
                      progress_color=COLORS["accent"]).pack(side="left")

        actions = ctk.CTkFrame(self.settings_frame, fg_color="transparent")
        actions.pack(fill="x", pady=(4, 0))
        self.settings_status = ctk.CTkLabel(actions, text="",
                                            font=("Arial", 11),
                                            text_color=COLORS["text_dim"])
        self.settings_status.pack(side="left")
        ctk.CTkButton(actions, text="Lưu cài đặt", width=130, height=36,
                      fg_color=COLORS["accent"], hover_color="#1D4ED8",
                      font=("Arial", 12, "bold"),
                      command=self.save_settings).pack(side="right")
        ctk.CTkButton(actions, text="Tải lại dữ liệu", width=130, height=36,
                      fg_color="transparent", text_color=COLORS["text"],
                      border_width=1, border_color=COLORS["border"],
                      hover_color=COLORS["hover"],
                      command=self.reload_all_data).pack(side="right", padx=(0, 8))

    def show_settings_frame(self):
        self.hide_all_frames()
        self.settings_frame.pack(fill="both", expand=True)
        self.set_active_nav("settings")

    def save_settings(self):
        for key, var in self.settings_vars.items():
            try:
                self.config_data[key] = var.get()
            except Exception:
                pass
        ok = AppConfig.save(self.config_data)
        if ok:
            self.settings_status.configure(
                text=f"Đã lưu · {datetime.now().strftime('%H:%M:%S')}",
                text_color=COLORS["success"])
            try:
                ctk.set_appearance_mode(self.config_data.get("appearance", "light"))
            except Exception:
                pass
        else:
            self.settings_status.configure(text="Lỗi khi lưu config.json",
                                           text_color=COLORS["error"])

    def reload_all_data(self):
        self.auto_load_mgmt_file()
        self.check_realtime_status()
        if hasattr(self, "month_tree"):
            self.render_month()
        if hasattr(self, "document_scroll"):
            self.render_documents()
        self.refresh_dashboard_stats()
        self.settings_status.configure(
            text=f"Đã tải lại · {datetime.now().strftime('%H:%M:%S')}",
            text_color=COLORS["success"])

    def setup_mgmt_ui(self):
        self.clear_right_frame()
        header = ctk.CTkFrame(self.mgmt_frame, fg_color="transparent")
        header.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(header, text="Thông tin giảng viên", font=("Arial", 26, "bold"),
                     text_color=COLORS["text"]).pack(side="left")
        self.mgmt_count = ctk.CTkLabel(header, text="", font=("Arial", 13),
                                       text_color=COLORS["text_dim"])
        self.mgmt_count.pack(side="right")

        self.mgmt_search = ctk.CTkEntry(self.mgmt_frame,
                                        placeholder_text="Tìm theo tên giảng viên...",
                                        height=40, font=("Arial", 13),
                                        border_color=COLORS["border"])
        self.mgmt_search.pack(fill="x", pady=(0, 10))
        self.mgmt_search.bind("<KeyRelease>", lambda e: self.render_mgmt())

        self.mgmt_scroll = ctk.CTkScrollableFrame(self.mgmt_frame, fg_color="transparent")
        self.mgmt_scroll.pack(fill="both", expand=True)
    def link_mgmt(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            self.mgmt_path = path
            self.refresh_mgmt()
    def refresh_mgmt(self):

        path = self.config_data.get("teacher_file", "danh sách k8.xlsx")
        if not os.path.exists(path):
            print(f"Không tìm thấy file {path}!")
            return

        try:
            # 1. Đọc file với engine openpyxl (quan trọng)
            # skiprows=2: Bỏ qua các dòng tiêu đề rỗng phía trên
            df = pd.read_excel(path, skiprows=2, engine='openpyxl', dtype=str)

            # 2. Chuẩn hóa tên cột: Xóa khoảng trắng và viết HOA toàn bộ
            df.columns = [str(c).strip().upper() for c in df.columns]

            # 3. Làm sạch dữ liệu: Xử lý gộp ô (ffill) và xóa dòng trống
            if 'HỌ VÀ TÊN' in df.columns:
                df['HỌ VÀ TÊN'] = df['HỌ VÀ TÊN'].ffill() # Điền tên cho các ô bị gộp
                df = df.dropna(subset=['HỌ VÀ TÊN']) # Xóa dòng rác

            # 4. QUAN TRỌNG: Lưu vào biến self để các hàm khác có thể dùng
            self.mgmt_data = df.to_dict('records')
            
            # 5. Sau khi nhận được dữ liệu, gọi hàm vẽ giao diện ngay
            self.render_mgmt()
            print(f"Đã nhận {len(self.mgmt_data)} giảng viên từ file.")

        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể nhận dữ liệu từ file: {e}")
    def auto_load_mgmt_file(self):
        """Tự động tìm tiêu đề và nạp dữ liệu chính xác"""
        file_name = self.config_data.get("teacher_file", "danh sách k8.xlsx")
        if not os.path.exists(file_name):
            print(f"Không tìm thấy file tại: {os.path.abspath(file_name)}")
            return

        try:
            # 1. Đọc nháp toàn bộ file (không dùng header) để tìm dòng tiêu đề
            raw_df = pd.read_excel(file_name, header=None, engine='openpyxl')
            
            header_row_index = None
            # Quét qua 20 dòng đầu tiên để tìm chữ "HỌ VÀ TÊN"
            for i, row in raw_df.head(20).iterrows():
                # Chuyển tất cả giá trị trong dòng thành chữ HOA, xóa khoảng trắng để so sánh
                row_values = [str(val).strip().upper() for val in row.values]
                if "HỌ VÀ TÊN" in row_values:
                    header_row_index = i
                    print(f"🎯 Đã tìm thấy tiêu đề 'HỌ VÀ TÊN' tại dòng thứ: {i + 1}")
                    break
            
            if header_row_index is None:
                print("❌ Vẫn không tìm thấy cột 'HỌ VÀ TÊN'.")
                print(f"Dữ liệu 5 dòng đầu đọc được:\n{raw_df.head(5)}")
                return

            # 2. Đọc lại file thật sự bắt đầu từ dòng tiêu đề đã tìm thấy
            df = pd.read_excel(file_name, skiprows=header_row_index,
                               engine='openpyxl', dtype=str)

            # 3. Chuẩn hóa tên cột một lần nữa cho chắc chắn
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            # 4. Làm sạch dữ liệu
            # Loại bỏ các cột "Unnamed" (cột thừa không có tên)
            df = df.loc[:, ~df.columns.str.contains('^UNNAMED')]
            
            # Điền đầy dữ liệu nếu có gộp ô (Merge Cells)
            if 'HỌ VÀ TÊN' in df.columns:
                df['HỌ VÀ TÊN'] = df['HỌ VÀ TÊN'].ffill()
                df = df.dropna(subset=['HỌ VÀ TÊN']) # Xóa dòng hoàn toàn trống
                
                # Chuyển đổi sang danh sách Dictionary để dùng cho app
                self.mgmt_data = df.to_dict('records')
                
                # 5. Cập nhật giao diện
                self.render_mgmt()
                print(f"✅ Nạp thành công {len(self.mgmt_data)} giảng viên.")
            
        except Exception as e:
            print(f"❌ Lỗi xử lý: {e}")
    def process_mgmt_file(self, path):
        try:
            # 1. Đọc toàn bộ file không bỏ qua dòng nào để dò tìm
            raw_df = pd.read_excel(path, header=None, engine='openpyxl')
            
            header_row_index = None
            
            # 2. Vòng lặp tìm dòng chứa từ khóa "HỌ VÀ TÊN"
            for i, row in raw_df.iterrows():
                # Chuyển dòng thành danh sách chữ HOA để so sánh
                row_values = [str(val).strip().upper() for val in row.values]
                if "HỌ VÀ TÊN" in row_values:
                    header_row_index = i
                    break
            
            if header_row_index is None:
                print(f"❌ Không tìm thấy dòng nào chứa cột 'HỌ VÀ TÊN' trong file {path}")
                return

            # 3. Đọc lại file với đúng dòng tiêu đề đã tìm thấy
            df = pd.read_excel(path, skiprows=header_row_index, engine='openpyxl')
            
            # 4. Chuẩn hóa tên cột (Xóa khoảng trắng, viết HOA)
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            # 5. Làm sạch dữ liệu rác
            # Điền đầy dữ liệu gộp ô (Merge cells)
            if 'HỌ VÀ TÊN' in df.columns:
                df['HỌ VÀ TÊN'] = df['HỌ VÀ TÊN'].ffill()
                df = df.dropna(subset=['HỌ VÀ TÊN']) # Bỏ dòng trống hoàn toàn
                
                # Chuyển thành List Dict để dùng cho App
                self.mgmt_data = df.to_dict('records')
                
                # Vẽ lên màn hình
                self.render_mgmt()
                print(f"✅ Đã tìm thấy tiêu đề ở dòng {header_row_index + 1} và nạp thành công!")
            else:
                print("❌ Lỗi logic: Đã tìm thấy dòng tiêu đề nhưng không khớp cột.")

        except Exception as e:
            print(f"❌ Lỗi xử lý file: {e}")
   
#--------------------------
    def show_plan_frame(self):
        self.hide_all_frames()
        self.plan_frame.pack(fill="both", expand=True)
        self.set_active_nav("plan")

    def render_plan(self):
        try:
            if not self.plan_scroll.winfo_exists():
                return
            for child in self.plan_scroll.winfo_children():
                child.destroy()

            if not self.plan_data:
                ctk.CTkLabel(self.plan_scroll, text="Chưa có dữ liệu kế hoạch ngày",
                             font=("Arial", 13),
                             text_color=COLORS["text_dim"]).pack(pady=40)
                return

            SUB_COLORS = {
                "BC": ("#E0F2FE", "#0369A1", "#2563EB"),
                "ĐH": ("#DCFCE7", "#15803D", "#10B981"),
                "KB": ("#F3E8FF", "#7E22CE", "#8B5CF6"),
                "ĐN": ("#FEF3C7", "#B45309", "#F59E0B"),
            }

            def get_v(r, k):
                v = str(r.get(k, "")).strip()
                return "" if v.lower() == "nan" or v == "" else v

            # Gom dữ liệu theo giảng viên, lọc bỏ entry không có tiết
            groups = []
            current = None
            for row in self.plan_data:
                name = get_v(row, "Họ và tên")
                subject = get_v(row, "môn học")
                slots = [get_v(row, s) for s in ("1 - 2", "3 - 4", "5 - 6", "7 - 8")]
                if not any(slots):
                    continue
                if not name and not subject:
                    continue

                if name and (current is None or current["name"] != name):
                    current = {"name": name, "rows": []}
                    groups.append(current)
                if current is None:
                    continue
                current["rows"].append({"subject": subject, "slots": slots})

            if not groups:
                ctk.CTkLabel(self.plan_scroll, text="Hôm nay không có tiết nào",
                             font=("Arial", 13),
                             text_color=COLORS["text_dim"]).pack(pady=40)
                return

            # Header bảng
            header_f = ctk.CTkFrame(self.plan_scroll, fg_color="#E0E7FF",
                                    height=40, corner_radius=0)
            header_f.pack(fill="x", pady=(0, 4))
            header_f.pack_propagate(False)
            COLS = [("Họ và tên", 0.02, 0.26), ("Môn", 0.29, 0.08),
                    ("Tiết 1-2", 0.40, 0.15), ("Tiết 3-4", 0.55, 0.15),
                    ("Tiết 5-6", 0.70, 0.15), ("Tiết 7-8", 0.85, 0.15)]
            for txt, rx, rw in COLS:
                ctk.CTkLabel(header_f, text=txt, font=("Arial", 13, "bold"),
                             text_color=COLORS["text"], anchor="w"
                             ).place(relx=rx, rely=0.5, anchor="w", relwidth=rw)

            # Render từng nhóm GV
            for g in groups:
                primary_sub = g["rows"][0]["subject"].upper() if g["rows"] else ""
                border_c = SUB_COLORS.get(primary_sub, (None, None, "#CBD5E1"))[2]

                card = ctk.CTkFrame(self.plan_scroll, fg_color="white",
                                     corner_radius=8,
                                     border_width=2,
                                     border_color=border_c)
                card.pack(fill="x", padx=2, pady=(0, 6))

                for ri, entry in enumerate(g["rows"]):
                    row_bg = "white" if ri % 2 == 0 else "#F8FAFC"
                    row_f = ctk.CTkFrame(card, fg_color=row_bg,
                                          height=42, corner_radius=0)
                    row_f.pack(fill="x", padx=2, pady=(2 if ri == 0 else 0,
                                                        2 if ri == len(g["rows"]) - 1 else 0))
                    row_f.pack_propagate(False)

                    display_name = g["name"].upper() if ri == 0 else ""
                    ctk.CTkLabel(row_f, text=display_name,
                                 font=("Arial", 15, "bold"),
                                 text_color=COLORS["text"], anchor="w"
                                 ).place(relx=0.02, rely=0.5, anchor="w",
                                         relwidth=0.28)

                    subject = entry["subject"]
                    if subject:
                        bg_c, fg_c, _ = SUB_COLORS.get(subject.upper(),
                                                        ("#F1F5F9", "#475569", "#94A3B8"))
                        badge = ctk.CTkFrame(row_f, fg_color=bg_c,
                                              corner_radius=10, height=26)
                        badge.place(relx=0.29, rely=0.5, anchor="w",
                                    relwidth=0.08)
                        ctk.CTkLabel(badge, text=subject,
                                      font=("Arial", 12, "bold"),
                                      text_color=fg_c).pack(expand=True)

                    for idx, val in enumerate(entry["slots"]):
                        if val:
                            tile = ctk.CTkFrame(row_f, fg_color="#EFF6FF",
                                                 corner_radius=6, height=30)
                            tile.place(relx=0.40 + (idx * 0.15), rely=0.5,
                                       anchor="w", relwidth=0.14)
                            ctk.CTkLabel(tile, text=val,
                                          font=("Arial", 13, "bold"),
                                          text_color=COLORS["accent"]
                                          ).pack(expand=True)

            self.update_idletasks()
        except Exception as e:
            print(f"Lỗi render_plan: {e}")
    def setup_plan_ui(self):
        self.clear_right_frame()
        header = ctk.CTkFrame(self.plan_frame, fg_color="transparent")
        header.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(header, text="Kế hoạch giảng dạy trong ngày",
                     font=("Arial", 24, "bold"), text_color=COLORS["text"]).pack(side="left")

        ctk.CTkButton(header, text="Làm mới", width=110, height=36,
                      fg_color=COLORS["accent"], hover_color="#1D4ED8",
                      font=("Arial", 13, "bold"),
                      command=self.check_realtime_status).pack(side="right")

        self.status_indicator = ctk.CTkLabel(header, text="● Sẵn sàng",
                                             font=("Arial", 13),
                                             text_color=COLORS["success"])
        self.status_indicator.pack(side="right", padx=(0, 12))

        self.plan_scroll = ctk.CTkScrollableFrame(self.plan_frame,
                                                  fg_color=COLORS["card"],
                                                  border_width=1,
                                                  border_color=COLORS["border"])
        self.plan_scroll.pack(fill="both", expand=True)
    def link_plan(self):
        """Hàm chọn file Excel từ máy tính"""
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.plan_path = path
            self.refresh_plan_data()
    def refresh_plan_data(self):
        if not self.plan_path:
            return

        try:
            df = pd.read_excel(self.plan_path, skiprows=3)
            df.columns = [str(col).strip() for col in df.columns]

            # Fill merged cells
            df['Họ và tên'] = df['Họ và tên'].ffill()

            slots = ["1 - 2", "3 - 4", "5 - 6", "7 - 8"]
            df = df.dropna(subset=['môn học'] + slots, how='all')

            # 🔥 GROUP BY TEACHER
            grouped = []

            for name, group in df.groupby('Họ và tên', sort=False):
                teacher = {
                    "name": name,
                    "subjects": [],
                    "rows": group.to_dict('records')
                }

                for _, r in group.iterrows():
                    subject = str(r.get('môn học', ''))
                    if subject != "nan" and subject not in teacher["subjects"]:
                        teacher["subjects"].append(subject)

                grouped.append(teacher)

            self.plan_data = df.to_dict('records')
            self.render_plan()

        except Exception as e:
            messagebox.showerror("Lỗi", str(e))
    def open_document(self, file_name):
        path = os.path.join("Document", file_name)
        try:
            if os.name == "nt":
                os.startfile(path)
            else:
                subprocess.call(["open", path])
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))
    def convert_excel_date(self, val):
        """Hàm phụ trợ để xử lý ngày tháng từ số Excel sang chuỗi dd/mm/yyyy"""
        try:
            if isinstance(val, (int, float)) and val > 1000:
                return pd.to_datetime(val, unit='D', origin='1899-12-30').strftime('%d/%m/%Y')
            return str(val) if str(val).lower() != 'nan' else ""
        except:
            return str(val)
    def update_time(self):
        self.lbl_time.configure(text=datetime.now().strftime("%H:%M:%S\n%A, %d/%m/%Y"))
        self.after(1000, self.update_time)   
    def load_monthly_plan(self, file_path):
        try:
            df = pd.read_excel(file_path)

            # Clear old content if reload
            for widget in self.tab_plan.winfo_children():
                widget.destroy()

            # Frame container
            frame = ctk.CTkFrame(self.tab_plan)
            frame.pack(fill="both", expand=True)

            # Create table
            tree = ttk.Treeview(frame)
            tree.pack(side="left", fill="both", expand=True)

            # Scrollbars
            scrollbar_y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
            scrollbar_y.pack(side="right", fill="y")

            scrollbar_x = ttk.Scrollbar(self.tab_plan, orient="horizontal", command=tree.xview)
            scrollbar_x.pack(fill="x")

            tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

            # Columns
            tree["columns"] = list(df.columns)
            tree["show"] = "headings"

            for col in df.columns:
                tree.heading(col, text=col)
                tree.column(col, anchor="center", width=120)

            # Rows
            for _, row in df.iterrows():
                tree.insert("", "end", values=list(row))

        except Exception as e:
            print("ERROR loading Excel:", e)   
    def load_documents(self, folder_path):
        frame = ctk.CTkFrame(self.tab_docs)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        files = [f for f in os.listdir(folder_path) if f.endswith(".pdf")]

        for file in files:
            btn = ctk.CTkButton(
                frame,
                text=file,
                anchor="w",
                command=lambda f=file: self.open_pdf(os.path.join(folder_path, f))
            )
            btn.pack(fill="x", pady=5)
    def check_realtime_status(self):
        import glob
        files = glob.glob("KeHoach_Ngay_*.xlsx")
        if not files: 
            print("❌ Không tìm thấy file Excel nào!")
            return

        latest_file = max(files, key=os.path.getctime)
        print(f"📂 Đang đọc file: {latest_file}")
        
        try:
            # Đọc từ dòng 4
            df = pd.read_excel(latest_file, skiprows=3)
            df.columns = [str(c).strip() for c in df.columns]
            
            # XỬ LÝ QUAN TRỌNG: Loại bỏ các dòng hoàn toàn trống
            # Chỉ giữ lại dòng có tên HOẶC có môn học
            df = df.dropna(subset=['Họ và tên', 'môn học'], how='all')
            
            self.plan_data = df.to_dict('records')
            
            # KIỂM TRA: In ra số lượng dòng Python đọc được
            print(f"✅ Đã nạp được {len(self.plan_data)} dòng dữ liệu.")
            
            if self.plan_scroll.winfo_exists():
                self.render_plan()
        except Exception as e:
            print(f"❌ Lỗi đọc file: {e}")


if __name__ == "__main__":
    app = TeacherManagerPro()
    app.mainloop()
