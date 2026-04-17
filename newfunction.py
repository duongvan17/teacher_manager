# CLEANED & FIXED VERSION (DOCUMENT SYSTEM WORKING)

import customtkinter as ctk
import pandas as pd
from tkinter import messagebox, filedialog
import os
import subprocess
from datetime import datetime

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

COLORS = {
    "bg": "#F5F7F9",
    "sidebar": "#FFFFFF",
    "accent": "#2563EB",
    "text": "#1E293B",
    "hover": "#F1F5F9",
    "border": "#E2E8F0",
}

class TeacherManagerPro(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("TSQ Teacher Manager Pro")
        self.geometry("1200x800")
        self.configure(fg_color=COLORS["bg"])

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.setup_sidebar()

        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=0, column=1, sticky="nsew")

        self.mgmt_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.plan_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.document_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")

        self.setup_mgmt_ui()
        self.setup_plan_ui()
        self.setup_document_ui()

        self.show_mgmt_frame()

    # ================= SIDEBAR =================
    def setup_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=250, fg_color=COLORS["sidebar"])
        self.sidebar.grid(row=0, column=0, sticky="nsew")

        ctk.CTkLabel(self.sidebar, text="TSQ QLGV", font=("Arial", 20, "bold")).pack(pady=30)

        ctk.CTkButton(self.sidebar, text="Quản lý", command=self.show_mgmt_frame).pack(pady=10, padx=20, fill="x")
        ctk.CTkButton(self.sidebar, text="Kế hoạch", command=self.show_plan_frame).pack(pady=10, padx=20, fill="x")
        ctk.CTkButton(self.sidebar, text="Tài liệu", command=self.show_document_frame).pack(pady=10, padx=20, fill="x")

    # ================= NAVIGATION =================
    def hide_all_frames(self):
        self.mgmt_frame.pack_forget()
        self.plan_frame.pack_forget()
        self.document_frame.pack_forget()

    def show_mgmt_frame(self):
        self.hide_all_frames()
        self.mgmt_frame.pack(fill="both", expand=True)

    def show_plan_frame(self):
        self.hide_all_frames()
        self.plan_frame.pack(fill="both", expand=True)

    def show_document_frame(self):
        self.hide_all_frames()
        self.document_frame.pack(fill="both", expand=True)
        self.render_documents()

    # ================= MANAGEMENT =================
    def setup_mgmt_ui(self):
        ctk.CTkLabel(self.mgmt_frame, text="Quản lý giảng viên", font=("Arial", 22, "bold")).pack(pady=20)

    # ================= PLAN =================
    def setup_plan_ui(self):
        ctk.CTkLabel(self.plan_frame, text="Kế hoạch giảng dạy", font=("Arial", 22, "bold")).pack(pady=20)

    # ================= DOCUMENT =================
    def setup_document_ui(self):
        header = ctk.CTkFrame(self.document_frame, fg_color="transparent")
        header.pack(fill="x", pady=10)

        ctk.CTkLabel(header, text="Tài liệu môn học", font=("Arial", 24, "bold")).pack(side="left", padx=20)

        ctk.CTkButton(header, text="Làm mới", command=self.render_documents).pack(side="right", padx=20)

        self.document_scroll = ctk.CTkScrollableFrame(self.document_frame)
        self.document_scroll.pack(fill="both", expand=True, padx=10, pady=10)

    def render_documents(self):
        # Xóa các widget cũ
        for child in self.document_scroll.winfo_children():
            child.destroy()

        folder = "documents"

        if not os.path.exists(folder):
            os.makedirs(folder)

        files = [f for f in os.listdir(folder) if not f.startswith('.')]

        # 1. Trạng thái trống (Empty State) hiện đại hơn
        if not files:
            empty_frame = ctk.CTkFrame(self.document_scroll, fg_color="transparent")
            empty_frame.pack(expand=True, fill="both", pady=40)
            
            ctk.CTkLabel(
                empty_frame, 
                text="📁", 
                font=("Segoe UI", 40)
            ).pack(pady=(0, 10))
            
            ctk.CTkLabel(
                empty_frame, 
                text="Chưa có tài liệu nào", 
                font=("Segoe UI", 16, "bold"),
                text_color=("gray50", "gray60")
            ).pack()
            return

        # 2. Render danh sách file
        for file in files:
            # Card chứa file (Bo góc, có viền mỏng, hỗ trợ Light/Dark mode)
            card = ctk.CTkFrame(
                self.document_scroll, 
                height=60,
                corner_radius=10,
                fg_color=("gray95", "gray15"),  # Màu nền nhạt cho Light, xám đậm cho Dark
                border_width=1,
                border_color=("gray85", "gray25") # Màu viền
            )
            card.pack(fill="x", padx=15, pady=6)
            card.pack_propagate(False)

            # Căn chỉnh Layout bên trong card
            content_frame = ctk.CTkFrame(card, fg_color="transparent")
            content_frame.pack(fill="both", expand=True, padx=15)

            # Icon và Tên file
            ctk.CTkLabel(
                content_frame, 
                text=f"📄  {file}", 
                font=("Segoe UI", 14, "bold"),
                text_color=("gray10", "gray90")
            ).pack(side="left")

            # Nút bấm hiện đại (Bo tròn nhiều hơn, màu Accent)
            ctk.CTkButton(
                content_frame,
                text="Mở",
                width=80,
                height=32,
                corner_radius=16, # Bo tròn dạng viên thuốc (pill shape)
                font=("Segoe UI", 12, "bold"),
                fg_color=("#2FA572", "#106A43"), # Màu xanh lá hiện đại
                hover_color=("#248259", "#0B4B2F"),
                command=lambda f=file: self.open_document(f)
            ).pack(side="right")

    def open_document(self, file_name):
        path = os.path.join("documents", file_name)
        try:
            if os.name == "nt":
                os.startfile(path)
            else:
                subprocess.call(["open", path])
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))


if __name__ == "__main__":
    app = TeacherManagerPro()
    app.mainloop()
def show_mgmt_frame(self):
        """Hiển thị tab Quản lý chung và ẩn tab Kế hoạch"""
        self.plan_frame.pack_forget()
        self.mgmt_frame.pack(fill="both", expand=True)
        self.btn_mgmt.configure(fg_color=COLORS["hover"], text_color=COLORS["accent"])
        self.btn_plan.configure(fg_color="transparent", text_color="blue")