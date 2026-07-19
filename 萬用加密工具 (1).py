# -*- coding: utf-8 -*-
"""
萬用加密工具 - 完整圖形介面版 v5.1（畫面清晰修正／UI 精緻化／FAQ 精簡）
=================================================
支援檔案類型：PDF、Word、Excel、PowerPoint、圖片(jpg/png/...)

v5.1 更新重點（相對 v5.0）：
  - 修正 Windows 高解析度螢幕文字模糊問題：改用更完整、更不容易衝突的
    DPI 感知設定順序（PerMonitorV2 → PerMonitor → System → 舊版 API），
    並在設定前先偵測目前狀態，避免重複設定導致系統退回模糊的點陣圖縮放。
  - 常見問題移除最後一題（畫面模糊 FAQ，因為現在已經直接修好了）。
  - 整體視覺精緻化：
      · 首頁新增「本機處理・不上傳」信任徽章
      · 區塊標題、步驟卡片改用圓形圖示徽章，質感更好
      · 步驟指示列的連接線會隨進度變色，一眼看出目前走到哪
      · 統一間距與圓角比例，卡片層次更清楚
"""

import os
import sys
import math
import queue
import secrets
import string
import threading
import platform
import subprocess

import tkinter as tk
from tkinter import filedialog, messagebox

# ------------------------------------------------------------------
# 選用套件檢查
# ------------------------------------------------------------------
try:
    import customtkinter as ctk
    CTK_AVAILABLE = True
except ImportError:
    CTK_AVAILABLE = False

try:
    from pypdf import PdfReader, PdfWriter
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

try:
    from PIL import Image
    IMAGE_AVAILABLE = True
except ImportError:
    IMAGE_AVAILABLE = False

OFFICE_AVAILABLE = False
if platform.system() == "Windows":
    try:
        import win32com.client as win32
        OFFICE_AVAILABLE = True
    except ImportError:
        OFFICE_AVAILABLE = False


APP_TITLE = "萬用加密工具 v5.1"

PDF_EXTENSIONS = (".pdf",)
WORD_EXTENSIONS = (".docx", ".doc")
EXCEL_EXTENSIONS = (".xlsx", ".xls")
PPT_EXTENSIONS = (".pptx", ".ppt")
IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".gif", ".webp")

FILE_ICONS = {
    "pdf": "📕", "word": "📘", "excel": "📗", "ppt": "📙",
    "image": "🖼️", "folder": "📁", "other": "📄",
}

FONT_FAMILY = "Microsoft JhengHei UI"

THEME = {
    "bg": "#F3F5FA",
    "sidebar": "#131A2B",
    "sidebar_hover": "#212B45",
    "sidebar_text": "#B9C2D9",
    "card": "#FFFFFF",
    "card_border": "#E4E8F1",
    "card_alt": "#F7F9FD",
    "accent": "#4F5BFF",
    "accent_hover": "#3C46D6",
    "accent_soft": "#EEF0FF",
    "text": "#161B26",
    "subtext": "#6B7280",
    "success": "#17A673",
    "success_soft": "#E5F8F1",
    "warn": "#D97706",
    "warn_soft": "#FEF3E2",
    "error": "#E5484D",
    "error_soft": "#FDECEC",
    "console_bg": "#0F1420",
    "console_text": "#D7E0F2",
    "step_inactive": "#DADFEC",
}


# ====================================================================
#  工具函式（純邏輯，與介面無關）
# ====================================================================
def password_strength(pw: str):
    if not pw:
        return 0, "尚未輸入", "subtext"
    classes = 0
    if any(c.islower() for c in pw):
        classes += 1
    if any(c.isupper() for c in pw):
        classes += 1
    if any(c.isdigit() for c in pw):
        classes += 1
    if any(c in string.punctuation for c in pw):
        classes += 1
    charset_size = {1: 26, 2: 52, 3: 62, 4: 94}.get(classes, 10)
    entropy = len(pw) * (math.log2(charset_size) if charset_size > 1 else 1)
    if len(pw) < 8:
        return 15, "太短，容易被破解（建議至少 12 碼）", "error"
    if entropy < 40:
        return 35, "弱：建議混合大小寫、數字、符號", "error"
    if entropy < 60:
        return 60, "中等：可以再加長或加符號更安全", "warn"
    if entropy < 80:
        return 82, "強：適合保護一般重要文件", "success"
    return 100, "非常強：適合保護高度機密文件", "success"


def generate_password(length=16, upper=True, lower=True, digits=True, symbols=True):
    pool = ""
    if upper:
        pool += string.ascii_uppercase
    if lower:
        pool += string.ascii_lowercase
    if digits:
        pool += string.digits
    if symbols:
        pool += "!@#$%^&*()-_=+[]{}"
    if not pool:
        pool = string.ascii_letters + string.digits
    while True:
        pw = "".join(secrets.choice(pool) for _ in range(length))
        ok = True
        if upper and not any(c in string.ascii_uppercase for c in pw):
            ok = False
        if lower and not any(c in string.ascii_lowercase for c in pw):
            ok = False
        if digits and not any(c in string.digits for c in pw):
            ok = False
        if symbols and not any(c in "!@#$%^&*()-_=+[]{}" for c in pw):
            ok = False
        if ok:
            return pw


def classify_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in PDF_EXTENSIONS:
        return "pdf"
    if ext in WORD_EXTENSIONS:
        return "word"
    if ext in EXCEL_EXTENSIONS:
        return "excel"
    if ext in PPT_EXTENSIONS:
        return "ppt"
    if ext in IMAGE_EXTENSIONS:
        return "image"
    return "other"


def open_in_file_explorer(path):
    try:
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])
    except Exception:
        pass


# ====================================================================
#  加密引擎（核心邏輯，與介面完全分離，方便測試）
# ====================================================================
class EncryptionEngine:

    @staticmethod
    def pdf_encrypt(input_path, output_path, password):
        reader = PdfReader(input_path)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.encrypt(user_password=password, algorithm="AES-256")
        with open(output_path, "wb") as f:
            writer.write(f)

    @staticmethod
    def images_to_encrypted_pdf(image_paths, output_pdf_path, password):
        images = []
        for p in image_paths:
            im = Image.open(p)
            if im.mode != "RGB":
                im = im.convert("RGB")
            images.append(im)
        if not images:
            raise ValueError("沒有可用的圖片")
        tmp_pdf = output_pdf_path + ".tmp.pdf"
        first, rest = images[0], images[1:]
        first.save(tmp_pdf, save_all=True, append_images=rest)
        EncryptionEngine.pdf_encrypt(tmp_pdf, output_pdf_path, password)
        os.remove(tmp_pdf)

    @staticmethod
    def office_encrypt(file_pairs, password, kind):
        results = []
        if kind == "word":
            app = win32.Dispatch("Word.Application")
        elif kind == "excel":
            app = win32.Dispatch("Excel.Application")
        else:
            app = win32.Dispatch("PowerPoint.Application")
        app.Visible = False
        try:
            app.DisplayAlerts = False
        except Exception:
            pass
        try:
            for input_path, output_path in file_pairs:
                filename = os.path.basename(input_path)
                try:
                    if kind == "word":
                        doc = app.Documents.Open(input_path)
                        doc.Password = password
                        doc.SaveAs2(output_path)
                        doc.Close()
                    elif kind == "excel":
                        wb = app.Workbooks.Open(input_path)
                        wb.Password = password
                        wb.SaveAs(output_path)
                        wb.Close()
                    else:
                        pres = app.Presentations.Open(input_path, WithWindow=False)
                        pres.Password = password
                        pres.Save()
                        if output_path != input_path:
                            pres.SaveAs(output_path)
                        pres.Close()
                    results.append((filename, True, None))
                except Exception as e:
                    results.append((filename, False, str(e)))
        finally:
            app.Quit()
        return results


# ====================================================================
#  小型可重用元件
# ====================================================================
def make_card(parent, **kwargs):
    defaults = dict(fg_color=THEME["card"], corner_radius=16,
                     border_width=1, border_color=THEME["card_border"])
    defaults.update(kwargs)
    return ctk.CTkFrame(parent, **defaults)


def icon_badge(parent, icon, size=40, bg=None, fg=None):
    """圓形圖示徽章，用於區塊標題、步驟卡片等，取代原本的細直條，質感更精緻。"""
    return ctk.CTkLabel(
        parent, text=icon, width=size, height=size, corner_radius=size // 2,
        fg_color=bg or THEME["accent_soft"], text_color=fg or THEME["accent"],
        font=(FONT_FAMILY, int(size * 0.45)),
    )


def section_title(parent, icon, text, subtitle=None):
    wrap = ctk.CTkFrame(parent, fg_color="transparent")
    wrap.pack(fill="x", padx=20, pady=(20, 6))
    row = ctk.CTkFrame(wrap, fg_color="transparent")
    row.pack(fill="x")
    icon_badge(row, icon, size=36).pack(side="left", padx=(0, 12))
    title_col = ctk.CTkFrame(row, fg_color="transparent")
    title_col.pack(side="left", fill="x", expand=True)
    ctk.CTkLabel(title_col, text=text, font=(FONT_FAMILY, 17, "bold"),
                 text_color=THEME["text"]).pack(anchor="w")
    if subtitle:
        ctk.CTkLabel(title_col, text=subtitle, font=(FONT_FAMILY, 13),
                     text_color=THEME["subtext"], justify="left",
                     wraplength=760).pack(anchor="w", pady=(4, 0))
    return wrap


def chip(parent, text, bg=None, fg=None, font_size=12):
    return ctk.CTkLabel(
        parent, text=text, font=(FONT_FAMILY, font_size, "bold"),
        fg_color=bg or THEME["accent_soft"], text_color=fg or THEME["accent"],
        corner_radius=16, padx=14, pady=8,
    )


def pill_button(parent, text, command=None, kind="primary", width=150, height=44, state="normal"):
    """
    kind:
      primary   → 主要動作（例如：下一步、開始加密）
      secondary → 次要但仍重要的動作（例如：上一步）→ 有邊框，比 ghost 更明顯
      ghost     → 輔助動作（例如：清除選擇）
      danger    → 危險動作
    """
    palette = {
        "primary":   (THEME["accent"],     THEME["accent_hover"], "#FFFFFF",      None),
        "secondary": (THEME["card"],       THEME["accent_soft"],  THEME["accent"], THEME["accent"]),
        "ghost":     (THEME["card_alt"],   THEME["card_border"],  THEME["text"],   None),
        "danger":    (THEME["error_soft"], THEME["error"],        THEME["error"],  None),
    }
    fg, hover, txt, border = palette.get(kind, palette["primary"])
    kwargs = dict(
        text=text, command=command, width=width, height=height,
        corner_radius=11, fg_color=fg, hover_color=hover, text_color=txt,
        font=(FONT_FAMILY, 13, "bold"), state=state,
    )
    if border:
        kwargs["border_width"] = 2
        kwargs["border_color"] = border
    return ctk.CTkButton(parent, **kwargs)


# ====================================================================
#  主要 GUI 應用程式
# ====================================================================
class App:
    NAV_ITEMS = [
        ("home", "🏠", "開始使用"),
        ("encrypt", "🔒", "加密檔案"),
        ("imgpdf", "🖼️", "圖片轉加密PDF"),
        ("pwtools", "🔑", "密碼工具"),
        ("help", "📖", "使用說明"),
    ]

    WIZARD_STEPS = [
        (1, "📁", "選擇檔案"),
        (2, "🌍", "加密方式"),
        (3, "🔑", "設定密碼"),
        (4, "📂", "輸出位置"),
        (5, "🚀", "開始加密"),
    ]

    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1240x830")
        self.root.minsize(1040, 660)

        self.selected_paths = []
        self.selected_images = []
        self.progress_queue = queue.Queue()
        self.pages = {}
        self.nav_buttons = {}

        self.output_mode = tk.StringVar(value="same")
        self.custom_output_dir = None

        self.wizard_frames = {}
        self.current_step = 1
        self.max_step_reached = 1
        self.step_dots = {}
        self.step_dot_wraps = {}
        self.step_connectors = {}

        self._build_shell()
        self._poll_queue()
        self.show_page("home")

    # ---------------------------------------------------------------
    def _build_shell(self):
        outer = ctk.CTkFrame(self.root, fg_color=THEME["bg"], corner_radius=0)
        outer.pack(fill="both", expand=True)
        outer.grid_columnconfigure(1, weight=1)
        outer.grid_rowconfigure(0, weight=1)

        # ---- 左側導覽列 ----
        sidebar = ctk.CTkFrame(outer, width=236, fg_color=THEME["sidebar"], corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="ns")
        sidebar.grid_propagate(False)

        logo_box = ctk.CTkFrame(sidebar, fg_color="transparent")
        logo_box.pack(fill="x", padx=22, pady=(28, 20))
        logo_circle = ctk.CTkLabel(logo_box, text="🔐", width=52, height=52, corner_radius=16,
                                    fg_color=THEME["accent"], font=(FONT_FAMILY, 24))
        logo_circle.pack(anchor="w")
        ctk.CTkLabel(logo_box, text="萬用加密工具", font=(FONT_FAMILY, 19, "bold"),
                     text_color="#FFFFFF").pack(anchor="w", pady=(10, 0))
        ctk.CTkLabel(logo_box, text="保護你的每一份重要檔案", font=(FONT_FAMILY, 12),
                     text_color=THEME["sidebar_text"]).pack(anchor="w")

        ctk.CTkFrame(sidebar, height=1, fg_color="#232D46").pack(fill="x", padx=20, pady=(0, 14))

        nav_wrap = ctk.CTkFrame(sidebar, fg_color="transparent")
        nav_wrap.pack(fill="x", padx=12)
        for key, icon, label in self.NAV_ITEMS:
            btn = ctk.CTkButton(
                nav_wrap, text=f"  {icon}   {label}", anchor="w",
                font=(FONT_FAMILY, 13, "bold"), height=46, corner_radius=11,
                fg_color="transparent", hover_color=THEME["sidebar_hover"],
                text_color=THEME["sidebar_text"],
                command=lambda k=key: self.show_page(k),
            )
            btn.pack(fill="x", pady=3)
            self.nav_buttons[key] = btn

        status_wrap = ctk.CTkFrame(sidebar, fg_color="#0C1120", corner_radius=14)
        status_wrap.pack(fill="x", side="bottom", padx=14, pady=16)
        ctk.CTkLabel(status_wrap, text="🟢 功能偵測狀態", font=(FONT_FAMILY, 11, "bold"),
                     text_color=THEME["sidebar_text"]).pack(anchor="w", padx=14, pady=(12, 6))
        deps = [
            ("PDF加密", PDF_AVAILABLE),
            ("圖片處理", IMAGE_AVAILABLE),
            ("Office原生", OFFICE_AVAILABLE),
        ]
        for name, ok in deps:
            row = ctk.CTkFrame(status_wrap, fg_color="transparent")
            row.pack(fill="x", padx=14, pady=2)
            dot = "🟢" if ok else "⚪"
            ctk.CTkLabel(row, text=f"{dot} {name}", font=(FONT_FAMILY, 11),
                         text_color=THEME["sidebar_text"] if ok else "#5B6376").pack(anchor="w")
        ctk.CTkLabel(status_wrap, text="", font=(FONT_FAMILY, 2)).pack(pady=3)

        # ---- 右側內容區 ----
        content_shell = ctk.CTkFrame(outer, fg_color=THEME["bg"], corner_radius=0)
        content_shell.grid(row=0, column=1, sticky="nsew")
        content_shell.grid_rowconfigure(1, weight=1)
        content_shell.grid_columnconfigure(0, weight=1)

        topbar = ctk.CTkFrame(content_shell, fg_color=THEME["bg"], height=56, corner_radius=0)
        topbar.grid(row=0, column=0, sticky="ew", padx=26, pady=(20, 0))
        self.page_title_label = ctk.CTkLabel(topbar, text="", font=(FONT_FAMILY, 23, "bold"),
                                              text_color=THEME["text"])
        self.page_title_label.pack(side="left")

        self.content_area = ctk.CTkFrame(content_shell, fg_color=THEME["bg"], corner_radius=0)
        self.content_area.grid(row=1, column=0, sticky="nsew", padx=26, pady=18)
        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)

        self._build_home_page()
        self._build_encrypt_page()
        self._build_imgpdf_page()
        self._build_pwtools_page()
        self._build_help_page()

    def show_page(self, key):
        for k, frame in self.pages.items():
            frame.grid_remove()
        self.pages[key].grid(row=0, column=0, sticky="nsew")
        titles = {k: label for k, _icon, label in self.NAV_ITEMS}
        self.page_title_label.configure(text=titles.get(key, ""))
        for k, btn in self.nav_buttons.items():
            if k == key:
                btn.configure(fg_color=THEME["accent"], text_color="#FFFFFF")
            else:
                btn.configure(fg_color="transparent", text_color=THEME["sidebar_text"])

    def _new_plain_page(self, key):
        page = ctk.CTkFrame(self.content_area, fg_color=THEME["bg"], corner_radius=0)
        page.grid(row=0, column=0, sticky="nsew")
        page.grid_columnconfigure(0, weight=1)
        self.pages[key] = page
        return page

    def _new_scroll_page(self, key):
        page = ctk.CTkScrollableFrame(self.content_area, fg_color=THEME["bg"],
                                       corner_radius=0, scrollbar_button_color=THEME["card_border"])
        page.grid(row=0, column=0, sticky="nsew")
        page.grid_columnconfigure(0, weight=1)
        self.pages[key] = page
        return page

    # =================================================================
    # 首頁
    # =================================================================
    def _build_home_page(self):
        page = self._new_scroll_page("home")

        hero = make_card(page, fg_color=THEME["accent_soft"], border_width=0, corner_radius=20)
        hero.pack(fill="x", pady=(4, 18))
        badge_row = ctk.CTkFrame(hero, fg_color="transparent")
        badge_row.pack(anchor="w", padx=24, pady=(22, 10))
        chip(badge_row, "", bg="#FFFFFF", fg=THEME["accent"], font_size=11).pack(side="left")
        ctk.CTkLabel(hero, text="歡迎使用！照著「🔒 加密檔案」的 5 個步驟就能完成",
                     font=(FONT_FAMILY, 20, "bold"), text_color=THEME["text"]).pack(anchor="w", padx=24, pady=(0, 4))
        ctk.CTkLabel(hero, text="所有處理都在你自己的電腦上完成，不會上傳任何檔案或密碼。",
                     font=(FONT_FAMILY, 13), text_color=THEME["subtext"]).pack(anchor="w", padx=24, pady=(0, 20))

        steps = [
            ("1", "📁", "選擇檔案", "挑出你要保護的檔案，或整個資料夾。"),
            ("2", "🌍", "選加密方式", "PDF 選 PDF 原生加密；Office 檔選 Office 原生加密。"),
            ("3", "🔑", "設定密碼", "自己輸入，或按一下自動產生高強度密碼。"),
            ("4", "📂", "選輸出位置", "預設存在原檔案旁邊，也可以另外指定資料夾。"),
            ("5", "🚀", "開始加密", "按下按鈕，等待完成即可。"),
        ]
        grid = ctk.CTkFrame(page, fg_color="transparent")
        grid.pack(fill="x", pady=(0, 18))
        grid.grid_columnconfigure(tuple(range(5)), weight=1, uniform="steps")
        for i, (num, icon, title, desc) in enumerate(steps):
            card = make_card(grid)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 8, 0))
            icon_badge(card, icon, size=42).pack(anchor="w", padx=16, pady=(16, 8))
            ctk.CTkLabel(card, text=f"步驟 {num}", font=(FONT_FAMILY, 10, "bold"),
                         text_color=THEME["accent"]).pack(anchor="w", padx=16)
            ctk.CTkLabel(card, text=title, font=(FONT_FAMILY, 14, "bold"),
                         text_color=THEME["text"]).pack(anchor="w", padx=16, pady=(2, 4))
            ctk.CTkLabel(card, text=desc, font=(FONT_FAMILY, 11), text_color=THEME["subtext"],
                         wraplength=190, justify="left").pack(anchor="w", padx=16, pady=(0, 16))

        note = make_card(page, fg_color=THEME["warn_soft"], border_width=0, corner_radius=16)
        note.pack(fill="x", pady=(0, 18))
        note_head = ctk.CTkFrame(note, fg_color="transparent")
        note_head.pack(anchor="w", padx=22, pady=(18, 4))
        ctk.CTkLabel(note_head, text="⚠️", font=(FONT_FAMILY, 16)).pack(side="left", padx=(0, 8))
        ctk.CTkLabel(note_head, text="重要提醒", font=(FONT_FAMILY, 15, "bold"),
                     text_color=THEME["warn"]).pack(side="left")
        ctk.CTkLabel(note, justify="left", wraplength=900, text=(
            "• 密碼一旦忘記，沒有任何人可以幫你救回檔案內容，請務必牢記或安全保存密碼。\n"
            "• 加密完成後，收到檔案的人只要輸入正確密碼，用一般 PDF 閱讀器 / Office\n"
            "  直接打開即可，不需要再安裝任何額外工具。"
        ), font=(FONT_FAMILY, 13), text_color=THEME["text"]).pack(anchor="w", padx=22, pady=(0, 18))

        pill_button(page, "開始加密 →", command=lambda: self.show_page("encrypt"),
                    width=180, height=48).pack(anchor="w", pady=(0, 24))

    # =================================================================
    # 加密檔案頁：5 步驟精靈流程
    # =================================================================
    def _build_encrypt_page(self):
        page = self._new_plain_page("encrypt")
        page.grid_rowconfigure(1, weight=1)

        # ---- 步驟指示條 ----
        step_bar = ctk.CTkFrame(page, fg_color="transparent")
        step_bar.grid(row=0, column=0, sticky="ew", pady=(4, 6))
        self._build_step_indicator(step_bar)
        ctk.CTkLabel(page, text="💡 已完成的步驟圖示可以直接點擊快速跳轉",
                     font=(FONT_FAMILY, 11), text_color=THEME["subtext"]).grid(
            row=0, column=0, sticky="ew", pady=(74, 0))

        # ---- 每一步的內容容器（用可捲動框架當安全網，避免內容被裁切）----
        self.wizard_container = ctk.CTkScrollableFrame(
            page, fg_color=THEME["bg"], corner_radius=0,
            scrollbar_button_color=THEME["card_border"])
        self.wizard_container.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        self.wizard_container.grid_columnconfigure(0, weight=1)

        self._build_wizard_step1()
        self._build_wizard_step2()
        self._build_wizard_step3()
        self._build_wizard_step4()
        self._build_wizard_step5()

        self._show_wizard_step(1)

    def _build_step_indicator(self, parent):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x")
        for i, (num, icon, label) in enumerate(self.WIZARD_STEPS):
            dot_wrap = ctk.CTkFrame(row, fg_color="transparent", cursor="hand2")
            dot_wrap.pack(side="left")
            circle = ctk.CTkLabel(dot_wrap, text=f"{icon}", width=44, height=44, corner_radius=22,
                                   fg_color=THEME["step_inactive"], text_color=THEME["text"],
                                   font=(FONT_FAMILY, 17), cursor="hand2")
            circle.pack()
            label_widget = ctk.CTkLabel(dot_wrap, text=f"{num}. {label}", font=(FONT_FAMILY, 11, "bold"),
                                         text_color=THEME["subtext"], cursor="hand2")
            label_widget.pack(pady=(4, 0))
            self.step_dots[num] = circle
            self.step_dot_wraps[num] = dot_wrap
            for widget in (dot_wrap, circle, label_widget):
                widget.bind("<Button-1>", lambda e, n=num: self._try_jump_step(n))
            if i < len(self.WIZARD_STEPS) - 1:
                connector = ctk.CTkFrame(row, height=3, width=56, fg_color=THEME["step_inactive"],
                                          corner_radius=2)
                connector.pack(side="left", padx=6, pady=(0, 24))
                self.step_connectors[num] = connector

    def _try_jump_step(self, n):
        """點擊上方步驟圖示：只能跳到「已經走過」的步驟，避免跳過必要的驗證。"""
        if n <= self.max_step_reached:
            self._show_wizard_step(n)
        else:
            messagebox.showinfo("尚未完成前面的步驟", "請先完成目前的步驟，才能前往後面的步驟喔。")

    def _update_step_indicator(self):
        for num, circle in self.step_dots.items():
            if num < self.current_step:
                circle.configure(fg_color=THEME["success"], text_color="#FFFFFF")
            elif num == self.current_step:
                circle.configure(fg_color=THEME["accent"], text_color="#FFFFFF")
            else:
                circle.configure(fg_color=THEME["step_inactive"], text_color=THEME["text"])
        for num, connector in self.step_connectors.items():
            connector.configure(fg_color=THEME["success"] if num < self.current_step else THEME["step_inactive"])

    def _show_wizard_step(self, n):
        for step_num, frame in self.wizard_frames.items():
            frame.pack_forget()
        self.wizard_frames[n].pack(fill="both", expand=True)
        self.current_step = n
        self.max_step_reached = max(self.max_step_reached, n)
        self._update_step_indicator()
        try:
            self.wizard_container._parent_canvas.yview_moveto(0)
        except Exception:
            pass

    def _wizard_nav_row(self, parent, back_step=None, next_step=None, next_text="下一步 ›",
                         next_command=None):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=(26, 6))
        if back_step:
            pill_button(row, "‹ 上一步", lambda: self._show_wizard_step(back_step),
                        "secondary", 130, 44).pack(side="left")
        if next_step or next_command:
            cmd = next_command if next_command else (lambda: self._show_wizard_step(next_step))
            pill_button(row, next_text, cmd, "primary", 160, 44).pack(side="right")
        return row

    # ---- 步驟 1：選擇檔案 ----
    def _build_wizard_step1(self):
        frame = ctk.CTkFrame(self.wizard_container, fg_color="transparent")
        self.wizard_frames[1] = frame

        card = make_card(frame)
        card.pack(fill="x")
        section_title(card, "📁", "選擇要加密的檔案或資料夾")
        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(4, 10))
        pill_button(btn_row, "📄 選擇檔案", self._pick_files, "ghost", 150, 42).pack(side="left", padx=(0, 8))
        pill_button(btn_row, "📁 選擇資料夾", self._pick_folder, "ghost", 150, 42).pack(side="left", padx=(0, 8))
        pill_button(btn_row, "🗑️ 清除選擇", self._clear_encrypt_selection, "ghost", 130, 42).pack(side="left")

        self.encrypt_file_list_frame = ctk.CTkFrame(card, fg_color=THEME["card_alt"], corner_radius=12)
        self.encrypt_file_list_frame.pack(fill="x", padx=20, pady=(0, 8))

        ctk.CTkLabel(card, text="💡 提示：可以多次點選「選擇檔案」來累加不同檔案", font=(FONT_FAMILY, 12),
                     text_color=THEME["subtext"]).pack(anchor="w", padx=20, pady=(0, 18))

        self._refresh_encrypt_file_list()
        self._wizard_nav_row(frame, back_step=None, next_command=self._go_step2)

    def _go_step2(self):
        if not self.selected_paths:
            messagebox.showwarning("尚未選擇檔案", "請先選擇至少一個要加密的檔案或資料夾。")
            return
        self._show_wizard_step(2)

    def _refresh_encrypt_file_list(self):
        for w in self.encrypt_file_list_frame.winfo_children():
            w.destroy()
        if not self.selected_paths:
            ctk.CTkLabel(self.encrypt_file_list_frame, text="尚未選擇任何檔案或資料夾",
                         font=(FONT_FAMILY, 12), text_color=THEME["subtext"]).pack(pady=20)
            return
        for path in self.selected_paths:
            is_dir = os.path.isdir(path)
            icon = FILE_ICONS["folder"] if is_dir else FILE_ICONS.get(classify_file(path), FILE_ICONS["other"])
            row = ctk.CTkFrame(self.encrypt_file_list_frame, fg_color=THEME["card"], corner_radius=10,
                                border_width=1, border_color=THEME["card_border"])
            row.pack(fill="x", padx=10, pady=5)
            ctk.CTkLabel(row, text=icon, font=(FONT_FAMILY, 17)).pack(side="left", padx=(14, 8), pady=9)
            name = os.path.basename(path.rstrip("/\\")) + ("　(資料夾)" if is_dir else "")
            ctk.CTkLabel(row, text=name, font=(FONT_FAMILY, 12), text_color=THEME["text"],
                         anchor="w").pack(side="left", fill="x", expand=True)
            ctk.CTkButton(row, text="✕", width=30, height=30, corner_radius=15,
                          fg_color="transparent", hover_color=THEME["error_soft"],
                          text_color=THEME["subtext"],
                          command=lambda p=path: self._remove_encrypt_path(p)).pack(side="right", padx=8)

    def _remove_encrypt_path(self, path):
        self.selected_paths = [p for p in self.selected_paths if p != path]
        self._refresh_encrypt_file_list()

    def _pick_files(self):
        paths = filedialog.askopenfilenames(title="選擇要加密的檔案")
        for p in paths:
            if p not in self.selected_paths:
                self.selected_paths.append(p)
        self._refresh_encrypt_file_list()

    def _pick_folder(self):
        path = filedialog.askdirectory(title="選擇要加密的資料夾")
        if path and path not in self.selected_paths:
            self.selected_paths.append(path)
        self._refresh_encrypt_file_list()

    def _clear_encrypt_selection(self):
        self.selected_paths = []
        self._refresh_encrypt_file_list()

    # ---- 步驟 2：加密方式 ----
    def _build_wizard_step2(self):
        frame = ctk.CTkFrame(self.wizard_container, fg_color="transparent")
        self.wizard_frames[2] = frame

        card = make_card(frame)
        card.pack(fill="x")
        section_title(card, "🌍", "選擇加密方式",
                      "兩種方式都不需要額外安裝解壓縮工具，收件人直接用 PDF 閱讀器或 Office 打開就會跳出密碼視窗。")

        default_method = "pdf_native" if PDF_AVAILABLE else "office_native"
        self.encrypt_method = tk.StringVar(value=default_method)
        method_wrap = ctk.CTkFrame(card, fg_color="transparent")
        method_wrap.pack(fill="x", padx=20, pady=(6, 4))
        methods = [
            ("pdf_native", "📕  PDF 原生加密（雙擊即跳出密碼視窗，僅限 PDF）", PDF_AVAILABLE),
            ("office_native", "📘  Word/Excel/PowerPoint 原生加密（需正版 Office，僅 Windows）", OFFICE_AVAILABLE),
        ]
        for value, text, enabled in methods:
            rb = ctk.CTkRadioButton(method_wrap, text=text, value=value, variable=self.encrypt_method,
                                     font=(FONT_FAMILY, 13), fg_color=THEME["accent"])
            rb.pack(anchor="w", pady=9)
            if not enabled:
                rb.configure(state="disabled")

        if not PDF_AVAILABLE and not OFFICE_AVAILABLE:
            warn_box = make_card(card, fg_color=THEME["error_soft"], border_width=0, corner_radius=12)
            warn_box.pack(fill="x", padx=20, pady=(4, 18))
            ctk.CTkLabel(warn_box, justify="left", wraplength=780, font=(FONT_FAMILY, 12),
                         text_color=THEME["error"], text=(
                "⚠️ 目前偵測不到可用的加密方式。請安裝 pypdf（PDF 加密）或在已安裝正版 "
                "Office 的 Windows 電腦上安裝 pywin32（Office 原生加密）。"
            )).pack(anchor="w", padx=14, pady=10)
        else:
            ctk.CTkLabel(card, text="", font=(FONT_FAMILY, 4)).pack(pady=2)

        self._wizard_nav_row(frame, back_step=1, next_step=3)

    # ---- 步驟 3：設定密碼 ----
    def _build_wizard_step3(self):
        frame = ctk.CTkFrame(self.wizard_container, fg_color="transparent")
        self.wizard_frames[3] = frame

        card = make_card(frame)
        card.pack(fill="x")
        section_title(card, "🔑", "設定密碼", "這是打開加密檔案時需要輸入的密碼，請務必牢記。")
        pw_wrap = ctk.CTkFrame(card, fg_color="transparent")
        pw_wrap.pack(fill="x", padx=20, pady=(6, 4))

        self.encrypt_pw1 = tk.StringVar()
        self.encrypt_pw2 = tk.StringVar()

        row1 = ctk.CTkFrame(pw_wrap, fg_color="transparent")
        row1.pack(fill="x", pady=5)
        ctk.CTkLabel(row1, text="密碼", width=90, anchor="w", font=(FONT_FAMILY, 13)).pack(side="left")
        self.pw_entry1 = ctk.CTkEntry(row1, textvariable=self.encrypt_pw1, show="●", width=320, height=42,
                                       corner_radius=10, font=(FONT_FAMILY, 13), placeholder_text="請輸入密碼")
        self.pw_entry1.pack(side="left")
        self.pw_entry1.bind("<KeyRelease>", self._update_strength_meter)

        row2 = ctk.CTkFrame(pw_wrap, fg_color="transparent")
        row2.pack(fill="x", pady=5)
        ctk.CTkLabel(row2, text="確認密碼", width=90, anchor="w", font=(FONT_FAMILY, 13)).pack(side="left")
        self.pw_entry2 = ctk.CTkEntry(row2, textvariable=self.encrypt_pw2, show="●", width=320, height=42,
                                       corner_radius=10, font=(FONT_FAMILY, 13), placeholder_text="請再輸入一次")
        self.pw_entry2.pack(side="left")

        ctrl_row = ctk.CTkFrame(pw_wrap, fg_color="transparent")
        ctrl_row.pack(fill="x", pady=(10, 4))
        self.show_pw = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(ctrl_row, text="顯示密碼", variable=self.show_pw, command=self._toggle_show_pw,
                         font=(FONT_FAMILY, 12), fg_color=THEME["accent"]).pack(side="left")
        pill_button(ctrl_row, "🎲 自動產生高強度密碼", self._autofill_password, "ghost", 210, 38).pack(side="left", padx=10)

        self.strength_label = ctk.CTkLabel(pw_wrap, text="密碼強度：尚未輸入", font=(FONT_FAMILY, 12),
                                            text_color=THEME["subtext"])
        self.strength_label.pack(anchor="w", pady=(16, 3))
        self.strength_bar = ctk.CTkProgressBar(pw_wrap, width=330, height=11, corner_radius=6)
        self.strength_bar.set(0)
        self.strength_bar.pack(anchor="w", pady=(0, 18))

        self._wizard_nav_row(frame, back_step=2, next_command=self._go_step4)

    def _go_step4(self):
        pw1, pw2 = self.encrypt_pw1.get(), self.encrypt_pw2.get()
        if not pw1:
            messagebox.showwarning("密碼空白", "請輸入密碼。")
            return
        if pw1 != pw2:
            messagebox.showwarning("密碼不一致", "兩次輸入的密碼不一致，請重新輸入。")
            return
        score, _, _ = password_strength(pw1)
        if score < 35:
            if not messagebox.askyesno("密碼強度偏弱", "目前的密碼強度偏弱，容易被破解。\n仍要繼續嗎？"):
                return
        self._show_wizard_step(4)

    def _toggle_show_pw(self):
        show = "" if self.show_pw.get() else "●"
        self.pw_entry1.configure(show=show)
        self.pw_entry2.configure(show=show)

    def _autofill_password(self):
        pw = generate_password(16)
        self.encrypt_pw1.set(pw)
        self.encrypt_pw2.set(pw)
        self.show_pw.set(True)
        self._toggle_show_pw()
        self._update_strength_meter()
        messagebox.showinfo("已產生密碼",
                             f"已自動填入密碼：\n\n{pw}\n\n請務必先複製並安全保存這組密碼，"
                             f"再繼續進行加密（忘記密碼將無法救回檔案）。")

    def _strength_color(self, key):
        return THEME.get(key, THEME["subtext"])

    def _update_strength_meter(self, *_):
        score, text, color_key = password_strength(self.encrypt_pw1.get())
        self.strength_label.configure(text=f"密碼強度：{text}", text_color=self._strength_color(color_key))
        self.strength_bar.set(score / 100)
        self.strength_bar.configure(progress_color=self._strength_color(color_key))

    # ---- 步驟 4：輸出位置 ----
    def _build_wizard_step4(self):
        frame = ctk.CTkFrame(self.wizard_container, fg_color="transparent")
        self.wizard_frames[4] = frame

        card = make_card(frame)
        card.pack(fill="x")
        section_title(card, "📂", "選擇加密後檔案要存在哪裡")

        opt_wrap = ctk.CTkFrame(card, fg_color="transparent")
        opt_wrap.pack(fill="x", padx=20, pady=(6, 4))
        ctk.CTkRadioButton(opt_wrap, text="存在原始檔案所在的資料夾（預設，推薦）", value="same",
                            variable=self.output_mode, font=(FONT_FAMILY, 13),
                            fg_color=THEME["accent"], command=self._on_output_mode_change).pack(anchor="w", pady=9)
        ctk.CTkRadioButton(opt_wrap, text="另外指定一個資料夾", value="custom",
                            variable=self.output_mode, font=(FONT_FAMILY, 13),
                            fg_color=THEME["accent"], command=self._on_output_mode_change).pack(anchor="w", pady=9)

        path_row = ctk.CTkFrame(card, fg_color="transparent")
        path_row.pack(fill="x", padx=20, pady=(6, 18))
        self.output_path_label = ctk.CTkLabel(path_row, text="尚未指定資料夾", font=(FONT_FAMILY, 12),
                                               text_color=THEME["subtext"])
        self.output_path_label.pack(side="left")
        self.output_browse_btn = pill_button(path_row, "📂 選擇資料夾", self._pick_output_folder,
                                              "ghost", 150, 40, state="disabled")
        self.output_browse_btn.pack(side="left", padx=10)

        self._wizard_nav_row(frame, back_step=3, next_command=self._go_step5)

    def _on_output_mode_change(self):
        if self.output_mode.get() == "custom":
            self.output_browse_btn.configure(state="normal")
            if not self.custom_output_dir:
                self.output_path_label.configure(text="請按右邊按鈕選擇資料夾")
        else:
            self.output_browse_btn.configure(state="disabled")
            self.output_path_label.configure(text="加密後檔案將存在原始檔案旁邊", text_color=THEME["subtext"])

    def _pick_output_folder(self):
        folder = filedialog.askdirectory(title="選擇加密檔案要存放的資料夾")
        if folder:
            self.custom_output_dir = folder
            self.output_path_label.configure(text=folder, text_color=THEME["text"])

    def _go_step5(self):
        if self.output_mode.get() == "custom" and not self.custom_output_dir:
            messagebox.showwarning("尚未選擇資料夾", "請先選擇輸出資料夾，或改選「存在原始檔案所在的資料夾」。")
            return
        self._update_summary()
        self._show_wizard_step(5)

    # ---- 步驟 5：確認並開始加密 ----
    def _build_wizard_step5(self):
        frame = ctk.CTkFrame(self.wizard_container, fg_color="transparent")
        self.wizard_frames[5] = frame

        card = make_card(frame)
        card.pack(fill="x")
        section_title(card, "🚀", "確認設定並開始加密")

        self.summary_label = ctk.CTkLabel(card, text="", font=(FONT_FAMILY, 13), text_color=THEME["text"],
                                           justify="left", wraplength=830)
        self.summary_label.pack(anchor="w", padx=20, pady=(6, 18))

        action_row = ctk.CTkFrame(card, fg_color="transparent")
        action_row.pack(fill="x", padx=20, pady=(0, 10))
        self.encrypt_start_btn = pill_button(action_row, "🔒  開始加密", self._start_encrypt, "primary", 180, 48)
        self.encrypt_start_btn.pack(side="left")

        self.encrypt_progress = ctk.CTkProgressBar(card, height=11, corner_radius=6)
        self.encrypt_progress.set(0)
        self.encrypt_progress.pack(fill="x", padx=20, pady=(4, 10))

        log_card = ctk.CTkFrame(card, fg_color=THEME["console_bg"], corner_radius=14)
        log_card.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.encrypt_log = ctk.CTkTextbox(log_card, height=200, font=("Consolas", 12),
                                           fg_color=THEME["console_bg"], text_color=THEME["console_text"],
                                           corner_radius=14)
        self.encrypt_log.pack(fill="both", expand=True, padx=4, pady=4)
        self.encrypt_log.configure(state="disabled")

        self._wizard_nav_row(frame, back_step=4)

    def _update_summary(self):
        method_names = {"pdf_native": "📕 PDF 原生加密",
                         "office_native": "📘 Word/Excel/PowerPoint 原生加密"}
        out_text = self.custom_output_dir if self.output_mode.get() == "custom" else "原始檔案所在的資料夾"
        lines = [
            f"📁  已選擇 {len(self.selected_paths)} 個項目",
            f"🌍  加密方式：{method_names.get(self.encrypt_method.get(), '')}",
            f"📂  輸出位置：{out_text}",
        ]
        self.summary_label.configure(text="\n".join(lines))

    def _log_encrypt(self, msg):
        self.encrypt_log.configure(state="normal")
        self.encrypt_log.insert("end", msg + "\n")
        self.encrypt_log.see("end")
        self.encrypt_log.configure(state="disabled")

    def _target_dir(self, original_dir):
        if self.output_mode.get() == "custom" and self.custom_output_dir:
            os.makedirs(self.custom_output_dir, exist_ok=True)
            return self.custom_output_dir
        return original_dir

    def _start_encrypt(self):
        if not self.selected_paths:
            messagebox.showwarning("尚未選擇檔案", "請先回到步驟 ① 選擇要加密的檔案或資料夾。")
            return
        pw1 = self.encrypt_pw1.get()
        if not pw1:
            messagebox.showwarning("密碼空白", "請先回到步驟 ③ 設定密碼。")
            return

        method = self.encrypt_method.get()
        paths = list(self.selected_paths)

        self.encrypt_start_btn.configure(state="disabled")
        self.encrypt_log.configure(state="normal")
        self.encrypt_log.delete("1.0", "end")
        self.encrypt_log.configure(state="disabled")
        self.encrypt_progress.set(0)

        t = threading.Thread(target=self._run_encrypt_job, args=(paths, method, pw1), daemon=True)
        t.start()

    def _run_encrypt_job(self, paths, method, password):
        try:
            if method == "pdf_native":
                self._job_pdf_encrypt(paths, password)
            elif method == "office_native":
                self._job_office_encrypt(paths, password)
        except Exception as e:
            self.progress_queue.put(("error", str(e)))
        finally:
            self.progress_queue.put(("done", None))

    def _job_pdf_encrypt(self, paths, password):
        if not PDF_AVAILABLE:
            self.progress_queue.put(("log", "❌ 缺少 pypdf 套件，請先安裝：pip install pypdf cryptography"))
            return
        all_files = []
        for path in paths:
            if os.path.isdir(path):
                for f in os.listdir(path):
                    if f.lower().endswith(".pdf"):
                        all_files.append(os.path.join(path, f))
            elif path.lower().endswith(".pdf"):
                all_files.append(path)
        if not all_files:
            self.progress_queue.put(("log", "⚠️ 沒有找到 PDF 檔案"))
            return
        total = len(all_files)
        last_out_dir = None
        for i, full in enumerate(all_files, 1):
            folder, filename = os.path.split(full)
            target = self._target_dir(folder)
            name, ext = os.path.splitext(filename)
            out_path = os.path.join(target, f"{name}_加密{ext}")
            try:
                EncryptionEngine.pdf_encrypt(full, out_path, password)
                self.progress_queue.put(("log", f"✅ {filename}"))
            except Exception as e:
                self.progress_queue.put(("log", f"❌ {filename} — {e}"))
            self.progress_queue.put(("progress", i / total))
            last_out_dir = target
        if last_out_dir:
            self.progress_queue.put(("open_folder", last_out_dir))

    def _job_office_encrypt(self, paths, password):
        if not OFFICE_AVAILABLE:
            self.progress_queue.put(("log", "❌ 找不到 Office 或 pywin32，此功能僅限已安裝正版 Office 的 Windows 電腦"))
            return
        groups = {"word": [], "excel": [], "ppt": []}
        last_out_dir = None
        for path in paths:
            targets = [path] if os.path.isfile(path) else [
                os.path.join(path, f) for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))
            ]
            for full in targets:
                kind = classify_file(full)
                if kind in groups:
                    folder, filename = os.path.split(full)
                    target = self._target_dir(folder)
                    name, ext = os.path.splitext(filename)
                    out_path = os.path.join(target, f"{name}_加密{ext}")
                    groups[kind].append((full, out_path))
                    last_out_dir = target
        total = sum(len(v) for v in groups.values())
        done = 0
        for kind, pairs in groups.items():
            if not pairs:
                continue
            try:
                results = EncryptionEngine.office_encrypt(pairs, password, kind)
                for filename, ok, err in results:
                    done += 1
                    if ok:
                        self.progress_queue.put(("log", f"✅ {filename}"))
                    else:
                        self.progress_queue.put(("log", f"❌ {filename} — {err}"))
                    if total:
                        self.progress_queue.put(("progress", done / total))
            except Exception as e:
                self.progress_queue.put(("log", f"❌ {kind} 處理發生錯誤：{e}"))
        if last_out_dir:
            self.progress_queue.put(("open_folder", last_out_dir))

    # =================================================================
    # 圖片轉加密 PDF 頁
    # =================================================================
    def _build_imgpdf_page(self):
        page = self._new_scroll_page("imgpdf")

        card1 = make_card(page)
        card1.pack(fill="x", pady=(4, 16))
        section_title(card1, "🖼️", "把多張圖片合併成一份加密 PDF",
                      "適合行政文書掃描件、身分證件等需要妥善保護的圖片資料。")
        btn_row = ctk.CTkFrame(card1, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(6, 6))
        pill_button(btn_row, "🖼️ 選擇圖片（可多選）", self._pick_images, "ghost", 200, 42).pack(side="left")
        pill_button(btn_row, "🗑️ 清除", self._clear_images, "ghost", 110, 42).pack(side="left", padx=8)

        self.image_list_frame = ctk.CTkFrame(card1, fg_color=THEME["card_alt"], corner_radius=12)
        self.image_list_frame.pack(fill="x", padx=20, pady=(6, 18))
        self._refresh_image_list()

        card2 = make_card(page)
        card2.pack(fill="x", pady=(0, 16))
        section_title(card2, "🔑", "設定 PDF 密碼")
        self.imgpdf_pw1 = tk.StringVar()
        self.imgpdf_pw2 = tk.StringVar()
        row1 = ctk.CTkFrame(card2, fg_color="transparent")
        row1.pack(fill="x", padx=20, pady=(6, 5))
        ctk.CTkLabel(row1, text="密碼", width=90, anchor="w", font=(FONT_FAMILY, 13)).pack(side="left")
        ctk.CTkEntry(row1, textvariable=self.imgpdf_pw1, show="●", width=300, height=40,
                     corner_radius=10, font=(FONT_FAMILY, 13)).pack(side="left")
        row2 = ctk.CTkFrame(card2, fg_color="transparent")
        row2.pack(fill="x", padx=20, pady=(5, 18))
        ctk.CTkLabel(row2, text="確認密碼", width=90, anchor="w", font=(FONT_FAMILY, 13)).pack(side="left")
        ctk.CTkEntry(row2, textvariable=self.imgpdf_pw2, show="●", width=300, height=40,
                     corner_radius=10, font=(FONT_FAMILY, 13)).pack(side="left")

        pill_button(page, "📄 產生加密 PDF", self._start_image_to_pdf, "primary", 190, 46).pack(anchor="w", pady=(0, 12))

        log_card = make_card(page, fg_color=THEME["console_bg"], border_width=0, corner_radius=16)
        log_card.pack(fill="both", expand=True, pady=(0, 22))
        self.imgpdf_log = ctk.CTkTextbox(log_card, height=160, font=("Consolas", 12),
                                          fg_color=THEME["console_bg"], text_color=THEME["console_text"],
                                          corner_radius=14)
        self.imgpdf_log.pack(fill="both", expand=True, padx=4, pady=4)
        self.imgpdf_log.configure(state="disabled")

        if not IMAGE_AVAILABLE:
            ctk.CTkLabel(page, text="⚠️ 未偵測到 Pillow 套件，此功能暫時無法使用（pip install pillow）",
                         font=(FONT_FAMILY, 12), text_color=THEME["error"]).pack(anchor="w", pady=(0, 10))

    def _refresh_image_list(self):
        for w in self.image_list_frame.winfo_children():
            w.destroy()
        if not self.selected_images:
            ctk.CTkLabel(self.image_list_frame, text="尚未選擇任何圖片", font=(FONT_FAMILY, 12),
                         text_color=THEME["subtext"]).pack(pady=20)
            return
        for idx, path in enumerate(self.selected_images, 1):
            row = ctk.CTkFrame(self.image_list_frame, fg_color=THEME["card"], corner_radius=10,
                                border_width=1, border_color=THEME["card_border"])
            row.pack(fill="x", padx=10, pady=5)
            ctk.CTkLabel(row, text=f"{idx}. 🖼️", font=(FONT_FAMILY, 12)).pack(side="left", padx=(14, 6), pady=9)
            ctk.CTkLabel(row, text=os.path.basename(path), font=(FONT_FAMILY, 12), text_color=THEME["text"],
                         anchor="w").pack(side="left", fill="x", expand=True)
            ctk.CTkButton(row, text="✕", width=30, height=30, corner_radius=15,
                          fg_color="transparent", hover_color=THEME["error_soft"],
                          text_color=THEME["subtext"],
                          command=lambda p=path: self._remove_image(p)).pack(side="right", padx=8)

    def _remove_image(self, path):
        self.selected_images = [p for p in self.selected_images if p != path]
        self._refresh_image_list()

    def _pick_images(self):
        paths = filedialog.askopenfilenames(
            title="選擇圖片",
            filetypes=[("圖片檔", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp"), ("所有檔案", "*.*")]
        )
        for p in paths:
            self.selected_images.append(p)
        self._refresh_image_list()

    def _clear_images(self):
        self.selected_images = []
        self._refresh_image_list()

    def _log_imgpdf(self, msg):
        self.imgpdf_log.configure(state="normal")
        self.imgpdf_log.insert("end", msg + "\n")
        self.imgpdf_log.see("end")
        self.imgpdf_log.configure(state="disabled")

    def _start_image_to_pdf(self):
        if not IMAGE_AVAILABLE or not PDF_AVAILABLE:
            messagebox.showerror("缺少套件", "此功能需要 pillow 及 pypdf 套件，請先安裝。")
            return
        if not self.selected_images:
            messagebox.showwarning("尚未選擇圖片", "請先選擇至少一張圖片。")
            return
        pw1, pw2 = self.imgpdf_pw1.get(), self.imgpdf_pw2.get()
        if not pw1 or pw1 != pw2:
            messagebox.showwarning("密碼有誤", "請輸入密碼，並確認兩次輸入一致。")
            return
        save_path = filedialog.asksaveasfilename(
            title="儲存加密 PDF", defaultextension=".pdf",
            filetypes=[("PDF 檔案", "*.pdf")], initialfile="合併加密文件.pdf"
        )
        if not save_path:
            return
        try:
            EncryptionEngine.images_to_encrypted_pdf(self.selected_images, save_path, pw1)
            self._log_imgpdf(f"✅ 已產生加密 PDF：{save_path}")
            open_in_file_explorer(os.path.dirname(save_path))
        except Exception as e:
            self._log_imgpdf(f"❌ 失敗：{e}")

    # =================================================================
    # 密碼工具頁
    # =================================================================
    def _build_pwtools_page(self):
        page = self._new_scroll_page("pwtools")

        card1 = make_card(page)
        card1.pack(fill="x", pady=(4, 16))
        section_title(card1, "🎲", "密碼產生器")

        opt_row = ctk.CTkFrame(card1, fg_color="transparent")
        opt_row.pack(fill="x", padx=20, pady=(6, 6))
        self.gen_upper = tk.BooleanVar(value=True)
        self.gen_lower = tk.BooleanVar(value=True)
        self.gen_digits = tk.BooleanVar(value=True)
        self.gen_symbols = tk.BooleanVar(value=True)
        for var, label in [(self.gen_upper, "大寫 A-Z"), (self.gen_lower, "小寫 a-z"),
                            (self.gen_digits, "數字 0-9"), (self.gen_symbols, "符號 !@#$")]:
            ctk.CTkCheckBox(opt_row, text=label, variable=var, font=(FONT_FAMILY, 12),
                             fg_color=THEME["accent"]).pack(side="left", padx=6)

        len_row = ctk.CTkFrame(card1, fg_color="transparent")
        len_row.pack(fill="x", padx=20, pady=(4, 6))
        ctk.CTkLabel(len_row, text="長度", font=(FONT_FAMILY, 12)).pack(side="left")
        self.gen_length = tk.IntVar(value=16)
        ctk.CTkSlider(len_row, from_=8, to=64, number_of_steps=56, variable=self.gen_length,
                      width=200, command=lambda v: None).pack(side="left", padx=10)
        self.gen_length_label = ctk.CTkLabel(len_row, text="16", font=(FONT_FAMILY, 12, "bold"),
                                              text_color=THEME["accent"])
        self.gen_length_label.pack(side="left", padx=(0, 10))
        self.gen_length.trace_add("write", lambda *_: self.gen_length_label.configure(text=str(self.gen_length.get())))
        pill_button(len_row, "🎲 產生密碼", self._generate_standalone_password, "primary", 140, 40).pack(side="left")

        result_row = ctk.CTkFrame(card1, fg_color="transparent")
        result_row.pack(fill="x", padx=20, pady=(6, 20))
        self.generated_pw_var = tk.StringVar()
        ctk.CTkEntry(result_row, textvariable=self.generated_pw_var, width=350, height=40,
                     corner_radius=10, font=("Consolas", 14)).pack(side="left")
        pill_button(result_row, "📋 複製", self._copy_generated_password, "ghost", 100, 40).pack(side="left", padx=8)

        card2 = make_card(page)
        card2.pack(fill="x", pady=(0, 22))
        section_title(card2, "🧪", "密碼強度檢測")
        self.check_pw_var = tk.StringVar()
        row = ctk.CTkFrame(card2, fg_color="transparent")
        row.pack(fill="x", padx=20, pady=(6, 4))
        ctk.CTkEntry(row, textvariable=self.check_pw_var, show="●", width=310, height=40,
                     corner_radius=10, font=(FONT_FAMILY, 13)).pack(side="left")
        self.check_strength_label = ctk.CTkLabel(card2, text="密碼強度：尚未輸入", font=(FONT_FAMILY, 12),
                                                  text_color=THEME["subtext"])
        self.check_strength_label.pack(anchor="w", padx=20, pady=(10, 3))
        self.check_strength_bar = ctk.CTkProgressBar(card2, width=310, height=11, corner_radius=6)
        self.check_strength_bar.set(0)
        self.check_strength_bar.pack(anchor="w", padx=20, pady=(0, 20))
        self.check_pw_var.trace_add("write", self._update_check_strength)

    def _generate_standalone_password(self):
        pw = generate_password(
            self.gen_length.get(), self.gen_upper.get(), self.gen_lower.get(),
            self.gen_digits.get(), self.gen_symbols.get()
        )
        self.generated_pw_var.set(pw)

    def _copy_generated_password(self):
        pw = self.generated_pw_var.get()
        if pw:
            self.root.clipboard_clear()
            self.root.clipboard_append(pw)
            messagebox.showinfo("已複製", "密碼已複製到剪貼簿。")

    def _update_check_strength(self, *_):
        score, text, color_key = password_strength(self.check_pw_var.get())
        self.check_strength_label.configure(text=f"密碼強度：{text}", text_color=self._strength_color(color_key))
        self.check_strength_bar.set(score / 100)
        self.check_strength_bar.configure(progress_color=self._strength_color(color_key))

    # =================================================================
    # 使用說明頁（分類卡片 + 可展開 FAQ）
    # =================================================================
    def _build_help_page(self):
        page = self._new_scroll_page("help")

        # ---- 快速上手 ----
        card_quick = make_card(page)
        card_quick.pack(fill="x", pady=(4, 16))
        section_title(card_quick, "🚀", "快速上手", "只要 5 步驟就能完成加密，每一步都可以按「上一步」回去修改。")
        flow_row = ctk.CTkFrame(card_quick, fg_color="transparent")
        flow_row.pack(fill="x", padx=20, pady=(4, 20))
        flow_steps = ["📁 選擇檔案", "🌍 加密方式", "🔑 設定密碼", "📂 輸出位置", "🚀 開始加密"]
        for i, s in enumerate(flow_steps):
            chip(flow_row, s).pack(side="left")
            if i < len(flow_steps) - 1:
                ctk.CTkLabel(flow_row, text="→", font=(FONT_FAMILY, 14), text_color=THEME["subtext"]).pack(
                    side="left", padx=6)

        # ---- 安裝需求 ----
        card_install = make_card(page)
        card_install.pack(fill="x", pady=(0, 16))
        section_title(card_install, "📦", "安裝需求")
        install_box = ctk.CTkFrame(card_install, fg_color=THEME["console_bg"], corner_radius=12)
        install_box.pack(fill="x", padx=20, pady=(4, 8))
        install_text = (
            "pip install customtkinter pypdf cryptography pillow\n\n"
            "# 選用（若要使用 Office 原生加密，僅限 Windows + 已安裝正版 Office）：\n"
            "pip install pywin32"
        )
        ctk.CTkLabel(install_box, text=install_text, font=("Consolas", 13), justify="left",
                     text_color=THEME["console_text"], anchor="w").pack(anchor="w", padx=16, pady=14)
        ctk.CTkLabel(card_install, text="安裝完成後，在終端機執行：python 萬用加密工具.py",
                     font=(FONT_FAMILY, 12), text_color=THEME["subtext"]).pack(anchor="w", padx=20, pady=(0, 18))

        # ---- 兩種加密方式比較 ----
        card_methods = make_card(page)
        card_methods.pack(fill="x", pady=(0, 16))
        section_title(card_methods, "🔑", "兩種加密方式怎麼選？")
        compare_row = ctk.CTkFrame(card_methods, fg_color="transparent")
        compare_row.pack(fill="x", padx=20, pady=(4, 20))
        compare_row.grid_columnconfigure((0, 1), weight=1, uniform="cmp")

        def method_card(parent, col, icon, title, points, enabled):
            box = ctk.CTkFrame(parent, fg_color=THEME["card_alt"], corner_radius=14,
                                border_width=1, border_color=THEME["card_border"])
            box.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 8, 0 if col == 1 else 8))
            head = ctk.CTkFrame(box, fg_color="transparent")
            head.pack(anchor="w", padx=16, pady=(16, 8))
            icon_badge(head, icon, size=34).pack(side="left", padx=(0, 10))
            ctk.CTkLabel(head, text=title, font=(FONT_FAMILY, 14, "bold"),
                         text_color=THEME["text"]).pack(side="left")
            for pt in points:
                ctk.CTkLabel(box, text=f"•  {pt}", font=(FONT_FAMILY, 12), text_color=THEME["subtext"],
                             justify="left", wraplength=340, anchor="w").pack(anchor="w", padx=16, pady=2)
            status_text = "🟢 目前可用" if enabled else "⚪ 目前未偵測到，需安裝對應套件"
            ctk.CTkLabel(box, text=status_text, font=(FONT_FAMILY, 11, "bold"),
                         text_color=THEME["success"] if enabled else THEME["subtext"]).pack(
                anchor="w", padx=16, pady=(8, 16))

        method_card(compare_row, 0, "📕", "PDF 原生加密", [
            "加密後用任何 PDF 閱讀器雙擊打開，會直接跳出密碼視窗。",
            "不需要額外安裝任何工具。",
            "僅適用 PDF 檔案。",
        ], PDF_AVAILABLE)
        method_card(compare_row, 1, "📘", "Office 原生加密", [
            "加密後用 Word/Excel/PowerPoint 打開會直接要求密碼。",
            "不需要額外安裝任何工具。",
            "限制：電腦需為 Windows，且需安裝正版 Microsoft Office。",
        ], OFFICE_AVAILABLE)

        # ---- 密碼安全建議 ----
        card_pw = make_card(page, fg_color=THEME["warn_soft"], border_width=0, corner_radius=16)
        card_pw.pack(fill="x", pady=(0, 16))
        head_pw = ctk.CTkFrame(card_pw, fg_color="transparent")
        head_pw.pack(anchor="w", padx=22, pady=(18, 6))
        ctk.CTkLabel(head_pw, text="🔒", font=(FONT_FAMILY, 16)).pack(side="left", padx=(0, 8))
        ctk.CTkLabel(head_pw, text="密碼安全建議", font=(FONT_FAMILY, 15, "bold"),
                     text_color=THEME["warn"]).pack(side="left")
        ctk.CTkLabel(card_pw, justify="left", wraplength=900, font=(FONT_FAMILY, 13), text_color=THEME["text"], text=(
            "• 建議使用 12 碼以上、混合大小寫字母、數字、符號的密碼。\n"
            "• 密碼一旦遺失，沒有任何方式可以救回加密檔案內容，請務必妥善保存。\n"
            "• 本程式所有運算都在你自己的電腦上完成，不會把檔案或密碼上傳到網路。"
        )).pack(anchor="w", padx=22, pady=(0, 20))

        # ---- 常見問題（可展開收合） ----
        card_faq = make_card(page)
        card_faq.pack(fill="x", pady=(0, 24))
        section_title(card_faq, "❓", "常見問題", "點一下問題即可展開／收合答案。")
        faqs = [
            ("這個工具會不會把我的檔案或密碼上傳到網路？",
             "不會。所有運算都在你自己的電腦上完成，沒有任何網路連線行為。"),
            ("為什麼「Word/Excel/PowerPoint 原生加密」選項是灰色、不能選？",
             "代表你的電腦目前偵測不到已安裝的正版 Office，或作業系統不是 Windows。"
             "請改用「PDF 原生加密」，或改用其他方式保護你的檔案。"),
            ("收到加密檔案的人該怎麼打開？",
             "PDF 原生加密：用任何 PDF 閱讀器打開，會直接跳出密碼視窗即可，不需要額外工具。\n"
             "Office 原生加密：用 Office 打開會直接要求密碼，不需要額外工具。"),
        ]
        for q, a in faqs:
            self._build_faq_item(card_faq, q, a)
        ctk.CTkLabel(card_faq, text="", font=(FONT_FAMILY, 4)).pack(pady=4)

    def _build_faq_item(self, parent, question, answer):
        wrap = ctk.CTkFrame(parent, fg_color=THEME["card_alt"], corner_radius=12, cursor="hand2")
        wrap.pack(fill="x", padx=20, pady=6)

        state = {"open": False}

        header = ctk.CTkFrame(wrap, fg_color="transparent", cursor="hand2")
        header.pack(fill="x")
        arrow_label = ctk.CTkLabel(header, text="▸", font=(FONT_FAMILY, 15, "bold"),
                                    text_color=THEME["accent"], width=26, cursor="hand2")
        arrow_label.pack(side="left", padx=(16, 4), pady=14)
        q_label = ctk.CTkLabel(header, text=question, font=(FONT_FAMILY, 13, "bold"),
                                text_color=THEME["text"], anchor="w", justify="left",
                                wraplength=700, cursor="hand2")
        q_label.pack(side="left", fill="x", expand=True, pady=14)

        answer_label = ctk.CTkLabel(wrap, text=answer, font=(FONT_FAMILY, 12), text_color=THEME["subtext"],
                                     anchor="w", justify="left", wraplength=680)

        def toggle(event=None):
            if state["open"]:
                answer_label.pack_forget()
                arrow_label.configure(text="▸")
                state["open"] = False
            else:
                answer_label.pack(fill="x", padx=(52, 20), pady=(0, 16))
                arrow_label.configure(text="▾")
                state["open"] = True

        for w in (wrap, header, arrow_label, q_label):
            w.bind("<Button-1>", toggle)

    # =================================================================
    # 佇列輪詢
    # =================================================================
    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.progress_queue.get_nowait()
                if kind == "log":
                    self._log_encrypt(payload)
                elif kind == "progress":
                    self.encrypt_progress.set(payload)
                elif kind == "error":
                    self._log_encrypt(f"❌ 發生錯誤：{payload}")
                elif kind == "open_folder":
                    open_in_file_explorer(payload)
                elif kind == "done":
                    self.encrypt_progress.set(1.0)
                    self.encrypt_start_btn.configure(state="normal")
                    self._log_encrypt("—— 全部處理完成，已為你打開輸出資料夾 ——")
        except queue.Empty:
            pass
        self.root.after(150, self._poll_queue)


def _fix_windows_dpi_blur():
    """
    修正 Windows 高解析度螢幕文字模糊的根本原因：若沒有正確宣告
    DPI 感知層級，Windows 會用「點陣圖縮放」硬把畫面放大，導致模糊。

    這裡改用更完整、由高到低的相容順序：
      1. SetProcessDpiAwarenessContext(PER_MONITOR_AWARE_V2) —— 效果最好，
         文字最銳利，Windows 10 1703 以後支援。
      2. shcore.SetProcessDpiAwareness(2)  —— Per-Monitor（較舊 Windows）
      3. shcore.SetProcessDpiAwareness(1)  —— System DPI Aware
      4. user32.SetProcessDPIAware()       —— 最舊版相容備援

    同時會先檢查目前是否「已經」被設定過（例如被其他模組或執行環境
    搶先設定），避免重複呼叫互相衝突而讓系統退回模糊的縮放模式。
    僅在 Windows 上執行，其他系統不受影響。
    """
    if platform.system() != "Windows":
        return
    try:
        import ctypes

        # 先偵測目前的 DPI 感知狀態，如果已經是「非預設」的感知模式，
        # 就不要再重複設定，避免衝突。
        already_aware = False
        try:
            awareness = ctypes.c_int()
            ctypes.windll.shcore.GetProcessDpiAwareness(0, ctypes.byref(awareness))
            if awareness.value != 0:  # 0 = DPI_AWARENESS_UNAWARE
                already_aware = True
        except Exception:
            pass

        if already_aware:
            return

        # 優先使用 PerMonitorV2（畫面最銳利）
        try:
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = ctypes.c_void_p(-4)
            if ctypes.windll.user32.SetProcessDpiAwarenessContext(
                    DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2):
                return
        except Exception:
            pass

        # 退而求其次：Per-Monitor DPI awareness
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
            return
        except Exception:
            pass

        # 再退一步：System DPI aware
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
            return
        except Exception:
            pass

        # 最後備援：舊版 API
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    except Exception:
        pass


def main():
    _fix_windows_dpi_blur()

    if not CTK_AVAILABLE:
        root = tk.Tk()
        root.title(APP_TITLE)
        tk.Label(root, text="缺少必要套件 customtkinter\n\n請在終端機執行：\npip install customtkinter\n\n"
                             "安裝完成後請重新開啟本程式。",
                 font=("Microsoft JhengHei UI", 13), padx=30, pady=30, justify="left").pack()
        root.mainloop()
        return

    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    app_root = ctk.CTk()
    App(app_root)
    app_root.mainloop()


if __name__ == "__main__":
    main()