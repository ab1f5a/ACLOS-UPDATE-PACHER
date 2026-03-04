import os
import sys
import ctypes
import shutil
import time
import requests
import webbrowser
import string
import threading
from ctypes import wintypes

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

import pythoncom
from win32com.client import Dispatch
from tkinter import filedialog, Tk, ttk, messagebox, scrolledtext
import tkinter as tk
from tkinter import font as tkfont


def _get_dpi_scale() -> float:
    try:
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
        ctypes.windll.user32.ReleaseDC(0, hdc)
        return dpi / 96.0
    except Exception:
        return 1.0

_SCALE = _get_dpi_scale()
UI_SCALE = _SCALE * 1.1
FONT_SCALE = _SCALE * 0.85

def S(value: int) -> int:
    return max(1, round(value * UI_SCALE))


def FS(size: int) -> int:
    return max(6, round(size * FONT_SCALE))

CURRENT_VERSION = "2.0.0"
VERSION_URL = "https://ab1f5a.net/v2/version.json"
WINDOW_TITLE = f"ACLOS Tools v{CURRENT_VERSION}"

COLORS = {
    'bg': '#F5F7FA',
    'bg_alt': '#EEF1F6',
    'fg': '#1A1D2E',
    'fg_muted': '#6B7280',
    'accent': '#4F6EF7',
    'accent_hover': '#3B5BE0',
    'accent_light': '#EBF0FF',
    'success': '#16A34A',
    'success_light': '#DCFCE7',
    'success_hover': '#15803D',
    'warning': '#D97706',
    'warning_light': '#FEF3C7',
    'error': '#DC2626',
    'error_light': '#FEE2E2',
    'info': '#0284C7',
    'info_light': '#E0F2FE',
    'surface': '#FFFFFF',
    'surface_hover': '#F0F4FF',
    'border': '#E2E6EE',
    'border_focus': '#4F6EF7',
    'shadow': '#00000015',
    'text_light': '#9CA3AF',
    'titlebar': '#13152A',
}

def get_resource_path(relative_path):
    """ 获取程序运行时的资源绝对路径 """
    # Nuitka/PyInstaller onefile 模式会将路径存放在这里
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def get_unicode_shell_path(csidl):
    buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetSpecialFolderPathW(None, buf, csidl, False)
    return buf.value

UNICODE_APPDATA = get_unicode_shell_path(0x001a)
DEPLOY_DIR = os.path.join(UNICODE_APPDATA, "ACLOS Tools")
PATCHER_EXE_NAME = "ACLOS UPDATE PATCHER.EXE"

VC_VERSION = "1.0"
ROOT_DIR_NAME = "ACLOS"
SUB_PATH = os.path.join("Launcher", "resources", "app.asar")
CONFIG_DIR = os.path.join(os.environ.get('APPDATA', UNICODE_APPDATA), "ACLOS")
PATH_CACHE = os.path.join(CONFIG_DIR, "asar_path_v2.cfg")

LANG_MAP = {
    "1": ("zh_CN", "中文_中国"),
    "2": ("en_US", "英语_美国"),
    "3": ("de_DE", "*德语_德国"),
    "4": ("it_IT", "意语_意大利"),
    "5": ("ja_JP", "日语_日本"),
    "6": ("ko_KR", "韩语_韩国"),
    "7": ("pt_PT", "*葡萄牙语_葡萄牙"),
    "8": ("ru_RU", "*俄语_俄国")
}

def make_rounded_button(parent, text, command, bg, fg, hover_bg=None,
                        font_size=FS(10), font_weight="normal", width=None,
                        padx=20, pady=8, radius=8):
    if hover_bg is None:
        hover_bg = COLORS['accent_hover']

    frame = tk.Frame(parent, bg=bg, cursor="hand2")
    if width:
        frame.config(width=width)

    label = tk.Label(
        frame, text=text,
        font=("Microsoft YaHei UI", font_size, font_weight),
        bg=bg, fg=fg, padx=padx, pady=pady, cursor="hand2"
    )
    label.pack()

    def on_enter(e):
        frame.config(bg=hover_bg)
        label.config(bg=hover_bg)

    def on_leave(e):
        frame.config(bg=bg)
        label.config(bg=bg)

    def on_click(e):
        command()

    for w in (frame, label):
        w.bind("<Enter>", on_enter)
        w.bind("<Leave>", on_leave)
        w.bind("<Button-1>", on_click)

    return frame


class RoundedButton(tk.Canvas):
    def __init__(self, parent, text, command, bg, fg, hover_bg=None,
                 font_size=FS(10), font_weight="normal", width=120, height=36,
                 radius=10, **kwargs):
        super().__init__(parent, width=width, height=height,
                         bg=parent.cget('bg'), highlightthickness=0, cursor="hand2", **kwargs)
        self.command = command
        self.bg_color = bg
        self.hover_color = hover_bg or self._darken(bg)
        self.current_fg = fg
        self.hover_fg = fg
        self._is_light_bg = self._is_light(bg)
        self.text_color = fg
        self.radius = radius
        self.width = width
        self.height = height
        self.font = tkfont.Font(family="Microsoft YaHei UI", size=font_size, weight=font_weight)
        self.text_str = text
        self._hovered = False
        self._draw(self.bg_color, self.text_color)

        self.bind("<Enter>", lambda e: self._animate_hover(True))
        self.bind("<Leave>", lambda e: self._animate_hover(False))
        self.bind("<Button-1>", lambda e: self._on_click())
        self.bind("<ButtonRelease-1>", lambda e: self._on_release())

    def _is_light(self, hex_color):
        try:
            r = int(hex_color[1:3], 16)
            g = int(hex_color[3:5], 16)
            b = int(hex_color[5:7], 16)
            luminance = (0.299 * r + 0.587 * g + 0.114 * b)
            return luminance > 180
        except:
            return False

    def _darken(self, hex_color):
        try:
            r = int(hex_color[1:3], 16)
            g = int(hex_color[3:5], 16)
            b = int(hex_color[5:7], 16)
            return f"#{max(0,r-20):02x}{max(0,g-20):02x}{max(0,b-20):02x}"
        except:
            return hex_color

    def _draw(self, fill, text_color):
        self.delete("all")
        r = self.radius
        w, h = self.width, self.height
        self.create_polygon(
            r, 0, w - r, 0,
            w, 0, w, r,
            w, h - r, w, h,
            w - r, h, r, h,
            0, h, 0, h - r,
            0, r, 0, 0,
            r, 0,
            smooth=True, fill=fill, outline=""
        )
        self.create_text(w // 2, h // 2, text=self.text_str,
                         font=self.font, fill=text_color)

    def _animate_hover(self, entering):
        self._hovered = entering
        if entering:
            self._draw(self.hover_color, self.text_color)
        else:
            self._draw(self.bg_color, self.text_color)

    def _on_click(self):
        self._draw(self._darken(self.hover_color), self.text_color)
        self.command()

    def _on_release(self):
        self._draw(self.hover_color if self._hovered else self.bg_color, self.text_color)


class ACLOSToolsGUI:

    F = "Microsoft YaHei UI"

    def __init__(self, root):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.configure(bg=COLORS['bg'])
        self.root.resizable(False, False)

        # 无边框 
        self.root.overrideredirect(True)
        self._drag_x = 0
        self._drag_y = 0

        w, h = S(900), S(560) + S(34)
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f'{w}x{h}+{x}+{y}')

        self._version_ok = False
        self._nav_frames  = {}
        self._nav_buttons = {}
        self._current_page = None

        self._setup_styles()
        self._build_titlebar()
        self._build_shell()
        self._show_page("home")
        self.root.after(400, self._check_version_startup)

    def _build_titlebar(self):
        F = self.F
        bar = tk.Frame(self.root, bg=COLORS['titlebar'], height=S(34))
        bar.pack(fill=tk.X, side=tk.TOP)
        bar.pack_propagate(False)

        left = tk.Frame(bar, bg=COLORS['titlebar'])
        left.pack(side=tk.LEFT, padx=S(14), fill=tk.Y)

        dot = tk.Frame(left, bg=COLORS['accent'], width=S(8), height=S(8))
        dot.pack(side=tk.LEFT, anchor="center")
        dot.place_configure(rely=0.5, anchor="w")
        dot.pack_forget()
        dot.place(relx=0, rely=0.5, anchor="w")

        tk.Label(left, text="  " + WINDOW_TITLE,
                 font=(F, FS(8)), bg=COLORS['titlebar'],
                 fg="#8B92B8").pack(side=tk.LEFT, anchor="center")

        right = tk.Frame(bar, bg=COLORS['titlebar'])
        right.pack(side=tk.RIGHT, fill=tk.Y)

        self.status_label = tk.Label(right, text="● 正在初始化...",
                                     font=(F, FS(8)), bg=COLORS['titlebar'],
                                     fg=COLORS['warning'], padx=S(10))
        self.status_label.pack(side=tk.LEFT, anchor="center")

        min_btn = tk.Label(right, text="  ─  ", font=(F, FS(9)),
                           bg=COLORS['titlebar'], fg="#8B92B8",
                           cursor="hand2", pady=0)
        min_btn.pack(side=tk.LEFT, fill=tk.Y)
        min_btn.bind("<Enter>",    lambda e: min_btn.config(bg="#2A2D45", fg="white"))
        min_btn.bind("<Leave>",    lambda e: min_btn.config(bg=COLORS['titlebar'], fg="#8B92B8"))
        min_btn.bind("<Button-1>", lambda e: self._minimize())

        close_btn = tk.Label(right, text="  ✕  ", font=(F, FS(9)),
                             bg=COLORS['titlebar'], fg="#8B92B8",
                             cursor="hand2", pady=0)
        close_btn.pack(side=tk.LEFT, fill=tk.Y)
        close_btn.bind("<Enter>",    lambda e: close_btn.config(bg=COLORS['error'], fg="white"))
        close_btn.bind("<Leave>",    lambda e: close_btn.config(bg=COLORS['titlebar'], fg="#8B92B8"))
        close_btn.bind("<Button-1>", lambda e: self.root.quit())

        for w in (bar, left):
            w.bind("<ButtonPress-1>", self._drag_start)
            w.bind("<B1-Motion>",     self._drag_move)
        for child in left.winfo_children():
            child.bind("<ButtonPress-1>", self._drag_start)
            child.bind("<B1-Motion>",     self._drag_move)

    def _drag_start(self, event):
        self._drag_x = event.x_root - self.root.winfo_x()
        self._drag_y = event.y_root - self.root.winfo_y()

    def _drag_move(self, event):
        x = event.x_root - self._drag_x
        y = event.y_root - self._drag_y
        self.root.geometry(f"+{x}+{y}")

    def _minimize(self):
        self.root.overrideredirect(False)
        self.root.iconify()
        def _restore(event):
            if self.root.state() == 'normal':
                self.root.overrideredirect(True)
                self.root.unbind('<Map>')
        self.root.bind('<Map>', _restore)

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TScrollbar',
                        background=COLORS['bg_alt'],
                        troughcolor=COLORS['bg'],
                        bordercolor=COLORS['border'],
                        arrowcolor=COLORS['fg_muted'])

    def _build_shell(self):
        F = self.F
        self.sidebar = tk.Frame(self.root, bg=COLORS['surface'], width=S(210))
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)

        tk.Frame(self.root, bg=COLORS['border'], width=1).pack(side=tk.LEFT, fill=tk.Y)

        self.content_host = tk.Frame(self.root, bg=COLORS['bg'])
        self.content_host.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        logo_wrap = tk.Frame(self.sidebar, bg=COLORS['surface'])
        logo_wrap.pack(fill=tk.X, pady=(S(20), S(6)))

        logo_box = tk.Frame(logo_wrap, bg=COLORS['accent'], width=S(40), height=S(40))
        logo_box.pack()
        logo_box.pack_propagate(False)
        tk.Label(logo_box, text="AT", font=(F, FS(13), "bold"),
                 bg=COLORS['accent'], fg="white").place(relx=.5, rely=.5, anchor="center")

        tk.Label(self.sidebar, text="ACLOS Tools",
                 font=(F, FS(11), "bold"), bg=COLORS['surface'], fg=COLORS['fg']).pack(pady=(6, 1))
        tk.Label(self.sidebar, text=f"v{CURRENT_VERSION}",
                 font=(F, FS(8)), bg=COLORS['surface'], fg=COLORS['text_light']).pack()

        tk.Frame(self.sidebar, bg=COLORS['border'], height=1).pack(fill=tk.X, padx=14, pady=12)
        tk.Label(self.sidebar, text="导 航",
                 font=(F, FS(8), "bold"), bg=COLORS['surface'],
                 fg=COLORS['text_light']).pack(anchor="w", padx=16, pady=(0, 4))

        self._add_nav("home",   "🏠  首页",       self._show_home,    locked=False)
        self._add_nav("deploy", "⚙  部署绕过补丁", self._show_deploy,  locked=True)
        self._add_nav("vc",     "🎵  配音修改",    self._show_vc,      locked=True)

        tk.Frame(self.sidebar, bg=COLORS['border'], height=1).pack(
            side=tk.BOTTOM, fill=tk.X, padx=14, pady=(0, 0))
        btm = tk.Frame(self.sidebar, bg=COLORS['surface'])
        btm.pack(side=tk.BOTTOM, fill=tk.X, padx=16, pady=10)
        tk.Label(btm, text="© 2025 ab1f5a", font=(F, FS(7)),
                 bg=COLORS['surface'], fg=COLORS['text_light']).pack(anchor="w")

        links_row = tk.Frame(self.sidebar, bg=COLORS['surface'])
        links_row.pack(side=tk.BOTTOM, fill=tk.X, padx=S(10), pady=(0, S(6)))

        def _link_btn(parent, text, url, bg, fg, hover_bg):
            btn = tk.Label(parent, text=text,
                           font=(F, FS(8)), bg=bg, fg=fg,
                           cursor="hand2", padx=S(6), pady=S(5),
                           relief=tk.FLAT)
            btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, S(4)))
            btn.bind("<Enter>",    lambda e: btn.config(bg=hover_bg, fg="white"))
            btn.bind("<Leave>",    lambda e: btn.config(bg=bg, fg=fg))
            btn.bind("<Button-1>", lambda e: webbrowser.open(url))
            return btn

        _link_btn(links_row,
                  text="哔哩哔哩",
                  url="https://space.bilibili.com/321503341",
                  bg="#F0F4FF", fg="#4F6EF7", hover_bg="#4F6EF7")

        _link_btn(links_row,
                  text="GitHub",
                  url="https://github.com/ab1f5a/ACLOS-Tools",
                  bg="#F0F4FF", fg="#4F6EF7", hover_bg="#4F6EF7")

        tk.Frame(self.sidebar, bg=COLORS['border'], height=1).pack(
            side=tk.BOTTOM, fill=tk.X, padx=14, pady=(0, S(4)))

        for name, builder in [("home",   self._build_home_page),
                               ("deploy", self._build_deploy_page),
                               ("vc",     self._build_vc_page)]:
            frame = tk.Frame(self.content_host, bg=COLORS['bg'])
            self._nav_frames[name] = frame
            builder(frame)

    def _add_nav(self, name, label, command, locked=False):
        F = self.F
        frame = tk.Frame(self.sidebar, bg=COLORS['surface'],
                         cursor="hand2", padx=S(12), pady=S(7))
        frame.pack(fill=tk.X, padx=S(8), pady=1)

        lbl = tk.Label(frame, text=label, font=(F, FS(9)),
                       bg=COLORS['surface'], fg=COLORS['fg_muted'], cursor="hand2")
        lbl.pack(side=tk.LEFT)

        lock_lbl = None
        if locked:
            lock_lbl = tk.Label(frame, text="🔒", font=(F, FS(7)),
                                bg=COLORS['surface'], fg=COLORS['text_light'])
            lock_lbl.pack(side=tk.RIGHT)

        self._nav_buttons[name] = (frame, lbl, lock_lbl)

        def clicked():
            if locked and not self._version_ok:
                self._flash_locked(name)
                return
            command()

        def on_enter(e):
            if self._current_page != name:
                frame.config(bg=COLORS['bg_alt'])
                lbl.config(bg=COLORS['bg_alt'])
                if lock_lbl:
                    lock_lbl.config(bg=COLORS['bg_alt'])

        def on_leave(e):
            if self._current_page != name:
                frame.config(bg=COLORS['surface'])
                lbl.config(bg=COLORS['surface'])
                if lock_lbl:
                    lock_lbl.config(bg=COLORS['surface'])

        for w in ([frame, lbl] + ([lock_lbl] if lock_lbl else [])):
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)
            w.bind("<Button-1>", lambda e: clicked())

    def _set_nav_active(self, name):
        for n, (frame, lbl, lock_lbl) in self._nav_buttons.items():
            if n == name:
                frame.config(bg=COLORS['accent_light'])
                lbl.config(bg=COLORS['accent_light'],
                           fg=COLORS['accent'],
                           font=(self.F, FS(9), "bold"))
                if lock_lbl:
                    lock_lbl.config(bg=COLORS['accent_light'])
            else:
                frame.config(bg=COLORS['surface'])
                lbl.config(bg=COLORS['surface'],
                           fg=COLORS['fg_muted'],
                           font=(self.F, FS(9), "normal"))
                if lock_lbl:
                    lock_lbl.config(bg=COLORS['surface'])

    def _unlock_nav(self):
        for name in ("deploy", "vc"):
            _, lbl, lock_lbl = self._nav_buttons[name]
            if lock_lbl:
                lock_lbl.config(text="")

    def _flash_locked(self, name):
        frame, lbl, lock_lbl = self._nav_buttons[name]
        orig_fg = lbl.cget('fg')
        lbl.config(fg=COLORS['error'])
        self.root.after(400, lambda: lbl.config(fg=orig_fg))

    def _show_page(self, name):
        if self._current_page:
            self._nav_frames[self._current_page].pack_forget()
        self._nav_frames[name].pack(fill=tk.BOTH, expand=True)
        self._current_page = name
        self._set_nav_active(name)

    def _show_home(self):   self._show_page("home")
    def _show_deploy(self): self._show_page("deploy")
    def _show_vc(self):     self._show_page("vc")
    def _build_home_page(self, parent):
        F = self.F
        self._page_topbar(parent, "首页")

        body = tk.Frame(parent, bg=COLORS['bg'])
        body.pack(fill=tk.BOTH, expand=True, padx=S(20), pady=S(16))

        hero = tk.Frame(body, bg=COLORS['accent'],
                        highlightbackground=COLORS['accent_hover'], highlightthickness=1)
        hero.pack(fill=tk.X, pady=(0, S(12)))

        hero_inner = tk.Frame(hero, bg=COLORS['accent'])
        hero_inner.pack(fill=tk.X, padx=S(20), pady=S(18))

        tk.Label(hero_inner, text="ACLOS Tools",
                 font=(F, FS(22), "bold"), bg=COLORS['accent'], fg="white").pack(anchor="w")
        tk.Label(hero_inner, text=f"Version {CURRENT_VERSION}  ·  专为国服无畏契约玩家设计",
                 font=(F, FS(9)), bg=COLORS['accent'], fg="#C7D7FE").pack(anchor="w", pady=(3, 0))

        self.home_version_row = tk.Frame(hero_inner, bg=COLORS['accent'])
        self.home_version_row.pack(anchor="w", pady=(S(10), 0))

        self.home_status_dot = tk.Label(self.home_version_row, text="●",
                                        font=(F, FS(9)), bg=COLORS['accent'], fg="#FEF3C7")
        self.home_status_dot.pack(side=tk.LEFT)
        self.home_status_text = tk.Label(self.home_version_row, text="正在验证版本授权...",
                                         font=(F, FS(9)), bg=COLORS['accent'], fg="#FEF3C7")
        self.home_status_text.pack(side=tk.LEFT, padx=(4, 0))

        cards_row = tk.Frame(body, bg=COLORS['bg'])
        cards_row.pack(fill=tk.X, pady=(0, S(12)))

        self._info_card(cards_row,
                        icon="🛡", title="绕过补丁部署",
                        lines=["一键将 ACLOS UPDATE PATCHER",
                               "部署到系统并创建桌面快捷方式",
                               "无需手动操作，自动完成全流程"])

        self._info_card(cards_row,
                        icon="🎧", title="配音语言修改",
                        lines=["修改 app.asar 注入语言标识",
                               "支持 8 种语言，含中日韩英德等",
                               "自动备份原文件，随时可还原"])

        note = tk.Frame(body, bg=COLORS['surface'],
                        highlightbackground=COLORS['border'], highlightthickness=1)
        note.pack(fill=tk.X)
        note_i = tk.Frame(note, bg=COLORS['surface'])
        note_i.pack(fill=tk.X, padx=S(14), pady=S(10))
        tk.Label(note_i, text="使用说明",
                 font=(F, FS(9), "bold"), bg=COLORS['surface'], fg=COLORS['fg']).pack(anchor="w")
        for line in [
            "• 本工具需要管理员权限运行",
            "• 使用前请确保软件版本已是最新（将自动验证）",
            "• 配音修改仅改变本地文件，不影响游戏账号安全",
            "• 如遇问题请前往 GitHub 提交 Issue，或通过 bilibili 视频反馈"
        ]:
            tk.Label(note_i, text=line, font=(F, FS(8)),
                     bg=COLORS['surface'], fg=COLORS['fg_muted'],
                     justify=tk.LEFT).pack(anchor="w", pady=1)

    def _info_card(self, parent, icon, title, lines):
        F = self.F
        card = tk.Frame(parent, bg=COLORS['surface'],
                        highlightbackground=COLORS['border'], highlightthickness=1)
        card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True,
                  padx=(0, S(8)) if len(parent.winfo_children()) == 1 else (0, 0))

        inner = tk.Frame(card, bg=COLORS['surface'])
        inner.pack(fill=tk.X, padx=S(14), pady=S(12))

        ico_bg = tk.Frame(inner, bg=COLORS['accent_light'], width=S(34), height=S(34))
        ico_bg.pack(anchor="w")
        ico_bg.pack_propagate(False)
        tk.Label(ico_bg, text=icon, font=(F, FS(15)),
                 bg=COLORS['accent_light']).place(relx=.5, rely=.5, anchor="center")

        tk.Label(inner, text=title, font=(F, FS(9), "bold"),
                 bg=COLORS['surface'], fg=COLORS['fg']).pack(anchor="w", pady=(S(6), S(2)))
        for line in lines:
            tk.Label(inner, text=line, font=(F, FS(8)),
                     bg=COLORS['surface'], fg=COLORS['fg_muted']).pack(anchor="w")

    def _build_deploy_page(self, parent):
        F = self.F
        self._page_topbar(parent, "部署绕过补丁")

        body = tk.Frame(parent, bg=COLORS['bg'])
        body.pack(fill=tk.BOTH, expand=True, padx=S(18), pady=S(14))

        self._feature_card(
            body, icon="🛡", title="一键部署绕过补丁",
            desc="将 ACLOS UPDATE PATCHER 部署到系统并自动在桌面创建快捷方式",
            btn_text="立即部署", btn_cmd=self.deploy_patcher,
            accent=COLORS['accent'], accent_light=COLORS['accent_light'])

        self._log_panel(body)

    def _build_vc_page(self, parent):
        F = self.F
        self._page_topbar(parent, "配音修改")
        self._vc_panel = VCEmbedded(parent, self)

    def _page_topbar(self, parent, title):
        F = self.F
        bar = tk.Frame(parent, bg=COLORS['surface'], height=S(50))
        bar.pack(fill=tk.X)
        bar.pack_propagate(False)
        tk.Label(bar, text=title, font=(F, FS(12), "bold"),
                 bg=COLORS['surface'], fg=COLORS['fg']).pack(side=tk.LEFT, padx=S(18), pady=S(12))
        tk.Frame(parent, bg=COLORS['border'], height=1).pack(fill=tk.X)

    def _feature_card(self, parent, icon, title, desc, btn_text, btn_cmd,
                      accent, accent_light):
        F = self.F
        card = tk.Frame(parent, bg=COLORS['surface'],
                        highlightbackground=COLORS['border'], highlightthickness=1)
        card.pack(fill=tk.X, pady=(0, S(8)))
        inner = tk.Frame(card, bg=COLORS['surface'])
        inner.pack(fill=tk.X, padx=S(14), pady=S(10))

        ico = tk.Frame(inner, bg=accent_light, width=S(36), height=S(36))
        ico.pack(side=tk.LEFT)
        ico.pack_propagate(False)
        tk.Label(ico, text=icon, font=(F, FS(16)),
                 bg=accent_light, fg=accent).place(relx=.5, rely=.5, anchor="center")

        tf = tk.Frame(inner, bg=COLORS['surface'])
        tf.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=S(10))
        tk.Label(tf, text=title, font=(F, FS(10), "bold"),
                 bg=COLORS['surface'], fg=COLORS['fg']).pack(anchor="w")
        tk.Label(tf, text=desc, font=(F, FS(8)),
                 bg=COLORS['surface'], fg=COLORS['fg_muted'],
                 wraplength=S(400), justify=tk.LEFT).pack(anchor="w", pady=(2, 0))

        RoundedButton(inner, text=btn_text, command=btn_cmd,
                      bg=accent, fg="white", hover_bg=self._darken(accent),
                      font_size=FS(9), font_weight="bold",
                      width=S(86), height=S(30), radius=S(8)).pack(side=tk.RIGHT)

    def _log_panel(self, parent):
        F = self.F
        log_card = tk.Frame(parent, bg=COLORS['surface'],
                            highlightbackground=COLORS['border'], highlightthickness=1)
        log_card.pack(fill=tk.BOTH, expand=True, pady=(0, 0))

        hdr = tk.Frame(log_card, bg=COLORS['surface'])
        hdr.pack(fill=tk.X, padx=S(12), pady=(S(8), 0))
        tk.Label(hdr, text="运行日志", font=(F, FS(9), "bold"),
                 bg=COLORS['surface'], fg=COLORS['fg']).pack(side=tk.LEFT)
        RoundedButton(hdr, text="清空", command=self.clear_log,
                      bg=COLORS['bg_alt'], fg=COLORS['fg_muted'], hover_bg=COLORS['border'],
                      font_size=FS(8), width=S(48), height=S(20), radius=S(4)).pack(side=tk.RIGHT)

        self.log_text = scrolledtext.ScrolledText(
            log_card, font=("Consolas", FS(9)),
            bg=COLORS['bg'], fg=COLORS['fg'],
            insertbackground=COLORS['fg'],
            relief=tk.FLAT, state=tk.DISABLED, padx=8, pady=4)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=S(10), pady=(S(6), S(10)))

    def _check_version_startup(self):
        self.set_status("正在验证版本...", COLORS['warning'])
        self.log(f"本地版本: {CURRENT_VERSION}")
        self.log(f"正在连接验证服务器...")

        def check():
            try:
                self.root.after(0, lambda: self.log("正在与服务器建立连接..."))
                headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
                res = requests.get(VERSION_URL, headers=headers, timeout=8)
                self.root.after(0, lambda: self.log(f"服务器响应: HTTP {res.status_code}"))

                if res.status_code != 200:
                    raise Exception(f"服务器返回异常状态码: {res.status_code}")

                data = res.json()
                version_pool = data.get("versions", [])
                self.root.after(0, lambda: self.log(
                    f"版本池已加载，共 {len(version_pool)} 条记录"))

                current_info = next(
                    (i for i in version_pool if i["v"] == CURRENT_VERSION), None)
                latest_info = (max(version_pool, key=lambda x: x["priority"])
                               if version_pool else None)

                if not current_info:
                    self.root.after(0, lambda: self.log(
                        f"版本 {CURRENT_VERSION} 不在云端列表，正在与云端版本比较...", "warning"))

                    def _parse_version(v_str):
                        """将版本字符串解析为可比较的整数元组"""
                        try:
                            return tuple(int(x) for x in str(v_str).split("."))
                        except Exception:
                            return (0,)

                    if latest_info:
                        local_tuple  = _parse_version(CURRENT_VERSION)
                        latest_tuple = _parse_version(latest_info.get("v", "0"))

                        if local_tuple < latest_tuple:
                            # 本地版本低于云端最新版本且不在列表中 无论 force_update 字段 一律强制更新
                            lv = latest_info.get("v", "?")
                            self.root.after(0, lambda: self.log(
                                f"本地版本不在列表且低于云端 {lv}，强制要求更新", "warning"))
                            _li = latest_info
                            self.root.after(0, lambda li=_li: self._on_version_rejected(
                                li, "force"))
                        else:
                            # 本地版本 >= 云端版本（预览版/开发版）
                            self.root.after(0, lambda: self.log(
                                f"版本 {CURRENT_VERSION} 不低于云端，视为有效版本", "info"))
                            self.root.after(0, lambda: self._on_version_ok())
                    else:
                        # 无法获取云端版本信息
                        self.root.after(0, lambda: self.log(
                            f"版本 {CURRENT_VERSION} 不在云端列表，且无法获取云端版本，拒绝访问", "error"))
                        self.root.after(0, lambda: self._on_version_rejected(
                            None, "unauthorized"))
                    return

                if latest_info and current_info["priority"] < latest_info["priority"]:
                    lv = latest_info.get("v", "?")
                    is_force = latest_info.get("force_update", False)
                    flag = "【强制】" if is_force else "【可选】"
                    self.root.after(0, lambda: self.log(
                        f"发现新版本: {lv} {flag}", "warning"))
                    if is_force:
                        _li = latest_info
                        self.root.after(0, lambda li=_li: self._on_version_rejected(
                            li, "force"))
                        return
                    else:
                        _li = latest_info
                        self.root.after(0, lambda: self._on_version_ok())
                        self.root.after(300, lambda li=_li: self.show_update_dialog(
                            li, False))
                else:
                    self.root.after(0, lambda: self._on_version_ok())

            except requests.exceptions.ConnectionError as e:
                self.root.after(0, lambda: self.log(f"网络连接失败: {e}", "error"))
                self.root.after(0, lambda: self._on_version_network_fail())
            except requests.exceptions.Timeout:
                self.root.after(0, lambda: self.log("请求超时: 服务器 8 秒内未响应", "error"))
                self.root.after(0, lambda: self._on_version_network_fail())
            except Exception as e:
                self.root.after(0, lambda: self.log(f"版本检查异常: {e}", "error"))
                self.root.after(0, lambda: self._on_version_network_fail())

        threading.Thread(target=check, daemon=True).start()

    def _on_version_ok(self):
        self._version_ok = True
        self._unlock_nav()
        self.set_status(f"系统就绪 · v{CURRENT_VERSION}", COLORS['success'])
        self.log("版本云端验证通过 ✓", "success")
        self._update_home_status(True, f"版本 {CURRENT_VERSION} 已验证 ✓")
        self._show_home_ok_card()

    def _on_version_rejected(self, latest_info, reason):
        self._version_ok = False
        if reason == "unauthorized":
            title = "版本未验证"
            body  = (f"当前版本 {CURRENT_VERSION} 不在云端列表中。\n"
                     "请前往 GitHub 下载最新版本后重新使用。")
            status_text = f"版本 {CURRENT_VERSION} 未验证，功能已锁定"
        else:
            title = "需要强制更新"
            lv = latest_info.get("v", "?") if latest_info else "?"
            body  = (f"检测到强制更新版本 {lv}，当前版本已停止支持。\n"
                     "请更新后重新使用。")
            status_text = f"需要更新至 {lv}"

        self.set_status(status_text, COLORS['error'])
        self.log(f"版本验证失败: {title}", "error")
        self._update_home_status(False, f"{title} — 功能已锁定")
        self._show_page("home")
        self._show_home_block_card(title, body, latest_info)

        if reason == "force" and latest_info:
            self.root.after(200, lambda: self.show_update_dialog(latest_info, True))

    def _on_version_network_fail(self):
        self.set_status("版本验证失败 (网络错误)", COLORS['error'])
        self._update_home_status(False, "无法连接验证服务器")
        self._show_page("home")
        self._show_home_block_card(
            "无法验证版本",
            ("无法连接版本验证服务器。\n"
             "为保证安全，功能模块已锁定。\n"
             "请检查网络连接后重试，或退出程序。"),
            None,
            show_retry=True)

    def _update_home_status(self, ok: bool, text: str):
        color = COLORS['success_light'] if ok else COLORS['warning_light']
        dot   = COLORS['success'] if ok else COLORS['error']
        try:
            self.home_status_dot.config(fg=dot)
            self.home_status_text.config(text=text, fg=color)
        except Exception:
            pass

    def _show_home_ok_card(self):
        F = self.F
        home_frame = self._nav_frames["home"]
        for w in home_frame.winfo_children():
            if getattr(w, '_is_block_card', False):
                w.destroy()

        card = tk.Frame(home_frame, bg=COLORS['success_light'],
                        highlightbackground=COLORS['success'], highlightthickness=1)
        card._is_block_card = True
        card.pack(fill=tk.X, padx=S(20), pady=(0, S(10)), side=tk.BOTTOM)

        inner = tk.Frame(card, bg=COLORS['success_light'])
        inner.pack(fill=tk.X, padx=S(14), pady=S(10))

        left = tk.Frame(inner, bg=COLORS['success_light'])
        left.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(left, text=f"✓  当前已是最新版本",
                 font=(F, FS(10), "bold"),
                 bg=COLORS['success_light'], fg=COLORS['success']).pack(anchor="w")

    def _show_home_block_card(self, title, body, latest_info, show_retry=False):
        F = self.F
        home_frame = self._nav_frames["home"]
        for w in home_frame.winfo_children():
            if getattr(w, '_is_block_card', False):
                w.destroy()

        card = tk.Frame(home_frame, bg=COLORS['error_light'],
                        highlightbackground=COLORS['error'], highlightthickness=1)
        card._is_block_card = True
        card.pack(fill=tk.X, padx=S(20), pady=(0, S(10)), side=tk.BOTTOM)

        inner = tk.Frame(card, bg=COLORS['error_light'])
        inner.pack(fill=tk.X, padx=S(14), pady=S(10))

        tk.Label(inner, text=f"🔒  {title}",
                 font=(F, FS(10), "bold"), bg=COLORS['error_light'],
                 fg=COLORS['error']).pack(anchor="w")
        tk.Label(inner, text=body, font=(F, FS(8)),
                 bg=COLORS['error_light'], fg=COLORS['fg'],
                 justify=tk.LEFT).pack(anchor="w", pady=(S(4), S(8)))

        btn_row = tk.Frame(inner, bg=COLORS['error_light'])
        btn_row.pack(anchor="w")

        def open_release():
            webbrowser.open("https://github.com/ab1f5a/ACLOS-Tools/releases/")

        RoundedButton(btn_row, text="前往下载", command=open_release,
                      bg=COLORS['accent'], fg="white", hover_bg=COLORS['accent_hover'],
                      font_size=FS(8), font_weight="bold",
                      width=S(80), height=S(26), radius=S(6)).pack(side=tk.LEFT, padx=(0, S(6)))

        if show_retry:
            RoundedButton(btn_row, text="重试",
                          command=lambda: [card.destroy(), self._check_version_startup()],
                          bg=COLORS['warning_light'], fg=COLORS['warning'],
                          hover_bg=COLORS['warning'],
                          font_size=FS(8), width=S(60), height=S(26), radius=S(6)).pack(side=tk.LEFT, padx=(0, S(6)))

        RoundedButton(btn_row, text="退出程序", command=self.root.quit,
                      bg=COLORS['error'], fg="white", hover_bg=self._darken(COLORS['error']),
                      font_size=FS(8), width=S(80), height=S(26), radius=S(6)).pack(side=tk.LEFT)

    # 更新弹窗

    def show_update_dialog(self, latest_info, is_force):
        F = self.F
        remote_v   = latest_info.get("v", "Unknown") if latest_info else "未知"
        update_log = latest_info.get("log", "暂无更新说明") if latest_info else "暂无说明"
        force_tag  = "强制更新" if is_force else "可选更新"
        force_color = COLORS['error'] if is_force else COLORS['warning']
        accent      = COLORS['error'] if is_force else COLORS['accent']
        title_text  = "需要强制更新" if is_force else "发现新版本"

        W, H = S(480), S(435)
        d = tk.Toplevel(self.root)
        d.configure(bg=COLORS['bg'])
        d.resizable(False, False)
        d.overrideredirect(True)
        d.update_idletasks()
        x = (d.winfo_screenwidth()  - W) // 2
        y = (d.winfo_screenheight() - H) // 2
        d.geometry(f'{W}x{H}+{x}+{y}')
        d.transient(self.root)
        d.update()
        d.grab_set()
        d.focus_force()
        dbar = tk.Frame(d, bg=COLORS['titlebar'], height=S(30))
        dbar.pack(fill=tk.X)
        dbar.pack_propagate(False)

        tk.Label(dbar, text="  🔔  版本更新提醒",
                 font=(F, FS(8)), bg=COLORS['titlebar'], fg="#8B92B8").pack(side=tk.LEFT)
        dclose = tk.Label(dbar, text="  ✕  ", font=(F, FS(9)),
                          bg=COLORS['titlebar'], fg="#8B92B8", cursor="hand2")
        dclose.pack(side=tk.RIGHT, fill=tk.Y)
        if is_force:
            dclose.bind("<Enter>",    lambda e: dclose.config(bg=COLORS['error'], fg="white"))
            dclose.bind("<Leave>",    lambda e: dclose.config(bg=COLORS['titlebar'], fg="#8B92B8"))
            dclose.bind("<Button-1>", lambda e: self.root.quit())
        else:
            dclose.bind("<Enter>",    lambda e: dclose.config(bg=COLORS['error'], fg="white"))
            dclose.bind("<Leave>",    lambda e: dclose.config(bg=COLORS['titlebar'], fg="#8B92B8"))
            dclose.bind("<Button-1>", lambda e: d.destroy())

        _dx, _dy = [0], [0]
        def _dbar_press(e):
            _dx[0] = e.x_root - d.winfo_x()
            _dy[0] = e.y_root - d.winfo_y()
        def _dbar_move(e):
            d.geometry(f"+{e.x_root - _dx[0]}+{e.y_root - _dy[0]}")
        dbar.bind("<ButtonPress-1>", _dbar_press)
        dbar.bind("<B1-Motion>",     _dbar_move)

        tk.Frame(d, bg=accent, height=S(4)).pack(fill=tk.X)

        ct = tk.Frame(d, bg=COLORS['bg'])
        ct.pack(fill=tk.BOTH, expand=True, padx=S(22), pady=S(14))

        title_row = tk.Frame(ct, bg=COLORS['bg'])
        title_row.pack(fill=tk.X, pady=(0, S(3)))

        badge_bg = COLORS['error_light'] if is_force else COLORS['warning_light']
        tk.Label(title_row, text=f"  {force_tag}  ",
                 font=(F, FS(8), "bold"), bg=badge_bg, fg=force_color).pack(side=tk.RIGHT, pady=S(2))
        tk.Label(title_row, text=title_text,
                 font=(F, FS(13), "bold"), bg=COLORS['bg'], fg=COLORS['fg']).pack(side=tk.LEFT, anchor="w")

        tk.Label(ct, text="检测到新版本，请查看详情后决定是否更新",
                 font=(F, FS(8)), bg=COLORS['bg'], fg=COLORS['fg_muted']).pack(anchor="w", pady=(0, S(10)))

        info_card = tk.Frame(ct, bg=COLORS['surface'],
                             highlightbackground=COLORS['border'], highlightthickness=1)
        info_card.pack(fill=tk.X, pady=(0, S(10)))

        rows_data = [
            ("本地版本", CURRENT_VERSION, COLORS['fg_muted']),
            ("云端版本", remote_v,        COLORS['success']),
            ("更新类型", force_tag,       force_color),
        ]
        for i, (label, value, vc) in enumerate(rows_data):
            row = tk.Frame(info_card, bg=COLORS['surface'])
            row.pack(fill=tk.X, padx=S(14), pady=S(5))
            if i < len(rows_data) - 1:
                tk.Frame(info_card, bg=COLORS['border'], height=1).pack(fill=tk.X, padx=S(14))

            tk.Label(row, text=label, font=(F, FS(9)),
                     bg=COLORS['surface'], fg=COLORS['fg_muted']).pack(side=tk.LEFT)

            if label == "本地版本":
                arrow_row = tk.Frame(row, bg=COLORS['surface'])
                arrow_row.pack(side=tk.RIGHT)
                tk.Label(arrow_row, text=CURRENT_VERSION,
                         font=(F, FS(9)), bg=COLORS['surface'],
                         fg=COLORS['fg_muted']).pack(side=tk.LEFT)
                tk.Label(arrow_row, text="  →  ",
                         font=(F, FS(9)), bg=COLORS['surface'],
                         fg=COLORS['text_light']).pack(side=tk.LEFT)
                tk.Label(arrow_row, text=remote_v,
                         font=(F, FS(9), "bold"), bg=COLORS['surface'],
                         fg=COLORS['success']).pack(side=tk.LEFT)
            else:
                tk.Label(row, text=value, font=(F, FS(9), "bold"),
                         bg=COLORS['surface'], fg=vc).pack(side=tk.RIGHT)

        tk.Label(ct, text="更新日志", font=(F, FS(9), "bold"),
                 bg=COLORS['bg'], fg=COLORS['fg']).pack(anchor="w", pady=(0, S(4)))

        log_frame = tk.Frame(ct, bg=COLORS['surface'],
                             highlightbackground=COLORS['border'], highlightthickness=1)
        log_frame.pack(fill=tk.X)

        tb = tk.Text(log_frame, font=(F, FS(8)), bg=COLORS['surface'], fg=COLORS['fg'],
                     relief=tk.FLAT, wrap=tk.WORD, state=tk.DISABLED,
                     height=6, padx=S(10), pady=S(8))
        tb_scroll = ttk.Scrollbar(log_frame, command=tb.yview)
        tb.configure(yscrollcommand=tb_scroll.set)
        tb_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        tb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tb.configure(state=tk.NORMAL)
        tb.insert(tk.END, update_log)
        tb.configure(state=tk.DISABLED)

        if is_force:
            warn = tk.Frame(ct, bg=COLORS['error_light'],
                            highlightbackground=COLORS['error'], highlightthickness=1)
            warn.pack(fill=tk.X, pady=(S(8), 0))
            tk.Label(warn, text="⚠  当前版本已停止支持，关闭此窗口将退出程序",
                     font=(F, FS(8), "bold"), bg=COLORS['error_light'],
                     fg=COLORS['error'], pady=S(5)).pack()

        bf = tk.Frame(ct, bg=COLORS['bg'])
        bf.pack(fill=tk.X, pady=(S(12), 0))

        RoundedButton(bf, text="前往下载",
                      command=lambda: webbrowser.open(
                          "https://github.com/ab1f5a/ACLOS-Tools/releases/"),
                      bg=COLORS['accent'], fg="white", hover_bg=COLORS['accent_hover'],
                      font_size=FS(9), font_weight="bold",
                      width=S(90), height=S(30), radius=S(8)).pack(side=tk.LEFT, padx=(0, S(8)))

        if is_force:
            RoundedButton(bf, text="退出程序",
                          command=self.root.quit,
                          bg=COLORS['error_light'],
                          fg=COLORS['error'],
                          hover_bg=self._darken_color(COLORS['error_light'], steps=30),
                          font_size=FS(9), width=S(90), height=S(30), radius=S(8)).pack(side=tk.LEFT)
        else:
            RoundedButton(bf, text="稍后更新",
                          command=d.destroy,
                          bg=COLORS['bg_alt'], fg=COLORS['fg'], hover_bg=COLORS['border'],
                          font_size=FS(9), width=S(90), height=S(30), radius=S(8)).pack(side=tk.LEFT)

    def deploy_patcher(self):
        self.log("开始部署绕过补丁...")
        self.set_status("正在部署...", COLORS['info'])

        def deploy():
            pythoncom.CoInitialize()
            try:
                os.makedirs(DEPLOY_DIR, exist_ok=True)
                self.root.after(0, lambda: self.log(f"部署目录: {DEPLOY_DIR}"))

                # --- PyInstaller 专用路径兼容逻辑 ---
                if hasattr(sys, '_MEIPASS'):
                    # 打包后的环境：资源在临时文件夹 _MEIPASS 中
                    base_dir = sys._MEIPASS
                else:
                    # 未打包的开发环境：资源在脚本同级目录
                    base_dir = os.path.dirname(os.path.abspath(__file__))

                src = os.path.join(base_dir, PATCHER_EXE_NAME)
                dst = os.path.join(DEPLOY_DIR, PATCHER_EXE_NAME)

                if not os.path.exists(src):
                    self.root.after(0, lambda: self.log(f"找不到源文件: {src}", "error"))
                    self.root.after(0, lambda: self.set_status("部署失败", COLORS['error']))
                    return

                shutil.copy2(src, dst)
                self.root.after(0, lambda: self.log(f"文件已复制: {dst}", "success"))

                desktop = get_unicode_shell_path(0x0000)
                lnk = os.path.join(desktop, "ACLOS.lnk")
                if os.path.exists(lnk):
                    try:
                        os.chmod(lnk, 0o777)
                        os.remove(lnk)
                    except Exception:
                        pass

                shell = Dispatch('WScript.Shell')
                sc = shell.CreateShortCut(str(lnk))
                sc.TargetPath = str(dst)
                sc.WorkingDirectory = str(DEPLOY_DIR)
                sc.IconLocation = str(dst)
                sc.save()

                self.root.after(0, lambda: self.log("桌面快捷方式创建成功", "success"))
                self.root.after(0, lambda: self.set_status("部署完成", COLORS['success']))
                self.root.after(0, lambda: messagebox.showinfo(
                    "部署成功", "绕过补丁已成功部署到桌面！"))

            except Exception as e:
                self.root.after(0, lambda: self.log(f"部署失败: {e}", "error"))
                self.root.after(0, lambda: self.set_status("部署失败", COLORS['error']))
                self.root.after(0, lambda: messagebox.showerror("部署失败", str(e)))
            
            finally:
                # --- 新增：释放 COM 资源 ---
                pythoncom.CoUninitialize()
        threading.Thread(target=deploy, daemon=True).start()

    def open_vc_module(self):
        self._show_vc()

    def log(self, message, level="info"):
        colors = {"info": COLORS['info'], "success": COLORS['success'],
                  "warning": COLORS['warning'], "error": COLORS['error']}
        try:
            self.log_text.configure(state=tk.NORMAL)
            ts = time.strftime("%H:%M:%S")
            self.log_text.insert(tk.END, f"[{ts}] ", "ts")
            self.log_text.insert(tk.END, f"{message}\n", level)
            self.log_text.tag_config("ts", foreground=COLORS['text_light'])
            self.log_text.tag_config(level, foreground=colors.get(level, COLORS['fg']))
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)
        except Exception:
            pass

    def clear_log(self):
        try:
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.configure(state=tk.DISABLED)
        except Exception:
            pass

    def set_status(self, text, color=None):
        try:
            self.status_label.config(
                text=f"● {text}",
                fg=color or COLORS['info'],
                bg=COLORS['titlebar'])
        except Exception:
            pass

    def _darken(self, hex_color):
        try:
            r = max(0, int(hex_color[1:3], 16) - 25)
            g = max(0, int(hex_color[3:5], 16) - 25)
            b = max(0, int(hex_color[5:7], 16) - 25)
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception:
            return hex_color

    def _darken_color(self, hex_color, steps=25):
        try:
            r = max(0, int(hex_color[1:3], 16) - steps)
            g = max(0, int(hex_color[3:5], 16) - steps)
            b = max(0, int(hex_color[5:7], 16) - steps)
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception:
            return hex_color

class VCEmbedded:
    """配音修改模块"""

    FONT = "Microsoft YaHei UI"

    def __init__(self, parent, main_app):
        self.main_app = main_app
        self.parent   = parent

        outer = tk.Frame(parent, bg=COLORS['bg'])
        outer.pack(fill=tk.BOTH, expand=True, padx=S(18), pady=S(12))
        self._outer = outer
        self._build(outer)
        self.auto_find()

    def _build(self, outer):
        F = self.FONT

        pc = tk.Frame(outer, bg=COLORS['surface'],
                      highlightbackground=COLORS['border'], highlightthickness=1)
        pc.pack(fill=tk.X, pady=(0, S(8)))
        pi = tk.Frame(pc, bg=COLORS['surface'])
        pi.pack(fill=tk.X, padx=S(12), pady=S(8))

        tk.Label(pi, text="目标文件 (app.asar)", font=(F, FS(9), "bold"),
                 bg=COLORS['surface'], fg=COLORS['fg']).pack(anchor="w")

        self.path_var = tk.StringVar(value="未选择")
        pd = tk.Frame(pi, bg=COLORS['bg_alt'],
                      highlightbackground=COLORS['border'], highlightthickness=1)
        pd.pack(fill=tk.X, pady=(S(4), S(6)))
        tk.Label(pd, textvariable=self.path_var, font=("Consolas", FS(8)),
                 bg=COLORS['bg_alt'], fg=COLORS['info'],
                 anchor="w", padx=S(8), pady=S(5)).pack(fill=tk.X)

        br = tk.Frame(pi, bg=COLORS['surface'])
        br.pack(anchor="w")
        RoundedButton(br, text="自动检索", command=self.auto_find,
                      bg=COLORS['accent'], fg="white", hover_bg=COLORS['accent_hover'],
                      font_size=FS(8), width=S(78), height=S(26), radius=S(6)).pack(side=tk.LEFT, padx=(0, S(6)))
        RoundedButton(br, text="手动选择", command=self.manual_select,
                      bg=COLORS['bg_alt'], fg=COLORS['fg'], hover_bg=COLORS['border'],
                      font_size=FS(8), width=S(78), height=S(26), radius=S(6)).pack(side=tk.LEFT)

        lc = tk.Frame(outer, bg=COLORS['surface'],
                      highlightbackground=COLORS['border'], highlightthickness=1)
        lc.pack(fill=tk.X, pady=(0, S(8)))
        li = tk.Frame(lc, bg=COLORS['surface'])
        li.pack(fill=tk.X, padx=S(12), pady=S(8))

        tk.Label(li, text="目标语言", font=(F, FS(9), "bold"),
                 bg=COLORS['surface'], fg=COLORS['fg']).pack(anchor="w", pady=(0, S(4)))

        gf = tk.Frame(li, bg=COLORS['surface'])
        gf.pack(fill=tk.X)
        self.lang_var = tk.StringVar(value="1")
        for i, (key, (code, name)) in enumerate(LANG_MAP.items()):
            self._lang_radio(gf, key, code, name, i // 4, i % 4)

        log_card = tk.Frame(outer, bg=COLORS['surface'],
                            highlightbackground=COLORS['border'], highlightthickness=1)
        log_card.pack(fill=tk.BOTH, expand=True, pady=(0, S(8)))

        tk.Label(log_card, text="操作日志", font=(F, FS(9), "bold"),
                 bg=COLORS['surface'], fg=COLORS['fg']).pack(anchor="w", padx=S(12), pady=(S(6), 0))

        self.vc_log = scrolledtext.ScrolledText(
            log_card, font=("Consolas", FS(8)),
            bg=COLORS['bg'], fg=COLORS['fg'],
            relief=tk.FLAT, height=4, state=tk.DISABLED, padx=6, pady=3)
        self.vc_log.pack(fill=tk.BOTH, expand=True, padx=S(10), pady=(S(4), S(8)))

        self.feedback_card = tk.Frame(outer, bg=COLORS['warning_light'],
                                      highlightbackground=COLORS['warning'],
                                      highlightthickness=1)

        self.action_row = tk.Frame(outer, bg=COLORS['bg'])
        self.action_row.pack(fill=tk.X, pady=(0, S(2)))

        RoundedButton(self.action_row, text="开始修改", command=self.start_patch,
                      bg=COLORS['success'], fg="white", hover_bg=COLORS['success_hover'],
                      font_size=FS(9), font_weight="bold",
                      width=S(100), height=S(32), radius=S(8)).pack(side=tk.LEFT, padx=(0, S(8)))

    def _show_feedback(self, bg_key, border_key, icon, title, title_color, body, buttons):
        F = self.FONT
        for w in self.feedback_card.winfo_children():
            w.destroy()
        bg = COLORS[bg_key]
        self.feedback_card.configure(bg=bg, highlightbackground=COLORS[border_key])

        inner = tk.Frame(self.feedback_card, bg=bg)
        inner.pack(fill=tk.X, padx=S(12), pady=S(8))

        tk.Label(inner, text=f"{icon}  {title}", font=(F, FS(9), "bold"),
                 bg=bg, fg=title_color).pack(anchor="w")
        if body:
            tk.Label(inner, text=body, font=(F, FS(8)),
                     bg=bg, fg=COLORS['fg'], justify=tk.LEFT,
                     wraplength=S(500)).pack(anchor="w", pady=(S(4), S(6)))

        btn_row = tk.Frame(inner, bg=bg)
        btn_row.pack(anchor="w")
        for (lbl, cmd, bb, bf, bh) in buttons:
            RoundedButton(btn_row, text=lbl, command=cmd,
                          bg=bb, fg=bf, hover_bg=bh,
                          font_size=FS(8), width=S(80), height=S(26), radius=S(6)
                          ).pack(side=tk.LEFT, padx=(0, S(6)))

        self.feedback_card.pack(fill=tk.X, pady=(0, S(6)), before=self.action_row)

    def _hide_feedback(self):
        self.feedback_card.pack_forget()

    def _show_confirm_card(self, file_path, target_lang, lang_name):
        body = (f"目标文件：{os.path.basename(file_path)}\n"
                f"注入语言：{target_lang}（{lang_name}）\n"
                f"原文件将自动备份为 .bak，可随时还原。")

        def do_confirm():
            self._hide_feedback()
            self._do_patch(file_path, target_lang)

        self._show_feedback(
            'warning_light', 'warning', "⚠", "确认操作", COLORS['warning'], body,
            [("确认修改", do_confirm,
              COLORS['success'], "white", COLORS['success_hover']),
             ("取消", self._hide_feedback,
              COLORS['bg_alt'], COLORS['fg'], COLORS['border'])])

    def _show_result_card(self, ok: bool, msg: str):
        if ok:
            self._show_feedback(
                'success_light', 'success', "✓", "操作成功", COLORS['success'], msg,
                [("关闭", self._hide_feedback,
                  COLORS['bg_alt'], COLORS['fg'], COLORS['border'])])
        else:
            self._show_feedback(
                'error_light', 'error', "✗", "操作失败", COLORS['error'], msg,
                [("关闭", self._hide_feedback,
                  COLORS['bg_alt'], COLORS['fg'], COLORS['border'])])

    def _lang_radio(self, parent, key, code, name, row, col):
        F = self.FONT
        frame = tk.Frame(parent, bg=COLORS['surface'], cursor="hand2")
        frame.grid(row=row, column=col, sticky="w", padx=S(4), pady=S(1))
        rb = tk.Radiobutton(frame, text=f"{code}  {name}",
                            variable=self.lang_var, value=key,
                            font=(F, FS(8)), bg=COLORS['surface'], fg=COLORS['fg'],
                            selectcolor=COLORS['accent_light'],
                            activebackground=COLORS['surface_hover'],
                            activeforeground=COLORS['accent'],
                            indicatoron=True, cursor="hand2")
        rb.pack()

    def vc_log_msg(self, message, level="info"):
        colors = {"info": COLORS['info'], "success": COLORS['success'],
                  "warning": COLORS['warning'], "error": COLORS['error']}
        self.vc_log.configure(state=tk.NORMAL)
        ts = time.strftime("%H:%M:%S")
        self.vc_log.insert(tk.END, f"[{ts}] ", "ts")
        self.vc_log.insert(tk.END, f"{message}\n", level)
        self.vc_log.tag_config("ts", foreground=COLORS['text_light'])
        self.vc_log.tag_config(level, foreground=colors.get(level, COLORS['fg']))
        self.vc_log.see(tk.END)
        self.vc_log.configure(state=tk.DISABLED)
        self.main_app.log(f"[VC] {message}", level)

    def auto_find(self):
        self.vc_log_msg("正在自动检索 ACLOS 根目录...")

        def find():
            if os.path.exists(PATH_CACHE):
                try:
                    with open(PATH_CACHE, 'r', encoding='utf-8') as f:
                        p = f.read().strip()
                    if os.path.exists(p):
                        self.parent.after(0, lambda: self.path_var.set(p))
                        self.parent.after(0, lambda: self.vc_log_msg(f"从缓存加载: {p}", "success"))
                        return
                except Exception:
                    pass

            drives = [f"{d}:\\" for d in string.ascii_uppercase if os.path.exists(f"{d}:\\")]
            for drive in drives:
                targets = [os.path.join(drive, ROOT_DIR_NAME)]
                try:
                    for item in os.listdir(drive):
                        if os.path.isdir(os.path.join(drive, item)):
                            targets.append(os.path.join(drive, item, ROOT_DIR_NAME))
                except Exception:
                    continue
                for rp in targets:
                    fp = os.path.join(rp, SUB_PATH)
                    if os.path.exists(fp):
                        self.parent.after(0, lambda p=fp: self.path_var.set(p))
                        self.parent.after(0, lambda p=fp: self.vc_log_msg(f"定位成功: {p}", "success"))
                        try:
                            with open(PATH_CACHE, 'w', encoding='utf-8') as f:
                                f.write(fp)
                        except Exception:
                            pass
                        return
            self.parent.after(0, lambda: self.vc_log_msg("自动检索失败，请手动选择", "warning"))

        threading.Thread(target=find, daemon=True).start()

    def manual_select(self):
        path = filedialog.askopenfilename(
            title="选择 app.asar 文件",
            filetypes=[("ASAR File", "app.asar"), ("All Files", "*.*")])
        if path:
            self.path_var.set(path)
            self.vc_log_msg(f"手动选择: {path}", "success")
            try:
                with open(PATH_CACHE, 'w', encoding='utf-8') as f:
                    f.write(path)
            except Exception:
                pass

    def start_patch(self):
        file_path = self.path_var.get()
        if file_path == "未选择" or not os.path.exists(file_path):
            self.vc_log_msg("请先选择有效的 app.asar 文件", "error")
            return
        target_lang, lang_name = LANG_MAP[self.lang_var.get()]
        self._show_confirm_card(file_path, target_lang, lang_name)

    def _do_patch(self, file_path, target_lang):
        self.vc_log_msg(f"开始注入语言: {target_lang}")

        def patch():
            bak = file_path + ".bak"
            src = bak if os.path.exists(bak) else file_path
            msg = "发现备份，从镜像还原注入" if src == bak else "未发现备份，正在创建备份..."
            lvl = "info" if src == bak else "warning"
            self.parent.after(0, lambda: self.vc_log_msg(msg, lvl))
            try:
                with open(src, 'rb') as f:
                    raw = f.read()
                if src == file_path:
                    shutil.copy2(file_path, bak)
                    self.parent.after(0, lambda: self.vc_log_msg("备份已就绪", "success"))

                if b"zh_CN" not in raw:
                    self.parent.after(0, lambda: self.vc_log_msg("校验失败: zh_CN 标识缺失", "error"))
                    self.parent.after(0, lambda: self._show_result_card(
                        False, "文件中未找到 zh_CN 标识\n可能已被修改或文件无效"))
                    return

                new_raw = raw.replace(b"zh_CN", target_lang.encode())
                with open(file_path, 'wb') as f:
                    f.write(new_raw)

                fname = os.path.basename(file_path)
                self.parent.after(0, lambda: self.vc_log_msg(f"写入完成: {fname}", "success"))
                self.parent.after(0, lambda: self._show_result_card(
                    True, f"语言已切换为 {target_lang}\n文件: {fname}\n请重启启动器以生效"))

            except Exception as e:
                self.parent.after(0, lambda: self.vc_log_msg(f"异常: {e}", "error"))
                self.parent.after(0, lambda: self._show_result_card(False, f"运行时异常:\n{e}"))

        threading.Thread(target=patch, daemon=True).start()

def main():
    if not ctypes.windll.shell32.IsUserAnAdmin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable,
                                             " ".join(sys.argv), None, 1)
        return

    root = tk.Tk()
    app = ACLOSToolsGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()