# gui/app.py -- main window
import os
import sys
import ctypes
import time
import threading
import queue
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config
from models import QuoteItem
from utils import fmt_money, save_hist, show_err, read_json
from engines.competitor import comp_items
from engines.library_manager import (
    get_party_b,
    list_party_a_names,
    list_party_b_names,
    load_contacts,
    load_private_customers,
    resolve_party_a,
    save_contact_entry,
)
from gui.autocomplete import AutocompleteInput

class ToolTip:
    """tkinter 按钮浮动提示（Tooltip）"""
    def __init__(self, widget, get_text):
        # get_text: 无参可调用对象，返回当前提示文本
        self.widget = widget
        self.get_text = get_text
        self.tipwindow = None
        widget.bind('<Enter>', self._show)
        widget.bind('<Leave>', self._hide)

    def _show(self, _event=None):
        if self.tipwindow:
            return
        text = self.get_text()
        if not text:
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(
            tw, text=text,
            background="#ffffe0", foreground="#333",
            relief="solid", borderwidth=1,
            font=("微软雅黑", 9), padx=6, pady=3
        ).pack()

    def _hide(self, _event=None):
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


BG = "#f6f7f9"
CARD = "#ffffff"
BORDER = "#dfe3e8"
TEXT = "#202124"
MUTED = "#6b7280"
ACCENT = "#10a37f"
ACCENT_DARK = "#0e8f70"
GRID_HEADER = "#F5F7FA"
GRID_ALT = "#FAFBFC"
GRID_LINE = "#E5E7EB"
GRID_HOVER = "#EAF3FF"
GRID_SELECTED = "#DCEBFF"
LINK = "#1677FF"
DANGER = "#FF4D4F"
DANGER_HOVER = "#FFF1F0"
EMPTY_TEXT = "#BFBFBF"
WINDOW_DEFAULT_WIDTH = 1500
WINDOW_DEFAULT_HEIGHT = 900
WINDOW_MIN_WIDTH = 1280
WINDOW_MIN_HEIGHT = 760
RESOURCE_TREE_MIN_BODY = 178
QUOTE_TREE_MIN_BODY = 150


class App(tk.Tk):
    def __init__(self):
        startup_start = time.perf_counter()
        super().__init__()
        self._log_startup("create_tk_root", startup_start)
        build_start = time.perf_counter()
        self.title("报价合同生成工具 Prod. BY LanZihan")
        scr_w = self.winfo_screenwidth()
        scr_h = self.winfo_screenheight()
        init_w = min(WINDOW_DEFAULT_WIDTH, max(WINDOW_MIN_WIDTH, scr_w - 80))
        init_h = min(WINDOW_DEFAULT_HEIGHT, max(WINDOW_MIN_HEIGHT, scr_h - 100))
        self.geometry(f"{init_w}x{init_h}")
        self.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)
        self.configure(bg=BG)
        try:
            self.tk.call("tk", "scaling", self.winfo_fpixels("1i") / 72)
        except tk.TclError:
            pass

        self.resources = []
        self.filtered_resources = []
        self.storage_resources = []
        self.storage_by_name = {}
        self.comp_items = []
        self.stor_items = []
        self.contract_info = {}
        self.generated_files = []
        self.unit_suggestions = []
        self.loading_prices = False
        self.price_load_error = None
        self.price_load_queue = queue.Queue()
        self._party_a_options_cache = None
        self._party_b_names_cache = None
        self._contacts_cache = None

        self._setup_style()
        self._build()
        # 强制立即渲染窗口，避免"未响应"白屏
        self.update_idletasks()
        self._log_private_contact_status()
        self._log_startup("create_main_window", build_start)
        self.after_idle(self._fit_to_screen)
        self.after_idle(self._show_startup_ready)
        self.after_idle(self._lazy_load_data)
        self.after(100, self.reload_prices)

    def _log_startup(self, label, start_time):
        elapsed = (time.perf_counter() - start_time) * 1000
        self._log_startup_elapsed(label, elapsed)

    def _log_startup_elapsed(self, label, elapsed):
        print(f"[startup] {label}: {elapsed:.1f} ms", flush=True)

    def _show_startup_ready(self):
        self.update_idletasks()
        print("[startup] main_window_visible", flush=True)

    def _log_private_contact_status(self):
        if config.private_contact_library_exists():
            return
        msg = "未检测到甲方联系人库，联系人联想功能已关闭，可手动填写联系人信息。"
        print(f"[privacy] {msg}", flush=True)
        try:
            with open(config.LOG, "a", encoding="utf-8") as f:
                f.write(f"[privacy] {msg}\n")
        except Exception:
            pass

    def _fit_to_screen(self):
        """将窗口居中显示，确保不超出屏幕范围"""
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        scr_w = self.winfo_screenwidth()
        scr_h = self.winfo_screenheight()
        max_w = max(WINDOW_MIN_WIDTH, scr_w - 40)
        max_h = max(WINDOW_MIN_HEIGHT, scr_h - 60)
        w = min(max(w, WINDOW_MIN_WIDTH), max_w)
        h = min(max(h, WINDOW_MIN_HEIGHT), max_h)
        x = max(0, (scr_w - w) // 2)
        y = max(0, (scr_h - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.after_idle(self._adapt_layout)

    def _lazy_load_data(self):
        """延迟加载重量级数据，避免阻塞主窗口初始化"""
        # 先强制窗口显示，再加载数据
        self.update_idletasks()
        # 缓存联系人库和甲方库
        try:
            self._contacts_cache = load_contacts()
        except Exception:
            self._contacts_cache = []
        try:
            self._party_a_options_cache = list(self._party_a_options())
        except Exception:
            self._party_a_options_cache = []
        try:
            self._party_b_names_cache = self._party_b_names()
        except Exception:
            self._party_b_names_cache = config.COMPANIES
        # 延迟填充默认值：先取联系人列表第一个作为默认业务经理
        try:
            managers = self._business_contact_suggestions("", include_all=True)
            if managers and not self.manager_var.get().strip():
                self.manager_var.set(managers[0]["value"])
                self._auto_fill_phone()
        except Exception:
            pass
        # 刷新 AutocompleteInput 的可用值
        if hasattr(self, 'company_input'):
            self.company_input.update_idletasks()

    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", font=("Microsoft YaHei UI", 10), background=BG, foreground=TEXT)
        style.configure("App.TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("Card.TLabelframe", background=CARD, bordercolor=BORDER, relief="solid", padding=4)
        style.configure("Card.TLabelframe.Label", background=BG, foreground=TEXT, font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("TLabel", background=CARD, foreground=TEXT)
        style.configure("Muted.TLabel", background=CARD, foreground=MUTED)
        style.configure("Total.TLabel", background=BG, foreground=TEXT, font=("Microsoft YaHei UI", 13, "bold"))
        style.configure("GridStat.TLabel", background=CARD, foreground=TEXT, font=("Microsoft YaHei UI", 10))
        style.configure("GridAmount.TLabel", background=CARD, foreground=LINK, font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("TEntry", fieldbackground="#ffffff", bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER, padding=3)
        style.configure("Readonly.TEntry", fieldbackground="#eef0f3", bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER, padding=3)
        style.configure("TCombobox", fieldbackground="#ffffff", bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER, padding=3)
        style.configure("TCheckbutton", background=CARD, foreground=TEXT)
        style.configure("TRadiobutton", background=CARD, foreground=TEXT)
        for tree_style, rowheight, heading_pad in (
            ("Resource.Treeview", 32, (8, 8)),
            ("Quote.Treeview", 34, (8, 8)),
        ):
            style.configure(
                tree_style,
                background=CARD,
                fieldbackground=CARD,
                foreground=TEXT,
                rowheight=rowheight,
                bordercolor=GRID_LINE,
                lightcolor=GRID_LINE,
                darkcolor=GRID_LINE,
                borderwidth=1,
                relief="flat",
            )
            style.configure(
                f"{tree_style}.Heading",
                background=GRID_HEADER,
                foreground=TEXT,
                font=("Microsoft YaHei UI", 10, "bold"),
                padding=heading_pad,
                relief="flat",
                bordercolor=GRID_LINE,
            )
            style.map(
                tree_style,
                background=[("selected", GRID_SELECTED)],
                foreground=[("selected", TEXT)],
            )
        style.configure(
            "Modern.Vertical.TScrollbar",
            width=8,
            gripcount=0,
            background="#C9CDD4",
            darkcolor="#C9CDD4",
            lightcolor="#C9CDD4",
            troughcolor="#F5F7FA",
            bordercolor="#F5F7FA",
            arrowcolor="#C9CDD4",
            relief="flat",
        )
        style.map("Modern.Vertical.TScrollbar", background=[("active", "#AAB2BD")])
        style.configure(
            "Modern.Horizontal.TScrollbar",
            width=8,
            gripcount=0,
            background="#C9CDD4",
            darkcolor="#C9CDD4",
            lightcolor="#C9CDD4",
            troughcolor="#F5F7FA",
            bordercolor="#F5F7FA",
            arrowcolor="#C9CDD4",
            relief="flat",
        )
        style.map("Modern.Horizontal.TScrollbar", background=[("active", "#AAB2BD")])
        style.configure("TButton", padding=(10, 4), background="#ffffff", bordercolor=BORDER)
        style.map("TButton", background=[("active", "#f1f3f5")])
        style.configure("Accent.TButton", padding=(10, 4), background=ACCENT, foreground="#ffffff", bordercolor=ACCENT)
        style.map("Accent.TButton", background=[("active", ACCENT_DARK)], foreground=[("active", "#ffffff")])
        style.configure("Danger.TButton", padding=(12, 4), background="#ffffff", foreground=DANGER, bordercolor=DANGER)
        style.map("Danger.TButton", background=[("active", DANGER_HOVER)], foreground=[("active", DANGER)])

    def _build(self):
        # ============================================================
        # 主窗口 grid 布局 — 响应式架构
        # ============================================================
        # 原则：
        #   - 只有 weight>0 的行参与剩余空间分配
        #   - 弹性区内部用嵌套 grid：上部固定控件 + 下部表格(weight=1, sticky=nsew)
        #   - 表格通过 _adapt_layout 动态调整可见行数
        #
        # 行分配（7行）：
        #   0 = 基础信息      weight=0  固定自然高度
        #   1 = 资源搜索      weight=55 弹性（55%剩余空间）
        #   2 = 存储设置      weight=0  固定（从资源搜索中拆出）
        #   3 = 输出设置      weight=0  固定自然高度
        #   4 = 报价明细      weight=45 弹性（45%剩余空间）
        #   5 = 底部操作栏    weight=0  固定始终可见
        # ============================================================
        self.columnconfigure(0, weight=1)

        self.rowconfigure(0, weight=0)
        self.rowconfigure(1, weight=55)
        self.rowconfigure(2, weight=0)
        self.rowconfigure(3, weight=0)
        self.rowconfigure(4, weight=45)
        self.rowconfigure(5, weight=0)

        self._build_basic()              # row=0
        self._build_resource_search()    # row=1
        self._build_storage_settings()   # row=2  ← 新增独立行
        self._build_add_detail()         # row=3
        self._build_quote_detail()       # row=4
        self._build_bottom()             # row=5

        # Configure 事件：窗口大小改变时重新计算表格可见行数
        self.bind("<Configure>", self._adapt_layout, add="+")
        # 显式初始化存储单价（tk StringVar 的 trace 不会在初始值上触发）
        self._on_storage_type_change()

    def _adapt_layout(self, _event=None):
        """响应式布局：仅处理非表格高度的事项（如窗口极大/极小时的padding微调）。
        【关键】不再动态修改 Treeview 的 height 属性——
        Treeview 的 height 控制"请求的数据行数"，而 grid 的 sticky="nsew"
        已经让 Treeview 填满父容器。动态修改 height 会在布局未稳定时
        把 height 写死为极小值，导致 body 区域高度为 0。
        """
        pass  # 布局已由 grid weight + sticky 完全处理，无需动态调整 height

    # 保留空方法，防止子类/其他代码引用时报错
    def _recalc_tree_heights(self):
        """历史遗留方法，已弃用。表格高度现在完全由 grid 布局管理。"""
        pass

    def _build_basic(self):
        box = ttk.Frame(self, style="Card.TFrame", padding=(8, 4))
        box.grid(row=0, column=0, sticky="nsew", pady=(0, 4))
        box.columnconfigure(0, weight=1)

        row1 = ttk.Frame(box, style="Card.TFrame")
        row1.grid(row=0, column=0, sticky="ew", padx=8, pady=2)
        row1.columnconfigure(1, weight=3)
        row1.columnconfigure(3, weight=2)
        row1.columnconfigure(5, weight=2)
        row1.columnconfigure(7, weight=1)

        ttk.Label(row1, text="报价单位").grid(row=0, column=0, sticky="w", padx=(0, 6))
        party_b_names = self._party_b_names()
        self.company_var = tk.StringVar(value=party_b_names[0] if party_b_names else "")
        self.company_input = AutocompleteInput(
            row1,
            self.company_var,
            self._party_b_company_suggestions,
            get_all_suggestions=self._party_b_names,
            on_select=self._on_quote_company_selected,
            width=30,
            style="Card.TFrame",
        )
        self.company_input.grid(row=0, column=1, sticky="ew", padx=(0, 14))

        ttk.Label(row1, text="业务经理").grid(row=0, column=2, sticky="w", padx=(0, 6))
        # 延迟加载：先设空，_lazy_load_data 中再填充
        self.manager_var = tk.StringVar(value="")
        self.manager_input = AutocompleteInput(
            row1,
            self.manager_var,
            self._manager_suggestions,
            get_all_suggestions=lambda: self._business_contact_suggestions("", include_all=True),
            on_select_item=lambda item: self._fill_manager_phone(overwrite=True, contact_item=item),
            width=16,
            style="Card.TFrame",
        )
        self.manager_input.grid(row=0, column=3, sticky="ew", padx=(0, 14))

        ttk.Label(row1, text="联系电话").grid(row=0, column=4, sticky="w", padx=(0, 6))
        self.phone_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.phone_var, width=20).grid(row=0, column=5, sticky="ew", padx=(0, 14))

        ttk.Label(row1, text="日期").grid(row=0, column=6, sticky="w", padx=(0, 6))
        self.date_var = tk.StringVar(value=datetime.now().strftime("%Y.%m.%d"))
        ttk.Entry(row1, textvariable=self.date_var, width=14).grid(row=0, column=7, sticky="ew")

        row2 = ttk.Frame(box, style="Card.TFrame")
        row2.grid(row=1, column=0, sticky="ew", padx=8, pady=2)
        row2.columnconfigure(1, weight=2)
        row2.columnconfigure(3, weight=5)

        ttk.Label(row2, text="项目名称").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.project_var = tk.StringVar(value="计算机时项目")
        ttk.Entry(row2, textvariable=self.project_var, width=30).grid(row=0, column=1, sticky="ew", padx=(0, 14))

        ttk.Label(row2, text="用户单位").grid(row=0, column=2, sticky="w", padx=(0, 6))
        self.unit_var = tk.StringVar()
        self.unit_input = AutocompleteInput(
            row2,
            self.unit_var,
            self._unit_suggestions,
            get_all_suggestions=lambda: [item["name"] for item in self._party_a_options()],
            on_select=self._on_user_unit_selected,
            width=40,
            style="Card.TFrame",
        )
        self.unit_input.grid(row=0, column=3, sticky="ew")

        row3 = ttk.Frame(box, style="Card.TFrame")
        row3.grid(row=2, column=0, sticky="ew", padx=8, pady=2)
        row3.columnconfigure(1, weight=3)
        row3.columnconfigure(5, weight=3)
        row3.columnconfigure(9, weight=2)

        ttk.Label(row3, text="三方公司1").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.tp1_company_var = tk.StringVar(value=config.COMPANIES[1] if len(config.COMPANIES) > 1 else "")
        ttk.Combobox(row3, textvariable=self.tp1_company_var, values=config.COMPANIES, width=26).grid(row=0, column=1, sticky="ew", padx=(0, 10))
        ttk.Label(row3, text="加价").grid(row=0, column=2, sticky="w", padx=(0, 6))
        self.tp1_rate_var = tk.StringVar(value="8")
        self._entry_with_suffix(row3, self.tp1_rate_var, "%", 6).grid(row=0, column=3, sticky="w", padx=(0, 14))

        ttk.Label(row3, text="三方公司2").grid(row=0, column=4, sticky="w", padx=(0, 6))
        self.tp2_company_var = tk.StringVar(value=config.COMPANIES[2] if len(config.COMPANIES) > 2 else "")
        ttk.Combobox(row3, textvariable=self.tp2_company_var, values=config.COMPANIES, width=26).grid(row=0, column=5, sticky="ew", padx=(0, 10))
        ttk.Label(row3, text="加价").grid(row=0, column=6, sticky="w", padx=(0, 6))
        self.tp2_rate_var = tk.StringVar(value="12")
        self._entry_with_suffix(row3, self.tp2_rate_var, "%", 6).grid(row=0, column=7, sticky="w", padx=(0, 14))

        self._auto_fill_phone()  # 初始化时根据默认经理名匹配电话
        self.company_var.trace_add("write", lambda *_args: self._on_quote_company_text_changed())
        self.unit_var.trace_add("write", lambda *_args: self._on_user_unit_text_changed())


    def _entry_with_suffix(self, parent, variable, suffix, width=8):
        frame = ttk.Frame(parent, style="Card.TFrame")
        ttk.Entry(frame, textvariable=variable, width=width).pack(side="left")
        ttk.Label(frame, text=suffix).pack(side="left", padx=(4, 0))
        return frame

    # 存储 List Price 兜底映射；优先使用价格表“存储资源服务”段读取到的真实名称/规格/单价。
    STORAGE_LIST_PRICE = {"分布式文件存储": 80, "高IO分布式文件存储": 300}

    def refresh_unit_suggestions(self, _event=None):
        current = self.unit_var.get()  # 记住当前输入，防止被清空
        suggestions = self._unit_suggestions(current)
        self.unit_suggestions = suggestions
        return suggestions

    def _unit_suggestions(self, text):
        query = text.strip().lower()
        if not query:
            self.unit_suggestions = []
            return []
        suggestions = [
            item["name"]
            for item in self._party_a_options()
            if self._match_text(query, [item["name"], *item.get("aliases", [])])
        ][:30]
        self.unit_suggestions = suggestions
        return suggestions

    def _party_a_options(self):
        if self._party_a_options_cache is not None:
            return self._party_a_options_cache
        names = list_party_a_names()
        public_data = read_json(config.F_CTRACT, {})
        options = {}
        for name in names:
            aliases = []
            if isinstance(public_data, dict):
                entry = public_data.get(name, {})
                if isinstance(entry, dict):
                    aliases = entry.get("aliases", []) or []
            options[name] = {"name": name, "aliases": aliases}
        for entry in load_private_customers():
            name = entry.get("PartyAName") or entry.get("party_a_name") or ""
            if name:
                options.setdefault(name, {"name": name, "aliases": []})
        self._party_a_options_cache = sorted(options.values(), key=lambda item: item["name"])
        return self._party_a_options_cache

    def _party_b_names(self):
        if self._party_b_names_cache is None:
            names = list_party_b_names()
            if not names:
                names = config.COMPANIES
            self._party_b_names_cache = list(dict.fromkeys(names))
        return self._party_b_names_cache

    def _party_b_company_suggestions(self, text):
        query = text.strip().lower()
        if not query:
            return []
        return [name for name in self._party_b_names() if self._match_text(query, [name])][:30]

    def _on_user_unit_selected(self, _value=None):
        """选择用户单位（甲方公司）时：更新合同信息 + 预加载甲方联系人"""
        name = self.unit_var.get().strip()
        if not name:
            return
        self.contract_info["party_a_name"] = name
        info = resolve_party_a(name)
        if info.get("source") == "new":
            return
        # 自动填充甲方公司的已知信息到合同字段
        for key in ("address", "credit", "bank", "account", "contact", "phone", "email"):
            value = info.get(key, "")
            if value:
                self.contract_info[f"party_a_{key}"] = value

    def _on_user_unit_text_changed(self):
        if not self.unit_var.get().strip():
            self.contract_info["party_a_name"] = ""

    def _on_quote_company_selected(self, _value=None):
        self._sync_quote_company_info()

    def _on_quote_company_text_changed(self):
        company = self.company_var.get().strip()
        self.contract_info["party_b_name"] = company
        if not company:
            return
        self._sync_quote_company_info()

    def _sync_quote_company_info(self):
        company = self.company_var.get().strip()
        if not company:
            return
        self.contract_info["party_b_name"] = company
        info = get_party_b(company)
        for key in ("address", "credit", "bank", "account"):
            value = info.get(key, "")
            if value:
                self.contract_info[f"party_b_{key}"] = value

    def _match_text(self, query, fields):
        return any(query in str(field or "").lower() for field in fields)

    def _build_resource_search(self):
        """资源搜索区：上半部固定搜索条件 + 下半部弹性表格（fill=BOTH, expand=True）"""
        box = ttk.Frame(self, style="Card.TFrame", padding=(8, 4))
        box.grid(row=1, column=0, sticky="nsew", pady=(0, 4))
        box.columnconfigure(0, weight=1)

        # ── 内部嵌套 grid：两行结构 ──
        # Row 0: 搜索条件区（关键词 + 资源控制条）— 固定高度
        # Row 1: 表格区 — 弹性，普通窗口至少可见 4 行资源
        # 【关键】没有 minsize 时，若父容器高度不足，grid 会把此行压缩到≈0
        box.rowconfigure(1, weight=1, minsize=RESOURCE_TREE_MIN_BODY)

        # ═══════════════════════════════════════════
        # Row 0: 搜索条件区（固定高度）
        # ═══════════════════════════════════════════
        search_area = ttk.Frame(box, style="Card.TFrame")
        search_area.grid(row=0, column=0, sticky="ew")
        search_area.columnconfigure(2, weight=4)
        search_area.columnconfigure(3, weight=3)

        # --- 关键词行 ---
        ttk.Label(search_area, text="资源搜索", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 12), pady=4)
        ttk.Label(search_area, text="关键词").grid(row=0, column=1, sticky="w", padx=(0, 6), pady=4)
        self.keyword_var = tk.StringVar()
        keyword = ttk.Entry(search_area, textvariable=self.keyword_var)
        keyword.grid(row=0, column=2, sticky="ew", padx=6, pady=4)
        keyword.bind("<KeyRelease>", lambda _e: self.refresh_resource_tree())

        self.price_file_var = tk.StringVar(value="价格表：")
        ttk.Label(search_area, textvariable=self.price_file_var, style="Muted.TLabel", anchor="center").grid(row=0, column=3, sticky="ew", padx=(6, 0), pady=4)
        ttk.Button(search_area, text="更新价格表", command=self.reload_prices, width=12).grid(row=0, column=4, sticky="e", padx=(6, 0), pady=4)

        # --- 资源控制条 ---
        resource_controls = ttk.Frame(search_area, style="Card.TFrame")
        resource_controls.grid(row=1, column=0, columnspan=5, sticky="ew", padx=0, pady=(1, 4))
        for i in range(12):
            resource_controls.columnconfigure(i, weight=0)
        resource_controls.columnconfigure(11, weight=1)

        GAP = 12

        self.mode_var = tk.StringVar(value="total")
        # 报价模式
        ttk.Label(resource_controls, text="报价模式").grid(row=0, column=0, sticky="w", padx=(0, 4))
        mode_frame = ttk.Frame(resource_controls, style="Card.TFrame")
        mode_frame.grid(row=0, column=1, sticky="w", padx=(0, GAP))
        ttk.Radiobutton(mode_frame, text="单价模式", variable=self.mode_var, value="total", command=self._on_mode_change).pack(side="left", padx=(0, 8))
        ttk.Radiobutton(mode_frame, text="数量模式", variable=self.mode_var, value="detail", command=self._on_mode_change).pack(side="left")
        # 资源折扣
        ttk.Label(resource_controls, text="资源折扣").grid(row=0, column=2, sticky="w", padx=(0, 4))
        self.resource_discount_var = tk.StringVar(value="3")
        self._entry_with_suffix(resource_controls, self.resource_discount_var, "折", 5).grid(row=0, column=3, sticky="w", padx=(0, GAP))
        # 折后单价
        ttk.Label(resource_controls, text="折后单价").grid(row=0, column=4, sticky="w", padx=(0, 4))
        self.final_price_var = tk.StringVar()
        ttk.Entry(resource_controls, textvariable=self.final_price_var, width=9).grid(row=0, column=5, sticky="w", padx=(0, GAP))
        # 总价 / 总预算
        self.total_budget_label = ttk.Label(resource_controls, text="总价")
        self.total_budget_label.grid(row=0, column=6, sticky="w", padx=(0, 4))
        self.total_budget_var = tk.StringVar()
        self.total_budget_entry = ttk.Entry(resource_controls, textvariable=self.total_budget_var, width=9)
        self.total_budget_entry.grid(row=0, column=7, sticky="w", padx=(0, GAP))
        # 本项金额 (仅数量模式可见)
        self.item_amount_label_var = tk.StringVar(value="本项金额")
        self.item_amount_label = ttk.Label(resource_controls, textvariable=self.item_amount_label_var)
        self.item_amount_label.grid(row=0, column=8, sticky="w", padx=(0, 4))
        self.item_amount_var = tk.StringVar()
        self.item_amount_entry = ttk.Entry(resource_controls, textvariable=self.item_amount_var, width=9)
        self.item_amount_entry.grid(row=0, column=9, sticky="w", padx=(0, GAP))
        # 自动数量 (仅数量模式可见)
        self.auto_qty_label = ttk.Label(resource_controls, text="数量")
        self.auto_qty_label.grid(row=0, column=10, sticky="w", padx=(0, 4))
        self.auto_qty_var = tk.StringVar(value="-")
        self.auto_qty_value = ttk.Label(resource_controls, textvariable=self.auto_qty_var)
        self.auto_qty_value.grid(row=0, column=11, sticky="w")

        # 初始化为总价模式，隐藏 detail 专属控件
        self._on_mode_change()

        # ══════════════════════════════════════════
        # Row 1: 表格区（弹性 — pack 内部布局）
        # ══════════════════════════════════════════
        tree_outer = ttk.Frame(box, style="Card.TFrame")
        tree_outer.grid(row=1, column=0, sticky="nsew", padx=0, pady=(2, 2))

        # ── tree_outer 内部使用 pack 布局 ──
        # 上部：Treeview(左,fill=both+expand) + 垂直滚动条(右,fill=y)
        # 下部：水平滚动条(fill=x)
        upper_area = ttk.Frame(tree_outer)
        upper_area.pack(side="top", fill="both", expand=True)

        columns = ("rtype", "name", "region", "spec", "price", "action")
        self.resource_tree = ttk.Treeview(upper_area, columns=columns, show="headings", height=4, style="Resource.Treeview")
        headings = {
            "rtype": "资源类型",
            "name": "资源名称",
            "region": "区域",
            "spec": "规格",
            "price": "目录单价",
            "action": "操作",
        }
        widths = {"rtype": 70, "name": 180, "region": 100, "spec": 450, "price": 100, "action": 80}
        for col in columns:
            self.resource_tree.heading(col, text=headings[col])
            anchor = "e" if col == "price" else "center" if col in ("rtype", "region", "action") else "w"
            self.resource_tree.column(col, width=widths[col], minwidth=widths[col], anchor=anchor, stretch=(col == "spec"))
        self.resource_tree.pack(side="left", fill="both", expand=True)
        self._configure_data_grid_tags(self.resource_tree)
        self.resource_full_specs = {}
        self.resource_full_names = {}
        vscroll = ttk.Scrollbar(upper_area, orient="vertical", command=self.resource_tree.yview, style="Modern.Vertical.TScrollbar")
        vscroll.pack(side="right", fill="y")
        hscroll = ttk.Scrollbar(tree_outer, orient="horizontal", command=self.resource_tree.xview, style="Modern.Horizontal.TScrollbar")
        hscroll.pack(side="bottom", fill="x")
        self.resource_tree.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        self._bind_tree_hover(self.resource_tree, self._resource_tree_hover_text)
        self.resource_tree.bind("<ButtonRelease-1>", self._on_resource_tree_click, add="+")
        self._bind_tree_autofit(self.resource_tree, widths, "spec")

    def _build_storage_settings(self):
        """存储设置区：固定高度，独立于资源搜索和报价明细"""
        box = ttk.Frame(self, style="Card.TFrame", padding=(8, 3))
        box.grid(row=2, column=0, sticky="ew", pady=(0, 4))
        box.columnconfigure(0, weight=0)
        box.columnconfigure(9, weight=1)

        ttk.Label(box, text="存储设置", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 12), pady=2)
        ttk.Label(box, text="存储折扣").grid(row=0, column=1, sticky="w", padx=(0, 4), pady=2)
        self.storage_discount_var = tk.StringVar(value="3")
        self._entry_with_suffix(box, self.storage_discount_var, "折", 6).grid(row=0, column=2, sticky="w", padx=(0, 8), pady=2)
        ttk.Label(box, text="折后单价").grid(row=0, column=3, sticky="w", padx=(0, 4), pady=2)
        self.storage_price_var = tk.StringVar(value="")
        ttk.Entry(box, textvariable=self.storage_price_var, width=8).grid(row=0, column=4, sticky="w", padx=(0, 8), pady=2)
        ttk.Label(box, text="存储类型").grid(row=0, column=5, sticky="w", padx=(0, 4), pady=2)
        self.storage_type_var = tk.StringVar(value=config.STORAGE_TYPES[0])
        self.storage_type_combo = ttk.Combobox(box, textvariable=self.storage_type_var, values=config.STORAGE_TYPES, width=18)
        self.storage_type_combo.grid(row=0, column=6, sticky="w", padx=(0, 8), pady=2)
        ttk.Label(box, text="赠送(T)").grid(row=0, column=7, sticky="w", padx=(0, 4), pady=2)
        self.free_storage_var = tk.StringVar(value=str(config.FREE_STORAGE))
        ttk.Entry(box, textvariable=self.free_storage_var, width=6).grid(row=0, column=8, sticky="w", padx=(0, 8), pady=2)
        self.gift_note_var = tk.StringVar(value=f"赠送：{config.FREE_STORAGE}T，免费额度内不收费")
        ttk.Label(box, textvariable=self.gift_note_var, style="Muted.TLabel").grid(row=0, column=9, sticky="w", padx=(0, 4), pady=2)
        action_buttons = ttk.Frame(box, style="Card.TFrame")
        action_buttons.grid(row=0, column=10, sticky="e", padx=(4, 8), pady=2)
        ttk.Button(action_buttons, text="添加资源", command=self.add_selected_resource, width=9).pack(side="left", padx=(0, 3))
        ttk.Button(action_buttons, text="添加存储", command=self.add_storage, width=9).pack(side="left")

    def _build_add_detail(self):
        box = ttk.Frame(self, style="Card.TFrame", padding=(8, 3))
        box.grid(row=3, column=0, sticky="ew", pady=(0, 4))
        box.columnconfigure(5, weight=1)

        self.out_excel_var = tk.BooleanVar(value=True)
        self.out_tp_var = tk.BooleanVar(value=True)
        self.out_contract_var = tk.BooleanVar(value=False)

        ttk.Label(box, text="输出设置", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 12), pady=3)
        ttk.Label(box, text="输出格式").grid(row=0, column=1, sticky="w", padx=(0, 6), pady=3)
        output_frame = ttk.Frame(box, style="Card.TFrame")
        output_frame.grid(row=0, column=2, sticky="w", padx=(0, 12), pady=3)
        ttk.Checkbutton(output_frame, text="Excel", variable=self.out_excel_var).pack(side="left", padx=(0, 18))
        ttk.Checkbutton(output_frame, text="三方Word", variable=self.out_tp_var).pack(side="left", padx=(0, 18))
        ttk.Checkbutton(output_frame, text="合同Word", variable=self.out_contract_var).pack(side="left", padx=(0, 0))
        ttk.Label(box, text="备注").grid(row=0, column=3, sticky="w", padx=(0, 6), pady=3)
        self.remark_var = tk.StringVar(value="1核时=1核心*1小时")
        ttk.Entry(box, textvariable=self.remark_var, width=48).grid(row=0, column=4, columnspan=2, sticky="ew", padx=(0, 8), pady=3)
        ttk.Button(box, text="合同信息", command=self.open_contract_dialog, width=12).grid(row=0, column=6, sticky="e", padx=(0, 8), pady=3)

        self.resource_discount_var.trace_add("write", lambda *_args: self.load_selected_price())
        for var in (self.final_price_var, self.item_amount_var):
            var.trace_add("write", lambda *_args: self.recalc_current())
        self.total_budget_var.trace_add("write", lambda *_args: self._on_total_budget_change())
        self.resource_tree.bind("<<TreeviewSelect>>", lambda _e: self.load_selected_price())
        self.free_storage_var.trace_add("write", lambda *_args: self.update_gift_note())
        self.free_storage_var.trace_add("write", lambda *_args: self._sync_storage_quantity())
        self.storage_price_var.trace_add("write", lambda *_args: self._update_storage_discounted())
        self.storage_discount_var.trace_add("write", lambda *_args: self._update_storage_discounted())
        self.storage_type_var.trace_add("write", lambda *_args: self._on_storage_type_change())
        self._on_mode_change()

    def _build_quote_detail(self):
        """报价明细区：三段式布局（工具栏固定 + 表格弹性fill=BOTH + 统计栏固定）"""
        box = ttk.Frame(self, style="Card.TFrame", padding=(8, 4))
        box.grid(row=4, column=0, sticky="nsew", pady=(4, 4))
        box.columnconfigure(0, weight=1)

        # ── 内部嵌套 grid：三行结构 ──
        # Row 0: 标题栏（固定高度）
        # Row 1: 表格区（弹性，weight=1, sticky=nsew）
        # Row 2: 统计栏（固定高度）
        box.rowconfigure(1, weight=1, minsize=QUOTE_TREE_MIN_BODY)

        # ═══════════════════════════════════════════
        # Row 0: 工具栏（固定高度）
        # ═══════════════════════════════════════════
        toolbar = ttk.Frame(box, style="Card.TFrame")
        toolbar.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 2))
        toolbar.columnconfigure(0, weight=1)
        ttk.Label(toolbar, text="报价明细", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Button(toolbar, text="删除选中", command=self.delete_selected_quote, width=12, style="Danger.TButton").grid(row=0, column=1, sticky="e")

        # ══════════════════════════════════════════
        # Row 1: 表格区（弹性 — pack 内部布局）
        # ══════════════════════════════════════════
        tree_outer = ttk.Frame(box, style="Card.TFrame")
        tree_outer.grid(row=1, column=0, sticky="nsew", padx=0, pady=(0, 2))

        # ── tree_outer 内部使用 pack 布局 ──
        # 上部：Treeview(左,fill=both+expand) + 垂直滚动条(右,fill=y)
        # 下部：水平滚动条(fill=x)
        upper_area = ttk.Frame(tree_outer)
        upper_area.pack(side="top", fill="both", expand=True)

        columns = ("rtype", "name", "spec", "price", "qty", "remark", "amount")
        self.quote_tree = ttk.Treeview(upper_area, columns=columns, show="headings", height=3, style="Quote.Treeview")
        headings = {
            "rtype": "资源类型",
            "name": "资源名称",
            "spec": "规格",
            "price": "单价",
            "qty": "自动数量",
            "remark": "备注",
            "amount": "含税金额",
        }
        widths = {"rtype": 80, "name": 180, "spec": 450, "price": 100, "qty": 100, "remark": 180, "amount": 120}
        for col in columns:
            self.quote_tree.heading(col, text=headings[col])
            anchor = "e" if col in ("price", "amount") else "center" if col in ("rtype", "qty") else "w"
            self.quote_tree.column(col, width=widths[col], minwidth=widths[col], anchor=anchor, stretch=(col == "spec"))
        self.quote_tree.pack(side="left", fill="both", expand=True)
        self._configure_data_grid_tags(self.quote_tree)
        self.quote_full_specs = {}
        self.quote_full_names = {}
        self.quote_full_remarks = {}
        vbar = ttk.Scrollbar(upper_area, orient="vertical", command=self.quote_tree.yview, style="Modern.Vertical.TScrollbar")
        vbar.pack(side="right", fill="y")
        hbar = ttk.Scrollbar(tree_outer, orient="horizontal", command=self.quote_tree.xview, style="Modern.Horizontal.TScrollbar")
        hbar.pack(side="bottom", fill="x")
        self.quote_tree.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        self._bind_tree_hover(self.quote_tree, self._quote_tree_hover_text)
        self.quote_tree.bind("<Configure>", lambda _e: self._update_quote_empty_state(), add="+")
        self._bind_tree_autofit(self.quote_tree, widths, "spec")

        # 空状态占位标签（放在表格容器内，place 定位）
        self.quote_empty_label = tk.Label(
            tree_outer,
            text="暂无报价资源\n请从上方资源列表添加资源",
            bg=CARD,
            fg=EMPTY_TEXT,
            font=("Microsoft YaHei UI", 11),
            justify="center",
        )

        # ═══════════════════════════════════════════
        # Row 2: 统计栏（固定高度）
        # ═══════════════════════════════════════════
        stats_bar = ttk.Frame(box, style="Card.TFrame")
        stats_bar.grid(row=2, column=0, sticky="ew", padx=0, pady=(0, 0))
        stats_bar.columnconfigure(0, weight=1)
        self.quote_count_var = tk.StringVar(value="资源总数：0项")
        self.quote_amount_var = tk.StringVar(value="当前总金额：¥0.00")
        ttk.Label(stats_bar, textvariable=self.quote_count_var, style="GridStat.TLabel").grid(row=0, column=0, sticky="e", padx=(0, 18))
        ttk.Label(stats_bar, textvariable=self.quote_amount_var, style="GridAmount.TLabel").grid(row=0, column=1, sticky="e")

    def _build_bottom(self):
        bar = ttk.Frame(self, style="App.TFrame")
        bar.grid(row=5, column=0, sticky="ew", pady=(4, 0))
        bar.columnconfigure(0, weight=1)  # 左侧扩展
        bar.columnconfigure(1, weight=0)  # 右侧固定

        # 左侧：价格
        left = ttk.Frame(bar, style="App.TFrame")
        left.grid(row=0, column=0, sticky="w", padx=12, pady=6)
        self.tax_total_var = tk.StringVar(value="含税总价： ¥0.00")
        self.no_tax_total_var = tk.StringVar(value="不含税价： ¥0.00")
        ttk.Label(left, textvariable=self.tax_total_var, style="Total.TLabel").pack(side="left", padx=(0, 18))
        ttk.Label(left, textvariable=self.no_tax_total_var, style="Total.TLabel").pack(side="left")

        # 右侧：按钮
        right = ttk.Frame(bar, style="App.TFrame")
        right.grid(row=0, column=1, sticky="e", padx=12, pady=6)

        btn_dir = ttk.Button(right, text="输出目录", command=self._change_output_dir, width=12)
        btn_dir.pack(side="left", padx=(0, 8))
        ToolTip(btn_dir, lambda: config.OUT_BASE)

        ttk.Button(right, text="生成文件", command=self.generate, width=12).pack(side="left", padx=4)
        ttk.Button(right, text="生成PDF", command=self.generate_pdf, width=12).pack(side="left", padx=4)

    def _bind_tree_autofit(self, tree, widths, flex_column):
        tree._autofit_widths = widths
        tree._autofit_column = flex_column
        tree.bind("<Configure>", lambda _event: self._autofit_tree_columns(tree), add="+")
        self.after_idle(lambda: self._autofit_tree_columns(tree))

    def _autofit_tree_columns(self, tree):
        widths = getattr(tree, "_autofit_widths", None)
        flex_column = getattr(tree, "_autofit_column", None)
        if not widths or not flex_column:
            return
        current_width = tree.winfo_width()
        if current_width <= 1:
            return
        fixed_width = sum(width for column, width in widths.items() if column != flex_column)
        flex_width = max(widths[flex_column], current_width - fixed_width - 24)
        for column, width in widths.items():
            tree.column(column, width=flex_width if column == flex_column else width)
        self._refresh_tree_display_text(tree)

    def _refresh_tree_display_text(self, tree):
        if not hasattr(self, "resource_tree") or not hasattr(self, "quote_tree"):
            return
        if tree is self.resource_tree:
            if self.loading_prices:
                return
            rows = tree.get_children()
            for index, row_id in enumerate(rows):
                if index >= len(getattr(self, "filtered_resources", [])):
                    break
                resource = self.filtered_resources[index]
                tree.set(row_id, "name", self._ellipsize_for_column(tree, "name", resource.name))
                tree.set(row_id, "spec", self._ellipsize_for_column(tree, "spec", resource.spec))
        elif tree is self.quote_tree:
            rows = tree.get_children()
            items = self.comp_items + self.stor_items
            for index, row_id in enumerate(rows):
                if index >= len(items):
                    break
                item = items[index]
                tree.set(row_id, "name", self._ellipsize_for_column(tree, "name", item.name))
                tree.set(row_id, "spec", self._ellipsize_for_column(tree, "spec", item.spec))
                tree.set(row_id, "remark", self._ellipsize_for_column(tree, "remark", item.remark))

    def _ellipsize_for_column(self, tree, column, text):
        try:
            width = int(tree.column(column, "width")) - 18
        except tk.TclError:
            width = 120
        return self._ellipsize_to_width(text, width)

    def _ellipsize_to_width(self, text, max_width):
        text = str(text or "")
        if not text:
            return ""
        if max_width <= 28:
            return "..."
        font = getattr(self, "_grid_text_font", None)
        if font is None:
            font = self._grid_text_font = tkfont.Font(family="Microsoft YaHei UI", size=10)
        if font.measure(text) <= max_width:
            return text
        suffix = "..."
        lo, hi = 0, len(text)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if font.measure(text[:mid] + suffix) <= max_width:
                lo = mid
            else:
                hi = mid - 1
        return text[:lo] + suffix

    def _configure_data_grid_tags(self, tree):
        tree.tag_configure("row_odd", background=CARD, foreground=TEXT)
        tree.tag_configure("row_even", background=GRID_ALT, foreground=TEXT)
        tree.tag_configure("hover", background=GRID_HOVER, foreground=TEXT)

    def _bind_tree_hover(self, tree, tooltip_getter):
        tree._hover_row = ""
        tree.bind("<Motion>", lambda event, t=tree, getter=tooltip_getter: self._on_tree_motion(event, t, getter), add="+")
        tree.bind("<Leave>", lambda event, t=tree: self._on_tree_leave(t), add="+")

    def _on_tree_motion(self, event, tree, tooltip_getter):
        row_id = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)
        self._set_tree_hover(tree, row_id)
        if tree is self.resource_tree and col_id == "#6" and row_id:
            tree.configure(cursor="hand2")
        else:
            tree.configure(cursor="")
        tooltip = tooltip_getter(row_id, col_id)
        if tooltip:
            self._show_tree_tooltip(tooltip, event.x_root + 12, event.y_root + 18)
        else:
            self._hide_tree_tooltip()

    def _on_tree_leave(self, tree):
        self._set_tree_hover(tree, "")
        tree.configure(cursor="")
        self._hide_tree_tooltip()

    def _set_tree_hover(self, tree, row_id):
        old_row = getattr(tree, "_hover_row", "")
        if old_row and tree.exists(old_row):
            tags = tuple(tag for tag in tree.item(old_row, "tags") if tag != "hover")
            tree.item(old_row, tags=tags)
        if row_id and tree.exists(row_id):
            tags = list(tree.item(row_id, "tags"))
            if "hover" not in tags:
                tags.append("hover")
            tree.item(row_id, tags=tuple(tags))
        tree._hover_row = row_id

    def _resource_tree_hover_text(self, row_id, col_id):
        if not row_id:
            return ""
        if col_id == "#2":
            full = self.resource_full_names.get(row_id, "")
            shown = self.resource_tree.set(row_id, "name")
            return full if full and full != shown else ""
        if col_id == "#4":
            full = self.resource_full_specs.get(row_id, "")
            shown = self.resource_tree.set(row_id, "spec")
            return full if full and full != shown else ""
        return ""

    def _quote_tree_hover_text(self, row_id, col_id):
        if not row_id:
            return ""
        if col_id == "#2":
            full = self.quote_full_names.get(row_id, "")
            shown = self.quote_tree.set(row_id, "name")
            return full if full and full != shown else ""
        if col_id == "#3":
            full = self.quote_full_specs.get(row_id, "")
            shown = self.quote_tree.set(row_id, "spec")
            return full if full and full != shown else ""
        if col_id == "#6":
            full = self.quote_full_remarks.get(row_id, "")
            shown = self.quote_tree.set(row_id, "remark")
            return full if full and full != shown else ""
        return ""

    def _show_tree_tooltip(self, text, x_root, y_root):
        if not hasattr(self, "_tree_tooltip"):
            self._tree_tooltip = tk.Label(
                self,
                bg="#FFFBE6",
                fg="#1F2937",
                padx=7,
                pady=4,
                wraplength=360,
                justify="left",
                font=("Microsoft YaHei UI", 9),
                relief="flat",
                borderwidth=0,
                highlightthickness=1,
                highlightbackground="#D9D9D9",
                highlightcolor="#D9D9D9",
            )
        root_x = self.winfo_rootx()
        root_y = self.winfo_rooty()
        self._tree_tooltip.configure(text=text)
        self._tree_tooltip.update_idletasks()
        x = x_root - root_x
        y = y_root - root_y
        req_w = self._tree_tooltip.winfo_reqwidth()
        req_h = self._tree_tooltip.winfo_reqheight()
        max_x = max(4, self.winfo_width() - req_w - 8)
        max_y = max(4, self.winfo_height() - req_h - 8)
        self._tree_tooltip.place(x=min(max(x, 4), max_x), y=min(max(y, 4), max_y))
        self._tree_tooltip.lift()

    def _hide_tree_tooltip(self):
        if hasattr(self, "_tree_tooltip"):
            self._tree_tooltip.place_forget()

    def _ellipsize(self, text, max_chars):
        text = str(text or "")
        return text if len(text) <= max_chars else text[: max_chars - 3] + "..."

    def _row_tag(self, index):
        return "row_odd" if index % 2 == 0 else "row_even"

    def _on_resource_tree_click(self, event):
        row_id = self.resource_tree.identify_row(event.y)
        col_id = self.resource_tree.identify_column(event.x)
        if row_id and col_id == "#6":
            self.resource_tree.selection_set(row_id)
            self.resource_tree.focus(row_id)
            self.load_selected_price()
            self.add_selected_resource()
            return "break"
        return None

    def _on_manager_changed(self, _event=None):
        """从下拉列表选中时，填入对应电话"""
        self._fill_manager_phone(overwrite=True)

    def _on_manager_typing(self, _event=None):
        """打字时：过滤下拉建议 + 匹配电话（不丢失当前输入）"""
        self._auto_fill_phone()

    def _all_manager_names(self):
        return self._business_contact_suggestions("", include_all=True)

    def _manager_suggestions(self, text):
        return self._business_contact_suggestions(text)

    def _auto_fill_phone(self):
        """根据当前业务经理名字自动匹配并填入联系电话"""
        self._fill_manager_phone(overwrite=False)

    def _fill_manager_phone(self, overwrite=False, contact_item=None):
        name = self.manager_var.get().strip()
        if not name:
            return
        contact = contact_item if isinstance(contact_item, dict) else None
        if not contact:
            for item in self._business_contact_records():
                if item.get("contact", "").strip() == name:
                    contact = item
                    break
        if not contact:
            return
        phone = contact.get("phone", "")
        # 只在电话为空时自动填充，避免覆盖用户手写的号码
        if phone and (overwrite or not self.phone_var.get().strip()):
            self.phone_var.set(phone)

    def _contacts(self):
        if self._contacts_cache is None:
            self._contacts_cache = load_contacts()
        return self._contacts_cache

    def _business_contact_records(self):
        company = self.company_var.get().strip() if hasattr(self, "company_var") else ""
        party_b_names = set(self._party_b_names())
        preferred = []
        other_party_b = []
        for contact in self._contacts():
            unit = contact.get("unit_name", "").strip()
            if company and unit == company:
                preferred.append(contact)
            elif unit in party_b_names:
                other_party_b.append(contact)
        return self._sort_contacts(preferred) + self._sort_contacts(other_party_b)

    def _sort_contacts(self, contacts):
        return sorted(contacts, key=lambda item: item.get("updated_at", ""), reverse=True)

    def _business_contact_suggestions(self, text, include_all=False):
        query = text.strip().lower()
        if not query and not include_all:
            return []
        results = []
        seen = set()
        for contact in self._business_contact_records():
            name = contact.get("contact", "").strip()
            if not name:
                continue
            if query and not self._match_text(
                query,
                [name, contact.get("phone", ""), contact.get("email", ""), contact.get("unit_name", "")],
            ):
                continue
            key = (contact.get("unit_name", "").strip(), name)
            if key in seen:
                continue
            seen.add(key)
            results.append(
                {
                    "label": self._contact_label(contact),
                    "value": name,
                    "raw": contact,
                }
            )
            if len(results) >= 30:
                break
        return results

    def _contact_label(self, contact):
        parts = [
            contact.get("contact", "").strip(),
            contact.get("phone", "").strip(),
            contact.get("unit_name", "").strip(),
        ]
        return "｜".join(part for part in parts if part)

    def _save_manager(self):
        """将当前业务经理保存到当前乙方公司的联系人库。"""
        name = self.manager_var.get().strip()
        phone = self.phone_var.get().strip()
        company = self.company_var.get().strip()
        if not name or not phone or not company:
            return  # 名字和电话都有才保存
        existing = {}
        for contact in self._contacts():
            if contact.get("unit_name") == company and contact.get("contact") == name:
                existing = dict(contact)
                break
        existing.update({"unit_name": company, "contact": name, "phone": phone})
        save_contact_entry(existing)
        self._contacts_cache = None

    def reload_prices(self):
        if self.loading_prices:
            return
        self.loading_prices = True
        self.price_load_error = None
        self.resources = []
        self.filtered_resources = []
        self.storage_resources = []
        self.storage_by_name = {}
        self.price_file_var.set("价格表：正在加载价格表，请稍候...")
        self._show_resource_loading()
        threading.Thread(target=self._load_prices_worker, daemon=True).start()
        self.after(100, self._poll_price_load_queue)

    def _load_prices_worker(self):
        load_start = time.perf_counter()
        try:
            import_start = time.perf_counter()
            from engines.price_loader import load as load_prices, load_storage, find_price
            self._log_startup("load_price_loader_module", import_start)
            price_file = find_price()
            self._log_startup("find_price_file", load_start)
            timings = {}
            total_load_start = time.perf_counter()
            resources = load_prices(price_file, timings=timings)
            storage_resources = load_storage(price_file, timings=timings)
            self._log_startup_elapsed("read_price_table", timings.get("read_price_table", 0))
            self._log_startup_elapsed("parse_resource_catalog", timings.get("parse_resource_catalog", 0))
            self._log_startup_elapsed("read_storage_table", timings.get("read_storage_table", 0))
            self._log_startup_elapsed("parse_storage_catalog", timings.get("parse_storage_catalog", 0))
            self._log_startup("load_price_resources_total", total_load_start)
            self.price_load_queue.put((price_file, resources, storage_resources, None))
        except Exception as exc:
            self.price_load_queue.put((None, [], [], exc))

    def _poll_price_load_queue(self):
        try:
            price_file, resources, storage_resources, error = self.price_load_queue.get_nowait()
        except queue.Empty:
            if self.loading_prices:
                self.after(100, self._poll_price_load_queue)
            return
        self._on_prices_loaded(price_file, resources, storage_resources, error)

    def _on_prices_loaded(self, price_file, resources, storage_resources, error):
        self.loading_prices = False
        self.resources = resources or []
        self.storage_resources = storage_resources or []
        if error:
            self.price_load_error = error
            self.price_file_var.set(f"价格表：加载失败 - {error}")
            self.refresh_resource_tree()
            return
        name = os.path.basename(price_file) if price_file else "未找到价格表"
        self.price_file_var.set(f"价格表：{name}（计算资源 {len(self.resources)} 条，存储 {len(self.storage_resources)} 条）")
        self._refresh_storage_options()
        render_start = time.perf_counter()
        self.refresh_resource_tree()
        self._log_startup("init_resource_table", render_start)

    def _show_resource_loading(self):
        for item in self.resource_tree.get_children():
            self.resource_tree.delete(item)
        self.resource_full_specs = {}
        self.resource_full_names = {}
        self.resource_tree.insert(
            "",
            "end",
            values=("", "正在加载价格表，请稍候...", "", "", "", ""),
            tags=("row_odd",),
        )

    def refresh_resource_tree(self):
        for item in self.resource_tree.get_children():
            self.resource_tree.delete(item)
        self.resource_full_specs = {}
        self.resource_full_names = {}
        if self.loading_prices:
            self._show_resource_loading()
            return
        self.filtered_resources = self._search_resources(self.keyword_var.get())
        visible_resources = self.filtered_resources[:50]
        for index, resource in enumerate(visible_resources):
            row_id = self.resource_tree.insert(
                "",
                "end",
                values=(
                    resource.rtype,
                    self._ellipsize_for_column(self.resource_tree, "name", resource.name),
                    resource.region,
                    self._ellipsize_for_column(self.resource_tree, "spec", resource.spec),
                    f"{resource.shared_price:g}",
                    "+ 添加",
                ),
                tags=(self._row_tag(index),),
            )
            self.resource_full_specs[row_id] = str(resource.spec or "")
            self.resource_full_names[row_id] = str(resource.name or "")

    def _search_resources(self, query):
        query = (query or "").strip().lower()
        if not query:
            return list(self.resources)
        return [
            resource for resource in self.resources
            if query in f"{resource.rtype} {resource.name} {resource.region} {resource.spec}".lower()
        ]

    def selected_resource(self):
        if self.loading_prices:
            return None
        selection = self.resource_tree.selection()
        if not selection:
            return None
        index = self.resource_tree.index(selection[0])
        if index >= len(self.filtered_resources):
            return None
        return self.filtered_resources[index]

    def load_selected_price(self):
        resource = self.selected_resource()
        if not resource:
            return
        discount = self.parse_discount(self.resource_discount_var.get())
        self.final_price_var.set(f"{resource.shared_price * discount:.6g}")
        self.remark_var.set("1卡时=1张卡*1小时" if resource.rtype in ("GPU", "异构加速卡") else "1核时=1核心*1小时")
        self.recalc_current()

    def parse_discount(self, text):
        try:
            value = float(text)
        except (TypeError, ValueError):
            return 1.0
        return value / 10 if value > 1 else value

    def parse_float(self, text, default=0.0):
        try:
            return float(str(text).replace(",", "").strip())
        except (TypeError, ValueError):
            return default

    def recalc_current(self):
        price = self.parse_float(self.final_price_var.get())
        amount = self.parse_float(self.item_amount_var.get())
        if price > 0 and amount > 0:
            self.auto_qty_var.set(f"{amount / price:.2f}")
        else:
            self.auto_qty_var.set("-")

    def _on_mode_change(self):
        if self.mode_var.get() == "detail":
            self.item_amount_label_var.set("本项金额/权重")
            self.item_amount_label.grid()
            self.item_amount_entry.grid()
            self.total_budget_label.configure(text="总预算")
            self.total_budget_label.grid()
            self.total_budget_entry.state(["!disabled"])
            self.total_budget_entry.grid()
            self.auto_qty_label.grid()
            self.auto_qty_value.grid()
        else:
            self.item_amount_label_var.set("本项金额")
            self.item_amount_label.grid_remove()
            self.item_amount_entry.grid_remove()
            self.total_budget_label.configure(text="总价")
            self.total_budget_label.grid()
            self.total_budget_entry.state(["!disabled"])
            self.total_budget_entry.grid()
            self.auto_qty_label.grid_remove()
            self.auto_qty_value.grid_remove()
        self.recalc_current()

    def _on_total_budget_change(self):
        self.recalc_current()
        self.distribute_detail_budget()
        self.refresh_quote_tree()

    def distribute_detail_budget(self):
        if self.mode_var.get() != "detail":
            return
        total_budget = self.parse_float(self.total_budget_var.get())
        if total_budget <= 0:
            return
        detail_items = [
            item
            for item in self.comp_items
            if getattr(item, "alloc_weight", None) is not None and (item.price or 0) > 0
        ]
        if not detail_items:
            return
        weights = [max(float(getattr(item, "alloc_weight", 0) or 0), 0.0) for item in detail_items]
        total_weight = sum(weights)
        if total_weight <= 0:
            weights = [1.0 for _item in detail_items]
            total_weight = float(len(detail_items))

        allocated = 0.0
        for index, item in enumerate(detail_items):
            if index == len(detail_items) - 1:
                amount = round(max(total_budget - allocated, 0), 2)
            else:
                amount = round(total_budget * weights[index] / total_weight, 2)
                allocated += amount
            item.amount = amount
            item.quantity = round(amount / item.price, 2) if item.price else 0

    def update_gift_note(self):
        tb = self.parse_float(self.free_storage_var.get(), config.FREE_STORAGE)
        self.gift_note_var.set(f"赠送：{tb:g}T，免费额度内不收费")

    def _sync_storage_quantity(self):
        """赠送(T)变化时同步更新已有存储项的数量"""
        tb = self.parse_float(self.free_storage_var.get(), config.FREE_STORAGE)
        changed = False
        for item in self.stor_items:
            if getattr(item, "free_quota", False):
                item.quantity = tb
                item.remark = f"赠送：{tb:g}T，免费额度内不收费"
                changed = True
        if changed:
            self.refresh_quote_tree()

    def _on_storage_type_change(self):
        """存储类型变化时自动填充折后单价（List Price × 折扣）"""
        storage = self._selected_storage_resource()
        stor_type = self.storage_type_var.get().strip()
        list_price = self._storage_list_price(storage, stor_type)
        discount = self.parse_discount(self.storage_discount_var.get())
        if list_price:
            discounted = round(list_price * discount, 4)
            self.storage_price_var.set(str(discounted))
            # 同步更新已有存储项的名称和规格
            name, spec = self._storage_name_spec(storage, stor_type)
            for item in self.stor_items:
                if getattr(item, "free_quota", False):
                    item.name = name
                    item.spec = spec
        self.refresh_quote_tree()

    def _update_storage_discounted(self):
        """折扣改变时，用当前 List Price 重算折后单价，并同步已有存储项"""
        storage = self._selected_storage_resource()
        stor_type = self.storage_type_var.get().strip()
        list_price = self._storage_list_price(storage, stor_type)
        discount = self.parse_discount(self.storage_discount_var.get())
        if list_price > 0 and discount > 0:
            discounted = round(list_price * discount, 4)
            self.storage_price_var.set(str(discounted))
            # 同步已有存储项的 orig_price / price / discount
            name, spec = self._storage_name_spec(storage, stor_type)
            for item in self.stor_items:
                if getattr(item, "free_quota", False):
                    item.name = name
                    item.spec = spec
                    item.orig_price = list_price
                    item.discount = discount
                    item.price = discounted
        self.refresh_quote_tree()

    def _refresh_storage_options(self):
        options = []
        self.storage_by_name = {}
        for resource in self.storage_resources:
            name = (resource.name or "").strip()
            if not name:
                continue
            if name not in self.storage_by_name:
                options.append(name)
                self.storage_by_name[name] = resource
        if not options:
            options = list(config.STORAGE_TYPES)
            self.storage_by_name = {}
        if hasattr(self, "storage_type_combo"):
            self.storage_type_combo.configure(values=options)
        if self.storage_type_var.get().strip() not in options and options:
            self.storage_type_var.set(options[0])
        else:
            self._on_storage_type_change()

    def _selected_storage_resource(self):
        return self.storage_by_name.get(self.storage_type_var.get().strip())

    def _storage_list_price(self, storage, stor_type):
        if storage and (storage.shared_price or 0) > 0:
            return storage.shared_price
        return self.STORAGE_LIST_PRICE.get(stor_type, 0)

    def _storage_name_spec(self, storage, stor_type):
        if storage:
            name = storage.name or stor_type
            spec = storage.spec or storage.name or stor_type
            return name, spec
        return stor_type, stor_type

    def add_selected_resource(self):
        if self.loading_prices:
            messagebox.showinfo("提示", "价格表正在加载，请稍候...")
            return
        resource = self.selected_resource()
        if not resource:
            messagebox.showwarning("提示", "请先在资源搜索中选择一项资源")
            return
        price = self.parse_float(self.final_price_var.get(), resource.shared_price)
        amount_input = self.parse_float(self.item_amount_var.get())
        if price <= 0:
            messagebox.showwarning("提示", "请填写有效的折后单价")
            return
        if self.mode_var.get() == "detail":
            amount = amount_input if amount_input > 0 else 0
            alloc_weight = amount_input if amount_input > 0 else 1
        else:
            amount = 0
            alloc_weight = None
        if price > 0 and amount > 0:
            quantity = round(amount / price, 2)
        else:
            quantity = 0
            amount = 0
        item = QuoteItem(
            item_type="Compute",
            rtype=resource.rtype,
            name=resource.name,
            spec=resource.spec,
            orig_price=resource.shared_price,
            discount=self.parse_discount(self.resource_discount_var.get()),
            price=price,
            quantity=quantity,
            amount=amount,
            remark=self.remark_var.get(),
        )
        if alloc_weight is not None:
            item.alloc_weight = alloc_weight
        item.quote_mode = self.mode_var.get()
        self.comp_items.append(item)
        self.distribute_detail_budget()
        self.refresh_quote_tree()

    def add_storage(self):
        tb = self.parse_float(self.free_storage_var.get(), config.FREE_STORAGE)
        storage = self._selected_storage_resource()
        stor_type = self.storage_type_var.get().strip()
        discount = self.parse_discount(self.storage_discount_var.get())
        price = self.parse_float(self.storage_price_var.get())  # 折后单价
        list_price = self._storage_list_price(storage, stor_type)
        orig_price = list_price if list_price > 0 else round(price / discount, 4) if discount > 0 else price
        name, spec = self._storage_name_spec(storage, stor_type)
        item = QuoteItem(
            item_type="Storage",
            rtype="存储",
            name=name,
            spec=spec,
            orig_price=orig_price,
            discount=discount,
            price=price,
            quantity=tb,
            amount=0,
            remark=f"赠送：{tb:g}T，免费额度内不收费",
            free_quota=True,
        )
        self.stor_items.append(item)
        self.refresh_quote_tree()

    def delete_selected_quote(self):
        for row_id in reversed(self.quote_tree.selection()):
            index = self.quote_tree.index(row_id)
            if index < len(self.comp_items):
                del self.comp_items[index]
            else:
                storage_index = index - len(self.comp_items)
                if storage_index < len(self.stor_items):
                    del self.stor_items[storage_index]
        self.distribute_detail_budget()
        self.refresh_quote_tree()

    def refresh_quote_tree(self):
        for row_id in self.quote_tree.get_children():
            self.quote_tree.delete(row_id)
        self.quote_full_specs = {}
        self.quote_full_names = {}
        self.quote_full_remarks = {}
        for index, item in enumerate(self.comp_items + self.stor_items):
            row_id = self.quote_tree.insert(
                "",
                "end",
                values=(
                    item.rtype,
                    self._ellipsize_for_column(self.quote_tree, "name", item.name),
                    self._ellipsize_for_column(self.quote_tree, "spec", item.spec),
                    f"{item.price:g}",
                    "-" if not item.quantity else f"{item.quantity:g}",
                    self._ellipsize_for_column(self.quote_tree, "remark", item.remark),
                    f"{item.amount:.2f}",
                ),
                tags=(self._row_tag(index),),
            )
            self.quote_full_specs[row_id] = str(item.spec or "")
            self.quote_full_names[row_id] = str(item.name or "")
            self.quote_full_remarks[row_id] = str(item.remark or "")
        total = sum((item.amount or 0) for item in self.comp_items + self.stor_items)
        # total模式：总价来自用户输入的total_budget_var
        if self.mode_var.get() == "total":
            total = self.parse_float(self.total_budget_var.get())
        self._update_quote_empty_state()
        self._update_quote_stats(total)
        self.tax_total_var.set(f"含税总价： ¥{fmt_money(total)}")
        self.no_tax_total_var.set(f"不含税价： ¥{fmt_money(total / config.TAX)}")

    def _update_quote_empty_state(self):
        if not hasattr(self, "quote_empty_label"):
            return
        if self.quote_tree.get_children():
            self.quote_empty_label.place_forget()
        else:
            self.quote_tree.update_idletasks()
            heading_height = 42
            body_height = max(self.quote_tree.winfo_height() - heading_height, 90)
            self.quote_empty_label.place(in_=self.quote_tree, relx=0.5, y=heading_height + body_height / 2, anchor="center")
            self.quote_empty_label.lift()

    def _update_quote_stats(self, total):
        if not hasattr(self, "quote_count_var"):
            return
        count = len(self.comp_items) + len(self.stor_items)
        self.quote_count_var.set(f"资源总数：{count}项")
        self.quote_amount_var.set(f"当前总金额：¥{fmt_money(total)}")

    def get_header_values(self):
        return {
            "company": self.company_var.get(),
            "date": self.date_var.get(),
            "manager": self.manager_var.get(),
            "phone": self.phone_var.get(),
            "unit": self.unit_var.get(),
            "project": self.project_var.get(),
        }

    def get_items(self):
        return self.comp_items, self.stor_items

    def _change_output_dir(self):
        new_dir = filedialog.askdirectory(
            title="选择生成文件保存位置",
            initialdir=config.OUT_BASE
        )
        if new_dir:
            config.set_output_dir(new_dir)

    def open_contract_dialog(self):
        import_start = time.perf_counter()
        from engines.contract_writer import ensure_university_contracts
        from gui.dialogs.contract_dialog import ContractDialog
        ensure_university_contracts()
        self._log_startup("load_contract_dialog_modules", import_start)

        # 保存当前联系人（新名字+新电话自动追加到列表）
        self._save_manager()

        # 首页字段同步到 contract_info，再传入对话框
        info = dict(self.contract_info)  # 拷贝避免直接修改原字典
        info["party_a_name"] = self.unit_var.get()   # 用户单位 → 甲方公司名称
        info["party_b_name"] = self.company_var.get()      # 报价单位 → 乙方公司名称
        info["party_b_contact"] = self.manager_var.get()   # 业务经理 → 乙方联系人姓名
        info["party_b_phone"] = self.phone_var.get()       # 联系电话 → 乙方联系电话

        ContractDialog(self, self.unit_var.get(), self._set_contract_info, info)

    def _set_contract_info(self, info):
        self.contract_info = info
        # 同步甲方名称到首页单位名称
        a_name = info.get("party_a_name", "").strip()
        if a_name:
            self.unit_var.set(a_name)
        b_name = info.get("party_b_name", "").strip()
        if b_name:
            self.company_var.set(b_name)
        b_contact = info.get("party_b_contact", "").strip()
        if b_contact:
            self.manager_var.set(b_contact)
        b_phone = info.get("party_b_phone", "").strip()
        if b_phone:
            self.phone_var.set(b_phone)

    def third_party_rule(self):
        """三方报价固定规则：同等总价下，单价更高并反算数量"""
        return "SameTotalMoreQty"

    def percent_to_markup(self, text):
        return self.parse_float(text) / 100

    def generate(self):
        try:
            self.generated_files = self._generate_files(export_pdf=False)
            if not self.generated_files:
                messagebox.showwarning("提示", "请至少选择一种输出格式")
                return
            messagebox.showinfo("完成", "生成成功：\n" + "\n".join(self.generated_files))
        except Exception as exc:
            show_err(self, exc, "生成失败")

    def generate_pdf(self):
        try:
            files = self._generate_files(export_pdf=False)
            import_start = time.perf_counter()
            from engines.pdf_exporter import convert_to_pdf
            self._log_startup("load_pdf_modules", import_start)
            pdf_files = []
            for path in files:
                pdf_files.append(convert_to_pdf(path))
            self.generated_files = files + pdf_files
            messagebox.showinfo("完成", "PDF生成成功：\n" + "\n".join(pdf_files))
        except Exception as exc:
            show_err(self, exc, "PDF生成失败")

    def _generate_files(self, export_pdf=False):
        import_start = time.perf_counter()
        from engines.excel_writer import generate as gen_excel
        from engines.word_writer import generate_standard, generate_yunda
        from engines.contract_writer import generate as gen_contract
        self._log_startup("load_template_modules", import_start)

        header = self.get_header_values()
        if not header["unit"]:
            raise ValueError("请填写用户单位")
        if not self.comp_items and not self.stor_items:
            raise ValueError("请先添加资源或存储")
        if not self.stor_items:
            raise ValueError("请先添加存储资源：如有免费赠送，也需要点击「添加存储」写入报价明细")

        save_hist(header["unit"])
        files = []
        is_total = self.mode_var.get() == "total"

        # detail（数量）模式：重新执行预算分配，确保使用最新的总预算
        if not is_total:
            detail_budget = self.parse_float(self.total_budget_var.get())
            if detail_budget <= 0:
                raise ValueError("数量模式请填写总预算后再生成文件")
            self.distribute_detail_budget()

        total = self.parse_float(self.total_budget_var.get()) if is_total else sum((item.amount or 0) for item in self.comp_items + self.stor_items)
        if is_total and total <= 0:
            raise ValueError("单价模式请填写总价")

        # ========== 调试：打印传递给 Excel 的数据 ==========
        print("[GUI_DBG] " + "=" * 60, flush=True)
        print(f"[GUI_DBG] mode={self.mode_var.get()} is_total={is_total} total={total}", flush=True)
        print(f"[GUI_DBG] comp_items ({len(self.comp_items)} 项):", flush=True)
        for i, it in enumerate(self.comp_items):
            print(f"[GUI_DBG]   [{i}] rtype={it.rtype} name={it.name} "
                  f"price={it.price} qty={it.quantity} amount={it.amount} "
                  f"disc={getattr(it,'discount','?')} "
                  f"alloc_weight={getattr(it,'alloc_weight','N/A')}", flush=True)
        print(f"[GUI_DBG] stor_items ({len(self.stor_items)} 项):", flush=True)
        for i, it in enumerate(self.stor_items):
            print(f"[GUI_DBG]   [{i}] rtype={it.rtype} name={it.name} "
                  f"price={it.price} qty={it.quantity} amount={it.amount} "
                  f"free_quota={getattr(it,'free_quota','?')}", flush=True)
        print("[GUI_DBG] " + "=" * 60, flush=True)
        # ====================================================

        if self.out_excel_var.get():
            files.append(
                gen_excel(
                    self.comp_items,
                    self.stor_items,
                    is_total,
                    total,
                    header["unit"],
                    header["date"],
                    header["company"],
                    header["manager"],
                    header["phone"],
                    header["project"],
                )
            )

        if self.out_tp_var.get():
            rule = self.third_party_rule()
            markup1 = self.percent_to_markup(self.tp1_rate_var.get())
            markup2 = self.percent_to_markup(self.tp2_rate_var.get())
            items_1 = comp_items(self.comp_items, markup1, rule, is_total)
            items_2 = comp_items(self.comp_items, markup2, rule, is_total)
            # 计算第三方报价单总价（固定规则：同等总价下单价更高）
            if is_total:
                tp1_total = total          # 总价不变
                tp2_total = total
            else:
                tp1_total = sum((it.amount or 0) for it in items_1) + sum((it.amount or 0) for it in self.stor_items)
                tp2_total = sum((it.amount or 0) for it in items_2) + sum((it.amount or 0) for it in self.stor_items)
            files.append(
                generate_standard(
                    items_1,
                    self.stor_items,
                    is_total,
                    tp1_total,
                    header["unit"],
                    header["date"],
                    self.tp1_company_var.get(),
                    header["manager"],
                    header["phone"],
                    header["project"],
                )
            )
            files.append(
                generate_yunda(
                    items_2,
                    self.stor_items,
                    is_total,
                    tp2_total,
                    header["unit"],
                    header["date"],
                    self.tp2_company_var.get(),
                    header["manager"],
                    header["phone"],
                    header["project"],
                )
            )

        if self.out_contract_var.get():
            if not self.contract_info:
                self.open_contract_dialog()
                raise ValueError("请先填写合同信息")
            ci = self.contract_info
            files.append(
                gen_contract(
                    header["unit"],
                    header["date"],
                    header["company"],
                    header["manager"],
                    header["phone"],
                    ci.get("party_a_name", ""),
                    ci.get("party_a_contact", ""),
                    ci.get("party_a_phone", ""),
                    ci.get("party_a_email", ""),
                    ci.get("party_a_address", ""),
                    ci.get("party_a_mailing_address", ""),
                    ci.get("party_a_credit", ""),
                    ci.get("party_a_bank", ""),
                    ci.get("party_a_account", ""),
                    ci.get("party_b_name", header["company"]),
                    ci.get("party_b_contact", header["manager"]),
                    ci.get("party_b_phone", header["phone"]),
                    ci.get("party_b_email", ""),
                    ci.get("party_b_address", ""),
                    ci.get("party_b_mailing_address", ""),
                    ci.get("party_b_credit", ""),
                    ci.get("party_b_bank", ""),
                    ci.get("party_b_account", ""),
                    ci.get("sign_place", ""),
                    ci.get("sign_time", header["date"]),
                    ci.get("service_term", "1年"),
                    ci.get("pay_days", "30"),
                    ci.get("invoice", "6%增值税普通发票"),
                    total,
                    self.comp_items,
                    self.stor_items,
                    header["project"],
                )
            )
        return files


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
