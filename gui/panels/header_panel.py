# gui/panels/header_panel.py — 顶部信息面板
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import config
from utils import search_units, save_hist, load_hist

class HeaderPanel(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, padding=5)
        self.app = app
        self._build()

    def _build(self):
        f = ttk.LabelFrame(self, text="基础信息")
        f.pack(fill="x", padx=5, pady=5)

        # Row 0: 报价单位
        ttk.Label(f, text="报价单位:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.company_var = tk.StringVar(value=config.COMPANIES[0] if config.COMPANIES else "")
        self.company_cb = ttk.Combobox(f, textvariable=self.company_var, values=config.COMPANIES, width=30)
        self.company_cb.grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        # Row 0: 报价日期
        ttk.Label(f, text="报价日期:").grid(row=0, column=2, sticky="w", padx=5, pady=2)
        self.date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        ttk.Entry(f, textvariable=self.date_var, width=15).grid(row=0, column=3, sticky="w", padx=5, pady=2)

        # Row 1: 业务经理
        ttk.Label(f, text="业务经理:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        mgr_names = [m["name"] for m in config.MANAGERS]
        self.manager_var = tk.StringVar(value=mgr_names[0] if mgr_names else "")
        self.manager_cb = ttk.Combobox(f, textvariable=self.manager_var, values=mgr_names, width=15)
        self.manager_cb.grid(row=1, column=1, sticky="w", padx=5, pady=2)
        self.manager_cb.bind("<<ComboboxSelected>>", self._on_manager)

        # Row 1: 电话 + 邮箱
        self.phone_var = tk.StringVar()
        ttk.Entry(f, textvariable=self.phone_var, width=15).grid(row=1, column=1, sticky="e", padx=5, pady=2)
        self.email_var = tk.StringVar()
        ttk.Entry(f, textvariable=self.email_var, width=20).grid(row=1, column=2, columnspan=2, sticky="w", padx=5, pady=2)
        self._fill_manager()

        # Row 2: 客户单位 (搜索)
        ttk.Label(f, text="客户单位:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.unit_var = tk.StringVar()
        self.unit_entry = ttk.Entry(f, textvariable=self.unit_var, width=30)
        self.unit_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=2)
        self.unit_entry.bind("<KeyRelease>", self._on_unit_type)
        self.unit_list = tk.Listbox(f, height=4, width=40)
        self.unit_list.grid(row=3, column=1, sticky="ew", padx=5, pady=2)
        self.unit_list.bind("<<ListboxSelect>>", self._on_unit_select)
        self.unit_list.bind("<Double-Button-1>", self._on_unit_dblclick)
        self._populate_unit_list("")

        # Row 2: 项目名称
        ttk.Label(f, text="项目名称:").grid(row=2, column=2, sticky="w", padx=5, pady=2)
        self.project_var = tk.StringVar()
        ttk.Entry(f, textvariable=self.project_var, width=20).grid(row=2, column=3, sticky="ew", padx=5, pady=2)

        f.columnconfigure(1, weight=1)

    def _on_manager(self, event=None):
        name = self.manager_var.get()
        for m in config.MANAGERS:
            if m["name"] == name:
                self.phone_var.set(m.get("phone", ""))
                self.email_var.set(m.get("email", ""))
                return

    def _fill_manager(self):
        name = self.manager_var.get()
        for m in config.MANAGERS:
            if m["name"] == name:
                self.phone_var.set(m.get("phone", ""))
                self.email_var.set(m.get("email", ""))
                return

    def _on_unit_type(self, event=None):
        q = self.unit_var.get()
        self._populate_unit_list(q)

    def _populate_unit_list(self, q):
        self.unit_list.delete(0, tk.END)
        results = search_units(q) if q else load_hist()[:15]
        for u in results:
            self.unit_list.insert(tk.END, u)

    def _on_unit_select(self, event=None):
        sel = self.unit_list.curselection()
        if sel:
            self.unit_var.set(self.unit_list.get(sel[0]))

    def _on_unit_dblclick(self, event=None):
        sel = self.unit_list.curselection()
        if sel:
            self.unit_var.set(self.unit_list.get(sel[0]))

    def get_values(self):
        return {
            "company": self.company_var.get(),
            "date": self.date_var.get(),
            "manager": self.manager_var.get(),
            "phone": self.phone_var.get(),
            "email": self.email_var.get(),
            "unit": self.unit_var.get(),
            "project": self.project_var.get(),
        }
