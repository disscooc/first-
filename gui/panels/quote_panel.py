# gui/panels/quote_panel.py — 配置+定价+输出面板
import tkinter as tk
from tkinter import ttk, messagebox
from models import QuoteItem

class QuotePanel(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, padding=5)
        self.app = app
        self.comp_items = []
        self.stor_items = []
        self._build()

    def _build(self):
        # === 报价模式 ===
        mf = ttk.LabelFrame(self, text="报价模式")
        mf.pack(fill="x", padx=5, pady=3)
        self.mode_var = tk.StringVar(value="unit_price")
        ttk.Radiobutton(mf, text="单价模式", variable=self.mode_var, value="unit_price", command=self._on_mode).pack(side="left", padx=10, pady=3)
        ttk.Radiobutton(mf, text="数量模式", variable=self.mode_var, value="quantity", command=self._on_mode).pack(side="left", padx=10, pady=3)

        # === 单价 & 数量 ===
        pf = ttk.Frame(mf)
        pf.pack(fill="x", padx=5, pady=3)
        ttk.Label(pf, text="单价:").pack(side="left")
        self.price_var = tk.StringVar()
        ttk.Entry(pf, textvariable=self.price_var, width=12).pack(side="left", padx=3)
        ttk.Label(pf, text="数量:").pack(side="left", padx=(20, 0))
        self.qty_var = tk.StringVar()
        self.qty_entry = ttk.Entry(pf, textvariable=self.qty_var, width=12)
        self.qty_entry.pack(side="left", padx=3)
        self.price_var.trace_add("write", self._calc)
        self.qty_var.trace_add("write", self._calc)
        self._on_mode()

        # === 折扣 ===
        df = ttk.LabelFrame(self, text="折扣")
        df.pack(fill="x", padx=5, pady=3)
        ttk.Label(df, text="折扣率:").grid(row=0, column=0, padx=5, pady=2)
        self.discount_var = tk.StringVar(value="1.0")
        ttk.Entry(df, textvariable=self.discount_var, width=8).grid(row=0, column=1, padx=5, pady=2)
        ttk.Label(df, text="(如0.3=3折)").grid(row=0, column=2, padx=5, pady=2)
        ttk.Label(df, text="折后单价:").grid(row=0, column=3, padx=5, pady=2)
        self.final_price_var = tk.StringVar()
        ttk.Entry(df, textvariable=self.final_price_var, width=10, state="readonly").grid(row=0, column=4, padx=5, pady=2)
        self.discount_var.trace_add("write", self._on_discount)

        # === 当前项操作 ===
        bf = ttk.Frame(self)
        bf.pack(fill="x", padx=5, pady=3)
        ttk.Button(bf, text="+ 添加资源", command=self._add_current).pack(side="left", padx=3)
        ttk.Button(bf, text="删除选中", command=self._del_selected).pack(side="left", padx=3)
        ttk.Button(bf, text="清空列表", command=self._clear_all).pack(side="left", padx=3)

        # === 已选资源列表 ===
        self.list_label = ttk.Label(self, text="已选资源 (0项)")
        self.list_label.pack(anchor="w", padx=5)
        cols = ("type", "name", "spec", "price", "qty", "amount")
        self.list_tree = ttk.Treeview(self, columns=cols, show="headings", height=6)
        self.list_tree.heading("type", text="类型")
        self.list_tree.heading("name", text="名称")
        self.list_tree.heading("spec", text="规格")
        self.list_tree.heading("price", text="单价")
        self.list_tree.heading("qty", text="数量")
        self.list_tree.heading("amount", text="金额")
        self.list_tree.column("type", width=50)
        self.list_tree.column("name", width=100)
        self.list_tree.column("spec", width=80)
        self.list_tree.column("price", width=70)
        self.list_tree.column("qty", width=60)
        self.list_tree.column("amount", width=80)
        self.list_tree.pack(fill="both", expand=True, padx=5, pady=3)

        # === 赠送存储 ===
        gsf = ttk.LabelFrame(self, text="赠送存储")
        gsf.pack(fill="x", padx=5, pady=3)
        ttk.Label(gsf, text="赠送容量(TB):").pack(side="left", padx=5)
        self.storage_var = tk.StringVar(value="")
        ttk.Entry(gsf, textvariable=self.storage_var, width=8).pack(side="left", padx=5)
        ttk.Button(gsf, text="更新赠送", command=self._update_storage).pack(side="left", padx=5)

        # === 输出选项 ===
        of = ttk.LabelFrame(self, text="输出选项")
        of.pack(fill="x", padx=5, pady=3)
        self.out_quote = tk.BooleanVar(value=True)
        self.out_contract = tk.BooleanVar(value=False)
        ttk.Checkbutton(of, text="标准报价单(Excel)", variable=self.out_quote).grid(row=0, column=0, sticky="w", padx=5)
        ttk.Checkbutton(of, text="合同(Word)", variable=self.out_contract).grid(row=0, column=1, sticky="w", padx=5)
        self.out_contract.trace_add("write", lambda *a: self.app._toggle_contract())

        # === 生成按钮 ===
        genf = ttk.Frame(self)
        genf.pack(fill="x", padx=5, pady=5)
        ttk.Button(genf, text="生成三方报价单", command=self.app.generate_tp).pack(side="right", padx=5)
        ttk.Button(genf, text="一键生成", command=self.app.generate).pack(side="right", padx=5)

    def _on_mode(self):
        mode = self.mode_var.get()
        if mode == "unit_price":
            self.qty_entry.configure(state="disabled")
            self.qty_var.set("")
        else:
            self.qty_entry.configure(state="normal")

    def _calc(self, *args):
        mode = self.mode_var.get()
        try:
            p = float(self.price_var.get()) if self.price_var.get() else 0
        except:
            p = 0
        if mode == "quantity":
            try:
                q = float(self.qty_var.get()) if self.qty_var.get() else 0
            except:
                q = 0

    def _on_discount(self, *args):
        try:
            disc = float(self.discount_var.get())
            price = float(self.price_var.get()) if self.price_var.get() else 0
            self.final_price_var.set(f"{price * disc:.6f}")
        except:
            self.final_price_var.set("")

    def _add_current(self):
        try:
            price = float(self.price_var.get()) if self.price_var.get() else 0
            disc = float(self.discount_var.get()) if self.discount_var.get() else 1.0
            final_price = price * disc
            qty = float(self.qty_var.get()) if self.qty_var.get() else 0
            amount = final_price * qty
        except:
            messagebox.showwarning("输入错误", "请填写有效的单价")
            return
        item = QuoteItem(
            item_type="Compute", rtype="", name="手动输入",
            spec="", orig_price=price, discount=disc, price=final_price,
            quantity=qty, amount=amount
        )
        self.comp_items.append(item)
        self._refresh_list()

    def add_item(self, resource):
        """从资源库添加"""
        qty = float(self.qty_var.get()) if self.qty_var.get() else 0
        disc = float(self.discount_var.get()) if self.discount_var.get() else 1.0
        price = resource.shared_price
        final_price = price * disc
        amount = final_price * qty
        item = QuoteItem(
            item_type="Compute", rtype=resource.rtype, name=resource.name,
            spec=resource.spec, orig_price=price, discount=disc, price=final_price,
            quantity=qty, amount=amount
        )
        self.comp_items.append(item)
        self._refresh_list()

    def add_storage(self, item):
        self.stor_items.append(item)
        self._refresh_list()

    def _del_selected(self):
        sel = self.list_tree.selection()
        for s in sel:
            idx = self.list_tree.index(s)
            if idx < len(self.comp_items):
                del self.comp_items[idx]
            else:
                si = idx - len(self.comp_items)
                if si < len(self.stor_items):
                    del self.stor_items[si]
        self._refresh_list()

    def _clear_all(self):
        self.comp_items.clear()
        self.stor_items.clear()
        self._refresh_list()

    def _refresh_list(self):
        for item in self.list_tree.get_children():
            self.list_tree.delete(item)
        for it in self.comp_items:
            self.list_tree.insert("", tk.END, values=(it.rtype, it.name, it.spec, f"{it.price:.6f}", f"{it.quantity:.2f}", f"{it.amount:.2f}"))
        for it in self.stor_items:
            self.list_tree.insert("", tk.END, values=("存储", it.name, it.spec, f"{it.price:.2f}", f"{it.quantity:.2f}", f"{it.amount:.2f}"))
        total = len(self.comp_items) + len(self.stor_items)
        self.list_label.configure(text=f"已选资源 ({total}项)")

    def _update_storage(self):
        try:
            tb = float(self.storage_var.get()) if self.storage_var.get() else 0
        except:
            tb = 0
        self.stor_items.clear()
        if tb > 0:
            from models import QuoteItem
            free = 0.5
            item = QuoteItem(
                item_type="Storage", rtype="存储", name="分布式文件存储",
                spec=f"免费赠送{tb}T", orig_price=0, price=0, discount=1.0,
                quantity=tb, amount=0, free_quota=(tb <= free)
            )
            self.stor_items.append(item)
        self._refresh_list()

    def get_items(self):
        return self.comp_items, self.stor_items
