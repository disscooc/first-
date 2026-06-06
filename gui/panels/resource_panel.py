# gui/panels/resource_panel.py — 资源选择面板
import tkinter as tk
from tkinter import ttk
from engines.price_loader import load, search, find_price
from utils import fmt_num

class ResourcePanel(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, padding=5)
        self.app = app
        self.resources = []
        self._build()

    def _build(self):
        f = ttk.LabelFrame(self, text="资源库")
        f.pack(fill="both", expand=True, padx=5, pady=5)

        # Search bar
        sf = ttk.Frame(f)
        sf.pack(fill="x", padx=5, pady=3)
        ttk.Label(sf, text="搜索:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(sf, textvariable=self.search_var, width=25)
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<KeyRelease>", self._on_search)
        ttk.Button(sf, text="刷新价格库", command=self._reload).pack(side="right", padx=5)

        # Treeview
        cols = ("type", "name", "spec", "price")
        self.tree = ttk.Treeview(f, columns=cols, show="headings", height=8)
        self.tree.heading("type", text="类型")
        self.tree.heading("name", text="名称")
        self.tree.heading("spec", text="规格")
        self.tree.heading("price", text="单价(共享)")
        self.tree.column("type", width=60)
        self.tree.column("name", width=120)
        self.tree.column("spec", width=100)
        self.tree.column("price", width=80)
        self.tree.pack(fill="both", expand=True, padx=5, pady=3)

        # Scrollbar
        sb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)

        # Action buttons
        bf = ttk.Frame(f)
        bf.pack(fill="x", padx=5, pady=3)
        ttk.Button(bf, text="添加选中 →", command=self._add_selected).pack(side="right", padx=5)
        ttk.Button(bf, text="添加存储", command=self._add_storage).pack(side="right", padx=5)

        self._reload()

    def _reload(self):
        self.resources = load()
        self._refresh_tree()

    def _on_search(self, event=None):
        self._refresh_tree()

    def _refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        q = self.search_var.get()
        results = search(self.resources, q)
        for r in results:
            self.tree.insert("", tk.END, values=(r.rtype, r.name, r.spec, fmt_num(r.shared_price, "0.######")))

    def _add_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        q = self.search_var.get()
        results = search(self.resources, q)
        for item in sel:
            idx = self.tree.index(item)
            if idx < len(results):
                r = results[idx]
                self.app.quote_panel.add_item(r)

    def _add_storage(self):
        from models import QuoteItem
        item = QuoteItem(
            item_type="Storage", rtype="存储", name="分布式文件存储",
            spec="免费赠送0.5T", orig_price=0, price=0, discount=1.0,
            quantity=0.5, amount=0, free_quota=True
        )
        self.app.quote_panel.add_storage(item)
