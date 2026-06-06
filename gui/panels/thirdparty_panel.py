# gui/panels/thirdparty_panel.py — 三方报价设置弹窗
import tkinter as tk
from tkinter import ttk
from engines.competitor import comp_items

class ThirdPartyDialog(tk.Toplevel):
    def __init__(self, parent, comp_items_list, stor_items_list, is_total, callback):
        super().__init__(parent)
        self.title("三方报价设置")
        self.geometry("480x320")
        self.callback = callback
        self.transient(parent)
        self.grab_set()

        self.comp = comp_items_list
        self.stor = stor_items_list
        self.is_total = is_total

        f = ttk.LabelFrame(self, text="三方报价参数", padding=10)
        f.pack(fill="both", expand=True, padx=10, pady=10)

        # 三方公司
        ttk.Label(f, text="三方公司名称:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.company_var = tk.StringVar(value="公司A")
        ttk.Entry(f, textvariable=self.company_var, width=25).grid(row=0, column=1, padx=5, pady=5)

        # 加价率
        ttk.Label(f, text="加价率:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.rate_var = tk.StringVar(value="0.3")
        ttk.Entry(f, textvariable=self.rate_var, width=8).grid(row=1, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(f, text="(0.3 = 加价30%)").grid(row=1, column=1, padx=(80, 0), pady=5, sticky="w")

        # 模式选择
        ttk.Label(f, text="三方模式:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.rule_var = tk.StringVar(value="MarkupPrice")
        rf = ttk.Frame(f)
        rf.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        ttk.Radiobutton(rf, text="同资源提价(数量不变)", variable=self.rule_var, value="MarkupPrice").pack(anchor="w")
        ttk.Radiobutton(rf, text="总价不变增数量", variable=self.rule_var, value="SameTotalMoreQty").pack(anchor="w")

        # 模板选择
        ttk.Label(f, text="输出模板:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.layout_var = tk.StringVar(value="StandardThirdParty")
        lf = ttk.Frame(f)
        lf.grid(row=3, column=1, sticky="w", padx=5, pady=5)
        ttk.Radiobutton(lf, text="标准模板", variable=self.layout_var, value="StandardThirdParty").pack(anchor="w")
        ttk.Radiobutton(lf, text="云达模板", variable=self.layout_var, value="YundaThirdParty").pack(anchor="w")

        # 按钮
        bf = ttk.Frame(self)
        bf.pack(fill="x", padx=10, pady=10)
        ttk.Button(bf, text="取消", command=self.destroy).pack(side="right", padx=5)
        ttk.Button(bf, text="生成", command=self._on_generate).pack(side="right", padx=5)

    def _on_generate(self):
        try:
            rate = float(self.rate_var.get())
        except:
            rate = 0.3
        rule = self.rule_var.get()
        layout = self.layout_var.get()

        adjusted = comp_items(self.comp, rate, rule, self.is_total)
        self.callback({
            "company": self.company_var.get(),
            "rate": rate,
            "rule": rule,
            "layout": layout,
            "comp": adjusted,
            "stor": self.stor,
        })
        self.destroy()
