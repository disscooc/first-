# gui/dialogs/party_editor.py — 甲方/乙方/联系人编辑表单
import tkinter as tk
from tkinter import ttk, messagebox


class PartyEditor(tk.Toplevel):
    """可复用的编辑表单：
    - mode: 'party_a' | 'party_b' | 'contact'
    - data: 现有数据 dict（编辑模式）或空（新增模式）
    - on_save: callback(name, data)
    """

    def __init__(self, parent, mode, data=None, on_save=None):
        super().__init__(parent)
        self.mode = mode
        self.data = data or {}
        self.on_save = on_save
        self.result = None

        titles = {"party_a": "编辑甲方", "party_b": "编辑乙方", "contact": "编辑联系人"}
        self.title(titles.get(mode, "编辑"))
        self.transient(parent)
        self.grab_set()

        self._build_form()
        self._load_data()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"+{x}+{y}")

    def _build_form(self):
        main = ttk.Frame(self, padding=15)
        main.pack(fill="both", expand=True)

        if self.mode == "party_a":
            fields = [
                ("单位名称:", "name", 35),
                ("注册地址/电话:", "address", 45),
                ("统一社会信用代码:", "credit", 25),
                ("开户银行:", "bank", 35),
                ("开户账号:", "account", 25),
            ]
        elif self.mode == "party_b":
            fields = [
                ("公司名称:", "name", 35),
                ("注册地址/电话:", "address", 45),
                ("统一社会信用代码:", "credit", 25),
                ("开户银行:", "bank", 35),
                ("开户账号:", "account", 25),
            ]
        else:  # contact
            fields = [
                ("所属单位:", "unit_name", 35),
                ("联系人:", "contact", 15),
                ("电话:", "phone", 15),
                ("邮箱:", "email", 25),
                ("备注/来源:", "note", 45),
            ]

        self.vars = {}
        for i, (label, key, width) in enumerate(fields):
            ttk.Label(main, text=label).grid(row=i, column=0, sticky="w", padx=5, pady=3)
            var = tk.StringVar()
            self.vars[key] = var
            ttk.Entry(main, textvariable=var, width=width).grid(row=i, column=1, padx=5, pady=3, sticky="ew")

        main.columnconfigure(1, weight=1)

        # 来源提示
        source = self.data.get("source", "")
        if source == "public" and self.mode == "party_a":
            note = ttk.Label(main, text="⚠ 这是公开库数据，保存后将写入私有库作为覆盖项", foreground="#e67e22")
            note.grid(row=len(fields), column=0, columnspan=2, sticky="w", padx=5, pady=(10, 0))

        # 按钮
        btn_frame = ttk.Frame(self, padding=10)
        btn_frame.pack(fill="x", padx=5, pady=5)
        ttk.Button(btn_frame, text="取消", command=self.destroy).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="保存", command=self._on_save).pack(side="right", padx=5)

        if self.mode == "party_a" and source == "public":
            del_btn = ttk.Button(btn_frame, text="删除（仅移除私有覆盖）", command=self._on_delete)
            del_btn.pack(side="left", padx=5)

    def _load_data(self):
        if not self.data:
            return
        for key, var in self.vars.items():
            value = self.data.get(key, "")
            if not value and self.mode == "party_a":
                # 尝试转换公开库字段名
                pass
            var.set(str(value) if value else "")

    def _on_save(self):
        result = {}
        for key, var in self.vars.items():
            result[key] = var.get().strip()

        # 传递旧名称（支持改名场景）
        if self.data.get("_old_name"):
            result["_old_name"] = self.data["_old_name"]

        # 验证必填
        name_key = "name" if self.mode != "contact" else "unit_name"
        if not result.get(name_key):
            messagebox.showwarning("提示", "请填写名称", parent=self)
            return

        if self.on_save:
            saved = self.on_save(result)
            if saved is False:
                return
        self.result = result
        self.destroy()

    def _on_delete(self):
        name = self.data.get("name", "")
        if not name:
            return
        if messagebox.askyesno("确认", f"确定要移除「{name}」的私有覆盖项吗？\n公开库原始数据不会受影响。", parent=self):
            if self.on_save:
                self.on_save({"_delete": True, "name": name})
            self.destroy()
