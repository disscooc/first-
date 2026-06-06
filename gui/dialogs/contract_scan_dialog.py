# gui/dialogs/contract_scan_dialog.py -- 历史合同扫描结果确认
import tkinter as tk
from tkinter import ttk, messagebox
from dataclasses import asdict

from engines.contract_scanner import import_scan_results, save_scan_log


class ContractScanConfirmDialog(tk.Toplevel):
    def __init__(self, parent, summary, on_imported=None):
        super().__init__(parent)
        self.title("合同扫描结果确认")
        self.geometry("1080x660")
        self.transient(parent)
        self.grab_set()
        self.summary = summary
        self.on_imported = on_imported
        self.results = [asdict(item) for item in summary.get("results", [])]
        self._build()
        self._refresh_tree()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"+{x}+{y}")

    def _build(self):
        top = ttk.Frame(self, padding=(8, 6))
        top.pack(fill="x")
        ttk.Label(
            top,
            text=(
                f"扫描合同数量：{self.summary.get('scanned_count', 0)}    "
                f"成功识别：{self.summary.get('success_count', 0)}    "
                f"待确认：{self.summary.get('pending_count', 0)}    "
                f"低置信度：{self.summary.get('low_confidence_count', 0)}    "
                f"跳过文件：{self.summary.get('skipped_count', 0)}    "
                f"未识别：{self.summary.get('unrecognized_count', 0)}"
            ),
        ).pack(side="left")

        body = ttk.Frame(self, padding=(8, 0))
        body.pack(fill="both", expand=True)
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        columns = ("file", "party_a", "party_b", "a_contact", "a_phone", "b_contact", "b_phone", "confidence", "status", "operation")
        self.tree = ttk.Treeview(body, columns=columns, show="headings", height=12)
        headings = {
            "file": "文件名",
            "party_a": "甲方单位",
            "party_b": "乙方单位",
            "a_contact": "甲方联系人",
            "a_phone": "甲方电话",
            "b_contact": "乙方联系人",
            "b_phone": "乙方电话",
            "confidence": "识别置信度",
            "status": "识别状态",
            "operation": "操作",
        }
        widths = {
            "file": 220,
            "party_a": 180,
            "party_b": 180,
            "a_contact": 90,
            "a_phone": 110,
            "b_contact": 90,
            "b_phone": 110,
            "confidence": 90,
            "status": 90,
            "operation": 70,
        }
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], minwidth=widths[col], anchor="center" if col in ("confidence", "status", "operation") else "w")
        self.tree.grid(row=0, column=0, sticky="nsew")
        vbar = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        vbar.grid(row=0, column=1, sticky="ns")
        hbar = ttk.Scrollbar(body, orient="horizontal", command=self.tree.xview)
        hbar.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        self.tree.bind("<Double-Button-1>", lambda _event: self._edit_selected())
        self.tree.bind("<<TreeviewSelect>>", lambda _event: self._refresh_pending_info())

        pending_frame = ttk.LabelFrame(body, text="未确认信息")
        pending_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        self.pending_text = tk.Text(pending_frame, height=4, wrap="word", relief="flat", background="#FAFBFC")
        self.pending_text.pack(fill="x", expand=True, padx=6, pady=4)
        self.pending_text.configure(state="disabled")

        bottom = ttk.Frame(self, padding=8)
        bottom.pack(fill="x")
        ttk.Button(bottom, text="编辑选中", command=self._edit_selected, width=12).pack(side="left", padx=3)
        ttk.Button(bottom, text="跳过/恢复", command=self._toggle_selected, width=12).pack(side="left", padx=3)
        ttk.Button(bottom, text="取消", command=self.destroy, width=10).pack(side="right", padx=3)
        ttk.Button(bottom, text="确认入库", command=self._import, width=12).pack(side="right", padx=3)

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for index, item in enumerate(self.results):
            self.tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    item.get("file_name", ""),
                    item.get("party_a_name", ""),
                    item.get("party_b_name", ""),
                    item.get("party_a_contact", ""),
                    item.get("party_a_phone", ""),
                    item.get("party_b_contact", ""),
                    item.get("party_b_phone", ""),
                    item.get("confidence", 0),
                    item.get("recognition_status", "待确认"),
                    item.get("operation", "导入"),
                ),
            )
        if self.results and not self.tree.selection():
            self.tree.selection_set("0")
        self._refresh_pending_info()

    def _refresh_pending_info(self):
        text = "无"
        selection = self.tree.selection()
        if selection:
            item = self.results[int(selection[0])]
            text = item.get("pending_info", "") or "无"
        self.pending_text.configure(state="normal")
        self.pending_text.delete("1.0", "end")
        self.pending_text.insert("1.0", text)
        self.pending_text.configure(state="disabled")

    def _selected_index(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择一条扫描结果。", parent=self)
            return None
        return int(selection[0])

    def _toggle_selected(self):
        index = self._selected_index()
        if index is None:
            return
        current = self.results[index].get("operation", "导入")
        self.results[index]["operation"] = "跳过" if current == "导入" else "导入"
        self._refresh_tree()
        self.tree.selection_set(str(index))

    def _edit_selected(self):
        index = self._selected_index()
        if index is None:
            return
        ScanResultEditor(self, self.results[index], lambda updated: self._save_edit(index, updated))

    def _save_edit(self, index, updated):
        self.results[index].update(updated)
        self.results[index]["recognition_status"] = "已人工确认"
        self.results[index]["operation"] = "导入"
        self.results[index]["confidence"] = max(int(self.results[index].get("confidence") or 0), 70)
        self._refresh_tree()
        self.tree.selection_set(str(index))
        return True

    def _import(self):
        active = [item for item in self.results if item.get("operation") == "导入"]
        if not active:
            messagebox.showwarning("提示", "没有可入库的扫描结果。", parent=self)
            return
        if not messagebox.askyesno("确认入库", "入库前会自动备份当前库。\n确认将勾选结果写入库吗？", parent=self):
            return
        try:
            stats = import_scan_results(active)
            log_path = save_scan_log(self.summary, stats)
        except Exception as exc:
            messagebox.showerror("入库失败", str(exc), parent=self)
            return
        messagebox.showinfo(
            "完成",
            (
                f"入库完成。\n\n"
                f"新增/补充单位：{stats.get('imported_units', 0)}\n"
                f"新增/补充联系人：{stats.get('imported_contacts', 0)}\n"
                f"重复记录：{stats.get('duplicates', 0)}\n"
                f"备份目录：{stats.get('backup_dir', '')}\n"
                f"扫描日志：{log_path}"
            ),
            parent=self,
        )
        if self.on_imported:
            self.on_imported()
        self.destroy()


class ScanResultEditor(tk.Toplevel):
    FIELDS = [
        ("甲方单位", "party_a_name"),
        ("甲方注册地址", "party_a_address"),
        ("甲方信用代码", "party_a_credit"),
        ("甲方开户行", "party_a_bank"),
        ("甲方账号", "party_a_account"),
        ("甲方联系人", "party_a_contact"),
        ("甲方电话", "party_a_phone"),
        ("甲方邮箱", "party_a_email"),
        ("甲方通讯地址", "party_a_mailing_address"),
        ("乙方单位", "party_b_name"),
        ("乙方注册地址", "party_b_address"),
        ("乙方信用代码", "party_b_credit"),
        ("乙方开户行", "party_b_bank"),
        ("乙方账号", "party_b_account"),
        ("乙方联系人", "party_b_contact"),
        ("乙方电话", "party_b_phone"),
        ("乙方邮箱", "party_b_email"),
        ("乙方通讯地址", "party_b_mailing_address"),
    ]

    def __init__(self, parent, data, on_save):
        super().__init__(parent)
        self.title("编辑扫描结果")
        self.transient(parent)
        self.grab_set()
        self.data = data
        self.on_save = on_save
        self.vars = {}
        self._build()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"+{x}+{y}")

    def _build(self):
        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)
        for index, (label, key) in enumerate(self.FIELDS):
            ttk.Label(main, text=label + "：").grid(row=index, column=0, sticky="w", padx=4, pady=2)
            var = tk.StringVar(value=str(self.data.get(key, "") or ""))
            self.vars[key] = var
            ttk.Entry(main, textvariable=var, width=56).grid(row=index, column=1, sticky="ew", padx=4, pady=2)
        main.columnconfigure(1, weight=1)
        buttons = ttk.Frame(self, padding=8)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="取消", command=self.destroy, width=10).pack(side="right", padx=3)
        ttk.Button(buttons, text="保存", command=self._save, width=10).pack(side="right", padx=3)

    def _save(self):
        updated = {key: var.get().strip() for key, var in self.vars.items()}
        if self.on_save(updated) is not False:
            self.destroy()
