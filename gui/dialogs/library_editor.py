# gui/dialogs/library_editor.py — 库编辑弹窗（甲方/乙方/联系人）
import tkinter as tk
from tkinter import ttk, messagebox
from engines.library_manager import (
    load_party_b, save_party_b_entry, delete_party_b, list_party_b_names, get_party_b,
    load_private_customers, save_private_customer, delete_private_customer,
    list_party_a_names, resolve_party_a,
    load_contacts, save_contact_entry, append_contact_entry, delete_contact,
    contact_duplicate_exists, contact_duplicate_summary_by_side, cleanup_duplicate_contacts,
)
from gui.dialogs.party_editor import PartyEditor


class LibraryEditorDialog(tk.Toplevel):
    def __init__(self, parent, on_close=None):
        super().__init__(parent)
        self.title("库编辑")
        self.geometry("700x520")
        self.transient(parent)
        self.grab_set()
        self.on_close = on_close

        self._build()
        self._refresh_all()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"+{x}+{y}")

    def destroy(self):
        if self.on_close:
            self.on_close()
        super().destroy()

    def _build(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=5, pady=5)

        # 甲方标签页
        af = ttk.Frame(nb)
        nb.add(af, text="甲方库")
        self._build_list_page(af, "party_a")

        # 乙方标签页
        bf = ttk.Frame(nb)
        nb.add(bf, text="乙方库")
        self._build_list_page(bf, "party_b")

        # 联系人标签页
        cf = ttk.Frame(nb)
        nb.add(cf, text="联系人")
        self._build_list_page(cf, "contact")

        # 底部关闭按钮
        bbar = ttk.Frame(self)
        bbar.pack(fill="x", padx=5, pady=5)
        ttk.Button(bbar, text="关闭", command=self.destroy).pack(side="right", padx=5)

    def _build_list_page(self, parent, mode):
        """创建左侧列表 + 新增按钮布局"""
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=0)

        # 左侧 Listbox
        list_frame = ttk.Frame(parent)
        list_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        scroll = ttk.Scrollbar(list_frame, orient="vertical")
        scroll.grid(row=0, column=1, sticky="ns")

        lb = tk.Listbox(list_frame, yscrollcommand=scroll.set, font=("Microsoft YaHei UI", 10))
        lb.grid(row=0, column=0, sticky="nsew")
        scroll.configure(command=lb.yview)

        # 右侧按钮
        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=0, column=1, sticky="n", padx=5, pady=5)
        ttk.Button(btn_frame, text="新增", command=lambda: self._add_entry(mode), width=10).pack(pady=2)
        ttk.Button(btn_frame, text="编辑", command=lambda: self._edit_entry(mode), width=10).pack(pady=2)
        ttk.Button(btn_frame, text="删除", command=lambda: self._delete_entry(mode), width=10).pack(pady=2)
        if mode == "contact":
            ttk.Separator(btn_frame, orient="horizontal").pack(fill="x", pady=(8, 6))
            ttk.Button(btn_frame, text="清理重复", command=self._cleanup_duplicate_contacts, width=10).pack(pady=2)

        if mode == "contact":
            status_var = tk.StringVar(value="")
            status = ttk.Label(list_frame, textvariable=status_var, foreground="#6b7280")
            status.grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))
            self.contact_status_var = status_var

        # 双击编辑
        lb.bind("<Double-Button-1>", lambda e: self._edit_entry(mode))

        # 存储引用
        setattr(self, f"{mode}_listbox", lb)

    def _refresh_all(self):
        self._refresh_party_a()
        self._refresh_party_b()
        self._refresh_contacts()

    def _refresh_party_a(self):
        lb = getattr(self, "party_a_listbox", None)
        if not lb:
            return
        lb.delete(0, "end")
        names = list_party_a_names()
        for name, source in sorted(names.items(), key=lambda x: x[0]):
            prefix = "[私]" if source == "private" else "[公]"
            lb.insert("end", f"{prefix} {name}")

    def _refresh_party_b(self):
        lb = getattr(self, "party_b_listbox", None)
        if not lb:
            return
        lb.delete(0, "end")
        names = sorted(list_party_b_names())
        for name in names:
            lb.insert("end", name)

    def _refresh_contacts(self):
        lb = getattr(self, "contact_listbox", None)
        if not lb:
            return
        lb.delete(0, "end")
        _groups, hidden_count = contact_duplicate_summary_by_side()
        data = load_contacts()
        self._contact_entries = []
        for entry in data:
            unit = entry.get("unit_name", "")
            contact = entry.get("contact", "")
            display = f"{unit} — {contact}" if unit else contact
            lb.insert("end", display)
            self._contact_entries.append(entry)
        if hasattr(self, "contact_status_var"):
            text = f"已隐藏 {hidden_count} 条重复联系人" if hidden_count else ""
            self.contact_status_var.set(text)

    def _get_selected_contact_entry(self):
        lb = getattr(self, "contact_listbox", None)
        if not lb:
            return None
        sel = lb.curselection()
        if not sel:
            return None
        entries = getattr(self, "_contact_entries", [])
        index = sel[0]
        return entries[index] if index < len(entries) else None

    def _get_selected_name(self, mode):
        lb = getattr(self, f"{mode}_listbox", None)
        if not lb:
            return None, None
        sel = lb.curselection()
        if not sel:
            return None, None
        text = lb.get(sel[0])
        if mode == "party_a":
            # "[私] xxx" or "[公] xxx"
            if text.startswith("[私] "):
                return text[4:], "private"
            elif text.startswith("[公] "):
                return text[4:], "public"
            return text, "unknown"
        elif mode == "contact":
            # "unit — contact"
            return text, None
        return text, None

    # === 操作 ===

    def _add_entry(self, mode):
        PartyEditor(self, mode, data={}, on_save=lambda result: self._handle_save(mode, result, is_new=True))

    def _edit_entry(self, mode):
        name, source = self._get_selected_name(mode)
        if not name:
            messagebox.showwarning("提示", "请先选择一个条目")
            return
        if mode == "party_a":
            info = resolve_party_a(name)
            info["_old_name"] = name
            PartyEditor(self, mode, data=info, on_save=lambda result: self._handle_save(mode, result))
        elif mode == "party_b":
            info = get_party_b(name)
            info["_old_name"] = name
            PartyEditor(self, mode, data=info, on_save=lambda result: self._handle_save(mode, result))
        else:  # contact
            info = self._get_selected_contact_entry()
            if not info:
                messagebox.showwarning("提示", "请先选择一个条目")
                return
            info = dict(info)
            info["_old_name"] = f"{info.get('unit_name', '')}|{info.get('contact', '')}"
            PartyEditor(self, mode, data=info, on_save=lambda result: self._handle_save(mode, result))

    def _delete_entry(self, mode):
        name, source = self._get_selected_name(mode)
        if not name:
            messagebox.showwarning("提示", "请先选择一个条目")
            return
        if not messagebox.askyesno("确认", f"确定要删除「{name}」吗？", parent=self):
            return

        if mode == "party_a":
            if source == "public":
                messagebox.showinfo("提示", "公开库条目不可删除。\n如需隐藏，可编辑后保存为空白私有覆盖。", parent=self)
                return
            delete_private_customer(name)
        elif mode == "party_b":
            delete_party_b(name)
        else:
            entry = self._get_selected_contact_entry()
            if entry:
                delete_contact(entry.get("unit_name", ""), entry.get("contact", ""))
        self._refresh_all()

    def _handle_save(self, mode, result, is_new=False):
        if result.get("_delete"):
            # 移除私有覆盖
            if mode == "party_a":
                delete_private_customer(result.get("name", ""))
            self._refresh_all()
            return True

        if mode == "party_a":
            name = result.get("name", "")
            if not name:
                return False
            save_private_customer(name, result)
            messagebox.showinfo("提示", f"甲方「{name}」已保存到私有库", parent=self)
        elif mode == "party_b":
            name = result.get("name", "")
            if not name:
                return False
            save_party_b_entry(name, result)
            messagebox.showinfo("提示", f"乙方「{name}」已保存", parent=self)
        else:
            unit_name = result.get("unit_name", "")
            contact = result.get("contact", "")
            if not unit_name:
                messagebox.showwarning("提示", "请填写所属单位", parent=self)
                return False
            if not contact:
                messagebox.showwarning("提示", "请填写联系人姓名", parent=self)
                return False
            old_key = result.get("_old_name")
            if is_new and contact_duplicate_exists(result):
                ok = messagebox.askyesno(
                    "重复联系人",
                    f"该公司下已存在联系人【{contact}】，是否继续添加？",
                    parent=self,
                    default=messagebox.NO,
                )
                if not ok:
                    return False
                append_contact_entry(result)
            else:
                if not is_new and contact_duplicate_exists(result, old_key=old_key):
                    messagebox.showwarning("提示", "该联系人已存在，请勿重复保存。", parent=self)
                    return False
                save_contact_entry(result)
            messagebox.showinfo("提示", f"联系人「{contact}」已保存", parent=self)
        self._refresh_all()
        return True

    def _cleanup_duplicate_contacts(self):
        group_count, duplicate_count = contact_duplicate_summary_by_side()
        if duplicate_count <= 0:
            messagebox.showinfo("提示", "未发现重复联系人。", parent=self)
            return
        ok = messagebox.askyesno(
            "清理重复联系人",
            f"发现 {group_count} 组重复联系人，共 {duplicate_count} 条重复记录。\n\n是否保留每组信息更完整的一条，删除其余重复项？\n\n清理前会自动备份联系人库。",
            parent=self,
            default=messagebox.NO,
        )
        if not ok:
            return
        result = cleanup_duplicate_contacts()
        backups = "\n".join(result.get("backups", [])) or "无"
        messagebox.showinfo(
            "完成",
            f"已清理 {result.get('duplicates', 0)} 条重复联系人。\n\n备份文件：\n{backups}",
            parent=self,
        )
        self._refresh_all()
