# gui/dialogs/contract_dialog.py — 合同信息弹窗
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import config
from utils import read_json
from gui.autocomplete import AutocompleteInput
from engines.contract_writer import load_customer
from engines.library_manager import (
    list_party_b_names, get_party_b,
    list_party_a_names, resolve_party_a,
    load_contacts, load_private_customers,
    save_contact_entry,
)


class ContractDialog(tk.Toplevel):
    def __init__(self, parent, unit_name, callback, existing_info=None):
        super().__init__(parent)
        self.title("合同信息")
        self.geometry("720x680")
        self.callback = callback
        self.transient(parent)
        self.grab_set()
        self._party_a_options_cache = None
        self._party_b_names_cache = None
        self._selected_contacts = {}

        # 顶部工具栏
        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", padx=5, pady=(5, 0))
        ttk.Button(toolbar, text="库编辑", command=self._open_library_editor, width=10).pack(side="left", padx=2)
        ttk.Button(toolbar, text="扫描历史合同", command=self._scan_history_contracts, width=14).pack(side="left", padx=2)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=5, pady=5)

        # === 4个标签页 ===
        af = ttk.Frame(nb)
        nb.add(af, text="甲方公司信息")
        self._party_company_frame(af, "party_a", "甲方公司")

        bf = ttk.Frame(nb)
        nb.add(bf, text="乙方公司信息")
        self._party_company_frame(bf, "party_b", "乙方公司")

        cf = ttk.Frame(nb)
        nb.add(cf, text="联系人信息")
        self._contact_frame(cf)

        sf = ttk.Frame(nb)
        nb.add(sf, text="签约信息")
        self._sign_frame(sf)

        # 底部按钮
        bbar = ttk.Frame(self)
        bbar.pack(fill="x", padx=5, pady=5)
        ttk.Button(bbar, text="取消", command=self.destroy).pack(side="right", padx=5)
        ttk.Button(bbar, text="确定", command=self._on_ok).pack(side="right", padx=5)

        # 加载已有客户：优先使用内存中的 existing_info
        if existing_info:
            cust = existing_info
        elif unit_name:
            cust = load_customer(unit_name)
        else:
            cust = {}
        self._load_defaults(cust)

        # 自动匹配联系人库：如果乙方联系人姓名已在库中存在，自动填入通讯地址/电话/邮箱
        self._auto_fill_contact("b")
        # 同理处理甲方联系人
        self._auto_fill_contact("a")

    # ======================== 公司信息页 ========================
    def _party_company_frame(self, parent, prefix, title):
        f = ttk.LabelFrame(parent, text=title, padding=10)
        f.pack(fill="x", padx=5, pady=5)

        fields = [
            ("单位名称:", "name", 35, True),
            ("注册地址:", "address", 45, False),
            ("信用代码:", "credit", 28, False),
            ("开户行:", "bank", 35, False),
            ("账号:", "account", 28, False),
        ]
        for i, (label, key, width, is_name) in enumerate(fields):
            ttk.Label(f, text=label).grid(row=i, column=0, sticky="w", padx=5, pady=3)
            var = tk.StringVar()
            setattr(self, f"{prefix}_{key}_var", var)

            if is_name:
                field = AutocompleteInput(
                    f,
                    var,
                    lambda text, p=prefix: self._company_suggestions(p, text),
                    get_all_suggestions=lambda p=prefix: self._company_suggestions(p, "", include_all=True),
                    on_select=lambda _value, p=prefix: self._on_company_selected(p),
                    width=width,
                )
                field.grid(row=i, column=1, padx=5, pady=3, sticky="ew")
                setattr(self, f"{prefix}_name_input", field)
            else:
                ttk.Entry(f, textvariable=var, width=width).grid(
                    row=i, column=1, padx=5, pady=3, sticky="ew")

        f.columnconfigure(1, weight=1)

    def _filter_company_combo(self, prefix):
        text = getattr(self, f"{prefix}_name_var").get()
        return self._company_suggestions(prefix, text)

    def _company_suggestions(self, prefix, text, include_all=False):
        query = text.strip().lower()
        if not query and not include_all:
            return []
        if prefix == "party_b":
            names = self._party_b_names()
            if not query:
                return names
            return [name for name in names if self._match_text(query, [name])][:30]
        options = self._party_a_options()
        if not query:
            return [item["name"] for item in options]
        return [
            item["name"]
            for item in options
            if self._match_text(query, [item["name"], *item.get("aliases", [])])
        ][:30]

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
            self._party_b_names_cache = list(dict.fromkeys(list_party_b_names()))
        return self._party_b_names_cache

    def _match_text(self, query, fields):
        return any(query in str(field or "").lower() for field in fields)

    def _on_company_selected(self, prefix):
        """选中公司名称后，自动带出所有公司信息"""
        name = getattr(self, f"{prefix}_name_var").get()
        if prefix == "party_b":
            info = get_party_b(name)
            if info:
                self.party_b_address_var.set(info.get("address", ""))
                self.party_b_credit_var.set(info.get("credit", ""))
                self.party_b_bank_var.set(info.get("bank", ""))
                self.party_b_account_var.set(info.get("account", ""))
        else:  # party_a
            info = resolve_party_a(name)
            if info.get("source") == "new":
                return
            self.party_a_address_var.set(info.get("address", ""))
            self.party_a_credit_var.set(info.get("credit", ""))
            self.party_a_bank_var.set(info.get("bank", ""))
            self.party_a_account_var.set(info.get("account", ""))

    # ======================== 联系人信息页（合并甲方+乙方） ========================
    def _contact_frame(self, parent):
        f = ttk.Frame(parent, padding=10)
        f.pack(fill="both", expand=True, padx=5, pady=5)
        f.columnconfigure(1, weight=1)

        self._contact_inputs = {}

        # 加载联系人库（用于联想）
        contacts = load_contacts()
        self._all_contacts = contacts

        # --- 甲方联系人 ---
        af = ttk.LabelFrame(f, text="甲方联系人", padding=8)
        af.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        af.columnconfigure(1, weight=1)

        ttk.Label(af, text="联系人:").grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.party_a_contact_var = tk.StringVar()
        a_input = AutocompleteInput(
            af,
            self.party_a_contact_var,
            lambda text: self._contact_suggestions("a", text),
            get_all_suggestions=lambda: self._contact_suggestions("a", "", include_all=True),
            on_select_item=lambda item: self._on_contact_item_selected("a", item),
            width=20,
        )
        a_input.grid(row=0, column=1, padx=5, pady=3, sticky="ew")
        self._contact_inputs["a"] = a_input
        self._refresh_contact_combo("a")

        ttk.Label(af, text="通讯地址:").grid(row=1, column=0, sticky="w", padx=5, pady=3)
        self.party_a_mailing_address_var = tk.StringVar()
        ttk.Entry(af, textvariable=self.party_a_mailing_address_var, width=45).grid(
            row=1, column=1, padx=5, pady=3, sticky="ew")

        ttk.Label(af, text="电话:").grid(row=2, column=0, sticky="w", padx=5, pady=3)
        self.party_a_phone_var = tk.StringVar()
        ttk.Entry(af, textvariable=self.party_a_phone_var, width=25).grid(
            row=2, column=1, padx=5, pady=3, sticky="w")

        ttk.Label(af, text="邮箱:").grid(row=3, column=0, sticky="w", padx=5, pady=3)
        self.party_a_email_var = tk.StringVar()
        ttk.Entry(af, textvariable=self.party_a_email_var, width=35).grid(
            row=3, column=1, padx=5, pady=3, sticky="w")

        # --- 乙方联系人 ---
        bf = ttk.LabelFrame(f, text="乙方联系人", padding=8)
        bf.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        bf.columnconfigure(1, weight=1)

        ttk.Label(bf, text="联系人:").grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.party_b_contact_var = tk.StringVar()
        b_input = AutocompleteInput(
            bf,
            self.party_b_contact_var,
            lambda text: self._contact_suggestions("b", text),
            get_all_suggestions=lambda: self._contact_suggestions("b", "", include_all=True),
            on_select_item=lambda item: self._on_contact_item_selected("b", item),
            width=20,
        )
        b_input.grid(row=0, column=1, padx=5, pady=3, sticky="ew")
        self._contact_inputs["b"] = b_input
        self._refresh_contact_combo("b")

        ttk.Label(bf, text="通讯地址:").grid(row=1, column=0, sticky="w", padx=5, pady=3)
        self.party_b_mailing_address_var = tk.StringVar()
        ttk.Entry(bf, textvariable=self.party_b_mailing_address_var, width=45).grid(
            row=1, column=1, padx=5, pady=3, sticky="ew")

        ttk.Label(bf, text="电话:").grid(row=2, column=0, sticky="w", padx=5, pady=3)
        self.party_b_phone_var = tk.StringVar()
        ttk.Entry(bf, textvariable=self.party_b_phone_var, width=25).grid(
            row=2, column=1, padx=5, pady=3, sticky="w")

        ttk.Label(bf, text="邮箱:").grid(row=3, column=0, sticky="w", padx=5, pady=3)
        self.party_b_email_var = tk.StringVar()
        ttk.Entry(bf, textvariable=self.party_b_email_var, width=35).grid(
            row=3, column=1, padx=5, pady=3, sticky="w")

    def _refresh_contact_combo(self, prefix):
        """刷新联系人候选（候选由当前甲/乙方公司动态决定）。"""
        return self._contact_suggestions(prefix, "", include_all=True)

    def _all_contact_names(self):
        """兼容旧调用：返回全部联系人候选对象。"""
        return self._contact_suggestions("a", "", include_all=True) + self._contact_suggestions("b", "", include_all=True)

    def _filter_contact_combo(self, prefix):
        """联系人输入框实时模糊匹配（防抖：仅处理长度>=1的输入）"""
        text = getattr(self, f"party_{prefix}_contact_var").get()
        return self._contact_suggestions(prefix, text)

    def _contact_suggestions(self, prefix, text, include_all=False):
        """联系人输入框实时模糊匹配（防抖：仅处理长度>=1的输入）"""
        query = text.strip().lower()
        if not query and not include_all:
            return []
        contacts = self._ordered_contacts(prefix)
        results = []
        seen = set()
        for c in contacts:
            name = c.get("contact", "").strip()
            if not name:
                continue
            if query and not self._match_text(
                query,
                [name, c.get("phone", ""), c.get("email", ""), c.get("unit_name", "")],
            ):
                continue
            key = (c.get("unit_name", "").strip(), name)
            if key in seen:
                continue
            seen.add(key)
            results.append({"label": self._contact_label(c), "value": name, "raw": c})
            if len(results) >= 30:
                break
        return results

    def _ordered_contacts(self, prefix):
        """按公司名称优先排序联系人：同公司 > 同侧(甲/乙方)"""
        current_unit = getattr(self, f"party_{prefix}_name_var").get().strip()
        party_b_names = set(self._party_b_names())
        same_company = []
        same_side = []
        for contact in getattr(self, "_all_contacts", []):
            unit = contact.get("unit_name", "").strip()
            is_party_b = unit in party_b_names
            if prefix == "b":
                side_match = is_party_b
            else:
                side_match = not is_party_b
            if current_unit and unit == current_unit:
                # 精确匹配当前选中公司 → 最高优先
                same_company.append(contact)
            elif side_match and not current_unit:
                # 未指定公司时，显示同侧所有联系人
                same_side.append(contact)
            elif side_match and current_unit:
                # 已指定公司但此联系人不属于该公司 → 降低优先级
                same_side.append(contact)
        # 有明确选中公司时：只展示该公司联系人 + 其他同侧作为备选
        if current_unit and same_company:
            return self._sort_contacts(same_company) + self._sort_contacts(same_side)[:5]
        return self._sort_contacts(same_company) + self._sort_contacts(same_side)

    def _sort_contacts(self, contacts):
        return sorted(contacts, key=lambda item: item.get("updated_at", ""), reverse=True)

    def _contact_label(self, contact):
        parts = [
            contact.get("contact", "").strip(),
            contact.get("phone", "").strip(),
            contact.get("unit_name", "").strip(),
        ]
        return "｜".join(part for part in parts if part)

    def _on_contact_item_selected(self, prefix, item):
        self._selected_contacts[prefix] = item if isinstance(item, dict) else {}
        self._fill_contact_fields(prefix, self._selected_contacts[prefix])

    def _on_contact_selected(self, prefix):
        """选中联系人后，自动填充通讯地址、电话和邮箱"""
        contact_name = getattr(self, f"party_{prefix}_contact_var").get()
        for c in self._ordered_contacts(prefix):
            if c.get("contact", "") == contact_name:
                self._fill_contact_fields(prefix, c)
                break

    def _fill_contact_fields(self, prefix, contact):
        if not contact:
            return
        getattr(self, f"party_{prefix}_mailing_address_var").set(contact.get("mailing_address", ""))
        getattr(self, f"party_{prefix}_phone_var").set(contact.get("phone", ""))
        getattr(self, f"party_{prefix}_email_var").set(contact.get("email", ""))

    # ======================== 签约信息页 ========================
    def _sign_frame(self, parent):
        f = ttk.LabelFrame(parent, text="签约信息", padding=10)
        f.pack(fill="x", padx=5, pady=5)
        fields = [
            ("签约地点:", "sign_place", 20),
            ("签约日期:", "sign_time", 15),
            ("服务期限:", "service_term", 15),
            ("付款天数:", "pay_days", 10),
            ("发票类型:", "invoice", 25),
        ]
        for i, (label, key, width) in enumerate(fields):
            ttk.Label(f, text=label).grid(row=i, column=0, sticky="w", padx=5, pady=2)
            var = tk.StringVar()
            setattr(self, f"sign_{key}_var", var)
            ttk.Entry(f, textvariable=var, width=width).grid(
                row=i, column=1, padx=5, pady=2, sticky="w")

    def _load_defaults(self, cust):
        # 甲方公司默认从客户档案加载；只有名称时再从甲方公司库补齐。
        party_a_name = cust.get("party_a_name", "")
        party_a_info = resolve_party_a(party_a_name) if party_a_name else {}
        for key in ["name", "address", "credit", "bank", "account"]:
            var = getattr(self, f"party_a_{key}_var", None)
            if var:
                var.set(cust.get(f"party_a_{key}", "") or party_a_info.get(key, ""))

        # 乙方公司默认；优先使用主界面传入的报价单位，再从乙方公司库补齐。
        d = config.DEF_CONTRACT
        party_b_name = cust.get("party_b_name", "") or d.get("party_b_name", "")
        party_b_info = get_party_b(party_b_name) if party_b_name else {}
        for key in ["name", "address", "credit", "bank", "account"]:
            var = getattr(self, f"party_b_{key}_var", None)
            if var:
                val = cust.get(f"party_b_{key}", "") or party_b_info.get(key, "") or d.get(f"party_b_{key}", "")
                if key == "address" and not val:
                    val = d.get("party_b_addr", "")
                var.set(val)

        # 甲方联系人默认值
        var = getattr(self, "party_a_contact_var", None)
        if var:
            var.set(cust.get("party_a_contact", ""))
        var = getattr(self, "party_a_mailing_address_var", None)
        if var:
            var.set(cust.get("party_a_mailing_address", ""))
        var = getattr(self, "party_a_phone_var", None)
        if var:
            var.set(cust.get("party_a_phone", ""))
        var = getattr(self, "party_a_email_var", None)
        if var:
            var.set(cust.get("party_a_email", ""))

        # 乙方联系人默认值
        var = getattr(self, "party_b_contact_var", None)
        if var:
            var.set(cust.get("party_b_contact", ""))
        var = getattr(self, "party_b_mailing_address_var", None)
        if var:
            var.set(cust.get("party_b_mailing_address", ""))
        var = getattr(self, "party_b_phone_var", None)
        if var:
            var.set(cust.get("party_b_phone", ""))
        var = getattr(self, "party_b_email_var", None)
        if var:
            var.set(cust.get("party_b_email", ""))

        # 签约信息默认
        for key in ["sign_place", "sign_time", "service_term", "pay_days", "invoice"]:
            var = getattr(self, f"sign_{key}_var", None)
            if var:
                default = d.get(key, "")
                if key == "sign_time" and not default:
                    default = datetime.now().strftime("%Y.%m.%d")
                var.set(default)

    def _auto_fill_contact(self, prefix):
        """根据当前联系人姓名，自动从联系人库匹配并填入通讯地址/邮箱（电话保留首页填入的值）"""
        var = getattr(self, f"party_{prefix}_contact_var", None)
        if not var:
            return
        contact_name = var.get().strip()
        if not contact_name:
            return
        for c in self._ordered_contacts(prefix):
            if c.get("contact", "") == contact_name:
                getattr(self, f"party_{prefix}_mailing_address_var").set(c.get("mailing_address", ""))
                getattr(self, f"party_{prefix}_email_var").set(c.get("email", ""))
                break

    def _open_library_editor(self):
        from gui.dialogs.library_editor import LibraryEditorDialog
        LibraryEditorDialog(self, on_close=self._on_library_changed)

    def _scan_history_contracts(self):
        folder = filedialog.askdirectory(title="选择历史合同文件夹", parent=self)
        if not folder:
            return
        try:
            self.configure(cursor="watch")
            self.update_idletasks()
            from engines.contract_scanner import scan_contract_folder, save_scan_debug, save_scan_log
            summary = scan_contract_folder(folder)
            save_scan_log(summary)
            save_scan_debug(summary)
        except Exception as exc:
            messagebox.showerror("扫描失败", str(exc), parent=self)
            return
        finally:
            self.configure(cursor="")
        if not summary.get("results"):
            messagebox.showinfo(
                "扫描完成",
                (
                    "未识别到可入库的历史合同。\n\n"
                    f"跳过文件：{summary.get('skipped_count', 0)}\n"
                    f"未识别：{summary.get('unrecognized_count', 0)}"
                ),
                parent=self,
            )
            return
        from gui.dialogs.contract_scan_dialog import ContractScanConfirmDialog
        ContractScanConfirmDialog(self, summary, on_imported=self._on_library_changed)

    def _on_library_changed(self):
        self._party_a_options_cache = None
        self._party_b_names_cache = None
        try:
            self._all_contacts = load_contacts()
        except Exception:
            self._all_contacts = []

    def _on_ok(self):
        result = {}
        # 甲方公司
        for key in ["name", "address", "credit", "bank", "account"]:
            var = getattr(self, f"party_a_{key}_var", None)
            result[f"party_a_{key}"] = var.get() if var else ""
        # 乙方公司
        for key in ["name", "address", "credit", "bank", "account"]:
            var = getattr(self, f"party_b_{key}_var", None)
            result[f"party_b_{key}"] = var.get() if var else ""

        # 甲方联系人（key名与 contract_writer.generate() 参数名一致）
        result["party_a_contact"] = self.party_a_contact_var.get() if hasattr(self, "party_a_contact_var") else ""
        result["party_a_phone"] = self.party_a_phone_var.get() if hasattr(self, "party_a_phone_var") else ""
        result["party_a_email"] = self.party_a_email_var.get() if hasattr(self, "party_a_email_var") else ""
        result["party_a_mailing_address"] = self.party_a_mailing_address_var.get() if hasattr(self, "party_a_mailing_address_var") else ""
        # 乙方联系人
        result["party_b_contact"] = self.party_b_contact_var.get() if hasattr(self, "party_b_contact_var") else ""
        result["party_b_phone"] = self.party_b_phone_var.get() if hasattr(self, "party_b_phone_var") else ""
        result["party_b_email"] = self.party_b_email_var.get() if hasattr(self, "party_b_email_var") else ""
        result["party_b_mailing_address"] = self.party_b_mailing_address_var.get() if hasattr(self, "party_b_mailing_address_var") else ""

        # 签约信息
        for key in ["sign_place", "sign_time", "service_term", "pay_days", "invoice"]:
            var = getattr(self, f"sign_{key}_var", None)
            result[key] = var.get() if var else ""

        # === 保存联系人到联系人库 ===
        self._save_contacts_to_library(result)

        self.callback(result)
        self.destroy()

    def _save_contacts_to_library(self, result):
        """将甲方/乙方联系人信息保存到联系人库"""
        # 甲方联系人
        a_name = result.get("party_a_contact", "").strip()
        if a_name:
            a_info = {
                "unit_name": result.get("party_a_name", ""),
                "contact": a_name,
                "mailing_address": result.get("party_a_mailing_address", ""),
                "phone": result.get("party_a_phone", ""),
                "email": result.get("party_a_email", ""),
            }
            save_contact_entry(a_info)

        # 乙方联系人
        b_name = result.get("party_b_contact", "").strip()
        if b_name:
            b_info = {
                "unit_name": result.get("party_b_name", ""),
                "contact": b_name,
                "mailing_address": result.get("party_b_mailing_address", ""),
                "phone": result.get("party_b_phone", ""),
                "email": result.get("party_b_email", ""),
            }
            save_contact_entry(b_info)
