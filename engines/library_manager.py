# engines/library_manager.py — 库编辑数据层
import os, json, shutil, config
from pathlib import Path
from datetime import datetime
from utils import read_json, write_json

# === 乙方库 ===
F_PARTY_B = config.F_PARTY_B_LIBRARY

def load_party_b():
    return read_json(F_PARTY_B, {})

def save_party_b(data):
    write_json(F_PARTY_B, data)

def list_party_b_names():
    data = load_party_b()
    return list(data.keys())

def get_party_b(name):
    data = load_party_b()
    return data.get(name, {})

def save_party_b_entry(name, info):
    data = load_party_b()
    old_name = info.pop("_old_name", None)
    if old_name and old_name != name and old_name in data:
        del data[old_name]
    data[name] = {
        "name": info.get("name", name),
        "address": info.get("address", ""),
        "credit": info.get("credit", ""),
        "bank": info.get("bank", ""),
        "account": info.get("account", ""),
    }
    save_party_b(data)

def delete_party_b(name):
    data = load_party_b()
    if name in data:
        del data[name]
        save_party_b(data)
        return True
    return False


# === 甲方 — 私有库 ===
def load_private_customers():
    raw = read_json(config.F_PRIVATE_CUSTOMERS, [])
    if not raw and Path(getattr(config, "F_LEGACY_PRIVATE_CUSTOMERS", "")).exists():
        raw = read_json(config.F_LEGACY_PRIVATE_CUSTOMERS, [])
    if isinstance(raw, list):
        return raw
    return list(raw.values())

def save_private_customers(data):
    write_json(config.F_PRIVATE_CUSTOMERS, data)

def find_private_customer(unit_name):
    data = load_private_customers()
    for entry in data:
        name = entry.get("PartyAName") or entry.get("party_a_name") or ""
        if name and (unit_name == name or unit_name in name or name in unit_name):
            return entry
    return None

def save_private_customer(unit_name, info):
    """保存甲方到私有库，如果同名则覆盖"""
    data = load_private_customers()
    old_name_key = info.pop("_old_name", None)
    # 查找并覆盖
    found = False
    for i, entry in enumerate(data):
        name = entry.get("PartyAName") or entry.get("party_a_name") or ""
        if old_name_key and old_name_key == name:
            data[i] = _private_party_a_record(unit_name, info)
            found = True
            break
        if name == unit_name:
            data[i] = _private_party_a_record(unit_name, info)
            found = True
            break
    if not found:
        data.append(_private_party_a_record(unit_name, info))
    save_private_customers(data)

def _private_party_a_record(unit_name, info):
    return {
        "PartyAName": unit_name,
        "PartyAAddressPhone": info.get("address", ""),
        "PartyACreditCode": info.get("credit", ""),
        "PartyAAccountName": unit_name,
        "PartyABank": info.get("bank", ""),
        "PartyAAccountNo": info.get("account", ""),
        "PartyAAddress": info.get("address", ""),
        "PartyBName": "曙光智算信息技术有限公司",
        "PartyBAddressPhone": "",
        "PartyBCreditCode": "",
        "PartyBAccountName": "",
        "PartyBBank": "",
        "PartyBAccountNo": "",
        "PartyBAddress": "",
        "PaymentDays": "",
        "InvoiceType": "",
        "UpdatedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "SourceFiles": [],
    }

def delete_private_customer(unit_name):
    data = load_private_customers()
    data = [e for e in data if (e.get("PartyAName") or e.get("party_a_name") or "") != unit_name]
    save_private_customers(data)

def list_party_a_names():
    """返回所有甲方名称列表，标注来源"""
    names = {}
    # 公开库
    pub_data = read_json(config.F_CTRACT, {})
    if isinstance(pub_data, dict):
        for name in pub_data:
            names[name] = "public"
    # 私有库
    priv_data = load_private_customers()
    for entry in priv_data:
        name = entry.get("PartyAName") or entry.get("party_a_name") or ""
        if name:
            names[name] = "private"  # 私有覆盖公开
    return names

def resolve_party_a(unit_name):
    """获取甲方的合并信息：私有 > 公开"""
    # 先查私有库
    priv_entry = find_private_customer(unit_name)
    if priv_entry:
        return {
            "name": priv_entry.get("PartyAName") or priv_entry.get("party_a_name", ""),
            "address": priv_entry.get("PartyAAddressPhone") or priv_entry.get("PartyAAddress") or priv_entry.get("party_a_address", ""),
            "credit": priv_entry.get("PartyACreditCode") or priv_entry.get("party_a_credit", ""),
            "bank": priv_entry.get("PartyABank") or priv_entry.get("party_a_bank", ""),
            "account": priv_entry.get("PartyAAccountNo") or priv_entry.get("party_a_account", ""),
            "contact": priv_entry.get("PartyAContact") or priv_entry.get("party_a_contact", ""),
            "phone": priv_entry.get("PartyAPhone") or priv_entry.get("party_a_phone", ""),
            "email": priv_entry.get("PartyAEmail", ""),
            "source": "private",
        }
    # 查公开库
    pub_data = read_json(config.F_CTRACT, {})
    if isinstance(pub_data, dict) and unit_name in pub_data:
        entry = pub_data[unit_name]
        return {
            "name": entry.get("party_a_name", unit_name),
            "address": entry.get("party_a_address", ""),
            "credit": entry.get("party_a_credit", ""),
            "bank": entry.get("party_a_bank", ""),
            "account": entry.get("party_a_account", ""),
            "contact": entry.get("party_a_contact", ""),
            "phone": entry.get("party_a_phone", ""),
            "email": "",
            "source": "public",
        }
    return {"name": unit_name, "source": "new"}


# === 联系人库 ===
def _party_b_name_set():
    return set(list_party_b_names())

def _is_seller_unit(unit_name):
    return str(unit_name or "").strip() in _party_b_name_set()

def _load_legacy_contacts():
    legacy = getattr(config, "F_LEGACY_PRIVATE_CONTACTS", "")
    if legacy and Path(legacy).exists():
        raw = read_json(legacy, [])
        return raw if isinstance(raw, list) else []
    return []

CONTACT_DEDUPE_FIELDS = ("phone", "email", "mailing_address", "unit_name", "contact")


def _contact_key(entry):
    return (
        str(entry.get("unit_name", "") or "").strip(),
        str(entry.get("contact", "") or "").strip(),
    )


def _contact_score(entry):
    return sum(1 for key in CONTACT_DEDUPE_FIELDS if str(entry.get(key, "") or "").strip())


def _dedupe_contact_records(records):
    groups = {}
    order = []
    for index, entry in enumerate(records or []):
        key = _contact_key(entry)
        if not key[0] or not key[1]:
            order.append((index, key, entry))
            continue
        if key not in groups:
            groups[key] = []
        groups[key].append((index, entry))
        order.append((index, key, entry))

    keep_indices = set()
    duplicate_groups = 0
    duplicate_records = 0
    for items in groups.values():
        if len(items) == 1:
            keep_indices.add(items[0][0])
            continue
        duplicate_groups += 1
        duplicate_records += len(items) - 1
        keep_index, _entry = max(items, key=lambda pair: (_contact_score(pair[1]), -pair[0]))
        keep_indices.add(keep_index)

    result = []
    for index, key, entry in order:
        if not key[0] or not key[1] or index in keep_indices:
            result.append(entry)
    return result, duplicate_groups, duplicate_records


def dedupe_contacts_for_display(records):
    return _dedupe_contact_records(records)[0]


def contact_duplicate_summary(records):
    _deduped, group_count, duplicate_count = _dedupe_contact_records(records)
    return group_count, duplicate_count


def contact_duplicate_summary_by_side():
    seller_groups, seller_duplicates = contact_duplicate_summary(load_seller_contacts(dedupe=False))
    buyer_groups, buyer_duplicates = contact_duplicate_summary(load_buyer_contacts(dedupe=False))
    return seller_groups + buyer_groups, seller_duplicates + buyer_duplicates


def load_seller_contacts(dedupe=True):
    raw = read_json(getattr(config, "F_SELLER_CONTACTS", ""), [])
    contacts = raw if isinstance(raw, list) else []
    if not contacts:
        # 开发环境兼容旧的混合联系人库：仅把乙方公司联系人视为公共乙方联系人。
        contacts = [entry for entry in _load_legacy_contacts() if _is_seller_unit(entry.get("unit_name", ""))]
    if dedupe:
        contacts = dedupe_contacts_for_display(contacts)
    return contacts

def load_buyer_contacts(dedupe=True):
    raw = read_json(config.F_PRIVATE_CONTACTS, [])
    contacts = raw if isinstance(raw, list) else []
    if not contacts:
        # 开发环境兼容旧的混合联系人库：过滤掉乙方联系人，只取甲方联系人。
        contacts = [entry for entry in _load_legacy_contacts() if not _is_seller_unit(entry.get("unit_name", ""))]
    if dedupe:
        contacts = dedupe_contacts_for_display(contacts)
    return contacts

def load_contacts(dedupe=True):
    return load_seller_contacts(dedupe=dedupe) + load_buyer_contacts(dedupe=dedupe)

def load_contacts_raw():
    return load_seller_contacts(dedupe=False) + load_buyer_contacts(dedupe=False)

def save_contacts(data):
    sellers = []
    buyers = []
    for entry in data:
        if _is_seller_unit(entry.get("unit_name", "")):
            sellers.append(entry)
        else:
            buyers.append(entry)
    write_json(getattr(config, "F_SELLER_CONTACTS", ""), sellers)
    write_json(config.F_PRIVATE_CONTACTS, buyers)

def find_contact_by_name(contact_name):
    data = load_contacts()
    results = []
    for entry in data:
        if contact_name and (contact_name in entry.get("contact", "")):
            results.append(entry)
    return results

def save_contact_entry(info):
    """保存或更新联系人"""
    path = getattr(config, "F_SELLER_CONTACTS", "") if _is_seller_unit(info.get("unit_name", "")) else config.F_PRIVATE_CONTACTS
    data = read_json(path, [])
    if not data and path == getattr(config, "F_SELLER_CONTACTS", ""):
        data = load_seller_contacts()
    elif not data and path == config.F_PRIVATE_CONTACTS:
        data = load_buyer_contacts()
    old_name = info.pop("_old_name", None)
    unit_name = info.get("unit_name", "")
    contact = info.get("contact", "")
    # 查找已存在的
    found = False
    for i, entry in enumerate(data):
        existing_unit = entry.get("unit_name", "")
        existing_contact = entry.get("contact", "")
        if old_name and f"{existing_unit}|{existing_contact}" == old_name:
            data[i] = _contact_record(info)
            found = True
            break
        if existing_unit == unit_name and existing_contact == contact:
            data[i] = _contact_record(info)
            found = True
            break
    if not found:
        data.append(_contact_record(info))
    write_json(path, data)

def append_contact_entry(info):
    path = getattr(config, "F_SELLER_CONTACTS", "") if _is_seller_unit(info.get("unit_name", "")) else config.F_PRIVATE_CONTACTS
    data = read_json(path, [])
    if not isinstance(data, list):
        data = []
    data.append(_contact_record(info))
    write_json(path, data)

def contact_duplicate_exists(info, old_key=None):
    unit_name = str(info.get("unit_name", "") or "").strip()
    contact = str(info.get("contact", "") or "").strip()
    if not unit_name or not contact:
        return False
    path = getattr(config, "F_SELLER_CONTACTS", "") if _is_seller_unit(unit_name) else config.F_PRIVATE_CONTACTS
    data = read_json(path, [])
    if not isinstance(data, list):
        data = []
    for entry in data:
        existing_key = f"{entry.get('unit_name', '')}|{entry.get('contact', '')}"
        if old_key and existing_key == old_key:
            continue
        if _contact_key(entry) == (unit_name, contact):
            return True
    return False

def _backup_contact_file(path, timestamp):
    if not path or not Path(path).exists():
        return None
    source = Path(path)
    backup = source.with_name(f"contacts_backup_{source.stem}_{timestamp}{source.suffix}")
    shutil.copy2(source, backup)
    return str(backup)

def cleanup_duplicate_contacts():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    results = {
        "groups": 0,
        "duplicates": 0,
        "backups": [],
    }
    for path in (getattr(config, "F_SELLER_CONTACTS", ""), config.F_PRIVATE_CONTACTS):
        raw = read_json(path, [])
        data = raw if isinstance(raw, list) else []
        deduped, group_count, duplicate_count = _dedupe_contact_records(data)
        results["groups"] += group_count
        results["duplicates"] += duplicate_count
        if duplicate_count <= 0:
            continue
        backup = _backup_contact_file(path, timestamp)
        if backup:
            results["backups"].append(backup)
        write_json(path, deduped)
    return results

def _contact_record(info):
    return {
        "unit_name": info.get("unit_name", ""),
        "contact": info.get("contact", ""),
        "mailing_address": info.get("mailing_address", ""),
        "phone": info.get("phone", ""),
        "email": info.get("email", ""),
        "source_files": info.get("source_files", []) or [],
        "note": info.get("note", ""),
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

def delete_contact(unit_name, contact):
    path = getattr(config, "F_SELLER_CONTACTS", "") if _is_seller_unit(unit_name) else config.F_PRIVATE_CONTACTS
    data = read_json(path, [])
    data = [e for e in data if not (e.get("unit_name") == unit_name and e.get("contact") == contact)]
    write_json(path, data)
