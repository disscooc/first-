"""Prepare sanitized public data for PyInstaller builds.

The build payload intentionally excludes buyer/customer private data. Seller
contacts are copied from the public seller contact file when present, or derived
from the legacy mixed contact file by keeping only contacts whose unit is a
seller company.
"""
from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
TEMPLATES = ROOT / "templates"
PAYLOAD = ROOT / "build_share_payload"
PUBLIC = PAYLOAD / "data" / "public"

PRIVATE_NAMES = {
    "contract_private_contacts.json",
    "contract_private_customers.json",
    "buyer_contacts.json",
}

PUBLIC_JSON_NAMES = {
    "contract_customers.json",
    "party_a_library.json",
    "party_b_library.json",
    "unit_aliases.json",
    "university_contracts.json",
}

SENSITIVE_BUYER_KEYS = {
    "party_a_contact",
    "party_a_phone",
    "party_a_email",
    "party_a_mailing_address",
    "party_a_contact_mailing_address",
    "party_a_contact_phone",
    "party_a_contact_email",
    "PartyAContact",
    "PartyAPhone",
    "PartyAEmail",
    "PartyAMailingAddress",
}


def read_json(path: Path, default):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return default


def write_json(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def copy_file(src: Path, dst: Path):
    if not src.exists() or src.name in PRIVATE_NAMES:
        return
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)


def copy_public_json():
    source_dirs = [DATA / "public", DATA]
    copied = set()
    for source_dir in source_dirs:
        if not source_dir.exists():
            continue
        for name in PUBLIC_JSON_NAMES:
            if name in copied:
                continue
            src = source_dir / name
            if src.exists():
                copy_file(src, PUBLIC / name)
                copied.add(name)
    sanitize_public_customer_json()


def sanitize_public_customer_json():
    for name in ("contract_customers.json", "party_a_library.json", "university_contracts.json"):
        path = PUBLIC / name
        data = read_json(path, None)
        if data is None:
            continue
        records = data.values() if isinstance(data, dict) else data if isinstance(data, list) else []
        for record in records:
            if not isinstance(record, dict):
                continue
            for key in SENSITIVE_BUYER_KEYS:
                if key in record:
                    record[key] = ""
        write_json(path, data)


def copy_templates():
    source = DATA / "public" / "templates" if (DATA / "public" / "templates").exists() else TEMPLATES
    if not source.exists():
        return
    dst = PUBLIC / "templates"
    if dst.exists():
        shutil.rmtree(dst)
    shutil.copytree(
        source,
        dst,
        ignore=shutil.ignore_patterns(
            "__pycache__",
            "*.pyc",
            "*.auto_backup.*",
            "*.backup_*",
            "yunda_unpacked",
        ),
    )


def copy_price_lists():
    for source_dir in [ROOT, DATA, DATA / "public"]:
        if not source_dir.exists():
            continue
        for src in source_dir.glob("*.xlsx"):
            if src.name.startswith("~$"):
                continue
            if any(key in src.name for key in ("List Price", "价格", "目录", "服务资源")):
                copy_file(src, PUBLIC / src.name)


def build_seller_contacts():
    party_b = read_json(PUBLIC / "party_b_library.json", {})
    seller_names = set(party_b.keys()) if isinstance(party_b, dict) else set()
    existing_public = read_json(DATA / "public" / "seller_contacts.json", [])
    if existing_public:
        write_json(PUBLIC / "seller_contacts.json", existing_public)
        return

    legacy_contacts = read_json(DATA / "contract_private_contacts.json", [])
    if not isinstance(legacy_contacts, list):
        legacy_contacts = []
    seller_contacts = [
        contact for contact in legacy_contacts
        if contact.get("unit_name", "") in seller_names
    ]
    write_json(PUBLIC / "seller_contacts.json", seller_contacts)


def assert_no_private_payload():
    bad = []
    for path in PAYLOAD.rglob("*"):
        if path.is_file() and (path.name in PRIVATE_NAMES or "private" in path.parts):
            bad.append(path)
    if bad:
        joined = "\n".join(str(path) for path in bad)
        raise SystemExit(f"Private data found in build payload:\n{joined}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--clean", action="store_true")
    args = parser.parse_args()

    if args.clean and PAYLOAD.exists():
        shutil.rmtree(PAYLOAD)
    PUBLIC.mkdir(parents=True, exist_ok=True)

    copy_public_json()
    copy_templates()
    copy_price_lists()
    build_seller_contacts()
    assert_no_private_payload()
    print(f"Prepared public build payload: {PAYLOAD}")


if __name__ == "__main__":
    main()
