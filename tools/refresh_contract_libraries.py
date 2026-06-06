import json
import re
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
PUBLIC = DATA / "contract_customers.json"
UNIV = DATA / "university_contracts.json"
PRIVATE = DATA / "contract_private_customers.json"
CONTACTS = DATA / "contract_private_contacts.json"
PARTY_A_LIBRARY = DATA / "party_a_library.json"

B_DEFAULT = {
    "party_b_name": "曙光智算信息技术有限公司",
    "party_b_contact": "",
    "party_b_phone": "",
    "party_b_address": "天津市滨海高新区华苑产业区(环外)海泰华科大街15号16层，022-23784462",
    "party_b_credit": "91110108MA02M0XT64",
    "party_b_bank": "中国建设银行北京中关村软件园支行",
    "party_b_account": "11050188380000002839",
}


def read_json(path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return default


def write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def rec(name, address="", credit="", bank="", account="", source="", category="北京重点高校/科研院所"):
    return {
        "party_a_name": name,
        "party_a_contact": "",
        "party_a_phone": "",
        "party_a_address": address,
        "party_a_credit": credit,
        "party_a_bank": bank,
        "party_a_account": account,
        **B_DEFAULT,
        "party_a_source_url": source,
        "category": category,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


PUBLIC_RECORDS = [
    rec("北京大学", "北京市海淀区颐和园路5号", category="北京985/211高校"),
    rec("清华大学", "北京市海淀区清华园", "12100000400000624D", "中国工商银行股份有限公司北京海淀西区支行营业室", "0200004509089131550", "https://wsbx.tsinghua.edu.cn/", "北京985/211高校"),
    rec("中国人民大学", "北京市海淀区中关村大街59号", category="北京985/211高校"),
    rec("北京航空航天大学", "北京市海淀区学院路37号", "12100000400011227Y", source="https://cwc.buaa.edu.cn/info/1017/5016.htm", category="北京985/211高校"),
    rec("北京理工大学", "北京市海淀区中关村南大街5号，010-68918558", "12100000400009127B", category="北京985/211高校"),
    rec("中国农业大学", "北京市海淀区圆明园西路2号", "12100000400018162G", source="https://kyy.cau.edu.cn/attach/0/b62329557ee84b2ab0f9057add915594.pdf", category="北京985/211高校"),
    rec("北京师范大学", "北京市海淀区新街口外大街19号", "12100000400010056C", source="https://atc.bnu.edu.cn/zyxz/114b5c51fe8d4f9c822a70d9756e505e.htm", category="北京985/211高校"),
    rec("中央民族大学", "北京市海淀区中关村南大街27号", category="北京985/211高校"),
    rec("北京交通大学", "北京市海淀区西直门外上园村3号", category="北京211高校"),
    rec("北京工业大学", "北京市朝阳区平乐园100号", category="北京211高校"),
    rec("北京科技大学", "北京市海淀区学院路30号", "121000004000022245", source="https://caiwu.ustb.edu.cn/shouji/shuihaozhanghao1/2018-08-01/601.html", category="北京211高校"),
    rec("北京化工大学", "北京市朝阳区北三环东路15号", category="北京211高校"),
    rec("北京邮电大学", "北京市海淀区西土城路10号", category="北京211高校"),
    rec("北京林业大学", "北京市海淀区清华东路35号", category="北京211高校"),
    rec("北京中医药大学", "北京市朝阳区北三环东路11号", category="北京211高校"),
    rec("北京外国语大学", "北京市海淀区西三环北路2号", category="北京211高校"),
    rec("中国传媒大学", "北京市朝阳区定福庄东街1号", category="北京211高校"),
    rec("中央财经大学", "北京市海淀区学院南路39号", category="北京211高校"),
    rec("对外经济贸易大学", "北京市朝阳区惠新东街10号", category="北京211高校"),
    rec("北京体育大学", "北京市海淀区信息路48号", category="北京211高校"),
    rec("中央音乐学院", "北京市西城区鲍家街43号", category="北京211高校"),
    rec("中国政法大学", "北京市昌平区府学路27号", category="北京211高校"),
    rec("华北电力大学", "北京市昌平区朱辛庄北农路2号，010-61772625", "1210000040000983X8", "建设银行北京沙河支行", "11001016000056055041", category="北京211高校"),
    rec("中国石油大学（北京）", "北京市昌平区府学路18号", "12100000400006110Y", category="北京211高校"),
    rec("中国矿业大学（北京）", "北京市海淀区学院路丁11号", category="北京211高校"),
    rec("中国地质大学（北京）", "北京市海淀区学院路29号", category="北京211高校"),
    rec("中国科学院大学", "北京市石景山区玉泉路19号甲", category="北京科研院所/高校"),
    rec("中国科学院计算技术研究所", "北京市海淀区中关村科学院南路6号", "12100000400012342E", category="中科院北京院所"),
    rec("中国科学院大气物理研究所", "北京市朝阳区德胜门外祁家豁子华严里40号", "121000004000119994", source="https://iap.cas.cn/gb/jgsz/glxt/jhcwc/xzc_cwc/202101/P020241030707944946136.pdf", category="中科院北京院所"),
    rec("中国科学院物理研究所", "北京市海淀区中关村南三街8号", category="中科院北京院所"),
    rec("中国科学院化学研究所", "北京市海淀区中关村北一街2号", category="中科院北京院所"),
    rec("中国科学院力学研究所", "北京市海淀区北四环西路15号", category="中科院北京院所"),
    rec("中国科学院高能物理研究所", "北京市石景山区玉泉路19号乙", category="中科院北京院所"),
    rec("中国科学院过程工程研究所", "北京市海淀区中关村北二街1号", category="中科院北京院所"),
    rec("中国科学院电工研究所", "北京市海淀区中关村北二条6号", category="中科院北京院所"),
    rec("中国科学院微电子研究所", "北京市朝阳区北土城西路3号", category="中科院北京院所"),
    rec("中国科学院半导体研究所", "北京市海淀区清华东路甲35号", category="中科院北京院所"),
    rec("中国科学院自动化研究所", "北京市海淀区中关村东路95号", category="中科院北京院所"),
    rec("中国科学院软件研究所", "北京市海淀区中关村南四街4号", category="中科院北京院所"),
    rec("中国科学院空天信息创新研究院", "北京市海淀区邓庄南路9号", category="中科院北京院所"),
    rec("中国科学院国家天文台", "北京市朝阳区大屯路甲20号", category="中科院北京院所"),
    rec("中国科学院数学与系统科学研究院", "北京市海淀区中关村东路55号", category="中科院北京院所"),
    rec("中国科学院理论物理研究所", "北京市海淀区中关村东路55号", category="中科院北京院所"),
    rec("中国科学院生物物理研究所", "北京市朝阳区大屯路15号", category="中科院北京院所"),
    rec("中国科学院遗传与发育生物学研究所", "北京市朝阳区北辰西路1号院2号", category="中科院北京院所"),
    rec("中国科学院植物研究所", "北京市海淀区香山南辛村20号", category="中科院北京院所"),
    rec("中国科学院动物研究所", "北京市朝阳区北辰西路1号院5号", category="中科院北京院所"),
    rec("中国科学院地理科学与资源研究所", "北京市朝阳区大屯路甲11号", category="中科院北京院所"),
    rec("中国科学院科技战略咨询研究院", "北京市海淀区中关村北一条15号", category="中科院北京院所"),
]

ALIASES = {
    "北京大学": ["北大", "beida", "bd", "bjdx"],
    "清华大学": ["清华", "qinghua", "qh", "tsinghua"],
    "中国人民大学": ["人大", "renmin", "renda", "rd"],
    "北京航空航天大学": ["北航", "beihang", "bh", "buaa"],
    "北京理工大学": ["北理工", "beiligong", "blg", "bit"],
    "中国农业大学": ["农大", "中农", "nongda", "nd", "cau"],
    "北京师范大学": ["北师大", "beishida", "bsd"],
    "中央民族大学": ["民大", "minzu", "minda"],
    "北京交通大学": ["北交大", "beijiaoda", "bjd"],
    "北京工业大学": ["北工大", "beigongda", "bgd"],
    "北京科技大学": ["北科大", "beikeda", "bkd", "ustb"],
    "北京化工大学": ["北化", "北化大", "beihua", "bhu"],
    "北京邮电大学": ["北邮", "beiyou", "by"],
    "北京林业大学": ["北林", "beilin", "bl"],
    "北京中医药大学": ["北中医", "beizhongyi"],
    "北京外国语大学": ["北外", "beiwai", "bw"],
    "中国传媒大学": ["中传", "zhongchuan", "zc", "cuc"],
    "中央财经大学": ["央财", "yangcai", "yc"],
    "对外经济贸易大学": ["贸大", "对外经贸", "uibe"],
    "北京体育大学": ["北体", "北体大", "btd"],
    "中央音乐学院": ["央音", "yangyin"],
    "中国政法大学": ["法大", "zhengfa", "fada", "cupl"],
    "华北电力大学": ["华电", "huadian", "hd"],
    "中国石油大学（北京）": ["石大", "中石大北京", "shida", "cupb"],
    "中国矿业大学（北京）": ["矿大北京", "kuangda"],
    "中国地质大学（北京）": ["地大北京", "dida"],
    "中国科学院大学": ["国科大", "guokeda", "gkd", "ucas"],
    "中国科学院计算技术研究所": ["计算所", "中科院计算所", "ict"],
    "中国科学院大气物理研究所": ["大气所", "中科院大气所", "iap"],
    "中国科学院物理研究所": ["物理所", "中科院物理所"],
    "中国科学院化学研究所": ["化学所", "中科院化学所"],
    "中国科学院自动化研究所": ["自动化所", "中科院自动化所"],
    "中国科学院软件研究所": ["软件所", "中科院软件所"],
}

NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def docx_text(path):
    with zipfile.ZipFile(path) as zf:
        out = []
        for name in zf.namelist():
            if not (name.startswith("word/") and name.endswith(".xml")):
                continue
            if not ("document" in name or "header" in name or "footer" in name):
                continue
            try:
                root = ET.fromstring(zf.read(name))
            except Exception:
                continue
            for text_node in root.iter(NS + "t"):
                if text_node.text:
                    out.append(text_node.text)
            out.append("\n")
        return "".join(out)


def clean(value):
    return re.sub(r"\s+", " ", value or "").strip(" ：:，,。")


def first(pattern, text, flags=re.S):
    match = re.search(pattern, text, flags)
    return clean(match.group(1)) if match else ""


def valid_unit_name(name):
    if not name or len(name) > 40:
        return False
    bad_words = ["乙 方", "乙方", "签订地点", "签订时间", "注册地址", "我公司名称"]
    return not any(word in name for word in bad_words)


def party_a_name_before_address(text):
    address_pos = text.find("注册地址、电话：")
    if address_pos < 0:
        return ""
    prefix = text[:address_pos]
    matches = list(re.finditer(r"甲\s*方：\s*([^甲乙签订注册]{1,50})", prefix))
    for match in reversed(matches):
        name = clean(match.group(1))
        if valid_unit_name(name):
            return name
    return ""


def extract_contract(text):
    party_a_name = party_a_name_before_address(text)
    if not valid_unit_name(party_a_name):
        return None

    party_a_address = first(r"甲\s*方：\s*.*?注册地址、电话：\s*(.*?)统一社会信用代码：", text)
    party_a_credit = first(r"统一社会信用代码：\s*([0-9A-Z]{8,30})\s*账号名称：", text)
    party_a_account_name = first(r"账号名称：\s*(.*?)开户银行：", text)
    party_a_bank = first(r"开户银行：\s*(.*?)开户账号：", text)
    party_a_account = first(r"开户账号：\s*(.*?)联\s*系\s*人：", text)
    party_a_contact = first(r"联\s*系\s*人：\s*(.*?)通信地址：", text)
    party_a_phone = first(r"电\s*话：\s*([0-9\-\s]{5,30})\s*乙\s*方：", text)
    party_a_email = first(r"对账函的邮箱为：\s*([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})", text)

    body_b_pos = text.find("乙    方：", text.find("电    话："))
    body_b = text[body_b_pos:] if body_b_pos >= 0 else text
    party_b_name = first(r"乙\s*方：\s*(.*?)注册地址、电话：", body_b)
    party_b_address = first(r"乙\s*方：\s*.*?注册地址、电话：\s*(.*?)统一社会信用代码：", body_b)
    party_b_credit = first(r"统一社会信用代码：\s*([0-9A-Z]{8,30})\s*账号名称：", body_b)
    party_b_bank = first(r"开户银行：\s*(.*?)开户账号：", body_b)
    party_b_account = first(r"开户账号：\s*(.*?)联\s*系\s*人：", body_b)
    party_b_contact = first(r"联\s*系\s*人：\s*(.*?)通信地址：", body_b)
    party_b_phone = first(r"电\s*话：\s*([0-9\-\s]{5,30})", body_b)

    return {
        "PartyAName": party_a_name,
        "PartyAAddressPhone": party_a_address,
        "PartyACreditCode": party_a_credit,
        "PartyAAccountName": party_a_account_name or party_a_name,
        "PartyABank": party_a_bank,
        "PartyAAccountNo": party_a_account,
        "PartyAAddress": party_a_address.split("，")[0] if party_a_address else "",
        "PartyAContact": party_a_contact,
        "PartyAPhone": party_a_phone,
        "PartyAEmail": party_a_email,
        "PartyBName": party_b_name or B_DEFAULT["party_b_name"],
        "PartyBAddressPhone": party_b_address or B_DEFAULT["party_b_address"],
        "PartyBCreditCode": party_b_credit or B_DEFAULT["party_b_credit"],
        "PartyBAccountName": party_b_name or B_DEFAULT["party_b_name"],
        "PartyBBank": party_b_bank or B_DEFAULT["party_b_bank"],
        "PartyBAccountNo": party_b_account or B_DEFAULT["party_b_account"],
        "PartyBAddress": party_b_address or B_DEFAULT["party_b_address"],
        "PartyBContact": party_b_contact,
        "PartyBPhone": party_b_phone,
        "PaymentDays": "30",
        "InvoiceType": "6%增值税普通发票",
    }


def normalize_private(raw):
    if not isinstance(raw, dict):
        return None
    name = raw.get("PartyAName") or raw.get("party_a_name")
    if not name:
        return None
    return {
        "PartyAName": name,
        "PartyAAddressPhone": raw.get("PartyAAddressPhone") or raw.get("party_a_address") or "",
        "PartyACreditCode": raw.get("PartyACreditCode") or raw.get("party_a_credit") or "",
        "PartyAAccountName": raw.get("PartyAAccountName") or name,
        "PartyABank": raw.get("PartyABank") or raw.get("party_a_bank") or "",
        "PartyAAccountNo": raw.get("PartyAAccountNo") or raw.get("party_a_account") or "",
        "PartyAAddress": raw.get("PartyAAddress") or raw.get("party_a_address") or "",
        "PartyAContact": raw.get("PartyAContact") or raw.get("party_a_contact") or "",
        "PartyAPhone": raw.get("PartyAPhone") or raw.get("party_a_phone") or "",
        "PartyAEmail": raw.get("PartyAEmail") or raw.get("party_a_email") or "",
        "PartyBName": raw.get("PartyBName") or raw.get("party_b_name") or B_DEFAULT["party_b_name"],
        "PartyBAddressPhone": raw.get("PartyBAddressPhone") or raw.get("party_b_address") or B_DEFAULT["party_b_address"],
        "PartyBCreditCode": raw.get("PartyBCreditCode") or raw.get("party_b_credit") or B_DEFAULT["party_b_credit"],
        "PartyBAccountName": raw.get("PartyBAccountName") or raw.get("party_b_name") or B_DEFAULT["party_b_name"],
        "PartyBBank": raw.get("PartyBBank") or raw.get("party_b_bank") or B_DEFAULT["party_b_bank"],
        "PartyBAccountNo": raw.get("PartyBAccountNo") or raw.get("party_b_account") or B_DEFAULT["party_b_account"],
        "PartyBAddress": raw.get("PartyBAddress") or raw.get("party_b_address") or B_DEFAULT["party_b_address"],
        "PartyBContact": raw.get("PartyBContact") or raw.get("party_b_contact") or "",
        "PartyBPhone": raw.get("PartyBPhone") or raw.get("party_b_phone") or "",
        "PaymentDays": raw.get("PaymentDays") or "30",
        "InvoiceType": raw.get("InvoiceType") or "6%增值税普通发票",
        "UpdatedAt": raw.get("UpdatedAt") or raw.get("updated_at") or "",
        "SourceFiles": raw.get("SourceFiles") or ([] if not raw.get("SourceFile") else [raw.get("SourceFile")]),
    }


def merge_private(existing, incoming):
    merged = dict(existing)
    for key, value in incoming.items():
        if not value:
            continue
        if key == "SourceFiles":
            merged[key] = sorted(set((existing.get(key) or []) + value))
        else:
            merged[key] = value
    return merged


def main():
    public = {item["party_a_name"]: item for item in PUBLIC_RECORDS}
    for name, aliases in ALIASES.items():
        if name in public:
            public[name]["aliases"] = aliases
    write_json(PUBLIC, dict(sorted(public.items())))

    party_a_library = {
        name: {
            "name": name,
            "aliases": sorted(set(item.get("aliases", []) + [name])),
            "category": item.get("category", ""),
        }
        for name, item in public.items()
    }
    write_json(PARTY_A_LIBRARY, dict(sorted(party_a_library.items())))

    university = read_json(UNIV, {})
    if not isinstance(university, dict):
        university = {}
    university = {
        key: value for key, value in university.items()
        if isinstance(key, str) and "?" not in key and valid_unit_name(key)
    }
    for item in PUBLIC_RECORDS:
        university[item["party_a_name"]] = item
    write_json(UNIV, dict(sorted(university.items())))

    old_private = read_json(PRIVATE, [])
    old_items = old_private if isinstance(old_private, list) else list(old_private.values()) if isinstance(old_private, dict) else []
    private_by_name = {}
    for raw in old_items:
        item = normalize_private(raw)
        if item and valid_unit_name(item["PartyAName"]):
            private_by_name[item["PartyAName"]] = item

    scanned = []
    for root in [Path.home() / "Desktop" / "合同", Path.home() / "Documents" / "合同"]:
        if not root.exists():
            continue
        for path in root.rglob("*.docx"):
            if path.name.startswith("~$"):
                continue
            try:
                item = extract_contract(docx_text(path))
            except Exception:
                continue
            if not item:
                continue
            item["UpdatedAt"] = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            item["SourceFiles"] = [str(path)]
            scanned.append(item)
            private_by_name[item["PartyAName"]] = merge_private(private_by_name.get(item["PartyAName"], {}), item)

    private_final = sorted(private_by_name.values(), key=lambda item: item.get("PartyAName", ""))
    contacts = []
    seen = set()
    for item in private_final:
        key = (item.get("PartyAName", ""), item.get("PartyAContact", ""), item.get("PartyAPhone", ""), item.get("PartyAEmail", ""))
        if key in seen or not any(key[1:]):
            continue
        seen.add(key)
        contacts.append({
            "unit_name": key[0],
            "contact": key[1],
            "phone": key[2],
            "email": key[3],
            "source_files": item.get("SourceFiles", []),
            "updated_at": item.get("UpdatedAt", ""),
        })

    write_json(PRIVATE, private_final)
    write_json(CONTACTS, contacts)
    print(json.dumps({
        "public_records": len(public),
        "party_a_library_records": len(party_a_library),
        "university_records": len(university),
        "private_records": len(private_final),
        "scanned_docx_records": len(scanned),
        "private_contacts": len(contacts),
        "scanned_names": sorted({item["PartyAName"] for item in scanned}),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
