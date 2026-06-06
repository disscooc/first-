# config.py — 全局配置
import json
import os, sys
from pathlib import Path
from datetime import datetime

APP_NAME = "报价合同生成工具"

def _resource_base():
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent

def _runtime_base():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return _resource_base()

BASE_DIR = _resource_base()
RUNTIME_DIR = _runtime_base()
DATA_DIR = BASE_DIR / "data"
PUBLIC_DATA_DIR = DATA_DIR / "public" if (DATA_DIR / "public").exists() else DATA_DIR
USER_DATA_DIR = Path.home() / "Documents" / APP_NAME
PRIVATE_DATA_DIR = USER_DATA_DIR / "private"
SETTINGS_FILE = USER_DATA_DIR / "settings.json"
DEFAULT_OUT_BASE_DIR = Path.home() / "Documents" / "报价单输出"
OUT_BASE_DIR = DEFAULT_OUT_BASE_DIR
TEMPLATE_DIR = PUBLIC_DATA_DIR / "templates" if (PUBLIC_DATA_DIR / "templates").exists() else BASE_DIR / "templates"

BASE = str(BASE_DIR)
RUNTIME_BASE = str(RUNTIME_DIR)
DATA = str(DATA_DIR)
PUBLIC_DATA = str(PUBLIC_DATA_DIR)
PRIVATE_DATA = str(PRIVATE_DATA_DIR)
USER_DATA = str(USER_DATA_DIR)
TMPL = str(TEMPLATE_DIR)
LOG = str(USER_DATA_DIR / "quote.log")

def _load_user_settings():
    if not SETTINGS_FILE.exists():
        return {}
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8-sig")) or {}
    except Exception:
        return {}

def _save_user_settings(settings):
    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
    SETTINGS_FILE.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")

_settings = _load_user_settings()
saved_out = _settings.get("output_dir")
if saved_out:
    try:
        OUT_BASE_DIR = Path(saved_out).expanduser()
    except Exception:
        OUT_BASE_DIR = DEFAULT_OUT_BASE_DIR

for d in [USER_DATA_DIR, OUT_BASE_DIR]:
    d.mkdir(parents=True, exist_ok=True)

OUT_BASE = str(OUT_BASE_DIR)

def _public_file(name):
    new_path = PUBLIC_DATA_DIR / name
    if PUBLIC_DATA_DIR != DATA_DIR or new_path.exists():
        return str(new_path)
    return str(DATA_DIR / name)

def _private_file(name):
    legacy_path = DATA_DIR / name
    if not (DATA_DIR / "public").exists() and legacy_path.exists():
        return str(legacy_path)
    return str(PRIVATE_DATA_DIR / name)

def _legacy_file(name):
    return str(DATA_DIR / name)

def private_contact_library_exists():
    return Path(F_PRIVATE_CONTACTS).exists() or Path(F_LEGACY_PRIVATE_CONTACTS).exists()

def set_output_dir(new_dir):
    """更新输出目录，同时刷新 OUT_BASE_DIR 和 OUT_BASE 字符串"""
    global OUT_BASE_DIR, OUT_BASE
    p = Path(new_dir)
    p.mkdir(parents=True, exist_ok=True)
    OUT_BASE_DIR = p
    OUT_BASE = str(p)
    settings = _load_user_settings()
    settings["output_dir"] = str(p)
    _save_user_settings(settings)

TAX = 1.06; FREE_STORAGE = 0.5

# 模板
EXCEL_TPL_QTY = os.path.join(TMPL, "机时服务报价单模板-数量.xlsx")
EXCEL_TPL_UNIT = os.path.join(TMPL, "机时服务报价单模板-单价.xlsx")
EXCEL_TPL = EXCEL_TPL_QTY
WORD_STD  = os.path.join(TMPL, "三方报价单模板-标准.docx")
WORD_YD   = os.path.join(TMPL, "三方报价单模板-云达.docx")
CONTRACT_TPL = os.path.join(TMPL, "计算资源服务合同（按单价）-20250115 (9).docx")

# 数据文件
F_ALIAS = _public_file("unit_aliases.json")
F_HIST  = str(USER_DATA_DIR / "unit_history.json")
F_CTRACT= _public_file("contract_customers.json")
F_UNIV_CONTRACTS = _public_file("university_contracts.json")
F_PARTY_A_LIBRARY = _public_file("party_a_library.json")
F_PARTY_B_LIBRARY = _public_file("party_b_library.json")
F_SELLER_CONTACTS = _public_file("seller_contacts.json")
F_PRIVATE_CUSTOMERS = _private_file("contract_private_customers.json")
F_PRIVATE_CONTACTS = _private_file("contract_private_contacts.json")
F_LEGACY_PRIVATE_CUSTOMERS = _legacy_file("contract_private_customers.json")
F_LEGACY_PRIVATE_CONTACTS = _legacy_file("contract_private_contacts.json")

# 价格搜索目录
PRICE_DIRS = [
    BASE,
    PUBLIC_DATA,
    RUNTIME_BASE,
    str(Path.home() / "Documents"),
]

# 公司 & 经理
COMPANIES = ["曙光智算信息技术有限公司","合肥先进计算中心运营管理有限公司","先进计算（天津）信息技术有限公司"]
MANAGERS = []
STORAGE_TYPES = ["分布式文件存储","高IO分布式文件存储"]
STORAGE_SIZES = ["0.5","1","2","5"]

# 高校 & 科研院所（模糊搜索用）
UNIVERSITIES = [
    "北京大学","清华大学","浙江大学","上海交通大学","复旦大学","南京大学","中国科学技术大学",
    "华中科技大学","武汉大学","中山大学","西安交通大学","哈尔滨工业大学","北京航空航天大学",
    "北京师范大学","同济大学","四川大学","东南大学","中国人民大学","南开大学","天津大学",
    "山东大学","中南大学","厦门大学","吉林大学","华南理工大学","大连理工大学","西北工业大学",
    "华东师范大学","中国农业大学","兰州大学","电子科技大学","重庆大学","湖南大学","东北大学",
    "郑州大学","云南大学","西北农林科技大学","新疆大学","国防科技大学",
    "北京理工大学","中央民族大学","中国海洋大学",
]
INSTITUTES = [
    "中国科学院","中国工程物理研究院","中国农业科学院","中国医学科学院","中国林业科学研究院",
    "中国地质科学院","中国环境科学研究院","中国水利水电科学研究院","中国气象科学研究院",
]

# 历史裁剪
HIST_MAX = 500; HIST_TRIM = 300

# 合同默认
DEF_CONTRACT = {
    "party_a_name":"","party_a_contact":"","party_a_phone":"","party_a_email":"","party_a_address":"","party_a_mailing_address":"","party_a_credit":"",
    "party_a_bank":"","party_a_account":"",
    "party_b_name":"曙光智算信息技术有限公司",
    "party_b_addr":"天津市滨海高新区华苑产业区(环外)海泰华科大街15号16层，022-23784462",
    "party_b_mailing_address":"",
    "party_b_credit":"91110108MA02M0XT64","party_b_bank":"中国建设银行北京中关村软件园支行",
    "party_b_account":"11050188380000002839","party_b_contact":"","party_b_phone":"","party_b_email":"",
    "sign_place":"北京","sign_time":"","service_term":"1年","pay_days":"30","invoice":"6%增值税普通发票",
}
