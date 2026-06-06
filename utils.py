# utils.py — 工具函数
import os,re,json,traceback
from datetime import datetime
import config

# === JSON ===
def read_json(path, default=None):
    if default is None:
        default = []
    if not path:
        return default
    if not os.path.exists(path):
        return default
    try:
        with open(path,'r',encoding='utf-8-sig') as f: return json.load(f) or default
    except: return default

def write_json(path, data):
    if not path:
        return
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path,'w',encoding='utf-8') as f: json.dump(data,f,ensure_ascii=False,indent=2)

# === 文件名 ===
def safe_fn(name):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]','_',str(name)).strip(' .') or 'untitled'

def out_dir():
    d = os.path.join(config.OUT_BASE, datetime.now().strftime("%Y-%m"))
    os.makedirs(d, exist_ok=True); return d

def new_path(safe_unit, suffix, ext):
    d = out_dir(); base = f"{safe_unit}-{suffix}-{datetime.now().strftime('%Y%m%d')}"
    p = os.path.join(d, f"{base}{ext}")
    i = 1
    while os.path.exists(p): p = os.path.join(d, f"{base}-{i}{ext}"); i += 1
    return p

# === 格式化 ===
def fmt_num(v, fmt="0.##"):
    """格式化数字，fmt 为 Excel 风格格式字符串。
    - '0' / '0.':    整数（无小数）
    - '0.00':         固定2位小数，四舍五入
    - '0.##':         最多2位小数，去末尾0（用于可变小数位的价格）
    - '0.####':       最多4位小数，去末尾0
    """
    if v is None: return ""
    try:
        fv = round(float(v), 6)
        if fmt == "0" or fmt == "0.":
            # 整数，无小数
            return str(int(round(fv)))
        if '.' not in fmt:
            return str(int(round(fv)))
        after_dot = fmt.split('.')[1]
        if '#' not in after_dot:
            # 固定小数位：如 '0.00' => 2位
            n = len(after_dot)
            return ("{:." + str(n) + "f}").format(round(fv, n))
        else:
            # 最多N位（去末尾0）：如 '0.##' => 最多2位，去末尾0
            n = len(after_dot)
            s = ("{:." + str(n) + "f}").format(round(fv, n)).rstrip('0').rstrip('.')
            return s if s else '0'
    except: return str(v)

def fmt_money(v):
    if v is None: return ""
    try: return f"{float(v):,.2f}"
    except: return str(v)

def try_num(s):
    if not s: return False,0
    try: return True, float(re.sub(r'[^\d.\-]','',str(s)))
    except: return False,0

# === 金额大写 ===
CN_D = ["零","壹","贰","叁","肆","伍","陆","柒","捌","玖"]
CN_U = ["","拾","佰","仟"]; CN_S = ["","万","亿","兆"]

def _sec(num):
    r=[]; n=num
    for i in range(4):
        d=n%10
        if d>0: r.append(CN_D[d]+CN_U[i])
        elif r and r[-1]!="零": r.append("零")
        n//=10
    while r and r[-1]=="零": r.pop()
    return "".join(reversed(r))

def _int_cn(num):
    if num==0: return "零"
    parts=[]; n=num; si=0
    while n>0:
        s=n%10000
        if s>0: parts.append(_sec(s)+CN_S[si])
        si+=1; n//=10000
    return "".join(reversed(parts))

def cny(amt):
    y=int(amt); j=int(round((amt-y)*10)); f=int(round((amt-y-j/10)*100))
    parts=[]
    if y>0: parts.append(_int_cn(y)+"元")
    if j>0: parts.append(CN_D[j]+"角")
    if f>0: parts.append(CN_D[f]+"分")
    return ("".join(parts)+"整") if parts else "零元整"

# === 文本辅助 ===
def price_header(items):
    g=any(getattr(i,'rtype','') in ('GPU','异构加速卡') for i in items)
    c=any(getattr(i,'rtype','')=='CPU' for i in items)
    if g and c: return "价格(元/核时或元/卡时)"
    if g: return "价格(元/卡时)"
    return "价格(元/核时)"

def unit_text(items):
    g=any(getattr(i,'rtype','') in ('GPU','异构加速卡') for i in items)
    c=any(getattr(i,'rtype','')=='CPU' for i in items)
    if g and c: return "元/核时或元/卡时"
    return "元/卡时" if g else "元/核时"

def comp_remark(rt): return "1卡时=1张卡*1小时" if rt in ("GPU","异构加速卡") else "1核时=1核心*1小时" if rt=="CPU" else ""
def qty_unit(rt): return "卡时" if rt in ("GPU","异构加速卡") else "核时"
def stor_remark(qty):
    try:
        qty = float(qty)
        if qty > 0:
            free = fmt_num(qty, '0.##')
            return f"免费赠送{free}T，免费额度内不收费"
    except:
        pass
    return ""

def quote_total(comp, stor, is_total, amt):
    if is_total: return amt
    return sum(getattr(i,'amount',0) or 0 for i in list(comp)+list(stor))

# === 历史记录 ===
def save_hist(name):
    data = read_json(config.F_HIST, [])
    if name not in data: data.insert(0, name)
    else: data.remove(name); data.insert(0, name)
    if len(data) > config.HIST_MAX: data = data[:config.HIST_TRIM]
    write_json(config.F_HIST, data)

def load_hist(): return read_json(config.F_HIST, [])

# === 甲方联想搜索 ===
COMMON_PINYIN = {
    "北京大学": ("beijingdaxue", "bjdx", "bd", "beida"),
    "清华大学": ("qinghuadaxue", "qhdx", "qh", "tsinghua"),
    "中国人民大学": ("zhongguorenmindaxue", "zgrmdx", "renda", "rd"),
    "北京航空航天大学": ("beijinghangkonghangtiandaxue", "bjhkdhtdx", "beihang", "bh"),
    "北京理工大学": ("beijingligongdaxue", "bjlgdx", "beiligong", "blg", "bit"),
    "中国农业大学": ("zhongguonongyedaxue", "zgnydx", "nongda", "nd"),
    "北京师范大学": ("beijingshifandaxue", "bjsfdx", "beishida", "bsd"),
    "中央民族大学": ("zhongyangminzudaxue", "zymzdx", "minzu", "minda"),
    "北京交通大学": ("beijingjiaotongdaxue", "bjjtdx", "beijiaoda", "bjd"),
    "北京工业大学": ("beijinggongyedaxue", "bjgydx", "beigongda", "bgd"),
    "北京科技大学": ("beijingkejidaxue", "bjkjdx", "beikedai", "bkd", "ustb"),
    "北京化工大学": ("beijinghuagongdaxue", "bjhgdx", "beihua", "bhu"),
    "北京邮电大学": ("beijingyoudiandaxue", "bjyddx", "beiyou", "by"),
    "北京林业大学": ("beijinglinyedaxue", "bjlydx", "beilinda", "bld"),
    "北京中医药大学": ("beijingzhongyiyaodaxue", "bjzyydx", "beizhongyi"),
    "北京外国语大学": ("beijingwaiguoyudaxue", "bjwgydx", "beiwai", "bw"),
    "中国传媒大学": ("zhongguochuanmeidaxue", "zgcma", "zhongchuan", "zc", "cuc"),
    "中央财经大学": ("zhongyangcaijingdaxue", "zycjdx", "yangcai", "yc"),
    "对外经济贸易大学": ("duiwaijingjimaoyidaxue", "dwjjmydx", "duiwaijingmao", "uibe"),
    "北京体育大学": ("beijingtiyudaxue", "bjtydx", "beitida", "btd"),
    "中央音乐学院": ("zhongyangyinyuexueyuan", "zyyyxy", "yangyin"),
    "中国政法大学": ("zhongguozhengfadaxue", "zgzfdx", "fazheng", "cupL".lower()),
    "华北电力大学": ("huabeidianlidaxue", "hbdldx", "huadian", "hd"),
    "中国石油大学（北京）": ("zhongguoshiyoudaxuebeijing", "zgsy dxbj".replace(" ", ""), "shida", "sy"),
    "中国矿业大学（北京）": ("zhongguokuangyedaxuebeijing", "zgkydxbj", "kuangda", "kd"),
    "中国地质大学（北京）": ("zhongguodizhidaxuebeijing", "zgdz dxbj".replace(" ", ""), "dida", "dd"),
    "中国科学院大学": ("zhongguokexueyuandaxue", "zgkxydx", "guokeda", "gkd", "ucas"),
}

PINYIN_INITIAL = {
    "北":"b","京":"j","大":"d","学":"x","清":"q","华":"h","中":"z","国":"g","人":"r","民":"m","航":"h","空":"k","天":"t","理":"l","工":"g",
    "农":"n","业":"y","师":"s","范":"f","央":"y","族":"z","交":"j","通":"t","科":"k","技":"j","化":"h","邮":"y","电":"d","林":"l",
    "外":"w","语":"y","传":"c","媒":"m","财":"c","经":"j","贸":"m","易":"y","体":"t","育":"y","音":"y","乐":"y","政":"z","法":"f",
    "石":"s","油":"y","矿":"k","地":"d","质":"z","院":"y","所":"s","计":"j","算":"s","物":"w","气":"q","植":"z","动":"d","微":"w",
    "软":"r","件":"j","自":"z","高":"g","能":"n","过":"g","程":"c","数":"s","与":"y","系":"x","统":"t","生":"s","半":"b","导":"d",
    "元":"y","初":"c","启":"q","智":"z","瑞":"r","锋":"f","讯":"x","香":"x","港":"g","燕":"y","山":"s","南":"n","开":"k","齐":"q","鲁":"l",
}

def _initials(text):
    return "".join(PINYIN_INITIAL.get(ch, ch.lower()) for ch in str(text) if ch.strip())

def _norm_query(text):
    return re.sub(r"[\s·（）()_\-—,，.。]+", "", str(text or "").lower())

def _record_name(record):
    if isinstance(record, dict):
        return record.get("party_a_name") or record.get("PartyAName") or record.get("unit_name") or record.get("name") or ""
    return ""

def _add_name(index, name, aliases=None, contact=None):
    name = str(name or "").strip()
    if not name:
        return
    entry = index.setdefault(name, {"aliases": set(), "contacts": set()})
    entry["aliases"].add(name)
    entry["aliases"].add(_initials(name))
    for alias in aliases or []:
        if alias:
            entry["aliases"].add(alias)
    if name in COMMON_PINYIN:
        entry["aliases"].update(COMMON_PINYIN[name])
    if contact:
        entry["contacts"].add(contact)
        entry["aliases"].add(contact)
        entry["aliases"].add(_initials(contact))

def _party_a_index():
    index = {}
    for name in getattr(config, "UNIVERSITIES", []) + getattr(config, "INSTITUTES", []):
        _add_name(index, name)
    for path in [getattr(config, "F_PARTY_A_LIBRARY", ""), config.F_CTRACT, config.F_UNIV_CONTRACTS]:
        data = read_json(path, {})
        values = data.values() if isinstance(data, dict) else data if isinstance(data, list) else []
        for record in values:
            _add_name(index, _record_name(record), record.get("aliases", []) if isinstance(record, dict) else [])
    private = read_json(getattr(config, "F_PRIVATE_CUSTOMERS", ""), [])
    values = private.values() if isinstance(private, dict) else private if isinstance(private, list) else []
    for record in values:
        _add_name(index, _record_name(record), contact=(record.get("PartyAContact") or record.get("party_a_contact")) if isinstance(record, dict) else "")
    contacts = read_json(getattr(config, "F_PRIVATE_CONTACTS", ""), [])
    for record in contacts if isinstance(contacts, list) else []:
        _add_name(index, record.get("unit_name", ""), contact=record.get("contact", ""))
    return index

def _unit_score(q, name, data):
    nq = _norm_query(name)
    aliases = [_norm_query(a) for a in data.get("aliases", set()) if a]
    contacts = [_norm_query(c) for c in data.get("contacts", set()) if c]
    if q == nq:
        return 0
    if nq.startswith(q):
        return 1
    if q in aliases:
        return 2
    if any(alias.startswith(q) for alias in aliases):
        return 3
    if any(q in contact or contact.startswith(q) for contact in contacts):
        return 4
    if q in nq or any(q in alias for alias in aliases):
        return 5
    return None

def search_units(q):
    q = _norm_query(q)
    index = _party_a_index()
    if not q:
        hist = [name for name in load_hist() if isinstance(name, str) and name in index]
        names = hist + [name for name in index if name not in hist]
        return names[:30]
    ranked = []
    for name, data in index.items():
        score = _unit_score(q, name, data)
        if score is not None:
            ranked.append((score, len(name), name))
    ranked.sort()
    return [name for _, __, name in ranked[:30]]


# === 联系人搜索（支持拼音/姓氏/子串联想）===
def _contact_index():
    """构建联系人搜索索引：{联系人姓名: {'unit': 单位, 'phone': 电话, 'email': 邮箱}}"""
    from config import F_PRIVATE_CONTACTS
    data = read_json(F_PRIVATE_CONTACTS, [])
    index = {}
    for entry in data:
        if not isinstance(entry, dict):
            continue
        name = entry.get("contact", "").strip()
        if not name:
            continue
        index[name] = {
            "unit": entry.get("unit_name", "").strip(),
            "phone": entry.get("phone", "").strip(),
            "email": entry.get("email", "").strip(),
        }
    return index


def search_contacts(q, limit=30):
    """搜索联系人，支持姓名子串、拼音首字母、单位名匹配。
    返回 [{'name': 姓名, 'unit': 单位, 'phone': 电话, 'email': 邮箱}, ...]
    """
    q = _norm_query(q)
    index = _contact_index()
    if not q:
        # 无输入时返回全部
        return [{"name": n, **v} for n, v in index.items()][:limit]

    results = []
    for name, info in index.items():
        nq = _norm_query(name)
        unit = _norm_query(info.get("unit", ""))
        # 1. 姓名完全匹配（最高优先级）
        if q == nq:
            results.append((0, len(name), {"name": name, **info}))
        # 2. 姓名以 q 开头
        elif nq.startswith(q):
            results.append((1, len(name), {"name": name, **info}))
        # 3. 拼音首字母匹配
        elif _initials(name).lower().startswith(q):
            results.append((2, len(name), {"name": name, **info}))
        # 4. 姓名包含 q（子串）
        elif q in nq:
            results.append((3, len(name), {"name": name, **info}))
        # 5. 单位名包含 q
        elif q in unit:
            results.append((4, len(name), {"name": name, **info}))
        # 6. 拼音首字母包含 q（松散匹配）
        elif q in _initials(name).lower():
            results.append((5, len(name), {"name": name, **info}))

    results.sort(key=lambda x: (x[0], x[1]))
    return [item[2] for item in results[:limit]]


# === 异常收口 ===
def handle_err(e, ctx=""):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    tb = traceback.format_exc()
    msg = f"[{ts}] {ctx}: {e}\n{tb}\n"
    try:
        with open(config.LOG,'a',encoding='utf-8') as f: f.write(msg)
    except: pass
    return f"{ctx}: {e}"

def show_err(parent, e, ctx=""):
    import tkinter.messagebox as mb
    msg = handle_err(e, ctx)
    mb.showerror("错误", msg, parent=parent)
