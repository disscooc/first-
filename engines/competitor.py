# competitor.py — 竞争对手计算
import copy; from models import TPSetting; import config

def comp_items(items, markup, rule, is_total):
    adj=[]
    for it in items:
        n=copy.copy(it)
        base_price = _base_price(it)
        if rule=="MarkupPrice":
            n.price=round(base_price*(1+markup),4)
            n.amount=round(n.price*it.quantity,2)
        elif rule=="SameTotalMoreQty":
            n.price=round(base_price*(1+markup),4)
            if not is_total:
                n.amount=it.amount
                n.quantity=round(it.amount/n.price,2) if n.price>0 and it.amount else it.quantity
        adj.append(n)
    return adj

def _base_price(item):
    price = getattr(item, "price", 0) or 0
    if price > 0:
        return price
    return getattr(item, "orig_price", 0) or 0

def tp_settings(c1,r1,c2,r2):
    return [
        TPSetting(company=c1, markup=r1, layout="StandardThirdParty"),
        TPSetting(company=c2, markup=r2, layout="YundaThirdParty"),
    ]
