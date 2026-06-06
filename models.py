# models.py
from dataclasses import dataclass

@dataclass
class Resource:
    raw_type:str=""; rtype:str=""; name:str=""; region:str=""; spec:str=""; shared_price:float=0; exclusive_price:float=0

@dataclass
class QuoteItem:
    item_type:str="Compute"; rtype:str=""; name:str=""; spec:str=""
    orig_price:float=0; discount:float=1.0; price:float=0
    quantity:float=0; amount:float=0; remark:str=""; free_quota:bool=False

@dataclass
class TPSetting:
    company:str; markup:float; layout:str="StandardThirdParty"
