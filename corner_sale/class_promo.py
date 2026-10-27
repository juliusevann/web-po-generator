from dataclasses import dataclass, field
from typing import Optional

@dataclass
class PromoRow:
    article: str
    colour: str
    harga_awal: Optional[int]
    harga_diskon: Optional[float]
    moved_from: Optional[str] = None
    moved_to:   Optional[str] = None
    harga_disesuaikan: bool = False


@dataclass
class PromoTable:
    sheet: str
    promo_name: str
    discount_key: str
    count: int = 0
    rows: list = field(default_factory=list)
    category_suffix: Optional[str] = None