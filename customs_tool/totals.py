from decimal import Decimal
from typing import Iterable
from customs_tool.domain import ProductRow, Totals


def calculate_totals(products: Iterable[ProductRow]) -> Totals:
    rows = list(products)
    return Totals(
        quantity=sum((p.quantity for p in rows), Decimal("0")),
        boxes=sum((p.boxes for p in rows), Decimal("0")),
        amount=sum((p.total_price for p in rows), Decimal("0")),
        net_weight=sum((p.net_weight for p in rows), Decimal("0")),
        gross_weight=sum((p.gross_weight for p in rows), Decimal("0")),
        volume=sum((p.volume for p in rows), Decimal("0")),
    )
