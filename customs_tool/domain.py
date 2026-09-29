from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from decimal import Decimal


ZERO = Decimal("0")


@dataclass(frozen=True)
class ShipmentHeader:
    domestic_shipper: str = ""
    overseas_consignee: str = ""
    production_sales_unit: str = ""
    contract_no: str = ""
    invoice_no: str = ""
    invoice_date: date | None = None
    trade_country: str = ""
    destination_country: str = ""
    destination_port: str = ""
    exit_port: str = ""
    transport_mode: str = ""
    incoterm: str = "FOB"
    packaging_type: str = "22-纸制或纤维板制盒/箱"
    loading_port: str = ""
    source_location: str = ""

    def output_name(self, today: date) -> str:
        stamp = today.strftime("%Y%m%d")
        contract = re.sub(r'[/\\:*?"<>|]+', "-", self.contract_no.strip())
        return f"报关资料_{contract}_{stamp}.xlsx" if contract else f"报关资料_{stamp}.xlsx"


@dataclass(frozen=True)
class ProductRow:
    chinese_name: str
    english_name: str
    sku: str
    hs_code: str
    quantity: Decimal
    quantity_unit: str
    boxes: Decimal
    unit_price: Decimal
    total_price: Decimal
    net_weight: Decimal
    gross_weight: Decimal
    volume: Decimal
    declaration_elements: str
    source_location: str

    @classmethod
    def empty(cls) -> "ProductRow":
        return cls(
            chinese_name="",
            english_name="",
            sku="",
            hs_code="",
            quantity=ZERO,
            quantity_unit="件",
            boxes=ZERO,
            unit_price=ZERO,
            total_price=ZERO,
            net_weight=ZERO,
            gross_weight=ZERO,
            volume=ZERO,
            declaration_elements="",
            source_location="",
        )


@dataclass(frozen=True)
class Totals:
    quantity: Decimal
    boxes: Decimal
    amount: Decimal
    net_weight: Decimal
    gross_weight: Decimal
    volume: Decimal
