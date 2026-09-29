from decimal import Decimal
from customs_tool.domain import ProductRow
from customs_tool.totals import calculate_totals


def product(quantity, boxes, total_price, net, gross, volume):
    return ProductRow(
        chinese_name="产品", english_name="Product", sku="SKU",
        hs_code="9403609990", quantity=Decimal(quantity), quantity_unit="件",
        boxes=Decimal(boxes), unit_price=Decimal("1"),
        total_price=Decimal(total_price), net_weight=Decimal(net),
        gross_weight=Decimal(gross), volume=Decimal(volume),
        declaration_elements="品名：产品", source_location="示例货源地",
    )


def test_calculate_totals_sums_all_products():
    totals = calculate_totals([
        product("10", "5", "20.5", "30", "35", "1.25"),
        product("3", "3", "9.5", "4", "5", "0.75"),
    ])
    assert totals.quantity == Decimal("13")
    assert totals.boxes == Decimal("8")
    assert totals.amount == Decimal("30.0")
    assert totals.net_weight == Decimal("34")
    assert totals.gross_weight == Decimal("40")
    assert totals.volume == Decimal("2.00")


def test_calculate_totals_handles_empty_list():
    totals = calculate_totals([])
    assert totals.quantity == Decimal("0")
    assert totals.boxes == Decimal("0")
