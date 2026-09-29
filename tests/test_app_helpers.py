from dataclasses import replace
from datetime import date
from decimal import Decimal

import pandas as pd
import pytest

from app import (
    cached_file_is_current,
    decimal_value,
    frame_to_products,
    generate_file_data,
    initial_product_frame,
    submission_fingerprint,
)
from customs_tool.domain import ProductRow, ShipmentHeader
from customs_tool.errors import CustomsToolError
from customs_tool.validation import validate_submission


PRODUCT_COLUMNS = [
    "中文品名",
    "英文品名",
    "SKU/规格型号",
    "HS编码",
    "数量",
    "数量单位",
    "箱数",
    "单价",
    "总价",
    "净重",
    "毛重",
    "体积",
    "申报要素",
    "货源地",
]


def product_frame(*rows: dict[str, object]) -> pd.DataFrame:
    return pd.DataFrame(rows, columns=PRODUCT_COLUMNS)


def valid_row(**overrides: object) -> dict[str, object]:
    row: dict[str, object] = {
        "中文品名": "示例产品",
        "英文品名": "Example Product",
        "SKU/规格型号": "SKU-001",
        "HS编码": "0940360999",
        "数量": "10",
        "数量单位": "件",
        "箱数": 2,
        "单价": "12.50",
        "总价": "125.00",
        "净重": 80,
        "毛重": 90,
        "体积": "0.5",
        "申报要素": "品牌类型：无品牌；用途：示例用途",
        "货源地": "示例货源地",
    }
    row.update(overrides)
    return row


@pytest.mark.parametrize("value", [None, "", "   ", pd.NA])
def test_decimal_value_converts_blank_values_to_zero(value: object) -> None:
    assert decimal_value(value) == Decimal("0")


def test_decimal_value_preserves_exact_decimal_text() -> None:
    assert decimal_value("12.50") == Decimal("12.50")


@pytest.mark.parametrize("value", ["not-a-number", object(), float("inf"), float("-inf")])
def test_decimal_value_rejects_invalid_or_non_finite_values(value: object) -> None:
    with pytest.raises(ValueError):
        decimal_value(value)


def test_frame_to_products_converts_valid_row_and_preserves_hs_leading_zero() -> None:
    products = frame_to_products(product_frame(valid_row()))

    assert len(products) == 1
    product = products[0]
    assert product.hs_code == "0940360999"
    assert product.quantity == Decimal("10")
    assert product.boxes == Decimal("2")
    assert product.unit_price == Decimal("12.50")
    assert product.total_price == Decimal("125.00")
    assert product.volume == Decimal("0.5")


def test_frame_to_products_skips_completely_empty_rows() -> None:
    blank_row = {column: None for column in PRODUCT_COLUMNS}

    products = frame_to_products(product_frame(blank_row, valid_row()))

    assert len(products) == 1
    assert products[0].sku == "SKU-001"


def test_initial_product_frame_is_an_empty_product_list() -> None:
    assert frame_to_products(initial_product_frame()) == []


def test_frame_to_products_treats_default_only_dynamic_row_as_empty() -> None:
    default_row = {column: None for column in PRODUCT_COLUMNS}
    default_row.update({"数量单位": "件", "货源地": "示例货源地"})
    for column in ["数量", "箱数", "单价", "总价", "净重", "毛重", "体积"]:
        default_row[column] = 0

    assert frame_to_products(product_frame(default_row)) == []


def test_frame_to_products_keeps_row_when_user_only_enters_quantity() -> None:
    quantity_only_row = {column: None for column in PRODUCT_COLUMNS}
    quantity_only_row.update({"数量": 1, "数量单位": "件", "货源地": "示例货源地"})

    products = frame_to_products(product_frame(quantity_only_row))

    assert len(products) == 1
    assert products[0].quantity == Decimal("1")
    errors = validate_submission(valid_header(), products)
    assert "第1行：中文品名不能为空" in errors
    assert "第1行：HS编码不能为空" in errors


def test_frame_to_products_reports_invalid_number_with_row_and_field() -> None:
    with pytest.raises(ValueError, match=r"第1行：数量不是有效数字"):
        frame_to_products(product_frame(valid_row(数量="ten")))


@pytest.mark.parametrize("value", [float("nan"), float("inf"), float("-inf"), "NaN", "Infinity"])
def test_frame_to_products_rejects_non_finite_numbers(value: object) -> None:
    with pytest.raises(ValueError, match=r"第1行：数量不是有效数字"):
        frame_to_products(product_frame(valid_row(数量=value)))


def valid_header() -> ShipmentHeader:
    return ShipmentHeader(
        domestic_shipper="示例境内发货人有限公司",
        overseas_consignee="EXAMPLE BUYER LTD.",
        production_sales_unit="示例生产单位有限公司",
        contract_no="TEST20260929",
        invoice_no="INV-TEST-001",
        invoice_date=date(2026, 9, 29),
        trade_country="示例国",
        destination_country="示例国",
        destination_port="EXAMPLE PORT",
        exit_port="示例口岸",
        transport_mode="水路运输",
        incoterm="FOB",
        packaging_type="22-纸制或纤维板制盒/箱",
        loading_port="EXAMPLE PORT",
        source_location="示例货源地",
    )


def valid_product() -> ProductRow:
    return frame_to_products(product_frame(valid_row()))[0]


def test_submission_fingerprint_is_stable_for_equivalent_input() -> None:
    header = valid_header()
    product = valid_product()
    equivalent_product = replace(product, quantity=Decimal("10.0"), total_price=Decimal("125.000"))

    assert submission_fingerprint(header, [product]) == submission_fingerprint(
        header, [equivalent_product]
    )


def test_submission_fingerprint_changes_with_header_or_product() -> None:
    header = valid_header()
    product = valid_product()
    original = submission_fingerprint(header, [product])

    changed_header = replace(header, contract_no="CONTRACT-002")
    changed_product = replace(product, quantity=Decimal("11"))

    assert submission_fingerprint(changed_header, [product]) != original
    assert submission_fingerprint(header, [changed_product]) != original


def test_cached_file_is_current_only_for_matching_fingerprint() -> None:
    header = valid_header()
    products = [valid_product()]
    fingerprint = submission_fingerprint(header, products)
    cached = {"bytes": b"xlsx", "name": "output.xlsx", "fingerprint": fingerprint}

    assert cached_file_is_current(cached, fingerprint)
    assert not cached_file_is_current(cached, "different")
    assert not cached_file_is_current(None, fingerprint)


def test_generate_file_data_wraps_os_error_as_user_facing_error(tmp_path) -> None:
    header = valid_header()
    products = [valid_product()]

    def generator(template_path, output_dir, shipment_header, product_rows):
        raise OSError("disk unavailable")

    with pytest.raises(CustomsToolError, match="无法读取或保存临时生成文件"):
        generate_file_data(tmp_path / "template.xls", header, products, generator=generator)
