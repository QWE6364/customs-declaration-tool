from dataclasses import replace
from datetime import date, datetime
from decimal import Decimal

import pytest

from customs_tool.domain import ProductRow, ShipmentHeader
from customs_tool.validation import validate_submission


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
    )


def valid_product(**overrides) -> ProductRow:
    values = {
        "chinese_name": "产品",
        "english_name": "Product",
        "sku": "SKU-001",
        "hs_code": "9403609990",
        "quantity": Decimal("10"),
        "quantity_unit": "件",
        "boxes": Decimal("2"),
        "unit_price": Decimal("3.50"),
        "total_price": Decimal("35.00"),
        "net_weight": Decimal("20"),
        "gross_weight": Decimal("22"),
        "volume": Decimal("1.2"),
        "declaration_elements": "品名：产品",
        "source_location": "示例货源地",
    }
    values.update(overrides)
    return ProductRow(**values)


def test_valid_header_and_product_have_no_errors():
    assert validate_submission(valid_header(), [valid_product()]) == []


def test_empty_header_reports_required_fields():
    errors = validate_submission(ShipmentHeader(), [valid_product()])

    assert "境内发货人不能为空" in errors
    assert "合同协议号不能为空" in errors


def test_empty_transport_mode_reports_required_field():
    header = replace(valid_header(), transport_mode="   ")

    errors = validate_submission(header, [valid_product()])

    assert "运输方式不能为空" in errors


def test_empty_loading_port_reports_required_field():
    header = replace(valid_header(), loading_port="")

    errors = validate_submission(header, [valid_product()])

    assert "装运港不能为空" in errors


def test_matching_product_amount_has_no_consistency_error():
    errors = validate_submission(valid_header(), [valid_product()])

    assert "第1行：总价与数量×单价不一致" not in errors


def test_product_amount_difference_at_tolerance_boundary_is_accepted():
    errors = validate_submission(
        valid_header(),
        [valid_product(total_price=Decimal("35.01"))],
    )

    assert "第1行：总价与数量×单价不一致" not in errors


def test_product_amount_difference_over_tolerance_is_rejected():
    errors = validate_submission(
        valid_header(),
        [valid_product(total_price=Decimal("35.0101"))],
    )

    assert "第1行：总价与数量×单价不一致" in errors


def test_product_amount_mismatch_reports_second_row_number():
    errors = validate_submission(
        valid_header(),
        [valid_product(), valid_product(total_price=Decimal("34.98"))],
    )

    assert "第2行：总价与数量×单价不一致" in errors


def test_invalid_numeric_value_does_not_add_amount_consistency_error():
    errors = validate_submission(
        valid_header(),
        [valid_product(quantity=Decimal("NaN"))],
    )

    assert "第1行：数量不是有效数字" in errors
    assert "第1行：总价与数量×单价不一致" not in errors


def test_negative_quantity_reports_product_row_error():
    errors = validate_submission(
        valid_header(),
        [valid_product(quantity=Decimal("-1"))],
    )

    assert "第1行：数量不能为负数" in errors


def test_empty_product_list_reports_required_product():
    assert validate_submission(valid_header(), []) == ["至少需要一个产品"]


@pytest.mark.parametrize("value", [Decimal("NaN"), Decimal("Infinity"), Decimal("-Infinity")])
def test_non_finite_decimal_reports_invalid_number(value):
    errors = validate_submission(valid_header(), [valid_product(quantity=value)])

    assert "第1行：数量不是有效数字" in errors


def test_non_decimal_reports_invalid_number_on_second_product():
    errors = validate_submission(
        valid_header(),
        [valid_product(), valid_product(quantity="10")],
    )

    assert "第2行：数量不是有效数字" in errors


def test_blank_product_text_reports_required_field():
    errors = validate_submission(
        valid_header(),
        [valid_product(chinese_name="   ")],
    )

    assert "第1行：中文品名不能为空" in errors


def test_header_source_location_supplies_blank_product_location():
    header = replace(valid_header(), source_location="示例货源地")

    errors = validate_submission(header, [valid_product(source_location="")])

    assert "第1行：货源地不能为空" not in errors


def test_non_string_product_text_reports_required_field():
    errors = validate_submission(
        valid_header(),
        [valid_product(chinese_name=123)],
    )

    assert "第1行：中文品名不能为空" in errors


def test_non_string_header_text_reports_required_field():
    header = replace(valid_header(), domestic_shipper=123)

    errors = validate_submission(header, [valid_product()])

    assert "境内发货人不能为空" in errors


def test_non_date_invoice_date_reports_invalid_format():
    header = replace(valid_header(), invoice_date="2026-09-28")

    errors = validate_submission(header, [valid_product()])

    assert "发票日期格式无效" in errors


def test_datetime_invoice_date_reports_invalid_format():
    header = replace(
        valid_header(),
        invoice_date=datetime(2026, 9, 28, 12, 0),
    )

    errors = validate_submission(header, [valid_product()])

    assert "发票日期格式无效" in errors
