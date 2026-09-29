from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

import customs_tool.excel_generator as generator
import customs_tool.output_verifier as output_verifier
from customs_tool.domain import ProductRow, ShipmentHeader, Totals
from customs_tool.errors import VerificationError
from customs_tool.layout import DECLARATION_FIRST_PRODUCT_ROW, TABLE_FIRST_PRODUCT_ROW, table_product_rows
from customs_tool.output_verifier import verify_totals, verify_workbook


EXPECTED = Totals(
    quantity=Decimal("30"),
    boxes=Decimal("6"),
    amount=Decimal("375"),
    net_weight=Decimal("240"),
    gross_weight=Decimal("270"),
    volume=Decimal("1.5"),
)


HEADER = ShipmentHeader(
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


PRODUCTS = [
    ProductRow(
        chinese_name="示例产品一",
        english_name="Example Product One",
        sku="SKU-001",
        hs_code="9403609991",
        quantity=Decimal("10"),
        quantity_unit="件",
        boxes=Decimal("2"),
        unit_price=Decimal("12.5"),
        total_price=Decimal("125"),
        net_weight=Decimal("80"),
        gross_weight=Decimal("90"),
        volume=Decimal("0.5"),
        declaration_elements="品牌类型：无品牌；用途：示例用途",
        source_location="示例货源地",
    ),
    ProductRow(
        chinese_name="示例产品二",
        english_name="Example Product Two",
        sku="SKU-002",
        hs_code="9403609992",
        quantity=Decimal("20"),
        quantity_unit="件",
        boxes=Decimal("4"),
        unit_price=Decimal("12.5"),
        total_price=Decimal("250"),
        net_weight=Decimal("160"),
        gross_weight=Decimal("180"),
        volume=Decimal("1"),
        declaration_elements="品牌类型：无品牌；用途：示例用途；序号：2",
        source_location="示例货源地",
    ),
]


def _workbook_path(tmp_path: Path, product_count: int = 2) -> Path:
    path = tmp_path / "generated.xlsx"
    workbook = openpyxl.Workbook()
    declaration = workbook.active
    declaration.title = "任意名称一"
    invoice = workbook.create_sheet("任意名称二")
    packing = workbook.create_sheet("任意名称三")
    total_row = table_product_rows(product_count)

    invoice[f"B{total_row}"] = "TOTAL"
    invoice[f"D{total_row}"] = float(EXPECTED.quantity)
    invoice[f"E{total_row}"] = float(EXPECTED.boxes)
    invoice[f"G{total_row}"] = float(EXPECTED.amount)

    packing[f"B{total_row}"] = "TOTAL"
    packing[f"D{total_row}"] = float(EXPECTED.quantity)
    packing[f"E{total_row}"] = float(EXPECTED.net_weight)
    packing[f"F{total_row}"] = float(EXPECTED.gross_weight)
    packing[f"G{total_row}"] = float(EXPECTED.volume)

    workbook.save(path)
    workbook.close()
    return path


def _verified_workbook_path(tmp_path: Path) -> Path:
    path = _workbook_path(tmp_path, len(PRODUCTS))
    workbook = openpyxl.load_workbook(path)
    declaration, invoice, packing = workbook.worksheets[:3]

    declaration["A10"] = HEADER.contract_no
    declaration["I12"] = float(EXPECTED.boxes)
    declaration["K12"] = float(EXPECTED.gross_weight)
    declaration["M12"] = float(EXPECTED.net_weight)
    declaration["U10"] = HEADER.exit_port

    for index, product in enumerate(PRODUCTS):
        main_row = DECLARATION_FIRST_PRODUCT_ROW + index * 3
        declaration[f"B{main_row}"] = product.hs_code
        declaration[f"E{main_row}"] = product.chinese_name
        declaration[f"F{main_row}"] = product.sku
        declaration[f"I{main_row}"] = float(product.quantity)
        declaration[f"J{main_row}"] = product.quantity_unit
        declaration[f"L{main_row}"] = float(product.unit_price)
        declaration[f"N{main_row}"] = float(product.total_price)
        declaration[f"U{main_row}"] = HEADER.destination_country
        declaration[f"W{main_row}"] = product.source_location
        declaration[f"I{main_row + 1}"] = float(product.gross_weight)
        declaration[f"J{main_row + 1}"] = "千克"
        declaration[f"I{main_row + 2}"] = float(product.boxes)
        declaration[f"J{main_row + 2}"] = "箱"

        table_row = TABLE_FIRST_PRODUCT_ROW + index
        invoice[f"C{table_row}"] = product.english_name
        invoice[f"D{table_row}"] = float(product.quantity)
        invoice[f"E{table_row}"] = float(product.boxes)
        invoice[f"F{table_row}"] = float(product.unit_price)
        invoice[f"G{table_row}"] = float(product.total_price)

        packing[f"C{table_row}"] = f"{product.english_name}\n{product.chinese_name}"
        packing[f"D{table_row}"] = float(product.quantity)
        packing[f"E{table_row}"] = float(product.net_weight)
        packing[f"F{table_row}"] = float(product.gross_weight)
        packing[f"G{table_row}"] = float(product.volume)
        packing[f"H{table_row}"] = product.hs_code
        packing[f"I{table_row}"] = product.declaration_elements

    workbook.save(path)
    workbook.close()
    return path


def test_verify_totals_accepts_matching_totals(tmp_path: Path) -> None:
    path = _workbook_path(tmp_path)

    verify_totals(path, 2, EXPECTED)


def test_verify_workbook_accepts_all_matching_fields(tmp_path: Path) -> None:
    path = _verified_workbook_path(tmp_path)

    verify_workbook(path, HEADER, PRODUCTS)


@pytest.mark.parametrize(
    ("sheet_index", "cell", "bad_value", "message"),
    [
        (0, "A10", "WRONG", "报关资料.*A10.*合同号"),
        (0, "I12", 999, "报关资料.*I12.*总箱数"),
        (0, "K12", 999, "报关资料.*K12.*总毛重"),
        (0, "M12", 999, "报关资料.*M12.*总净重"),
        (0, "U10", "WRONG", "报关资料.*U10.*离境口岸"),
        (0, "B18", "WRONG", "报关资料.*18.*HS编码"),
        (0, "E18", "WRONG", "报关资料.*18.*中文名"),
        (0, "F18", "WRONG", "报关资料.*18.*SKU"),
        (0, "I18", 999, "报关资料.*18.*数量"),
        (0, "J18", "WRONG", "报关资料.*18.*单位"),
        (0, "L18", 999, "报关资料.*18.*单价"),
        (0, "N18", 999, "报关资料.*18.*总价"),
        (0, "U18", "WRONG", "报关资料.*18.*目的国"),
        (0, "W18", "WRONG", "报关资料.*18.*货源地"),
        (0, "I19", 999, "报关资料.*19.*毛重"),
        (0, "J19", "WRONG", "报关资料.*19.*毛重单位"),
        (0, "I20", 999, "报关资料.*20.*箱数"),
        (0, "J20", "WRONG", "报关资料.*20.*箱数单位"),
        (1, "C12", "WRONG", "INVOICE.*12.*英文名"),
        (1, "D12", 999, "INVOICE.*12.*数量"),
        (1, "E12", 999, "INVOICE.*12.*箱数"),
        (1, "F12", 999, "INVOICE.*12.*单价"),
        (1, "G12", 999, "INVOICE.*12.*总价"),
        (2, "C12", "WRONG", "PACKING.*12.*中英文名称"),
        (2, "D12", 999, "PACKING.*12.*数量"),
        (2, "E12", 999, "PACKING.*12.*净重"),
        (2, "F12", 999, "PACKING.*12.*毛重"),
        (2, "G12", 999, "PACKING.*12.*体积"),
        (2, "H12", "WRONG", "PACKING.*12.*HS编码"),
        (2, "I12", "WRONG", "PACKING.*12.*申报要素"),
    ],
)
def test_verify_workbook_detects_each_field_mismatch(
    tmp_path: Path,
    sheet_index: int,
    cell: str,
    bad_value: object,
    message: str,
) -> None:
    path = _verified_workbook_path(tmp_path)
    workbook = openpyxl.load_workbook(path)
    workbook.worksheets[sheet_index][cell] = bad_value
    workbook.save(path)
    workbook.close()

    with pytest.raises(VerificationError, match=message):
        verify_workbook(path, HEADER, PRODUCTS)


def test_verify_workbook_uses_decimal_tolerance_for_product_values(tmp_path: Path) -> None:
    path = _verified_workbook_path(tmp_path)
    workbook = openpyxl.load_workbook(path)
    workbook.worksheets[1]["F12"] = float(PRODUCTS[0].unit_price + Decimal("0.000001"))
    workbook.save(path)
    workbook.close()

    verify_workbook(path, HEADER, PRODUCTS)

    workbook = openpyxl.load_workbook(path)
    workbook.worksheets[1]["F12"] = float(PRODUCTS[0].unit_price + Decimal("0.0000011"))
    workbook.save(path)
    workbook.close()

    with pytest.raises(VerificationError, match="INVOICE.*12.*单价"):
        verify_workbook(path, HEADER, PRODUCTS)


def test_verify_totals_detects_invoice_amount_mismatch(tmp_path: Path) -> None:
    path = _workbook_path(tmp_path)
    workbook = openpyxl.load_workbook(path)
    workbook.worksheets[1][f"G{table_product_rows(2)}"] = 374
    workbook.save(path)
    workbook.close()

    with pytest.raises(VerificationError, match="INVOICE 总金额"):
        verify_totals(path, 2, EXPECTED)


@pytest.mark.parametrize(
    ("sheet_index", "column", "message"),
    [
        (1, "D", "INVOICE 数量"),
        (1, "E", "INVOICE 箱数"),
        (2, "D", "PACKING 数量"),
        (2, "E", "PACKING 总净重"),
        (2, "F", "PACKING 总毛重"),
        (2, "G", "PACKING 总体积"),
    ],
)
def test_verify_totals_detects_each_common_field_mismatch(
    tmp_path: Path,
    sheet_index: int,
    column: str,
    message: str,
) -> None:
    path = _workbook_path(tmp_path)
    workbook = openpyxl.load_workbook(path)
    workbook.worksheets[sheet_index][f"{column}{table_product_rows(2)}"] = -1
    workbook.save(path)
    workbook.close()

    with pytest.raises(VerificationError, match=message):
        verify_totals(path, 2, EXPECTED)


def test_verify_totals_converts_missing_file_to_verification_error(tmp_path: Path) -> None:
    with pytest.raises(VerificationError, match="文件"):
        verify_totals(tmp_path / "missing.xlsx", 2, EXPECTED)


def test_verify_totals_converts_insufficient_worksheets_to_verification_error(tmp_path: Path) -> None:
    path = tmp_path / "too-few-sheets.xlsx"
    workbook = openpyxl.Workbook()
    workbook.save(path)
    workbook.close()

    with pytest.raises(VerificationError, match="工作表不足"):
        verify_totals(path, 2, EXPECTED)


@pytest.mark.parametrize("sheet_index", [1, 2])
def test_verify_totals_converts_missing_total_row_to_verification_error(
    tmp_path: Path,
    sheet_index: int,
) -> None:
    path = _workbook_path(tmp_path)
    workbook = openpyxl.load_workbook(path)
    workbook.worksheets[sheet_index][f"B{table_product_rows(2)}"] = None
    workbook.save(path)
    workbook.close()

    with pytest.raises(VerificationError, match="总计行缺失"):
        verify_totals(path, 2, EXPECTED)


def test_verify_totals_uses_decimal_tolerance_of_one_millionth(tmp_path: Path) -> None:
    path = _workbook_path(tmp_path)
    total_row = table_product_rows(2)
    workbook = openpyxl.load_workbook(path)
    workbook.worksheets[1][f"G{total_row}"] = float(EXPECTED.amount + Decimal("0.000001"))
    workbook.worksheets[2][f"G{total_row}"] = float(EXPECTED.volume - Decimal("0.000001"))
    workbook.save(path)
    workbook.close()

    verify_totals(path, 2, EXPECTED)

    workbook = openpyxl.load_workbook(path)
    workbook.worksheets[1][f"G{total_row}"] = float(EXPECTED.amount + Decimal("0.0000011"))
    workbook.save(path)
    workbook.close()

    with pytest.raises(VerificationError, match="INVOICE 总金额"):
        verify_totals(path, 2, EXPECTED)


def test_generate_workbook_deletes_only_newly_published_file_when_verification_fails(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    template = tmp_path / "template.xls"
    template.write_bytes(b"template")
    existing = tmp_path / "existing.xlsx"
    existing.write_bytes(b"user workbook")
    published = tmp_path / "published.xlsx"

    def create_temporary(template_path, temporary_path, header, products) -> None:
        temporary_path.write_bytes(b"generated workbook")

    def publish(temporary_path: Path, output_dir: Path, filename: str) -> Path:
        temporary_path.replace(published)
        return published

    def fail_verification(path: Path, header: ShipmentHeader, products: list[ProductRow]) -> None:
        assert path == published
        raise VerificationError("INVOICE 总金额不匹配")

    monkeypatch.setattr(generator, "_invoke_worker", create_temporary)
    monkeypatch.setattr(generator, "_publish_without_overwrite", publish)
    monkeypatch.setattr(generator, "verify_workbook", fail_verification, raising=False)

    product = ProductRow.empty()
    with pytest.raises(VerificationError, match="INVOICE 总金额"):
        generator.generate_workbook(template, tmp_path, ShipmentHeader(), [product])

    assert not published.exists()
    assert existing.read_bytes() == b"user workbook"


def test_verify_totals_closes_workbook_when_validation_fails(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    path = _workbook_path(tmp_path)
    real_workbook = openpyxl.load_workbook(path, data_only=True)
    closed = False
    real_close = real_workbook.close

    def close() -> None:
        nonlocal closed
        closed = True
        real_close()

    monkeypatch.setattr(real_workbook, "close", close)
    monkeypatch.setattr(output_verifier.openpyxl, "load_workbook", lambda *args, **kwargs: real_workbook)
    real_workbook.worksheets[1][f"G{table_product_rows(2)}"] = -1

    with pytest.raises(VerificationError):
        verify_totals(path, 2, EXPECTED)

    assert closed
