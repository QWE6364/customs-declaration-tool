from __future__ import annotations

from decimal import Decimal, InvalidOperation
from pathlib import Path

import openpyxl

from customs_tool.domain import ProductRow, ShipmentHeader, Totals
from customs_tool.errors import VerificationError
from customs_tool.layout import DECLARATION_FIRST_PRODUCT_ROW, TABLE_FIRST_PRODUCT_ROW, table_product_rows
from customs_tool.totals import calculate_totals


_TOLERANCE = Decimal("0.000001")


def _decimal(value: object, label: str) -> Decimal:
    try:
        return Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError) as exc:
        raise VerificationError(f"{label}无法读取：{value!r}") from exc


def _check(actual: object, expected: Decimal, label: str) -> None:
    if abs(_decimal(actual, label) - expected) > _TOLERANCE:
        raise VerificationError(f"{label}不匹配：实际 {actual!r}，预期 {expected}")


def _check_text(actual: object, expected: str, label: str) -> None:
    if actual != expected:
        raise VerificationError(f"{label}不匹配：实际 {actual!r}，预期 {expected!r}")


def verify_totals(path: Path, product_count: int, expected: Totals) -> None:
    workbook = None
    try:
        workbook = openpyxl.load_workbook(Path(path), data_only=True)
        if len(workbook.worksheets) < 3:
            raise VerificationError("工作表不足：至少需要3张工作表")

        invoice = workbook.worksheets[1]
        packing = workbook.worksheets[2]
        total_row = table_product_rows(product_count)

        if invoice[f"B{total_row}"].value != "TOTAL":
            raise VerificationError("INVOICE 总计行缺失")
        if packing[f"B{total_row}"].value != "TOTAL":
            raise VerificationError("PACKING 总计行缺失")

        _check(invoice[f"D{total_row}"].value, expected.quantity, "INVOICE 数量")
        _check(invoice[f"E{total_row}"].value, expected.boxes, "INVOICE 箱数")
        _check(invoice[f"G{total_row}"].value, expected.amount, "INVOICE 总金额")
        _check(packing[f"D{total_row}"].value, expected.quantity, "PACKING 数量")
        _check(packing[f"E{total_row}"].value, expected.net_weight, "PACKING 总净重")
        _check(packing[f"F{total_row}"].value, expected.gross_weight, "PACKING 总毛重")
        _check(packing[f"G{total_row}"].value, expected.volume, "PACKING 总体积")
    except VerificationError:
        raise
    except FileNotFoundError as exc:
        raise VerificationError(f"复核文件不存在：{path}") from exc
    except (IndexError, KeyError) as exc:
        raise VerificationError(f"生成结果结构异常：{exc}") from exc
    except Exception as exc:
        raise VerificationError(f"生成结果复核失败：{exc}") from exc
    finally:
        if workbook is not None:
            workbook.close()


def verify_workbook(path: Path, header: ShipmentHeader, products: list[ProductRow]) -> None:
    verify_totals(path, len(products), calculate_totals(products))

    workbook = None
    try:
        workbook = openpyxl.load_workbook(Path(path), data_only=True)
        if len(workbook.worksheets) < 3:
            raise VerificationError("工作表不足：至少需要3张工作表")

        declaration, invoice, packing = workbook.worksheets[:3]
        _check_text(declaration["A10"].value, header.contract_no, "报关资料 A10 合同号")
        totals = calculate_totals(products)
        _check(declaration["I12"].value, totals.boxes, "报关资料 I12 总箱数")
        _check(declaration["K12"].value, totals.gross_weight, "报关资料 K12 总毛重")
        _check(declaration["M12"].value, totals.net_weight, "报关资料 M12 总净重")
        _check_text(declaration["U10"].value, header.exit_port, "报关资料 U10 离境口岸")

        for index, product in enumerate(products):
            main_row = DECLARATION_FIRST_PRODUCT_ROW + index * 3
            gross_row = main_row + 1
            boxes_row = main_row + 2
            _check_text(declaration[f"B{main_row}"].value, product.hs_code, f"报关资料 第{main_row}行 HS编码")
            _check_text(declaration[f"E{main_row}"].value, product.chinese_name, f"报关资料 第{main_row}行 中文名")
            _check_text(declaration[f"F{main_row}"].value, product.sku, f"报关资料 第{main_row}行 SKU")
            _check(declaration[f"I{main_row}"].value, product.quantity, f"报关资料 第{main_row}行 数量")
            _check_text(declaration[f"J{main_row}"].value, product.quantity_unit, f"报关资料 第{main_row}行 单位")
            _check(declaration[f"L{main_row}"].value, product.unit_price, f"报关资料 第{main_row}行 单价")
            _check(declaration[f"N{main_row}"].value, product.total_price, f"报关资料 第{main_row}行 总价")
            _check_text(declaration[f"U{main_row}"].value, header.destination_country, f"报关资料 第{main_row}行 目的国")
            source_location = product.source_location or header.source_location
            _check_text(declaration[f"W{main_row}"].value, source_location, f"报关资料 第{main_row}行 货源地")
            _check(declaration[f"I{gross_row}"].value, product.gross_weight, f"报关资料 第{gross_row}行 毛重")
            _check_text(declaration[f"J{gross_row}"].value, "千克", f"报关资料 第{gross_row}行 毛重单位")
            _check(declaration[f"I{boxes_row}"].value, product.boxes, f"报关资料 第{boxes_row}行 箱数")
            _check_text(declaration[f"J{boxes_row}"].value, "箱", f"报关资料 第{boxes_row}行 箱数单位")

            table_row = TABLE_FIRST_PRODUCT_ROW + index
            _check_text(invoice[f"C{table_row}"].value, product.english_name, f"INVOICE 第{table_row}行 英文名")
            _check(invoice[f"D{table_row}"].value, product.quantity, f"INVOICE 第{table_row}行 数量")
            _check(invoice[f"E{table_row}"].value, product.boxes, f"INVOICE 第{table_row}行 箱数")
            _check(invoice[f"F{table_row}"].value, product.unit_price, f"INVOICE 第{table_row}行 单价")
            _check(invoice[f"G{table_row}"].value, product.total_price, f"INVOICE 第{table_row}行 总价")

            _check_text(
                packing[f"C{table_row}"].value,
                f"{product.english_name}\n{product.chinese_name}",
                f"PACKING 第{table_row}行 中英文名称",
            )
            _check(packing[f"D{table_row}"].value, product.quantity, f"PACKING 第{table_row}行 数量")
            _check(packing[f"E{table_row}"].value, product.net_weight, f"PACKING 第{table_row}行 净重")
            _check(packing[f"F{table_row}"].value, product.gross_weight, f"PACKING 第{table_row}行 毛重")
            _check(packing[f"G{table_row}"].value, product.volume, f"PACKING 第{table_row}行 体积")
            _check_text(packing[f"H{table_row}"].value, product.hs_code, f"PACKING 第{table_row}行 HS编码")
            _check_text(packing[f"I{table_row}"].value, product.declaration_elements, f"PACKING 第{table_row}行 申报要素")
    except VerificationError:
        raise
    except FileNotFoundError as exc:
        raise VerificationError(f"复核文件不存在：{path}") from exc
    except (IndexError, KeyError) as exc:
        raise VerificationError(f"生成结果结构异常：{exc}") from exc
    except Exception as exc:
        raise VerificationError(f"生成结果复核失败：{exc}") from exc
    finally:
        if workbook is not None:
            workbook.close()
