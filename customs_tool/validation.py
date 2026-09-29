from datetime import date
from decimal import Decimal

from customs_tool.domain import ProductRow, ShipmentHeader


HEADER_FIELDS = {
    "domestic_shipper": "境内发货人",
    "overseas_consignee": "境外收货人",
    "production_sales_unit": "生产销售单位",
    "contract_no": "合同协议号",
    "invoice_no": "发票号",
    "invoice_date": "发票日期",
    "trade_country": "贸易国",
    "destination_country": "运抵国",
    "destination_port": "指运港",
    "exit_port": "离境口岸",
    "transport_mode": "运输方式",
    "incoterm": "成交方式",
    "packaging_type": "包装种类",
    "loading_port": "装运港",
}

PRODUCT_TEXT_FIELDS = {
    "chinese_name": "中文品名",
    "english_name": "英文品名",
    "sku": "SKU/规格型号",
    "hs_code": "HS编码",
    "quantity_unit": "数量单位",
    "declaration_elements": "申报要素",
    "source_location": "货源地",
}

PRODUCT_NUMBER_FIELDS = {
    "quantity": "数量",
    "boxes": "箱数",
    "unit_price": "单价",
    "total_price": "总价",
    "net_weight": "净重",
    "gross_weight": "毛重",
    "volume": "体积",
}


def _is_invalid_text(value: object) -> bool:
    return not isinstance(value, str) or not value.strip()


def validate_submission(
    header: ShipmentHeader,
    products: list[ProductRow],
) -> list[str]:
    errors: list[str] = []

    for field_name, label in HEADER_FIELDS.items():
        value = getattr(header, field_name)
        if field_name == "invoice_date":
            if value is None:
                errors.append(f"{label}不能为空")
            elif type(value) is not date:
                errors.append(f"{label}格式无效")
        elif _is_invalid_text(value):
            errors.append(f"{label}不能为空")

    if not products:
        errors.append("至少需要一个产品")
        return errors

    for row_number, product in enumerate(products, start=1):
        for field_name, label in PRODUCT_TEXT_FIELDS.items():
            value = getattr(product, field_name)
            if field_name == "source_location" and not _is_invalid_text(header.source_location):
                continue
            if _is_invalid_text(value):
                errors.append(f"第{row_number}行：{label}不能为空")

        valid_numbers: dict[str, Decimal] = {}
        for field_name, label in PRODUCT_NUMBER_FIELDS.items():
            value = getattr(product, field_name)
            if not isinstance(value, Decimal) or not value.is_finite():
                errors.append(f"第{row_number}行：{label}不是有效数字")
            else:
                valid_numbers[field_name] = value
                if value < 0:
                    errors.append(f"第{row_number}行：{label}不能为负数")

        if {"quantity", "unit_price", "total_price"} <= valid_numbers.keys():
            calculated_total = valid_numbers["quantity"] * valid_numbers["unit_price"]
            if abs(calculated_total - valid_numbers["total_price"]) > Decimal("0.01"):
                errors.append(f"第{row_number}行：总价与数量×单价不一致")

    return errors
