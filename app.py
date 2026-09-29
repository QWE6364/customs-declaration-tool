from __future__ import annotations

from dataclasses import asdict
from datetime import date
from decimal import Decimal, InvalidOperation
import hashlib
import json
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any

import pandas as pd
import streamlit as st

from customs_tool.domain import ProductRow, ShipmentHeader
from customs_tool.errors import CustomsToolError
from customs_tool.excel_generator import generate_workbook
from customs_tool.totals import calculate_totals
from customs_tool.validation import validate_submission


APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = APP_DIR / "templates" / "报关资料模板.xls"
EDITOR_KEY = "products_editor"
PRODUCT_FRAME_KEY = "product_frame"
GENERATED_FILE_KEY = "generated_file"

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
CORE_IDENTITY_COLUMNS = (
    "中文品名",
    "英文品名",
    "SKU/规格型号",
    "HS编码",
    "申报要素",
)

TEXT_FIELDS = {
    "中文品名": "chinese_name",
    "英文品名": "english_name",
    "SKU/规格型号": "sku",
    "HS编码": "hs_code",
    "数量单位": "quantity_unit",
    "申报要素": "declaration_elements",
    "货源地": "source_location",
}
NUMBER_FIELDS = {
    "数量": "quantity",
    "箱数": "boxes",
    "单价": "unit_price",
    "总价": "total_price",
    "净重": "net_weight",
    "毛重": "gross_weight",
    "体积": "volume",
}

DEFAULT_DOMESTIC_SHIPPER = ""
DEFAULT_OVERSEAS_CONSIGNEE = ""
DEFAULT_PRODUCTION_SALES_UNIT = ""
DEFAULT_TRADE_COUNTRY = ""
DEFAULT_DESTINATION_COUNTRY = ""
DEFAULT_DESTINATION_PORT = ""
DEFAULT_EXIT_PORT = ""
DEFAULT_TRANSPORT_MODE = "水路运输"
DEFAULT_INCOTERM = "FOB"
DEFAULT_PACKAGING_TYPE = "22-纸制或纤维板制盒/箱"
DEFAULT_SOURCE_LOCATION = ""
DEFAULT_LOADING_PORT = ""


def _is_blank(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    try:
        missing = pd.isna(value)
    except (TypeError, ValueError):
        return False
    return bool(missing) if isinstance(missing, bool) else False


def decimal_value(value: object) -> Decimal:
    """Convert an editor cell to a finite Decimal; blank cells become zero."""
    if value is None or (isinstance(value, str) and not value.strip()) or value is pd.NA:
        return Decimal("0")
    if isinstance(value, bool):
        raise ValueError("不是有效数字")
    try:
        converted = value if isinstance(value, Decimal) else Decimal(str(value).strip())
    except (InvalidOperation, TypeError, ValueError) as exc:
        raise ValueError("不是有效数字") from exc
    if not converted.is_finite():
        raise ValueError("不是有效数字")
    return converted


def _text_value(value: object) -> str:
    if _is_blank(value):
        return ""
    return str(value).strip()


def _numeric_cell_is_blank_or_zero(value: object) -> bool:
    if _is_blank(value):
        return True
    try:
        return decimal_value(value) == 0
    except ValueError:
        return False


def _row_is_empty(row: pd.Series) -> bool:
    return (
        all(_is_blank(row.get(column)) for column in CORE_IDENTITY_COLUMNS)
        and all(_numeric_cell_is_blank_or_zero(row.get(column)) for column in NUMBER_FIELDS)
    )


def frame_to_products(frame: pd.DataFrame) -> list[ProductRow]:
    """Convert non-empty data-editor rows to domain ProductRow objects."""
    products: list[ProductRow] = []
    for display_row, (_, row) in enumerate(frame.iterrows(), start=1):
        if _row_is_empty(row):
            continue

        values: dict[str, Any] = {
            field_name: _text_value(row.get(column))
            for column, field_name in TEXT_FIELDS.items()
        }
        for column, field_name in NUMBER_FIELDS.items():
            try:
                values[field_name] = decimal_value(row.get(column))
            except ValueError as exc:
                raise ValueError(f"第{display_row}行：{column}不是有效数字") from exc
        products.append(ProductRow(**values))
    return products


def _json_value(value: object) -> object:
    if isinstance(value, Decimal):
        return format(value.normalize(), "f") if value else "0"
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, dict):
        return {key: _json_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_value(item) for item in value]
    return value


def submission_fingerprint(header: ShipmentHeader, products: list[ProductRow]) -> str:
    payload = {
        "header": _json_value(asdict(header)),
        "products": _json_value([asdict(product) for product in products]),
    }
    serialized = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(serialized.encode("utf-8")).hexdigest()


def cached_file_is_current(cached_file: object, current_fingerprint: str | None) -> bool:
    return (
        isinstance(cached_file, dict)
        and current_fingerprint is not None
        and cached_file.get("fingerprint") == current_fingerprint
        and isinstance(cached_file.get("bytes"), bytes)
        and isinstance(cached_file.get("name"), str)
    )


def generate_file_data(
    template_path: Path,
    header: ShipmentHeader,
    products: list[ProductRow],
    *,
    generator=generate_workbook,
) -> dict[str, object]:
    try:
        with TemporaryDirectory(prefix="customs_tool_") as temporary_dir:
            output_path = generator(template_path, Path(temporary_dir), header, products)
            return {
                "bytes": output_path.read_bytes(),
                "name": output_path.name,
                "fingerprint": submission_fingerprint(header, products),
            }
    except OSError as exc:
        raise CustomsToolError(f"无法读取或保存临时生成文件：{exc}") from exc


def initial_product_frame() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "中文品名": "",
                "英文品名": "",
                "SKU/规格型号": "",
                "HS编码": "",
                "数量": 0.0,
                "数量单位": "件",
                "箱数": 0.0,
                "单价": 0.0,
                "总价": 0.0,
                "净重": 0.0,
                "毛重": 0.0,
                "体积": 0.0,
                "申报要素": "",
                "货源地": DEFAULT_SOURCE_LOCATION,
            }
        ],
        columns=PRODUCT_COLUMNS,
    )


def _column_config() -> dict[str, Any]:
    numeric_columns = {
        label: st.column_config.NumberColumn(
            label,
            min_value=0.0,
            default=0.0,
            format="%.4f" if label == "体积" else "%.2f",
        )
        for label in NUMBER_FIELDS
    }
    return {
        "中文品名": st.column_config.TextColumn("中文品名"),
        "英文品名": st.column_config.TextColumn("英文品名"),
        "SKU/规格型号": st.column_config.TextColumn("SKU/规格型号"),
        "HS编码": st.column_config.TextColumn(
            "HS编码",
            help="按文本录入，可保留前导零。",
        ),
        **numeric_columns,
        "数量单位": st.column_config.TextColumn("数量单位", default="件"),
        "申报要素": st.column_config.TextColumn("申报要素", width="large"),
        "货源地": st.column_config.TextColumn("货源地", default=DEFAULT_SOURCE_LOCATION),
    }


def _reset_products() -> None:
    st.session_state[PRODUCT_FRAME_KEY] = initial_product_frame()
    st.session_state.pop(EDITOR_KEY, None)
    st.session_state.pop(GENERATED_FILE_KEY, None)


def _display_decimal(value: Decimal) -> str:
    if value == value.to_integral():
        return format(value, "f").split(".")[0]
    return format(value.normalize(), "f")


def _build_header(
    *,
    domestic_shipper: str,
    overseas_consignee: str,
    production_sales_unit: str,
    contract_no: str,
    invoice_no: str,
    invoice_date: date,
    trade_country: str,
    destination_country: str,
    destination_port: str,
    exit_port: str,
    transport_mode: str,
    incoterm: str,
    packaging_type: str,
    source_location: str,
    loading_port: str,
) -> ShipmentHeader:
    return ShipmentHeader(
        domestic_shipper=domestic_shipper.strip(),
        overseas_consignee=overseas_consignee.strip(),
        production_sales_unit=production_sales_unit.strip(),
        contract_no=contract_no.strip(),
        invoice_no=invoice_no.strip(),
        invoice_date=invoice_date,
        trade_country=trade_country.strip(),
        destination_country=destination_country.strip(),
        destination_port=destination_port.strip(),
        exit_port=exit_port.strip(),
        transport_mode=transport_mode.strip(),
        incoterm=incoterm.strip(),
        packaging_type=packaging_type.strip(),
        loading_port=loading_port.strip(),
        source_location=source_location.strip(),
    )


def main() -> None:
    st.set_page_config(page_title="报关资料填写工具", layout="wide")
    st.title("报关资料填写工具")
    st.caption("填写本票基础信息和产品明细，校验无误后生成可下载的报关资料工作簿。")

    if PRODUCT_FRAME_KEY not in st.session_state:
        st.session_state[PRODUCT_FRAME_KEY] = initial_product_frame()

    st.subheader("基础信息")
    row1 = st.columns(3)
    domestic_shipper = row1[0].text_input("境内发货人", value=DEFAULT_DOMESTIC_SHIPPER)
    overseas_consignee = row1[1].text_input("境外收货人", value=DEFAULT_OVERSEAS_CONSIGNEE)
    production_sales_unit = row1[2].text_input("生产销售单位", value=DEFAULT_PRODUCTION_SALES_UNIT)

    row2 = st.columns(3)
    contract_no = row2[0].text_input("合同协议号")
    invoice_no = row2[1].text_input("发票号")
    invoice_date = row2[2].date_input("发票日期", value=date.today())

    row3 = st.columns(4)
    trade_country = row3[0].text_input("贸易国", value=DEFAULT_TRADE_COUNTRY)
    destination_country = row3[1].text_input("运抵国", value=DEFAULT_DESTINATION_COUNTRY)
    destination_port = row3[2].text_input("指运港", value=DEFAULT_DESTINATION_PORT)
    exit_port = row3[3].text_input("离境口岸", value=DEFAULT_EXIT_PORT)

    row4 = st.columns(5)
    transport_mode = row4[0].text_input("运输方式", value=DEFAULT_TRANSPORT_MODE)
    incoterm = row4[1].text_input("成交方式", value=DEFAULT_INCOTERM)
    packaging_type = row4[2].text_input("包装种类", value=DEFAULT_PACKAGING_TYPE)
    source_location = row4[3].text_input("默认货源地", value=DEFAULT_SOURCE_LOCATION)
    loading_port = row4[4].text_input("装运港", value=DEFAULT_LOADING_PORT)

    st.subheader("产品明细")
    st.caption("可在表格底部添加或删除行；HS编码按文本处理，所有数值列只允许非负数。")
    edited_frame = st.data_editor(
        st.session_state[PRODUCT_FRAME_KEY],
        key=EDITOR_KEY,
        num_rows="dynamic",
        hide_index=True,
        column_order=PRODUCT_COLUMNS,
        column_config=_column_config(),
        width="stretch",
    )

    header = _build_header(
        domestic_shipper=domestic_shipper,
        overseas_consignee=overseas_consignee,
        production_sales_unit=production_sales_unit,
        contract_no=contract_no,
        invoice_no=invoice_no,
        invoice_date=invoice_date,
        trade_country=trade_country,
        destination_country=destination_country,
        destination_port=destination_port,
        exit_port=exit_port,
        transport_mode=transport_mode,
        incoterm=incoterm,
        packaging_type=packaging_type,
        source_location=source_location,
        loading_port=loading_port,
    )
    try:
        products = frame_to_products(edited_frame)
        totals = calculate_totals(products)
        current_fingerprint = submission_fingerprint(header, products)
        preview_error = None
    except ValueError as exc:
        products = []
        totals = calculate_totals([])
        current_fingerprint = None
        preview_error = str(exc)

    if not cached_file_is_current(st.session_state.get(GENERATED_FILE_KEY), current_fingerprint):
        st.session_state.pop(GENERATED_FILE_KEY, None)

    st.subheader("汇总预览")
    metrics = st.columns(6)
    metrics[0].metric("总数量", _display_decimal(totals.quantity))
    metrics[1].metric("总箱数", _display_decimal(totals.boxes))
    metrics[2].metric("总金额（USD）", _display_decimal(totals.amount))
    metrics[3].metric("总净重（kg）", _display_decimal(totals.net_weight))
    metrics[4].metric("总毛重（kg）", _display_decimal(totals.gross_weight))
    metrics[5].metric("总体积（m³）", _display_decimal(totals.volume))
    if preview_error:
        st.warning(preview_error)

    generate_col, clear_col = st.columns([1, 1])
    generate_clicked = generate_col.button("生成报关资料", type="primary")
    clear_col.button("清空产品", on_click=_reset_products)

    if generate_clicked:
        st.session_state.pop(GENERATED_FILE_KEY, None)
        try:
            products = frame_to_products(edited_frame)
            validation_errors = validate_submission(header, products)
            if validation_errors:
                for error in validation_errors:
                    st.error(error)
            else:
                with st.spinner("正在生成并复核工作簿，请稍候……"):
                    st.session_state[GENERATED_FILE_KEY] = generate_file_data(
                        TEMPLATE_PATH,
                        header,
                        products,
                    )
                st.success("报关资料已生成并通过汇总复核。")
        except CustomsToolError as exc:
            st.error(str(exc))
        except (TypeError, ValueError, InvalidOperation) as exc:
            st.error(f"输入转换失败：{exc}")

    generated_file = st.session_state.get(GENERATED_FILE_KEY)
    if cached_file_is_current(generated_file, current_fingerprint):
        st.download_button(
            "下载报关资料",
            data=generated_file["bytes"],
            file_name=generated_file["name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
