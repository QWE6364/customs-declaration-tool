from __future__ import annotations

import ast
from datetime import date
import logging
from pathlib import Path

from streamlit.testing.v1 import AppTest

from app import (
    DEFAULT_DESTINATION_COUNTRY,
    DEFAULT_DESTINATION_PORT,
    DEFAULT_DOMESTIC_SHIPPER,
    DEFAULT_EXIT_PORT,
    DEFAULT_INCOTERM,
    DEFAULT_LOADING_PORT,
    DEFAULT_OVERSEAS_CONSIGNEE,
    DEFAULT_PACKAGING_TYPE,
    DEFAULT_PRODUCTION_SALES_UNIT,
    DEFAULT_SOURCE_LOCATION,
    DEFAULT_TRADE_COUNTRY,
    DEFAULT_TRANSPORT_MODE,
    submission_fingerprint,
)
from customs_tool.domain import ShipmentHeader


logging.getLogger("streamlit.runtime.scriptrunner_utils.script_run_context").setLevel(logging.ERROR)


def load_app() -> AppTest:
    return AppTest.from_file("app.py", default_timeout=10).run()


def test_initial_page_has_no_exception() -> None:
    app = load_app()

    assert list(app.exception) == []
    assert app.title[0].value == "报关资料填写工具"


def test_initial_generate_reports_at_least_one_product() -> None:
    app = load_app()

    app.button[0].click().run()

    assert list(app.exception) == []
    assert "至少需要一个产品" in [error.value for error in app.error]


def test_product_data_editor_uses_dynamic_rows() -> None:
    source = Path("app.py").read_text(encoding="utf-8")
    tree = ast.parse(source)
    data_editor_calls = [
        node
        for node in ast.walk(tree)
        if isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute)
        and isinstance(node.func.value, ast.Name)
        and node.func.value.id == "st"
        and node.func.attr == "data_editor"
    ]

    assert len(data_editor_calls) == 1
    num_rows = next(
        (keyword.value for keyword in data_editor_calls[0].keywords if keyword.arg == "num_rows"),
        None,
    )
    assert isinstance(num_rows, ast.Constant)
    assert num_rows.value == "dynamic"


def test_data_editor_adds_two_products_then_deletes_one_and_updates_totals() -> None:
    app = load_app()
    first_product = {
        "中文品名": "示例产品一",
        "英文品名": "Example Product One",
        "SKU/规格型号": "SKU-001",
        "HS编码": "9403609990",
        "数量": 10.0,
        "数量单位": "件",
        "箱数": 2.0,
        "单价": 12.5,
        "总价": 125.0,
        "净重": 80.0,
        "毛重": 90.0,
        "体积": 0.5,
        "申报要素": "品牌类型：无品牌；用途：示例用途一",
        "货源地": "示例货源地",
    }
    second_product = {
        "中文品名": "示例产品二",
        "英文品名": "Example Product Two",
        "SKU/规格型号": "SKU-002",
        "HS编码": "9403609991",
        "数量": 20.0,
        "数量单位": "件",
        "箱数": 4.0,
        "单价": 15.0,
        "总价": 300.0,
        "净重": 160.0,
        "毛重": 180.0,
        "体积": 1.0,
        "申报要素": "品牌类型：无品牌；用途：示例用途二",
        "货源地": "示例货源地",
    }

    app.session_state["products_editor"] = {
        "edited_rows": {0: first_product},
        "added_rows": [second_product],
        "deleted_rows": [],
    }
    app.run()

    assert list(app.exception) == []
    assert [metric.value for metric in app.metric] == ["30", "6", "425", "240", "270", "1.5"]

    app.session_state["products_editor"] = {
        "edited_rows": {},
        "added_rows": [second_product],
        "deleted_rows": [0],
    }
    app.run()

    assert list(app.exception) == []
    assert [metric.value for metric in app.metric] == ["20", "4", "300", "160", "180", "1"]


def initial_header(contract_no: str = "") -> ShipmentHeader:
    return ShipmentHeader(
        domestic_shipper=DEFAULT_DOMESTIC_SHIPPER,
        overseas_consignee=DEFAULT_OVERSEAS_CONSIGNEE,
        production_sales_unit=DEFAULT_PRODUCTION_SALES_UNIT,
        contract_no=contract_no,
        invoice_no="",
        invoice_date=date.today(),
        trade_country=DEFAULT_TRADE_COUNTRY,
        destination_country=DEFAULT_DESTINATION_COUNTRY,
        destination_port=DEFAULT_DESTINATION_PORT,
        exit_port=DEFAULT_EXIT_PORT,
        transport_mode=DEFAULT_TRANSPORT_MODE,
        incoterm=DEFAULT_INCOTERM,
        packaging_type=DEFAULT_PACKAGING_TYPE,
        loading_port=DEFAULT_LOADING_PORT,
        source_location=DEFAULT_SOURCE_LOCATION,
    )


def preload_current_download(app: AppTest) -> None:
    app.session_state["generated_file"] = {
        "bytes": b"xlsx",
        "name": "报关资料_TEST.xlsx",
        "fingerprint": submission_fingerprint(initial_header(), []),
    }


def test_changed_contract_invalidates_preloaded_download() -> None:
    app = load_app()
    preload_current_download(app)
    app.run()
    assert len(app.get("download_button")) == 1

    app.text_input[3].set_value("NEW-CONTRACT").run()

    assert list(app.exception) == []
    assert len(app.get("download_button")) == 0
    assert "generated_file" not in app.session_state


def test_clear_removes_preloaded_download() -> None:
    app = load_app()
    preload_current_download(app)
    app.run()
    assert len(app.get("download_button")) == 1

    app.button[1].click().run()

    assert list(app.exception) == []
    assert len(app.get("download_button")) == 0
    assert "generated_file" not in app.session_state
