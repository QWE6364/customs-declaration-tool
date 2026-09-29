from __future__ import annotations

import hashlib
import os
import time
from datetime import date
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest
import pywintypes

import customs_tool.excel_generator as generator
from customs_tool.domain import ProductRow, ShipmentHeader
from customs_tool.errors import GenerationError, TemplateNotFoundError, VerificationError, WorkbookInUseError
from customs_tool.output_verifier import verify_totals, verify_workbook
from customs_tool.totals import calculate_totals


pytestmark = [pytest.mark.windows_com, pytest.mark.skipif(os.name != "nt", reason="Windows COM only")]


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _template() -> Path:
    return Path(__file__).parents[1] / "templates" / "报关资料模板.xls"


def _header(contract_no: str = "TEST20260929") -> ShipmentHeader:
    return ShipmentHeader(
        domestic_shipper="示例境内发货人有限公司",
        overseas_consignee="EXAMPLE BUYER LTD.",
        production_sales_unit="示例生产单位有限公司",
        contract_no=contract_no,
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


def _product(index: int = 1) -> ProductRow:
    return ProductRow(
        chinese_name=f"示例产品{index}",
        english_name=f"Example Product {index}",
        sku=f"SKU-{index:03d}",
        hs_code=f"940360999{index}",
        quantity=Decimal(str(index * 10)),
        quantity_unit="件",
        boxes=Decimal(str(index * 2)),
        unit_price=Decimal("12.50"),
        total_price=Decimal(str(index * 125)),
        net_weight=Decimal(str(index * 80)),
        gross_weight=Decimal(str(index * 90)),
        volume=Decimal(str(index * Decimal("0.5"))),
        declaration_elements=f"品牌类型：无品牌；用途：示例用途；序号：{index}",
        source_location="示例货源地",
    )


def _temporary_outputs(directory: Path) -> list[Path]:
    return [path for path in directory.iterdir() if ".tmp." in path.name]


def _worker_raises_runtime_error(template_path, temporary_path, header, products) -> None:
    raise RuntimeError("remote runtime failure")


def _worker_raises_permission_error(template_path, temporary_path, header, products) -> None:
    raise PermissionError("remote access denied")


def _worker_sleeps_forever(template_path, temporary_path, header, products) -> None:
    temporary_path.write_bytes(b"partial workbook")
    while True:
        time.sleep(1)


def test_com_generate_disables_macros_and_events_before_open(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    events: list[tuple[str, object]] = []

    class FakeWorkbooks:
        def Open(self, *args, **kwargs):
            events.append(("open", None))
            raise RuntimeError("stop after open")

    class FakeExcel:
        def __init__(self) -> None:
            object.__setattr__(self, "Workbooks", FakeWorkbooks())

        def __setattr__(self, name, value) -> None:
            events.append((name, value))
            object.__setattr__(self, name, value)

        def Quit(self) -> None:
            events.append(("quit", None))

    monkeypatch.setattr(generator.pythoncom, "CoInitialize", lambda: None)
    monkeypatch.setattr(generator.pythoncom, "CoUninitialize", lambda: None)
    monkeypatch.setattr(generator, "DispatchEx", lambda _name: FakeExcel())

    with pytest.raises(RuntimeError, match="stop after open"):
        generator._com_generate(
            _template(),
            tmp_path / "unused.xlsx",
            _header("SECURITY"),
            [_product()],
        )

    open_index = events.index(("open", None))
    assert events.index(("AutomationSecurity", 3)) < open_index
    assert events.index(("EnableEvents", False)) < open_index


def test_invoke_worker_timeout_ends_owned_child_and_reports_timeout(tmp_path: Path) -> None:
    temporary_path = tmp_path / "timeout.xlsx"
    before_children = {child.pid for child in generator.multiprocessing.active_children()}

    with pytest.raises(GenerationError, match="超时"):
        generator._invoke_worker(
            _template(),
            temporary_path,
            _header("TIMEOUT"),
            [_product()],
            worker_callable=_worker_sleeps_forever,
            timeout_seconds=0.05,
        )

    after_children = {child.pid for child in generator.multiprocessing.active_children()}
    assert after_children <= before_children


def test_generate_workbook_timeout_removes_temp_and_final_outputs(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    real_invoke = generator._invoke_worker
    header = _header("TIMEOUT-CLEANUP")
    expected_final = tmp_path / header.output_name(date.today())

    def invoke_timeout_worker(template_path, temporary_path, shipment_header, products):
        return real_invoke(
            template_path,
            temporary_path,
            shipment_header,
            products,
            worker_callable=_worker_sleeps_forever,
            timeout_seconds=0.05,
        )

    monkeypatch.setattr(generator, "_invoke_worker", invoke_timeout_worker)

    with pytest.raises(GenerationError, match="超时"):
        generator.generate_workbook(_template(), tmp_path, header, [_product()])

    assert not expected_final.exists()
    assert _temporary_outputs(tmp_path) == []
    assert list(tmp_path.iterdir()) == []


def test_generate_single_product_preserves_template_and_writes_required_fields(
    tmp_path: Path,
    capfd: pytest.CaptureFixture[str],
) -> None:
    template = _template()
    original_hash = _sha256(template)
    original_mtime = template.stat().st_mtime_ns

    output = generator.generate_workbook(template, tmp_path, _header(), [_product()])
    captured = capfd.readouterr()

    assert captured.err == ""
    assert output.exists()
    assert output.suffix == ".xlsx"
    assert output.parent == tmp_path
    assert output.name == _header().output_name(date.today())
    assert _temporary_outputs(tmp_path) == []

    workbook = openpyxl.load_workbook(output, data_only=False)
    try:
        assert len(workbook.worksheets) == 3
        declaration, invoice, packing = workbook.worksheets
        assert declaration["U10"].value == "示例口岸"
        assert declaration["S18"].value == "中国"
        assert declaration["U18"].value == "示例国"
        assert declaration["W18"].value == "示例货源地"
        assert declaration["I12"].value == 2
        assert declaration["K12"].value == 90
        assert declaration["M12"].value == 80
        assert invoice["C12"].value == "Example Product 1"
        assert invoice["B13"].value == "TOTAL"
        assert invoice["D13"].value == 10
        assert invoice["E13"].value == 2
        assert invoice["G13"].value == 125
        assert packing["H12"].value == "9403609991"
        assert packing["H12"].number_format == "@"
        assert packing["A13"].value is None
        assert packing["B13"].value == "TOTAL"
        assert packing["D13"].value == 10
        assert packing["E13"].value == 80
        assert packing["F13"].value == 90
        assert packing["G13"].value == 0.5
    finally:
        workbook.close()

    assert _sha256(template) == original_hash
    assert template.stat().st_mtime_ns == original_mtime


def test_five_products_expand_all_sheets_and_place_totals_on_row_17(tmp_path: Path) -> None:
    products = [_product(index) for index in range(1, 6)]

    output = generator.generate_workbook(_template(), tmp_path, _header("FIVE"), products)

    workbook = openpyxl.load_workbook(output, data_only=False)
    try:
        declaration, invoice, packing = workbook.worksheets
        assert declaration["E30"].value == "示例产品5"
        assert declaration["B30"].value == "9403609995"
        assert declaration["W30"].value == "示例货源地"
        assert declaration["I12"].value == 30
        assert declaration["K12"].value == 1350
        assert declaration["M12"].value == 1200
        assert invoice["C16"].value == "Example Product 5"
        assert invoice["B17"].value == "TOTAL"
        assert invoice["D17"].value == 150
        assert invoice["E17"].value == 30
        assert invoice["G17"].value == 1875
        assert packing["C16"].value == "Example Product 5\n示例产品5"
        assert packing["H16"].value == "9403609995"
        assert packing["B17"].value == "TOTAL"
        assert packing["D17"].value == 150
        assert packing["E17"].value == 1200
        assert packing["F17"].value == 1350
        assert packing["G17"].value == 7.5
    finally:
        workbook.close()


def test_six_products_expand_all_sheets_move_footer_and_verify_totals(tmp_path: Path) -> None:
    template = _template()
    original_hash = _sha256(template)
    original_mtime = template.stat().st_mtime_ns
    products = [_product(index) for index in range(1, 7)]
    expected_totals = calculate_totals(products)

    output = generator.generate_workbook(template, tmp_path, _header("SIX"), products)

    assert output.exists()
    assert output.suffix == ".xlsx"
    assert output.parent == tmp_path
    assert _temporary_outputs(tmp_path) == []
    verify_totals(output, len(products), expected_totals)

    workbook = openpyxl.load_workbook(output, data_only=False)
    try:
        assert len(workbook.worksheets) == 3
        declaration, invoice, packing = workbook.worksheets
        assert declaration["E33"].value == "示例产品6"
        assert declaration["B33"].value == "9403609996"
        assert declaration["W33"].value == "示例货源地"
        assert declaration["I12"].value == 42
        assert declaration["K12"].value == 1890
        assert declaration["M12"].value == 1680
        assert declaration["A37"].value == "申报要素 & 内包装种类"
        assert declaration["A38"].value == "特殊关系确认："
        assert declaration["A39"].value == "报关人员"
        assert invoice["C17"].value == "Example Product 6"
        assert invoice["B18"].value == "TOTAL"
        assert invoice["D18"].value == 210
        assert invoice["E18"].value == 42
        assert invoice["G18"].value == 2625
        assert packing["C17"].value == "Example Product 6\n示例产品6"
        assert packing["H17"].value == "9403609996"
        assert packing["B18"].value == "TOTAL"
        assert packing["D18"].value == 210
        assert packing["E18"].value == 1680
        assert packing["F18"].value == 1890
        assert packing["G18"].value == 10.5
        assert packing["B21"].value == "境内货源地"
        assert packing["C21"].value == "示例货源地"
    finally:
        workbook.close()

    assert _sha256(template) == original_hash
    assert template.stat().st_mtime_ns == original_mtime


def test_eleven_products_preserve_template_and_all_dynamic_sections(tmp_path: Path) -> None:
    template = _template()
    original_hash = _sha256(template)
    original_mtime = template.stat().st_mtime_ns
    products = [_product(index) for index in range(1, 12)]
    header = _header("ELEVEN")
    totals = calculate_totals(products)

    output = generator.generate_workbook(template, tmp_path, header, products)

    verify_workbook(output, header, products)
    workbook = openpyxl.load_workbook(output, data_only=False)
    try:
        declaration, invoice, packing = workbook.worksheets
        assert declaration["E48"].value == "示例产品11"
        assert declaration["B48"].value == "94036099911"
        assert declaration["W48"].value == "示例货源地"
        assert declaration["A52"].value == "申报要素 & 内包装种类"
        assert declaration["A53"].value == "特殊关系确认："
        assert declaration["A54"].value == "报关人员"
        assert invoice["C22"].value == "Example Product 11"
        assert invoice["B23"].value == "TOTAL"
        assert invoice["D23"].value == float(totals.quantity)
        assert invoice["E23"].value == float(totals.boxes)
        assert invoice["G23"].value == float(totals.amount)
        assert packing["C22"].value == "Example Product 11\n示例产品11"
        assert packing["H22"].value == "94036099911"
        assert packing["B23"].value == "TOTAL"
        assert packing["D23"].value == float(totals.quantity)
        assert packing["E23"].value == float(totals.net_weight)
        assert packing["F23"].value == float(totals.gross_weight)
        assert packing["G23"].value == float(totals.volume)
        assert packing["B26"].value == "境内货源地"
        assert packing["C26"].value == "示例货源地"
    finally:
        workbook.close()

    assert _temporary_outputs(tmp_path) == []
    assert _sha256(template) == original_hash
    assert template.stat().st_mtime_ns == original_mtime


def test_verify_workbook_rejects_tampered_generated_product_cell(tmp_path: Path) -> None:
    products = [_product(1), _product(2)]
    header = _header("TAMPERED")
    output = generator.generate_workbook(_template(), tmp_path, header, products)
    workbook = openpyxl.load_workbook(output)
    workbook.worksheets[2]["I13"] = "tampered"
    workbook.save(output)
    workbook.close()

    with pytest.raises(VerificationError, match="PACKING.*13.*申报要素"):
        verify_workbook(output, header, products)


def test_second_generation_uses_suffix_without_overwriting_first(tmp_path: Path) -> None:
    header = _header("DUPLICATE")
    first = generator.generate_workbook(_template(), tmp_path, header, [_product()])
    first_hash = _sha256(first)

    second = generator.generate_workbook(_template(), tmp_path, header, [_product(2)])

    assert second.name == f"{first.stem}_2.xlsx"
    assert first.exists() and second.exists()
    assert _sha256(first) == first_hash
    assert _temporary_outputs(tmp_path) == []


def test_file_created_at_final_publish_is_preserved_and_next_suffix_is_used(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    header = _header("RACE")
    first_path = tmp_path / header.output_name(date.today())
    real_move = generator._move_no_replace
    raced = False

    def create_user_file_before_first_publish(source: Path, destination: Path) -> None:
        nonlocal raced
        if not raced and destination == first_path:
            raced = True
            first_path.write_bytes(b"user workbook")
            raise FileExistsError(destination)
        real_move(source, destination)

    monkeypatch.setattr(generator, "_move_no_replace", create_user_file_before_first_publish)

    output = generator.generate_workbook(_template(), tmp_path, header, [_product()])

    assert raced
    assert first_path.read_bytes() == b"user workbook"
    assert output.name == f"{first_path.stem}_2.xlsx"
    assert output.exists()
    assert _temporary_outputs(tmp_path) == []


def test_publish_failure_removes_temporary_file_and_leaves_no_final_xlsx(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    header = _header("PUBLISHFAIL")
    expected_final = tmp_path / header.output_name(date.today())

    def fail_move(source: Path, destination: Path) -> None:
        raise OSError("injected move failure")

    monkeypatch.setattr(generator, "_move_no_replace", fail_move)

    with pytest.raises(GenerationError, match="无法发布输出文件"):
        generator.generate_workbook(_template(), tmp_path, header, [_product()])

    assert not expected_final.exists()
    assert _temporary_outputs(tmp_path) == []
    assert list(tmp_path.iterdir()) == []


def test_output_directory_that_is_a_file_maps_to_generation_error(tmp_path: Path) -> None:
    output_file = tmp_path / "not-a-directory"
    output_file.write_text("user data", encoding="utf-8")

    with pytest.raises(GenerationError):
        generator.generate_workbook(_template(), output_file, _header(), [_product()])

    assert output_file.read_text(encoding="utf-8") == "user data"


def test_missing_template_raises_template_not_found(tmp_path: Path) -> None:
    with pytest.raises(TemplateNotFoundError):
        generator.generate_workbook(tmp_path / "missing.xls", tmp_path, _header(), [_product()])

    assert list(tmp_path.iterdir()) == []


def test_failed_worker_removes_partial_temporary_output_and_never_creates_final(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    header = _header("FAILURE")
    expected_final = tmp_path / header.output_name(date.today())

    def fail_after_partial_save(template_path, temporary_path, shipment_header, products):
        temporary_path.write_bytes(b"partial xlsx")
        raise GenerationError("injected SaveAs failure")

    monkeypatch.setattr(generator, "_invoke_worker", fail_after_partial_save)

    with pytest.raises(GenerationError, match="injected SaveAs failure"):
        generator.generate_workbook(_template(), tmp_path, header, [_product()])

    assert not expected_final.exists()
    assert _temporary_outputs(tmp_path) == []
    assert list(tmp_path.iterdir()) == []


def test_real_worker_protocol_preserves_runtime_error_as_generation_cause(
    tmp_path: Path,
) -> None:
    with pytest.raises(GenerationError, match="remote runtime failure") as captured:
        generator._invoke_worker(
            _template(),
            tmp_path / "remote-runtime.xlsx",
            _header("REMOTE-RUNTIME"),
            [_product()],
            worker_callable=_worker_raises_runtime_error,
        )

    assert isinstance(captured.value.__cause__, RuntimeError)
    assert str(captured.value.__cause__) == "remote runtime failure"


def test_real_worker_protocol_preserves_in_use_business_error(
    tmp_path: Path,
) -> None:
    with pytest.raises(WorkbookInUseError, match="remote access denied"):
        generator._invoke_worker(
            _template(),
            tmp_path / "remote-denied.xlsx",
            _header("REMOTE-DENIED"),
            [_product()],
            worker_callable=_worker_raises_permission_error,
        )


def test_generate_workbook_preserves_real_worker_in_use_protocol(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    real_invoke = generator._invoke_worker

    def invoke_permission_worker(template_path, temporary_path, header, products):
        return real_invoke(
            template_path,
            temporary_path,
            header,
            products,
            worker_callable=_worker_raises_permission_error,
        )

    monkeypatch.setattr(generator, "_invoke_worker", invoke_permission_worker)

    with pytest.raises(WorkbookInUseError, match="remote access denied") as captured:
        generator.generate_workbook(_template(), tmp_path, _header("PUBLIC-DENIED"), [_product()])

    assert isinstance(captured.value.__cause__, PermissionError)
    assert list(tmp_path.iterdir()) == []


def test_generate_workbook_preserves_real_worker_runtime_cause(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    real_invoke = generator._invoke_worker

    def invoke_runtime_worker(template_path, temporary_path, header, products):
        return real_invoke(
            template_path,
            temporary_path,
            header,
            products,
            worker_callable=_worker_raises_runtime_error,
        )

    monkeypatch.setattr(generator, "_invoke_worker", invoke_runtime_worker)

    with pytest.raises(GenerationError, match="remote runtime failure") as captured:
        generator.generate_workbook(_template(), tmp_path, _header("PUBLIC-RUNTIME"), [_product()])

    assert isinstance(captured.value.__cause__, RuntimeError)
    assert list(tmp_path.iterdir()) == []


def test_generate_workbook_wraps_unexpected_worker_error_with_cause(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original = RuntimeError("unexpected worker protocol failure")

    def fail_unexpectedly(template_path, temporary_path, shipment_header, products):
        raise original

    monkeypatch.setattr(generator, "_invoke_worker", fail_unexpectedly)

    with pytest.raises(GenerationError, match="Excel生成失败") as captured:
        generator.generate_workbook(_template(), tmp_path, _header("RUNTIME"), [_product()])

    assert captured.value.__cause__ is original
    assert list(tmp_path.iterdir()) == []


def test_generate_workbook_maps_worker_access_denied_to_workbook_in_use(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    access_denied = pywintypes.com_error(
        -2147352567,
        "Exception occurred.",
        (0, "Microsoft Excel", "Access denied", None, 0, -2147024891),
        None,
    )

    def deny_access(template_path, temporary_path, shipment_header, products):
        raise access_denied

    monkeypatch.setattr(generator, "_invoke_worker", deny_access)

    with pytest.raises(WorkbookInUseError):
        generator.generate_workbook(_template(), tmp_path, _header("DENIED"), [_product()])

    assert list(tmp_path.iterdir()) == []


def test_generate_workbook_maps_worker_com_failure_to_generation_error(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    com_failure = pywintypes.com_error(
        -2147352567,
        "Exception occurred.",
        (0, "Microsoft Excel", "Unexpected COM failure", None, 0, -2146827284),
        None,
    )

    def fail_generation(template_path, temporary_path, shipment_header, products):
        raise com_failure

    monkeypatch.setattr(generator, "_invoke_worker", fail_generation)

    with pytest.raises(GenerationError):
        generator.generate_workbook(_template(), tmp_path, _header("COMFAIL"), [_product()])

    assert list(tmp_path.iterdir()) == []
