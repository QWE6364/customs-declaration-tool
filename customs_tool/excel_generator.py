from __future__ import annotations

import ctypes
import gc
import multiprocessing
from contextlib import suppress
from datetime import date
from decimal import Decimal
from pathlib import Path
from uuid import uuid4

import pythoncom
import pywintypes
from win32com.client import DispatchEx

from customs_tool.domain import ProductRow, ShipmentHeader
from customs_tool.errors import GenerationError, TemplateNotFoundError, VerificationError, WorkbookInUseError
from customs_tool.output_verifier import verify_workbook
from customs_tool.layout import (
    DECLARATION_BASE_CAPACITY,
    DECLARATION_FIRST_PRODUCT_ROW,
    DECLARATION_INSERT_ROW,
    DECLARATION_PRODUCT_HEIGHT,
    TABLE_BASE_CAPACITY,
    TABLE_FIRST_PRODUCT_ROW,
    TABLE_TOTAL_ROW,
)
from customs_tool.totals import calculate_totals


_XL_DOWN = -4121
_XL_OPEN_XML_WORKBOOK = 51
WORKER_TIMEOUT_SECONDS = 120
_WORKER_STOP_TIMEOUT_SECONDS = 5


def _number(value: Decimal) -> float:
    return float(value)


def _cell(sheet, address: str):
    cell = sheet.Range(address)
    if cell.MergeCells:
        return cell.MergeArea.Cells(1, 1)
    return cell


def _write(sheet, address: str, value, *, text: bool = False) -> None:
    cell = _cell(sheet, address)
    if text:
        cell.NumberFormat = "@"
    cell.Value = value


def _clear(sheet, addresses: tuple[str, ...]) -> None:
    for address in addresses:
        _cell(sheet, address).Value = None


def _copy_insert_rows(sheet, source_start: int, source_end: int, insert_at: int) -> None:
    sheet.Rows(f"{source_start}:{source_end}").Copy()
    sheet.Rows(f"{insert_at}:{insert_at + source_end - source_start}").Insert(Shift=_XL_DOWN)


def _resize_declaration(sheet, product_count: int) -> None:
    if product_count > DECLARATION_BASE_CAPACITY:
        for index in range(product_count - DECLARATION_BASE_CAPACITY):
            insert_at = DECLARATION_INSERT_ROW + index * DECLARATION_PRODUCT_HEIGHT
            _copy_insert_rows(sheet, 27, 29, insert_at)
    elif product_count < DECLARATION_BASE_CAPACITY:
        delete_start = DECLARATION_FIRST_PRODUCT_ROW + product_count * DECLARATION_PRODUCT_HEIGHT
        delete_end = DECLARATION_FIRST_PRODUCT_ROW + DECLARATION_BASE_CAPACITY * DECLARATION_PRODUCT_HEIGHT - 1
        if delete_start <= delete_end:
            sheet.Rows(f"{delete_start}:{delete_end}").Delete()


def _resize_table(sheet, product_count: int) -> int:
    if product_count > TABLE_BASE_CAPACITY:
        for index in range(product_count - TABLE_BASE_CAPACITY):
            _copy_insert_rows(sheet, TABLE_TOTAL_ROW - 1, TABLE_TOTAL_ROW - 1, TABLE_TOTAL_ROW + index)
    elif product_count < TABLE_BASE_CAPACITY:
        delete_start = TABLE_FIRST_PRODUCT_ROW + product_count
        delete_end = TABLE_TOTAL_ROW - 1
        if delete_start <= delete_end:
            sheet.Rows(f"{delete_start}:{delete_end}").Delete()
    return TABLE_FIRST_PRODUCT_ROW + product_count


def _write_declaration(sheet, header: ShipmentHeader, products: list[ProductRow]) -> None:
    _resize_declaration(sheet, len(products))
    totals = calculate_totals(products)

    _clear(
        sheet,
        (
            "A4", "I4", "M4", "Q4", "A6", "I6", "M6", "Q6",
            "A8", "I8", "M8", "Q8", "A10", "I10", "M10", "Q10", "U10",
            "A12", "I12", "K12", "M12", "O12", "Q12", "A14",
        ),
    )
    _write(sheet, "A4", header.domestic_shipper)
    _write(sheet, "A6", header.overseas_consignee)
    _write(sheet, "I6", header.transport_mode)
    _write(sheet, "A8", header.production_sales_unit)
    _write(sheet, "A10", header.contract_no)
    _write(sheet, "I10", header.trade_country)
    _write(sheet, "M10", header.destination_country)
    _write(sheet, "Q10", header.destination_port)
    _write(sheet, "U10", header.exit_port)
    _write(sheet, "A12", header.packaging_type)
    _write(sheet, "I12", _number(totals.boxes))
    _write(sheet, "K12", _number(totals.gross_weight))
    _write(sheet, "M12", _number(totals.net_weight))
    _write(sheet, "O12", header.incoterm)
    _write(sheet, "A14", header.invoice_no)

    for index, product in enumerate(products):
        main_row = DECLARATION_FIRST_PRODUCT_ROW + index * DECLARATION_PRODUCT_HEIGHT
        gross_row = main_row + 1
        second_quantity_row = main_row + 2
        addresses = (
            f"A{main_row}", f"B{main_row}", f"D{main_row}", f"E{main_row}",
            f"F{main_row}", f"I{main_row}", f"J{main_row}", f"L{main_row}",
            f"N{main_row}", f"Q{main_row}", f"S{main_row}", f"U{main_row}",
            f"W{main_row}", f"Y{main_row}", f"A{gross_row}", f"F{gross_row}",
            f"I{gross_row}", f"J{gross_row}", f"I{second_quantity_row}",
            f"J{second_quantity_row}",
        )
        _clear(sheet, addresses)
        _write(sheet, f"A{main_row}", f"{index + 1:02d}", text=True)
        _write(sheet, f"B{main_row}", product.hs_code, text=True)
        _write(sheet, f"E{main_row}", product.chinese_name)
        _write(sheet, f"F{main_row}", product.sku)
        _write(sheet, f"I{main_row}", _number(product.quantity))
        _write(sheet, f"J{main_row}", product.quantity_unit)
        _write(sheet, f"L{main_row}", _number(product.unit_price))
        _write(sheet, f"N{main_row}", _number(product.total_price))
        _write(sheet, f"Q{main_row}", "USD")
        _write(sheet, f"S{main_row}", "中国")
        _write(sheet, f"U{main_row}", header.destination_country)
        _write(sheet, f"W{main_row}", product.source_location or header.source_location)
        details = product.declaration_elements
        if product.source_location:
            details = f"{details}\n货源地：{product.source_location}" if details else f"货源地：{product.source_location}"
        _write(sheet, f"A{gross_row}", details)
        _write(sheet, f"F{gross_row}", product.sku)
        _write(sheet, f"I{gross_row}", _number(product.gross_weight))
        _write(sheet, f"J{gross_row}", "千克")
        _write(sheet, f"I{second_quantity_row}", _number(product.boxes))
        _write(sheet, f"J{second_quantity_row}", "箱")


def _write_invoice(sheet, header: ShipmentHeader, products: list[ProductRow]) -> None:
    total_row = _resize_table(sheet, len(products))
    totals = calculate_totals(products)

    _clear(sheet, ("B1", "B5", "E5", "E6", "E7", "E8", "B9", "E9"))
    _write(sheet, "B1", header.domestic_shipper)
    _write(sheet, "B5", f"BUYER: {header.overseas_consignee}")
    _write(sheet, "E5", header.invoice_no)
    _write(sheet, "E6", header.contract_no)
    _write(sheet, "E7", header.invoice_date.strftime("%Y-%m-%d") if header.invoice_date else "")
    _write(sheet, "E8", header.incoterm)
    _write(sheet, "B9", f"LOADING PORT: {header.loading_port}")
    _write(sheet, "E9", header.destination_port)

    for index, product in enumerate(products):
        row = TABLE_FIRST_PRODUCT_ROW + index
        _clear(sheet, tuple(f"{column}{row}" for column in "BCDEFG"))
        _write(sheet, f"B{row}", "N/M")
        _write(sheet, f"C{row}", product.english_name)
        _write(sheet, f"D{row}", _number(product.quantity))
        _write(sheet, f"E{row}", _number(product.boxes))
        _write(sheet, f"F{row}", _number(product.unit_price))
        _write(sheet, f"G{row}", _number(product.total_price))

    _clear(sheet, tuple(f"{column}{total_row}" for column in "BCDEFG"))
    _write(sheet, f"B{total_row}", "TOTAL")
    _write(sheet, f"D{total_row}", _number(totals.quantity))
    _write(sheet, f"E{total_row}", _number(totals.boxes))
    _write(sheet, f"G{total_row}", _number(totals.amount))


def _packing_source_location(header: ShipmentHeader, products: list[ProductRow]) -> str:
    if header.source_location:
        return header.source_location
    locations = {product.source_location for product in products if product.source_location}
    return locations.pop() if len(locations) == 1 else ""


def _write_packing_list(sheet, header: ShipmentHeader, products: list[ProductRow]) -> None:
    total_row = _resize_table(sheet, len(products))
    totals = calculate_totals(products)

    _clear(sheet, ("B1", "B5", "E5", "E6", "E7", "E8", "B9", "E9"))
    _write(sheet, "B1", f"{header.domestic_shipper}\n{header.domestic_shipper}")
    _write(sheet, "B5", f"BUYER: {header.overseas_consignee}")
    _write(sheet, "E5", header.invoice_no)
    _write(sheet, "E6", header.contract_no)
    _write(sheet, "E7", header.invoice_date.strftime("%Y-%m-%d") if header.invoice_date else "")
    _write(sheet, "E8", header.incoterm)
    _write(sheet, "B9", f"LOADING PORT: {header.loading_port}")
    _write(sheet, "E9", header.destination_port)

    for index, product in enumerate(products):
        row = TABLE_FIRST_PRODUCT_ROW + index
        _clear(sheet, tuple(f"{column}{row}" for column in "BCDEFGHI"))
        _write(sheet, f"B{row}", "N/M")
        _write(sheet, f"C{row}", f"{product.english_name}\n{product.chinese_name}")
        _write(sheet, f"D{row}", _number(product.quantity))
        _write(sheet, f"E{row}", _number(product.net_weight))
        _write(sheet, f"F{row}", _number(product.gross_weight))
        _write(sheet, f"G{row}", _number(product.volume))
        _write(sheet, f"H{row}", product.hs_code, text=True)
        _write(sheet, f"I{row}", product.declaration_elements)

    _clear(sheet, tuple(f"{column}{total_row}" for column in "ABCDEFGHI"))
    _write(sheet, f"B{total_row}", "TOTAL")
    _write(sheet, f"D{total_row}", _number(totals.quantity))
    _write(sheet, f"E{total_row}", _number(totals.net_weight))
    _write(sheet, f"F{total_row}", _number(totals.gross_weight))
    _write(sheet, f"G{total_row}", _number(totals.volume))

    source_row = 19 + len(products) - TABLE_BASE_CAPACITY
    _clear(sheet, (f"C{source_row}",))
    _write(sheet, f"C{source_row}", _packing_source_location(header, products))


def _map_com_error(exc: pywintypes.com_error, output_path: Path) -> Exception:
    details = " ".join(str(part) for part in exc.args).lower()
    access_codes = ("-2147024891", "0x80070005", "access denied", "permission denied", "拒绝访问")
    if any(code in details for code in access_codes):
        return WorkbookInUseError(f"输出文件正在使用或无写入权限：{output_path}")
    return GenerationError(f"Excel生成失败：{exc}")


def _com_generate(
    template_path: Path,
    temporary_path: Path,
    header: ShipmentHeader,
    products: list[ProductRow],
) -> None:
    initialized = False
    excel = None
    workbook = None
    declaration = None
    invoice = None
    packing_list = None
    try:
        pythoncom.CoInitialize()
        initialized = True
        excel = DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(
            str(template_path.resolve()),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
        )
        declaration = workbook.Sheets(1)
        invoice = workbook.Sheets(2)
        packing_list = workbook.Sheets(3)
        _write_declaration(declaration, header, products)
        _write_invoice(invoice, header, products)
        _write_packing_list(packing_list, header, products)
        excel.CutCopyMode = False
        workbook.SaveAs(str(temporary_path.resolve()), FileFormat=_XL_OPEN_XML_WORKBOOK)
        declaration = None
        invoice = None
        packing_list = None
        workbook.Close(SaveChanges=False)
        workbook = None
    finally:
        declaration = None
        invoice = None
        packing_list = None
        if workbook is not None:
            with suppress(Exception):
                workbook.Close(SaveChanges=False)
            workbook = None
        if excel is not None:
            with suppress(Exception):
                excel.CutCopyMode = False
            with suppress(Exception):
                excel.Quit()
            excel = None
        gc.collect()
        if initialized:
            pythoncom.CoUninitialize()


class _RemoteWorkerError(Exception):
    pass


_REMOTE_EXCEPTION_TYPES = {
    ("builtins", "RuntimeError"): RuntimeError,
    ("builtins", "EOFError"): EOFError,
    ("builtins", "PermissionError"): PermissionError,
    ("builtins", "OSError"): OSError,
}


def _exception_message(kind: str, exc: BaseException) -> tuple[str, str, str, str]:
    return kind, exc.__class__.__module__, exc.__class__.__name__, str(exc)


def _remote_exception(module_name: str, type_name: str, message: str) -> BaseException:
    exception_type = _REMOTE_EXCEPTION_TYPES.get((module_name, type_name), _RemoteWorkerError)
    return exception_type(message)


def _worker_entry(
    connection,
    template_path,
    temporary_path,
    header,
    products,
    worker_callable,
) -> None:
    try:
        worker_callable(Path(template_path), Path(temporary_path), header, products)
        connection.send(("ok", "", "", ""))
    except PermissionError as exc:
        connection.send(_exception_message("in_use", exc))
    except pywintypes.com_error as exc:
        mapped = _map_com_error(exc, Path(temporary_path))
        kind = "in_use" if isinstance(mapped, WorkbookInUseError) else "generation"
        connection.send(_exception_message(kind, exc))
    except Exception as exc:
        connection.send(_exception_message("generation", exc))
    finally:
        connection.close()


def _invoke_worker(
    template_path: Path,
    temporary_path: Path,
    header: ShipmentHeader,
    products: list[ProductRow],
    worker_callable=_com_generate,
    *,
    timeout_seconds: float = WORKER_TIMEOUT_SECONDS,
) -> None:
    context = multiprocessing.get_context("spawn")
    parent, child = context.Pipe(duplex=False)
    process = context.Process(
        target=_worker_entry,
        args=(child, str(template_path), str(temporary_path), header, products, worker_callable),
    )
    try:
        process.start()
        child.close()
        process.join(timeout_seconds)
        if process.is_alive():
            process.terminate()
            process.join(_WORKER_STOP_TIMEOUT_SECONDS)
            if process.is_alive():
                process.kill()
                process.join(_WORKER_STOP_TIMEOUT_SECONDS)
            if process.is_alive():
                raise GenerationError("Excel生成超时，子进程无法结束")
            raise GenerationError(f"Excel生成超时（{timeout_seconds}秒）")

        message = (
            parent.recv()
            if parent.poll()
            else ("generation", "builtins", "RuntimeError", f"Excel生成进程异常退出：{process.exitcode}")
        )
    finally:
        child.close()
        parent.close()

    kind, module_name, type_name, detail = message
    if kind == "ok" and process.exitcode == 0:
        return

    cause = _remote_exception(module_name, type_name, detail)
    if kind == "in_use":
        raise WorkbookInUseError(detail) from cause
    raise GenerationError(f"Excel生成失败：{detail}") from cause


def _move_no_replace(source: Path, destination: Path) -> None:
    move_file = ctypes.WinDLL("kernel32", use_last_error=True).MoveFileExW
    move_file.argtypes = (ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_uint32)
    move_file.restype = ctypes.c_int
    if move_file(str(source), str(destination), 0):
        return

    error_code = ctypes.get_last_error()
    if error_code in (80, 183):
        raise FileExistsError(error_code, "目标文件已存在", str(destination))
    if error_code == 5:
        raise PermissionError(error_code, "拒绝访问", str(destination))
    raise OSError(error_code, ctypes.FormatError(error_code), str(destination))


def _publish_without_overwrite(temporary_path: Path, output_dir: Path, filename: str) -> Path:
    name = Path(filename)
    candidate = output_dir / name
    sequence = 2
    while True:
        try:
            _move_no_replace(temporary_path, candidate)
            return candidate
        except FileExistsError:
            candidate = output_dir / f"{name.stem}_{sequence}{name.suffix}"
            sequence += 1


def generate_workbook(
    template_path: Path,
    output_dir: Path,
    header: ShipmentHeader,
    products: list[ProductRow],
) -> Path:
    template_path = Path(template_path)
    output_dir = Path(output_dir)
    if not template_path.is_file():
        raise TemplateNotFoundError(f"模板不存在：{template_path}")

    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except PermissionError as exc:
        raise WorkbookInUseError(f"无法写入输出目录：{output_dir}") from exc
    except OSError as exc:
        raise GenerationError(f"无法创建输出目录：{output_dir}") from exc

    filename = header.output_name(date.today())
    temporary_path = output_dir / f".{filename}.tmp.{uuid4().hex}.xlsx"
    try:
        _invoke_worker(template_path, temporary_path, header, products)
        if not temporary_path.is_file():
            raise GenerationError("Excel生成失败：未生成临时工作簿")
        output_path = _publish_without_overwrite(temporary_path, output_dir, filename)
        try:
            verify_workbook(output_path, header, products)
        except VerificationError as exc:
            try:
                output_path.unlink()
            except FileNotFoundError:
                pass
            except OSError as cleanup_error:
                exc.add_note(f"无法删除复核失败的输出文件：{cleanup_error}")
            raise
        return output_path
    except (TemplateNotFoundError, WorkbookInUseError, GenerationError, VerificationError):
        raise
    except PermissionError as exc:
        raise WorkbookInUseError(f"输出文件正在使用或无写入权限：{output_dir / filename}") from exc
    except pywintypes.com_error as exc:
        raise _map_com_error(exc, output_dir / filename) from exc
    except OSError as exc:
        raise GenerationError(f"无法发布输出文件：{output_dir / filename}") from exc
    except Exception as exc:
        raise GenerationError(f"Excel生成失败：{exc}") from exc
    finally:
        with suppress(FileNotFoundError, PermissionError):
            temporary_path.unlink()
