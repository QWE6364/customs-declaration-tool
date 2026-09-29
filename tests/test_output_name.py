from datetime import date

from customs_tool.domain import ShipmentHeader


def test_output_name_uses_contract_and_date():
    header = ShipmentHeader(contract_no="TEST20260929")
    assert header.output_name(date(2026, 9, 29)) == "报关资料_TEST20260929_20260929.xlsx"


def test_output_name_falls_back_to_date():
    header = ShipmentHeader(contract_no="")
    assert header.output_name(date(2026, 9, 29)) == "报关资料_20260929.xlsx"


def test_output_name_falls_back_for_whitespace_contract():
    header = ShipmentHeader(contract_no="   ")
    assert header.output_name(date(2026, 9, 29)) == "报关资料_20260929.xlsx"


def test_output_name_strips_surrounding_contract_whitespace():
    header = ShipmentHeader(contract_no="  TEST20260929  ")
    assert header.output_name(date(2026, 9, 29)) == "报关资料_TEST20260929_20260929.xlsx"


def test_output_name_replaces_invalid_filename_characters():
    header = ShipmentHeader(contract_no="TEST/2026:09*29")
    assert header.output_name(date(2026, 9, 29)) == "报关资料_TEST-2026-09-29_20260929.xlsx"


def test_output_name_collapses_consecutive_invalid_characters():
    header = ShipmentHeader(contract_no='TEST/\\:*?"<>|20260929')
    assert header.output_name(date(2026, 9, 29)) == "报关资料_TEST-20260929_20260929.xlsx"
