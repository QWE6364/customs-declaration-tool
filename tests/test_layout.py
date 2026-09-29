from customs_tool.layout import declaration_product_rows, table_product_rows


def test_declaration_product_rows_use_three_rows_each():
    assert declaration_product_rows(0) == (18, 19, 20)
    assert declaration_product_rows(4) == (30, 31, 32)


def test_table_product_rows_are_contiguous():
    assert table_product_rows(0) == 12
    assert table_product_rows(5) == 17
