DECLARATION_BASE_CAPACITY = 4
DECLARATION_FIRST_PRODUCT_ROW = 18
DECLARATION_PRODUCT_HEIGHT = 3
DECLARATION_INSERT_ROW = 30
TABLE_BASE_CAPACITY = 4
TABLE_FIRST_PRODUCT_ROW = 12
TABLE_TOTAL_ROW = 16


def declaration_product_rows(index: int) -> tuple[int, int, int]:
    start = DECLARATION_FIRST_PRODUCT_ROW + index * DECLARATION_PRODUCT_HEIGHT
    return start, start + 1, start + 2


def table_product_rows(index: int) -> int:
    return TABLE_FIRST_PRODUCT_ROW + index
