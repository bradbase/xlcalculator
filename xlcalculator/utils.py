import collections
import re
from openpyxl.utils.cell import COORD_RE, SHEET_TITLE
from openpyxl.utils.cell import range_boundaries, get_column_letter

MAX_COL = 18278
MAX_ROW = 1048576


def resolve_sheet(sheet_str):
    sheet_str = sheet_str.strip()
    sheet_match = re.match(SHEET_TITLE.strip(), sheet_str + '!')
    if sheet_match is None:
        # Internally, sheets are not properly quoted, so consider the entire
        # string.
        return sheet_str

    return sheet_match.group("quoted") or sheet_match.group("notquoted")


def resolve_address(addr):
    # Addresses without sheet name are not supported.
    sheet_str, addr_str = addr.split('!')
    sheet = resolve_sheet(sheet_str)
    coord_match = COORD_RE.split(addr_str)
    col, row = coord_match[1:3]
    return sheet, col, row


def resolve_ranges(ranges, default_sheet='Sheet1'):
    sheet = None
    range_cells = collections.defaultdict(set)
    for rng in ranges.split(','):
        # Handle sheets in range.
        if '!' in rng:
            sheet_str, rng = rng.split('!')
            rng_sheet = resolve_sheet(sheet_str)
            if sheet is not None and sheet != rng_sheet:
                raise ValueError(
                    f'Got multiple different sheets in ranges: '
                    f'{sheet}, {rng_sheet}'
                )
            sheet = rng_sheet
        min_col, min_row, max_col, max_row = range_boundaries(rng)

        if (
                min_col is None or min_row is None
                or max_col is None or max_row is None
        ):
            continue

        # Excel ranges are boundaries inclusive!
        for row_idx in range(min_row, max_row + 1):
            row_cells = range_cells[row_idx]
            for col_idx in range(min_col, max_col + 1):
                row_cells.add(col_idx)

    # Now convert the internal structure to a matrix of cell addresses.
    sheet = default_sheet if sheet is None else sheet
    sheet_str = sheet + '!' if sheet else ''
    return sheet, [
        [
            f'{sheet_str}{get_column_letter(col_idx)}{row_idx}'
            for col_idx in sorted(row_cells)
        ]
        for row_idx, row_cells in sorted(range_cells.items())
    ]
