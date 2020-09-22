from . import xl, xlerrors, func_xltypes


@xl.register()
@xl.validate_args
def CHOOSE(
        index_num: func_xltypes.XlNumber,
        *values,
) -> func_xltypes.XlAnything:
    """Uses index_num to return a value from the list of value arguments.

    https://support.office.com/en-us/article/
        choose-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc
    """
    if index_num <= 0 or index_num > 254:
        raise xlerrors.ValueExcelError(
            f"`index_num` {index_num} must be between 1 and 254")

    if index_num > len(values):
        raise xlerrors.ValueExcelError(
            f"`index_num` {index_num} must not be larger than the number of "
            f"values: {len(values)}")

    idx = int(index_num) - 1
    return values[idx]


@xl.register()
@xl.validate_args
def VLOOKUP(
        lookup_value: func_xltypes.XlAnything,
        table_array: func_xltypes.XlArray,
        col_index_num: func_xltypes.XlNumber,
        range_lookup=False
) -> func_xltypes.XlAnything:
    """Looks in the first column of an array and moves across the row to
    return the value of a cell.

    https://support.office.com/en-us/article/
        vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1
    """
    if range_lookup:
        raise NotImplementedError("Excact match only supported at the moment.")

    col_index_num = int(col_index_num)

    if col_index_num > len(table_array.values[0]):
        raise xlerrors.ValueExcelError(
            'col_index_num is greater than the number of cols in table_array')

    table_array = table_array.set_index(0)

    if lookup_value not in table_array.index:
        raise xlerrors.NaExcelError(
            '`lookup_value` not in first column of `table_array`.')

    return table_array.loc[lookup_value].values[0]


@xl.register()
@xl.validate_args
def MATCH(
        lookup_value: func_xltypes.XlAnything,
        lookup_array: func_xltypes.XlArray,
        match_type: func_xltypes.XlAnything = 1,
) -> func_xltypes.XlAnything:
    assert len(lookup_array.values[0]) == 1

    lookup_array = lookup_array.flat

    if match_type == 1:
        if lookup_array != sorted(lookup_array):
            return xlerrors.NaExcelError(
                "Values must be sorted in ascending order"
            )
    if match_type == -1:
        if lookup_array != sorted(lookup_array, reverse=True):
            return xlerrors.NaExcelError(
                "Values must be sorted in descending order"
            )

    for i, val in enumerate(lookup_array):
        if val == lookup_value:
            return i + 1
        if match_type == 1 and val > lookup_value:
            return i or xlerrors.NaExcelError(
                "No lesser value found."
            )
        if match_type == -1 and val < lookup_value:
            return i or xlerrors.NaExcelError(
                "No greater value found."
            )
    return xlerrors.NaExcelError("No match found.")
