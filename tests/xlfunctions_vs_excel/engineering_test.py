from ..testing import workbook_test_cases, assert_equivalent

@workbook_test_cases("BASES.xlsx")
def test_bases(sheet_value, calculated_value):
    assert_equivalent(result=calculated_value, expected=sheet_value)
