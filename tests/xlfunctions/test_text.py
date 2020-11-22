import unittest

from xlcalculator.xlfunctions import text, xl, xlerrors, func_xltypes


class TextModuleTest(unittest.TestCase):

    def test_CONCAT(self):
        self.assertEqual(
            text.CONCAT("SPAM", " ", "SPAM"), 'SPAM SPAM')
        self.assertEqual(
            text.CONCAT(
                "SPAM",
                " ",
                func_xltypes.Array([[1, 2]]),
                4
            ),
            'SPAM 124'
        )

    def test_CONCAT_with_too_many_items(self):
        self.assertIsInstance(
            text.CONCAT(*[0] * 300),
            xlerrors.ValueExcelError
        )

    def test_CONCAT_MS_cases(self):

        self.assertEqual(
            text.CONCAT(
                "The",
                " ",
                "sun",
                " ",
                "will",
                " ",
                "come",
                " ",
                "up",
                " ",
                "tomorrow."
            ),
            "The sun will come up tomorrow.")

        range0 = func_xltypes.Array(
            ["A's", 'a1', 'a2', 'a4', 'a5', 'a6', 'a7'])
        range1 = func_xltypes.Array(
            ["B's", 'b1', 'b2', 'b4', 'b5', 'b6', 'b7'])

        self.assertEqual(text.CONCAT(range0, range1),
                         "A'sa1a2a4a5a6a7B'sb1b2b4b5b6b7")

    def test_EXACT(self):
        self.assertEqual(text.EXACT("word", "word"), True)
        self.assertEqual(text.EXACT("Word", "word"), False)
        self.assertEqual(text.EXACT("w ord", "word"), False)

    def test_FIND(self):
        self.assertEqual(text.FIND('M', 'Miriam McGovern'), 1)
        self.assertEqual(text.FIND('m', 'Miriam McGovern'), 6)
        self.assertEqual(text.FIND('M', 'Miriam McGovern', 3), 8)

    def test_FIND_ValueError(self):
        self.assertIsInstance(
            text.FIND('B', 'Miriam McGovern'), xlerrors.ValueExcelError)

    def test_LEFT(self):
        self.assertEqual(text.LEFT('Sale Price', 4), "Sale")
        self.assertEqual(text.LEFT('Sweden'), "S")

    def test_LEN(self):
        self.assertEqual(text.LEN("Phoenix, AZ"), 11)
        self.assertEqual(text.LEN(""), 0)
        self.assertEqual(text.LEN("        One"), 11)

    def test_LOWER(self):
        self.assertEqual(text.LOWER("E. E. Cummings"), "e. e. cummings")
        self.assertEqual(text.LOWER("Apt. 2B"), "apt. 2b")

    def test_MID(self):
        self.assertEqual(text.MID('Romain', 3, 4), 'main')
        self.assertEqual(text.MID('Romain', 1, 2), 'Ro')
        self.assertEqual(text.MID('Romain', 3, 6), 'main')

    def test_MID_start_num_converted_to_integer(self):
        self.assertEqual(text.MID('Romain', 1.1, 2), 'Ro')

    def test_MID_num_chars_converted_to_integer(self):
        self.assertEqual(text.MID('Romain', 1, 2.1), 'Ro')

    def test_MID_start_num_must_be_superior_or_equal_to_1(self):
        self.assertIsInstance(
            text.MID('Romain', 0, 3), xlerrors.NumExcelError)

    def test_MID_num_chars_must_be_positive(self):
        self.assertIsInstance(
            text.MID('Romain', 1, -1), xlerrors.NumExcelError)

    def test_MID_with_too_large_text(self):
        self.assertIsInstance(
            text.MID('foo' + ' ' * xl.CELL_CHARACTER_LIMIT, 1, 3),
            xlerrors.ValueExcelError)

    def test_REPLACE(self):
        self.assertEqual(text.REPLACE("abcdefghijk", 6, 5, "*"), "abcde*k")
        self.assertEqual(text.REPLACE("2009", 3, 2, "10"), "2010")
        self.assertEqual(text.REPLACE("123456", 1, 3, "@"), "@456")

    def test_RIGHT(self):
        self.assertEqual(text.RIGHT("Sale Price", 5), "Price")
        self.assertEqual(text.RIGHT("Stock Number"), "r")

    def test_TRIM(self):
        self.assertEqual(text.TRIM(" First Quarter Earnings "),
                         "First Quarter Earnings")

    def test_UPPER(self):
        self.assertEqual(text.UPPER("total"), "TOTAL")
        self.assertEqual(text.UPPER("Yield"), "YIELD")
