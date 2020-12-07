import unittest

from xlcalculator.xlfunctions import xlcriteria


class XlCriteriaModuleTest(unittest.TestCase):

    def test_parse_criteria(self):
        check = xlcriteria.parse_criteria('>3')
        self.assertTrue(check(4))
        self.assertFalse(check(2))

    def test_parse_criteria_asBool(self):
        check = xlcriteria.parse_criteria('>3')
        self.assertTrue(check(4))
        self.assertFalse(check(2))

    def test_parse_criteria_with_implicit_operator(self):
        # Assumes equality.
        check = xlcriteria.parse_criteria('1')
        self.assertTrue(check(1))
        self.assertFalse(check(2))

    def test_parse_criteria_with_implicit_operator_string_value(self):
        # Assumes equality.
        check = xlcriteria.parse_criteria('data')
        self.assertTrue(check('data'))
        self.assertFalse(check(2))

    def test_parse_criteria_with_simple_number(self):
        check = xlcriteria.parse_criteria(1)
        self.assertTrue(check(1))
        self.assertFalse(check(2))
