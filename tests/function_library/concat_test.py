
import unittest

import pandas as pd

from koala_xlcalculator.function_library import Concat


class TestConcat(unittest.TestCase):

    # def setUp(self):
    #     pass

    # def teardown(self):
    #     pass

    def test_concat(self):

        concat_result_00 = Concat.concat("SPAM", " ", "SPAM", " ", "SPAM", " ", "SPAM")
        result_00 = "SPAM SPAM SPAM SPAM"
        self.assertTrue(result_00, concat_result_00)

        concat_result_01 = Concat.concat("SPAM", " ", pd.DataFrame([[1, 2],[3, 4]]), " ", "SPAM", " ", "SPAM")
        result_01 = "SPAM 1234 SPAM SPAM"
        self.assertTrue(result_01, concat_result_01)

        concat_result_02 = Concat.concat("SPAM", "SPAM", "SPAM")
        result_02 = "SPAMSPAMSPAM"
        self.assertTrue(result_02, concat_result_02)
