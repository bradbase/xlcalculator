import datetime
import mock
import unittest

from xlcalculator.xlfunctions import date, xlerrors, func_xltypes, utils

dt = datetime.datetime


class DateModuleTest(unittest.TestCase):

    def test_DATE(self):
        self.assertEqual(date.DATE(2000, 1, 1), dt(2000, 1, 1))
        self.assertEqual(date.DATE(2008, 11, 3), dt(2008, 11, 3))
        self.assertEqual(date.DATE(2024, 1, 1), dt(2024, 1, 1))
        self.assertEqual(date.DATE(2025, 1, 1), dt(2025, 1, 1))
        self.assertEqual(date.DATE(2026, 1, 1), dt(2026, 1, 1))
        self.assertEqual(date.DATE(1911, 1, 15), dt(1911, 1, 15))

    def test_DATE_with_short_year(self):
        self.assertEqual(date.DATE(99, 1, 1), dt(1999, 1, 1))

    def test_DATE_with_string_components(self):
        self.assertEqual(date.DATE('2026', '1', '1'), dt(2026, 1, 1))

    def test_DATE_year_must_be_number(self):
        self.assertIsInstance(date.DATE('bad', 1, 1), xlerrors.ValueExcelError)

    def test_DATE_month_must_be_number(self):
        self.assertIsInstance(
            date.DATE(2000, 'bad', 1), xlerrors.ValueExcelError)

    def test_DATE_day_must_be_integer(self):
        self.assertIsInstance(
            date.DATE(2000, 1, 'bad'), xlerrors.ValueExcelError)

    def test_DATE_year_must_be_positive(self):
        self.assertIsInstance(date.DATE(-1, 1, 1), xlerrors.NumExcelError)

    def test_DATE_year_must_have_less_than_10000(self):
        self.assertIsInstance(date.DATE(10000, 1, 1), xlerrors.NumExcelError)

    def test_DATE_result_must_be_positive(self):
        self.assertIsInstance(date.DATE(1900, 1, -1), xlerrors.NumExcelError)

    def test_DATE_not_stricly_positive_month_substracts(self):
        self.assertEqual(date.DATE(2009, -1, 1), dt(2008, 11, 1))

    def test_DATE_not_stricly_positive_day_substracts(self):
        self.assertEqual(date.DATE(2009, 1, -1), dt(2008, 12, 30))

    def test_DATE_month_superior_to_12_change_year(self):
        self.assertEqual(date.DATE(2009, 14, 1), dt(2010, 2, 1))

    def test_DATE_day_superior_to_365_change_year(self):
        self.assertEqual(date.DATE(2009, 1, 400), dt(2010, 2, 4))

    def test_DATEDIF_year(self):
        start_date = date.DATE(2011, 1, 1)
        end_date = date.DATE(2011, 12, 31)

        self.assertEqual(date.DATEDIF(start_date, end_date, 'Y'), 0)

        start_date = date.DATE(2011, 1, 1)
        end_date = date.DATE(2012, 12, 31)
        self.assertEqual(date.DATEDIF(start_date, end_date, 'Y'), 1)

    def test_DATEDIF_month(self):
        start_date = date.DATE(2011, 1, 1)
        end_date = date.DATE(2011, 12, 31)

        self.assertEqual(date.DATEDIF(start_date, end_date, 'M'), 11)

    def test_DATEDIF_day(self):
        start_date = date.DATE(2001, 6, 1)
        end_date = date.DATE(2002, 8, 15)

        self.assertEqual(date.DATEDIF(start_date, end_date, 'D'), 440)

    def test_DATEDIF_MD(self):
        start_date = date.DATE(2011, 1, 1)
        end_date = date.DATE(2011, 12, 31)

        self.assertEqual(date.DATEDIF(start_date, end_date, 'MD'), 30)

    def test_DATEDIF_YM(self):
        start_date = date.DATE(2011, 1, 1)
        end_date = date.DATE(2011, 12, 31)

        self.assertEqual(date.DATEDIF(start_date, end_date, 'YM'), 11)

    def test_DATEDIF_YD(self):
        start_date = date.DATE(2001, 6, 1)
        end_date = date.DATE(2002, 8, 15)

        self.assertEqual(date.DATEDIF(start_date, end_date, 'YD'), 75)

    def test_DATEDIF_M_start_beyond_end(self):
        start_date = date.DATE(2002, 8, 15)
        end_date = date.DATE(2001, 6, 1)

        self.assertIsInstance(date.DATEDIF(start_date, end_date, 'M'),
                              xlerrors.NumExcelError)

    def test_DAY(self):
        sample_date = date.DATE(1911, 4, 15)
        serial_number = int(sample_date)

        self.assertEqual(date.DAY(serial_number), 15)

    def test_DAYS(self):
        end_date = date.DATE(2011, 12, 31)
        start_date = date.DATE(2011, 1, 1)

        self.assertEqual(date.DAYS(end_date, start_date), 364)

    def test_EDATE(self):
        the_date = date.DATE(2011, 1, 15)
        target_serial_number = utils.datetime_to_number(dt(2011, 2, 15))
        self.assertEqual(date.EDATE(the_date, 1), target_serial_number)

    def test_EDATE_must_be_greater_than_epoch(self):
        the_date = date.DATE(1901, 1, 15)
        self.assertIsInstance(date.EDATE(the_date, -24),
                              xlerrors.NumExcelError)

    def test_EOMONTH(self):
        the_date = date.DATE(2011, 1, 1)
        target_serial_number = utils.datetime_to_number(dt(2011, 2, 28))
        self.assertEqual(date.EOMONTH(the_date, 1), target_serial_number)

    def test_EOMONTH_must_be_greater_than_epoch(self):
        the_date = date.DATE(1901, 1, 1)
        self.assertIsInstance(date.EOMONTH(the_date, -24),
                              xlerrors.NumExcelError)

    def test_ISOWEEKNUM(self):
        sample_date = date.DATE(2012, 3, 9)

        self.assertEqual(date.ISOWEEKNUM(sample_date), 10)

    def test_MONTH(self):
        sample_date = date.DATE(1911, 4, 15)
        serial_number = int(sample_date)

        self.assertEqual(date.MONTH(serial_number), 4)

    def test_NOW(self):
        with mock.patch.object(date, 'now', lambda: dt(2000, 1, 1, 12, 30)):
            self.assertEqual(
                date.NOW(), utils.datetime_to_number(dt(2000, 1, 1, 12, 30))
            )

    def test_TODAY(self):
        with mock.patch.object(date, 'now', lambda: dt(2000, 1, 1, 12, 30)):
            self.assertEqual(
                date.TODAY(), utils.datetime_to_number(dt(2000, 1, 1))
            )

    def test_WEEKDAY_return_type_omitted(self):
        sample_date = date.DATE(2008, 2, 14)
        serial_number = int(sample_date)

        self.assertEqual(date.WEEKDAY(serial_number), 5)

    def test_WEEKDAY_return_type_1(self):
        sample_date = date.DATE(2008, 2, 14)
        serial_number = int(sample_date)

        self.assertEqual(date.WEEKDAY(serial_number, 1), 5)

    def test_WEEKDAY_return_type_2(self):
        sample_date = date.DATE(2008, 2, 14)
        serial_number = int(sample_date)

        self.assertEqual(date.WEEKDAY(serial_number, 2), 4)

    def test_WEEKDAY_return_type_3(self):
        sample_date = date.DATE(2008, 2, 14)
        serial_number = int(sample_date)

        self.assertEqual(date.WEEKDAY(serial_number, 3), 3)

    def test_WEEKDAY_return_type_11(self):
        sample_date = date.DATE(2008, 2, 14)
        serial_number = int(sample_date)

        self.assertEqual(date.WEEKDAY(serial_number, 11), 4)

    def test_WEEKDAY_return_type_12(self):
        sample_date = date.DATE(2008, 2, 14)
        serial_number = int(sample_date)

        self.assertEqual(date.WEEKDAY(serial_number, 12), 3)

    def test_WEEKDAY_return_type_13(self):
        sample_date = date.DATE(2008, 2, 14)
        serial_number = int(sample_date)

        self.assertEqual(date.WEEKDAY(serial_number, 13), 2)

    def test_WEEKDAY_return_type_14(self):
        sample_date = date.DATE(2008, 2, 14)
        serial_number = int(sample_date)

        self.assertEqual(date.WEEKDAY(serial_number, 14), 1)

    def test_WEEKDAY_return_type_15(self):
        sample_date = date.DATE(2008, 2, 14)
        serial_number = int(sample_date)

        self.assertEqual(date.WEEKDAY(serial_number, 15), 7)

    def test_WEEKDAY_return_type_16(self):
        sample_date = date.DATE(2008, 2, 14)
        serial_number = int(sample_date)

        self.assertEqual(date.WEEKDAY(serial_number, 16), 6)

    def test_WEEKDAY_return_type_17(self):
        sample_date = date.DATE(2008, 2, 14)
        serial_number = int(sample_date)

        self.assertEqual(date.WEEKDAY(serial_number, 17), 5)

    def test_WEEKDAY_error(self):
        sample_date = date.DATE(2008, 2, 14)
        serial_number = int(sample_date)

        self.assertIsInstance(date.WEEKDAY(serial_number, 0),
                              xlerrors.NumExcelError)
        self.assertIsInstance(date.WEEKDAY(serial_number, 4),
                              xlerrors.NumExcelError)
        self.assertIsInstance(date.WEEKDAY(serial_number, 5),
                              xlerrors.NumExcelError)
        self.assertIsInstance(date.WEEKDAY(serial_number, 6),
                              xlerrors.NumExcelError)
        self.assertIsInstance(date.WEEKDAY(serial_number, 7),
                              xlerrors.NumExcelError)
        self.assertIsInstance(date.WEEKDAY(serial_number, 8),
                              xlerrors.NumExcelError)
        self.assertIsInstance(date.WEEKDAY(serial_number, 9),
                              xlerrors.NumExcelError)
        self.assertIsInstance(date.WEEKDAY(serial_number, 10),
                              xlerrors.NumExcelError)
        self.assertIsInstance(date.WEEKDAY(serial_number, 18),
                              xlerrors.NumExcelError)

    def test_YEAR(self):
        sample_date_2008 = date.DATE(2008, 5, 7)
        serial_number_2008 = int(sample_date_2008)
        self.assertEqual(date.YEAR(serial_number_2008), 2008)

        sample_date_2010 = date.DATE(2010, 5, 7)
        serial_number_2010 = int(sample_date_2010)
        self.assertEqual(date.YEAR(serial_number_2010), 2010)

    def test_YEAR_before_epoch(self):
        # is text as a "date" entry into Excel of a date pre-epoch
        # is stored as datatype General
        sample_date = func_xltypes.XlText('12/31/1899')

        self.assertIsInstance(
            date.YEAR(sample_date), xlerrors.ValueExcelError)

    def test_YEAR_after_9999(self):
        # is text as a "date" entry into Excel of a date pre-epoch
        # is stored as datatype General
        sample_date = func_xltypes.XlText('1/1/10000')

        self.assertIsInstance(
            date.YEAR(sample_date), xlerrors.ValueExcelError)

    def test_YEARFRAC_start_date_must_be_datetime(self):
        self.assertIsInstance(
            date.YEARFRAC('bad', 1), xlerrors.ValueExcelError)

    def test_YEARFRAC_end_date_must_be_datetime(self):
        self.assertIsInstance(
            date.YEARFRAC(1, 'bad'), xlerrors.ValueExcelError)

    def test_YEARFRAC_start_date_must_after_epoch(self):
        self.assertIsInstance(
            date.YEARFRAC(dt(1800, 1, 1), 0), xlerrors.ValueExcelError)

    def test_YEARFRAC_end_date_must_be_after_epoch(self):
        self.assertIsInstance(
            date.YEARFRAC(dt(1900, 1, 1), -1), xlerrors.ValueExcelError)

    def test_YEARFRAC_basis_must_be_between_0_and_4(self):
        self.assertIsInstance(
            date.YEARFRAC(1, 2, 5), xlerrors.ValueExcelError)
        self.assertIsInstance(
            date.YEARFRAC(1, 2, -1), xlerrors.ValueExcelError)

    def test_YEARFRAC_yearfrac_basis_0(self):
        # From Excel docs.
        self.assertAlmostEqual(
            date.YEARFRAC(dt(2012, 1, 1), dt(2012, 7, 30), 0),
            0.58055556)
        self.assertAlmostEqual(
            date.YEARFRAC(dt(2008, 1, 1), dt(2015, 4, 20)),
            7.30277777777778)
        self.assertAlmostEqual(
            date.YEARFRAC(dt(2008, 1, 1), dt(2015, 4, 20), 0),
            7.30277777777778)
        self.assertAlmostEqual(
            date.YEARFRAC(dt(2024, 1, 1), dt(2025, 1, 1), 0), 1)

    def test_YEARFRAC_yearfrac_basis_1(self):
        # From Excel docs.
        self.assertAlmostEqual(
            date.YEARFRAC(dt(2012, 1, 1), dt(2012, 7, 30), 1),
            0.57650273)

    def test_YEARFRAC_yearfrac_basis_2(self):
        self.assertAlmostEqual(
            date.YEARFRAC(dt(2008, 1, 1), dt(2015, 4, 20), 2),
            7.405555556)
        self.assertAlmostEqual(
            date.YEARFRAC(dt(2024, 1, 1), dt(2025, 1, 1), 2),
            1.01666666666667)

    def test_YEARFRAC_yearfrac_basis_3(self):
        # From Excel docs.
        self.assertAlmostEqual(
            date.YEARFRAC(dt(2012, 1, 1), dt(2012, 7, 30), 3),
            0.57808219)
        self.assertAlmostEqual(
            date.YEARFRAC(dt(2008, 1, 1), dt(2015, 4, 20), 3),
            7.304109589)
        self.assertAlmostEqual(
            date.YEARFRAC(dt(2024, 1, 1), dt(2025, 1, 1), 3),
            1.0027397260274)

    def test_YEARFRAC_yearfrac_basis_4(self):
        self.assertAlmostEqual(
            date.YEARFRAC(dt(2008, 1, 1), dt(2015, 4, 20), 4),
            7.302777778)
        self.assertAlmostEqual(
            date.YEARFRAC(dt(2024, 1, 1), dt(2025, 1, 1), 4), 1)

    def test_YEARFRAC_yearfrac_inverted(self):
        self.assertAlmostEqual(
            date.YEARFRAC(dt(2015, 4, 20), dt(2008, 1, 1)),
            date.YEARFRAC(dt(2008, 1, 1), dt(2015, 4, 20)))
