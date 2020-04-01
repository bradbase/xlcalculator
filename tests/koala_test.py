from unittest import TestCase

from decimal import Decimal
from decimal import ROUND_UP
from decimal import ROUND_DOWN


class KoalaTestCase(TestCase):

    def assertEqualRounded(self, lhs, rhs, rounding_precision = None):

        if rounding_precision is None:
            lhs_split = str(lhs).split('.')
            rhs_split = str(rhs).split('.')

            if len(lhs_split) > 1:
                len_lhs_after_decimal = len(lhs_split[1])
            else:
                len_lhs_after_decimal = None

            if len(rhs_split) > 1:
                len_rhs_after_decimal = len(rhs_split[1])
            else:
                len_rhs_after_decimal = None

            if len_lhs_after_decimal is None or len_rhs_after_decimal is None:
                return self.assertEqual( round(lhs), round(rhs) )

            rounding_precision = min(len_lhs_after_decimal, len_rhs_after_decimal)

        precision_mask = "{0:." + str(rounding_precision - 1) + "f}1"
        precision = precision_mask.format(0.0)

        if lhs > rhs:
            lhs_value = Decimal(lhs).quantize(Decimal(precision), rounding=ROUND_DOWN)
            rhs_value = Decimal(rhs).quantize(Decimal(precision), rounding=ROUND_UP)
        else:
            lhs_value = Decimal(rhs).quantize(Decimal(precision), rounding=ROUND_UP)
            rhs_value = Decimal(lhs).quantize(Decimal(precision), rounding=ROUND_DOWN)

        return self.assertEqual( lhs_value, rhs_value )


    def assertEqualTruncated(self, lhs, rhs, truncating_places = None):

        lhs_before_dec, lhs_after_dec = str(lhs).split('.')
        rhs_before_dec, rhs_after_dec = str(rhs).split('.')

        if truncating_places is None:
            truncating_places = min(len(str(lhs).split('.')[1]), len(str(rhs).split('.')[1]))

        if 'E' in lhs_after_dec:
            lhs_value = float('.'.join((lhs_before_dec, lhs_after_dec)))
        else:
            lhs_value = float('.'.join((lhs_before_dec, lhs_after_dec[0:truncating_places])))

        if 'E' in lhs_after_dec:
            rhs_value = float('.'.join((rhs_before_dec, rhs_after_dec)))
        else:
            rhs_value = float('.'.join((rhs_before_dec, rhs_after_dec[0:truncating_places])))

        return self.assertAlmostEqual( lhs_value, rhs_value, truncating_places )
