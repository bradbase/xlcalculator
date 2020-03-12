
# Excel reference: https://support.office.com/en-us/article/MATCH-function-e8dffd45-c762-47d6-bf89-533f4a37673a


from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError
from ..koala_types import XLRange

class Match(KoalaBaseFunction):
    """"""

    def match(self, lookup_value, lookup_range, match_type=1):
        """"""

        raise Exception("MATCH DOESN'T WORK, XLRANGE DOESN'T SUPPORT .values")

        if not isinstance(lookup_range, XLRange):
            return ExcelError('#VALUE!', 'Lookup_range is not a Range')

        def type_convert(value):
            if type(value) == str:
                value = value.lower()

            elif type(value) == int:
                value = float(value)

            elif value is None:
                value = 0

            return value;

        lookup_value = type_convert(lookup_value)

        range_values = [x for x in lookup_range.values if x is not None] # filter None values to avoid asc/desc order errors
        range_length = len(range_values)

        if match_type == 1:
            # Verify ascending sort

            posMax = -1
            for i in range(range_length):
                current = type_convert(range_values[i])

                if i is not range_length-1 and current > type_convert(range_values[i+1]):
                    return ExcelError('#VALUE!', 'for match_type 1, lookup_range must be sorted ascending')

                if current <= lookup_value:
                    posMax = i
            if posMax == -1:
                return ExcelError('#VALUE!','no result in lookup_range for match_type 1')

            return posMax +1 #Excel starts at 1

        elif match_type == 0:
            # No string wildcard
            try:

                return [type_convert(x) for x in range_values].index(lookup_value) + 1

            except:
                return ExcelError('#VALUE!', '%s not found' % lookup_value)

        elif match_type == -1:
            # Verify descending sort
            posMin = -1
            for i in range((range_length)):
                current = type_convert(range_values[i])

                if i is not range_length-1 and current < type_convert(range_values[i+1]):
                   return ExcelError('#VALUE!','for match_type -1, lookup_range must be sorted descending')

                if current >= lookup_value:
                   posMin = i

            if posMin == -1:
                return ExcelError('#VALUE!', 'no result in lookup_range for match_type -1')

            return posMin +1 #Excel starts at 1
