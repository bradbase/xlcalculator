
# Excel reference: https://support.office.com/en-us/article/VDB-function-dde4e207-f3fa-488d-91d2-66d55e861d73


from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError
from .sln import SLN

class VDB(KoalaBaseFunction):
    """"""

    def vdb(self, cost, salvage, life, start_period, end_period, factor = 2, no_switch = False):
        """"""

        for arg in [cost, salvage, life, start_period, end_period, factor, no_switch]:
            if isinstance(arg, ExcelError) or arg in KoalaBaseFunction.ErrorCodes:
                return arg

        for arg in [cost, salvage, life, start_period, end_period, factor]:
            if not isinstance(arg, (float, int)):
                raise ExcelError('#VALUE', 'Arg %s should be an int, float or long, instead: %s' % (arg, type(arg)))

        start_period = start_period
        end_period = end_period
        sln_depr = SLN.sln(cost, salvage, life)
        depr_rate = factor / life
        acc_depr = 0
        depr = 0
        switch_to_sln = False
        sln_depr = 0
        result = 0
        start_life = 0
        delta_life = life % 1
        if delta_life > 0: # to handle cases when life is not an integer
            end_life = int(life + 1)

        else:
            end_life = int(life)
        periods = list(range(start_life, end_life))

        if int(start_period) != start_period:
            delta_start = abs(int(start_period) - start_period)
            depr = (cost - acc_depr) * depr_rate * delta_start
            acc_depr += depr
            start_life = 1
            periods = [x + 0.5 for x in periods]

        for index, current_year in enumerate(periods):

            if not no_switch: # no_switch = False (Default Case)
                if switch_to_sln:
                    depr = sln_depr

                else:
                    depr = (cost - acc_depr) * depr_rate
                    acc_depr += depr

                    temp_sln_depr = SLN.sln(cost, salvage, life)

                    if depr < temp_sln_depr:
                        switch_to_sln = True
                        fixed_remaining_years = life - current_year - 1
                        fixed_remaining_cost = cost - acc_depr

                         # we need to check future sln: current depr should never be smaller than sln to come
                        sln_depr = SLN.sln(fixed_remaining_cost, salvage, fixed_remaining_years)

                        if sln_depr > depr: # if it's the case, we switch to sln earlier than the regular case
                            # cancel what has been done
                            acc_depr -= depr
                            fixed_remaining_years += 1
                            fixed_remaining_cost = cost - acc_depr

                            # recalculate depreciation
                            sln_depr = SLN.sln(fixed_remaining_cost, salvage, fixed_remaining_years)
                            depr = sln_depr
                            acc_depr += depr

            else: # no_switch = True
                depr = (cost - acc_depr) * depr_rate
                acc_depr += depr

            delta_start = abs(current_year - start_period)

            if delta_start < 1 and delta_start != 0:
                result += depr * (1 - delta_start)

            elif current_year >= start_period and current_year < end_period:

                delta_end = abs(end_period - current_year)

                if delta_end < 1 and delta_end != 0:
                    result += depr * delta_end

                else:
                    result += depr

        return result
