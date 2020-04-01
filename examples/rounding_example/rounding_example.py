
from decimal import Decimal
from decimal import ROUND_FLOOR
from math import e


def floating_point_rounded():

    print("e", e * 10**-17)
    precision_mask = "{0:." + str(17) + "f}1"
    precision = precision_mask.format(e * 10**-17)


    the_float_lhs = Decimal( float(0.5805555555555556) )
    the_dec_lhs = Decimal('0.5805555555555556')
    lhs_result = Decimal(the_float_lhs).quantize(Decimal(precision), rounding=ROUND_FLOOR)
    print("rounded_float", lhs_result)




    the_float_rhs = Decimal( float(0.5805555555555) )
    the_dec_rhs = Decimal('0.5805555555555')
    rhs_result = Decimal(the_float_rhs).quantize(Decimal(precision), rounding=ROUND_FLOOR)
    print("rounded_float", rhs_result )

    print("             ", precision)
    print("rounded_float",str(lhs_result))
    print("rounded_float", round(rhs_result, 13))
    print("LHS error ", the_float_lhs - the_dec_lhs)
    print("RHS error", the_float_rhs - the_dec_rhs)



if __name__ == '__main__':
    floating_point_rounded()
