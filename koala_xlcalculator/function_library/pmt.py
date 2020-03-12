
# Excel reference: https://support.office.com/en-us/article/PMT-function-0214da64-9a63-4996-bc20-214433fa6441


import numpy as np

from .excel_lib import KoalaBaseFunction

class PMT(KoalaBaseFunction):
    """"""

    def pmt(self, *args):
        """"""

        rate = args[0]
        num_payments = args[1]
        present_value = args[2]
        # WARNING fv & type not used yet - both are assumed to be their defaults (0)
        # fv = args[3]
        # type = args[4]

        return -present_value * rate / (1 - np.power(1 + rate, -num_payments))
