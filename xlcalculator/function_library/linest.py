
# Excel reference: https://support.office.com/en-us/article/LINEST-function-84d7d0d9-6e50-4101-977a-fa7abf772b6d

import numpy as np

from .excel_lib import XlCalculatorBaseFunction

class Linest(XlCalculatorBaseFunction):
    """"""

    def linest(self, *args, **kwargs):
        """"""

        Y = list(args[0].values())
        X = list(args[1].values())

        if len(args) == 3:
            const = args[2]
            if isinstance(const, str):
                const = (const.lower() == "true")

        else:
            const = True

        degree = kwargs.get('degree', 1)

        # build the vandermonde matrix
        A = np.vander(X, degree+1)

        if not const:
            # force the intercept to zero
            A[:,-1] = np.zeros((1, len(X)))

        # perform the fit
        (coefs, residuals, rank, sing_vals) = np.linalg.lstsq(A, Y)

        return coefs
