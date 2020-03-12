
# https://support.office.com/en-us/article/concat-function-9b1a9a3f-94ff-41af-9736-694cbd6b4ca2


from .excel_lib import KoalaBaseFunction

class Concat(KoalaBaseFunction):
    """"""

    def concat(self, *args):
        """"""

        return concatenate(*tuple( KoalaBaseFunction.flatten(args) ))
