
import logging

class ExcelError(Exception):
    """"""

    def __init__(self, value, info=None):

        self.value = value
        self.info = info
        logging.error("ExcelError, {}, {}".format(value, info))


    def __str__(self):
        
        return self.value
