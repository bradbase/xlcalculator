import datetime

EXCEL_EPOCH = datetime.datetime(1900, 1, 1)


def number_to_datetime(value):
    offset = 2 if value > 58 else 1
    delta = datetime.timedelta(
        days=int(value) - offset, seconds=(value % 1) * 24 * 60 * 60)
    return EXCEL_EPOCH + delta


def datetime_to_number(value):
    delta = value - EXCEL_EPOCH
    # Excel treats 1900 as a leap year.
    offset = 2 if delta.days > 58 else 1
    return (delta.days + offset) + (delta.seconds / 24 * 60 * 60)
