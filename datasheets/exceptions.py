class DatasheetsException(Exception):
    """Base Exception for all other datasheets exceptions

    This is intended to make catching exceptions from this library easier.
    """


class FolderNotFound(DatasheetsException):
    """ Attempting to open non-existent or inaccessible folder """


class MultipleWorkbooksFound(DatasheetsException):
    """ Multiple workbooks found for the given filename """


class PermissionNotFound(DatasheetsException):
    """ Trying to retrieve non-existent permission for workbook """


class TabNotFound(DatasheetsException):
    """ Trying to open non-existent tab """


class WorkbookNotFound(DatasheetsException):
    """ Trying to open non-existent or inaccessible workbook """
