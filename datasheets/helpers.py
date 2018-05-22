"""
Functionality used elsewhere. Almost all of these functions and classes are not
intended to be utilized by the end-user and are not exposed in the external API.
"""
import collections
import contextlib
import copy
import datetime as dt
import sys

import numpy as np
from oauth2client.file import Storage
from oauth2client.client import AccessTokenCredentialsError
import pandas as pd


# Note: dates, times, and datetimes in Google Sheets are represented in 'serial number' format
# as explained here: https://developers.google.com/sheets/reference/rest/v4/DateTimeRenderOption
_TYPE_CONVERSIONS = {'numberValue': float,
                     'stringValue': str,
                     'boolValue': bool,
                     'NUMBER_FORMAT_TYPE_UNSPECIFIED': lambda x: x,
                     'TEXT': str,
                     'NUMBER': float,
                     'PERCENT': float,
                     'CURRENCY': float,
                     'DATE': lambda x: dt.date(1899, 12, 30) + dt.timedelta(days=x),
                     'TIME': lambda x: (dt.datetime.min + dt.timedelta(days=x)).time(),
                     'DATE_TIME': lambda x: dt.datetime(1899, 12, 30) + dt.timedelta(days=x),
                     'SCIENTIFIC': float,
                     None: lambda x: x
                     }

ASCII_CHAR_OFFSET = ord('A') - 1
NUMBER_OF_LETTERS_IN_ALPHABET = 26


def _convert_nan_and_datelike_values(values):
    """Make all items JSON serializable

    Args:
        values (list)

    Returns:
        list: A copy of the list, with datelike-object converted to strings and np.nans
            converted to None
    """
    output = []
    for row in values:
        new_row = []
        for item in row:
            if isinstance(item, np.float) and np.isnan(item):
                new_row.append(None)
            elif isinstance(item, (dt.date, dt.datetime, dt.time)):
                new_row.append(str(item))
            else:
                new_row.append(item)
        output.append(new_row)

    return output


def _escape_query(query):
    return query.replace("\\", "\\\\").replace("'", r"\'")


def _find_max_nonempty_row(data):
    """Identify the index of largest row in a table that is not all Nones

    Empty rows sandwiched between populated rows will not mistakenly be identified as indicating
    the end of the populated data.

    Args:
        data (list): A list of lists, with each sublist representing a row in the table

    Returns:
        int: Index associated with the last non-empty row (i.e. the last list that is not all Nones)
    """
    is_not_empty = lambda row: row.count(None) != len(row)
    nonempty_rows = list(map(is_not_empty, data))

    if nonempty_rows.count(True):
        max_nonempty_index = max(loc for loc, val in enumerate(nonempty_rows) if val)
        return max_nonempty_index


def _get_column_letter(col_idx):
    """ Convert a column number into a label, e.g. 3 -> C, 27 -> AA, 53 -> BA, etc. """
    quotient, remainder = divmod(col_idx, NUMBER_OF_LETTERS_IN_ALPHABET)
    suffix = chr(remainder + ASCII_CHAR_OFFSET)
    if quotient == 0:
        return suffix

    return _get_column_letter(quotient) + suffix


def _make_list_of_lists(data, index):
    """Convert the input data to a list of lists, which Google Sheets requires for uploads.

    Note that the headers is a list of lists because we may have multiple rows of headers for
    DataFrames

    Args:
        data (pandas.DataFrame or list): The data set to be converted
        index (bool): Whether the index should be processed as well. This argument is only
            applicable if `data` is a pandas.DataFrame

    Returns:
        list: A list of lists representing the input data set
    """
    if isinstance(data, pd.DataFrame):
        headers = _process_df_headers(data, index)
        values = _process_df_values(data, index)
    elif isinstance(data, collections.Sequence) and isinstance(data[0], collections.Mapping):
        headers = [list(data[0].keys())]
        values = []
        for row in data:
            # We have to ensure we return the values in the same order, i.e. get them by key
            row_values = [row[key] for key in headers[0]]
            values.append(row_values)
    elif isinstance(data, list) and isinstance(data[0], list):
        headers = []
        values = data
    else:
        raise ValueError('Input data must be a pandas.DataFrame, a list of dicts, or a list of lists')

    return headers, values


def _process_df_headers(data, index):
    """Create a list containing the row(s) of column headers for a pandas.DataFrame

    Args:
        data (pandas.DataFrame): The data set to process column headers from
        index (bool): If True, include index names in each row of headers. If `data` has a
            MultiIndex for columns, multiple rows of headers will be returned

    Returns:
        list: A list of lists representing the column headers for the input data set
    """
    is_multiindex = isinstance(data.columns, pd.MultiIndex)
    idx_names = _process_df_index_names(data)

    if index and is_multiindex:
        column_names = map(list, zip(*data.columns))
        return [idx_names + row for row in column_names]
    elif index and not is_multiindex:
        return [idx_names + data.columns.tolist()]
    elif (not index) and is_multiindex:
        return list(map(list, zip(*data.columns)))
    else:
        # not is_multiindex and not index
        return [data.columns.tolist()]


def _process_df_index_names(data):
    """Create a list containing the index names of a pandas.DataFrame

    If the index does not have names, assign them as index0, index1, etc.

    Args:
        data (pandas.DataFrame): The data set to process index names for

    Returns:
        list: A list of index names for the in put data set
    """
    index_names = list(data.index.names)
    if index_names[0]:
        return index_names
    else:
        num_cols = len(index_names)
        return ['index' + str(i) for i in range(num_cols)]


def _process_df_values(data, index):
    """Create a list containing the row(s) of a pandas.DataFrame

    Args:
        data (pandas.DataFrame): The data se to process values from
        index (bool): Whether the index should be processed as well

    Returns:
        list: A list of lists, with each sublist representing one row in the input data set
    """
    if index and isinstance(data.index, pd.MultiIndex):
        mapping = lambda row: list(row[0] + row[1:])
        return list(map(mapping, data.itertuples(index=index)))
    else:
        return list(map(list, data.itertuples(index=index)))


def _remove_trailing_nones(array):
    """ Trim any trailing Nones from a list """
    while array and array[-1] is None:
        array.pop()
    return array


def _resize_row(array, new_len):
    """Alter the size of a list to match a specified length

    If the list is too long, trim it. If it is too short, pad it with Nones

    Args:
        array (list): The data set to pad or trim
        new_len int): The desired length for the data set

    Returns:
        list: A copy of the input `array` that has been extended or trimmed
    """
    current_len = len(array)
    if current_len > new_len:
        return array[:new_len]
    else:
        # pad the row with Nones
        padding = [None] * (new_len - current_len)
        return array + padding


@contextlib.contextmanager
def _remove_sys_argv():
    """Temporarily removing sys.argv

    The authentication flow within Jupyter Notebooks errors because argparse tries to parse
    sys.argv. Based on the following two github issues it seems like one way around this is
    to just remove all arguments to sys.argv. However, since it's not clear whether permanently
    removing these arguments might introduce unwanted behavior, we only temporarily remove them
    while running our auth flow and then restore their value.

    Sources:
        * https://github.com/google/oauth2client/issues/695
        * https://github.com/spyder-ide/spyder/issues/3883
    """
    original_argv = copy.deepcopy(sys.argv)
    sys.argv = []
    yield
    sys.argv = original_argv


@contextlib.contextmanager
def _suppress_stdout():
    """ Suppress stdout to eliminate unneeded print statements from oauth2client's run_flow """
    class DummyStdOut(object):
        def write(self, x):
            pass

        def flush(self):
            pass

    original_stdout = sys.stdout
    sys.stdout = DummyStdOut()
    yield
    sys.stdout = original_stdout


def convert_cell_index_to_label(row, col):
    """Convert two cell indexes to a string address

    Args:
        row (int): The cell row number, starting from 1
        col (int): The cell column number, starting from 1

    Note that Google Sheets starts both the row and col indexes at 1.

    Example:
        >>> sheets.convert_cell_index_to_label(1, 1)
        A1
        >>> sheets.convert_cell_index_to_label(10, 40)
        BH10

    Returns:
        str: The cell reference as an address (e.g. 'B6')
    """
    row = int(row)
    col = int(col)

    if row < 1 or col < 1:
        raise ValueError('Row and column values must be >= 1')

    column_label = _get_column_letter(col)
    label = '{}{}'.format(column_label, row)
    return label


def convert_cell_label_to_index(label):
    """Convert a cell label in string form into one based cell indexes of the form (row, col).

    Args:
        label (str): The cell label in string form

    Note that Google Sheets starts both the row and col indexes at 1.

    Example:
        >>> sheets.convert_cell_label_to_index('A1')
        (1, 1)
        >>> sheets.convert_cell_label_to_index('BH10')
        (10, 40)

    Returns:
        tuple: The cell reference in (row_int, col_int) form
    """
    if not isinstance(label, str):
        raise ValueError('Input must be a string')

    import re

    # Split out the letters from the numbers
    pattern = r'([A-Za-z]+)([1-9]\d*)'
    match = re.match(pattern, label)

    if not match:
        raise ValueError('Unable to parse user-provided label')

    column_label = match.group(1).upper()
    row = int(match.group(2))

    col = 0
    for c in column_label:
        col = col * NUMBER_OF_LETTERS_IN_ALPHABET + (ord(c) - ASCII_CHAR_OFFSET)

    return (row, col)


class _MockStorage(Storage):
    """A mock of oauth2client.file.Storage to prevent credentials being stored to disk.

    This is intended to facilitate use of the library in multi-user environments where cached
    credentials would be undesirable
    """
    def __init__(self):
        super(_MockStorage, self).__init__(None)

    def put(self, _):
        return

    def locked_get(self):
        raise AccessTokenCredentialsError('This one-time access token is expired and can not be '
                                          'refreshed. Please create a new datasheets.Client instance.')
