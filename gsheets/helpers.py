'''
Functionality used within core.py. These functions and classes are not intended to be
utilized by the end-user and are not exposed in the external API.
'''
import collections
import contextlib
import datetime as dt
import sys

import oauth2client
import pandas as pd


class WorkbookNotFound(Exception):
    ''' Trying to open non-existent or inaccessible workbook '''


class FolderNotFound(Exception):
    ''' Trying to open non-existent or inaccessible folder '''


class MultipleWorkbooksFound(Exception):
    ''' Multiple workbooks found for the given filename '''


class PermissionNotFound(Exception):
    ''' Trying to retrieve non-existent permission for workbook '''


class TabNotFound(Exception):
    ''' Trying to open non-existent tab '''


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


def _convert_datelike_to_str(item):
    ''' This is needed to make Python datetime-like objects JSON serializable '''
    if isinstance(item, (dt.date, dt.datetime, dt.time)):
        return str(item)
    else:
        return item


def _find_max_nonempty_row(data):
    '''
    Take a list of lists and return the index of the last row in the table that is
    not a list of Nones. If empty rows are sandwiched between populated rows, these
    will be included.
    '''
    is_not_empty = lambda row: row.count(None) != len(row)
    nonempty_rows = list(map(is_not_empty, data))

    if not nonempty_rows.count(True):
        return None  # All rows were empty
    else:
        max_nonempty_index = max(loc for loc, val in enumerate(nonempty_rows) if val)
        return max_nonempty_index


def _make_list_of_lists(data, index):
    ''' Convert the input data to a list of lists, which Google Sheets requires for uploads  '''
    if isinstance(data, pd.DataFrame):
        headers = _process_df_headers(data, index)
        values = _process_df_values(data, index)
    elif isinstance(data, list) and isinstance(data[0], (dict, collections.OrderedDict)):
        headers = [data[0].keys()]
        values = [row.values() for row in data]
    elif isinstance(data, list) and isinstance(data[0], list):
        headers = []
        values = data
    else:
        raise ValueError('Input data must be a pd.DataFrame, a list of dicts, or a list of lists')

    return headers, values


def _process_df_index_names(data):
    '''
    Return the pd.DataFrame's index names as a list. If the index does
    not have names, assign them as index0, index1, etc.
    '''
    index_names = list(data.index.names)
    if index_names[0]:
        return index_names
    else:
        num_cols = len(index_names)
        return ['index' + str(i) for i in range(num_cols)]


def _process_df_headers(data, index):
    '''
    Create a list containing the row(s) of headers for a pd.DataFrame. If index=True, include
    index names in each row of headers. If `data` has a MultiIndex for columns, multiple rows
    of headers will be returned
    '''
    is_multiindex = isinstance(data.columns, pd.indexes.multi.MultiIndex)
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


def _process_df_values(data, index):
    ''' Include a pd.DataFrame's indexes in row values if specified '''
    if isinstance(data.index, pd.indexes.multi.MultiIndex) and index:
        mapping = lambda row: list(row[0] + row[1:])
        return list(map(mapping, data.itertuples(index=index)))
    else:
        return list(map(list, data.itertuples(index=index)))


def _remove_trailing_nones(array):
    ''' Trim any trailing Nones from a list '''
    while array and array[-1] is None:
        array.pop()
    return array


def _resize_row(array, new_len):
    '''
    Alter the size of a list to match a specified length. If it is too long,
    trim it. If it is too short, pad it with Nones
    '''
    current_len = len(array)
    if current_len > new_len:
        return array[:new_len]
    else:
        # pad the row with Nones
        padding = [None] * (new_len - current_len)
        return array + padding


@contextlib.contextmanager
def _suppress_stdout():
    ''' Suppress stdout to eliminate unneeded print statements from oauth2client's run_flow '''
    class DummyStdOut(object):
        def write(self, x):
            pass

    original_stdout = sys.stdout
    sys.stdout = DummyStdOut()
    yield
    sys.stdout = original_stdout


class _MockStorage(oauth2client.file.Storage):
    '''
    A mock of oauth2client.file.Storage to prevent credentials being stored to disk. This
    is intended to facilitate use of the library in multi-user environments where cached
    credentials would be undesirable
    '''
    def put(self, _):
        return None

    def locked_get(self):
        msg = ('This one-time access token is expired and can not be refreshed. '
               'Please create a new gsheets.Client instance')
        raise oauth2client.file.AccessTokenCredentialsError(msg)
