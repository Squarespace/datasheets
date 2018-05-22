from collections import OrderedDict
import types

import apiclient
import pandas as pd

from datasheets import exceptions, helpers


class Tab(object):
    def __init__(self, tabname, workbook, drive_svc, sheets_svc):
        """Create a datasheets.Tab instance of an existing Google Sheets tab.

        This class in not intended to be directly instantiated; it is created by
        datasheets.Workbook.fetch_tab().

        Args:
            tabname (str): The name of the tab
            workbook (datasheets.Workbook): The workbook instance that instantiated this tab
            drive_svc (googleapiclient.discovery.Resource): An instance of Google Drive
            sheets_svc (googleapiclient.discovery.Resource): An instance of Google Sheets
        """
        self.tabname = tabname
        self._workbook = workbook
        self.drive_svc = drive_svc
        self.sheets_svc = sheets_svc

        # Get basic properties of the tab. We do this here partly
        # to force failures early if tab can't be found
        try:
            self._update_tab_properties()
        except apiclient.errors.HttpError as e:
            if 'Unable to parse range'.encode() in e.content:
                raise exceptions.TabNotFound('The given tab could not be found. Error generated: {}'.format(e))
            else:
                raise

        self.url = 'https://docs.google.com/spreadsheets/d/{}#gid={}'.format(self.workbook.file_id, self.tab_id)

    def __getattribute__(self, attr):
        """Get an attribute (variable or method) of this instance of this class

        For client OAuth, before each user-facing method call this method will verify that the
        access token is not expired and refresh it if it is.

        We only refresh on user-facing method calls since otherwise we'd be refreshing multiple
        times per user action (once for the user call, possibly multiple times for the private
        method calls invoked by it).
        """
        requested_attr = super(Tab, self).__getattribute__(attr)

        if isinstance(requested_attr, types.MethodType) \
           and not attr.startswith('_'):
            self.workbook.client._refresh_token_if_needed()

        return requested_attr

    def __repr__(self):
        msg = "<{module}.{name}(filename='{filename}', tabname='{tabname}')>"
        return msg.format(module=self.__class__.__module__,
                          name=self.__class__.__name__,
                          filename=self.workbook.filename,
                          tabname=self.tabname)

    @staticmethod
    def _process_rows(raw_data):
        """Prepare a tab's raw data so that a pandas.DataFrame can be produced from it

        Args:
            raw_data (dict): The raw data from a tab

        Returns:
            list: A list of lists representing the raw_data, with one list per row in the tab
        """
        raw_rows = raw_data['sheets'][0]['data'][0].get('rowData', {})
        rows = []
        for i, row in enumerate(raw_rows):
            row_values = []
            for cell in row.get('values', {}):
                # If the cell is empty, use None
                value = cell.get('effectiveValue', {None: None})
                # value is a dict with only 1 key so this next(iter()) is safe
                base_fmt, cell_value = next(iter(value.items()))

                num_fmt = cell.get('effectiveFormat', {}).get('numberFormat')
                if num_fmt:
                    cell_format = num_fmt['type']
                else:
                    cell_format = base_fmt

                formatting_fn = helpers._TYPE_CONVERSIONS[cell_format]
                if cell_value:
                    try:
                        cell_value = formatting_fn(cell_value)
                    except ValueError:
                        pass
                    except TypeError:
                        raise TypeError(
                            "Mismatch exists in expected and actual data types for cell with "
                            "value '{value}'. Cell format is '{cell_format}' but cell value type "
                            "is '{value_type}'. To correct this, in Google Sheets set the "
                            "appropriate cell format or set it to Automatic".format(
                                value=cell_value,
                                cell_format=cell_format,
                                value_type=type(cell_value))
                        )
                row_values.append(cell_value)

            rows.append(row_values)
        return rows

    @property
    def ncols(self):
        """ Property for the number (int) of columns in the tab """
        return self.properties['gridProperties']['columnCount']

    @property
    def nrows(self):
        """ Property for the number (int) of rows in the tab """
        return self.properties['gridProperties']['rowCount']

    @property
    def tab_id(self):
        """ Property that gives the ID for the tab """
        return self.properties['sheetId']

    @property
    def workbook(self):
        """ Property for the workbook instance that this tab belongs to """
        return self._workbook

    def _add_rows_or_columns(self, kind, n):
        request_body = {'appendDimension': {
                            'sheetId': self.tab_id,
                            'dimension': kind,
                            'length': n
                            }
                        }
        body = {'requests': [request_body]}
        self.workbook.batch_update(body)
        self._update_tab_properties()

    def _update_tab_properties(self):
        raw_properties = self.sheets_svc.get(spreadsheetId=self.workbook.file_id,
                                             ranges=self.tabname + '!A1',
                                             fields='sheets/properties').execute()
        self.properties = raw_properties['sheets'][0]['properties']

    def add_rows(self, n):
        """Add n rows to the given tab

        Args:
            n (int): The number of rows to add

        Returns:
            None
        """
        self._add_rows_or_columns(kind='ROWS', n=n)

    def add_columns(self, n):
        """Add n columns to the given tab

        Args:
            n (int): The number of columns to add

        Returns:
            None
        """
        self._add_rows_or_columns(kind='COLUMNS', n=n)

    def align_cells(self, horizontal='LEFT', vertical='MIDDLE'):
        """Align all cells in the tab

        Args:
            horizontal (str): The horizontal alignment for cells. May be one of 'LEFT',
                'CENTER', or 'RIGHT'

            vertical (str): The vertical alignment for cells. May be one of 'TOP',
                'MIDDLE', 'BOTTOM'

        Returns:
            None
        """
        request_body = {'repeatCell': {
                             'range': {
                                  'sheetId': self.tab_id,
                                  'startRowIndex': 0,
                                  'endRowIndex': self.nrows
                              },
                             'cell': {
                                  'userEnteredFormat': {
                                       'horizontalAlignment': horizontal,
                                       'verticalAlignment': vertical,
                                        }
                                   },
                             'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment)'
                              }
                        }
        body = {'requests': [request_body]}
        self.workbook.batch_update(body)

    def alter_dimensions(self, nrows=None, ncols=None):
        """Alter the dimensions of the current tab.

        If either dimension is left to None, that dimension will not be altered. Note that it is
        possible to set nrows or ncols to smaller than the current tab dimensions, in which case
        that data will be eliminated.

        Args:
            nrows (int): The number of rows for the tab to have
            ncols (int): The number of columns for the tab to have

        Returns:
            None
        """
        request_body = {'updateSheetProperties': {
                             'properties': {
                                 'sheetId': self.tab_id,
                                 'gridProperties': {
                                     'columnCount': ncols or self.ncols,
                                     'rowCount': nrows or self.nrows
                                     }
                                 },
                             'fields': 'gridProperties(columnCount, rowCount)'
                             }
                        }
        body = {'requests': [request_body]}
        self.workbook.batch_update(body)
        self._update_tab_properties()

    def append_data(self, data, index=True, autoformat=True):
        """Append data to the existing data in this tab.

        If the new data exceeds the tab's current dimensions the tab will be resized to
        accommodate it. Data headers will not be included among the appended data as they are
        assumed to already be among the existing tab data.

        If the dimensions of `data` are larger than the tab's current dimensions,
        the tab will automatically be resized to fit it.

        Args:
            data (pandas.DataFrame or dict or list): The data to be uploaded, formatted as a
                pandas.DataFrame, a dict of lists, or a list of lists
            index (bool): If `data` is a pandas.DataFrame, whether to upload the index as well

        Returns:
            None
        """
        # Convert everything to lists of lists, which Google Sheets requires
        headers, values = helpers._make_list_of_lists(data, index)
        values = helpers._convert_nan_and_datelike_values(values)

        body = {'values': values}
        self.sheets_svc.values().append(spreadsheetId=self.workbook.file_id, range=self.tabname,
                                        valueInputOption='USER_ENTERED', body=body).execute()

        if autoformat:
            self.autoformat(len(headers))

        self._update_tab_properties()

    def autoformat(self, n_header_rows):
        """Apply default stylings to the tab

        This will apply the following stylings to the tab:

            - Header rows will be formatted to a dark gray background and off-white text
            - Font for all cells will be set to size 10 Proxima Nova
            - Cells will be horizontally left-aligned and vertically middle-aligned
            - Columns will be resized to display their largest entry
            - Empty columns and rows will be trimmed from the tab

        Args:
            n_header_rows (int): The number of header rows (i.e. row of labels / metadata)

        Returns:
            None
        """
        self.format_headers(nrows=n_header_rows)
        self.format_font()
        self.align_cells()
        self.autosize_columns()

        populated_cells = self.sheets_svc.values().get(spreadsheetId=self.workbook.file_id,
                                                       range=self.tabname).execute()
        nrows = len(populated_cells['values'])
        ncols = max(map(len, populated_cells['values']))
        self.alter_dimensions(nrows=nrows, ncols=ncols)
        self._update_tab_properties()

    def autosize_columns(self):
        """Resize the widths of all columns in the tab to fit their data

        Returns:
            None
        """
        request_body = {'autoResizeDimensions': {
                            'dimensions': {
                                  'sheetId': self.tab_id,
                                  'dimension': 'COLUMNS',
                                  'startIndex': 0,
                                  'endIndex': self.ncols
                                  }
                            }
                        }
        body = {'requests': [request_body]}
        self.workbook.batch_update(body)

    def clear_data(self):
        """Clear all data from the tab while leaving formatting intact

        Returns:
            None
        """
        self.sheets_svc.values().clear(spreadsheetId=self.workbook.file_id,
                                       range=self.tabname,
                                       body={}).execute()

    def format_font(self, font='Proxima Nova', size=10):
        """Set the font and size for all cells in the tab

        Args:
            font (str): The name of the font to use
            size (int): The size to set the font to

        Returns:
            None
        """
        request_body = {'repeatCell': {
                            'range': {'sheetId': self.tab_id},
                            'cell': {
                                'userEnteredFormat': {
                                    'textFormat': {
                                        'fontSize': size,
                                        'fontFamily': font
                                        }
                                    }
                                },
                            'fields': 'userEnteredFormat(textFormat(fontSize,fontFamily))'
                            }
                        }
        body = {'requests': [request_body]}
        self.workbook.batch_update(body)

    def format_headers(self, nrows):
        """Format the first n rows of a tab.

        The following stylings will be applied to these rows:

            - Background will be set to dark gray with off-white text
            - Font will be set to size 10 Proxima Nova
            - Text will be horizontally left-aligned and vertically middle-aligned
            - Rows will be made "frozen" so that when the user scrolls these rows stay visible

        Args:
            nrows (int): The number of rows of headers in the tab

        Returns:
            None
        """
        body = {
          'requests': [
            {
              'repeatCell': {
                'range': {
                  'sheetId': self.tab_id,
                  'startRowIndex': 0,
                  'endRowIndex': nrows
                },
                'cell': {
                  'userEnteredFormat': {
                    'backgroundColor': {
                      'red': 0.26274511,
                      'green': 0.26274511,
                      'blue': 0.26274511
                    },
                    'horizontalAlignment': 'LEFT',
                    'textFormat': {
                      'foregroundColor': {
                        'red': 0.95294118,
                        'green': 0.95294118,
                        'blue': 0.95294118
                      },
                      'fontSize': 10,
                      'fontFamily': 'Proxima Nova',
                      'bold': False
                    }
                  }
                },
                'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
              }
            },
            {
              'updateSheetProperties': {
                'properties': {
                  'sheetId': self.tab_id,
                  'gridProperties': {
                    'frozenRowCount': nrows
                  }
                },
                'fields': 'gridProperties(frozenRowCount)'
              }
            }
          ]
        }
        self.workbook.batch_update(body)

    def fetch_data(self, headers=True, fmt='df'):
        """Retrieve the data within this tab.

        Efforts are taken to ensure that returned rows are always the same length. If
        headers=True, this length will be equal to the length of the headers. If headers=False,
        this length will be equal to the longest row.

        In either case, shorter rows will be padded with Nones and longer rows will be
        truncated (i.e. if there are 3 headers then all rows will have 3 entries regardless
        of the amount of populated cells they have).

        Args:
            headers (bool): If True, the first row will be used as the column names for the
                pandas.DataFrame. Otherwise, a 0-indexed range will be used instead

            fmt (str): The format in which to return the data. Accepted values: 'df', 'dict', 'list'

        Returns:
            When fmt='df' --> pandas.DataFrame

            When fmt='dict' --> list of dicts, e.g.::

                [{header1: row1cell1, header2: row1cell2},
                 {header1: row2cell1, header2: row2cell2},
                 ...]

            When fmt='list' --> tuple of header names, list of lists with row data, e.g.::

                ([header1, header2, ...],
                 [[row1cell1, row1cell2, ...], [row2cell1, row2cell2, ...], ...])
        """
        if fmt not in ('df', 'dict', 'list'):
            raise ValueError("Unexpected value '{}' for parameter `fmt`. "
                             "Accepted values are 'df', 'dict', and 'list'".format(fmt))

        fields = 'sheets/data/rowData/values(effectiveValue,effectiveFormat/numberFormat/type)'
        raw_data = self.sheets_svc.get(spreadsheetId=self.workbook.file_id, ranges=self.tabname,
                                       includeGridData=True, fields=fields).execute()
        processed_rows = self._process_rows(raw_data)

        # filter out empty rows
        max_idx = helpers._find_max_nonempty_row(processed_rows)

        if max_idx is None:
            if fmt == 'df':
                return pd.DataFrame([])
            elif fmt == 'dict':
                return []
            else:
                return ([], [])

        processed_rows = processed_rows[:max_idx+1]

        # remove trailing Nones on rows
        processed_rows = list(map(helpers._remove_trailing_nones, processed_rows))

        if headers:
            header_names = processed_rows.pop(0)
            max_width = len(header_names)
        else:
            # Iterate through rows to find widest one
            max_width = max(map(len, processed_rows))
            header_names = list(range(max_width))

        # resize the rows to match the number of column headers
        processed_rows = [helpers._resize_row(row, max_width) for row in processed_rows]

        if fmt == 'df':
            df = pd.DataFrame(data=processed_rows, columns=header_names)
            return df
        elif fmt == 'dict':
            make_row_dict = lambda row: OrderedDict(zip(header_names, row))
            return list(map(make_row_dict, processed_rows))
        else:
            return header_names, processed_rows

    def insert_data(self, data, index=True, autoformat=True):
        """Overwrite all data in this tab with the provided data.

        All existing data in the tab will be removed, even if it might not have been overwritten
        (for example, if there is 4x2 data already in the tab and only 2x2 data is being inserted).

        If the dimensions of `data` are larger than the tab's current dimensions,
        the tab will automatically be resized to fit it.

        Args:
            data (pandas.DataFrame or dict or list): The data to be uploaded, formatted as a
                pandas.DataFrame, a dict of lists, or a list of lists
            index (bool): If `data` is a pandas.DataFrame, whether to upload the index as well

        Returns:
            None
        """
        # Convert everything to lists of lists, which Google Sheets requires
        headers, values = helpers._make_list_of_lists(data, index)

        values = headers + values  # Include headers for inserts but not for appends
        self.clear_data()
        values = helpers._convert_nan_and_datelike_values(values)

        body = {'values': values}
        self.sheets_svc.values().update(spreadsheetId=self.workbook.file_id, range=self.tabname,
                                        valueInputOption='USER_ENTERED', body=body).execute()

        if autoformat:
            self.autoformat(len(headers))

        self._update_tab_properties()
