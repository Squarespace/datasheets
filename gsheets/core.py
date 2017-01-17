'''
Core functionality for the package
'''
import argparse
from collections import OrderedDict
import httplib2
import json
import os

import apiclient
import oauth2client
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

from gsheets import helpers


def convert_cell_index_to_label(row, col):
    '''
    Convert a tuple of cell indexes to a string address. For example:

    >>> sheets.convert_cell_index_to_label(1, 1)
    A1
    >>> sheets.convert_cell_index_to_label(10, 40)
    BH10

    Note that Google Sheets starts both the row and col indexes at 1.

    This function is based off of opeypyxl.util._get_column_letter
    and gspread.Worksheet.get_addr_int
    '''
    MAGIC_NUMBER = 64

    row = int(row)
    col = int(col)

    if row < 1 or col < 1:
        raise ValueError('row and column values must be >= 1')

    col_idx = col
    column_label = ''

    while col_idx:
        (col_idx, remainder) = divmod(col_idx, 26)
        if remainder == 0:
            remainder = 26
            col_idx -= 1
        prefix_letter = chr(remainder + MAGIC_NUMBER)
        column_label = prefix_letter + column_label

    label = '{}{}'.format(column_label, row)
    return label


def convert_cell_label_to_index(label):
    '''
    Convert a cell label in string form into one-based cell indexes of the form (row, col).
    For example:

    >>> sheets.convert_cell_label_to_index('A1')
    (1, 1)
    >>> sheets.convert_cell_label_to_index('BH10')
    (10, 40)

    This function is based off of gspread.Worksheet.get_int_addr
    '''
    if not isinstance(label, str):
        raise ValueError('Input must be a string')

    import re

    # Split out the letters from the numbers
    pattern = r'([A-Za-z]+)([1-9]\d*)'
    match = re.match(pattern, label)

    if not match:
        raise ValueError('Unable to parse user-provided label')

    column_label, row = match.groups()
    column_label, row = column_label.upper(), int(row)

    col = 0
    for i, c in enumerate(reversed(column_label)):
        MAGIC_NUMBER = 64
        col += (ord(c) - MAGIC_NUMBER) * (26 ** i)

    return (row, col)


class Client(object):
    def __init__(self, service=False, storage=True, testing_config={}):
        '''
        Create an authenticated client for interacting with Google Drive and Google Sheets.

        Inputs:

            service

                Whether to authenticate as a user or as a service account. If service=False,
                you will be prompted to authorize this instance to access the Google Drive
                attached to one of your Gmail accounts. Service-based authorization proceeds
                from ~/.gsheets/service_key.json. User-based authorization proceeds from
                the client secrets stored at ~/.gsheets/client_secrets.json. By default this
                authorization is only tied to this instance of the class and access disappears
                when the instance is deleted. However, authorization can be stored using
                storage=True, which persists the credentials to ~/.gsheets/client_credentials.json

            storage:

                Whether to use authorized credentials stored in ~/.gsheets/client_credentials.json
                (or if credentials have not been stored yet or are invalid, obtain new credentials
                and store them at that location).

                If False, authorization will be requested every time a new Client instance is
                created. This mode deliberately does not store credentials to disk in order
                to allow for use of the library in multi-user environments.

            testing_config

                An optional dict of the form:

                    {'drive_http': googleapiclient.http.HttpMock,
                     'sheets_http': googleapiclient.http.HttpMock,
                     'request_builder': googleapiclient.http.RequestMockBuilder
                     }

                This is primarily used as a testing vehicle. See the documentation for
                apiclient.http.RequestMockBuilder for more details:

                    https://github.com/google/google-api-python-client/blob/master/googleapiclient/http.py#L1486
        '''
        drive_http = testing_config.get('drive_http')
        sheets_http = testing_config.get('sheets_http')
        request_builder = testing_config.get('request_builder', apiclient.http.HttpRequest)

        if not (drive_http and sheets_http):
            if service:
                credentials = self._get_service_credentials()
            else:
                credentials = self._get_client_credentials(storage=storage)
            http = credentials.authorize(httplib2.Http())
            drive_http = drive_http or http
            sheets_http = sheets_http or http
        else:
            self.email = None

        self.drive_svc = apiclient.discovery.build('drive', 'v3', http=drive_http,
                                                   requestBuilder=request_builder)
        # bind sheets_svc directly to .spreadsheets() as the API exposes no other functionality
        self.sheets_svc = apiclient.discovery.build('sheets', 'v4', http=sheets_http,
                                                    requestBuilder=request_builder).spreadsheets()

    def __repr__(self):
        msg = "<{module}.{name}(email='{email}') at {mem_addr}>"
        return msg.format(module=self.__class__.__module__,
                          name=self.__class__.__name__,
                          email=self.email,
                          mem_addr=hex(id(self)))

    def _get_service_credentials(self):
        '''
        Return an instance of oauth2client.service_account.ServiceAccountCredentials; this instance
        has not yet been authenticated / authorized
        '''
        filepath = os.path.expanduser('~/.gsheets/service_key.json')
        with open(filepath) as f:
            keyfile_dict = json.load(f)

        self.email = keyfile_dict['client_email']  # used in __repr__
        return ServiceAccountCredentials.from_json_keyfile_dict(keyfile_dict=keyfile_dict,
                                                                scopes=('https://www.googleapis.com/auth/drive',))

    def _get_client_credentials(self, storage):
        """
        Gets valid user credentials. If storage=True, then first look in ~/.gsheets/client_credentials.json
        to see if authorized credentials have been stored. If nothing has been stored, or if
        the stored credentials are invalid, the OAuth2 flow is completed to obtain
        the new credentials.

        If storage=False, execute the OAuth2 flow to obtain new user credentials. This mode
        deliberately does not store credentials to disk in order to allow for use of the library
        in multi-user environments.
        """
        if storage:
            credential_path = os.path.expanduser('~/.gsheets/client_credentials.json')
            store = oauth2client.file.Storage(credential_path)
            credentials = store.get()
        else:
            store = helpers._MockStorage(None)
            credentials = None

        if not credentials or credentials.invalid:
            client_secrets_path = os.path.expanduser('~/.gsheets/client_secrets.json')
            scopes = ('https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/userinfo.email')
            flow = oauth2client.client.flow_from_clientsecrets(client_secrets_path,
                                                               scope=scopes)
            flow.user_agent = 'Python gsheets library'
            # use 'offline' to prevent InvalidGrant on refreshes after 1-hour access expiration
            flow.params['access_type'] = 'offline'
            # User the default flags provided by oauth2client.tools.argparser
            flags = argparse.Namespace(auth_host_name='localhost',
                                       auth_host_port=[8080],
                                       logging_level='ERROR',
                                       noauth_local_webserver=False)
            # Store the credentials at ~/.gsheets/client_credentials.json; if storage=False
            # then this won't happen because we setattr store.put to do nothing
            with helpers._suppress_stdout():
                credentials = oauth2client.tools.run_flow(flow, store, flags)

        self.email = credentials.id_token['email']
        return credentials

    def _get_file_id(self, filename, kind):
        '''
        Return the file_id for the Google Drive file with the specified filename (i.e. title).

        If multiple files exist with the specified filename, an exception is raised with
        information about each file, including its filename, file_id, the datetime it was last
        modified, and a webViewLink URL to look at the file. The user will be directed to use
        the appropriate file_id in their function instead.
        '''
        matches = []
        for f in self._get_info_on_items(kind=kind):
            if f['name'] == filename:
                matches.append(f)

        if len(matches) == 1:
            return matches[0]['id']
        elif len(matches) == 0 and kind == 'spreadsheet':
            msg = 'Workbook not found. Verify that it is shared with {}'
            raise helpers.WorkbookNotFound(msg.format(self.email))
        elif len(matches) == 0 and kind == 'folder':
            msg = 'Folder not found. Verify that it is shared with {}'
            raise helpers.FolderNotFound(msg.format(self.email))
        else:
            # Format the matches for readability when printed with the exception
            template = '''\n\n{}\nfilename: {}\nfile_id: {}\nmodifiedTime: {}\nwebViewLink: {}'''
            formatted_output = ''
            for i, row in enumerate(matches):
                cleaned_time = row['modifiedTime'].replace('T', ' ').replace('Z', '')
                this_row = template.format(i, row['name'], row['id'], cleaned_time, row['webViewLink'])
                formatted_output += this_row

            msg = ('Multiple workbooks founds. Please choose the correct file_id below '
                   'and provide it to your function instead of a filename:')
            raise helpers.MultipleWorkbooksFound(msg + formatted_output)

    def _get_info_on_items(self, kind, folder=None, fields='files(name,id,modifiedTime,webViewLink)'):
        '''
        Return a list of dicts, each one representing one workbook or folder shared with the user

        Inputs:

            kind
                'spreadsheet' or 'folder'

            folder
                An optional folder to limit the results to

            fields
                The fields to return in the results

        '''
        fields = 'nextPageToken, {}'.format(fields)
        query = "mimeType='application/vnd.google-apps.{}'".format(kind)
        if folder:
            folder_id = self._get_file_id(folder, kind='folder')
            query += " and '{}' in parents".format(folder_id)

        page_token = None
        raw_info = []

        while True:
            response = self.drive_svc.files().list(fields=fields, q=query,
                                                   orderBy='viewedByMeTime desc',
                                                   pageSize=1000,
                                                   pageToken=page_token).execute()
            raw_info += response.get('files', [])
            page_token = response.get('nextPageToken')
            if page_token is None:
                break

        return raw_info

    def create_workbook(self, filename, folders=()):
        '''
        Create a blank workbook with the specific filename and return a gsheets.Workbook
        instance of it. Optionally provide a tuple of folders to add this file to.
        '''
        root_file_id = self.drive_svc.files().get(fileId='root', fields='id').execute()['id']
        folders = [self._get_file_id(filename=f, kind='folder') for f in folders]
        folders = [root_file_id] + list(folders)

        body = {
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'name': filename,
            'parents': folders
        }
        file_id = self.drive_svc.files().create(body=body).execute()['id']
        return Workbook(filename, file_id, self.drive_svc, self.sheets_svc)

    def delete_workbook(self, filename=None, file_id=None):
        '''
        Delete the workbook with the given filename (i.e. title) or file_id. If both are provided,
        the file_id will be used as it is more precise
        '''
        file_id = file_id or self._get_file_id(filename=filename, kind='spreadsheet')
        self.drive_svc.files().delete(fileId=file_id).execute()
        return None

    def get_workbook(self, filename=None, file_id=None):
        '''
        Return a gsheets.Workbook instance of the given filename (i.e. title) or file_id. If both
        are provided, the file_id will be used as it is more precise
        '''
        file_id = file_id or self._get_file_id(filename=filename, kind='spreadsheet')
        return Workbook(filename, file_id, self.drive_svc, self.sheets_svc)

    def show_workbooks(self, folder=None):
        '''
        Show all workbooks shared with this account, optionally limiting the results to
        only those workbooks in the specified folder
        '''
        raw_info = self._get_info_on_items(kind='spreadsheet', folder=folder)
        return pd.DataFrame(raw_info, columns=['name', 'id', 'modifiedTime', 'webViewLink'])

    def show_folders(self, only_mine=False):
        '''
        Show all folders shared with this account, optionally limiting the results to
        only those folders owned by this user
        '''
        fields = 'files(name,id,modifiedTime,webViewLink,owners(me))'
        raw_info = self._get_info_on_items(kind='folder', fields=fields)

        if only_mine:
            raw_info = [folder for folder in raw_info if folder['owners'][0]['me']]

        return pd.DataFrame(raw_info, columns=['name', 'id', 'modifiedTime', 'webViewLink'])


class Workbook(object):
    def __init__(self, filename, file_id, drive_svc, sheets_svc):
        '''
        Create a gsheets.Workbook instance of an existing Google Sheets doc that has
        the given file_id. This class in not intended to be directly instantiated; it is
        created by the get_workbook() method of gsheets.Client

        Inputs:
            filename
                The name of the workbook

            file_id
                The Google Sheets-assigned ID for the file

            drive_svc
                A googleapiclient.discovery.Resource instance for Google Drive

            sheets_svc
                A googleapiclient.discovery.Resource instance for Google Sheets
        '''
        # Note that work is not done here to verify that the Workbook exists because
        # the assumed mechanism of instantiation via sheets.Client already assures this
        # by searching for the file_id through the sheets.Client._get_file_id method.
        self.filename = filename
        self.file_id = file_id
        self.drive_svc = drive_svc
        self.sheets_svc = sheets_svc
        self.url = 'https://docs.google.com/spreadsheets/d/{}'.format(self.file_id)

    def __repr__(self):
        msg = "<{module}.{name}(filename='{filename}') at {mem_addr}>"
        return msg.format(module=self.__class__.__module__,
                          name=self.__class__.__name__,
                          filename=self.filename,
                          mem_addr=hex(id(self)))

    def _get_permission_id(self, email):
        ''' Return the permission_id associated with the given email '''
        permissions = self.drive_svc.permissions().list(fileId=self.file_id).execute()
        for perm in permissions['permissions']:
            this_entry = self.drive_svc.permissions().get(fileId=self.file_id,
                                                          permissionId=perm['id'],
                                                          fields='emailAddress').execute()
            if this_entry['emailAddress'] == email:
                return perm['id']

        msg = "Permission for email '{}' not found for workbook '{}'"
        raise helpers.PermissionNotFound(msg.format(email, self.filename))

    def add_permission(self, email, role='reader', notify=True, message=None):
        '''
        Add a permission to this workbook.

        Inputs:

            email
                The email address to grant the permission to

            role
                The type of permission to grant. Values can be one of:
                    - 'owner'
                    - 'writer'
                    - 'reader'

            notify
                Whether to notify the recipient who was granted the permission

            message
                If notify==True, the message to send with the email notification
        '''
        new_permission = {
            'emailAddress': email,
            'type': 'user',
            'role': role
        }
        self.drive_svc.permissions().create(fileId=self.file_id, body=new_permission,
                                            emailMessage=message, sendNotificationEmail=notify).execute()

    def batch_update(self, body):
        '''
        Apply one or more updates to a workbook or tab using the Google Sheets API's
        spreadsheets.batchUpdate method. This method permits effectively all functionality
        allowed for by the Google Sheets.

        Inputs:
            body
                A list of requests, with each request provided as a dict

        For more details on how to use this method, see the following:
            - Explanation of spreadsheets.batchUpdate method:
                https://developers.google.com/sheets/reference/rest/v4/spreadsheets/batchUpdate

            - List of available request types and the parameters they take:
                https://developers.google.com/sheets/reference/rest/v4/spreadsheets/request

        An example request that would perform only one update operation
        (the 'repeatCell' operation) is the following:

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
        '''
        self.sheets_svc.batchUpdate(spreadsheetId=self.file_id, body=body).execute()

    def create_tab(self, tabname, nrows=1000, ncols=26):
        ''' Create a new tab in the given workbook

        Inputs:

            tabname
                The name to display for the tab

            nrows
                An integer number of rows for the tab to have

            ncols
                An integer number of columns for the tab to have
        '''
        request_body = {'addSheet': {
                            'properties': {
                                  'title': tabname,
                                  'gridProperties': {
                                        'rowCount': nrows,
                                        'columnCount': ncols
                                        },
                                  }
                            }
                        }
        body = {'requests': [request_body]}
        self.sheets_svc.batchUpdate(spreadsheetId=self.file_id, body=body).execute()
        return self.get_tab(tabname)

    def delete_tab(self, tabname):
        ''' Delete a tab with the given name from the current workbook '''
        tab_id = Tab(tabname, self.filename, self.file_id, self.drive_svc, self.sheets_svc).tab_id
        request_body = {'deleteSheet': {'sheetId': tab_id}}
        body = {'requests': [request_body]}
        self.sheets_svc.batchUpdate(spreadsheetId=self.file_id, body=body).execute()

    def get_tab(self, tabname):
        '''
        Return a gsheets.Tab instance of the given tab associated with
        this workbook
        '''
        return Tab(tabname, self.filename, self.file_id, self.drive_svc, self.sheets_svc)

    def show_permissions(self):
        '''
        Return a pd.DataFrame of all email addresses shared on this workbook and
        the permission role they've been granted
        '''
        req = self.drive_svc.permissions().list(fileId=self.file_id, fields='permissions(id,role,type)')
        perm_ids = req.execute()['permissions']

        for perm in perm_ids:
            if perm['type'] == 'user':
                perm['email'] = self.drive_svc.permissions().get(
                    fileId=self.file_id, permissionId=perm['id'],
                    fields='emailAddress').execute().get('emailAddress')
            else:
                # type is 'group', 'domain', or 'anyone'
                perm['email'] = "User Type: '{}'".format(perm['type'])

        output = pd.DataFrame(data=perm_ids)
        return output[['email', 'role']]

    def show_all_tabs(self):
        ''' Show the names of the tabs within the given workbook, returned as a pd.DataFrame. '''
        workbook = self.sheets_svc.get(spreadsheetId=self.file_id,
                                       fields='sheets/properties/title').execute()
        tab_names = [tab['properties']['title'] for tab in workbook['sheets']]
        return pd.DataFrame(tab_names, columns=['Tabs'])

    def remove_permission(self, email):
        '''
        Remove a permission from this workbook.

        Inputs:

            email
                The email address whose permission will be removed
        '''
        permission_id = self._get_permission_id(email)
        self.drive_svc.permissions().delete(fileId=self.file_id, permissionId=permission_id).execute()


class Tab(object):
    def __init__(self, tabname, filename, file_id, drive_svc, sheets_svc):
        '''
        Create a gsheets.Tab instance of an existing Google Sheets tab with the given
        tabname in the file with the given file_id. This class in not intended to be directly
        instantiated; it is created by the get_tab() method of gsheets.Workbook
        '''
        self.tabname = tabname
        self.filename = filename
        self.file_id = file_id
        self.drive_svc = drive_svc
        self.sheets_svc = sheets_svc

        # Get basic properties of the tab. We do this here partly
        # to force failures early if tab can't be found
        try:
            self.properties = self._get_properties()
        except apiclient.errors.HttpError as e:
            if 'Unable to parse range' in e.content:
                raise helpers.TabNotFound('The given tab could not be found. Error generated: {}'.format(e))
            else:
                raise e

        self.url = 'https://docs.google.com/spreadsheets/d/{}#gid={}'.format(self.file_id, self.tab_id)

    def __repr__(self):
        msg = "<{module}.{name}(filename='{filename}', tabname='{tabname}') at {mem_addr}>"
        return msg.format(module=self.__class__.__module__,
                          name=self.__class__.__name__,
                          filename=self.filename,
                          tabname=self.tabname,
                          mem_addr=hex(id(self)))

    def _add_rows_or_columns(self, kind, n):
        request_body = {'appendDimension': {
                            'sheetId': self.tab_id,
                            'dimension': kind,
                            'length': n
                            }
                        }
        body = {'requests': [request_body]}
        self.sheets_svc.batchUpdate(spreadsheetId=self.file_id, body=body).execute()
        # Update tab properties
        self.properties = self._get_properties()

    def _get_properties(self):
        ''' Get properties of the given tab '''
        raw_properties = self.sheets_svc.get(spreadsheetId=self.file_id,
                                             ranges=self.tabname + '!A1',
                                             fields='sheets/properties').execute()
        return raw_properties['sheets'][0]['properties']

    @staticmethod
    def _process_rows(raw_data):
        ''' Prepare a tab's raw data so that a pd.DataFrame can be produced from it '''
        raw_rows = raw_data['sheets'][0]['data'][0].get('rowData', {})
        rows = []
        for i, row in enumerate(raw_rows):
            row_values = []
            for cell in row.get('values', {}):
                # If the cell is empty, use None
                value = cell.get('effectiveValue', {None: None})
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
                row_values.append(cell_value)

            rows.append(row_values)
        return rows

    def add_rows(self, n):
        ''' Add n rows to the given tab '''
        self._add_rows_or_columns(kind='ROWS', n=n)

    def add_columns(self, n):
        ''' Add n columns to the given tab '''
        self._add_rows_or_columns(kind='COLUMNS', n=n)

    def align_cells(self, horizontal='LEFT', vertical='MIDDLE'):
        '''
        Align all cells in the tab

        Inputs:

            horizontal
                The horizontal alignment for cells. May be one of 'LEFT', 'CENTER', or 'RIGHT'

            vertical
                The vertical alignment for cells. May be one of 'TOP', 'MIDDLE', 'BOTTOM'
        '''
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
        self.sheets_svc.batchUpdate(spreadsheetId=self.file_id, body=body).execute()

    def alter_dimensions(self, nrows=None, ncols=None):
        '''
        Alter the dimensions of the current tab. If either dimension is left to None, that
        dimension will not be altered. Note that it is possible to set nrows or ncols to smaller
        than the current tab dimensions, in which case that data will be eliminated

        nrows
            The number of rows for the tab to have

        ncols
            The number of columns for the tab to have
        '''
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
        self.sheets_svc.batchUpdate(spreadsheetId=self.file_id, body=body).execute()

        # Update tab properties
        self.properties = self._get_properties()

    def autosize_columns(self):
        ''' Resize the widths of all columns in the tab to fit their data '''
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
        self.sheets_svc.batchUpdate(spreadsheetId=self.file_id, body=body).execute()

    def clear_data(self):
        ''' Clear all data from the tab while leaving formatting intact '''
        _ = self.sheets_svc.values().clear(spreadsheetId=self.file_id,
                                           range=self.tabname,
                                           body={}).execute()

    def format_font(self, font='Proxima Nova', size=10):
        ''' Set the font and size for all cells in the tab '''
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
        self.sheets_svc.batchUpdate(spreadsheetId=self.file_id, body=body).execute()

    def format_headers(self, nrows):
        ''' Format the first n rows of a tab. '''
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
        self.sheets_svc.batchUpdate(spreadsheetId=self.file_id, body=body).execute()

    def get_data(self, headers=True, fmt='df'):
        '''
        Retrieve the data within this tab.

        Efforts are taken to ensure that returned rows are always the same length. If
        headers=True, this length will equal the number of headers. If headers=False, this
        length will be equal to the longest row.

        In either case, shorter rows will be padded with Nones and longer rows will be
        truncated (i.e. if there are 3 headers then all rows will have 3 entries regardless
        of the amount of populated cells they have).

        Inputs:

            headers
                If True, the first row will be used as the column names for the pd.DataFrame.
                Otherwise, a 0-indexed range will be used instead

            fmt
                The format in which to return the data. Accepted values: 'df', 'dict', 'list'

        Returns:

            if fmt='df' --> pd.DataFrame
            if fmt='dict' --> list of dicts, e.g.:
                [{header1: row1cell1, header2: row1cell2},
                 {header1: row2cell1, header2: row2cell2},
                 ...]
            if fmt='list' --> tuple of header names, list of lists with row data, e.g.:
                ([header1, header2, ...],
                 [[row1cell1, row1cell2, ...], [row2cell1, row2cell2, ...], ...])
        '''
        if fmt not in ('df', 'dict', 'list'):
            raise ValueError("Unexpected value '{}' for parameter `fmt`. "
                             "Accepted values are 'df', 'dict', and 'list'".format(fmt))

        fields = 'sheets/data/rowData/values(effectiveValue,effectiveFormat/numberFormat/type)'
        raw_data = self.sheets_svc.get(spreadsheetId=self.file_id, ranges=self.tabname,
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

    def get_parent_workbook(self):
        ''' Return the workbook that contains this tab '''
        return Workbook(self.filename, self.file_id, self.drive_svc, self.sheets_svc)

    @property
    def ncols(self):
        return self.properties['gridProperties']['columnCount']

    @property
    def nrows(self):
        return self.properties['gridProperties']['rowCount']

    @property
    def tab_id(self):
        return self.properties['sheetId']

    def upload_data(self, data, index=True, mode='insert', autoformat=True):
        '''
        Add data to to the given tab.

        If the dimensions of `data` are larger than the tab's current dimensions,
        the tab will automatically be resized to fit it

        Inputs:

            data
                The data to be uploaded, formatted as a pd.DataFrame, a dict of lists,
                or a list of lists

            index
                If `data` is a pd.DataFrame, whether to upload the index as well

            mode
                Either 'insert' or 'append'.

                If insert mode is chosen, all existing data in the tab will be removed,
                even if it might not have been overwritten (for example, if there is
                4x2 data already in the tab and only 2x2 data is being inserted).

                If append mode is chosen, the data will be appended to the existing
                data in the tab (and the tab will be resized to accommodate it).
                Data headers will not be included among the appended data as they
                are assumed to already be among the existing tab data

            autoformat
                If True, the following stylings will be applied to the tab:
                    - Header formatting will be applied (dark gray background, off-white text)
                    - Font for all cells will be set to size 10 Proxima Nova
                    - Cells will be horizontally left-aligned and vertically middle-aligned
                    - Columns will be resized to display their largest entry
                    - Empty columns and rows will be trimmed from the tab
        '''
        mode = mode.lower()
        if mode not in ('insert', 'append'):
            raise ValueError("Unexpected value '{}' for parameter `mode`. "
                             "Accepted values are 'insert' and 'append'".format(mode))

        # Convert everything to lists of lists, which Google Sheets requires
        headers, values = helpers._make_list_of_lists(data, index)

        if mode == 'insert':
            values = headers + values  # Include headers for inserts but not for appends
            self.clear_data()

        # Convert datelike items to strings
        values = list(map(lambda row: list(map(helpers._convert_datelike_to_str, row)), values))

        if mode == 'insert':
            fn = self.sheets_svc.values().update
        else:
            fn = self.sheets_svc.values().append

        body = {'values': values}
        fn(spreadsheetId=self.file_id, range=self.tabname,
           valueInputOption='USER_ENTERED', body=body).execute()

        if autoformat:
            self.format_headers(nrows=len(headers))
            self.format_font()
            self.align_cells()
            self.autosize_columns()

            populated_cells = self.sheets_svc.values().get(spreadsheetId=self.file_id,
                                                           range=self.tabname).execute()
            nrows = len(populated_cells['values'])
            ncols = max(map(len, populated_cells['values']))
            self.alter_dimensions(nrows=nrows, ncols=ncols)

        # Update tab properties
        self.properties = self._get_properties()
