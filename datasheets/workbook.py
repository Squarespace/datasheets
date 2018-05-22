import types

import pandas as pd

from datasheets import exceptions
from datasheets.tab import Tab


class Workbook(object):
    def __init__(self, filename, file_id, client, drive_svc, sheets_svc):
        """Create a datasheets.Workbook instance of an existing Google Sheets doc

        This class in not intended to be directly instantiated; it is created by
        datasheets.Client.fetch_workbook().

        Args:
            filename (str): The name of the workbook
            file_id (str): The Google Sheets-assigned ID for the file
            client (datasheets.Client): The client instance that instantiated this workbook
            drive_svc (googleapiclient.discovery.Resource): An instance of Google Drive
            sheets_svc (googleapiclient.discovery.Resource): An instance of Google Sheets
        """
        # Note that work is not done here to verify that the Workbook exists because
        # the assumed mechanism of instantiation via sheets.Client already assures this
        # by searching for the file_id through the sheets.Client._get_file_id method.
        self.filename = filename
        self.file_id = file_id
        self._client = client
        self.drive_svc = drive_svc
        self.sheets_svc = sheets_svc
        self.url = 'https://docs.google.com/spreadsheets/d/{}'.format(self.file_id)

    def __getattribute__(self, attr):
        """Get an attribute (variable or method) of this instance of this class

        For client OAuth, before each user-facing method call this method will verify that the
        access token is not expired and refresh it if it is.

        We only refresh on user-facing method calls since otherwise we'd be refreshing multiple
        times per user action (once for the user call, possibly multiple times for the private
        method calls invoked by it).
        """
        requested_attr = super(Workbook, self).__getattribute__(attr)

        if isinstance(requested_attr, types.MethodType) \
           and not attr.startswith('_'):
            self.client._refresh_token_if_needed()

        return requested_attr

    def __repr__(self):
        msg = "<{module}.{name}(filename='{filename}')>"
        return msg.format(module=self.__class__.__module__,
                          name=self.__class__.__name__,
                          filename=self.filename)

    @property
    def client(self):
        """ Property for the client instance that instantiated this workbook """
        return self._client

    def _fetch_permission_id(self, email):
        """ Return the permission_id associated with the given email address """
        permissions = self.drive_svc.permissions().list(fileId=self.file_id).execute()
        for perm in permissions.get('permissions', tuple()):
            this_entry = self.drive_svc.permissions().get(fileId=self.file_id,
                                                          permissionId=perm['id'],
                                                          fields='emailAddress').execute()
            if this_entry['emailAddress'] == email:
                return perm['id']

        msg = "Permission for email '{}' not found for workbook '{}'"
        raise exceptions.PermissionNotFound(msg.format(email, self.filename))

    def share(self, email, role='reader', notify=True, message=None):
        """Share this workbook with someone.

        Args:
            email (str): The email address to share the workbook with

            role (str): The type of permission to grant. Values must be one of 'owner',
                'writer', or 'reader'

            notify (bool): If True, send an email notifying the recipient of their granted
                permission.  These notification emails are the same as what Google sends when
                a document is shared through Google Drive

            message (str): If notify is True, the message to send with the email notification

        Returns:
            None
        """
        new_permission = {
            'emailAddress': email,
            'type': 'user',
            'role': role
        }
        self.drive_svc.permissions().create(fileId=self.file_id,
                                            body=new_permission,
                                            emailMessage=message,
                                            sendNotificationEmail=notify).execute()

    def batch_update(self, body):
        """Apply updates to a workbook or tab using Google Sheets' spreadsheets.batchUpdate method

        Args:
            body (list): A list of requests, with each request provided as a dict

        The Google Sheets batch update method is a flexible method exposed by the Google Sheets
        API that permits effectively all functionality allowed for by the Google Sheets.

        For more details on how to use this method, see the following:

            Explanation of spreadsheets.batchUpdate method:

                https://developers.google.com/sheets/reference/rest/v4/spreadsheets/batchUpdate

            List of available request types and the parameters they take:

                https://developers.google.com/sheets/reference/rest/v4/spreadsheets/request

        Example:
            The following is an example request that would perform only one update
            operation (the 'repeatCell' operation)::

                request_body = {
                    'repeatCell': {
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

        """
        self.sheets_svc.batchUpdate(spreadsheetId=self.file_id, body=body).execute()

    def create_tab(self, tabname, nrows=1000, ncols=26):
        """Create a new tab in the given workbook

        Args:
            tabname (str): The name for the tab

            nrows (int): An integer number of rows for the tab to have. The Google Sheets
                default is 1000

            ncols (int): An integer number of columns for the tab to have. The Google Sheets
                default is 26

        Returns:
            datasheets.Tab: An instance of the newly created tab
        """
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
        self.batch_update(body=body)
        return self.fetch_tab(tabname)

    def delete_tab(self, tabname):
        """Delete a tab with the given name from the current workbook

        Args:
            tabname (str): The name of the tab to delete

        Returns:
            None
        """
        tab_id = self.fetch_tab(tabname).tab_id
        request_body = {'deleteSheet': {'sheetId': tab_id}}
        body = {'requests': [request_body]}
        self.batch_update(body=body)

    def fetch_tab(self, tabname):
        """Return a datasheets.Tab instance of the given tab associated with this workbook

        Args:
            tabname (str): The name of the tab to fetch

        Returns:
            datasheets.Tab: An instance of the requested tab

        """
        return Tab(tabname, self, self.drive_svc, self.sheets_svc)

    def unshare(self, email):
        """Unshare this workbook with someone.

        Args:
            email (str): The email address that will be unshared

        Returns:
            None
        """
        permission_id = self._fetch_permission_id(email)
        self.drive_svc.permissions().delete(fileId=self.file_id,
                                            permissionId=permission_id).execute()

    def fetch_tab_names(self):
        """Show the names of the tabs within the workbook, returned as a pandas.DataFrame.

        Returns:
            pandas.DataFrame: One row per tabname within the workbook
        """
        workbook = self.sheets_svc.get(spreadsheetId=self.file_id,
                                       fields='sheets/properties/title').execute()
        tab_names = [tab['properties']['title'] for tab in workbook['sheets']]
        return pd.DataFrame(tab_names, columns=['Tabs'])

    def fetch_permissions(self):
        """Fetch information on who is shared on this workbook and their permission level

        Returns:
            pandas.DataFrame: One row per email address shared, including the permission level
            that that email has been granted
        """
        req = self.drive_svc.permissions().list(fileId=self.file_id,
                                                fields='permissions(id,role,type)')
        perm_ids = req.execute()['permissions']

        permissions = []
        for perm in perm_ids:
            email = self.drive_svc.permissions().get(
                fileId=self.file_id, permissionId=perm['id'],
                fields='emailAddress').execute().get('emailAddress')

            if not email:
                email = "User Type: '{}'".format(perm['type'])

            permissions.append(dict(
                role=perm['role'],
                email=email
            ))

        return pd.DataFrame(data=permissions)
