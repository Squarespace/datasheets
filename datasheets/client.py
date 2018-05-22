import datetime as dt
import httplib2
import json
import os
import types
import warnings

import apiclient
import oauth2client
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

from datasheets import exceptions, helpers
from datasheets.workbook import Workbook


class Client(object):
    def __init__(self, service=False, storage=True, user_agent='Python datasheets library'):
        """Create an authenticated client for interacting with Google Drive and Google Sheets

        Args:
            service (bool): Whether to authenticate as a user or as a service account. If
                service=False, you will be prompted to authorize this instance to access the
                Google Drive attached to one of your Gmail accounts.

                Service-based authorization proceeds using the JSON file located at
                ``$DATASHEETS_SERVICE_PATH`` (default: ``~/.datasheets/service_key.json``).

                User-based authorization proceeds using the client secrets stored at
                ``$DATASHEETS_SECRETS_PATH`` (default: ``~/.datasheets/client_secrets.json``).
                Successful authentication by this method creates a set of credentials. By
                default these credentials are stored at ``$DATASHEETS_CREDENTIALS_PATH``
                (default: ``~/.datasheets/client_credentials.json``), though storage of these
                credentials can be disabled with storage=False.

            storage (bool): Whether to use authorized credentials stored at
                ``$DATASHEETS_CREDENTIALS_PATH`` (or the default of
                ``~/.datasheets/client_credentials.json``).  If credentials have not been stored
                yet or are invalid, whether to store newly obtained credentials at this location.

                If False, authorization will be requested every time a new Client instance is
                created. This mode deliberately does not store credentials to disk in order
                to allow for use of the library in multi-user environments.

            user_agent (str): The user agent tied to new credentials, if new credentials are
                required. This is primarily metadata, and thus unless you have a reason to change
                this the default value is probably fine.
        """
        self.is_service = service
        self.use_storage = storage
        self.user_agent = user_agent

        self.http = self._authenticate()
        self.drive_svc = apiclient.discovery.build('drive', 'v3', http=self.http)
        # Bind sheets_svc directly to .spreadsheets() as the API exposes no other functionality
        self.sheets_svc = apiclient.discovery.build('sheets', 'v4', http=self.http).spreadsheets()

        self._refresh_token_if_needed()

    def __getattribute__(self, attr):
        """Get an attribute (variable or method) of this instance of this class

        For client OAuth, before each user-facing method call this method will verify that the
        access token is not expired and refresh it if it is.

        We only refresh on user-facing method calls since otherwise we'd be refreshing multiple
        times per user action (once for the user call, possibly multiple times for the private
        method calls invoked by it).
        """
        requested_attr = super(Client, self).__getattribute__(attr)

        if isinstance(requested_attr, types.MethodType) \
           and not attr.startswith('_') \
           and hasattr(self, 'credentials'):
            self._refresh_token_if_needed()

        return requested_attr

    def __repr__(self):
        msg = "<{module}.{name}(email='{email}')>"
        return msg.format(module=self.__class__.__module__,
                          name=self.__class__.__name__,
                          email=self.email)

    def _authenticate(self):
        if self.is_service:
            self.credentials = self._get_service_credentials()
        else:
            self.credentials = self._retrieve_client_credentials()
        return self.credentials.authorize(httplib2.Http())

    def _refresh_token_if_needed(self):
        credentials_expired = self.credentials.token_expiry < dt.datetime.utcnow()
        if credentials_expired and not self.is_service:
            self.credentials.refresh(self.http)

    def _retrieve_client_credentials(self):
        """Get valid user credentials

        If storage=True, then first look in ``$DATASHEETS_CREDENTIALS_PATH``
        (default: ``~/.datasheets/client_credentials.json``) to see if authorized credentials have
        been stored. The client credentials are basically an access token that expires. If nothing
        has been stored or if the stored credentials are invalid, the OAuth2 flow is completed to
        obtain new credentials using the client secrets as the means of verifying that the user
        should be granted access. The resulting credentials will be stored at that same path unless
        storage=False, in which case the OAuth2 flow will be executed to obtain new user
        credentials regardless but the credentials will not be stored. The storage=False mode
        deliberately does not store credentials to disk in order to allow for use of the library
        in multi-user environments.
        """
        if self.use_storage:
            unexpanded_credential_path = os.environ.get('DATASHEETS_CREDENTIALS_PATH',
                                                        '~/.datasheets/client_credentials.json')
            credential_path = os.path.expanduser(unexpanded_credential_path)
            store = oauth2client.file.Storage(credential_path)
            # If the client credentials file does not yet exist then oauth2client will throw
            # a warning that isn't useful; to avoid end-user confusion we suppress it
            with warnings.catch_warnings():
                warnings.simplefilter('ignore')
                credentials = store.get()
        else:
            store = helpers._MockStorage()
            credentials = None

        if not credentials or credentials.invalid:
            credentials = self._fetch_new_client_credentials(store)

        self.email = credentials.id_token['email']
        return credentials

    def _fetch_new_client_credentials(self, store):
        """Fetch new user credentials

        Uses the secrets stored at ``$DATASHEETS_SECRETS_PATH``
        (default: ``~/.datasheets/client_secrets.json``).
        """
        unexpanded_client_secrets_path = os.environ.get('DATASHEETS_SECRETS_PATH',
                                                        '~/.datasheets/client_secrets.json')
        client_secrets_path = os.path.expanduser(unexpanded_client_secrets_path)
        scope = ('https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/userinfo.email')
        flow = oauth2client.client.flow_from_clientsecrets(client_secrets_path, scope=scope)
        flow.params['access_type'] = 'offline'  # Allow refreshing expired access tokens
        flow.user_agent = self.user_agent

        with helpers._suppress_stdout():
            with helpers._remove_sys_argv():
                return oauth2client.tools.run_flow(flow, store)

    def _fetch_file_id(self, filename, kind):
        """Return the file_id for the Google Drive file with the specified filename (i.e. title).

        If multiple files exist with the specified filename, an exception is raised with
        information about each file, including its filename, file_id, the datetime it was last
        modified, and a webViewLink URL to look at the file. The user will be directed to use
        the appropriate file_id in their function instead.

        Args:
            filename (str): The name of the workbook we want to fetch the file_id for
            kind (str): Either 'spreadsheet' or 'folder'

        Returns:
            str: The file ID for the specified file
        """
        matches = []
        for f in self._fetch_info_on_items(kind=kind, name=filename):
            matches.append(f)

        if len(matches) == 1:
            return matches[0]['id']
        elif len(matches) == 0 and kind == 'spreadsheet':
            msg = 'Workbook not found. Verify that it is shared with {}'
            raise exceptions.WorkbookNotFound(msg.format(self.email))
        elif len(matches) == 0 and kind == 'folder':
            msg = 'Folder not found. Verify that it is shared with {}'
            raise exceptions.FolderNotFound(msg.format(self.email))

        # Multple matches occurred; format the matches for printing with the exception
        template = """\n\n{}\nfilename: {}\nfile_id: {}\nmodifiedTime: {}\nwebViewLink: {}"""
        formatted_output = ''
        for i, row in enumerate(matches):
            cleaned_time = row['modifiedTime'].replace('T', ' ').replace('Z', '')
            this_row = template.format(i, row['name'], row['id'], cleaned_time, row['webViewLink'])
            formatted_output += this_row

        msg = ('Multiple workbooks founds. Please choose the correct file_id below '
               'and provide it to your function instead of a filename:')
        raise exceptions.MultipleWorkbooksFound(msg + formatted_output)

    def _fetch_info_on_items(self, kind, folder=None, name=None, only_mine=False,
                             fields='files(name,id,modifiedTime,webViewLink)'):
        """Return info on workbooks or folders shared with the user

        Return a list of dicts, with each list representing one workbook or folder shared
        with the user


        Args:
            kind (str): Either 'spreadsheet' or 'folder'

            folder (str): An optional folder name to limit the results to

            name (str): An optional file name to limit the results to

            only_mine (bool): If True, only fetch items owned by the current user

            fields (str): The fields to return in the results


        Returns:
            list: A list of dicts, one dict per workbook or folder shared with the user
        """
        fields = 'nextPageToken, {}'.format(fields)
        query = "mimeType='application/vnd.google-apps.{}'".format(kind)
        if folder:
            folder_id = self._fetch_file_id(folder, kind='folder')
            query += " and '{}' in parents".format(folder_id)

        if name:
            escaped_name = helpers._escape_query(name)
            query += " and name = '{}'".format(escaped_name)

        if only_mine:
            escaped_email = helpers._escape_query(self.email)
            query += " and '{}' in owners".format(escaped_email)

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

    def _get_service_credentials(self):
        """Get credentials for a service account

        Uses the service key stored at ``$DATASHEETS_SERVICE_PATH``
        (default: ``~/.datasheets/service_key.json``) to get service account credentials. At the
        time this method returns the instance has not yet been authenticated / authorized.

        Returns:
            oauth2client.service_account.ServiceAccountCredentials: instance of service credentials
        """
        unexpanded_service_key_path = os.environ.get('DATASHEETS_SERVICE_PATH',
                                                     '~/.datasheets/service_key.json')
        service_key_path = os.path.expanduser(unexpanded_service_key_path)
        with open(service_key_path) as f:
            keyfile_dict = json.load(f)

        self.email = keyfile_dict['client_email']  # used in __repr__
        return ServiceAccountCredentials.from_json_keyfile_dict(
            keyfile_dict=keyfile_dict,
            scopes=('https://www.googleapis.com/auth/drive',)
        )

    def create_workbook(self, filename, folders=()):
        """Create a blank workbook with the specific filename

        Args:
            filename (str): The name to give to the new workbook
            folders (tuple): An optional tuple of folders to add the new workbook to

        Returns:
            datasheets.Workbook: An instance of the newly created workbook
        """
        root_file_id = self.drive_svc.files().get(fileId='root', fields='id').execute()['id']
        folders = [self._fetch_file_id(filename=f, kind='folder') for f in folders]
        folders = [root_file_id] + list(folders)

        body = {
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'name': filename,
            'parents': folders
        }
        self.drive_svc.files().create(body=body).execute()
        return self.fetch_workbook(filename=filename)

    def delete_workbook(self, filename=None, file_id=None):
        """Delete a workbook from Google Drive

        Either filename (i.e. title) or file_id should be provided. Providing file_id is preferred
        as it is more precise.

        Args:
            filename (str): The name of the workbook to delete
            file_id (str): The ID of the workbook to delete

        Returns:
            None
        """
        if (filename is None) == (file_id is None):
            raise ValueError('Either filename or file_id must be provided, but not both.')
        if not file_id:
            file_id = self._fetch_file_id(filename=filename, kind='spreadsheet')
        self.drive_svc.files().delete(fileId=file_id).execute()

    def fetch_folders(self, only_mine=False):
        """Fetch all folders shared with this account

        Args:
            only_mine (bool): If True, limit results to only those folders owned by this user

        Returns:
            pandas.DataFrame: One row per folder listing folder name, ID, most recent modified
            time, and webview link to the folder
        """
        fields = 'files(name,id,modifiedTime,webViewLink,owners(me))'
        raw_info = self._fetch_info_on_items(kind='folder', only_mine=only_mine, fields=fields)
        return pd.DataFrame(raw_info, columns=['name', 'id', 'modifiedTime', 'webViewLink'])

    def fetch_workbook(self, filename=None, file_id=None):
        """Fetch a workbook

        Either filename (i.e. title) or file_id should be provided. Providing file_id
        is preferred as it is more precise.

        Args:
            filename (str): The name of the workbook to fetch
            file_id (str): The ID of the workbook to fetch

        Returns:
            datasheets.Workbook: An instance of the requested workbook
        """
        if (filename is None) == (file_id is None):
            raise ValueError('Either filename or file_id must be provided, but not both.')
        if not file_id:
            file_id = self._fetch_file_id(filename=filename, kind='spreadsheet')
        return Workbook(filename, file_id, self, self.drive_svc, self.sheets_svc)

    def fetch_workbooks_info(self, folder=None):
        """Fetch information on all workbooks shared with this account

        Args:
            folder (str): An optional folder name to limit the results to

        Returns:
            pandas.DataFrame: One row per workbook listing workbook name, ID, most recent
            modified time, and webview link to the workbook
        """
        raw_info = self._fetch_info_on_items(kind='spreadsheet', folder=folder)
        return pd.DataFrame(raw_info, columns=['name', 'id', 'modifiedTime', 'webViewLink'])
