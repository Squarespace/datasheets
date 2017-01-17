"""
Tests for gsheets.core module, all of which are mocked and require no network access
"""
import sys

import apiclient
import pandas as pd
import pytest

import gsheets


def yaml_resource(filename):
    ''' A decorator to pass correct .yaml file to test function '''
    return pytest.mark.parametrize('mock_client', [filename], indirect=True)


class TestClient:
    @yaml_resource('drive.yaml')
    def test_init(self, mock_client):
        # sheets_svc variable properly created
        assert isinstance(mock_client.sheets_svc, apiclient.discovery.Resource)
        assert hasattr(mock_client.sheets_svc, 'sheets') and hasattr(mock_client.sheets_svc, 'values')

        # drive_svc variable properly created
        assert isinstance(mock_client.drive_svc, apiclient.discovery.Resource)
        assert hasattr(mock_client.drive_svc, 'files') and hasattr(mock_client.drive_svc, 'permissions')

        assert mock_client.email is None

        repr_start = "<gsheets.core.Client(email='None')"
        assert repr(mock_client).startswith(repr_start)

    @yaml_resource('drive.yaml')
    def test_get_file_id_findable_workbook(self, mock_client):
        file_id = mock_client._get_file_id(filename='gsheets_test_2', kind='spreadsheet')
        assert file_id == 'xyz2345'

    @yaml_resource('drive.yaml')
    def test_get_file_id_findable_folder(self, mock_client):
        file_id = mock_client._get_file_id(filename='gsheets_test_folder_2', kind='folder')
        assert file_id == 'xyz6789'

    @yaml_resource('drive.yaml')
    def test_get_file_id_workbook_not_found(self, mock_client):
        with pytest.raises(gsheets.helpers.WorkbookNotFound) as err:
            mock_client._get_file_id(filename='nonexistent_folder', kind='spreadsheet')
        assert err.match('Workbook not found. Verify that it is shared with {}'.format(mock_client.email))

    @yaml_resource('drive.yaml')
    def test_get_file_id_folder_not_found(self, mock_client):
        with pytest.raises(gsheets.helpers.FolderNotFound) as err:
            mock_client._get_file_id(filename='nonexistent_file', kind='folder')
        assert err.match('Folder not found. Verify that it is shared with {}'.format(mock_client.email))

    @yaml_resource('drive.yaml')
    def test_get_file_id_with_duplicates(self, mock_client):
        filename = 'gsheets_test_3'
        with pytest.raises(gsheets.helpers.MultipleWorkbooksFound) as err:
            mock_client._get_file_id(filename=filename, kind='spreadsheet')

        # Rather than checking the whole message, just check that each component is there
        base_msg = ('Multiple workbooks founds. Please choose the correct file_id below '
                    'and provide it to your function instead of a filename:')
        assert err.match(base_msg)
        assert err.match('filename: {}'.format(filename))
        assert err.match('file_id: xyz3456')
        assert err.match('file_id: xyz4567')
        assert err.match('modifiedTime')
        assert err.match('webViewLink')

    @yaml_resource('drive.yaml')
    def test_delete_workbook_which_exists(self, mock_client):
        mock_client.delete_workbook(filename='gsheets_test_2')

    @yaml_resource('drive.yaml')
    def test_delete_workbook_nonexistent(self, mock_client):
        with pytest.raises(gsheets.helpers.WorkbookNotFound) as err:
            mock_client.delete_workbook(filename='test_delete_workbook')
        assert err.match('Workbook not found. Verify that it is shared with {}'.format(mock_client.email))

    @yaml_resource('drive.yaml')
    def test_delete_workbook_by_file_id(self, mock_client):
        mock_client.delete_workbook(file_id='xyz3456')

    @yaml_resource('drive.yaml')
    def test_get_workbook(self, mock_client):
        filename = 'gsheets_test_1'
        wkb = mock_client.get_workbook(filename)
        assert isinstance(wkb, gsheets.Workbook)
        assert wkb.filename == filename
        assert wkb.file_id == 'xyz1234'

    @yaml_resource('drive.yaml')
    def test_get_workbook_by_file_id(self, mock_client):
        mock_client.get_workbook(file_id='xyz3456')

    @yaml_resource('drive.yaml')
    def test_show_workbooks(self, mock_client):
        # Note: It's not currently possible to test the `folder` parameter for show_workbooks
        # because that is ultimately implemented as part of the query to the API, which isn't
        # contained in the HTTP body and thus isn't checkable via apiclient.http.RequestMockBuilder
        results = mock_client.show_workbooks()
        assert isinstance(results, pd.DataFrame)
        assert set(results.columns) == set(['name', 'modifiedTime', 'webViewLink', 'id'])
        assert results.shape == (6, 4)

    @yaml_resource('test_show_folders.yaml')
    def test_show_folders_all(self, mock_client):
        results = mock_client.show_folders()
        assert isinstance(results, pd.DataFrame)
        assert set(results.columns) == set(['name', 'modifiedTime', 'webViewLink', 'id'])
        assert results.shape == (5, 4)

    @yaml_resource('test_show_folders.yaml')
    def test_show_folders_only_mine(self, mock_client):
        results = mock_client.show_folders(only_mine=True)
        assert isinstance(results, pd.DataFrame)
        assert set(results.columns) == set(['name', 'modifiedTime', 'webViewLink', 'id'])
        assert results.shape == (2, 4)


class TestCreateWorkbook:
    @yaml_resource('test_create_workbook_no_folder.yaml')
    def test_create_workbook_no_folder(self, mock_client):
        filename = 'test_create_workbook'
        wkb = mock_client.create_workbook(filename)
        assert isinstance(wkb, gsheets.Workbook)
        assert wkb.filename == filename
        assert wkb.file_id == 'xyz4567'

    @yaml_resource('test_create_workbook_with_folder.yaml')
    def test_create_workbook_with_folder(self, mock_client):
        filename = 'test_create_workbook'
        wkb = mock_client.create_workbook(filename, folders=('gsheets_test_folder_1',))
        assert isinstance(wkb, gsheets.Workbook)
        assert wkb.filename == filename
        assert wkb.file_id == 'xyz4567'

    @yaml_resource('test_create_workbook_multiple_folders.yaml')
    def test_create_workbook_multiple_folders(self, mock_client):
        filename = 'test_create_workbook'
        wkb = mock_client.create_workbook(filename, folders=('gsheets_test_folder_1', 'gsheets_test_folder_2'))
        assert isinstance(wkb, gsheets.Workbook)
        assert wkb.filename == filename
        assert wkb.file_id == 'xyz4567'


class TestWorkbook:
    @yaml_resource('drive.yaml')
    def test_init(self, mock_client):
        mock_workbook = mock_client.get_workbook('gsheets_test_1')
        assert mock_workbook.filename == 'gsheets_test_1'
        assert mock_workbook.file_id == 'xyz1234'

        # sheets_svc variable properly created
        assert isinstance(mock_workbook.sheets_svc, apiclient.discovery.Resource)
        assert hasattr(mock_workbook.sheets_svc, 'sheets') and hasattr(mock_workbook.sheets_svc, 'values')

        # drive_svc variable properly created
        assert isinstance(mock_workbook.drive_svc, apiclient.discovery.Resource)
        assert hasattr(mock_workbook.drive_svc, 'files') and hasattr(mock_workbook.drive_svc, 'permissions')

        assert mock_workbook.url == 'https://docs.google.com/spreadsheets/d/xyz1234'

        repr_start = "<gsheets.core.Workbook(filename='gsheets_test_1')"
        assert repr(mock_workbook).startswith(repr_start)

    @yaml_resource('drive.yaml')
    def test_get_permission_id(self, mock_workbook):
        permission_id = mock_workbook._get_permission_id('get_permission_id@testdomain.test')
        assert permission_id == 15375041453037031415

    @yaml_resource('drive.yaml')
    def test_get_permission_id_nonexistent(self, mock_workbook):
        with pytest.raises(gsheets.helpers.PermissionNotFound) as err:
            mock_workbook._get_permission_id('nonexistent@testdomain.test')
        err.match("Permission for email 'nonexistent@testdomain.test' not found for workbook 'gsheets_test_1'")

    @yaml_resource('drive.yaml')
    def test_add_permission(self, mock_workbook):
        mock_workbook.add_permission(email='add_permission@testdomain.test')

    @yaml_resource('drive.yaml')
    def test_show_permissions(self, mock_workbook):
        expected = pd.DataFrame([{'email': 'get_permission_id@testdomain.test', 'role': 'owner'},
                                 {'email': "User Type: 'domain'", 'role': 'writer'}])
        output = mock_workbook.show_permissions()
        assert output.equals(expected)

    @yaml_resource('drive.yaml')
    def test_remove_permission(self, mock_workbook):
        mock_workbook.remove_permission('get_permission_id@testdomain.test')

    @yaml_resource('test_batch_update.yaml')
    def test_batch_update(self, mock_workbook):
        request_body = {'addSheet': {
                            'properties': {
                                  'title': 'test_batch_update',
                                  'gridProperties': {
                                        'rowCount': 15,
                                        'columnCount': 5
                                        },
                                  }
                            }
                        }
        body = {'requests': [request_body]}
        mock_workbook.batch_update(body=body)

    @yaml_resource('test_create_tab.yaml')
    def test_create_tab(self, mock_workbook):
        tab = mock_workbook.create_tab('test_create_tab', nrows=20, ncols=10)
        assert isinstance(tab, gsheets.core.Tab)
        assert tab.tabname == 'test_create_tab'

    @yaml_resource('test_delete_tab.yaml')
    def test_delete_tab(self, mock_workbook):
        mock_workbook.delete_tab('test_delete_tab')

    @yaml_resource('test_show_all_tabs.yaml')
    def test_show_all_tabs(self, mock_workbook):
        expected = pd.DataFrame(['test_tab_1', 'test_tab_2', 'test_tab_3'], columns=['Tabs'])
        assert mock_workbook.show_all_tabs().equals(expected)

    @yaml_resource('test_get_tab.yaml')
    def test_get_tab(self, mock_workbook):
        tab = mock_workbook.get_tab('test_get_tab')
        assert isinstance(tab, gsheets.core.Tab)
        assert tab.tabname == 'test_get_tab'
        assert tab.filename == mock_workbook.filename

    @pytest.mark.skipif(sys.version_info.major == 3, reason="Breaks on Python3")
    @yaml_resource('test_get_tab_not_found.yaml')
    def test_get_tab_not_found(self, mock_workbook):
        with pytest.raises(gsheets.helpers.TabNotFound) as err:
            mock_workbook.get_tab('nonexistent_tab')
        assert err.match('The given tab could not be found. Error generated:')


class TestTab:
    @yaml_resource('test_tab_init.yaml')
    def test_init(self, mock_tab):
        assert isinstance(mock_tab, gsheets.core.Tab)
        assert mock_tab.tabname == 'test_tab'

        # sheets_svc variable properly created
        assert isinstance(mock_tab.sheets_svc, apiclient.discovery.Resource)
        assert hasattr(mock_tab.sheets_svc, 'sheets') and hasattr(mock_tab.sheets_svc, 'values')

        # drive_svc variable properly created
        assert isinstance(mock_tab.drive_svc, apiclient.discovery.Resource)
        assert hasattr(mock_tab.drive_svc, 'files') and hasattr(mock_tab.drive_svc, 'permissions')

        expected_keys = ['sheetType', 'index', 'sheetId', 'gridProperties', 'title']
        assert set(mock_tab.properties.keys()) == set(expected_keys)

        assert mock_tab.ncols == 26
        assert mock_tab.nrows == 1000
        assert mock_tab.tab_id == 1945916913
        assert mock_tab.url == 'https://docs.google.com/spreadsheets/d/xyz1234#gid=1945916913'

        repr_start = "<gsheets.core.Tab(filename='gsheets_test_1', tabname='test_tab') at"
        assert repr(mock_tab).startswith(repr_start)

    @yaml_resource('test_add_rows_or_columns.yaml')
    def test_add_rows_or_columns(self, mock_tab):
        mock_tab._add_rows_or_columns(kind='ROWS', n='1234')

    @yaml_resource('test_align_cells.yaml')
    def test_align_cells(self, mock_tab):
        mock_tab.align_cells()

    @yaml_resource('test_alter_dimensions.yaml')
    def test_alter_dimensions(self, mock_tab):
        mock_tab.alter_dimensions(nrows=123)

    @yaml_resource('test_autosize_columns.yaml')
    def test_autosize_columns(self, mock_tab):
        mock_tab.autosize_columns()

    @yaml_resource('test_format_font.yaml')
    def test_format_font(self, mock_tab):
        mock_tab.format_font()

    @yaml_resource('test_format_headers.yaml')
    def test_format_headers(self, mock_tab):
        mock_tab.format_headers(nrows=3)

    @yaml_resource('test_clear_data.yaml')
    def test_clear_data(self, mock_tab):
        mock_tab.clear_data()

    @yaml_resource('test_tab_init.yaml')
    def test_get_parent_workbook(self, mock_workbook):
        mock_tab = mock_workbook.get_tab('test_tab')
        parent = mock_tab.get_parent_workbook()
        assert parent.filename == mock_workbook.filename
        assert parent.file_id == mock_workbook.file_id
        assert isinstance(parent, gsheets.core.Workbook)

    @yaml_resource('test_tab_init.yaml')
    def test_process_rows_empty_tab(self, mock_tab):
        ''' Verify the raw_data an empty tab would return is processed correctly '''
        raw_data = {u'sheets': [{u'data': [{}]}]}
        assert [] == mock_tab._process_rows(raw_data)


class TestGetData:
    @yaml_resource('test_tab_init.yaml')
    @pytest.mark.parametrize('get_raw_resource', ['test_process_rows.yaml'], indirect=True)
    def test_process_rows(self, mock_tab, get_raw_resource, expected_data):
        assert mock_tab._process_rows(get_raw_resource) == expected_data

    @yaml_resource('test_get_data_empty.yaml')
    def test_get_data_empty(self, mock_tab):
        assert mock_tab.get_data().equals(pd.DataFrame([]))
        assert mock_tab.get_data(fmt='dict') == []
        assert mock_tab.get_data(fmt='list') == ([], [])
        assert mock_tab.get_data(headers=False).equals(pd.DataFrame([]))
        assert mock_tab.get_data(headers=False, fmt='dict') == []
        assert mock_tab.get_data(headers=False, fmt='list') == ([], [])

    @yaml_resource('test_tab_init.yaml')
    def test_get_data_unexpected_fmt(self, mock_tab):
        with pytest.raises(ValueError) as err:
            mock_tab.get_data(fmt='foo')
        assert err.match("Unexpected value 'foo' for parameter `fmt`. "
                         "Accepted values are 'df', 'dict', and 'list'")

    @yaml_resource('test_get_data.yaml')
    def test_get_data_df_with_headers(self, mock_tab, expected_cleaned_data_with_headers):
        expected = pd.DataFrame(data=expected_cleaned_data_with_headers[1],
                                columns=expected_cleaned_data_with_headers[0])
        assert mock_tab.get_data().equals(expected)

    @yaml_resource('test_get_data.yaml')
    def test_get_data_df_no_headers(self, mock_tab, expected_cleaned_data_no_headers):
        expected = pd.DataFrame(data=expected_cleaned_data_no_headers[1],
                                columns=expected_cleaned_data_no_headers[0])
        assert mock_tab.get_data(headers=False).equals(expected)

    @yaml_resource('test_get_data.yaml')
    def test_get_data_list_with_headers(self, mock_tab, expected_cleaned_data_with_headers):
        assert mock_tab.get_data(fmt='list') == expected_cleaned_data_with_headers

    @yaml_resource('test_get_data.yaml')
    def test_get_data_list_no_headers(self, mock_tab, expected_cleaned_data_no_headers):
        assert mock_tab.get_data(fmt='list', headers=False) == expected_cleaned_data_no_headers

    @yaml_resource('test_get_data.yaml')
    def test_get_data_dict_with_headers(self, mock_tab, expected_cleaned_data_with_headers):
        keys = expected_cleaned_data_with_headers[0]
        values = expected_cleaned_data_with_headers[1]
        expected = [dict(zip(keys, row)) for row in values]
        assert mock_tab.get_data(fmt='dict') == expected

    @yaml_resource('test_get_data.yaml')
    def test_get_data_dict_no_headers(self, mock_tab, expected_cleaned_data_no_headers):
        keys = expected_cleaned_data_no_headers[0]
        values = expected_cleaned_data_no_headers[1]
        expected = [dict(zip(keys, row)) for row in values]
        assert mock_tab.get_data(fmt='dict', headers=False) == expected


class TestUploadData:
    @yaml_resource('test_upload_data.yaml')
    def test_upload_data_incorrect_mode(self, mock_tab):
        with pytest.raises(ValueError) as err:
            mock_tab.upload_data(data=[], mode='nonexistent')
        assert err.match("Unexpected value 'nonexistent' for parameter `mode`. "
                         "Accepted values are 'insert' and 'append'")

    @yaml_resource('test_upload_data.yaml')
    def test_upload_data_insert(self, mock_tab, expected_data):
        mock_tab.upload_data(data=expected_data, mode='insert', autoformat=False)

    @yaml_resource('test_upload_data.yaml')
    def test_upload_data_append(self, mock_tab, expected_data):
        mock_tab.upload_data(data=expected_data, mode='append', autoformat=False)


class TestConvenience:
    @yaml_resource('test_get_tab.yaml')
    def test_open_tab(self, mock_client):
        tab = gsheets.open_tab('gsheets_test_1', 'test_get_tab', client=mock_client)
        assert isinstance(tab, gsheets.core.Tab)
        assert tab.tabname == 'test_get_tab'
        assert tab.filename == 'gsheets_test_1'

    @yaml_resource('test_create_tab_in_existing_workbook.yaml')
    def test_create_tab_in_existing_workbook(self, mock_client):
        tab = gsheets.create_tab_in_existing_workbook('existing_workbook', 'new_tab', client=mock_client)
        assert isinstance(tab, gsheets.core.Tab)
        assert tab.tabname == 'new_tab'
        assert tab.filename == 'existing_workbook'

    @yaml_resource('test_create_tab_in_new_workbook.yaml')
    def test_create_tab_in_new_workbook(self, mock_client):
        # Note: this test doesn't check the request body for sheets.spreadsheets.batchUpdate
        # as two requests to it are made (addSheet to create a tab and deleteSheet to remove Sheet1)
        tab = gsheets.create_tab_in_new_workbook('new_workbook', 'new_tab', client=mock_client)
        assert isinstance(tab, gsheets.core.Tab)
        assert tab.tabname == 'new_tab'
        assert tab.filename == 'new_workbook'
