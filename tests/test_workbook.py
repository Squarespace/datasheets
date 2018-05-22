import apiclient
import httplib2
import pandas as pd
import pytest

import datasheets


def test_getattribute_for_non_method(mock_workbook):
    # Also make sure we actually get something back from the non-method call
    assert mock_workbook.filename == 'datasheets_test_1'
    # _refresh_token_if_needs is called in __init__(); verify it wasn't called again
    assert mock_workbook.client._refresh_token_if_needed.call_count == 1


def test_getattribute_for_private_method(mocker, mock_workbook):
    mocker.patch.object(mock_workbook, '_fetch_permission_id')
    mock_workbook._fetch_permission_id('test@test.test')
    # _refresh_token_if_needs is called in __init__(); verify it wasn't called again
    assert mock_workbook.client._refresh_token_if_needed.call_count == 1


def test_getattribute_for_user_facing_method_with_credentials(mocker, mock_workbook):
    mocker.patch.object(mock_workbook, 'share')
    mock_workbook.share('test@test.test')
    # _refresh_token_if_needs is called in __init__(); verify it was called a second time
    assert mock_workbook.client._refresh_token_if_needed.call_count == 2


def test_init(mocker, mock_client):
    mocker.patch.object(mock_client, '_fetch_file_id', autospec=True, return_value='xyz1234')

    mock_workbook = mock_client.fetch_workbook('datasheets_test_1')
    assert mock_workbook.filename == 'datasheets_test_1'
    assert mock_workbook.file_id == 'xyz1234'

    # sheets_svc variable properly created
    assert isinstance(mock_workbook.sheets_svc, apiclient.discovery.Resource)
    assert hasattr(mock_workbook.sheets_svc, 'sheets')
    assert hasattr(mock_workbook.sheets_svc, 'values')

    # drive_svc variable properly created
    assert isinstance(mock_workbook.drive_svc, apiclient.discovery.Resource)
    assert hasattr(mock_workbook.drive_svc, 'files')
    assert hasattr(mock_workbook.drive_svc, 'permissions')

    assert mock_workbook.url == 'https://docs.google.com/spreadsheets/d/xyz1234'

    repr_start = "<datasheets.workbook.Workbook(filename='datasheets_test_1')>"
    assert repr(mock_workbook).startswith(repr_start)


def test_fetch_permission_id(mocker, mock_workbook):
    mocked_drive_svc = mocker.patch.object(mock_workbook, 'drive_svc', autospec=True)
    mocked_drive_svc.permissions().list().execute.return_value = {
        'kind': 'drive#permissionList',
        'permissions': [
            {'kind': 'drive#permission',
             'id': '48004950760004877923',
             'type': 'user',
             'role': 'owner'},
            {'kind': 'drive#permission',
             'id': '15012643990489651114',
             'type': 'user',
             'role': 'reader'}
        ]
    }
    mocked_drive_svc.permissions().get().execute.side_effect = [
        {'emailAddress': 'wrong@email.test'},
        {'emailAddress': 'get_permission_id@testdomain.test'},
    ]

    permission_id = mock_workbook._fetch_permission_id('get_permission_id@testdomain.test')
    assert permission_id == '15012643990489651114'

    mocked_drive_svc.permissions().list.assert_any_call(fileId=mock_workbook.file_id)
    mocked_drive_svc.permissions().get.assert_any_call(fileId=mock_workbook.file_id,
                                                       permissionId='48004950760004877923',
                                                       fields='emailAddress')
    mocked_drive_svc.permissions().get.assert_any_call(fileId=mock_workbook.file_id,
                                                       permissionId='15012643990489651114',
                                                       fields='emailAddress')
    assert mocked_drive_svc.permissions().get().execute.call_count == 2


def test_fetch_permission_id_nonexistent(mocker, mock_workbook):
    mocked_drive_svc = mocker.patch.object(mock_workbook, 'drive_svc', autospec=True)
    mocked_drive_svc.permissions().list().execute.return_value = {}

    with pytest.raises(datasheets.exceptions.PermissionNotFound) as err:
        mock_workbook._fetch_permission_id('nonexistent@testdomain.test')
    err.match("Permission for email 'nonexistent@testdomain.test' not found for workbook 'datasheets_test_1'")
    mocked_drive_svc.permissions().list.assert_any_call(fileId=mock_workbook.file_id)


def test_share(mocker, mock_workbook):
    mocked_drive_svc = mocker.patch.object(mock_workbook, 'drive_svc', autospec=True)
    email = 'add_permission@testdomain.test'

    mock_workbook.share(email=email)

    assert mocked_drive_svc.permissions().create.call_count == 1
    _, _, kwargs = mocked_drive_svc.permissions().create.mock_calls[0]
    assert kwargs['fileId'] == mock_workbook.file_id
    assert kwargs['sendNotificationEmail'] is True
    assert kwargs['emailMessage'] is None
    assert kwargs['body']['emailAddress'] == email


def test_fetch_permissions(mocker, mock_workbook):
    mocked_drive_svc = mocker.patch.object(mock_workbook, 'drive_svc', autospec=True)
    mocked_drive_svc.permissions().list().execute.return_value = {
        'permissions': [
            {'id': '12604950761524962923', 'type': 'user', 'role': 'owner'},
            {'id': '13232714134634019830', 'type': 'group', 'role': 'writer'},
            {'id': '13845842511136751920k', 'type': 'domain', 'role': 'commenter'},
        ]
    }
    mocked_drive_svc.permissions().get().execute.side_effect = [
        {'emailAddress': 'fetch_permission_id@testdomain.test'},
        {'emailAddress': 'some_group@testdomain.test'},
        {},
    ]

    expected = pd.DataFrame([
        {'email': 'fetch_permission_id@testdomain.test', 'role': 'owner'},
        {'email': 'some_group@testdomain.test', 'role': 'writer'},
        {'email': "User Type: 'domain'", 'role': 'commenter'},
    ])
    output = mock_workbook.fetch_permissions()
    assert output.equals(expected)

    mocked_drive_svc.permissions().list.assert_called_with(fileId=mock_workbook.file_id,
                                                           fields='permissions(id,role,type)')
    assert mocked_drive_svc.permissions().get().execute.call_count == 3


def test_unshare(mocker, mock_workbook):
    mocked_fetch_permission_id = mocker.patch.object(mock_workbook, '_fetch_permission_id',
                                                     autospec=True, return_value='15383')
    mocked_drive_svc = mocker.patch.object(mock_workbook, 'drive_svc', autospec=True)

    mock_workbook.unshare('get_permission_id@testdomain.test')

    mocked_fetch_permission_id.assert_called_with(email='get_permission_id@testdomain.test')
    mocked_drive_svc.permissions().delete.assert_any_call(fileId=mock_workbook.file_id,
                                                          permissionId='15383')


def test_batch_update(mocker, mock_workbook):
    mocked_sheets_svc = mocker.patch.object(mock_workbook, 'sheets_svc', autospec=True)
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

    mocked_sheets_svc.batchUpdate.assert_any_call(spreadsheetId=mock_workbook.file_id, body=body)
    mocked_sheets_svc.batchUpdate().execute.assert_called_once()


def test_create_tab(mocker, mock_workbook):
    filename = 'test_create_tab'
    mocked_batch_update = mocker.patch.object(mock_workbook, 'batch_update', autospec=True)
    mocked_fetch_tab = mocker.patch.object(mock_workbook, 'fetch_tab')

    mock_workbook.create_tab(filename, nrows=20, ncols=10)

    mocked_fetch_tab.assert_any_call(filename)
    _, _, kwargs = mocked_batch_update.mock_calls[0]
    assert len(kwargs['body']['requests']) == 1
    properties = kwargs['body']['requests'][0]['addSheet']['properties']
    assert properties['gridProperties'] == {'columnCount': 10, 'rowCount': 20}
    assert properties['title'] == 'test_create_tab'


def test_delete_tab(mocker, mock_workbook):
    mocked_fetch_tab = mocker.patch.object(mock_workbook, 'fetch_tab')
    mocked_fetch_tab().tab_id = '1234'
    mocked_batch_update = mocker.patch.object(mock_workbook, 'batch_update')

    result = mock_workbook.delete_tab('test_delete_tab')
    assert result is None

    mocked_fetch_tab.assert_any_call('test_delete_tab')
    assert mocked_batch_update.call_count == 1
    _, _, kwargs = mocked_batch_update.mock_calls[0]
    assert kwargs['body']['requests'][0] == {'deleteSheet': {'sheetId': '1234'}}


def test_fetch_tab_names(mocker, mock_workbook):
    mocked_sheets_svc = mocker.patch.object(mock_workbook, 'sheets_svc', autospec=True)
    mocked_sheets_svc.get().execute.return_value = {
        'sheets': [
            {'properties': {'title': 'test_tab_1'}},
            {'properties': {'title': 'test_tab_2'}},
            {'properties': {'title': 'test_tab_3'}},
        ]
    }

    expected = pd.DataFrame(['test_tab_1', 'test_tab_2', 'test_tab_3'], columns=['Tabs'])
    assert mock_workbook.fetch_tab_names().equals(expected)


def test_fetch_tab(mocker, mock_workbook):
    tabname = 'test_fetch_tab'
    mocked_sheets_svc = mocker.patch.object(mock_workbook, 'sheets_svc', autospec=True)
    mocked_sheets_svc.get().execute.return_value = {
        'sheets': [{
            'properties': {
                'sheetId': 1504104867,
                'title': 'My Test Tab',
                'index': 3,
                'sheetType': 'GRID',
                'gridProperties': {'rowCount': 15, 'columnCount': 2}
            }
        }]
    }
    tab = mock_workbook.fetch_tab(tabname)
    assert isinstance(tab, datasheets.Tab)
    assert tab.tabname == tabname
    assert tab.workbook.filename == mock_workbook.filename


def test_fetch_tab_not_found(mocker, mock_workbook):
    # Mock response generated by adding a pdb breakpoint in workbook.py to get a real error
    side_effect = apiclient.errors.HttpError(
        resp=httplib2.Response({
            'vary': 'Origin, X-Origin, Referer',
            'content-type': 'application/json; charset=UTF-8',
            'date': 'Thu, 12 Apr 2018 20:16:57 GMT',
            'server': 'ESF',
            'cache-control': 'private',
            'x-xss-protection': '1; mode=block', 'x-frame-options': 'SAMEORIGIN',
            'alt-svc': 'hq=":443"; ma=2592000; quic=51303432; quic=51303431; quic=51303339; quic=51303335,quic=":443"; ma=2592000; v="42,41,39,35"',
            'transfer-encoding': 'chunked',
            'status': '400',
            'content-length': '275',
            '-content-encoding': 'gzip'}),
        content=b'{\n  "error": {\n    "code": 400,\n    "message": "Unable to parse range: foofoo!A1",\n    "errors": [\n      {\n        "message": "Unable to parse range: foofoo!A1",\n        "domain": "global",\n        "reason": "badRequest"\n      }\n    ],\n    "status": "INVALID_ARGUMENT"\n  }\n}\n',
        uri='https://sheets.googleapis.com/v4/spreadsheets/2bJOzX8qQIajUdlaFKfb3HtfyD5Ihy2MOFpR67RwG1SQ?ranges=foofoo%21A1&fields=sheets%2Fproperties&alt=json',
    )
    mocker.patch.object(datasheets.Tab, '_update_tab_properties', autospec=True, side_effect=side_effect)

    with pytest.raises(datasheets.exceptions.TabNotFound) as err:
        mock_workbook.fetch_tab('nonexistent_tab')
    assert err.match('The given tab could not be found. Error generated:')
