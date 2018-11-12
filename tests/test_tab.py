import apiclient
import httplib2
import pandas as pd
import pytest
from conftest import get_data_from_yaml

import datasheets


def test_init(mock_tab):
    assert isinstance(mock_tab, datasheets.Tab)
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
    assert mock_tab.tab_id == 7704104867
    assert mock_tab.url == 'https://docs.google.com/spreadsheets/d/xyz1234#gid=7704104867'

    repr_start = "<datasheets.tab.Tab(filename='datasheets_test_1', tabname='test_tab')>"
    assert repr(mock_tab).startswith(repr_start)


def test_getattribute_for_non_method(mock_tab):
    # Also make sure we actually get something back from the non-method call
    assert mock_tab.tabname == 'test_tab'
    # There will be 2 calls already because mock_tab is created from mock_workbook.fetch_tab()
    # and _refresh_token_if_needs is called in __init__(); verify it wasn't called again
    assert mock_tab.workbook.client._refresh_token_if_needed.call_count == 2


def test_getattribute_for_private_method(mocker, mock_tab):
    mocker.patch.object(mock_tab, '_process_rows')
    mock_tab._process_rows()
    # There will be 2 calls already because mock_tab is created from mock_workbook.fetch_tab()
    # and _refresh_token_if_needs is called in __init__(); verify it wasn't called again
    assert mock_tab.workbook.client._refresh_token_if_needed.call_count == 2


def test_getattribute_for_user_facing_method_with_credentials(mocker, mock_tab):
    mocker.patch.object(mock_tab, 'add_rows')
    mock_tab.add_rows()
    # There will be 2 calls already because mock_tab is created from mock_workbook.fetch_tab()
    # and _refresh_token_if_needs is called in __init__(); verify it was called a third time
    assert mock_tab.workbook.client._refresh_token_if_needed.call_count == 3


def test_fetch_tab_not_found(mocker):
    # Exception built by pasting content from a real exception generated
    # via Ipython + a pdb trace put in the datasheets code
    exception = apiclient.errors.HttpError(
        resp=httplib2.Response({
            'vary': 'Origin, X-Origin, Referer',
            'content-type': 'application/json; charset=UTF-8',
            'date': 'Tue, 10 Apr 2018 19:17:12 GMT',
            'server': 'ESF',
            'cache-control': 'private',
            'x-xss-protection': '1; mode=block',
            'x-frame-options': 'SAMEORIGIN',
            'alt-svc': 'hq=":443"; ma=2592000; quic=51303432; quic=51303431; quic=51303339; quic=51303335,quic=":443"; ma=2592000; v="42,41,39,35"',
            'transfer-encoding': 'chunked',
            'status': '400',
            'content-length': '271',
            '-content-encoding': 'gzip'
        }),
        content=b"""
            {
                "error": {
                    "code": 400,
                    "message": "Unable to parse range: flib!A1",
                    "errors": [
                        {
                            "message": "Unable to parse range: flib!A1",
                            "domain": "global",
                            "reason": "badRequest"
                        }
                    ],
                    status": "INVALID_ARGUMENT"
                }
            }
        """,
        uri='https://sheets.googleapis.com/v4/spreadsheets/1bKOzXaaaaaaaaaaaaaa2FftyD5Ihy2MOFqR67rWG0SQ?ranges=flib%21A1&fields=sheets%2Fproperties&alt=json'
    )
    mocker.patch('datasheets.Tab._update_tab_properties', side_effect=exception)

    DUMMY = '_'
    with pytest.raises(datasheets.exceptions.TabNotFound) as err:
        datasheets.Tab('test_fetch_tab', DUMMY, DUMMY, DUMMY)

    assert err.match('The given tab could not be found. Error generated: ')


def test_add_rows_or_columns(mocker, mock_tab):
    mocked_batch_update = mocker.patch.object(mock_tab.workbook, 'batch_update', autospec=True)
    mocked_update_tab_properties = mocker.patch.object(mock_tab, '_update_tab_properties',
                                                       autospec=True)

    mock_tab._add_rows_or_columns(kind='ROWS', n='1234')

    assert mocked_batch_update.call_count == 1
    _, call_args, _ = mocked_batch_update.mock_calls[0]
    requests = call_args[0]['requests']
    assert len(requests) == 1
    assert requests[0]['appendDimension'] == {'length': '1234',
                                              'sheetId': mock_tab.tab_id,
                                              'dimension': 'ROWS'}
    mocked_update_tab_properties.assert_called_once_with()


def test_align_cells(mocker, mock_tab):
    mocked_batch_update = mocker.patch.object(mock_tab.workbook, 'batch_update', autospec=True)

    mock_tab.align_cells()

    assert mocked_batch_update.call_count == 1
    _, call_args, _ = mocked_batch_update.mock_calls[0]
    assert len(call_args[0]['requests']) == 1
    assert call_args[0]['requests'][0]['repeatCell']['range']['sheetId'] == mock_tab.tab_id


def test_alter_dimensions(mocker, mock_tab):
    mocked_batch_update = mocker.patch.object(mock_tab.workbook, 'batch_update', autospec=True)
    mocked_update_tab_properties = mocker.patch.object(mock_tab, '_update_tab_properties',
                                                       autospec=True)

    mock_tab.alter_dimensions(nrows=123)

    assert mocked_batch_update.call_count == 1
    _, call_args, _ = mocked_batch_update.mock_calls[0]
    requests = call_args[0]['requests']
    assert len(requests) == 1
    properties = requests[0]['updateSheetProperties']['properties']
    assert properties['sheetId'] == mock_tab.tab_id
    assert properties['gridProperties']['columnCount'] == mock_tab.ncols
    assert properties['gridProperties']['rowCount'] == 123
    mocked_update_tab_properties.assert_called_once_with()


def test_autosize_columns(mocker, mock_tab):
    mocked_batch_update = mocker.patch.object(mock_tab.workbook, 'batch_update', autospec=True)

    mock_tab.autosize_columns()

    assert mocked_batch_update.call_count == 1
    _, call_args, _ = mocked_batch_update.mock_calls[0]
    requests = call_args[0]['requests']
    assert len(requests) == 1
    assert requests[0]['autoResizeDimensions']['dimensions'] == {
        'sheetId': mock_tab.tab_id,
        'dimension': 'COLUMNS',
        'startIndex': 0,
        'endIndex': mock_tab.ncols
    }


def test_format_font(mocker, mock_tab):
    mocked_batch_update = mocker.patch.object(mock_tab.workbook, 'batch_update', autospec=True)

    mock_tab.format_font()

    assert mocked_batch_update.call_count == 1
    _, call_args, _ = mocked_batch_update.mock_calls[0]
    requests = call_args[0]['requests']
    assert len(requests) == 1
    assert requests[0]['repeatCell']['range']['sheetId'] == mock_tab.tab_id
    assert requests[0]['repeatCell']['cell']['userEnteredFormat']['textFormat'] == {
        'fontSize': 10,
        'fontFamily': 'Proxima Nova',
    }


def test_format_headers(mocker, mock_tab):
    mocked_batch_update = mocker.patch.object(mock_tab.workbook, 'batch_update', autospec=True)

    mock_tab.format_headers(nrows=3)

    assert mocked_batch_update.call_count == 1
    _, call_args, _ = mocked_batch_update.mock_calls[0]
    requests = call_args[0]['requests']
    assert len(requests) == 2
    repeat_cell, update_sheet = requests
    assert repeat_cell['repeatCell']['range']['endRowIndex'] == 3
    assert repeat_cell['repeatCell']['range']['sheetId'] == mock_tab.tab_id
    properties = update_sheet['updateSheetProperties']['properties']
    assert properties['gridProperties']['frozenRowCount'] == 3
    assert properties['sheetId'] == mock_tab.tab_id


def test_clear_data(mocker, mock_tab):
    mock_tab.clear_data()
    mock_tab.sheets_svc.values().clear.assert_called_with(spreadsheetId=mock_tab.workbook.file_id,
                                                          range=mock_tab.tabname, body={})
    mock_tab.sheets_svc.values().clear().execute.assert_called_once_with()


def test_get_parent_workbook(mock_tab):
    parent = mock_tab.workbook
    assert parent == mock_tab._workbook


def test_process_rows_empty_tab(mock_tab):
    """ Verify the raw_data an empty tab would return is processed correctly """
    raw_data = {u'sheets': [{u'data': [{}]}]}
    assert [] == mock_tab._process_rows(raw_data)


def test_process_rows(mock_tab, expected_data):
    data = get_data_from_yaml('test_process_rows.yaml')
    assert mock_tab._process_rows(data) == expected_data


def test_process_rows_type_errors(mock_tab, expected_data):
    data = get_data_from_yaml('test_process_rows_type_error.yaml')
    with pytest.raises(TypeError) as err:
        mock_tab._process_rows(data)
    assert err.match('Mismatch exists in expected and actual data types')


def test_process_rows_cell_error_values(mock_tab):
    data = {
        'sheets': [
            {'data': [
                {'rowData': [
                    {'values': [{'effectiveValue': {'numberValue': 2}}]},
                    {'values': [{}, {'effectiveValue': {'stringValue': 'foo'}}]},
                    {'values': [{}, {
                        'effectiveValue': {
                            'errorValue': {
                                'type': 'NAME',
                                'message': "Unknown range name: 'FOO'.",
                            }
                        }
                    }]}
                ]}
            ]}
        ]
    }
    with pytest.raises(datasheets.exceptions.FetchDataError) as err:
        mock_tab._process_rows(data)
    assert err.match(
        'Error of type "NAME" within cell B3 prevents fetching data. Message: "Unknown range name: \'FOO\'."'
    )


def test_fetch_data_empty(mock_tab):
    mock_tab.sheets_svc.get().execute.return_value = {'sheets': [{'data': [{}]}]}

    assert mock_tab.fetch_data().equals(pd.DataFrame([]))
    assert mock_tab.fetch_data(fmt='dict') == []
    assert mock_tab.fetch_data(fmt='list') == ([], [])
    assert mock_tab.fetch_data(headers=False).equals(pd.DataFrame([]))
    assert mock_tab.fetch_data(headers=False, fmt='dict') == []
    assert mock_tab.fetch_data(headers=False, fmt='list') == ([], [])


def test_fetch_data_unexpected_fmt(mock_tab):
    with pytest.raises(ValueError) as err:
        mock_tab.fetch_data(fmt='foo')
    assert err.match("Unexpected value 'foo' for parameter `fmt`. "
                     "Accepted values are 'df', 'dict', and 'list'")


def test_fetch_data_df_with_headers(mock_tab, expected_cleaned_data_with_headers):
    data = get_data_from_yaml('test_fetch_data.yaml')
    mock_tab.sheets_svc.get().execute.return_value = data

    expected = pd.DataFrame(data=expected_cleaned_data_with_headers[1],
                            columns=expected_cleaned_data_with_headers[0])
    assert mock_tab.fetch_data().equals(expected)


def test_fetch_data_df_no_headers(mock_tab, expected_cleaned_data_no_headers):
    data = get_data_from_yaml('test_fetch_data.yaml')
    mock_tab.sheets_svc.get().execute.return_value = data

    expected = pd.DataFrame(data=expected_cleaned_data_no_headers[1],
                            columns=expected_cleaned_data_no_headers[0])
    assert mock_tab.fetch_data(headers=False).equals(expected)


def test_fetch_data_list_with_headers(mock_tab, expected_cleaned_data_with_headers):
    data = get_data_from_yaml('test_fetch_data.yaml')
    mock_tab.sheets_svc.get().execute.return_value = data

    assert mock_tab.fetch_data(fmt='list') == expected_cleaned_data_with_headers


def test_fetch_data_list_no_headers(mock_tab, expected_cleaned_data_no_headers):
    data = get_data_from_yaml('test_fetch_data.yaml')
    mock_tab.sheets_svc.get().execute.return_value = data

    assert mock_tab.fetch_data(fmt='list', headers=False) == expected_cleaned_data_no_headers


def test_fetch_data_dict_with_headers(mock_tab, expected_cleaned_data_with_headers):
    data = get_data_from_yaml('test_fetch_data.yaml')
    mock_tab.sheets_svc.get().execute.return_value = data

    keys = expected_cleaned_data_with_headers[0]
    values = expected_cleaned_data_with_headers[1]
    expected = [dict(zip(keys, row)) for row in values]
    assert mock_tab.fetch_data(fmt='dict') == expected


def test_fetch_data_dict_no_headers(mock_tab, expected_cleaned_data_no_headers):
    data = get_data_from_yaml('test_fetch_data.yaml')
    mock_tab.sheets_svc.get().execute.return_value = data

    keys = expected_cleaned_data_no_headers[0]
    values = expected_cleaned_data_no_headers[1]
    expected = [dict(zip(keys, row)) for row in values]
    assert mock_tab.fetch_data(fmt='dict', headers=False) == expected


def test_insert_data(mocker, mock_tab, expected_data):
    mocker.patch.object(mock_tab, 'clear_data', autospec=True)
    mocked_update_tab_properties = mocker.patch.object(mock_tab, '_update_tab_properties',
                                                       autospec=True)
    transformed_data = [
        ['a', 1.23, 'this is a sentence.'],
        [3.23, 7.0, '2016-01-01', None],
        [],
        ['2010-08-07 16:13:00', True, 0.19, 0.54, None],
        [None, '10:03:00', 3.23, None, None],
        [None, None, None, None, None]
    ]

    mock_tab.insert_data(data=expected_data, autoformat=False)

    mock_tab.sheets_svc.values().update.assert_called_with(
        spreadsheetId=mock_tab.workbook.file_id,
        range=mock_tab.tabname,
        valueInputOption='USER_ENTERED',
        body={'values': transformed_data})

    mocked_update_tab_properties.assert_called_once_with()


def test_append_data(mocker, mock_tab, expected_data):
    mocker.patch.object(mock_tab, 'clear_data', autospec=True)
    mocked_update_tab_properties = mocker.patch.object(mock_tab, '_update_tab_properties',
                                                       autospec=True)
    transformed_data = [
        ['a', 1.23, 'this is a sentence.'],
        [3.23, 7.0, '2016-01-01', None],
        [],
        ['2010-08-07 16:13:00', True, 0.19, 0.54, None],
        [None, '10:03:00', 3.23, None, None],
        [None, None, None, None, None]
    ]

    mock_tab.append_data(data=expected_data, autoformat=False)

    mock_tab.sheets_svc.values().append.assert_called_with(
        spreadsheetId=mock_tab.workbook.file_id,
        range=mock_tab.tabname,
        valueInputOption='USER_ENTERED',
        body={'values': transformed_data})

    mocked_update_tab_properties.assert_called_once_with()
