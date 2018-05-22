"""
drive_svc and sheets_svc are copies of the Google Drive and Google Sheets discovery resources which
define the high-level schema for each service: which endpoints exist, what fields they take, and how
those fields are expected to be configured. These are direct copies of the actual discovery
resources pulled from the Google API in November 2016 using code listed below. These resources are
mocked using `apiclient.http.HttpMock` within the fixtures `drive_http` and `sheets_http`, which are
used in the `mock_client` fixture.  These discovery resources are unlikely to require modifications
until a new version of of the APIs comes out and datasheets switches to it.

The discovery resources were generated using the following:

    import apiclient
    from httplib2 import Http

    # For Google Sheets, use /apis/sheets/v4/rest below
    request_uri = 'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest'
    discovery_doc = apiclient.discovery._retrieve_discovery_doc(url=request_uri, http=Http(),
                                                                cache_discovery=False)
"""
import datetime as dt
import os

import apiclient
import pytest
import yaml

import datasheets


def build_path(path):
    return os.path.join(os.path.dirname(__file__), 'resources', path)


def get_data_from_yaml(path):
    filepath = build_path(path)
    with open(filepath, 'r') as f:
        return yaml.load(f)


@pytest.fixture(scope='session')
def drive_svc():
    request_builder = apiclient.http.RequestMockBuilder(None, check_unexpected=True)
    discovery_doc = os.path.join(os.path.dirname(__file__), 'resources/drive_discovery.json')
    drive_http = apiclient.http.HttpMock(discovery_doc)
    return apiclient.discovery.build('drive', 'v3', http=drive_http, requestBuilder=request_builder)


@pytest.fixture(scope='session')
def sheets_svc():
    request_builder = apiclient.http.RequestMockBuilder(None, check_unexpected=True)
    discovery_doc = os.path.join(os.path.dirname(__file__), 'resources/sheets_discovery.json')
    sheets_http = apiclient.http.HttpMock(discovery_doc)
    return apiclient.discovery.build('sheets', 'v4', http=sheets_http,
                                     requestBuilder=request_builder)


@pytest.fixture
def mock_client(mocker, drive_svc, sheets_svc):
    mocker.patch('datasheets.Client._authenticate')
    mocker.patch('datasheets.Client.http', create=True)
    mocker.patch('apiclient.discovery.build', autospec=True,
                 side_effect=[drive_svc, sheets_svc])
    mocker.patch('datasheets.Client._refresh_token_if_needed')

    client = datasheets.Client()
    client.email = 'test@email.com'
    return client


@pytest.fixture
def mock_workbook(mock_client, drive_svc, sheets_svc):
    return datasheets.Workbook(filename='datasheets_test_1', file_id='xyz1234', client=mock_client,
                               drive_svc=drive_svc, sheets_svc=sheets_svc.spreadsheets())


@pytest.fixture
def mock_tab(mocker, mock_workbook):
    mocked_sheets_svc = mocker.patch.object(mock_workbook, 'sheets_svc', autospec=True)
    mocked_sheets_svc.get().execute.return_value = {
        'sheets': [{
            'properties': {
                'sheetId': 7704104867,
                'title': 'test_tab',
                'index': 3,
                'sheetType': 'GRID',
                'gridProperties': {'rowCount': 1000, 'columnCount': 26}
            }
        }]
    }
    return mock_workbook.fetch_tab('test_tab')


@pytest.fixture
def expected_data():
    """
    A data set with:
        - multiple date types: strings, floats, dates, datetimes, times
        - an empty row in the middle
        - empty trailing row
        - empty trailing column
        - 1 row with 4 populated cells vs. header and other rows having 3 populated cells
    """
    return [['a', 1.23, 'this is a sentence.'],
            [3.23, 7.0, dt.date(2016, 1, 1), None],
            [],
            [dt.datetime(2010, 8, 7, 16, 13), True, 0.19, 0.54, None],
            [None, dt.time(10, 3), 3.23, None, None],
            [None, None, None, None, None]]


@pytest.fixture
def expected_cleaned_data_no_headers():
    return ([0, 1, 2, 3],
            [['a', 1.23, 'this is a sentence.', None],
            [3.23, 7.0, dt.date(2016, 1, 1), None],
            [None, None, None, None],
            [dt.datetime(2010, 8, 7, 16, 13), True, 0.19, 0.54],
            [None, dt.time(10, 3), 3.23, None]])


@pytest.fixture
def expected_cleaned_data_with_headers():
    return (['a', 1.23, 'this is a sentence.'],
            [[3.23, 7.0, dt.date(2016, 1, 1)],
            [None, None, None],
            [dt.datetime(2010, 8, 7, 16, 13), True, 0.19],
            [None, dt.time(10, 3), 3.23]])
