'''
Create mocked Google Drive and Google Sheets APIs to permit testing gsheets without
requiring network access

Discovery docs were generated using the following:

    import apiclient
    from httplib2 import Http

    # For Google Sheets, use /apis/sheets/v4/rest below
    request_uri = 'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest'
    discovery_doc = apiclient.discovery._retrieve_discovery_doc(url=request_uri, http=Http(),
                                                                cache_discovery=False)
'''
import datetime as dt
import httplib2
import json
import os

import apiclient
import pytest
import yaml

import gsheets as sheets


@pytest.fixture
def build_path(request):
    return os.path.join(os.path.dirname(__file__), 'mock_responses', request.param)


@pytest.fixture
def get_raw_resource(request):
    filepath = build_path(request)
    with open(filepath, 'r') as f:
        return yaml.load(f)


def process_resource(request):
    '''
    apiclient.http.RequestMockBuilder expects the 2nd and (if it exists) 3rd entries
    to be in json format. Convert those entries.
    '''
    raw_resource = get_raw_resource(request)

    if not raw_resource:
        return

    processed = {}
    for k, v in raw_resource.items():
        first, to_convert = v[0], v[1:]
        if first:
            first = httplib2.Response(first)
        jsoned_items = list(map(json.dumps, to_convert))
        processed[k] = [first] + jsoned_items
    return processed


@pytest.fixture
def drive_http():
    discovery_doc = os.path.join(os.path.dirname(__file__), 'drive_discovery.json')
    return apiclient.http.HttpMock(discovery_doc)


@pytest.fixture
def sheets_http():
    discovery_doc = os.path.join(os.path.dirname(__file__), 'sheets_discovery.json')
    return apiclient.http.HttpMock(discovery_doc)


@pytest.fixture
def request_builder(request):
    processed = process_resource(request)
    return apiclient.http.RequestMockBuilder(processed, check_unexpected=True)


@pytest.fixture
def mock_client(drive_http, sheets_http, request):
    '''
    The `request` parameter contains indirect parameters passed by the test
    function that uses mock_client
    '''
    testing_config = {'drive_http': drive_http,
                      'sheets_http': sheets_http,
                      'request_builder': request_builder(request)
                      }
    return sheets.Client(testing_config=testing_config)


@pytest.fixture
def mock_workbook(mock_client):
    return mock_client.get_workbook('gsheets_test_1')


@pytest.fixture
def mock_tab(mock_workbook):
    return mock_workbook.get_tab('test_tab')


@pytest.fixture
def expected_data():
    ''' A data set with:
            - multiple date types: strings, floats, dates, datetimes, times
            - an empty row in the middle
            - empty trailing row
            - empty trailing column
            - 1 row with 4 populated cells vs. header and other rows having 3 populated cells
    '''
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
