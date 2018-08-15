import apiclient
import os
import pandas as pd
import pytest
from google.oauth2.credentials import Credentials as base_credentials
from google.oauth2.service_account import Credentials as service_credentials
from google_auth_oauthlib.flow import InstalledAppFlow

import datasheets


@pytest.fixture
def clear_envvars():
    """
    Remove all environmental variables that have DATASHEETS in them. We do this both to prevent our
    tests from contaminating each other as well as to make sure we don't add cruft to our real env
    from testing. We remove these envvars using a fixture so test failures don't leave junk around.

    Note that because we want to clean up after the test is done we have to initially yield
    """
    yield
    for key in os.environ.keys():
        if 'DATASHEETS' in key:
            del os.environ[key]


def test_init(mock_client):
    # sheets_svc variable properly created
    assert isinstance(mock_client.sheets_svc, apiclient.discovery.Resource)
    assert hasattr(mock_client.sheets_svc, 'sheets') and hasattr(mock_client.sheets_svc, 'values')

    # drive_svc variable properly created
    assert isinstance(mock_client.drive_svc, apiclient.discovery.Resource)
    assert hasattr(mock_client.drive_svc, 'files') and hasattr(mock_client.drive_svc, 'permissions')

    assert mock_client.email == 'test@email.com'

    repr_start = "<datasheets.client.Client(email='test@email.com')>"
    assert repr(mock_client).startswith(repr_start)


def test_getattribute_for_non_method(mock_client):
    mock_client.credentials = 'foo'
    # Also make sure we actually get something back from the non-method call
    assert mock_client.email == 'test@email.com'
    # _refresh_token_if_needs is called in __init__(); verify it wasn't called again
    assert mock_client._refresh_token_if_needed.call_count == 1


def test_getattribute_for_private_method(mock_client):
    mock_client.credentials = 'foo'
    assert mock_client._authenticate()
    # _refresh_token_if_needs is called in __init__(); verify it wasn't called again
    assert mock_client._refresh_token_if_needed.call_count == 1


def test_getattribute_for_user_facing_method_with_credentials(mocker, mock_client):
    mock_client.credentials = 'foo'
    mocker.patch.object(mock_client, 'fetch_workbook', return_value='foo')
    mock_client.fetch_workbook()
    # _refresh_token_if_needs is called in __init__(); verify it was called a second time
    assert mock_client._refresh_token_if_needed.call_count == 2


def test_fetch_file_id_findable_workbook(mocker, mock_client):
    mocked_fetch_info_on_items = mocker.patch.object(
        mock_client, '_fetch_info_on_items', autospec=True, return_value=[
            {'id': 'xyz2345',
             'name': 'datasheets_test',
             'webViewLink': 'https://docs.google.com/spreadsheets/d/xyz2345/edit?usp=drivesdk',
             'modifiedTime': '2018-03-25T03:41:53.640Z'
             },
        ]
    )

    file_id = mock_client._fetch_file_id(filename='datasheets_test', kind='spreadsheet')
    assert file_id == 'xyz2345'
    mocked_fetch_info_on_items.assert_called_once_with(kind='spreadsheet', name='datasheets_test')


def test_fetch_file_id_findable_folder(mocker, mock_client):
    mocked_fetch_info_on_items = mocker.patch.object(
        mock_client, '_fetch_info_on_items', autospec=True, return_value=[
            {'id': 'xyz6789',
             'name': 'datasheets_test_folder',
             'webViewLink': 'https://docs.google.com/spreadsheets/d/xyz6789/edit?usp=drivesdk',
             'modifiedTime': '2018-03-24T02:41:53.640Z'
             },
        ]
    )

    file_id = mock_client._fetch_file_id(filename='datasheets_test_folder', kind='folder')
    assert file_id == 'xyz6789'
    mocked_fetch_info_on_items.assert_called_once_with(kind='folder', name='datasheets_test_folder')


def test_fetch_file_id_workbook_not_found(mocker, mock_client):
    mocked_fetch_info_on_items = mocker.patch.object(
        mock_client, '_fetch_info_on_items', autospec=True, return_value=[]
    )

    with pytest.raises(datasheets.exceptions.WorkbookNotFound) as err:
        mock_client._fetch_file_id(filename='missing_file', kind='spreadsheet')

    mocked_fetch_info_on_items.assert_called_once_with(kind='spreadsheet', name='missing_file')
    err_message = 'Workbook not found. Verify that it is shared with {}'.format(mock_client.email)
    assert err.match(err_message)


def test_fetch_file_id_folder_not_found(mocker, mock_client):
    mocked_fetch_info_on_items = mocker.patch.object(
        mock_client, '_fetch_info_on_items', autospec=True, return_value=[]
    )

    with pytest.raises(datasheets.exceptions.FolderNotFound) as err:
        mock_client._fetch_file_id(filename='missing_folder', kind='folder')

    mocked_fetch_info_on_items.assert_called_once_with(kind='folder', name='missing_folder')
    err_message = 'Folder not found. Verify that it is shared with {}'.format(mock_client.email)
    assert err.match(err_message)


def test_fetch_file_id_with_duplicates(mocker, mock_client):
    mocked_fetch_info_on_items = mocker.patch.object(
        mock_client, '_fetch_info_on_items', autospec=True, return_value=[
            {'id': 'xyz3456',
             'name': 'duplicate_file',
             'webViewLink': 'https://docs.google.com/spreadsheets/d/xyz3456/edit?usp=drivesdk',
             'modifiedTime': '2018-01-14T02:41:53.640Z'
             },
            {'id': 'xyz4567',
             'name': 'duplicate_file',
             'webViewLink': 'https://docs.google.com/spreadsheets/d/xyz4567/edit?usp=drivesdk',
             'modifiedTime': '2018-01-15T08:41:53.640Z'
             },
        ]
    )

    with pytest.raises(datasheets.exceptions.MultipleWorkbooksFound) as err:
        mock_client._fetch_file_id(filename='duplicate_file', kind='spreadsheet')

    mocked_fetch_info_on_items.assert_called_once_with(kind='spreadsheet', name='duplicate_file')

    # Rather than checking the whole message, just check that each component is there
    base_msg = ('Multiple workbooks founds. Please choose the correct file_id below '
                'and provide it to your function instead of a filename:')
    assert err.match(base_msg)
    assert err.match('filename: duplicate_file')
    assert err.match('file_id: xyz3456')
    assert err.match('file_id: xyz4567')
    assert err.match('modifiedTime')
    assert err.match('webViewLink')


def test_fetch_info_on_items(mocker, mock_client):
    mocked_drive_svc = mocker.patch.object(mock_client, 'drive_svc')
    mocked_drive_svc.files().list().execute.return_value = {}

    raw_info = mock_client._fetch_info_on_items(kind='spreadsheet')

    assert raw_info == []
    expected_query = "mimeType='application/vnd.google-apps.spreadsheet'"
    mocked_drive_svc.files().list.assert_called_with(
        fields='nextPageToken, files(name,id,modifiedTime,webViewLink)',
        q=expected_query,
        orderBy='viewedByMeTime desc',
        pageSize=1000,
        pageToken=None,
    )


def test_fetch_info_on_items_with_folder_name_and_only_mine(mocker, mock_client):
    mocked_drive_svc = mocker.patch.object(mock_client, 'drive_svc')
    mocked_drive_svc.files().list().execute.return_value = {}
    mocker.patch.object(mock_client, '_fetch_file_id', autospec=True, return_value='xyz0123')

    # We explicitly test a workbook name with an apostrophe to ensure proper escaping
    raw_info = mock_client._fetch_info_on_items(kind='spreadsheet', folder='my_folder',
                                                name="Test's Workbook", only_mine=True)

    assert raw_info == []
    expected_query = (
        "mimeType='application/vnd.google-apps.spreadsheet'"
        " and 'xyz0123' in parents"
        " and name = 'Test\\'s Workbook'"
        " and '{}' in owners".format(mock_client.email)
    )
    mocked_drive_svc.files().list.assert_called_with(
        fields='nextPageToken, files(name,id,modifiedTime,webViewLink)',
        q=expected_query,
        orderBy='viewedByMeTime desc',
        pageSize=1000,
        pageToken=None,
    )


def test_delete_workbook_which_exists(mocker, mock_client):
    mocked_drive_svc = mocker.patch.object(mock_client, 'drive_svc')
    mocked_fetch_file_id = mocker.patch.object(
        mock_client, '_fetch_file_id', autospec=True, return_value='xyz1234'
    )

    result = mock_client.delete_workbook(filename='testfile')

    mocked_drive_svc.files().delete.assert_called_with(fileId='xyz1234')
    mocked_drive_svc.files().delete().execute.assert_called_once()
    mocked_fetch_file_id.assert_called_once_with(filename='testfile', kind='spreadsheet')
    assert result is None


def test_delete_workbook_error_passing_filename_and_file_id(mock_client):
    with pytest.raises(ValueError) as err:
        mock_client.delete_workbook(filename='foo', file_id='bar')
    assert err.match('Either filename or file_id must be provided, but not both.')


def test_fetch_workbook_error_passing_filename_and_file_id(mock_client):
    with pytest.raises(ValueError) as err:
        mock_client.fetch_workbook(filename='foo', file_id='bar')
    assert err.match('Either filename or file_id must be provided, but not both.')


def test_fetch_workbook(mocker, mock_client):
    mocked_fetch_file_id = mocker.patch.object(
        mock_client, '_fetch_file_id', autospec=True, return_value='xyz1234'
    )

    filename = 'test_fetch_workbook'
    workbook = mock_client.fetch_workbook(filename)

    mocked_fetch_file_id.assert_called_once_with(filename=filename, kind='spreadsheet')
    assert isinstance(workbook, datasheets.Workbook)
    assert workbook.filename == filename
    assert workbook.file_id == 'xyz1234'


def test_fetch_workbook_by_file_id(mock_client):
    workbook = mock_client.fetch_workbook(file_id='xyz1234')
    assert isinstance(workbook, datasheets.Workbook)
    assert workbook.file_id == 'xyz1234'


def test_fetch_workbooks_info_no_folder(mocker, mock_client):
    mocked_fetch_info_on_items = mocker.patch.object(
        mock_client, '_fetch_info_on_items', autospec=True, return_value=[
            {'id': 'xyz1234',
             'name': 'workbook1',
             'webViewLink': 'https://docs.google.com/spreadsheets/d/xyz1234/edit?usp=drivesdk',
             'modifiedTime': '2018-04-07T17:35:16.895Z'},
            {'id': 'xyz2345',
             'name': 'workbook2',
             'webViewLink': 'https://docs.google.com/spreadsheets/d/xyz2345/edit?usp=drivesdk',
             'modifiedTime': '2018-04-06T15:10:04.566Z'},
            {'id': 'xyz3456',
             'name': 'workbook3',
             'webViewLink': 'https://docs.google.com/spreadsheets/d/xyz3456/edit?usp=drivesdk',
             'modifiedTime': '2018-03-23T17:48:52.967Z'},
        ]
    )

    results = mock_client.fetch_workbooks_info()

    mocked_fetch_info_on_items.assert_called_once_with(kind='spreadsheet', folder=None)
    assert isinstance(results, pd.DataFrame)
    assert set(results.columns) == set(['name', 'modifiedTime', 'webViewLink', 'id'])
    assert results.shape == (3, 4)


def test_fetch_workbooks_info_in_folder(mocker, mock_client):
    mocked_fetch_info_on_items = mocker.patch.object(
        mock_client, '_fetch_info_on_items', autospec=True, return_value=[
            {'id': 'xyz1234',
             'name': 'workbook1',
             'webViewLink': 'https://docs.google.com/spreadsheets/d/xyz1234/edit?usp=drivesdk',
             'modifiedTime': '2018-04-07T17:35:16.895Z'},
        ]
    )

    results = mock_client.fetch_workbooks_info(folder='my folder')

    mocked_fetch_info_on_items.assert_called_once_with(kind='spreadsheet', folder='my folder')
    assert isinstance(results, pd.DataFrame)
    assert set(results.columns) == set(['name', 'modifiedTime', 'webViewLink', 'id'])
    assert results.shape == (1, 4)


def test_fetch_folders_all(mocker, mock_client):
    mocker.patch.object(
        mock_client, '_fetch_info_on_items', autospec=True, return_value=[
            {'id': 'xyz5678',
             'name': 'datasheets_test_folder1',
             'webViewLink': 'https://docs.google.com/spreadsheets/d/xyz5678/edit?usp=drivesdk',
             'modifiedTime': '2018-02-13T02:12:53.640Z'
             },
            {'id': 'xyz6789',
             'name': 'datasheets_test_folder2',
             'webViewLink': 'https://docs.google.com/spreadsheets/d/xyz6789/edit?usp=drivesdk',
             'modifiedTime': '2018-03-24T02:41:53.640Z'
             },
        ]
    )

    results = mock_client.fetch_folders()
    assert isinstance(results, pd.DataFrame)
    assert set(results.columns) == set(['name', 'modifiedTime', 'webViewLink', 'id'])
    assert set(results['name']) == set(['datasheets_test_folder2', 'datasheets_test_folder1'])
    assert results.shape == (2, 4)


@pytest.mark.usefixtures('clear_envvars')
def test_get_service_credentials_envvar_set(mocker, tmpdir):
    """
    Only the envvar-based version of running Client()._get_service_credentials()
    is tested as the non-envvar version simply uses a different path
    """
    # Use a non-standard filename and file ending to ensure they work
    file_path = tmpdir.join('my_service_key_file.foo')
    # Credentials were built by taking an existing key and manually smudging it
    file_path.write(r"""
        {
          "type": "service_account",
          "project_id": "datasheets-etl",
          "private_key_id": "199689b78c435a8d2416d166dd3c8f816dbe9837",
          "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCJsDzi0xK3dJza\nsT/sx3Bu+3kXhhpld0BDfQnngU948JjXWlco+svezVXL9fjvaaA5eIhKhZvtAr+2\nJROykKfrER899zPcZxTrUkZQi+T8I5NpufXU68Sx0/mUXpZMx+RZtU12O+YkUqKu\n4694P98vTnY7IkGmp8roAxHo9qe6wn9l6/bNprfuhvp3NAyUK5abcQGHHGJaY7PW\nP129j/fKfXaaB/b7QhdfsbRr4VOMt5eJh2qt8baEJbPbasDfA0h1Q4IMYzPP0UHf\nf9ORsdBklPFpCO56L3rVFd4zbj2JOv2gG63gIPWgAS34FfL1G+s0nN18OrSmg9Yv\nJIjo4WQ7AgMBAAECggEAHt/Jb1VEK6x11cU1wBrij8hIhV8LmRQcDjPq/HqEuoiI\nxTVpOPO1Vpqw2iOxuakkScOfT3A0T088mjI08wEPlmfyYV0MaH4m3ZpFN75+qLkQ\n6ZyuixpbpMJzAv0q3/7sq0321MowK3HqL8vr7TfBpqCWG2Dz0zuQbG3MTqmlaIiR\ndBB0s+qYAJtjXxxRBI7/h1Lkky0DXgw47xI2gdA1COIC+URX49xiamyXcU2M67AY\n1ehdE+98X/I6TpIxm34OFHFLijxi3YAP1Ro0EpDN+xG0CJxbqRh+whdgZ30pu7fQ\nYhtmBaKpaBSALkKsM0nTX/hHtgx1MLQFy/0nGZJ49QKBgQC9YOJVUaUwTObhIimA\nJo0KmquneE/uO0TZf0Abn10aABfKFk5vkJyYqX0vFbm1VA/w7uY579gx3xnFntNL\n3b6WlT87Ffm2PqIPM9v9Lyw2hcRme0AaFalHf0pAYN1civ0Apzd/+z3eRp7JL/8M\n4kppgth1d0rmuGVKKZ1BA8WEVQKBgQC6IDvs2naERYUHRF9tKr0rGKnJSJUxz2UV\naG/Y/xCsSm0pqaAPkY8UQfZp/4iOaXBA5310pGlVG9qBnBAjJl04ivXNphwLMCSe\nInkEVV08eUASR0f0sniP1RR8VSsQSx8aIaxtLyrE0YRQsOSMXYCUK8GdGTbRRrHh\nDv9L45VWTwKBgEFyjyW/PqBvooBx1IjKRl8Lc+W9g9xqlLpY6lNgLnM87jkBl75g\ns1RsSlllrIz+DzhLX12xEbo/MpjgQoskdjbnaxldwudHGYhJetiICfaZRuhHr1jA\nyKzp1+Jwk18bhsAK3L1PsbDmrcYH0VQpjFOw91Jsk58Ih/DWauNCI2u1AoGBAInE\nJDb7uS/Mulsb1kwFngGRPtsjb01BzFEOcW5/1VBdIXDRfwkbh1pwl083v58Rwkn4\nTAbRgza102mnK62MzwF0Md6nLijGk5Ud6Q91hBxk2GnfH8e5r08zvWeOS1LDF+Fi\nvVsQBSBjk82A1881wl9qR+Q8OiFxvO5GCIemi7oXAoGAV/7SZexPjcskjb1jPrHX\nT+cbxVpnoJ7T2m9x5u/53MEKu0JafjMtn84t/2HS5cak5f256bmpFTF9KgvF0daz\nQATAgphZsjbUoMWfTC6Bsaj1kbqQC41mhdzMjz23mIYQbrpxlB11rVu+u6NmMNIB\nzgL2IXskmK19pKzia4FlKVw=\n-----END PRIVATE KEY-----\n",
          "client_email": "datasheets-service@datasheets-etl.iam.gserviceaccount.com",
          "client_id": "109046144667258184477",
          "auth_uri": "https://accounts.google.com/o/oauth2/auth",
          "token_uri": "https://accounts.google.com/o/oauth2/token",
          "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
          "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/datasheets-service%40datasheets-etl.iam.gserviceaccount.com"
        }
        """)
    os.environ['DATASHEETS_SERVICE_PATH'] = file_path.strpath

    mocker.patch.object(datasheets.Client, '__init__', return_value=None)
    client = datasheets.Client()
    credentials = client._get_service_credentials()
    assert isinstance(credentials, service_credentials)
    assert client.email == 'datasheets-service@datasheets-etl.iam.gserviceaccount.com'


@pytest.mark.usefixtures('clear_envvars')
def test_fetch_new_client_credentials_envvar_set(tmpdir):
    # Use a non-standard filename and file ending to ensure they work
    file_path = tmpdir.join('my_client_secrets_file.foo')
    # Credentials were built by taking an existing secrets file and manually smudging it
    file_path.write(r"""
        {
            "installed": {
                "client_id":"562803761647-1lj6fdt4rk27qde3f61slphbqcr9mieh.apps.googleusercontent.com",
                "project_id":"gsheets-etl",
                "auth_uri":"https://accounts.google.com/o/oauth2/auth",
                "token_uri":"https://www.googleapis.com/oauth2/v3/token",
                "auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs",
                "client_secret":"yMWIX9SijX-nUgvFGqkzoSBb",
                "redirect_uris":["urn:ietf:wg:oauth:2.0:oob","http://localhost"]}}
    """)
    os.environ['DATASHEETS_SECRETS_PATH'] = file_path.strpath

    flow = InstalledAppFlow.from_client_secrets_file(os.environ['DATASHEETS_SECRETS_PATH'] , [])
    config = flow.client_config
    assert isinstance(flow, InstalledAppFlow)
    assert config["client_id"] == '562803761647-1lj6fdt4rk27qde3f61slphbqcr9mieh.apps.googleusercontent.com'


@pytest.mark.usefixtures('clear_envvars')
def test_retrieve_client_credentials_use_storage_and_envvar_set(mocker, tmpdir):
    # Use a non-standard filename and file ending to ensure they work
    file_path = tmpdir.join('my_client_credentials_file.foo')
    # Credentials were built by taking an existing credentials file and manually smudging it
    file_path.write(r"""{
        "access_token": "ya29.GlycBQBd0i9bxu2F1DZ4kPhk4ahwcayAVNEzo1aLFcLVIRFevIJXCvG7WtKDT7jX3nnSTdI69nprY6W27AfEgBHlDRKOGI1VkyDgtV8OidAP5wutTduMoVd8pqrpmw",
        "client_id": "561903281647-9kt18bal218sblb1b4b76uj0b5vq7e0o.apps.googleusercontent.com",
        "client_secret": "acGIZrpwf18djbk1EUVyzpjq",
        "refresh_token": null,
        "token_expiry": "2018-04-13T05:47:47Z",
        "token_uri": "https://accounts.google.com/o/oauth2/token",
        "user_agent": "Python datasheets library",
        "revoke_uri": "https://accounts.google.com/o/oauth2/revoke",
        "id_token": {
            "azp": "561903281647-9kt18bal218sblb1b4b76uj0b5vq7e0o.apps.googleusercontent.com",
            "aud": "561903281647-9kt18bal218sblb1b4b76uj0b5vq7e0o.apps.googleusercontent.com",
            "sub": "101102810221472303039",
            "hd": "squarespace.com",
            "email": "datasheets@test.com",
            "email_verified": true,
            "at_hash": "PZQvzsWx-wOvZGgCAQaJeQ",
            "exp": 1523599737,
            "iss": "accounts.google.com",
            "iat": 1523494867},
        "id_token_jwt": "eyJhbGciOiJSUz19283jbj1kdk1jbN1dNTQ3ODg2ZmY4NWEzNDI4ZGY0ZjYxZGI3M2MxYzIzOTgyYTkyOGUifQ.eyJhenAsdkLS18dBpwQ28BjsNDctOWt0MGNmdXZiOGgwOWxiMWI0Yjc2dWowYjV2cTdlMG8uYXBwcy5nb29nbGV1c2VyY29udGVudC5jb20iLCsp1PRlbj4KNjI4MDM3NjE2NDctOWt0MGNmdXZiOGgwOWxiMWI0Yjc2dWowYjV2cTdlMG8uYXBwcy5nb29nbGV1c2VyY29udGVudC5jb20iLCJzdWIiOiIxMDEzNzgyMzQzODczMjMzMDMwMzkiLCJoZCI6InNxdWFyZXNwYWNlLmNvbSIsImVtYWlsIjoiem1hcmluZUBzcXVhcmVzcGFjZS5jb20iLCJlbWFpbF92ZXJpZ18GPfp1jbL1ZSwiYXRfaGFzaCI6IlBPWnN2eld4LXdPWnZnR0FDcUFKZVEiLCJleHAiOjE1MjM1OTg0NjcsImlzcyI6ImFjY291bnRzLmdvb2dsZS5jb20iLCJpYXQiOjE1MjM1OTQ4Njd9.VP-LSRBnBx87YcvWi5kV1SEZlg3AKky7o_qIBo8Q9KT7nPwihBDdE0uBk5GraKFwGIKu-Xx95AisUEJdnWnJQZZg-RXyINCVHiEzutskPL3jBKlL0EJWnre2IISJxmIqrz6yAJcQD-buWTk1J7zf4Sbhk7EzVvpI1kQJO_pSWgRCdglgFQXJ4ozdBmIQbd76WUXA8-juElea9NkRjCKW8t_dXKvbj-1okR-YOczgmYAoQOfnJ19jplGK7qrQ9sP06ALon993yhbW4Ah37wMEEX3EcHoxcjciH6Z_373ZyVyjf2ZHOZKgqkHZqzefUteEMMdG3phiNd0h6ro12DrGMw",
        "token_response": {
            "access_token": "ya29.GlycBQBd0i9bxu2F1DZ4kPpwl284J3klLNEzo1aLFcLVIRFevIJXCvG7WtKDT7jX3npwlbHJknprY6W27AfEgBHlDRKOGI1VkyDgtV8OidAP5wutTduMoVd8pqrpmw",
        "expires_in": 3600,
        "id_token": "eyJhbGciOiJSUzILPL1jk294MvI6IjNiNTQ3ODg2ZmY4NWEzNDI4ZGY0ZjYxZGI3M2MxYzIzOTgyYTkyOGUifQ.eyJhenAiOiI1NjI4MDM3NjE2NDctOWt0MGNmdXZiOGgwOWxiMWI0Yjc2dWowYjV2cTdlMG8uYXBwcy5nb29nbGV1c2VyY29udGVudC5jb20iLCJhdWQiOiI1NjI4MDM3NjE2NDctOWt0MGNmdXZiOkshb1KLvWI0Yjc2dWowYjV2cTdlMG8uYXBwcy5nb29nbGV1c2VyY29udGplwMN4820iLCJzdWIiOiIxMDEzNzgyMzQzODczMjMzMDMwMzkiLCJoZCI6InNxdWFyZXNwYWNlLmNvbSIsImVtYWlsk1MNl83hcmluZUBzcXVhcmVzcGFjZS5jb20iLCJlbWFpbF92ZXJpZmllZCI6dHJ1ZSwiYXRfaGFzaCI6IlBPWnN2eld4LXdPWnZnR0FDcUFKZVEiLCJleHAiOjE1MjM1OTg0NjcsImlzcyI6ImFjY291bnRzLmdvb2dsZS5jb20iLCJpYXQiOjE1MjM1OTQ4Njd9.VP-LSRBnBx87YcvWi5kV1SEZlg3AKky7o_qIBplKm28FbmsihBDdE0uBk5GraKFwGIKu-Xx95AisUEJdnWnJQZZg-RXyINCVHiEzutuBg2kKl84pmEJWnre2IISJxmIqrz6yAJcQD-buWTk1J7zf4Sbhk7EzVvpI1kQJO_pSWgRCdglgFQXJ4ozdBmIQbd76WUXA8-juElea9NkRjCKW8t_dXKvbj-1okR-YOczgmYkmpvinJ8HTSNFl7qrQ9sP06ALon993yhbW4Ah37wMEEX3EcHoxcjciH6Z_373ZyVyjf2ZHbj1kKplZqzefUteEMMdG3phiNd0h6ro12DrGMw",
        "token_type": "Bearer"},
        "scopes": ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/userinfo.email"],
        "token_info_uri": "https://www.googleapis.com/oauth2/v3/tokeninfo",
        "invalid": false,
        "_class": "OAuth2Credentials",
        "_module": "oauth2client.client"
    }""")
    os.environ['DATASHEETS_CREDENTIALS_PATH'] = file_path.strpath

    mocker.patch.object(datasheets.Client, '__init__', return_value=None)
    mocked_fetch_new = mocker.patch.object(datasheets.Client, '_fetch_new_client_credentials',
                                           return_value='test_return', autospec=True)
    client = datasheets.Client()
    client.use_storage = True
    credentials = client._retrieve_client_credentials()
    assert mocked_fetch_new.call_count == 0
    assert isinstance(credentials, base_credentials)


def test_retrieve_client_credentials_no_storage(mocker):
    mocker.patch.object(datasheets.Client, '__init__', return_value=None)
    client = datasheets.Client()
    client.use_storage = False
    mocked_fetch_new = mocker.patch.object(client, '_fetch_new_client_credentials')
    credentials = client._retrieve_client_credentials()
    assert credentials == mocked_fetch_new()


def test_stores_credentials_when_not_found(mocker):
    credentials = base_credentials("token", refresh_token="refresh_token", client_id="client_id",
                                   client_secret="client_secret")

    os.environ['DATASHEETS_CREDENTIALS_PATH'] = "./test_stores_credentials_when_not_found.json"
    mocker.patch.object(datasheets.Client, '__init__', return_value=None)
    mocker.patch.object(datasheets.Client, '_fetch_new_client_credentials',
                        return_value=credentials, autospec=True)
    client = datasheets.Client()
    client.use_storage = True
    client.email = "Test"
    client._retrieve_client_credentials()
    with open(os.environ['DATASHEETS_CREDENTIALS_PATH']) as file:
        assert file.read() == '{"refresh_token": "refresh_token", "client_id": "client_id", "client_secret": "client_secret"}'
    os.remove(os.environ['DATASHEETS_CREDENTIALS_PATH'])


def test_create_workbook_no_folder(mocker, mock_client):
    filename = 'test_create_workbook'
    file_id = 'xyz1234'
    root_id = '0AP2cy554S5hyUk9PVA'
    mocked_drive_svc = mocker.patch.object(mock_client, 'drive_svc', autospec=True)
    mocked_drive_svc.files().get().execute.return_value = {'id': root_id}
    mocker.patch.object(mock_client, '_fetch_file_id', autospec=True, return_value=file_id)

    workbook = mock_client.create_workbook(filename)

    mocked_drive_svc.files().get.assert_called_with(fileId='root', fields='id')
    mocked_drive_svc.files().get().execute.assert_called_once()

    mocked_drive_svc.files().create.assert_called_once()
    _, _, kwargs = mocked_drive_svc.files().create.mock_calls[0]
    assert kwargs['body']['name'] == filename
    assert kwargs['body']['parents'] == [root_id]

    assert isinstance(workbook, datasheets.Workbook)


def test_create_workbook_with_folder(mocker, mock_client):
    filename = 'test_create_workbook'
    file_id = 'xyz1234'
    foldername = 'datasheets_test_folder_1'
    folder_id = 'xyz1234'
    root_id = '0AP2cy554S5hyUk9PVA'
    mocked_drive_svc = mocker.patch.object(mock_client, 'drive_svc', autospec=True)
    mocked_drive_svc.files().get().execute.return_value = {'id': root_id}
    mocked_fetch_file_id = mocker.patch.object(
        mock_client, '_fetch_file_id', autospec=True, side_effect=[folder_id, file_id]
    )

    workbook = mock_client.create_workbook(filename, folders=(foldername,))

    mocked_drive_svc.files().get.assert_called_with(fileId='root', fields='id')
    mocked_drive_svc.files().get().execute.assert_called_once()

    assert mocked_fetch_file_id.call_count == 2
    mocked_fetch_file_id.assert_any_call(filename=foldername, kind='folder')
    mocked_fetch_file_id.assert_any_call(filename=filename, kind='spreadsheet')

    mocked_drive_svc.files().create.assert_called_once()
    _, _, kwargs = mocked_drive_svc.files().create.mock_calls[0]
    assert kwargs['body']['name'] == filename
    assert set(kwargs['body']['parents']) == set([root_id, folder_id])

    assert isinstance(workbook, datasheets.Workbook)


def test_create_workbook_multiple_folders(mocker, mock_client):
    filename = 'test_create_workbook'
    file_id = 'xyz1234'
    foldernames = ('datasheets_test_folder_1', 'datasheets_test_folder_2')
    folder_ids = ['abc1234', 'efg2345']
    root_id = '0AP2cy554S5hyUk9PVA'
    mocked_drive_svc = mocker.patch.object(mock_client, 'drive_svc', autospec=True)
    mocked_drive_svc.files().get().execute.return_value = {'id': root_id}
    mocked_fetch_file_id = mocker.patch.object(
        mock_client, '_fetch_file_id', autospec=True, side_effect=folder_ids + [file_id]
    )

    workbook = mock_client.create_workbook(filename, folders=foldernames)

    mocked_drive_svc.files().get.assert_called_with(fileId='root', fields='id')
    mocked_drive_svc.files().get().execute.assert_called_once()

    mocked_fetch_file_id.assert_any_call(filename=foldernames[0], kind='folder')
    mocked_fetch_file_id.assert_any_call(filename=foldernames[1], kind='folder')
    mocked_fetch_file_id.assert_any_call(filename=filename, kind='spreadsheet')

    mocked_drive_svc.files().create.assert_called_once()
    _, _, kwargs = mocked_drive_svc.files().create.mock_calls[0]
    assert kwargs['body']['name'] == filename
    assert set(kwargs['body']['parents']) == set([root_id, folder_ids[0], folder_ids[1]])

    assert isinstance(workbook, datasheets.Workbook)
    assert workbook.file_id == file_id
