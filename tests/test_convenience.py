import datasheets


def test_create_tab_in_existing_workbook_by_filename(mocker):
    mocked_client = mocker.patch('datasheets.convenience.Client', autospec=True)

    _ = datasheets.convenience.create_tab_in_existing_workbook('existing_workbook', 'new_tab')

    mocked_client().fetch_workbook.assert_any_call(filename='existing_workbook')
    mocked_client().fetch_workbook().create_tab.assert_any_call('new_tab')


def test_create_tab_in_existing_workbook_by_file_id(mocker):
    mocked_client = mocker.patch('datasheets.convenience.Client', autospec=True)

    _ = datasheets.convenience.create_tab_in_existing_workbook(filename='existing_workbook',
                                                               tabname='new_tab', file_id='xyz0012')

    mocked_client().fetch_workbook.assert_any_call(file_id='xyz0012')
    mocked_client().fetch_workbook().create_tab.assert_any_call('new_tab')


def test_create_tab_in_new_workbook(mocker):
    mocked_client = mocker.patch('datasheets.convenience.Client', autospec=True)

    _ = datasheets.convenience.create_tab_in_new_workbook('new_workbook', 'new_tab')

    mocked_client().create_workbook.assert_any_call('new_workbook')
    mocked_workbook = mocked_client().create_workbook('new_workbook')
    mocked_workbook.delete_tab.assert_any_call('Sheet1')
    mocked_workbook.create_tab.assert_any_call('new_tab')


def test_create_tab_in_new_workbook_share_with_emails_one_role(mocker):
    mocked_client = mocker.patch('datasheets.convenience.Client', autospec=True)
    emails = ('email1@datasheets.test', 'email2@datasheets.test')

    _ = datasheets.convenience.create_tab_in_new_workbook('new_workbook', 'new_tab', emails=emails,
                                                          role='writer', notify=False)

    mocked_client().create_workbook.assert_any_call('new_workbook')
    mocked_workbook = mocked_client().create_workbook('new_workbook')
    mocked_workbook.delete_tab.assert_any_call('Sheet1')
    mocked_workbook.create_tab.assert_any_call('new_tab')

    for email in emails:
        mocked_workbook.share.assert_any_call(email=email, message=None,
                                              notify=False, role='writer')


def test_create_tab_in_new_workbook_share_with_emails_multiple_roles(mocker):
    mocked_client = mocker.patch('datasheets.convenience.Client', autospec=True)
    emails = ('email1@datasheets.test', 'email2@datasheets.test', 'email2@datasheets.test')
    roles = ('owner', 'reader', 'writer')
    message = 'Here is a spreadsheet for you'

    _ = datasheets.convenience.create_tab_in_new_workbook('new_workbook', 'new_tab', emails=emails,
                                                          role=roles, message=message)

    mocked_client().create_workbook.assert_any_call('new_workbook')
    mocked_workbook = mocked_client().create_workbook('new_workbook')
    mocked_workbook.delete_tab.assert_any_call('Sheet1')
    mocked_workbook.create_tab.assert_any_call('new_tab')

    for email, role in zip(emails, roles):
        mocked_workbook.share.assert_any_call(email=email, message=message,
                                              notify=True, role=role)
