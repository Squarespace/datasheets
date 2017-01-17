'''
Convenience functions to simplify common end-user tasks like creating or opening workbooks and tabs
'''
from gsheets import core


def create_tab_in_new_workbook(filename, tabname, emails=None, role='reader', notify=True,
                               message=None, client=None):
    '''
    filename
        The name of the workbook to be created

    tabname
        The name of the tab to be created

    emails
        The email address(es) to grant the permission to. This may be one address in string form
        or a series of addresses in list form

    role
        The type of permission(s) to grant. This can be either a list of the same size as `emails`
        or a single value, in which case all emails are granted that permission level. Values
        must be one of:
            - 'owner'
            - 'writer'
            - 'reader'

    notify
        Whether to notify via email the recipient(s) granted the permission(s)

    message
        If notify==True, the message to send with the email notification

    client
        An optional gsheets.Client instance to use. This is primarily used
        for unit testing
    '''
    client = client or core.Client(storage=False)
    wkb = client.create_workbook(filename)
    tab = wkb.create_tab(tabname)
    wkb.delete_tab('Sheet1')

    if emails:
        emails = [emails] if isinstance(emails, str) else emails
        roles = [role] * len(emails) if isinstance(role, str) else role

        for this_email, this_role in zip(emails, roles):
            wkb.add_permission(email=this_email, role=this_role, notify=notify, message=message)

    return tab


def create_tab_in_existing_workbook(filename, tabname, file_id=None, client=None):
    '''
    filename
        The name of the workbook to be created

    tabname
        The name of the tab to be created

    file_id
        The unique file_id for the workbook. If provided, the file_id will be used as
        it is more precise

    client
        An optional gsheets.Client instance to use. This is primarily used
        for unit testing
    '''
    client = client or core.Client(storage=False)
    wkb = client.get_workbook(filename=filename, file_id=file_id)
    return wkb.create_tab(tabname)


def open_tab(filename, tabname, file_id=None, client=None):
    '''
    filename
        The name of the workbook to be created

    tabname
        The name of the tab to be created

    file_id
        The unique file_id for the workbook. If provided, the file_id will be used as
        it is more precise

    client
        An optional gsheets.Client instance to use. This is primarily used
        for unit testing
    '''
    client = client or core.Client(storage=False)
    return client.get_workbook(filename=filename, file_id=file_id).get_tab(tabname)
