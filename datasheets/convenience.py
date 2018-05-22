"""
Convenience functions to simplify common end-user tasks like creating or opening workbooks and tabs

These can be imported accessed directly from the datasheets module, e.g.::

    import datasheets
    tab = datasheets.create_tab_in_existing_workbook(myfilename, mytabname)
"""
from datasheets.client import Client


def create_tab_in_existing_workbook(filename, tabname, file_id=None):
    """Create a new tab in an existing workbook and return an instance of that tab

    Either filename (i.e. title) or file_id should be provided. Providing file_id is preferred
    as it is more precise.


    Args:
        filename (str): The name of the existing workbook in which the tab will be created

        tabname (str): The name of the tab to be created

        file_id (str): The unique file_id for the workbook. If provided, the file_id will be
            used instead of the filename


    Returns:
        datasheets.Tab: An instance of the newly created tab
    """
    kwargs = {'file_id': file_id} if file_id is not None else {'filename': filename}
    return (
        Client()
        .fetch_workbook(**kwargs)
        .create_tab(tabname)
    )


def create_tab_in_new_workbook(filename, tabname, emails=(), role='reader', notify=True,
                               message=None):
    """Create a new tab in a new workbook and return an instance of that tab


    Args:
        filename (str): The name of the workbook to be created

        tabname (str): The name of the tab to be created

        emails (str or tuple): The email address(es) to grant the permission to. This may be one
            address in string form or a series of addresses in tuple form

        role (str or tuple): The type of permission(s) to grant. This can be either a tuple of the
            same size as `emails` or a single value, in which case all emails are granted that
            permission level. Values must be one of 'owner', 'writer', or 'reader'

        notify (bool): If True, send an email notifying the recipient(s) of their granted
            permissions.  These notification emails are the same as what Google sends when a
            document is shared through Google Drive

        message (str): If notify is True, the message to send with the email notification


    Returns:
        datasheets.Tab: An instance of the newly created tab

    """
    workbook = (
        Client()
        .create_workbook(filename)
    )

    tab = (
        workbook
        .create_tab(tabname)
    )

    workbook.delete_tab('Sheet1')

    emails = [emails] if isinstance(emails, str) else emails
    roles = [role] * len(emails) if isinstance(role, str) else role

    for this_email, this_role in zip(emails, roles):
        workbook.share(email=this_email, role=this_role, notify=notify, message=message)

    return tab
