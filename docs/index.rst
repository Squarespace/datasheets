Overview
========

datasheets is a library for interfacing with Google Sheets, including reading data from, writing
data to, and modifying the formatting of Google Sheets. It is built on top of Google's
`google-api-python-client`_ and `oauth2client`_ libraries using the `Google Drive v3`_ and
`Google Sheets v4`_ REST APIs.

.. _google-api-python-client: https://github.com/google/google-api-python-client
.. _oauth2client: https://github.com/google/oauth2client) libraries using the
.. _Google Drive v3: https://developers.google.com/drive/v3/reference/
.. _Google Sheets v4: https://developers.google.com/sheets/reference/rest/

It can be installed with pip via ``pip install datasheets``.


Basic Usage
-----------

Get the necessary OAuth credentials from the Google Developer Console as described
in :ref:`Getting OAuth Credentials`.

After that, using datasheets looks like: ::

    import datasheets

    # Create a data set to upload
    import pandas as pd
    df = pd.DataFrame([('a', 1.3), ('b', 2.7), ('c', 3.9)], columns=['letter', 'number'])

    client = datasheets.Client()
    workbook = client.create_workbook('my_new_workbook')
    tab = workbook.create_tab('my_new_tab')

    # Upload a data set
    tab.insert_data(df, index=False)

    # Fetch the data again
    df_again = tab.fetch_data()

    # Show workbooks you have access to; this may be slow if you are shared on many workbooks
    client.fetch_workbooks_info()

    # Show tabs within a given workbook
    workbook.fetch_tab_names()


Documentation Contents
----------------------
.. toctree::
    :maxdepth: 3

    self
    functionality
    getting_oauth_credentials
    comparison_to_gspread
    api_reference
    development

Indices and tables
------------------

* :ref:`genindex`
