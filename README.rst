datasheets
==========
|pip_versions| |travis_ci| |coveralls|

.. |pip_versions| image:: https://img.shields.io/pypi/pyversions/datasheets.svg
    :target: https://pypi.python.org/pypi/datasheets

.. |travis_ci| image:: https://travis-ci.org/Squarespace/datasheets.svg?branch=master
    :target: https://travis-ci.org/Squarespace/datasheets

.. |coveralls| image:: https://coveralls.io/repos/github/Squarespace/datasheets/badge.svg?branch=master
    :target: https://coveralls.io/github/Squarespace/datasheets?branch=master


datasheets is a library for interfacing with Google Sheets, including reading data from, writing
data to, and modifying the formatting of Google Sheets. It is built on top of Google's
`google-api-python-client`_ and `oauth2client`_ libraries using the `Google Drive v3`_ and
`Google Sheets v4`_ REST APIs.

.. _google-api-python-client: https://github.com/google/google-api-python-client
.. _oauth2client: https://github.com/google/oauth2client
.. _Google Drive v3: https://developers.google.com/drive/v3/reference/
.. _Google Sheets v4: https://developers.google.com/sheets/reference/rest/

It can be installed with pip via ``pip install datasheets``.

Detailed information can be found in the `documentation`_.

.. _documentation: https://datasheets.readthedocs.io/en/latest/


Basic Usage
-----------
Get the necessary OAuth credentials from the Google Developer Console as described
in `Getting OAuth Credentials`_.

.. _Getting OAuth Credentials: https://datasheets.readthedocs.io/en/latest/getting_oauth_credentials.html

After that, using datasheets looks like:

.. code-block:: python

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

For further information, see the `documentation`_.


License
-------
Copyright 2018 Squarespace, INC.

Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in
compliance with the License. You may obtain a copy of the License at:

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is
distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or
implied. See the License for the specific language governing permissions and limitations under the
License.
