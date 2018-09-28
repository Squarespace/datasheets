Comparison To gspread
======================================
`gspread`_ is a popular library for reading data from Google Sheets, and it was a big inspiration
for this project. datasheets attempts to improve on gspread by:

* Supporting uploading of pandas DataFrames into Google Sheets.
* Ensuring that data pulled from Google Sheets keeps the data type it had within Google Sheets, e.g.
  datetimes will come in as datetimes, numbers as numbers, etc. Within gspread, non-numbers are
  generally all converted to strings.
* Allowing users to authenticate with their own Google account, meaning there is no need to create a
  service account and share all your files with it (though you can still do both of those things).
  Service accounts can be a security liability (as described below under "OAuth Service Account
  Access"); being able to use OAuth Client ID access diminishes that concern.
* Providing a number of additional tools for interacting with Google Sheets: format them, add/remove
  rows and columns, create or delete tabs and workbooks, share or unshare a workbook with users,
  etc. See below for more details.
* Using more modern, Google-maintained tools (e.g. Google's `google-api-python-client`_ and `google_auth`_
  libraries) as opposed to parsing XML feeds.

.. _gspread: https://github.com/burnash/gspread
.. _google-api-python-client: https://github.com/google/google-api-python-client
.. _google_auth: https://github.com/GoogleCloudPlatform/google-auth-library-python
