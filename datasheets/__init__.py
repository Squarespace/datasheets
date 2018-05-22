"""
datasheets is a library for interfacing with Google Sheets, including reading from, writing to,
and modifying the formatting of Google Sheets. It is built on top of Google's
google-api-python-client and oauth2client libraries using the Google Drive v3 and Google Sheets
v4 REST APIs. Further details on these libraries and APIs can be found here:

    google-api-python-client: https://github.com/google/google-api-python-client
    oauth2client: https://github.com/google/oauth2client
    Drive v3: https://developers.google.com/drive/v3/reference/
    Sheets v4: https://developers.google.com/sheets/reference/rest/
"""
from datasheets import exceptions
from datasheets.client import Client
from datasheets.convenience import create_tab_in_new_workbook, create_tab_in_existing_workbook
from datasheets.helpers import convert_cell_index_to_label, convert_cell_label_to_index
from datasheets.tab import Tab
from datasheets.workbook import Workbook

__all__ = (
    'Client',
    'Tab',
    'Workbook',
    '__version__',
    'convert_cell_index_to_label',
    'convert_cell_label_to_index',
    'create_tab_in_existing_workbook',
    'create_tab_in_new_workbook',
    'exceptions',
)

__version__ = '0.1.0'
