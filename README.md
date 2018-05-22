## Overview
datasheets is a library for interfacing with Google Sheets, including reading data from, writing
data to, and modifying the formatting of Google Sheets. It is built on top of Google's
[google-api-python-client](https://github.com/google/google-api-python-client) and
[oauth2client](https://github.com/google/oauth2client) libraries using the
[Google Drive v3](https://developers.google.com/drive/v3/reference/) and
[Google Sheets v4](https://developers.google.com/sheets/reference/rest/) REST APIs.

It can be installed with pip via `pip install datasheets`.

Detailed documentation can be found at [here](https://datasheets.readthedocs.io/en/latest/).


## Basic Usage

Get the necessary OAuth credentials from the Google Developer Console as described in "Getting OAuth
Credentials" doc page.

After that, using datasheets looks like:
```python
import datasheets

client = datasheets.Client()
workbook = client.create_workbook('my_new_workbook')
tab = workbook.create_tab('my_new_tab')

# Create a data set and upload it
import pandas as pd
df = pd.DataFrame([('a', 1.3), ('b', 2.7), ('c', 3.9)], columns=['letter', 'number'])
tab.insert_data(df, index=False)

# Fetch the data again
df_again = tab.fetch_data()
df_again.equals(df)

# Show workbooks you have access to; this may be slow if you are shared on many workbooks
client.fetch_workbooks_info()

# Show tabs within a given workbook
workbook.fetch_tab_names()
```

For further information, see the documentation [here](https://datasheets.readthedocs.io/en/latest/).


## Copyright and License
Copyright 2018 Squarespace, INC.

Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in
compliance with the License. You may obtain a copy of the License at:

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is
distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or
implied. See the License for the specific language governing permissions and limitations under the
License.
