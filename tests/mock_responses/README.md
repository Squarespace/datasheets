Each file is a dictionary (in yaml form) of methodIds that map to (httplib2.Response, content, optional_expected_body)

A null httplib2.Response will return {'status': '200'}. If optional_expected_body is
included, it will be compared against the request's body and UnxpectedBodyError raised
on inequality.

 For more info, see googlapiclient.http.RequestMockBuilder documentation
here:
    https://github.com/google/google-api-python-client/blob/master/googleapiclient/http.py#L1486

