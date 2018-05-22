import ast
import re
from setuptools import setup


def ensure_one_level_of_quotes(text):
    # Converts '"foo"' to 'foo'
    return str(ast.literal_eval(text))


def get_version():
    """ Based on the functionality in pallets/click's setup.py
    (https://github.com/pallets/click/blob/master/setup.py) """
    _version_re = re.compile(r'__version__\s+=\s+(.*)')
    with open('datasheets/__init__.py', 'rb') as f:
        lines = f.read().decode('utf-8')
        version = ensure_one_level_of_quotes(_version_re.search(lines).group(1))
        return version


required = [
    'pandas',
    'numpy',
    'oauth2client>=3.0.0',
    'google-api-python-client>=1.5.4',
    'six>=1.10.0',  # required by google-api-python-client but not installed by it
    'PyOpenSSL',  # used by oauth2client
    'httplib2==0.9.2',  # pin this; 0.10.2 causes httplib2.CertificateValidationUnsupported error
]

setup(
    name='datasheets',
    description='Read data from, write data to, and format Google Sheets from Python',
    version=get_version(),
    author='Squarespace Data Engineering',
    url='https://github.com/Squarespace/datasheets',
    download_url='https://github.com/Squarespace/datasheets/tarball/{}'.format(get_version()),
    packages=['datasheets'],
    install_requires=required,
    license='Apache License 2.0',
    classifiers=[
        'Programming Language :: Python',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3.6',
    ],
)
