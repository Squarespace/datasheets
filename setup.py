from setuptools import setup


__version__ = '0.1'

version_suffix = ''
try:
    with open('.dev_version_suffix') as f:
        version_suffix = f.readline().strip()
except IOError:
    pass

INSTALL_REQUIRES = [
    'pandas',
    'oauth2client>=3.0.0',
    'google-api-python-client>=1.5.4',
    'PyOpenSSL',  # used by oauth2client
]

setup(
    name='gsheets',
    version=__version__ + version_suffix,
    description='Library for reading from, writing to, and formatting Google Sheets from Python',
    author='Squarespace Analytics',
    author_email='analytics-all@squarespace.com',
    url='https://stash.nyc.squarespace.net/projects/STRAT/repos/gsheets/browse',
    packages=[
        'gsheets',
    ],
    package_dir={'gsheets':
                 'gsheets'},
    include_package_data=True,

    install_requires=INSTALL_REQUIRES,

    license='Copyright',
    zip_safe=False,
    classifiers=[
        'Private :: Do Not Upload',
    ],
)
