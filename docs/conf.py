# -*- coding: utf-8 -*-
import datetime as dt

from datasheets import __version__


project = 'datasheets'
copyright = 'Squarespace Data Engineering, {}'.format(dt.datetime.utcnow().year)
author = 'Squarespace Data Engineering'

version = __version__  # The short X.Y version
release = __version__  # The full version, including alpha/beta/rc tags

extensions = [
    'sphinx.ext.autodoc',
    'sphinx.ext.autosectionlabel',
    'sphinx.ext.intersphinx',
    'sphinx.ext.napoleon',
    'sphinx.ext.viewcode',
]

exclude_patterns = ['_build', 'Thumbs.db', '.DS_Store']
html_theme = 'sphinx_rtd_theme'
master_doc = 'index'
nitpick_ignore = [('py:class', 'googleapiclient.discovery.Resource')]
source_suffix = '.rst'

intersphinx_mapping = {
    'python': ('https://docs.python.org/', None),
    'pandas': ('https://pandas.pydata.org/pandas-docs/stable/', None),
}
