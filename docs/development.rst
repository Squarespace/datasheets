Developing / Testing
====================

Getting Set Up
--------------
First, get your Python environment set up:

.. code-block:: bash

    mkvirtualenv datasheets
    pip install -e . -r requirements-dev.txt

Various testing functionality exists:

    * ``make test`` - Run tests for both Python 2 and 3
    * ``pytest`` - Run tests for whichever Python version is in your virtualenv
    * ``make coverage`` - Check code coverage

Manual tests also exist in the ``tests/manual_testing.ipynb`` Jupyter Notebook. To run the manual
tests, install Jupyter Notebook (``pip install jupyter notebook``), then run ``jupyter notebook``,
open the file in the browser, and execute each cell.

Releasing A New Version
-----------------------
If you make a PR that gets merged into master, a new version of datasheets can be created as follows.

1. Increment the ``__version__`` in the ``datasheets/__init__.py`` file and commit that change.
2. Push a new git tag to the repo by doing:

    * Write the tag message in a dummy file called ``tag_message``. We do this to allow multi-line tag
      messages
    * ``git tag x.x.x -F tag_message``
    * ``git push --tags origin master``

3. Run ``make release_pypitest`` to test that you can release to pypi.
4. Run ``make release_pypi`` to actually push the release to pypi.
