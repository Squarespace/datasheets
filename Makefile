.PHONY: clean coverage docs release_pypi release_pypitest test view_docs

VERSION := `grep "^__version__" datasheets/__init__.py | cut -d "'" -f 2`


clean:
	@find . -name '__pycache__' -type d -exec rm -rf {} +
	@find . -name '*.pyc' -delete
	@find . -name '*.retry' -delete

docs:
	$(MAKE) -C docs html O=-nW

coverage:
	pytest --cov datasheets/ --cov-report=term-missing:skip-covered

release_pypi: test
	rm -rf dist/
	python setup.py sdist bdist_wheel upload -r pypi
	rm -rf dist/

release_pypitest: test
	rm -rf dist/
	python setup.py sdist bdist_wheel upload -r pypitest
	rm -rf dist/

test:
	@tox

view_docs: docs
	open docs/_build/html/index.html
