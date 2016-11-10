
TEST_PATH=./test/online_test.py

clean-pyc:
	find . -name '*.pyc' -exec rm --force {} +
	find . -name '*.pyo' -exec rm --force {} +
	find . -name '*~' -exec rm --force  {} +
	find . -name '__pycache__' -exec rm -rf {} +

clean-build:
	rm --force --recursive build/
	rm --force --recursive dist/
	rm --force --recursive *.egg-info

lint:
	flake8 --filename = ./pygsheets/*.py

test: clean-pyc
	cd test;py.test -vs $(TEST_PATH);cd ..

install:
	python setup.py install

.PHONY: clean-pyc clean-build