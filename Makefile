XLSX ?= MM19 Gita.xlsx
OUT ?= output

.PHONY: refresh run web test coverage

refresh:
	python scripts/regenerate.py "$(XLSX)" --out "$(OUT)"

web:
	gunicorn gita_reader.web:app --bind 0.0.0.0:8000 --reload

run: refresh web

test:
	python -m unittest discover -s tests -p 'test_*.py'

coverage:
	python scripts/coverage.py
