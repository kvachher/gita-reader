XLSX ?= MM19 Gita.xlsx
OUT ?= output

.PHONY: refresh run web

refresh:
	python scripts/regenerate.py "$(XLSX)" --out "$(OUT)"

web:
	gunicorn gita_reader.web:app --bind 0.0.0.0:8000 --reload

run: refresh web
