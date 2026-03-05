# Gita Reader

Deployable workbook-to-dashboard app for competition logistics.

## Project structure

- `gita_reader/pipeline.py`: extraction + parsing + page generation
- `gita_reader/cli.py`: command-line entrypoint
- `gita_reader/web.py`: Flask app for Render/local hosting
- `scripts/regenerate.py`: easy refresh command for new `.xlsx`
- `output/`: generated dashboard pages + JSON
- `render.yaml`: Render service config

## Regenerate from spreadsheet (easy refresh)

```bash
python scripts/regenerate.py "MM19 Gita.xlsx" --out output
```

Or:

```bash
make refresh XLSX="MM19 Gita.xlsx"
```

This rebuilds:

- `output/index.html` (calendar home)
- `output/personal.html` (personal gitas)
- `output/info.html` (contacts + important info)
- `output/all_todos.json`

## Run locally

```bash
pip install -r requirements.txt
make run
```

Then open `http://localhost:8000`.

## Deploy to Render

This repo is ready for Render:

- Build command: `pip install -r requirements.txt && python scripts/regenerate.py "MM19 Gita.xlsx" --out output`
- Start command: `gunicorn gita_reader.web:app`
- `render.yaml` is already included.

## Notes

- Priority is zero information loss: if parsing templates do not match, full cell text is still retained/displayed.
- Re-run `scripts/regenerate.py` any time the workbook changes.
