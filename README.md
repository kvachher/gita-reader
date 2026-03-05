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

## Testing

Recommended day-to-day command:

```bash
GITA_TEST_VERBOSE=1 make test
```

This runs all unit tests with per-person validation logs.

How validation works (expected vs actual):

- `Expected` is computed directly from the workbook logistics sheets (`Wednesday` to `Sunday`) by scanning each non-empty cell and assigning it to people via:
  - explicit name matches in cell text
  - group propagation for `ALL_BOARD`, `ALL_LIAISONS`, and `ALL_VOLUNTEERS`
- `Actual` is computed from the generated dashboard data model (`TodoBuilder(...).build()`), using each person’s extracted logistics tasks.
- The tests compare normalized `(sheet, role/column, full cell text)` signatures for every person, and also compare the `ALL_*` subset separately.

Other useful commands:

```bash
make test
```

Runs unit tests only.

```bash
make coverage
```

Runs unit tests and prints per-file + total line coverage summary for `gita_reader`.

## Deploy to Render

This repo is ready for Render:

- Build command: `pip install -r requirements.txt && python scripts/regenerate.py "MM19 Gita.xlsx" --out output`
- Start command: `gunicorn gita_reader.web:app`
- `render.yaml` is already included.

## Notes

- Priority is zero information loss: if parsing templates do not match, full cell text is still retained/displayed.
- Re-run `scripts/regenerate.py` any time the workbook changes.
