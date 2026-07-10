# BateRooming

BateRooming is a Windows desktop tool for matching guest names across Excel rooming lists and exporting the reviewed result.

The complete public documentation is available in [README_full.md](README_full.md).

## Current behavior

- Matches names using normalized text, identity compatibility, and fuzzy similarity rules.
- Keeps each name assigned at most once and selects the highest score within the applicable identity group.
- Uses a conservative cross-identity threshold to reduce false matches.
- Exports the existing result sheets plus a `LOG` sheet with the decision, score, candidates, conflicts, and reason used for each result.
- Preserves the existing result format and workbook workflow.

## Development

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest -q
python -m ruff check app.py core tests
```

The executable is built locally with PyInstaller and is intentionally not committed to this repository.
