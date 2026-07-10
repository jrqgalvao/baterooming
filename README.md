# BateRooming

> Local Windows desktop tool for comparing Excel rooming lists, matching names, and exporting review-ready results.

BateRooming provides two spreadsheet workflows for operational teams:

- **Bate-Rooming:** compares two rooming lists and highlights matches, divergences, missing records, and repeated names.
- **Name Matching:** matches names between two workbooks while preserving the target spreadsheet structure.

All processing runs locally. The application does not require an external API.

## Highlights

- Excel input and formatted Excel output.
- Name normalization for accents, case, punctuation, and repeated spaces.
- Exact matches prioritized before similarity matching.
- One-to-one assignment: a reference record can be used only once.
- Highest-score selection among compatible candidates.
- Conservative cross-identity matching to reduce false positives.
- Duplicate and placeholder detection.
- Auditable matching decisions in the exported `LOG` worksheet.

## Bate-Rooming output

The exported workbook contains the following worksheets:

| Worksheet | Purpose |
| --- | --- |
| `RESUMO` | Operational summary and key metrics |
| `RESULTADO COMPLETO` | Complete comparison result |
| `DIVERGENCIAS` | Records with relevant differences |
| `SEM CORRESPONDENCIA` | Records without a valid match |
| `REPETIDOS` | Duplicate or placeholder records |
| `LOG` | Matching stage, score, threshold, candidates, conflicts, and decision reason |

### Matching rules

1. Names are normalized before comparison.
2. Exact normalized matches take priority.
3. Repeated names and placeholders are excluded from fuzzy matching.
4. For compatible identities, the candidate with the highest score is selected.
5. Each reference record can be assigned only once.
6. Cross-identity fuzzy matches require a conservative score and compatible first and last significant tokens.
7. Ambiguous candidates and one-to-one conflicts remain visible in `LOG`.

## Requirements

- Windows 10 or later
- Python 3.11 recommended
- WebView2 Runtime

## Quick start

Install the runtime dependencies:

```powershell
python -m pip install -r requirements.txt
```

Run the application locally:

```powershell
python app.py
```

## Development

Install development dependencies:

```powershell
python -m pip install -r requirements-dev.txt
```

Run the automated checks:

```powershell
python -m pytest -q
python -m ruff check app.py core tests
```

## Build

Create the Windows executable with PyInstaller:

```powershell
pyinstaller --noconfirm --clean app.spec
```

The executable is generated at `dist/app/app.exe`. Build output is intentionally not committed to the repository.

## Project structure

```text
.
|-- app.py                  Application entry point
|-- core/                   Matching and Excel processing logic
|-- ui/                     Desktop interface files
|-- tests/                  Automated tests
|-- assets/                 Public generic assets
|-- requirements.txt        Runtime dependencies
|-- requirements-dev.txt    Development dependencies
|-- app.spec                PyInstaller configuration
`-- README_full.md          Extended project documentation
```

## Documentation

See [README_full.md](README_full.md) for the complete public documentation, including detailed workflow behavior, configuration, testing, and publishing guidance.

## Public repository guidance

This repository is a sanitized public mirror. Before distributing an identified version:

- Replace `{{PRODUCT_NAME}}` and `{{COMPANY_NAME}}` placeholders.
- Use approved branding assets.
- Do not commit real spreadsheets, personal data, build folders, caches, or temporary files.
- Run the tests and validate the final executable locally.
