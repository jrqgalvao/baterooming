# BateRooming

A desktop application that automates the comparison, validation, and organization of hospitality spreadsheets.

BateRooming was designed to reduce the manual effort involved in reconciling hotel rooming lists and matching guest names across different data sources. It identifies matches, discrepancies, duplicates, and missing records, then exports structured Excel reports for operational review.

All spreadsheet processing runs locally. No external API or cloud service is required.

## Documentation

For architecture details, business rules, workflows, build instructions, and release guidance, see the [full project documentation](README_full.md).

## Business Problem

Hospitality teams often receive guest and rooming data in spreadsheets with inconsistent layouts, headers, names, and room information. Reviewing these files manually is time-consuming and increases the risk of operational errors.

BateRooming provides a repeatable workflow to:

- compare records from different sources;
- identify inconsistencies before operations begin;
- standardize name matching;
- organize exceptions for faster review;
- produce formatted reports for decision-making.

## Core Features

### Rooming List Reconciliation

Compares an internal system spreadsheet against a hotel rooming list.

- Automatically detects relevant columns across different layouts.
- Supports spreadsheets with or without headers.
- Identifies name and room discrepancies.
- Separates duplicate and unmatched records.
- Provides an option to ignore room differences when required.
- Exports complete results or only the currently filtered records.
- Generates dedicated Excel sheets for each result category.

### Name Matching

Finds the most likely matches between names from two spreadsheets.

- Normalizes names before comparison.
- Uses fuzzy text matching to handle spelling and formatting differences.
- Allows users to configure the minimum similarity threshold.
- Classifies matched, unmatched, and empty records.
- Preserves the original structure and formatting of the reference workbook.

## Technical Highlights

- Desktop interface built with `pywebview` and local HTML, CSS, and JavaScript.
- Business rules isolated from the interface layer.
- Excel reading, transformation, and export implemented with `openpyxl`.
- Fuzzy matching implemented with `RapidFuzz`.
- Automated tests covering business rules, exports, UI integration, and packaging.
- Windows executable generated with PyInstaller.
- Optimized runtime package without unnecessary `pandas` or `numpy` dependencies.

## Technology Stack

- Python 3.11
- pywebview
- openpyxl
- RapidFuzz
- HTML, CSS, and JavaScript
- PyInstaller
- pytest
- Ruff

## Project Structure

```text
.
├── app.py                 # Desktop application and UI integration
├── core/                  # Business rules and Excel exporters
├── ui/                    # HTML, CSS, and JavaScript interfaces
├── assets/                # Generic icons and visual assets
├── tests/                 # Automated test suite
├── app.spec               # PyInstaller build configuration
├── requirements.txt       # Runtime dependencies
└── requirements-dev.txt   # Development and build dependencies
```

## Requirements

- Windows 10 or later
- Python 3.11 recommended
- Microsoft Edge WebView2 Runtime

## Installation

```bash
git clone https://github.com/jrqgalvao/baterooming.git
cd baterooming
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Running the Application

```bash
python app.py
```

The application opens a desktop menu with access to the rooming reconciliation and name-matching workflows.

## Building the Windows Executable

```bash
pip install -r requirements-dev.txt
pyinstaller --noconfirm --clean app.spec
```

The executable package is created at:

```text
dist/app/app.exe
```

## Testing and Code Quality

Run the automated test suite:

```bash
pytest -q
```

Run static analysis:

```bash
ruff check . --exclude build --exclude dist
```

The test suite covers matching rules, rooming reconciliation, exports, UI API behavior, accessibility checks, and executable package structure.

## Privacy

- All spreadsheets are processed locally.
- No data is automatically uploaded or shared.
- The core workflows do not depend on external APIs.
- Real guest spreadsheets and generated reports should never be committed to the repository.

## Customization

This public repository uses generic visual assets and company/product placeholders. Before distributing a customized version, review:

- interface labels and titles;
- executable metadata in `app_version_info.txt`;
- icons and images in `assets/`;
- the application title in `app.py`.

## Contributing

1. Create a dedicated branch.
2. Keep changes small and focused.
3. Add or update tests when changing business rules.
4. Run the test suite and lint checks before opening a pull request.

## License

This project does not currently include a license. Add a `LICENSE` file before reusing or distributing the code.
