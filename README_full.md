# BateRooming: Full Project Documentation

This document provides detailed technical and operational documentation for BateRooming.

For a concise project overview, see the [main README](README.md).

## Project Overview

BateRooming is a Windows desktop application for reconciling hospitality spreadsheets. It combines two workflows in a single interface:

- **Rooming List Reconciliation:** compares internal system data against a hotel rooming list.
- **Name Matching:** identifies likely name matches between two spreadsheets.

The application processes files locally and exports structured Excel workbooks for operational review.

## Current Version

- Application version: **1.4.4**
- Recommended Python version: **3.11**
- Target platform: **Windows 10 or later**
- Packaging mode: PyInstaller `onedir`
- Main executable output: `dist/app/app.exe`

## Application Architecture

The project separates the desktop interface, application API, business rules, and Excel export logic.

```text
UI pages
   |
   v
pywebview API in app.py
   |
   v
Business rules in core/
   |
   v
Excel readers and exporters
```

### Interface Layer

The `ui/` directory contains local HTML, CSS, and JavaScript pages:

- `menu_ui.html`: main navigation and changelog.
- `bate_rooming_ui.html`: rooming reconciliation workflow.
- `match_nomes_ui.html`: name-matching workflow.

The interface communicates with Python through the API exposed by `pywebview`.

### Application Layer

`app.py` is responsible for:

- creating and configuring the desktop window;
- loading local UI pages;
- exposing Python methods to JavaScript;
- managing temporary application state;
- validating user-selected files;
- coordinating business rules and exports;
- opening exported files and folders;
- translating exceptions into user-friendly messages.

### Business Rule Layer

The `core/` directory contains the main processing logic:

- `bate_rooming.py`: rooming reconciliation rules.
- `bate_rooming_export.py`: formatted reconciliation workbook export.
- `match_nomes.py`: name-matching workflow.
- `match_nomes_export.py`: reference-workbook-preserving export.
- `matching.py`: shared normalization and similarity utilities.

## Rooming List Reconciliation Workflow

### Purpose

Compare an internal source spreadsheet against a hotel rooming list and organize differences for review.

### User Flow

1. Select the internal system spreadsheet.
2. Select the hotel rooming list.
3. Choose whether room differences should be ignored.
4. Run the comparison.
5. Review summary indicators and filtered records.
6. Export all records or only visible results.

### Processing Rules

- Spreadsheet columns are detected from known header aliases when possible.
- Files without headers are supported through fallback layout detection.
- Names and room values are normalized before comparison.
- Duplicate records are classified separately.
- Records without a valid match are classified as unmatched.
- Name and room differences are identified independently.
- When **Ignore Room** is enabled, room differences do not affect the overall status.

### Exported Workbook

The generated workbook contains dedicated sheets for:

- complete results;
- discrepancies;
- unmatched records;
- duplicate records.

Formatting, headers, column widths, status values, and summary information are applied during export.

## Name Matching Workflow

### Purpose

Find the most likely correspondence between names from two spreadsheets while preserving the reference workbook layout.

### User Flow

1. Select the system-name spreadsheet.
2. Select the reference or hotel spreadsheet.
3. Configure the minimum similarity threshold.
4. Run name matching.
5. Review matched, unmatched, and empty records.
6. Export results into a copy of the reference workbook.

### Processing Rules

- The name column is detected automatically when possible.
- Text is trimmed and normalized before comparison.
- Significant tokens are used to improve matching quality.
- Fuzzy similarity is calculated with RapidFuzz.
- A configurable threshold determines accepted matches.
- Empty names remain explicitly classified.
- Matching results include status and similarity score.

### Export Behavior

The exporter uses the selected reference workbook as a template and preserves:

- existing rows;
- worksheet structure;
- cell formatting;
- dimensions and column widths;
- the original name positions.

Matching results are written back to their corresponding rows without rebuilding the workbook from scratch.

## Project Structure

```text
.
├── app.py                     # Desktop entry point and pywebview API
├── app.spec                   # PyInstaller configuration
├── app_version_info.txt       # Windows executable metadata
├── app.exe.config             # Windows runtime configuration
├── assets/
│   ├── app_icon.ico
│   ├── logo_generic_color.png
│   └── logo_generic_white.png
├── core/
│   ├── bate_rooming.py
│   ├── bate_rooming_export.py
│   ├── match_nomes.py
│   ├── match_nomes_export.py
│   └── matching.py
├── tests/
├── ui/
├── requirements.txt
├── requirements-dev.txt
├── pytest.ini
├── README.md
└── README_full.md
```

## Dependencies

### Runtime Dependencies

Defined in `requirements.txt`:

- `pywebview`: desktop window and JavaScript/Python integration.
- `openpyxl`: Excel workbook reading and writing.
- `rapidfuzz`: fuzzy text comparison.
- `xlrd`: compatibility support for older spreadsheet formats.

### Development Dependencies

Defined in `requirements-dev.txt`:

- runtime dependencies;
- `pytest`;
- `ruff`;
- `pyinstaller`.

`pandas` and `numpy` are intentionally not required by the application.

## Development Setup

Clone the repository:

```bash
git clone https://github.com/urukrehn/baterooming.git
cd baterooming
```

Create and activate a virtual environment:

```bash
python -m venv .venv
.venv\Scripts\activate
```

Install runtime dependencies:

```bash
pip install -r requirements.txt
```

Install development dependencies when testing or building:

```bash
pip install -r requirements-dev.txt
```

Run the application:

```bash
python app.py
```

## Testing and Quality Checks

Run the complete automated test suite:

```bash
pytest -q
```

Run static analysis:

```bash
ruff check . --exclude build --exclude dist
```

The automated tests cover:

- shared name-matching rules;
- rooming reconciliation behavior;
- export behavior;
- application API responses;
- UI structure and accessibility attributes;
- runtime and development dependency separation;
- required packaging assets;
- executable version metadata.

## Building the Windows Package

Generate a clean executable package:

```bash
pyinstaller --noconfirm --clean app.spec
```

Expected output:

```text
dist/app/app.exe
```

The PyInstaller configuration includes:

- local HTML interfaces;
- generic visual assets;
- Windows version metadata;
- executable icon;
- runtime configuration;
- exclusions for unused heavy dependencies.

## Release Checklist

Before publishing a new version:

1. Update the visible version in the UI.
2. Update the version reference in `app.py`.
3. Update `app_version_info.txt`.
4. Run `pytest -q`.
5. Run `ruff check . --exclude build --exclude dist`.
6. Generate the executable with `pyinstaller --noconfirm --clean app.spec`.
7. Open the generated executable.
8. Validate navigation between all screens.
9. Run one controlled example through each workflow.
10. Open and review the exported Excel files.
11. Confirm that no sensitive spreadsheets or generated files are included.

## Troubleshooting

### The Application Does Not Open

Check that:

- runtime dependencies are installed;
- Microsoft Edge WebView2 Runtime is available;
- all files under `ui/` and `assets/` exist;
- the application is being run from a complete package.

### Spreadsheet Processing Fails

Check that:

- the selected spreadsheet is valid and accessible;
- the file is not locked by Excel;
- the destination folder is writable;
- the spreadsheet contains recognizable data columns.

### The Executable Package Is Unexpectedly Large

Check that:

- `pandas` and `numpy` were not added to runtime dependencies;
- `app.spec` still excludes unused packages;
- the package was generated with `--clean`.

### Images Are Missing

Check that:

- all files under `assets/` exist;
- `app.spec` includes the assets;
- UI paths still reference `../assets/...`.

## Privacy and Data Handling

- All core processing runs locally.
- No spreadsheet data is automatically uploaded.
- External APIs are not required for reconciliation or matching.
- Real guest data, generated workbooks, and temporary files must not be committed.

## Maintenance Guidelines

- Add or update tests before changing matching rules.
- Preserve workbook structures unless a requirement explicitly changes them.
- Keep runtime and development dependencies separated.
- Do not remove UI assets without validating the packaged application.
- Prefer small, reviewable changes over broad rewrites.

## License

This project does not currently include a license. Add a `LICENSE` file before reusing or distributing the code.
