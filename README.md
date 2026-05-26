# Bate Rooming

Desktop tool for automating Excel-based hotel and rooming workflows.

## Tools

- Bate-Rooming: compares two rooming spreadsheets and highlights divergences, missing records, duplicates, check-ins, and check-outs.
- Match de Nomes: compares name lists using fuzzy matching and exports a corrected spreadsheet.

## Tech Stack

- Python
- pywebview
- HTML/CSS/JavaScript
- pandas
- openpyxl
- RapidFuzz

## Install

pip install -r requirements.txt

## Run

python app.py

## Build EXE

pyinstaller --noconfirm app.spec
