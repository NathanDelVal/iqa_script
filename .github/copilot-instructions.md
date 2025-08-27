# Copilot Instructions for IQA Script Project

## Project Overview
- This project processes IQA (quality analysis) Excel files, extracting and transforming data for reporting.
- Main logic is in `index.py`, which loads multiple Excel files, processes specific sheets, and outputs a consolidated Excel file.
- Configuration and constants are in `params.py` (e.g., file patterns, block/empresa names).

## Key Files
- `index.py`: Main entry point. Handles CLI args, file validation, data processing, and output.
- `params.py`: Stores regex patterns, block/empresa names, and example file references.

## Data Flow
- Input: Excel files named with a pattern like `(a).xlsx`, `(b).xlsx`, `(c).xlsx`.
- Sheets processed: `05-PLN_AMT_VRF` and `08-RST_ANL_VRF`.
- Data is cleaned (columns dropped, types inferred), annotated (block, empresa, month), and merged.
- Output: Single Excel file with processed sheets, named by month/year.

## Conventions & Patterns
- File pattern for input: `params.file_pattern` (e.g., `(a)`, `(b)`, `(c)` in filename).
- Block/empresa mapping: `params.bloco_a`, `params.bloco_b`, `params.bloco_c`.
- Sheet names are mapped via `sheet_reference` in `index.py`.
- Uses pandas for all data manipulation.
- CLI arguments control input files, reference month, output file, and log file.
- If a required column is missing, code handles with `None` or default values.
- All code expects to be run as a script: `python index.py [files] [--referencia N] [--output name]`.

## Developer Workflows
- No explicit test or build system; run `python index.py ...` to process data.
- To add new blocks or empresas, update `params.py` and `blocos` in `index.py`.
- To change input file validation, update `params.file_pattern`.
- For new sheet logic, extend the `for sheet in sheet_reference_name` loop in `index.py`.

## External Dependencies
- Requires: `pandas`, `dateparser`, `openpyxl` (for Excel I/O).
- Install dependencies with: `pip install -r requirements.txt`

## Example Usage
```sh
python index.py (A).xlsx (B).xlsx --referencia 8 --output IQA_08-2025.xlsx
```

## Notes
- All data transformations are explicit in `index.py`.
- No hidden side effects or background processes.
- If you encounter a new file pattern or block, update both `params.py` and `index.py` accordingly.
