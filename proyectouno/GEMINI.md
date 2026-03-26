# Project Overview
This is a data processing and automation project designed to consolidate monthly financial or expense reports. It takes multiple legacy Excel (`.xls`) files from a specific month's folder (e.g., `FEBRERO`) and consolidates their data into a single master Excel template (`librouno_ejemplo.xlsx`). The consolidation is performed using Python scripts that read the master file's template formulas, resolve external links to the monthly files, extract the corresponding values, and generate a new consolidated file for that month.

## Main Technologies
- **Python**: Core programming language.
- **openpyxl**: Used for reading and writing modern Excel (`.xlsx`) files, particularly the master template.
- **xlrd**: Used for reading data from legacy Excel (`.xls`) files.
- **Batch Scripting**: Used to provide a simple, interactive CLI wrapper for end-users on Windows.

## Directory Structure and Key Files
- `PROCESAR_MES.bat`: An interactive Windows batch script that asks the user for the month they want to process and executes the main Python script.
- `consolidate_month.py`: The main automation script. It reads the target month, looks up the corresponding folder, evaluates cell formulas in the master template, resolves external references to extract values from the `.xls` files, and outputs a consolidated `_FINAL.xlsx` file.
- `librouno_ejemplo.xlsx`: The master Excel template containing the base structure and formulas (including external links) used as a reference for consolidation.
- `analyze_relation.py`, `check_math.py`, `debug_c45.py`, `inspect_excel.py`: Auxiliary Python scripts likely used for debugging, analyzing cell relationships, inspecting formulas, or verifying mathematical operations within the Excel files.
- `<MONTH_NAME>/` (e.g., `FEBRERO/`): Folders containing the source legacy `.xls` files for the respective month.

## Building and Running
The primary way to use this tool is via the provided batch script.

### Prerequisites
You need Python installed, along with the `openpyxl` and `xlrd` libraries.
```bash
pip install openpyxl xlrd
```

### Running the Consolidation
1. Double-click the `PROCESAR_MES.bat` file from the Windows file explorer, or run it from the command line:
   ```cmd
   .\PROCESAR_MES.bat
   ```
2. When prompted, enter the name of the month you want to process (e.g., `FEBRERO`, `MARZO`). The script will look for a folder with that exact name.
3. Alternatively, you can run the Python script directly from the terminal, passing the month as an argument:
   ```bash
   python consolidate_month.py FEBRERO
   ```

## Development Conventions
- **Naming Conventions**: Legacy source files seem to follow a pattern (e.g., `5105_CHVS_FEB_2026.xls`), and the script uses partial string matching based on the links in the template to identify the correct files.
- **Data Extraction**: The `consolidate_month.py` uses column `C` (Febrero) in `librouno_ejemplo.xlsx` as the structural template for formulas and translates references and external links relative to the target month column.
- **Error Handling**: The scripts implement basic fallback mechanisms (e.g., generating a `_V2` file if there's a permission error when saving the Excel output). When developing, ensure you maintain robust file handling since Excel files might be locked by the OS or the user.