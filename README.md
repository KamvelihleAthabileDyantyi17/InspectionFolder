# Folder Inspection Report

This script recursively scans a configured folder and writes a text inventory of the files it finds. The report groups files by a three-letter filename prefix and includes file paths, sizes, and sample lines from the first file in each group.

## Requirements

- Python 3
- No third-party packages required

## Usage

1. Open `inspect_folder.py` and set `TARGET_FOLDER` to the folder you want to scan.
2. Run the script:

   ```powershell
   python inspect_folder.py
   ```

The report is saved as `folder_inventory_report.txt` inside the target folder.

## Exclusions

The scan skips directories whose paths contain `Backup`, `Live`, or `_unzipped_temp`. It also excludes `.zip` files and existing `folder_inventory_report.txt` files.

---

# Paltrack Masterfile Builder (`paltrack_masterfile_builder.py`)

Turns Paltrack master file exports into one structured Excel workbook.

## Requirements

- Python 3
- `pip install xlsxwriter openpyxl`

## Usage

1. Put each new delivery in its own folder inside `C:\Users\Kamelihled\Desktop\MaterFiles`, named `Batch 1`, `Batch 2`, `Batch 3`, ... A batch can hold all files or only one or two.
2. Run:

   ```powershell
   python paltrack_masterfile_builder.py
   ```

3. Open `MaterFiles\Masterfile_Combined.xlsx`.

## Files it reads

- Fixed-width Paltrack exports (`.B18` etc.). The table is taken from the first two letters of the file name: CG, CO, FA, GR, GT, MK, NC, PE, PK, SI, TK, VA, VG.
- Paltrack **R2 Gtin Export** spreadsheets (`.xlsx`), which go into the GTINs tab. Any other spreadsheet is skipped and noted in the load log. Save `.xls`/`.csv` files as `.xlsx` first.

## What you get

- **Index**: links to every tab, record counts, source files and a load log.
- **Data Issues**: every record that uses a code missing from its master table (e.g. a GTIN with pack `04I`), with the batch and file it came from.
- **GTINs**: GTIN, Orgzn, Commodity, Variety, Variety Name, Variety Group, Variety Group Name, Target Market, Pack, Grade, Mark, Inventory Code, Size Count, Country, Active. Target Market lists every linked market from an R2 GTIN export (e.g. `EU, TE`), or the GTIN's default market otherwise. Change the `VIEWS` setting at the top of the script to alter these columns.
- One tab per other table, with all fields, filters, a frozen header and descriptions next to codes.

## Rules

- The workbook is rebuilt from **all** batch folders on every run. When a record appears more than once, the newest version by Paltrack transaction date is kept. Fields that an Excel export lacks are kept from the `.B18` version.
- Inactive records (Active = N) are kept, so filter on the Active column when needed.
- Close the workbook in Excel before re-running.
