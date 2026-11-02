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
