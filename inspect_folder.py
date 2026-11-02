import os
import zipfile
from pathlib import Path

# Point this to your main folder containing the nested master files
TARGET_FOLDER = r"C:\Users\Kamelihled\Desktop\MaterFiles"
OUTPUT_REPORT = os.path.join(TARGET_FOLDER, "folder_inventory_report.txt")

print(f"Scanning folder recursively: {TARGET_FOLDER}")

all_data_files = []

# Walk through all directories and subdirectories recursively
for root, dirs, files in os.walk(TARGET_FOLDER):
    # Skip backup or temp system folders if needed
    if "Backup" in root or "Live" in root or "_unzipped_temp" in root:
        continue
        
    for file in files:
        # Ignore zips and report files
        if file.lower().endswith(".zip") or file == "folder_inventory_report.txt":
            continue
            
        full_path = os.path.join(root, file)
        all_data_files.append(full_path)

print(f"Found {len(all_data_files)} data files across nested folders. Generating report...")

grouped_summary = {}

for f_path in all_data_files:
    f_obj = Path(f_path)
    fname = f_obj.name
    
    # Extract 3-letter prefix (fallback to parent folder name if filename doesn't start with 3 letters)
    prefix = fname[:3].upper() if len(fname) >= 3 and fname[:3].isalpha() else f_obj.parent.name[:3].upper()
    
    if prefix not in grouped_summary:
        grouped_summary[prefix] = []
    
    # Read first 3 lines of the file to inspect headers & delimiters
    sample_rows = []
    try:
        with open(f_path, 'r', encoding='utf-8', errors='ignore') as file:
            for _ in range(3):
                line = file.readline().strip()
                if line:
                    sample_rows.append(line)
    except Exception as e:
        sample_rows.append(f"<Error reading file: {e}>")
        
    grouped_summary[prefix].append({
        "filename": fname,
        "relative_path": str(f_obj.relative_to(TARGET_FOLDER)),
        "size_bytes": os.path.getsize(f_path),
        "samples": sample_rows
    })

# Format the inventory report
report_lines = [
    "="*80,
    "RECURSIVE FOLDER INSPECTION REPORT",
    f"Target Directory: {TARGET_FOLDER}",
    f"Total Data Files Discovered: {len(all_data_files)}",
    "="*80 + "\n"
]

for prefix, file_list in sorted(grouped_summary.items()):
    report_lines.append(f"--- PREFIX GROUP: [{prefix}] ({len(file_list)} files) ---")
    
    # Show sample from the first file in this group
    template_file = file_list[0]
    report_lines.append(f"Sample File: {template_file['relative_path']} ({template_file['size_bytes']} bytes)")
    report_lines.append("Header / Top Rows:")
    for idx, sample in enumerate(template_file['samples'], 1):
        report_lines.append(f"   Line {idx}: {sample}")
    
    report_lines.append("All files in this prefix group:")
    for item in file_list:
        report_lines.append(f"   - {item['relative_path']} ({item['size_bytes']} bytes)")
    report_lines.append("\n" + "-"*60 + "\n")

# Write report
with open(OUTPUT_REPORT, "w", encoding="utf-8") as f:
    f.write("\n".join(report_lines))

print(f"\nInspection complete! Clean summary generated at:\n{OUTPUT_REPORT}")