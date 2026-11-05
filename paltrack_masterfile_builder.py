"""
Paltrack masterfile export -> one structured Excel workbook.

Reads every "Batch ..." folder inside MASTER_ROOT, decodes the fixed-width
Paltrack export files (e.g. FALC4028.B18, GTLC7157.B18) using the R2 table
layouts, also reads Paltrack "R2 Gtin Export" Excel files (.xlsx), merges all
batches, and writes one workbook:

    Index            - links to every table, record counts, load log
    Data Issues      - every record that uses a code missing from its master table
    one tab per table - Excel Table with filters, frozen header,
                        code descriptions looked up from the other tables

Re-run it whenever a new batch folder is added: the workbook is rebuilt from
all batches. If the same record (e.g. the same GTIN or farm code) appears in
more than one batch, the newest version (by Paltrack transaction date) is kept.

Requires:  pip install xlsxwriter openpyxl
Run:       python paltrack_masterfile_builder.py
"""

import os
import re
import sys
from datetime import datetime, date

try:
    import xlsxwriter
    import openpyxl
except ImportError:
    sys.exit("Missing package. Run:  pip install xlsxwriter openpyxl")

# --------------------------------------------------------------------------
# SETTINGS
# --------------------------------------------------------------------------
MASTER_ROOT = r"C:\Users\Kamelihled\Desktop\MaterFiles"
OUTPUT_FILE = os.path.join(MASTER_ROOT, "Masterfile_Combined.xlsx")
BATCH_FOLDER_PATTERN = re.compile(r"^batch\s*\d+", re.IGNORECASE)
FILE_ENCODING = "cp1252"          # Paltrack exports are Windows-1252 (™, –, ’)

# --------------------------------------------------------------------------
# TABLE LAYOUTS  (from the "R2 - PALTRACK Table Layouts" PDFs)
# Each field: (column name, from position, to position, type)
# Positions are 1-based and inclusive, exactly as printed in the layout PDFs.
# Types: text | int | num | date (YYYYMMDD)
# Standard fields (TRAN_USER, TRAN_DATE/TIME, ACTIVE, ORGZN, LANG) are added
# automatically - see build_record().
# --------------------------------------------------------------------------
TABLES = {
    "GT": {
        "sheet": "GTINs", "title": "GTIN", "orgzn": (54, 55), "lang": False,
        "key": ["GTIN", "START_DATE"],
        "fields": [
            ("GTIN", 109, 122, "text"),
            ("COMMODITY", 56, 57, "text"),
            ("COUNTRY", 58, 59, "text"),
            ("VARIETY", 60, 62, "text"),
            ("PACK", 63, 66, "text"),
            ("GRADE", 67, 70, "text"),
            ("MARK", 71, 75, "text"),
            ("INV_CODE", 76, 77, "text"),
            ("SIZE_COUNT", 78, 82, "text"),   # layout lists MASS on the same positions
            ("MKTSPEC", 83, 92, "text"),
            ("START_DATE", 93, 100, "date"),
            ("END_DATE", 101, 108, "date"),
            ("TRAN_NO", 49, 53, "int"),
            ("SEASON", 123, 126, "text"),
            ("DIAMETERA", 127, 135, "text"),
            ("COMM_GRP", 136, 137, "text"),
            ("NETT_MASS", 138, 147, "num"),
            ("GROSS_MASS", 148, 157, "num"),
            ("EQV_CONV", 158, 165, "num"),
            ("EQV_SIZE", 166, 170, "text"),
            ("EQV_PACK", 171, 174, "text"),
            ("EQV_VARIETY", 175, 177, "text"),
            ("UPC_CODE", 178, 192, "text"),
            ("DEFLT_TARG_MKT", 193, 194, "text"),
            ("PICK_TYPE", 195, 195, "text"),
        ],
    },
    "FA": {
        "sheet": "Farms", "title": "Farm", "key": ["FARM"],
        "fields": [
            ("FARM", 432, 438, "text"),
            ("NAME", 67, 106, "text"),
            ("LONG_DESC", 41, 56, "text"),
            ("MED_DESC", 457, 463, "text"),
            ("CONTACT", 347, 376, "text"),
            ("PHONE", 378, 397, "text"),
            ("FAX", 398, 417, "text"),
            ("POST_ADR1", 107, 146, "text"),
            ("POST_ADR2", 147, 186, "text"),
            ("POST_ADR3", 187, 226, "text"),
            ("STR_ADR1", 227, 266, "text"),
            ("STR_ADR2", 267, 306, "text"),
            ("STR_ADR3", 307, 346, "text"),
            ("COUNTRY", 424, 425, "text"),
            ("MAIN_AREA", 426, 427, "text"),
            ("STAT_AREA", 428, 429, "text"),
            ("FARM_AREA", 430, 431, "text"),
            ("ORCH_FLAG", 377, 377, "text"),
            ("FURN", 418, 423, "text"),
            ("VAT_NO", 464, 483, "text"),
            ("GLOBAL_LOCATION_NUMBER", 484, 503, "text"),
            ("GLOBAL_GAP_NUMBER", 504, 523, "text"),
        ],
    },
    "PK": {
        "sheet": "Packs", "title": "Pack", "key": ["COMMODITY", "PACK"],
        "fields": [
            ("COMMODITY", 34, 35, "text"),
            ("PACK", 36, 39, "text"),
            ("LONG_DESC", 40, 55, "text"),
            ("MED_DESC", 56, 65, "text"),
            ("ALONG", 66, 81, "text"),
            ("AMED", 82, 91, "text"),
            ("LABEL_DESC", 258, 267, "text"),
            ("SORT_SEQ", 92, 96, "int"),
            ("LAYERS", 97, 97, "text"),
            ("STYLE", 98, 98, "text"),
            ("UNIT", 99, 99, "text"),
            ("UNIT_TYPE", 257, 257, "text"),
            ("IN_LENGTH", 100, 107, "num"),
            ("IN_DEPTH", 108, 115, "num"),
            ("IN_HEIGHT", 116, 123, "num"),
            ("OUT_LENGTH", 124, 131, "num"),
            ("OUT_DEPTH", 132, 139, "num"),
            ("OUT_HEIGHT", 140, 147, "num"),
            ("EQV_PACK", 148, 151, "text"),
            ("EQV_CONV", 152, 159, "num"),
            ("GROSS_MASS", 160, 167, "num"),
            ("NET_MASS", 168, 177, "num"),
            ("MIN_MASS", 219, 232, "num"),
            ("MAX_MASS", 233, 246, "num"),
            ("UPP_AIR", 178, 182, "int"),
            ("UPP_SEA", 183, 187, "int"),
            ("CPP_MINI", 214, 218, "int"),
            ("M2_FACTOR", 188, 197, "num"),
            ("M3_FACTOR", 198, 207, "num"),
            ("SIZE_TABLE", 211, 213, "text"),
        ],
    },
    "VA": {
        "sheet": "Varieties", "title": "Variety", "key": ["COMMODITY", "VARIETY"],
        "fields": [
            ("COMMODITY", 39, 40, "text"),
            ("VARIETY", 34, 36, "text"),
            ("VAR_GRP", 37, 38, "text"),
            ("LONG_DESC", 41, 56, "text"),
            ("MED_DESC", 57, 66, "text"),
            ("SHORT_DESC", 67, 72, "text"),
            ("EXT_DESC", 117, 186, "text"),
            ("SORT_SEQ", 73, 77, "int"),
            ("ACT_VAR", 78, 78, "text"),
            ("SUB_VAR", 79, 79, "text"),
            ("EST_FLAG", 80, 80, "text"),
            ("FDA_CODE", 187, 189, "text"),
        ],
    },
    "VG": {
        "sheet": "Variety Groups", "title": "Variety Group", "key": ["COMMODITY", "VAR_GRP"],
        "fields": [
            ("COMMODITY", 34, 35, "text"),
            ("VAR_GRP", 36, 37, "text"),
            ("LONG_DESC", 38, 53, "text"),
            ("MED_DESC", 54, 63, "text"),
            ("SHORT_DESC", 64, 69, "text"),
            ("SORT_SEQ", 70, 74, "int"),
        ],
    },
    "CO": {
        "sheet": "Commodities", "title": "Commodity", "key": ["COMMODITY"],
        "fields": [
            ("COMMODITY", 34, 35, "text"),
            ("COMM_GRP", 36, 37, "text"),
            ("LONG_DESC", 38, 53, "text"),
            ("MED_DESC", 54, 63, "text"),
            ("SHORT_DESC", 64, 69, "text"),
            ("EXT_DESC", 186, 225, "text"),
            ("SORT_SEQ", 102, 106, "int"),
            ("BOTANY_NAME", 107, 136, "text"),
            ("CUSTOMS_CODE", 137, 146, "text"),
            ("EQV_PACK", 147, 150, "text"),
            ("EST_YEAR1", 153, 156, "text"),
            ("EST_YEAR2", 157, 160, "text"),
            ("EST_YEAR3", 161, 164, "text"),
        ],
    },
    "CG": {
        "sheet": "Commodity Groups", "title": "Commodity Group", "key": ["COMM_GRP"],
        "fields": [
            ("COMM_GRP", 34, 35, "text"),
            ("LONG_DESC", 36, 51, "text"),
            ("MED_DESC", 52, 61, "text"),
            ("SHORT_DESC", 62, 67, "text"),
            ("SORT_SEQ", 68, 72, "int"),
        ],
    },
    "GR": {
        "sheet": "Grades", "title": "Grade", "key": ["GRADE"],
        "fields": [
            ("GRADE", 34, 37, "text"),
            ("LONG_DESC", 38, 53, "text"),
            ("MED_DESC", 54, 63, "text"),
            ("SHORT_DESC", 64, 69, "text"),
            ("SORT_SEQ", 70, 74, "int"),
        ],
    },
    "MK": {
        "sheet": "Marks", "title": "Mark", "key": ["MARK"],
        "fields": [
            ("MARK", 34, 38, "text"),
            ("LONG_DESC", 39, 54, "text"),
            ("MED_DESC", 55, 64, "text"),
            ("GRADE", 70, 73, "text"),
            ("SORT_SEQ", 65, 69, "int"),
        ],
    },
    "NC": {
        "sheet": "Inventory Codes", "title": "Inventory Code", "key": ["INV_CODE"],
        "fields": [
            ("INV_CODE", 34, 35, "text"),
            ("LONG_DESC", 36, 51, "text"),
            ("MED_DESC", 52, 61, "text"),
            ("SHORT_DESC", 62, 67, "text"),
            ("SORT_SEQ", 68, 72, "int"),
        ],
    },
    "TK": {
        "sheet": "Target Markets", "title": "Target Market", "key": ["TARG_MKT"],
        "fields": [
            ("TARG_MKT", 34, 35, "text"),
            ("LONG_DESC", 36, 51, "text"),
            ("MED_DESC", 52, 61, "text"),
            ("SHORT_DESC", 62, 67, "text"),
            ("TARGET_COUNTRY", 69, 70, "text"),
            ("TARGET_REGION", 71, 73, "text"),
            ("SELL_BY_CODE_YN", 68, 68, "text"),
        ],
    },
    "SI": {
        "sheet": "Sizes & Counts", "title": "Size / Count", "lang": False,
        "key": ["COMMODITY", "PACK", "SEASON", "SIZE_COUNT"],
        "fields": [
            ("COMMODITY", 34, 35, "text"),
            ("PACK", 36, 39, "text"),
            ("SEASON", 40, 43, "text"),
            ("SIZE_COUNT", 44, 48, "text"),
            ("SORT_SEQ", 49, 53, "int"),
            ("SIZEREF", 54, 62, "text"),
            ("DIAMETERA", 63, 71, "text"),
        ],
    },
    "PE": {
        "sheet": "Periods", "title": "Period", "lang": False,
        "key": ["SEASON", "PERIOD"],
        "fields": [
            ("SEASON", 34, 37, "text"),
            ("PERIOD", 38, 43, "text"),
            ("START_DATE", 44, 51, "date"),
            ("END_DATE", 52, 59, "date"),
        ],
    },
}

# Look-ups: add columns to table X taken from another table T.
#   (insert after column, new column, lookup table, [key columns in X], field in T)
# Rules are applied in order, so a later rule can use a column added earlier.
def _desc(col, target, keys):
    return (col, col + "_DESC", target, keys, "LONG_DESC")

LOOKUPS = {
    "GT": [_desc("COMMODITY", "CO", ["COMMODITY"]),
           ("VARIETY", "VARIETY_NAME", "VA", ["COMMODITY", "VARIETY"], "LONG_DESC"),
           ("VARIETY_NAME", "VAR_GRP", "VA", ["COMMODITY", "VARIETY"], "VAR_GRP"),
           ("VAR_GRP", "VAR_GRP_NAME", "VG", ["COMMODITY", "VAR_GRP"], "LONG_DESC"),
           _desc("PACK", "PK", ["COMMODITY", "PACK"]),
           _desc("GRADE", "GR", ["GRADE"]),
           _desc("MARK", "MK", ["MARK"]),
           _desc("INV_CODE", "NC", ["INV_CODE"]),
           _desc("COMM_GRP", "CG", ["COMM_GRP"]),
           _desc("DEFLT_TARG_MKT", "TK", ["DEFLT_TARG_MKT"])],
    "VA": [_desc("COMMODITY", "CO", ["COMMODITY"]),
           _desc("VAR_GRP", "VG", ["COMMODITY", "VAR_GRP"])],
    "VG": [_desc("COMMODITY", "CO", ["COMMODITY"])],
    "CO": [_desc("COMM_GRP", "CG", ["COMM_GRP"])],
    "PK": [_desc("COMMODITY", "CO", ["COMMODITY"])],
    "SI": [_desc("COMMODITY", "CO", ["COMMODITY"]),
           _desc("PACK", "PK", ["COMMODITY", "PACK"])],
    "MK": [_desc("GRADE", "GR", ["GRADE"])],
}

# Tabs with a fixed, approved column list: (heading shown in Excel, source column).
# Tabs not listed here show every field.
VIEWS = {
    "GT": [("GTIN", "GTIN"),
           ("Orgzn", "ORGZN"),
           ("Commodity", "COMMODITY"),
           ("Variety", "VARIETY"),
           ("Variety Name", "VARIETY_NAME"),
           ("Variety Group", "VAR_GRP"),
           ("Variety Group Name", "VAR_GRP_NAME"),
           ("Target Market", "TARGET_MARKETS"),
           ("Pack", "PACK"),
           ("Grade", "GRADE"),
           ("Mark", "MARK"),
           ("Inventory Code", "INV_CODE"),
           ("Size Count", "SIZE_COUNT"),
           ("Country", "COUNTRY"),
           ("Active", "ACTIVE")],
}

# Paltrack "R2 Gtin Export" spreadsheet: Excel heading -> GTIN field.
# A GTIN linked to several target markets appears on several rows; the
# Target Market column on the GTINs tab lists all of them ("EU, TE").
R2_GTIN_COLUMNS = {
    "GTIN": "GTIN", "Orgzn": "ORGZN", "Commodity": "COMMODITY", "Variety": "VARIETY",
    "Pack": "PACK", "Grade": "GRADE", "Mark": "MARK", "Inventory Code": "INV_CODE",
    "Size Count": "SIZE_COUNT", "Start Date": "START_DATE", "End Date": "END_DATE",
    "Season": "SEASON", "Country": "COUNTRY", "Active": "ACTIVE",
    "Tran User": "TRAN_USER", "Tran Date": "TRAN_DATETIME",
}

STANDARD_COLS = ["ACTIVE", "ORGZN", "LANG", "TRAN_USER", "TRAN_DATETIME", "BATCH", "SOURCE_FILE"]
SHEET_ORDER = ["GT", "FA", "PK", "VA", "VG", "CO", "CG", "GR", "MK", "NC", "TK", "SI", "PE"]


# --------------------------------------------------------------------------
# PARSING
# --------------------------------------------------------------------------
def cut(line, a, b):
    return line[a - 1:b].strip()


def convert(value, kind):
    if value == "":
        return None
    try:
        if kind == "int":
            return int(float(value))
        if kind == "num":
            return float(value)
        if kind == "date":
            return datetime.strptime(value, "%Y%m%d").date()
    except ValueError:
        return value          # keep the raw text rather than lose it
    return value


def build_record(code, line, batch, source):
    spec = TABLES[code]
    rec = {}
    for name, a, b, kind in spec["fields"]:
        rec[name] = convert(cut(line, a, b), kind)
    o1, o2 = spec.get("orgzn", (32, 33))
    rec["ACTIVE"] = cut(line, 31, 31)
    rec["ORGZN"] = cut(line, o1, o2)
    rec["LANG"] = cut(line, 6, 6) if spec.get("lang", True) else None
    rec["TRAN_USER"] = cut(line, 8, 14)
    d, t = cut(line, 15, 22), cut(line, 23, 30)
    try:
        rec["TRAN_DATETIME"] = datetime.strptime(d + t, "%Y%m%d%H:%M:%S")
    except ValueError:
        rec["TRAN_DATETIME"] = (d + " " + t).strip() or None
    rec["BATCH"] = batch
    rec["SOURCE_FILE"] = source
    return rec


def read_lines(path):
    raw = open(path, "rb").read().decode(FILE_ENCODING, errors="replace")
    # Split on line feeds only. Stray carriage returns inside a field are
    # replaced with a space so the character positions stay correct.
    return [l.replace("\r", " ") for l in raw.split("\n")]


def natural_key(text):
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r"(\d+)", text)]


def find_batches(root):
    folders = [d for d in os.listdir(root)
               if os.path.isdir(os.path.join(root, d)) and BATCH_FOLDER_PATTERN.match(d)]
    return sorted(folders, key=natural_key)


def add_record(data, stats, code, rec, order):
    """Add a record, keeping the newest version when the same key already exists."""
    spec = TABLES[code]
    rec["_order"] = order
    key = tuple(rec.get(k) for k in spec["key"]) + (rec["ORGZN"], rec["LANG"])
    stats[code]["read"] += 1
    old = data[code].get(key)
    if old is None:
        data[code][key] = rec
        return
    stats[code]["replaced"] += 1
    newer, older = (rec, old) if sort_stamp(rec) >= sort_stamp(old) else (old, rec)
    if newer.get("_partial"):
        # An Excel export doesn't carry every Paltrack field: keep the ones it lacks
        merged = dict(older)
        merged.update({k: v for k, v in newer.items() if v not in (None, "")})
        newer = merged
    data[code][key] = newer


def load_fixed_width(path, code, batch, fname, order, data, stats, log):
    skipped = 0
    for line in read_lines(path):
        if len(line.rstrip()) < 31:         # blank lines and the "BT" trailer
            if line.strip():
                skipped += 1
            continue
        add_record(data, stats, code, build_record(code, line, batch, fname), order)
    if skipped:
        log.append(f"{batch}/{fname}: {skipped} short/trailer line(s) ignored")


def as_text(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    v = str(v).strip()
    return v or None


def as_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    v = as_text(v)
    if not v:
        return None
    for fmt in ("%Y%m%d", "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(v[:10], fmt).date()
        except ValueError:
            pass
    return v


def load_excel(path, batch, fname, order, data, stats, links, log):
    """Reads a Paltrack R2 Gtin Export spreadsheet into the GTINs table."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        log.append(f"{batch}/{fname}: could not open Excel file ({e}) - skipped")
        return
    used = False
    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if not header:
            continue
        head = [as_text(h) or "" for h in header]
        if "GTIN" not in head or "Target Market" not in head:
            continue
        used = True
        col = {h: i for i, h in enumerate(head)}
        tm_i = col.get("Target Market")
        tm_active_i = col.get("Target Market Active")
        stats["GT"]["files"].append(f"{batch}/{fname}")
        seen = set()
        for row in rows:
            row = list(row) + [None] * (len(head) - len(row))
            gtin = as_text(row[col["GTIN"]])
            if not gtin:
                continue
            gtin = gtin.zfill(14)
            rec = {f: as_text(row[col[h]]) for h, f in R2_GTIN_COLUMNS.items() if h in col}
            rec["GTIN"] = gtin
            rec["START_DATE"] = as_date(row[col["Start Date"]]) if "Start Date" in col else None
            rec["END_DATE"] = as_date(row[col["End Date"]]) if "End Date" in col else None
            t = row[col["Tran Date"]] if "Tran Date" in col else None
            rec["TRAN_DATETIME"] = t if isinstance(t, datetime) else as_text(t)
            rec["LANG"] = None
            rec["BATCH"], rec["SOURCE_FILE"], rec["_partial"] = batch, fname, True
            tm = as_text(row[tm_i]) if tm_i is not None else None
            tm_ok = tm_active_i is None or as_text(row[tm_active_i]) in (None, "Y")
            if tm and tm_ok:
                prev = links.get(gtin)
                if prev is None or prev[0] < order:
                    links[gtin] = (order, set())
                links[gtin][1].add(tm)
            k = (gtin, rec["START_DATE"])
            if k in seen:                   # extra target-market row for the same GTIN
                continue
            seen.add(k)
            add_record(data, stats, "GT", rec, order)
    wb.close()
    if not used:
        log.append(f"{batch}/{fname}: Excel file is not an R2 GTIN export "
                   f"(no 'GTIN' and 'Target Market' headings) - skipped")


def load_all(root, log):
    batches = find_batches(root)
    if not batches:
        sys.exit(f"No 'Batch ...' folders found in {root}")
    data = {code: {} for code in TABLES}      # code -> {key: record}
    stats = {code: {"read": 0, "replaced": 0, "files": []} for code in TABLES}
    links = {}                                # GTIN -> (batch order, {target markets})

    for order, batch in enumerate(batches):
        folder = os.path.join(root, batch)
        for fname in sorted(os.listdir(folder)):
            path = os.path.join(folder, fname)
            low = fname.lower()
            if os.path.isdir(path) or fname.startswith("~$") or \
                    low.endswith((".pdf", ".zip", ".txt", ".py", ".png", ".jpg", ".docx", ".msg")):
                continue
            if low.endswith((".xlsx", ".xlsm")):
                load_excel(path, batch, fname, order, data, stats, links, log)
                continue
            if low.endswith((".xls", ".csv")):
                log.append(f"{batch}/{fname}: save it as .xlsx and run again - skipped")
                continue
            code = fname[:2].upper()
            if code not in TABLES:
                log.append(f"{batch}/{fname}: unknown table prefix '{code}' - skipped")
                continue
            stats[code]["files"].append(f"{batch}/{fname}")
            load_fixed_width(path, code, batch, fname, order, data, stats, log)

    # Target Market on the GTINs tab: all linked markets if an R2 export gave
    # them, otherwise the GTIN's default target market.
    for rec in data["GT"].values():
        linked = links.get(rec["GTIN"])
        rec["TARGET_MARKETS"] = ", ".join(sorted(linked[1])) if linked else rec.get("DEFLT_TARG_MKT")
    for code in TABLES:
        if not stats[code]["files"]:
            log.append(f"No file found for {TABLES[code]['title']} ({code}) in any batch")
    return batches, data, stats


def sort_stamp(rec):
    t = rec["TRAN_DATETIME"]
    return (t if isinstance(t, datetime) else datetime.min, rec["_order"])


# --------------------------------------------------------------------------
# LOOK-UPS
# --------------------------------------------------------------------------
def build_lookup_index(data, code, key_cols, field):
    """Map lookup key -> field value, preferring active, then English, then newest."""
    index, rank = {}, {}
    for rec in data[code].values():
        k = tuple(rec.get(c) for c in key_cols)
        r = (rec["ACTIVE"] == "Y", rec["LANG"] == "L", sort_stamp(rec))
        if k not in rank or r > rank[k]:
            rank[k], index[k] = r, rec.get(field)
    return index


def apply_lookups(data, log):
    """Fills look-up columns. Returns a list of data issues (codes with no match)."""
    issues, done = [], set()
    for code, rules in LOOKUPS.items():
        for after, new_col, target, key_cols, field in rules:
            index = build_lookup_index(data, target, TABLES[target]["key"], field)
            missing = set()
            for rec in data[code].values():
                k = tuple(rec.get(c) for c in key_cols)
                if any(v in (None, "") for v in k):
                    rec[new_col] = None
                    continue
                rec[new_col] = index.get(k)
                if k in index:
                    continue
                missing.add(k)
                rec_id = " / ".join(str(rec.get(c)) for c in TABLES[code]["key"] if rec.get(c) is not None)
                tag = (code, rec_id, tuple(key_cols), target)
                if tag not in done:
                    done.add(tag)
                    issues.append((TABLES[code]["sheet"], rec_id, rec.get("ACTIVE"), key_cols[-1],
                                   rec.get(key_cols[-1]), f"Not in {TABLES[target]['sheet']}",
                                   rec.get("BATCH"), rec.get("SOURCE_FILE")))
            if missing and (code, tuple(key_cols), target) not in done:
                done.add((code, tuple(key_cols), target))
                shown = ", ".join("/".join(str(x) for x in k) for k in sorted(missing, key=str)[:10])
                log.append(f"{TABLES[code]['sheet']}: {len(missing)} {'/'.join(key_cols)} code(s) not found in "
                           f"{TABLES[target]['sheet']}: {shown} - see Data Issues tab")
    return issues


def column_order(code):
    """Returns [(heading, source column)] for a tab."""
    if code in VIEWS:
        return VIEWS[code]
    cols = [name for name, *_ in TABLES[code]["fields"]]
    for after, new_col, *_ in LOOKUPS.get(code, []):
        cols.insert(cols.index(after) + 1, new_col)
    std = [c for c in STANDARD_COLS if not (c == "LANG" and not TABLES[code].get("lang", True))]
    return [(c, c) for c in cols + std]


# --------------------------------------------------------------------------
# EXCEL OUTPUT
# --------------------------------------------------------------------------
def write_workbook(path, batches, data, stats, log, issues=()):
    wb = xlsxwriter.Workbook(path, {"strings_to_numbers": False, "strings_to_urls": False})
    f_title = wb.add_format({"bold": True, "font_size": 16})
    f_sub = wb.add_format({"italic": True, "font_color": "#595959"})
    f_head = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white", "border": 1})
    f_link = wb.add_format({"font_color": "#0563C1", "underline": 1})
    f_date = wb.add_format({"num_format": "yyyy-mm-dd"})
    f_dt = wb.add_format({"num_format": "yyyy-mm-dd hh:mm"})
    f_num = wb.add_format({"num_format": "#,##0.000"})
    f_int = wb.add_format({"num_format": "0"})
    f_text = wb.add_format({"num_format": "@"})

    index = wb.add_worksheet("Index")
    issues_ws = wb.add_worksheet("Data Issues")
    sheets_meta = []

    for code in SHEET_ORDER:
        spec = TABLES[code]
        cols = column_order(code)
        kinds = {n: k for n, _, _, k in spec["fields"]}
        kinds["TRAN_DATETIME"] = "datetime"
        rows = sorted(data[code].values(),
                      key=lambda r: tuple("" if r.get(k) is None else str(r.get(k)) for k in spec["key"]))
        ws = wb.add_worksheet(spec["sheet"])
        sheets_meta.append((code, spec, len(rows),
                            sum(1 for r in rows if r["ACTIVE"] == "Y")))

        for r_i, rec in enumerate(rows, start=1):
            for c_i, (_, col) in enumerate(cols):
                v = rec.get(col)
                if v is None:
                    continue
                k = kinds.get(col, "text")
                if isinstance(v, datetime):
                    ws.write_datetime(r_i, c_i, v, f_dt)
                elif isinstance(v, date):
                    ws.write_datetime(r_i, c_i, datetime(v.year, v.month, v.day), f_date)
                elif k == "num" and isinstance(v, float):
                    ws.write_number(r_i, c_i, v, f_num)
                elif k == "int" and isinstance(v, int):
                    ws.write_number(r_i, c_i, v, f_int)
                else:
                    ws.write_string(r_i, c_i, str(v))

        last_row = max(len(rows), 1)
        ws.add_table(0, 0, last_row, len(cols) - 1, {
            "name": "tbl_" + re.sub(r"\W", "", spec["sheet"]),
            "style": "Table Style Medium 2",
            "columns": [{"header": h, "header_format": f_head} for h, _ in cols],
        })
        n_key = len(spec["key"]) if code != "GT" else 1
        ws.freeze_panes(1, n_key)
        for c_i, (head, col) in enumerate(cols):
            sample = [len(str(r.get(col))) for r in rows[:3000] if r.get(col) is not None]
            width = min(max([len(head) + 4] + sample) + 2, 45)
            if kinds.get(col) == "datetime":
                width = 17
            elif kinds.get(col) == "date":
                width = 12
            ws.set_column(c_i, c_i, width, f_text if kinds.get(col, "text") == "text" else None)
        ws.write_url(0, len(cols) + 1, "internal:'Index'!A1", f_link, string="<< Back to Index")

    # ---- Data Issues sheet ----
    i_heads = ["Tab", "Record", "Active", "Field", "Code used", "Problem", "Batch", "Source file"]
    issues = sorted(issues, key=lambda x: (x[0], x[3], str(x[4]), x[1]))
    for r_i, row in enumerate(issues, start=1):
        for c_i, v in enumerate(row):
            if v is not None:
                issues_ws.write_string(r_i, c_i, str(v))
    issues_ws.add_table(0, 0, max(len(issues), 1), len(i_heads) - 1, {
        "name": "tbl_DataIssues", "style": "Table Style Medium 3",
        "columns": [{"header": h, "header_format": f_head} for h in i_heads]})
    issues_ws.freeze_panes(1, 0)
    for c_i, w in enumerate([16, 26, 8, 14, 12, 26, 10, 18]):
        issues_ws.set_column(c_i, c_i, w)
    issues_ws.write_url(0, len(i_heads) + 1, "internal:'Index'!A1", f_link, string="<< Back to Index")
    if not issues:
        issues_ws.write(1, 0, "No issues found.")

    # ---- Index sheet ----
    index.write(0, 0, "Paltrack Masterfile - Combined", f_title)
    index.write(1, 0, f"Built {datetime.now():%Y-%m-%d %H:%M} from: {', '.join(batches)}", f_sub)
    index.write(2, 0, "Click a table name to jump to it. Every tab has filter buttons on the header row.", f_sub)
    heads = ["Table", "Paltrack code", "Records", "Active", "Inactive",
             "Rows read", "Duplicates merged", "Source files"]
    for c, h in enumerate(heads):
        index.write(4, c, h, f_head)
    for r, (code, spec, n, active) in enumerate(sheets_meta, start=5):
        index.write_url(r, 0, f"internal:'{spec['sheet']}'!A1", f_link, string=spec["sheet"])
        index.write(r, 1, code)
        index.write_number(r, 2, n)
        index.write_number(r, 3, active)
        index.write_number(r, 4, n - active)
        index.write_number(r, 5, stats[code]["read"])
        index.write_number(r, 6, stats[code]["replaced"])
        index.write(r, 7, "; ".join(stats[code]["files"]))
    r = 5 + len(sheets_meta)
    index.write_url(r, 0, "internal:'Data Issues'!A1", f_link, string="Data Issues")
    index.write(r, 2, f"{len(issues)} record(s) use a code that is missing from its master table")
    r += 2
    index.write(r, 0, "Load log", f_head)
    for i, msg in enumerate(log or ["No issues."], start=1):
        index.write(r + i, 0, msg)
    index.set_column(0, 0, 22)
    index.set_column(1, 6, 14)
    index.set_column(7, 7, 60)
    index.activate()
    wb.close()


def main():
    log = []
    print(f"Reading batches in {MASTER_ROOT} ...")
    batches, data, stats = load_all(MASTER_ROOT, log)
    print("Batches:", ", ".join(batches))
    issues = apply_lookups(data, log)
    for code in SHEET_ORDER:
        print(f"  {TABLES[code]['sheet']:<18} {len(data[code]):>7,} records")
    try:
        write_workbook(OUTPUT_FILE, batches, data, stats, log, issues)
    except PermissionError:
        sys.exit(f"Cannot write {OUTPUT_FILE} - close it in Excel and run again.")
    for msg in log:
        print("  note:", msg)
    print(f"\nDone. Workbook saved to:\n{OUTPUT_FILE}")


if __name__ == "__main__":
    if len(sys.argv) > 1:                 # optional: python paltrack_masterfile_builder.py <root folder>
        MASTER_ROOT = sys.argv[1]
        OUTPUT_FILE = os.path.join(MASTER_ROOT, "Masterfile_Combined.xlsx")
    main()
