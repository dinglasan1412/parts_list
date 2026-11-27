#!/usr/bin/env python3
"""Build data.js for the Material Usage Finder from Excel parts lists.

    python tools/import_parts_lists.py --input parts_lists --output data.js

The parts lists (.xlsx) are described in tools/parts_lists.json: which file is
which product and model. From every workbook the newest revision sheet is read.
For products marked "track_superseded", parts that the newest revision replaced
with a different part number are kept too and tagged as superseded.
Processes are guessed from the material name (PROCESS_RULES below) unless
tools/process_overrides.json says otherwise.
"""
import argparse
import datetime
import json
import re
from collections import Counter
from pathlib import Path

import openpyxl

# The first matching rule wins. A direct material that matches no rule is taken
# to be an assembly component; any other material falls back to general support.
PROCESS_RULES = [
    ('Coil Winding', r"copper wire"),
    ('Soldering', r"solder|flux|diluent|thermometer"),
    ('Wire Bonding', r"^wire"),
    ('Molding / Encapsulation', r"mold epoxy|mole epoxy"),
    ('Lapping / Polishing', r"lapping"),
    ('Leak Testing', r"helium"),
    ('Adhesive Dispensing', r"adhesive|hardener|resin|loctite|barrel|tapered|receiver|piston|o\W+ring|syringe|nozzle|dispenser|chemiseal|nicbond|bonding|rubber bond"),
    ('Packing', r"plastic bag|embossed|reel"),
    ('Cleaning', r"soap|alcohol|ipa|methanol|swab|cleanroom|cloth|chamois|wiper|statinol|tech ?clean"),
    ('Packing', r"^tray$|cushion|layer pad|carton|box|tape|zipper|bubble|hi-wrap|wrapping|packaging|silica|rag paper|shipment tray"),
    ('Marking & Documentation', r"marker|sharpie|sign pen|stamp|^file"),
    ('Assembly', r"case|magnet|pcb|spacer|lead frame|chip|seal|plunger|bane|orifice|gaitho|guide|coil assembl|coil assy|ferrite|ppw|assembly wire|dummy wire|die cast|main lead|ground lead|cover|sokoita|futa|screw"),
]
FALLBACK_PROCESS = "Handling & General Support"
DATE_FORMATS = ("%d.%b.%y", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y", "%d-%b-%Y", "%B %d, %Y")
SECTION_RE = re.compile(r"(Direct Materials|Indirect Materials|Factory Supplies)")
SECTION_NAMES = {"Direct": "Direct", "Indirect": "Indirect", "Factory": "Factory Supply"}
HEADERS = {"No.": "no", "Supplier": "supplier", "Remarks": "remarks"}


def text(value):
    """Cell text with whitespace collapsed."""
    return re.sub(r"\s+", " ", str(value)).strip() if value is not None else ""


def item_text(value):
    """Cell text of a material row: numbers are shortened, dashes mean blank."""
    out = f"{value:.6g}" if isinstance(value, (int, float)) else text(value)
    return "" if re.fullmatch(r"-*", out) else out


def blank_if_dash(value):
    return "" if re.fullmatch(r"-*", value) else value


def format_date(value):
    raw = text(value)
    for fmt in DATE_FORMATS:
        try:
            return datetime.datetime.strptime(raw.title(), fmt).strftime("%d %b %Y").lstrip("0")
        except ValueError:
            pass
    return raw


def guess_process(name, section, overrides):
    """Return (process, is_guess)."""
    if name.lower() in overrides:
        return overrides[name.lower()], 0
    for process, pattern in PROCESS_RULES:
        if re.search(pattern, name.lower()):
            return process, 1
    return ("Assembly" if section == "Direct" else FALLBACK_PROCESS), 1


def read_rows(sheet):
    """Every row as {column: value}, keeping only the cells that have content."""
    return [{c.column: c.value for c in row if c.value is not None and text(c.value)}
            for row in sheet.iter_rows()]


def row_below(rows, label):
    """The row right under the first row that has a cell equal to `label`."""
    for i, row in enumerate(rows):
        if any(text(v) == label for v in row.values()):
            return rows[i + 1]
    raise ValueError(f"'{label}' not found in the sheet")


def read_document(rows):
    """Parts list number, revision, description, customer and last change."""
    labels_row = next(r for r in rows if "NIPHCO Part No." in map(text, r.values()))
    labels = [text(v) for _, v in sorted(labels_row.items())]
    values = [text(v) for _, v in sorted(row_below(rows, "NIPHCO Part No.").items())]
    info = dict(zip(labels, values))
    number = re.sub(r"PL-\s+", "PL-", text(row_below(rows, "Parts List No.:")[1]))
    issued = format_date(list(row_below(rows, "Issue Date:").values())[0])

    # Revision table: the last numbered row is the newest revision.
    start = next(i for i, r in enumerate(rows)
                 if text(r.get(1)) == "Rev" and "Nature of Change" in map(text, r.values()))
    latest = None
    for row in rows[start + 1:]:
        first = text(row.get(1))
        if re.fullmatch(r"\d+", first):
            latest = row
        elif first:
            break
    change = text(latest.get(4))
    reference = text(list(latest.values())[-1])
    if reference and reference != "---" and reference != change and len(reference) > 2:
        change += f" ({reference.replace('As per 5M1ECA No.: ', '5M1ECA ')})"
    return {
        "pl": number,
        "rev": "Rev " + str(int(text(latest[1]))),
        "date": issued,
        "desc": info.get("Description:", ""),
        "pn": blank_if_dash(info.get("NIPHCO Part No.", "")),
        "cpn": blank_if_dash(info.get("Customer P/N:", "")),
        "cust": blank_if_dash(info.get("Customer:", "")),
        "chg": change,
    }


def column_map(header_row):
    """Which column holds the name, part number, quantity... of a material table."""
    cols = {}
    for col, value in header_row.items():
        label = text(value)
        if label in HEADERS:
            cols[HEADERS[label]] = col
        elif label.startswith(("Direct", "Indirect", "Factory")):
            cols["name"] = col
        elif label.startswith("Part No"):
            cols["code"] = col
        elif label.startswith("Qty"):
            cols["qty"] = col
        elif label.startswith("Unit"):
            cols["unit"] = col
    return cols


def read_items(rows):
    """(category, name, part no., qty, unit, supplier, remarks) for every material."""
    items, section, cols, last_name = [], None, {}, ""
    for row in rows:
        line = " ".join(text(v) for v in row.values())
        match = SECTION_RE.search(line)
        if match and any(text(v) == "No." or text(v).startswith("Part No") for v in row.values()):
            section = SECTION_NAMES[match.group(1).split()[0]]
            cols = column_map(row)
            continue
        if line.startswith(("Reference Doc", "Rev ")):
            section = None
        if not section:
            continue

        def cell(key):
            return item_text(row.get(cols[key])) if key in cols else ""

        name, code = cell("name"), cell("code")
        if not name and not code:
            continue
        name = name or last_name  # second line of a two-part material (Part A / Part B)
        last_name = name
        if name.startswith("Joy"):
            name = "Joy Soap"
        items.append((section, name, code, cell("qty"), cell("unit"), cell("supplier"), cell("remarks")))
    return items


def find_superseded(old_items, new_items):
    """Old parts that the newest revision replaced with a different part number."""
    old_count = Counter(i[1] for i in old_items)
    new_count = Counter(i[1] for i in new_items)
    new_code = {i[1]: i[2] for i in new_items}
    new_codes = {i[2] for i in new_items}
    for item in old_items:
        name, code = item[1], item[2]
        if old_count[name] == 1 == new_count[name] and code and code != new_code[name] and code not in new_codes:
            yield item, new_code[name]


def build(config, input_dir, overrides):
    products = {code: {"name": p["name"], "models": {m: [] for m in p["models"]}}
                for code, p in config["products"].items()}
    records = []
    for entry in config["files"]:
        path = input_dir / entry["file"]
        if not path.exists():
            print(f"  missing: {entry['file']}")
            continue
        product, model = entry["product"], entry["model"]
        book = openpyxl.load_workbook(path, data_only=True)
        names = book.sheetnames
        sheet = entry.get("sheet") or names[-1]
        rows = read_rows(book[sheet])
        doc, items = read_document(rows), read_items(rows)
        label = f"{doc['pl']} · {doc['rev']}"
        products[product]["models"].setdefault(model, []).append(doc)
        for section, name, code, qty, unit, supplier, remarks in items:
            process, guessed = guess_process(name, section, overrides)
            records.append([product, model, section, name, code, qty, unit, supplier, remarks,
                            process, guessed, 0, label])
        if config["products"][product].get("track_superseded") and names.index(sheet) > 0:
            old_rows = read_rows(book[names[names.index(sheet) - 1]])
            old_doc, old_items = read_document(old_rows), read_items(old_rows)
            for item, new_code in find_superseded(old_items, items):
                section, name, code, qty, unit, supplier, _ = item
                process, guessed = guess_process(name, section, overrides)
                records.append([product, model, section, name, code, qty, unit, supplier,
                                f"Superseded by {new_code} ({doc['rev']})", process, guessed, 1,
                                f"{old_doc['pl']} · {old_doc['rev']}"])
        print(f"  {product} {model}: {len(items)} materials from {label}")
    return {"products": products, "recs": records}


def main():
    here = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description="Build data.js from Excel parts lists.")
    parser.add_argument("--input", default="parts_lists", help="folder with the .xlsx parts lists")
    parser.add_argument("--config", default=here / "parts_lists.json")
    parser.add_argument("--overrides", default=here / "process_overrides.json")
    parser.add_argument("--output", default="data.js")
    args = parser.parse_args()

    config = json.loads(Path(args.config).read_text(encoding="utf-8"))
    overrides = {k.lower(): v for k, v in json.loads(Path(args.overrides).read_text(encoding="utf-8")).items()}
    input_dir = Path(args.input)
    print(f"Reading parts lists from {input_dir}/")
    data = build(config, input_dir, overrides)

    listed = {e["file"] for e in config["files"]}
    for path in sorted(input_dir.glob("*.xlsx")):
        if path.name not in listed:
            print(f"  skipped (not in parts_lists.json): {path.name}")

    header = ("// Generated by tools/import_parts_lists.py. Do not edit by hand.\n"
              "// products: product > model > parts list documents.\n"
              "// recs: [product, model, category, material, part no., qty, unit, supplier, remarks, "
              "process, process-is-a-guess, superseded, parts list]\n")
    Path(args.output).write_text(header + "window.SEED=" + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";\n",
                                 encoding="utf-8")
    print(f"Wrote {args.output}: {len(data['recs'])} records, {sum(len(p['models']) for p in data['products'].values())} models")


if __name__ == "__main__":
    main()
