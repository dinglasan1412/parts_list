#!/usr/bin/env python3
"""Search the Material Usage Finder database from the command line.

    python tools/search.py solder
    python tools/search.py hv998 --product MRS
    python tools/search.py oring --model VSM --process Assembly --summary
"""
import argparse
import json
import re
from collections import defaultdict
from pathlib import Path

FIELDS = ("product", "model", "category", "name", "code", "qty", "unit", "supplier",
          "remarks", "process", "guess", "superseded", "doc")
# Different files name the same material differently.
ALIASES = {"hardener": "adhesive", "mainresin": "adhesive", "moleepoxy": "moldepoxy"}


def load(path):
    """data.js is JSON behind a 'window.SEED=' prefix."""
    text = Path(path).read_text(encoding="utf-8")
    data = json.loads(text[text.index("{"): text.rindex("}") + 1])
    return data["products"], [dict(zip(FIELDS, r)) for r in data["recs"]]


def group_key(name):
    """Same grouping as the website: ignore brackets, numbers and wording variants."""
    key = re.sub(r"\(.*?\)|\d+$|nitrile |hand |dispenser |anti-static |finger ?cots?",
                 lambda m: "fingercots" if "finger" in m.group(0) else "", name.lower())
    key = re.sub(r"[^a-z]", "", key) or name.lower()
    return ALIASES.get(key, key)


def matches(rec, terms, args, products):
    haystack = " ".join([rec["name"], rec["code"], rec["remarks"], rec["process"], rec["product"],
                         rec["model"], rec["supplier"], rec["doc"], products[rec["product"]]["name"]]).lower()
    return (all(t in haystack for t in terms)
            and (not args.product or rec["product"].lower() == args.product.lower())
            and (not args.model or rec["model"].lower() == args.model.lower())
            and (not args.process or args.process.lower() in rec["process"].lower())
            and (not args.category or args.category.lower() in rec["category"].lower()))


def main():
    parser = argparse.ArgumentParser(description="Search the material database.")
    parser.add_argument("query", nargs="*", help="words that must all appear (name, part number, remarks...)")
    parser.add_argument("--product")
    parser.add_argument("--model")
    parser.add_argument("--process")
    parser.add_argument("--category", help="Direct, Indirect or Factory")
    parser.add_argument("--summary", action="store_true", help="only show the common usages")
    parser.add_argument("--data", default="data.js")
    args = parser.parse_args()

    products, records = load(args.data)
    groups = defaultdict(list)
    for rec in records:
        if matches(rec, [w.lower() for w in args.query], args, products):
            groups[group_key(rec["name"])].append(rec)
    if not groups:
        print("No match.")
        return
    for rows in sorted(groups.values(), key=lambda r: r[0]["name"].lower()):
        models = {(r["product"], r["model"]) for r in rows}
        by_process = defaultdict(set)
        for r in rows:
            by_process[r["process"]].add((r["product"], r["model"]))
        print(f"\n{rows[0]['name']}  -  {len(models)} model(s), {len(by_process)} process(es)")
        for process, found in sorted(by_process.items(), key=lambda x: -len(x[1])):
            print(f"  {process}: {len(found)} model(s)")
        if args.summary:
            continue
        for r in sorted(rows, key=lambda r: (r["product"], r["model"], r["process"])):
            qty = " ".join(x for x in (r["qty"], r["unit"]) if x) or "-"
            tag = " [superseded]" if r["superseded"] else ""
            print(f"    {r['product']} {r['model']:<9} {r['process']:<26} {r['code'] or '-'}{tag}  ({qty})")


if __name__ == "__main__":
    main()
