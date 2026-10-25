"""Run with ``python -m demo [--output NEW_DIRECTORY]``."""

import argparse
from copy import deepcopy
import json
from pathlib import Path

from .costing import calculate, load_fixtures
from .workbook import create_fixture, edit_copy


def run(output=None):
    request, catalog = load_fixtures()
    baseline = calculate(request, catalog)
    missing = deepcopy(catalog)
    missing["records"] = [record for record in missing["records"] if record["resource"] != "plywood_12mm"]
    repriced = deepcopy(catalog)
    repriced["records"][0]["price"] = "13000"
    changed = calculate(request, repriced)
    report = {
        "classification": "SYNTHETIC", "purpose": "DEMO_ONLY", "implementation": "PUBLIC_REDUCED_DEMO",
        "notice": "Invented prices and coefficients; not the private system or official construction rates.",
        "complete_case": baseline, "missing_price_case": calculate(request, missing),
        "price_change_case": changed,
        "validation_scope": "Synthetic Decimal costing and restricted OOXML links; no external API or native Excel used.",
    }
    if output is not None:
        folder = Path(output)
        if folder.exists():
            raise FileExistsError("Use a new output directory to preserve previous runs")
        folder.mkdir(parents=True)
        source = folder / "synthetic_original.xlsx"
        revision = folder / "synthetic_repriced.xlsx"
        original_hash = create_fixture(source, request, catalog)
        report["workbook_verification"] = edit_copy(
            source, revision, {("Prices", "B2"): "13000"}, original_hash, changed)
        report["outputs"] = ["synthetic_original.xlsx", "synthetic_repriced.xlsx", "result.json"]
        (folder / "result.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", metavar="NEW_DIRECTORY", help="Write synthetic XLSX originals/copies and a verification JSON")
    args = parser.parse_args()
    try:
        result = run(args.output)
    except (ValueError, OSError) as error:
        parser.exit(2, f"demo: {error}\n")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
