"""A deliberately small OOXML editor for this generated synthetic fixture only.

It preserves the original bytes, untouched ZIP parts, styles and formula text.
It supports this fixture's arithmetic, SUM and TRUNC only, rejects other workbook
shapes, and compares cached links with a separate Decimal calculation. This is
not a general Excel recalculation engine or native Excel verification.
"""

import ast
from decimal import Decimal, localcontext
import hashlib
from pathlib import Path
import re
import tempfile
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

from .costing import decimal, trunc

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SHEETS = {"Prices": "xl/worksheets/sheet1.xml", "Estimate": "xl/worksheets/sheet2.xml",
          "Summary": "xl/worksheets/sheet3.xml"}
ALLOWED_INPUTS = {("Prices", "B2"), ("Prices", "B3"), ("Prices", "B4"), ("Estimate", "D2")}
ET.register_namespace("", NS)


class WorkbookError(ValueError):
    """Unsupported structure, changed source or failed independent check."""


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _q(name):
    return f"{{{NS}}}{name}"


def _worksheet(cells):
    root = ET.Element(_q("worksheet"))
    columns = ET.SubElement(root, _q("cols"))
    ET.SubElement(columns, _q("col"), min="1", max="1", width="48", customWidth="1")
    ET.SubElement(columns, _q("col"), min="2", max="12", width="18", customWidth="1")
    data, rows = ET.SubElement(root, _q("sheetData")), {}
    for address, value in sorted(cells.items(), key=lambda item: (int(re.search(r"\d+", item[0])[0]), item[0])):
        row_number = re.search(r"\d+", address)[0]
        if row_number not in rows:
            rows[row_number] = ET.SubElement(data, _q("row"), r=row_number)
        cell = ET.SubElement(rows[row_number], _q("c"), r=address)
        if isinstance(value, Decimal):
            ET.SubElement(cell, _q("v")).text = format(value, "f")
        elif value.startswith("="):
            ET.SubElement(cell, _q("f")).text = value[1:]
            ET.SubElement(cell, _q("v")).text = "0"
        else:
            cell.set("t", "inlineStr")
            ET.SubElement(ET.SubElement(cell, _q("is")), _q("t")).text = value
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _fixture_cells(request, catalog):
    prices = {record["resource"]: decimal(record["price"]) for record in catalog["records"]}
    return {
        "Prices": {"A1": "SYNTHETIC / DEMO_ONLY", "B1": "KRW, VAT excluded",
                   "A2": "DEMO plywood 12 mm / m2", "B2": prices["plywood_12mm"],
                   "A3": "DEMO carpenter / person_day", "B3": prices["carpenter"],
                   "A4": "DEMO helper / person_day", "B4": prices["helper"],
                   "A6": "Invented demo prices; not an official source."},
        "Estimate": {"A1": "SYNTHETIC / DEMO_ONLY", "D1": "Quantity m2", "E1": "Material / m2",
                     "F1": "Material amount", "G1": "Labor / m2", "H1": "Labor amount", "I1": "Expense / m2",
                     "J1": "Expense amount", "K1": "Unit total", "L1": "Direct amount",
                     "A2": "DEMO plywood installation 12 mm", "D2": decimal(request["quantity"]),
                     "E2": "=TRUNC(Prices!B2*1.03,0)+TRUNC(TRUNC(Prices!B2*1.03,0)*0.03,0)",
                     "F2": "=TRUNC(D2*E2,0)", "G2": "=TRUNC((Prices!B3*2+Prices!B4)/40,0)",
                     "H2": "=TRUNC(D2*G2,0)", "I2": "=TRUNC(G2*0.02,0)", "J2": "=TRUNC(D2*I2,0)",
                     "K2": "=SUM(E2,G2,I2)", "L2": "=SUM(F2,H2,J2)",
                     "A4": "Synthetic coefficients only. VAT, overhead, profit and delivery excluded."},
        "Summary": {"A1": "SYNTHETIC / DEMO_ONLY", "A2": "Material amount", "B2": "=Estimate!F2",
                    "A3": "Labor amount", "B3": "=Estimate!H2", "A4": "Expense amount", "B4": "=Estimate!J2",
                    "A5": "Known direct total (complete fixture)", "B5": "=SUM(B2:B4)",
                    "A7": "Restricted OOXML checks only; native Excel not verified."},
    }


def create_fixture(path, request, catalog):
    """Generate an explicit synthetic fixture; never ingest a real workbook."""
    path = Path(path)
    if path.exists():
        raise WorkbookError("Refusing to overwrite an existing fixture")
    # Formula coefficients are fixed to this demo's documented contract.
    expected = {"material_loss_rate": "0.03", "fixings_rate": "0.03", "carpenter_person_days": "2",
                "helper_person_days": "1", "daily_productivity_m2": "40", "expense_rate_on_labor": "0.02"}
    if request.get("classification") != "SYNTHETIC" or request.get("purpose") != "DEMO_ONLY":
        raise WorkbookError("Synthetic demo fixture required")
    if any(decimal(request["assumptions"].get(key)) != Decimal(value) for key, value in expected.items()):
        raise WorkbookError("Workbook fixture supports its fixed synthetic coefficients only")
    cells = _fixture_cells(request, catalog)
    content_types = ('<?xml version="1.0" encoding="UTF-8"?>'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                     + ''.join(f'<Override PartName="/{part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                               for part in SHEETS.values()) + '</Types>')
    workbook = (f'<workbook xmlns="{NS}" xmlns:r="{REL}"><sheets>'
                + ''.join(f'<sheet name="{name}" sheetId="{index}" r:id="rId{index}"/>'
                          for index, name in enumerate(SHEETS, 1))
                + '</sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')
    relations = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 + ''.join(f'<Relationship Id="rId{index}" Type="{REL}/worksheet" Target="worksheets/sheet{index}.xml"/>'
                           for index in range(1, 4)) + '</Relationships>')
    parts = {"[Content_Types].xml": content_types.encode(), "xl/workbook.xml": workbook.encode(),
             "xl/_rels/workbook.xml.rels": relations.encode(),
             "_rels/.rels": (f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                             f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/>'
                             '</Relationships>').encode()}
    parts.update({SHEETS[name]: _worksheet(values) for name, values in cells.items()})
    book = _Book(parts)
    book.recalculate()
    path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(path, "x", compression=ZIP_DEFLATED) as archive:
        for name, payload in book.serialized().items():
            archive.writestr(name, payload)
    return sha256(path)


class _Book:
    def __init__(self, parts):
        self.parts = dict(parts)
        if set(parts) != {"[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "_rels/.rels", *SHEETS.values()}:
            raise WorkbookError("Only the generated synthetic fixture's seven ZIP parts are supported")
        self.trees = {name: ET.fromstring(parts[part]) for name, part in SHEETS.items()}
        self.cells = {name: {cell.get("r"): cell for cell in tree.iter(_q("c"))}
                      for name, tree in self.trees.items()}
        if any('SYNTHETIC / DEMO_ONLY' not in ET.tostring(tree, encoding="unicode") for tree in self.trees.values()):
            raise WorkbookError("Synthetic marker missing")
        self.values, self.visiting = {}, set()

    def read(self, sheet, address):
        key = sheet, address
        if key in self.values:
            return self.values[key]
        if key in self.visiting:
            raise WorkbookError("Cyclic dependency")
        try:
            cell = self.cells[sheet][address]
        except KeyError as error:
            raise WorkbookError(f"Unknown reference: {sheet}!{address}") from error
        if cell.get("t"):
            raise WorkbookError("A numeric dependency refers to text")
        self.visiting.add(key)
        formula = cell.find(_q("f"))
        if formula is None:
            value = decimal(cell.findtext(_q("v")))
        else:
            value = self.evaluate(formula.text, sheet)
        self.visiting.remove(key)
        self.values[key] = value
        return value

    def evaluate(self, formula, sheet):
        if not formula or len(formula) > 500:
            raise WorkbookError("Formula outside the supported demo scope")
        refs = []

        def replace(match):
            target_sheet = match.group("sheet") or sheet
            start, end = match.group("start"), match.group("end")
            refs.append((target_sheet, start, end))
            return f"_r{len(refs) - 1}"

        expression = re.sub(r"(?<![\w.])(?:(?P<sheet>[A-Za-z_][A-Za-z_0-9]*)!)?(?P<start>[A-Z]{1,3}[1-9][0-9]*)(?::(?P<end>[A-Z]{1,3}[1-9][0-9]*))?(?![\w])", replace, formula)
        try:
            tree = ast.parse(expression, mode="eval")
        except SyntaxError as error:
            raise WorkbookError("Unsupported formula syntax") from error

        def visit(node):
            if isinstance(node, ast.Constant) and type(node.value) in (int, float):
                return Decimal(ast.get_source_segment(expression, node))
            if isinstance(node, ast.Name) and re.fullmatch(r"_r\d+", node.id):
                target, first, last = refs[int(node.id[2:])]
                if last is None:
                    return self.read(target, first)
                first_match, last_match = re.fullmatch(r"([A-Z]+)(\d+)", first), re.fullmatch(r"([A-Z]+)(\d+)", last)
                if first_match[1] != last_match[1]:
                    raise WorkbookError("Only single-column ranges are supported")
                start, end = int(first_match[2]), int(last_match[2])
                if not 0 <= end - start <= 100:
                    raise WorkbookError("Range outside demo limits")
                return [self.read(target, f"{first_match[1]}{row}") for row in range(start, end + 1)]
            if isinstance(node, ast.BinOp) and type(node.op) in (ast.Add, ast.Sub, ast.Mult, ast.Div):
                left, right = visit(node.left), visit(node.right)
                if type(node.op) is ast.Add:
                    return left + right
                if type(node.op) is ast.Sub:
                    return left - right
                if type(node.op) is ast.Mult:
                    return left * right
                return left / right
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and not node.keywords:
                args = [visit(arg) for arg in node.args]
                if node.func.id == "TRUNC" and len(args) == 2 and args[1] == 0:
                    return trunc(args[0])
                if node.func.id == "SUM":
                    return sum((sum(value, Decimal("0")) if isinstance(value, list) else value for value in args), Decimal("0"))
            raise WorkbookError("Unsupported formula node; arbitrary code is never evaluated")

        with localcontext() as context:
            context.prec = 50
            return visit(tree.body)

    def recalculate(self):
        self.values.clear()
        for sheet, cells in self.cells.items():
            for address, cell in cells.items():
                if cell.find(_q("f")) is not None:
                    value = self.read(sheet, address)
                    cached = cell.find(_q("v"))
                    if cached is None:
                        cached = ET.SubElement(cell, _q("v"))
                    cached.text = format(value, "f")

    def serialized(self):
        result = dict(self.parts)
        result.update({SHEETS[name]: ET.tostring(tree, encoding="utf-8", xml_declaration=True)
                       for name, tree in self.trees.items()})
        return result


def edit_copy(source, destination, updates, expected_sha256, expected_result):
    """Apply allowlisted numeric inputs, recalculate and verify on a new copy."""
    source, destination = Path(source), Path(destination)
    if source.resolve() == destination.resolve() or destination.exists():
        raise WorkbookError("A new destination is required; originals and existing files are never overwritten")
    if sha256(source) != expected_sha256:
        raise WorkbookError("Source SHA256 mismatch")
    if expected_result.get("status") != "COMPLETE":
        raise WorkbookError("Incomplete prices cannot be written as a complete fixture")
    with ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    book = _Book(parts)
    formulas_before = {(sheet, address): cell.findtext(_q("f")) for sheet, cells in book.cells.items()
                       for address, cell in cells.items() if cell.find(_q("f")) is not None}
    for key, raw in updates.items():
        if key not in ALLOWED_INPUTS:
            raise WorkbookError("Only explicitly allowlisted numeric inputs may change")
        cell = book.cells[key[0]][key[1]]
        cell.find(_q("v")).text = format(decimal(raw), "f")
    book.recalculate()
    expected = {
        ("Estimate", "E2"): expected_result["unit_components"]["material"],
        ("Estimate", "G2"): expected_result["unit_components"]["labor"],
        ("Estimate", "I2"): expected_result["unit_components"]["expense"],
        ("Estimate", "K2"): expected_result["unit_total"],
        ("Estimate", "L2"): expected_result["total"],
        ("Summary", "B2"): expected_result["amount_components"]["material"],
        ("Summary", "B3"): expected_result["amount_components"]["labor"],
        ("Summary", "B4"): expected_result["amount_components"]["expense"],
        ("Summary", "B5"): expected_result["total"],
    }
    checks = []
    for (sheet, address), raw in expected.items():
        actual, wanted = book.read(sheet, address), decimal(raw)
        checks.append({"sheet": sheet, "cell": address, "actual": format(actual, "f"),
                       "expected": format(wanted, "f"), "passed": actual == wanted})
    if not all(check["passed"] for check in checks):
        raise WorkbookError("OOXML linked values disagree with the independent cost calculation")
    formulas_after = {(sheet, address): cell.findtext(_q("f")) for sheet, cells in book.cells.items()
                      for address, cell in cells.items() if cell.find(_q("f")) is not None}
    if formulas_before != formulas_after:
        raise WorkbookError("Original formula text changed unexpectedly")
    updated = book.serialized()
    changed = sorted(name for name in parts if parts[name] != updated[name])
    if any(name not in SHEETS.values() for name in changed):
        raise WorkbookError("An unapproved OOXML part changed")
    if sha256(source) != expected_sha256:
        raise WorkbookError("Source changed while the edit was being prepared")
    destination.parent.mkdir(parents=True, exist_ok=True)
    # Exclusive creation avoids silently replacing another revision.
    with tempfile.TemporaryFile() as temporary:
        with ZipFile(temporary, "w", compression=ZIP_DEFLATED) as archive:
            for name, payload in updated.items():
                archive.writestr(name, payload)
        temporary.seek(0)
        with destination.open("xb") as output:
            output.write(temporary.read())
    if sha256(source) != expected_sha256:
        destination.unlink()
        raise WorkbookError("Original preservation check failed")
    return {
        "status": "PASS_SELECTED_SYNTHETIC_SCOPE", "classification": "SYNTHETIC", "purpose": "DEMO_ONLY",
        "source_sha256_before": expected_sha256, "source_sha256_after": sha256(source),
        "source_unchanged": True, "output_sha256": sha256(destination),
        "formulas_preserved": True, "changed_ooxml_parts": changed,
        "untouched_parts_preserved_byte_for_byte": True, "check_count": len(checks), "checks": checks,
        "validation_method": "Restricted OOXML evaluator vs independent Decimal costing",
        "native_excel_verified": False, "full_contract_amount_verified": False,
    }
