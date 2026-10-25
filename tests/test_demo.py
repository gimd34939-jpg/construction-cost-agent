"""Selected synthetic cases; these tests do not validate the private engine."""

from copy import deepcopy
from decimal import Decimal
import json
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

from demo.__main__ import run
from demo.costing import InputError, calculate, decimal, load_fixtures
from demo.workbook import NS, SHEETS, WorkbookError, create_fixture, edit_copy, sha256


class CostingTests(unittest.TestCase):
    def setUp(self):
        self.request, self.catalog = load_fixtures()

    def result(self):
        return calculate(self.request, self.catalog)

    def test_complete_synthetic_case(self):
        result = self.result()
        self.assertEqual(result["status"], "COMPLETE")
        self.assertEqual(result["unit_components"], {"material": "12730", "labor": "18250", "expense": "365"})
        self.assertEqual(result["unit_total"], "31345")
        self.assertEqual(result["intermediates"], {"material_with_loss": "12360", "fixings": "370"})
        self.assertEqual(result["classification"], "SYNTHETIC")
        self.assertTrue(all(source["purpose"] == "DEMO_ONLY" for source in result["sources"]))

    def test_quantity_truncates_each_component_before_sum(self):
        self.request["quantity"] = "2.01"
        result = self.result()
        self.assertEqual(result["amount_components"], {"material": "25587", "labor": "36682", "expense": "733"})
        self.assertEqual(result["total"], "63002")
        self.assertNotEqual(result["total"], "63003")  # A single final truncation would differ.

    def test_missing_material_is_null_with_known_subtotal(self):
        self.catalog["records"].pop(0)
        result = self.result()
        self.assertEqual(result["status"], "INCOMPLETE")
        self.assertIsNone(result["unit_components"]["material"])
        self.assertIsNone(result["total"])
        self.assertEqual(result["known_unit_subtotal"], "18615")

    def test_missing_labor_also_blocks_dependent_expense(self):
        self.catalog["records"].pop(1)
        result = self.result()
        self.assertIsNone(result["unit_components"]["labor"])
        self.assertIsNone(result["unit_components"]["expense"])
        self.assertEqual(result["known_unit_subtotal"], "12730")

    def test_unit_conflict_is_not_implicitly_converted(self):
        self.catalog["records"][0]["unit"] = "sheet"
        result = self.result()
        self.assertEqual(result["issues"][0]["code"], "UNIT_CONFLICT")
        self.assertIsNone(result["total"])

    def test_out_of_period_cannot_become_current_price(self):
        self.request["pricing_date"] = "2027-01-01"
        result = self.result()
        self.assertEqual(result["status"], "INCOMPLETE")
        self.assertTrue(all(issue["code"] == "DATE_OUT_OF_SCOPE" for issue in result["issues"]))
        self.assertIsNone(result["total"])

    def test_missing_validity_does_not_invent_an_effective_period(self):
        del self.catalog["records"][0]["effective_end"]
        self.assertEqual(self.result()["issues"][0]["code"], "DATE_OUT_OF_SCOPE")

    def test_missing_source_is_blocked(self):
        del self.catalog["records"][0]["source_id"]
        self.assertEqual(self.result()["issues"][0]["code"], "MISSING_SOURCE")

    def test_tax_scope_and_specification_gates(self):
        for key, value, code in (("tax_basis", "UNKNOWN", "TAX_BASIS_UNCONFIRMED"),
                                 ("scope", "SUPPLY_AND_INSTALL", "SCOPE_CONFLICT"),
                                 ("specification", "DEMO plywood 9 mm", "SPECIFICATION_CONFLICT"),
                                 ("delivery", "INCLUDED", "DELIVERY_CONFLICT")):
            with self.subTest(key=key):
                catalog = deepcopy(self.catalog)
                catalog["records"][0][key] = value
                self.assertEqual(calculate(self.request, catalog)["issues"][0]["code"], code)

    def test_ambiguous_price_is_not_arbitrarily_selected(self):
        self.catalog["records"].append(deepcopy(self.catalog["records"][0]))
        self.assertEqual(self.result()["issues"][0]["code"], "MISSING_OR_AMBIGUOUS_PRICE")

    def test_non_demo_source_cannot_be_presented_as_verified(self):
        self.catalog["records"][0]["classification"] = "OFFICIAL"
        self.assertEqual(self.result()["issues"][0]["code"], "NOT_DEMO_ONLY")

    def test_float_nonfinite_bool_and_negative_inputs_rejected(self):
        for value in (1.01, True, "NaN", "Infinity", "-1", "1e5"):
            with self.subTest(value=value), self.assertRaises(InputError):
                decimal(value)

    def test_zero_is_an_explicit_value_not_a_missing_price(self):
        self.catalog["records"][0]["price"] = "0"
        result = self.result()
        self.assertEqual(result["status"], "COMPLETE")
        self.assertEqual(result["unit_components"]["material"], "0")

    def test_new_material_price_changes_total_without_contract_fallback(self):
        before = self.result()
        self.catalog["records"][0]["price"] = "13000"
        after = self.result()
        self.assertEqual(after["unit_total"], "32406")
        self.assertEqual(Decimal(after["unit_total"]) - Decimal(before["unit_total"]), Decimal("1061"))

    def test_zero_productivity_blocks_request(self):
        self.request["assumptions"]["daily_productivity_m2"] = "0"
        with self.assertRaises(InputError):
            self.result()


class WorkbookTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.folder = Path(self.temporary.name)
        self.request, self.catalog = load_fixtures()
        self.source = self.folder / "original.xlsx"
        self.output = self.folder / "revision.xlsx"
        self.original_hash = create_fixture(self.source, self.request, self.catalog)

    def test_price_change_reaches_estimate_and_summary_and_preserves_original(self):
        catalog = deepcopy(self.catalog)
        catalog["records"][0]["price"] = "13000"
        expected = calculate(self.request, catalog)
        before_bytes = self.source.read_bytes()
        report = edit_copy(self.source, self.output, {("Prices", "B2"): "13000"}, self.original_hash, expected)
        self.assertEqual(self.source.read_bytes(), before_bytes)
        self.assertEqual(sha256(self.source), self.original_hash)
        self.assertEqual(report["check_count"], 9)
        self.assertTrue(all(check["passed"] for check in report["checks"]))
        self.assertFalse(report["native_excel_verified"])
        self.assertFalse(report["full_contract_amount_verified"])
        with ZipFile(self.source) as source, ZipFile(self.output) as output:
            for name in source.namelist():
                if name not in SHEETS.values():
                    self.assertEqual(source.read(name), output.read(name))
                else:
                    original = ET.fromstring(source.read(name))
                    revision = ET.fromstring(output.read(name))
                    self.assertEqual([node.text for node in original.iter(f"{{{NS}}}f")],
                                     [node.text for node in revision.iter(f"{{{NS}}}f")])

    def test_quantity_change_reaches_all_component_amounts(self):
        request = deepcopy(self.request)
        request["quantity"] = "2.01"
        report = edit_copy(self.source, self.output, {("Estimate", "D2"): "2.01"}, self.original_hash,
                           calculate(request, self.catalog))
        total = next(check for check in report["checks"] if check["sheet"] == "Summary" and check["cell"] == "B5")
        self.assertEqual(total["actual"], "63002")

    def test_wrong_hash_blocks_before_writing(self):
        with self.assertRaisesRegex(WorkbookError, "SHA256"):
            edit_copy(self.source, self.output, {}, "0" * 64, calculate(self.request, self.catalog))
        self.assertFalse(self.output.exists())

    def test_existing_destination_and_original_overwrite_blocked(self):
        for destination in (self.source,):
            with self.assertRaisesRegex(WorkbookError, "new destination"):
                edit_copy(self.source, destination, {}, self.original_hash, calculate(self.request, self.catalog))
        self.output.write_bytes(b"previous revision")
        with self.assertRaises(WorkbookError):
            edit_copy(self.source, self.output, {}, self.original_hash, calculate(self.request, self.catalog))
        self.assertEqual(self.output.read_bytes(), b"previous revision")

    def test_formula_cell_is_not_an_allowed_input(self):
        with self.assertRaisesRegex(WorkbookError, "allowlisted"):
            edit_copy(self.source, self.output, {("Estimate", "L2"): "0"}, self.original_hash,
                      calculate(self.request, self.catalog))
        self.assertFalse(self.output.exists())

    def test_incomplete_result_cannot_be_written_as_complete(self):
        self.catalog["records"].pop(0)
        with self.assertRaisesRegex(WorkbookError, "Incomplete"):
            edit_copy(self.source, self.output, {}, self.original_hash, calculate(self.request, self.catalog))

    def test_independent_mismatch_blocks_output(self):
        expected = calculate(self.request, self.catalog)
        expected["total"] = "999999"
        with self.assertRaisesRegex(WorkbookError, "independent"):
            edit_copy(self.source, self.output, {}, self.original_hash, expected)
        self.assertFalse(self.output.exists())

    def test_unsupported_formula_is_not_evaluated(self):
        with ZipFile(self.source) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        parts[SHEETS["Estimate"]] = parts[SHEETS["Estimate"]].replace(b"SUM(F2,H2,J2)", b"UNSUPPORTED(F2)")
        altered = self.folder / "unsupported.xlsx"
        with ZipFile(altered, "x", compression=ZIP_DEFLATED) as archive:
            for name, payload in parts.items():
                archive.writestr(name, payload)
        with self.assertRaisesRegex(WorkbookError, "Unsupported"):
            edit_copy(altered, self.output, {}, sha256(altered), calculate(self.request, self.catalog))

    def test_external_link_parts_are_outside_demo_scope(self):
        altered = self.folder / "external.xlsx"
        with ZipFile(self.source) as source, ZipFile(altered, "x") as output:
            for name in source.namelist():
                output.writestr(name, source.read(name))
            output.writestr("xl/externalLinks/externalLink1.xml", b"<externalLink/>")
        with self.assertRaisesRegex(WorkbookError, "seven ZIP parts"):
            edit_copy(altered, self.output, {}, sha256(altered), calculate(self.request, self.catalog))

    def test_cli_report_and_optional_output_folder(self):
        output = self.folder / "run"
        result = run(output)
        self.assertEqual(result["complete_case"]["unit_total"], "31345")
        self.assertIsNone(result["missing_price_case"]["total"])
        self.assertEqual(result["price_change_case"]["unit_total"], "32406")
        self.assertEqual(set(path.name for path in output.iterdir()),
                         {"synthetic_original.xlsx", "synthetic_repriced.xlsx", "result.json"})
        stored = json.loads((output / "result.json").read_text(encoding="utf-8"))
        self.assertTrue(stored["workbook_verification"]["source_unchanged"])
        with self.assertRaises(FileExistsError):
            run(output)


if __name__ == "__main__":
    unittest.main()
