"""Offline checks of selected public engine modules using synthetic inputs.

These tests exercise the published source, not the private databases, a live
API, workbook recalculation, or the applicability of any field estimate.
"""

import json
import unittest

from engine import api_clients, pricing_policy, quantity_engine


class PublicQuantityEngineTests(unittest.TestCase):
    def setUp(self):
        self.rule = {
            "expression": "length * width * thickness",
            "output_unit": "m3",
            "variables": {
                "length": {"unit": "m"},
                "width": {"unit": "m"},
                "thickness": {"unit": "m"},
            },
        }
        self.inputs = {
            name: {
                "value": value,
                "unit": unit,
                "evidence": {"type": "TEST_ONLY", "note": "Synthetic dimensions"},
            }
            for name, value, unit in (
                ("length", "2500", "mm"),
                ("width", "2", "m"),
                ("thickness", "10", "cm"),
            )
        }

    def test_mixed_length_units_produce_cubic_metres_and_retain_test_scope(self):
        result = quantity_engine.evaluate_quantity(self.rule, self.inputs)
        # 2.5 m * 2 m * 0.1 m = 0.5 m3, independently of input units.
        self.assertEqual(result["status"], "COMPLETE")
        self.assertEqual(result["value"], "0.5")
        self.assertEqual(result["unit"], "m3")
        self.assertTrue(result["test_only"])

    def test_missing_dimension_stays_incomplete_without_a_numeric_result(self):
        self.inputs.pop("thickness")
        result = quantity_engine.evaluate_quantity(self.rule, self.inputs)
        self.assertEqual(result["status"], "INCOMPLETE")
        self.assertIsNone(result["value"])
        self.assertEqual(result["missing"][0]["variable"], "thickness")

    def test_function_call_is_rejected_instead_of_executed(self):
        self.rule["expression"] = '__import__("os")'
        result = quantity_engine.evaluate_quantity(self.rule, self.inputs)
        self.assertEqual(result["status"], "BLOCKED")
        self.assertIsNone(result["value"])
        self.assertIn("UNSAFE_EXPRESSION", [error["code"] for error in result["errors"]])


class PublicApiParserTests(unittest.TestCase):
    def test_pps_success_envelope_preserves_row_and_pagination(self):
        response = {
            "response": {
                "header": {"resultCode": "00"},
                "body": {
                    "items": {"item": [{"itemName": "SYNTHETIC", "price": 1234.56}]},
                    "totalCount": 1,
                    "pageNo": 1,
                    "numOfRows": 10,
                },
            }
        }
        result = api_clients.parse_response(json.dumps(response).encode(), "PPS_PRICE")
        self.assertEqual(result["rows"], [{"itemName": "SYNTHETIC", "price": "1234.56"}])
        self.assertEqual(result["total_count"], 1)
        self.assertEqual(result["page_no"], 1)
        self.assertEqual(result["num_of_rows"], 10)

    def test_xml_doctype_is_rejected(self):
        response = b'<!DOCTYPE response [<!ENTITY example "synthetic">]><response/>'
        with self.assertRaises(api_clients.ApiError) as caught:
            api_clients.parse_response(response, "PPS_PRICE")
        self.assertEqual(caught.exception.code, "UNSAFE_XML")


class PublicPricingPolicyTests(unittest.TestCase):
    def test_default_policy_requires_independent_external_prices(self):
        self.assertEqual(pricing_policy.validate_policy({}), "EXTERNAL_PRIMARY")

    def test_existing_workbook_prices_are_not_authorized_by_default(self):
        self.assertFalse(pricing_policy.reference_authorized({}))
        with self.assertRaises(pricing_policy.PricingPolicyError):
            pricing_policy.require_reference({})


if __name__ == "__main__":
    unittest.main()
