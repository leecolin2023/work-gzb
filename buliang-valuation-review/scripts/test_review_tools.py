"""Regression tests for arithmetic error localization and evidence navigation.

Run: python3 -m unittest discover -s scripts -p 'test_*.py' -v
"""
import copy
import json
import unittest
from decimal import Decimal
from pathlib import Path

from anchors import CHECKS, EXTENDED_CHECKS, recall
from verify_formulas import verify


def input_value(value, unit="元"):
    return {"value": value, "unit": unit, "evidence": ["fixture:source"]}


def node(name, formula, reported, unit="元", tolerance="0.01", **extra):
    return {"id": name, "formula": formula, "reported": reported,
            "unit": unit, "tolerance": tolerance, "evidence": ["fixture:report"], **extra}


def ledger(nodes):
    return {"model": "formula_graph", "inputs": {"a": input_value("100")},
            "nodes": nodes, "result": nodes[-1]["id"]}


class FormulaTests(unittest.TestCase):
    def test_upstream_error_is_not_misclassified_as_local_error(self):
        p = ledger([node("subtotal", "a + a", "199"), node("total", "subtotal * 2", "398")])
        rows = verify(p)["checks"]
        self.assertEqual(rows[0]["difference"], Decimal(1))
        self.assertEqual(rows[1]["difference"], Decimal(2))
        self.assertEqual(rows[1]["local_difference"], Decimal(0))

    def test_bad_reported_divisor_does_not_discard_checks(self):
        p = ledger([node("subtotal", "a + a", "0"), node("ratio", "a / subtotal", "0.5", unit="比例")])
        result = verify(p)
        self.assertEqual(result["result"], Decimal("0.5"))
        self.assertIn("DivisionByZero", result["checks"][1]["local_error"])
        self.assertEqual(result["checks"][0]["difference"], Decimal(200))

    def test_missing_unknown_and_wrong_model_are_rejected(self):
        original = ledger([node("total", "a * 2", "200")])
        for value in (None, "NaN", "Infinity", True):
            p = copy.deepcopy(original)
            p["inputs"]["a"]["value"] = value
            with self.assertRaises(ValueError):
                verify(p)
        for formula in ("unknown + 1", "total + a", "__import__('os').getcwd()"):
            p = copy.deepcopy(original)
            p["nodes"][0]["formula"] = formula
            with self.assertRaises(ValueError):
                verify(p)
        original["model"] = "unrecognized"
        with self.assertRaises(ValueError):
            verify(original)

    def test_decimal_rounding_and_explicit_tolerance(self):
        p = ledger([node("total", "1.005", "1.01", round_to="0.01")])
        self.assertEqual(verify(p)["result"], Decimal("1.01"))
        del p["nodes"][0]["tolerance"]
        with self.assertRaises(ValueError):
            verify(p)

    def test_formula_and_wrong_substitution_are_separate(self):
        p = ledger([node("expected", "a * 16", "1800",
                         reported_formula_text="面积×租金，租金取16",
                         reported_substitution_text="100×18=1800"),
                    node("as_substituted", "a * 18", "1800")])
        rows = verify(p)["checks"]
        self.assertTrue(rows[0]["exceeds_tolerance"])
        self.assertFalse(rows[1]["exceeds_tolerance"])
        self.assertEqual(rows[0]["reported_substitution_text"], "100×18=1800")

    def test_supported_deduction_changes_asset_then_remaining_claim(self):
        p = {"model": "formula_graph", "inputs": {
            "claim": input_value("1000"), "asset": input_value("500"),
            "deduction": input_value("0"), "rate": input_value("0.1", "比例")},
            "nodes": [node("priority", "max(0, asset - deduction)", "500"),
                      node("general", "(claim - priority) * rate", "50"),
                      node("total", "min(claim, priority + general)", "550")],
            "result": "total", "scenarios": [{"name": "confirmed burden", "reason": "fixture confirmed asset burden",
                "overrides": {"deduction": {"value": "100", "basis": "evidence", "evidence": ["fixture:settlement"]}}}]}
        result = verify(p, True)
        self.assertEqual(result["result"], Decimal(550))
        self.assertEqual(result["sensitivity"][0]["result"], Decimal(460))
        self.assertEqual(result["sensitivity"][0]["delta"], Decimal(-90))

    def test_lingnan_model_matches_independent_known_amounts(self):
        p = json.loads((Path(__file__).resolve().parents[1] / "assets/params_lingnan_example.json").read_text())
        result = verify(p, True)
        self.assertEqual(result["values"]["debtor"], Decimal("2162.0498015196"))
        self.assertEqual(result["values"]["legal_total"], Decimal("3230.7157509317064"))
        self.assertEqual(result["values"]["natural_total"], Decimal("51.112288452"))
        self.assertEqual(result["result"], Decimal("5972.817218676207208"))
        income = result["values"]["income_example"]
        self.assertLess(abs(income - Decimal("352922.7415")), Decimal("0.0001"))
        self.assertEqual(len(result["sensitivity"]), 1)


class EvidenceTests(unittest.TestCase):
    def test_remote_numeric_case_retains_row_and_table_references(self):
        rows = [("可比案例", "距离"), ("案例甲", "5公里"), ("案例乙", "8公里"), ("某异地商业铺位", "80公里")]
        items = [{"id": "T01-R%02d-C%02d" % (r, c), "type": "table_cell",
                  "table": 1, "row": r, "col": c, "text": text}
                 for r, row in enumerate(rows, 1) for c, text in enumerate(row, 1)]
        check = next(c for c in EXTENDED_CHECKS if c["id"] == "C02")
        candidates = recall(items, check)
        self.assertTrue(any("80公里" in c["text"] for c in candidates))
        self.assertTrue(all("T01-R04-C02" in c["table_evidence_ids"] for c in candidates))
        self.assertTrue(any("T01-R04-C01" in c["evidence_ids"] for c in candidates))

    def test_zero_recall_remains_only_navigation(self):
        self.assertEqual(recall([], CHECKS[0]), [])
        self.assertEqual(len(CHECKS), 40)
        self.assertEqual(len(EXTENDED_CHECKS), 11)
        self.assertEqual(len({c["id"] for c in CHECKS + EXTENDED_CHECKS}), 51)


if __name__ == "__main__":
    unittest.main()
