"""Authority tests proving Run Audit calculates from the standard catalog."""

import unittest

import pandas as pd

from excel_engine.product_formatter import _product_type
from excel_engine.workbook_report import calculate_raw_sheet


def row_for(description, *, W=360, H=360, R=180, K=2, outer="", inner="", plan="", up_down="", vertical=""):
    """Build one minimal raw worksheet row for authority tests."""
    values = [None] * 25
    values[2] = description
    values[5] = W
    values[6] = H
    values[12] = R
    values[16] = K
    values[17] = plan
    values[18] = up_down
    values[19] = vertical
    values[20] = 0.0
    values[21] = 0.0
    values[22] = outer
    values[23] = inner
    return values


def catalog(*entries):
    """Build the smallest catalog payload accepted by the report workflow."""
    return {"entries": list(entries), "count": len(entries), "sheet": "COMBINE"}


def entry(product_type, target, formula, category, family):
    """Build one standard catalog entry with the requested lookup keys."""
    return {
        "name": "Co Vuông",
        "product_type": product_type,
        "target": target,
        "formula": formula,
        "formula_block": f"{target}={formula}",
        "category": category,
        "family": family,
        "signature": None,
    }


class StandardAuthorityTest(unittest.TestCase):
    """Verify contractor text classifies rows but never becomes authority."""
    # Confirm contractor text is not the numeric calculation authority.
    def test_contractor_formula_is_not_used_for_calculation(self):
        data = pd.DataFrame([
            row_for("Square elbow 360x360", outer="S=(W+H)/1000000")
        ])
        reference = catalog(
            entry("square_elbow", "S", "K*(W+H)/1000000", "Cách nhiệt", "standard")
        )

        calculated, output, errors = calculate_raw_sheet(data, reference_catalog=reference)

        self.assertEqual([], errors)
        self.assertEqual("square_elbow", calculated[0]["product_type"])
        self.assertAlmostEqual(0.00144, output[0]["S"])
        self.assertNotAlmostEqual(output[0]["S"], 0.00072)
        self.assertEqual("different", output[0]["formula_audits"]["S"]["reference_status"])

    # Confirm shared SCC/TA is evaluated once as one inside quantity.
    def test_scc_ta_is_one_inside_quantity(self):
        data = pd.DataFrame([
            row_for("Square elbow 360x360", outer="S=(W+H)*K/1000000", inner="SCC/TA=(W+H)*K/1000000")
        ])
        reference = catalog(
            entry("square_elbow", "S", "(W+H)*K/1000000", "Cách nhiệt", "standard"),
            entry("square_elbow", "SCN", "(W+H)*K/1000000", "Cách nhiệt", "acoustic"),
            entry("square_elbow", "SCC/TA", "(W+H)*K/1000000", "Tiêu âm", "acoustic"),
        )

        _calculated, output, errors = calculate_raw_sheet(data, reference_catalog=reference)

        self.assertEqual([], errors)
        self.assertIn("S_inner", output[0]["formula_audits"])
        self.assertNotIn("S_fireproof", output[0]["formula_audits"])
        self.assertAlmostEqual(output[0]["S_inner"], output[0]["S"])

    # Confirm SCC falls back to the shared inside column in a two-column sheet.
    def test_scc_uses_shared_inner_target_in_two_column_layout(self):
        data = pd.DataFrame([
            row_for(
                "Square elbow 360x360",
                outer="S=(W+H)*K/1000000",
                inner="SCC=(W+H)*K/1000000",
            )
        ])
        reference = catalog(
            entry("square_elbow", "SCN", "(W+H)*K/1000000", "Cách nhiệt", "fireproof"),
            entry("square_elbow", "SCC", "(W+H)*K/1000000", "Chống cháy", "fireproof"),
        )

        _calculated, output, errors = calculate_raw_sheet(data, reference_catalog=reference)

        self.assertEqual([], errors)
        self.assertIn("S_inner", output[0]["formula_audits"])
        self.assertNotIn("S_fireproof", output[0]["formula_audits"])
        self.assertAlmostEqual(output[0]["S_inner"], output[0]["S"])

    # Confirm cosmetic parentheses do not create a false numeric mismatch.
    def test_parentheses_only_difference_is_ignored_when_value_matches(self):
        data = pd.DataFrame([
            row_for("Square elbow 360x360", outer="S=((K*(W+H)))/1000000")
        ])
        data.iat[0, 20] = 0.00144
        reference = catalog(
            entry("square_elbow", "S", "(K*(W+H))/1000000", "Cách nhiệt", "standard")
        )

        _calculated, output, errors = calculate_raw_sheet(data, reference_catalog=reference)

        self.assertEqual([], errors)
        self.assertEqual("", output[0]["formula_audits"]["S"]["reference_status"])
        self.assertIsNone(output[0]["formula_audits"]["S"]["formula_comparison"])
        self.assertTrue(output[0]["matches"]["S"])

    # Confirm a matching saved value stays valid when detail parsing fails.
    def test_matching_value_stays_green_when_detail_text_is_unverifiable(self):
        data = pd.DataFrame([
            row_for("Square elbow 360x360", outer="S=((W+H)*K/1000000")
        ])
        data.iat[0, 20] = 0.00144
        reference = catalog(
            entry("square_elbow", "S", "(W+H)*K/1000000", "Cách nhiệt", "standard")
        )

        _calculated, output, errors = calculate_raw_sheet(data, reference_catalog=reference)

        self.assertEqual([], errors)
        self.assertTrue(output[0]["matches"]["S"])
        self.assertTrue(output[0]["formula_audits"]["S"]["matches"])
        self.assertFalse(output[0]["formula_audits"]["S"]["error"])
        self.assertIsNone(output[0]["formula_audits"]["S"]["formula_comparison"])

    # Confirm the MFD included-in-valve exception chooses one outer quantity.
    def test_mfd_scn_with_inside_quantity_included_uses_one_outer_quantity(self):
        data = pd.DataFrame([
            row_for(
                "Fire damper (FD) 300x300",
                W=300,
                H=300,
                K=2,
                outer=(
                    "SCN=((W+H+50*4)*2*L*K+"
                    "(H+W+50*4+4*25)*2*K*Q*100)/1000000"
                ),
                inner="Bao gồm trong van",
            )
        ])
        data.iat[0, 11] = 210
        data.iat[0, 13] = 25
        data.iat[0, 15] = 2
        data.iat[0, 20] = 1.392
        reference = catalog(
            entry(
                "mfd_valve",
                "SCN",
                "((W+H+50*4)*2*L*K+(H+W+50*4+4*25)*2*K*Q*100)/1000000",
                "Cách nhiệt",
                "acoustic",
            )
        )

        _calculated, output, errors = calculate_raw_sheet(data, reference_catalog=reference)

        self.assertEqual([], errors)
        self.assertAlmostEqual(1.392, output[0]["S"])
        self.assertEqual(0.0, output[0]["S_inner"])
        self.assertTrue(output[0]["matches"]["S"])
        self.assertNotIn("S_inner", output[0]["formula_audits"])

    # Confirm non-MFD family mapping follows labels, not formula appearance.
    def test_non_mfd_family_comes_from_labels_not_formula_shape(self):
        data = pd.DataFrame([
            row_for(
                "Square elbow 360x360",
                outer="SCN=(K*a/25)+999",
                inner="STA=K*2",
            )
        ])
        reference = catalog(
            entry("square_elbow", "SCN", "K*100", "Cách nhiệt", "fireproof"),
            entry("square_elbow", "SCN", "K*10", "Cách nhiệt", "acoustic"),
            entry("square_elbow", "STA", "K*2", "Tiêu âm", "acoustic"),
        )

        _calculated, output, errors = calculate_raw_sheet(data, reference_catalog=reference)

        self.assertEqual([], errors)
        self.assertEqual(20.0, output[0]["S"])
        self.assertEqual("STA", output[0]["formula_audits"]["S_inner"]["label"])

    # Confirm K is recomputed only when a K component is populated.
    def test_k_uses_components_only_when_any_component_is_present(self):
        data = pd.DataFrame([
            row_for("Square elbow 360x360", K=4, outer="S=(W+H)*K/1000000", plan="", up_down="", vertical=""),
            row_for("Square elbow 360x360", K=4, outer="S=(W+H)*K/1000000", plan="1", up_down="2", vertical="3"),
        ])
        reference = catalog(
            entry("square_elbow", "S", "(W+H)*K/1000000", "Cách nhiệt", "standard")
        )

        _calculated, output, errors = calculate_raw_sheet(data, reference_catalog=reference)

        self.assertEqual([], errors)
        self.assertEqual(4, output[0]["K"])
        self.assertEqual(6, output[1]["K"])

    # Confirm a volume-control damper maps to its own product type.
    def test_volume_control_damper_is_not_an_elbow(self):
        self.assertEqual("nrd_valve", _product_type("Volume control damper (VCD) 400x200"))


if __name__ == "__main__":
    unittest.main()
