import importlib.util
import unittest
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("app_tsc_matrix_under_test", ROOT / "app.py")
APP = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(APP)


class TscFreshFrozenMatrixTests(unittest.TestCase):
    def test_fixed_layouts_and_special_uses_breast_specs(self):
        leg = APP._tsc_matrix_columns("腿肉")
        breast = APP._tsc_matrix_columns("胸肉")
        special = APP._tsc_matrix_columns("特殊")

        self.assertEqual(leg[0], "鲜品|无规格")
        self.assertIn("鲜品|80g", leg)
        self.assertIn("冻品|120g以上", leg)
        self.assertIn("鲜品|其他规格", leg)
        self.assertEqual(leg[-2:], ["冻品合计", "累计"])
        self.assertEqual(special, breast)
        self.assertEqual(APP._tsc_matrix_columns("其他"), ["鲜品", "冻品", "鲜冻品合计"])

    def test_39000308_and_39000172_bucket_values_reconcile(self):
        metadata = {
            "31002330": {"spec": "80g", "description": "鸡去皮腿肉/80g/鲜品", "fresh_frozen": "鲜品"},
            "31003723": {"spec": "120g以上", "description": "鸡去皮腿肉/120g以上/鲜品", "fresh_frozen": "鲜品"},
            "31001666": {"spec": "80g", "description": "鸡去皮腿肉/80g/冻品", "fresh_frozen": "冻品"},
        }
        rows_308 = pd.DataFrame([
            {"原料号_raw": "31002330", "正向数量": 129984.7, "正向金额": 1267338.04},
            {"原料号_raw": "31003723", "正向数量": 1720.0, "正向金额": 16855.06},
        ])
        qty_308, amt_308 = APP._aggregate_tsc_usage_rows(rows_308, "腿肉", metadata)
        self.assertAlmostEqual(qty_308["鲜品|80g"] / 1000, 129.9847)
        self.assertAlmostEqual(qty_308["鲜品|120g以上"] / 1000, 1.72)
        self.assertAlmostEqual(amt_308["鲜品|80g"] / 1000, 1267.33804)
        self.assertAlmostEqual(amt_308["鲜品|120g以上"] / 1000, 16.85506)
        self.assertAlmostEqual(qty_308["鲜品|80g"] / sum(qty_308.values()), 0.986940481243266)
        self.assertAlmostEqual(amt_308["鲜品|80g"] / qty_308["鲜品|80g"], 9.749901642270206)

        rows_172 = pd.DataFrame([
            {"原料号_raw": "31002330", "正向数量": 7210.0, "正向金额": 73644.54},
            {"原料号_raw": "31001666", "正向数量": 840.0, "正向金额": 8475.6},
        ])
        qty_172, amt_172 = APP._aggregate_tsc_usage_rows(rows_172, "腿肉", metadata)
        self.assertAlmostEqual(qty_172["鲜品|80g"] / 1000, 7.21)
        self.assertAlmostEqual(qty_172["冻品|80g"] / 1000, 0.84)
        self.assertAlmostEqual(qty_172["鲜品|80g"] / sum(qty_172.values()), 0.8956521739130435)
        self.assertAlmostEqual(sum(amt_172.values()) / sum(qty_172.values()), 10.20125962732919)

    def test_unknown_spec_goes_to_other_but_unknown_fresh_frozen_blocks(self):
        rows = pd.DataFrame([
            {"原料号_raw": "31009998", "正向数量": 1000.0, "正向金额": 9000.0},
        ])
        metadata = {
            "31009998": {"spec": "999g", "description": "鸡腿肉/999g/鲜品", "fresh_frozen": "鲜品"},
        }
        qty, _ = APP._aggregate_tsc_usage_rows(rows, "腿肉", metadata)
        self.assertEqual(qty["鲜品|其他规格"], 1000.0)

        with self.assertRaisesRegex(ValueError, "31009998"):
            APP._aggregate_tsc_usage_rows(
                rows,
                "腿肉",
                {"31009998": {"spec": "80g", "description": "鸡腿肉/80g", "fresh_frozen": None}},
            )

    def test_new_reference_matrix_ratio_is_read_by_stable_key(self):
        data = [
            ["产品族", "修行后原料", "使用半成品规格", "行类型", "影响口径", "鲜品", None, None, "冻品", None, "累计", "综合单价"],
            [None, None, None, None, None, "80g", "120g以上", "鲜品合计", "80g", "冻品合计", "累计", "修形前原料综合耗用单价"],
            ["香酥炸鸡", "39000172", "规格", "Q3规格占比", None, 0.90, 0.0, 0.90, 0.10, 0.10, 1.0, None],
        ]
        values = APP._find_tsc_matrix_reference_values(
            pd.DataFrame(data),
            "39000172",
            "规格占比",
            "Q3",
            "腿肉",
            {},
            label_kind="ratio",
        )
        self.assertEqual(values["鲜品|80g"], 0.90)
        self.assertEqual(values["冻品|80g"], 0.10)

    def test_legacy_raw_code_reference_is_converted_to_matrix_keys(self):
        data = [
            ["产品族", "修行后原料", "使用半成品规格", "行类型", "影响口径", "31002330", "31001666", "综合单价", "修形前原料综合耗用单价"],
            [None, None, None, None, None, "80g", "80g", None, None],
            ["香酥炸鸡", "39000172", "规格", "Q3规格占比", None, 0.90, 0.10, None, None],
        ]
        metadata = {
            "31002330": {"spec": "80g", "description": "鸡腿肉/80g/鲜品", "fresh_frozen": "鲜品"},
            "31001666": {"spec": "80g", "description": "鸡腿肉/80g/冻品", "fresh_frozen": "冻品"},
        }
        values = APP._find_tsc_matrix_reference_values(
            pd.DataFrame(data),
            "39000172",
            "规格占比",
            "Q3",
            "腿肉",
            metadata,
            [(5, "31002330"), (6, "31001666")],
            {"31002330": "80g", "31001666": "80g"},
            label_kind="ratio",
        )
        self.assertEqual(values, {"鲜品|80g": 0.90, "冻品|80g": 0.10})


if __name__ == "__main__":
    unittest.main()
