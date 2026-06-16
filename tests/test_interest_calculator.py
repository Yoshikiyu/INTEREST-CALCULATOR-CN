import unittest
from decimal import Decimal

from interest_calculator import (
    calculate_user_annual_rate,
    extract_current_interest_calculation,
    normalize_year_days,
)


class RateSettingsTests(unittest.TestCase):
    def test_daily_rate_uses_selected_year_days_for_annualized_rate(self):
        self.assertEqual(
            calculate_user_annual_rate("0", "0", "0.05", Decimal("360")),
            Decimal("0.1800"),
        )
        self.assertEqual(
            calculate_user_annual_rate("0", "0", "0.05", Decimal("365")),
            Decimal("0.1825"),
        )

    def test_user_rate_is_highest_annualized_rate(self):
        self.assertEqual(
            calculate_user_annual_rate("10", "1", "0.02", Decimal("360")),
            Decimal("0.12"),
        )

    def test_normalize_year_days_accepts_display_values(self):
        self.assertEqual(normalize_year_days("360天"), Decimal("360"))
        self.assertEqual(normalize_year_days("365"), Decimal("365"))


class ExportDetailTests(unittest.TestCase):
    def test_extracts_only_current_interest_calculation_section(self):
        detail = "\n".join(
            [
                "第1条 还款记录",
                "流水日期: 2024-02-01",
                "",
                "本期利息计算:",
                "2024-01-01 至 2024-02-01:",
                "  适用上限: 无单独上限",
                "  计算公式: 10000.00 × 12.0000% ÷ 360 × 31天 = 103.33",
                "",
                "本期新增应付利息: 103.33",
                "本条为还款，按先息后本处理。",
                "实付利息: 103.33",
            ]
        )

        result = extract_current_interest_calculation(detail)

        self.assertIn("本期利息计算:", result)
        self.assertIn("计算公式", result)
        self.assertNotIn("流水日期", result)
        self.assertNotIn("本期新增应付利息", result)
        self.assertNotIn("实付利息", result)


if __name__ == "__main__":
    unittest.main()
