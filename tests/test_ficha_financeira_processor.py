from decimal import Decimal
import unittest

from processors.ficha_financeira_processor import FichaFinanceiraProcessor


class ApplyVacationAdjustmentsTest(unittest.TestCase):
    def test_applies_adjustment_when_single_vacation_code_has_value(self) -> None:
        processor = FichaFinanceiraProcessor()
        aggregated = {
            "173-Ferias": {},
            "174-Ferias": { (2024, 1): Decimal("2000") },
            "527-INSS-Comp": { (2024, 1): Decimal("3000") },
            "527-INSS-Valor": { (2024, 1): Decimal("300") },
        }

        processor._apply_vacation_adjustments(aggregated)

        base_values = aggregated.get("3123-Base", {})
        self.assertIn((2024, 1), base_values)
        self.assertEqual(Decimal("10"), base_values[(2024, 1)])

    def test_applies_adjustment_using_inss_months_when_vacation_values_zero(self) -> None:
        processor = FichaFinanceiraProcessor()
        aggregated = {
            "167-Ferias": { (2024, 2): Decimal("0") },
            "168-Ferias": { (2024, 2): Decimal("0") },
            "173-Ferias": { (2024, 2): Decimal("0") },
            "174-Ferias": { (2024, 2): Decimal("0") },
            "527-INSS-Comp": { (2024, 2): Decimal("3000") },
            "527-INSS-Valor": { (2024, 2): Decimal("300") },
        }

        processor._apply_vacation_adjustments(aggregated)

        base_values = aggregated.get("3123-Base", {})
        self.assertIn((2024, 2), base_values)
        self.assertEqual(Decimal("10"), base_values[(2024, 2)])


if __name__ == "__main__":
    unittest.main()
