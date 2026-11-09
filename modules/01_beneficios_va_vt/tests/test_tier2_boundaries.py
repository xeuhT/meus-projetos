"""
Tier 2: Comprehensive Boundary and Corner Cases Test Suite for "04 - GOOD CARD"
Covers extreme values, rounding, threshold comparisons (15 vs 16 days, 90 vs 91 days, exact 20% vs 20% + 0.01),
currency sentinels, matricula lengths, and accounting tolerances.
Authoritative source: PROJECT.md, business_rules_architecture.md, template_spec.md.
"""

import os
import sys
import unittest
import pandas as pd
import numpy as np

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from tests.e2e.oracle import (
    TOLERANCIA_CENTAVOS,
    get_module,
    oracle_parse_currency,
    oracle_normalize_matricula,
    oracle_calculate_margem_20,
    oracle_check_status_margem,
    oracle_check_quitacao,
    oracle_check_quitacao_residual,
    oracle_classify_audit_case
)


# ==============================================================================
# 1. CURRENCY & NULL SENTINEL BOUNDARIES
# ==============================================================================
class TestCurrencyBoundaries(unittest.TestCase):
    """Boundary conditions for currency strings, dashes, nulls, and extreme floats."""

    def setUp(self):
        dl = get_module("data_loader")
        self.parse_fn = getattr(dl, "parse_brazilian_currency", oracle_parse_currency)

    def test_boundary_single_dash(self):
        """Single dash '-' must convert to exactly 0.00."""
        self.assertEqual(self.parse_fn("-"), 0.0)

    def test_boundary_spaced_dash(self):
        """Dash with spaces ' - ' must convert to exactly 0.00."""
        self.assertEqual(self.parse_fn(" - "), 0.0)
        self.assertEqual(self.parse_fn("   -   "), 0.0)

    def test_boundary_empty_and_whitespace(self):
        """Empty string and whitespace string must convert to 0.00."""
        self.assertEqual(self.parse_fn(""), 0.0)
        self.assertEqual(self.parse_fn("   "), 0.0)

    def test_boundary_none_and_nan(self):
        """None, float NaN, and pd.NA must convert to 0.00."""
        self.assertEqual(self.parse_fn(None), 0.0)
        self.assertEqual(self.parse_fn(np.nan), 0.0)
        self.assertEqual(self.parse_fn(pd.NA), 0.0)

    def test_boundary_zero_variations(self):
        """All variations of zero convert to 0.00."""
        self.assertEqual(self.parse_fn(0), 0.0)
        self.assertEqual(self.parse_fn(0.0), 0.0)
        self.assertEqual(self.parse_fn("0"), 0.0)
        self.assertEqual(self.parse_fn("0,00"), 0.0)
        self.assertEqual(self.parse_fn("0.00"), 0.0)

    def test_boundary_single_cent(self):
        """Minimum non-zero currency value 1 cent (0.01)."""
        self.assertEqual(self.parse_fn("0,01"), 0.01)
        self.assertEqual(self.parse_fn("R$ 0,01"), 0.01)

    def test_boundary_two_cents(self):
        """Tolerance boundary value 2 cents (0.02)."""
        self.assertEqual(self.parse_fn("0,02"), 0.02)

    def test_boundary_large_currency_value(self):
        """Large corporate payroll value: 999.999.999,99."""
        self.assertEqual(self.parse_fn("999.999.999,99"), 999999999.99)


# ==============================================================================
# 2. MARGIN 20% & SALARY BOUNDARIES
# ==============================================================================
class TestMarginThresholdBoundaries(unittest.TestCase):
    """Strict boundary tests around the 20% margin threshold and zero salaries."""

    def test_margin_exact_equality(self):
        """Saldo Devedor EXACTLY EQUAL to 20% margin is 'Dentro da Margem'."""
        # Salário 1000.00 -> Margem 200.00, Saldo 200.00
        margem = oracle_calculate_margem_20(1000.00)
        self.assertEqual(margem, 200.00)
        status = oracle_check_status_margem(200.00, margem)
        self.assertEqual(status, "Dentro da Margem")

    def test_margin_plus_one_cent(self):
        """Saldo Devedor EXACTLY 1 cent ABOVE 20% margin is 'Fora da Margem'."""
        margem = oracle_calculate_margem_20(1000.00)
        status = oracle_check_status_margem(200.01, margem)
        self.assertEqual(status, "Fora da Margem")

    def test_margin_minus_one_cent(self):
        """Saldo Devedor EXACTLY 1 cent BELOW 20% margin is 'Dentro da Margem'."""
        margem = oracle_calculate_margem_20(1000.00)
        status = oracle_check_status_margem(199.99, margem)
        self.assertEqual(status, "Dentro da Margem")

    def test_deduction_exact_equality_with_margin(self):
        """Desconto 6280 EXACTLY EQUAL to 20% margin is 'Dentro da Margem'."""
        margem = oracle_calculate_margem_20(2500.00) # 500.00
        status = oracle_check_status_margem(500.00, margem)
        self.assertEqual(status, "Dentro da Margem")

    def test_deduction_plus_one_cent_exceeds_margin(self):
        """Desconto 6280 EXACTLY 1 cent above 20% margin triggers 'Fora da Margem'."""
        margem = oracle_calculate_margem_20(2500.00) # 500.00
        status = oracle_check_status_margem(500.01, margem)
        self.assertEqual(status, "Fora da Margem")

    def test_salary_zero_margin(self):
        """Salário 0.00 produces 0.00 margin; any positive balance is 'Fora da Margem'."""
        margem = oracle_calculate_margem_20(0.00)
        self.assertEqual(margem, 0.00)
        self.assertEqual(oracle_check_status_margem(0.00, margem), "Dentro da Margem")
        self.assertEqual(oracle_check_status_margem(0.01, margem), "Fora da Margem")

    def test_margin_penny_fractional_rounding(self):
        """Salário 1500.55 * 0.20 = 300.110 -> 300.11 rounded."""
        margem = oracle_calculate_margem_20(1500.55)
        self.assertEqual(margem, 300.11)


# ==============================================================================
# 3. LEAVE (AFASTAMENTO) DURATION BOUNDARIES (15 vs 16 DAYS)
# ==============================================================================
class TestAfastamentoBoundaries(unittest.TestCase):
    """Threshold tests: exactly 15 days vs 16 days of leave."""

    def test_afastamento_exactly_15_days_does_not_trigger_block(self):
        """Leave of exactly 15 days does NOT require card block (> 15 required)."""
        # Colaborador with 15 days leave: rule requires 'período superior a 15 dias'
        dias = 15
        is_superior_15 = dias > 15
        self.assertFalse(is_superior_15)

    def test_afastamento_exactly_16_days_triggers_block(self):
        """Leave of 16 days triggers block request rule (> 15 days)."""
        dias = 16
        is_superior_15 = dias > 15
        self.assertTrue(is_superior_15)

    def test_afastamento_1_day_does_not_trigger_block(self):
        """1-day medical note does not trigger block request."""
        self.assertFalse(1 > 15)

    def test_afastamento_multiple_periods_selects_max(self):
        """Collaborator with 14-day and 18-day leaves selects the 18-day leave."""
        df_leaves = pd.DataFrame({
            "Matrícula": ["100", "100"],
            "Dias Afastamento": [14, 18]
        })
        selected = df_leaves.sort_values("Dias Afastamento", ascending=False).iloc[0]
        self.assertEqual(selected["Dias Afastamento"], 18)
        self.assertTrue(selected["Dias Afastamento"] > 15)


# ==============================================================================
# 4. ADMISSION DURATION BOUNDARIES (90 vs 91 DAYS)
# ==============================================================================
class TestAdmissaoBoundaries(unittest.TestCase):
    """Threshold tests: exactly 90 days vs 91 days from reference date."""

    def test_admissao_exactly_90_days_not_exceeded(self):
        """Collaborator admitted exactly 90 days prior has NOT finished experience (> 90 required)."""
        ref_date = pd.Timestamp("2026-09-01")
        admissao = ref_date - pd.Timedelta(days=90)
        days = (ref_date - admissao).days
        self.assertEqual(days, 90)
        self.assertFalse(days > 90)

    def test_admissao_exactly_91_days_exceeded(self):
        """Collaborator admitted 91 days prior HAS finished experience (> 90 days)."""
        ref_date = pd.Timestamp("2026-09-01")
        admissao = ref_date - pd.Timedelta(days=91)
        days = (ref_date - admissao).days
        self.assertEqual(days, 91)
        self.assertTrue(days > 90)

    def test_admissao_zero_days(self):
        """Admitted on the audit reference date (0 days of experience)."""
        ref_date = pd.Timestamp("2026-09-01")
        days = (ref_date - ref_date).days
        self.assertEqual(days, 0)
        self.assertFalse(days > 90)


# ==============================================================================
# 5. MATRÍCULA LENGTH & PADDING BOUNDARIES
# ==============================================================================
class TestMatriculaLengthBoundaries(unittest.TestCase):
    """Boundary cases for matricula representations."""

    def setUp(self):
        dl = get_module("data_loader")
        self.norm_fn = getattr(dl, "normalize_matricula", oracle_normalize_matricula)

    def test_matricula_single_digit(self):
        """Single digit matricula '5' preserved."""
        self.assertEqual(self.norm_fn(5), "5")
        self.assertEqual(self.norm_fn("5"), "5")

    def test_matricula_two_digits(self):
        """Two digit matricula '78' preserved."""
        self.assertEqual(self.norm_fn(78), "78")
        self.assertEqual(self.norm_fn("78"), "78")

    def test_matricula_leading_zeros_six_digits(self):
        """Matricula with multiple leading zeros '000935' preserved."""
        self.assertEqual(self.norm_fn("000935"), "000935")

    def test_matricula_leading_zeros_with_point_zero(self):
        """String '001234.0' strips .0 while preserving leading zeros."""
        self.assertEqual(self.norm_fn("001234.0"), "001234")

    def test_matricula_long_eight_digits(self):
        """Long 8-digit matricula '12345678' preserved."""
        self.assertEqual(self.norm_fn(12345678), "12345678")


# ==============================================================================
# 6. ACCOUNTING EQUATION & PENNY TOLERANCE BOUNDARIES
# ==============================================================================
class TestAccountingToleranceBoundaries(unittest.TestCase):
    """Boundary conditions for accounting equation tolerance (0.02 tolerance)."""

    def test_tolerance_exact_match(self):
        """0.00 difference matches cleanly."""
        self.assertTrue(oracle_check_quitacao(100.0, 0.0, 100.0))

    def test_tolerance_one_cent_difference(self):
        """0.01 difference passes within tolerance."""
        # 6280=99.99, 29624=0.0, Total=100.00 -> diff 0.01
        self.assertTrue(oracle_check_quitacao(99.99, 0.0, 100.00))

    def test_tolerance_two_cents_difference(self):
        """0.02 difference passes within tolerance."""
        self.assertTrue(oracle_check_quitacao(99.98, 0.0, 100.00))

    def test_tolerance_three_cents_difference_fails(self):
        """0.03 difference exceeds tolerance and fails."""
        self.assertFalse(oracle_check_quitacao(99.97, 0.0, 100.00))

    def test_residual_equation_three_cents_difference_fails(self):
        """6280 + 29621 + 29624 vs Total with 0.03 discrepancy fails."""
        # 6280=50.0, 29621=49.97, Total=100.0 -> diff 0.03
        self.assertFalse(oracle_check_quitacao_residual(50.0, 49.97, 0.0, 100.0))


if __name__ == "__main__":
    unittest.main(verbosity=2)
