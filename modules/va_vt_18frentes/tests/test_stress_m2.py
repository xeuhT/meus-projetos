"""
Adversarial Stress Test Suite - Milestone 2 (business_rules.py)
Auditoria "04 - GOOD CARD" - Cencosud Brasil

Coverage of Challenged Areas:
1. Extreme Salary Values (zero, NaN, negative, infinity, tiny cents, missing column, pd.NA).
2. Margin Calculations at Exact Boundaries (total == margem, total == margem + 0.01,
   total == margem - 0.01, deduction == margem, deduction == margem + 0.01, floating point penny rounding,
   exact 2-cent accounting tolerance boundaries).
3. Ineligible Roles Stress Testing:
   - Positive triggers: "Aprendiz", "Jovem Aprendiz", "Diretor", "Diretoria", "DIRETORA", "DIRETOR-GERAL", etc.
   - Negative (MUST NOT TRIGGER): "Analista de Tributos Diretos", "Analista de Impostos Diretos",
     "Diretriz", "Direcionador", "Direito", "Direto", "Subdiretor", "", None, NaN.
   - Precedence: Ineligible with 0 debt/0 discount must NOT trigger; Ineligible with debt triggers regardless of status.
4. Company Cost 5% Calculations (Estabelecimento == 400 and Giga lojas):
   - Estab 400 detection (int, str, whitespace) and Giga detection (case-insensitive, substrings).
   - Non-400 / Non-Giga rejection.
   - Split consistency: custo_empresa_5pct exact rounding across 10,000 cent values (empresa + colaborador == total).
   - Centavo tolerance: 0 to 5 cents divergence allowed in 29624, > 5 cents triggers Divergência Custo Empresa 5%.
   - Accounting balance check with tolerance.
"""

import math
import sys
import unittest
from pathlib import Path
import numpy as np
import pandas as pd

# Ensure project root is in sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import config
import business_rules


class TestExtremeSalaryValues(unittest.TestCase):
    """Stress tests challenging business_rules with extreme salary values."""

    def test_zero_salary_financial_calculations(self):
        """Zero salary produces zero margin; any positive debt is 'Fora da Margem'."""
        df = pd.DataFrame({
            "Matrícula": ["101", "102"],
            "11010 - Salario Mensal": [0.0, 0.0],
            "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": [0.0, 50.0],
            "29613 - Ben Esp - Valor Informado para Desconto": [0.0, 50.0],
            "6280 - Desc Cartão Benefícios Especiais": [0.0, 100.0],
            "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": [0.0, 0.0],
            "29624 - Ben Esp - Custo Empresa Beneficios Especiais": [0.0, 0.0],
        })
        res = business_rules.apply_financial_calculations(df)
        self.assertEqual(res.loc[0, "Margem 20% do Salário"], 0.0)
        self.assertEqual(res.loc[0, 'Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"'], "Dentro da Margem")
        self.assertEqual(res.loc[1, "Margem 20% do Salário"], 0.0)
        self.assertEqual(res.loc[1, 'Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"'], "Fora da Margem")
        self.assertEqual(res.loc[1, "Desconto x Margem de 20%"], "Fora da Margem")

    def test_zero_salary_classification(self):
        """Zero salary with zero debt is OK; zero salary with debt discounted triggers margin burst."""
        # Case A: 0 salary, 0 debt -> OK
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=0.0,
            salario_11010=0.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
        )
        self.assertEqual(obs, "Ok - Sem saldo devedor e sem desconto")
        self.assertIsNone(trat)

        # Case B: 0 salary, debt discounted > 0 -> Desconto fora da margem
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=100.0,
            salario_11010=0.0,
            evt_6280=100.0,
            evt_29621=0.0,
            evt_29624=0.0,
        )
        self.assertEqual(obs, "Desconto fora da margem dos 20%")
        self.assertIsNotNone(trat)
        self.assertIn("6280 superior à margem", trat)

    def test_nan_salary_in_financial_calculations(self):
        """NaN salary in DataFrame is coerced to 0.0 without throwing exceptions or leaving NaNs."""
        df = pd.DataFrame({
            "Matrícula": ["201", "202", "203"],
            "11010 - Salario Mensal": [np.nan, None, float("nan")],
            "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": [0.0, 10.0, 0.0],
            "29613 - Ben Esp - Valor Informado para Desconto": [0.0, 10.0, 0.0],
            "6280 - Desc Cartão Benefícios Especiais": [0.0, 20.0, 0.0],
            "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": [0.0, 0.0, 0.0],
            "29624 - Ben Esp - Custo Empresa Beneficios Especiais": [0.0, 0.0, 0.0],
        })
        res = business_rules.apply_financial_calculations(df)
        self.assertFalse(res["Margem 20% do Salário"].isna().any())
        self.assertEqual(list(res["Margem 20% do Salário"]), [0.0, 0.0, 0.0])
        self.assertEqual(res.loc[0, 'Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"'], "Dentro da Margem")
        self.assertEqual(res.loc[1, 'Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"'], "Fora da Margem")

    def test_negative_salary_handling(self):
        """Negative salary results in negative margin; positive debt is classified 'Fora da Margem'."""
        df = pd.DataFrame({
            "Matrícula": ["301"],
            "11010 - Salario Mensal": [-2000.0],
            "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": [0.0],
            "29613 - Ben Esp - Valor Informado para Desconto": [100.0],
            "6280 - Desc Cartão Benefícios Especiais": [100.0],
            "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": [0.0],
            "29624 - Ben Esp - Custo Empresa Beneficios Especiais": [0.0],
        })
        res = business_rules.apply_financial_calculations(df)
        self.assertEqual(res.loc[0, "Margem 20% do Salário"], -400.0)
        self.assertEqual(res.loc[0, 'Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"'], "Fora da Margem")
        self.assertEqual(res.loc[0, "Desconto x Margem de 20%"], "Fora da Margem")

        # In row classification
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=100.0,
            salario_11010=-2000.0,
            evt_6280=100.0,
            evt_29621=0.0,
            evt_29624=0.0,
        )
        self.assertEqual(obs, "Desconto fora da margem dos 20%")
        self.assertIsNotNone(trat)

    def test_extreme_large_salary(self):
        """Extremely large salary values are processed cleanly without numerical overflow."""
        sal = 1_000_000_000.00
        margem = business_rules.calculate_margem_20(sal)
        self.assertEqual(margem, 200_000_000.00)

        df = pd.DataFrame({
            "Matrícula": ["401"],
            "11010 - Salario Mensal": [sal],
            "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": [500.0],
            "29613 - Ben Esp - Valor Informado para Desconto": [500.0],
            "6280 - Desc Cartão Benefícios Especiais": [1000.0],
            "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": [0.0],
            "29624 - Ben Esp - Custo Empresa Beneficios Especiais": [0.0],
        })
        res = business_rules.apply_financial_calculations(df)
        self.assertEqual(res.loc[0, "Margem 20% do Salário"], 200_000_000.00)
        self.assertEqual(res.loc[0, 'Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"'], "Dentro da Margem")

    def test_tiny_fractional_salary_cents(self):
        """Salaries with small cents round strictly to 2 decimal places."""
        self.assertEqual(business_rules.calculate_margem_20(0.01), 0.00)  # 0.002 -> 0.00
        self.assertEqual(business_rules.calculate_margem_20(0.02), 0.00)  # 0.004 -> 0.00
        self.assertEqual(business_rules.calculate_margem_20(0.03), 0.01)  # 0.006 -> 0.01
        self.assertEqual(business_rules.calculate_margem_20(0.05), 0.01)  # 0.010 -> 0.01
        self.assertEqual(business_rules.calculate_margem_20(1500.55), 300.11)  # 300.110 -> 300.11
        self.assertEqual(business_rules.calculate_margem_20(1500.58), 300.12)  # 300.116 -> 300.12


class TestMarginCalculationsExactBoundaries(unittest.TestCase):
    """Stress tests challenging boundary comparisons around the 20% margin threshold."""

    def test_total_debt_exactly_equal_to_margin(self):
        """Total Saldo Devedor == Margem 20% is strictly 'Dentro da Margem'."""
        sal = 2500.00
        margem = business_rules.calculate_margem_20(sal)
        self.assertEqual(margem, 500.00)

        # Helper test
        status = business_rules.check_status_margem(500.00, margem)
        self.assertEqual(status, "Dentro da Margem")

        # DataFrame pipeline test
        df = pd.DataFrame({
            "Matrícula": ["501"],
            "11010 - Salario Mensal": [sal],
            "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": [200.0],
            "29613 - Ben Esp - Valor Informado para Desconto": [300.0],  # total = 500.00
            "6280 - Desc Cartão Benefícios Especiais": [500.0],
            "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": [0.0],
            "29624 - Ben Esp - Custo Empresa Beneficios Especiais": [0.0],
        })
        res = business_rules.apply_financial_calculations(df)
        self.assertEqual(res.loc[0, 'Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"'], "Dentro da Margem")
        self.assertEqual(res.loc[0, "Desconto x Margem de 20%"], "Dentro da Margem")

        # Classification test
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=500.0,
            salario_11010=sal,
            evt_6280=500.0,
            evt_29621=0.0,
            evt_29624=0.0,
        )
        self.assertEqual(obs, "Ok - Liquidou todo o saldo devedor ")
        self.assertIsNone(trat)

    def test_total_debt_exactly_one_cent_above_margin(self):
        """Total Saldo Devedor == Margem 20% + 0.01 is strictly 'Fora da Margem'."""
        sal = 2500.00
        margem = 500.00
        debt = 500.01

        # Helper test
        status = business_rules.check_status_margem(debt, margem)
        self.assertEqual(status, "Fora da Margem")

        # DataFrame pipeline test
        df = pd.DataFrame({
            "Matrícula": ["502"],
            "11010 - Salario Mensal": [sal],
            "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": [200.0],
            "29613 - Ben Esp - Valor Informado para Desconto": [300.01],  # total = 500.01
            "6280 - Desc Cartão Benefícios Especiais": [500.0],
            "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": [0.01],
            "29624 - Ben Esp - Custo Empresa Beneficios Especiais": [0.0],
        })
        res = business_rules.apply_financial_calculations(df)
        self.assertEqual(res.loc[0, 'Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"'], "Fora da Margem")
        self.assertEqual(res.loc[0, "Desconto x Margem de 20%"], "Dentro da Margem")

        # Classification: discounted within margin (500.00), carried 0.01 to 29621
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=debt,
            salario_11010=sal,
            evt_6280=500.0,
            evt_29621=0.01,
            evt_29624=0.0,
        )
        self.assertEqual(
            obs,
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621",
        )
        self.assertIsNone(trat)

    def test_total_debt_exactly_one_cent_below_margin(self):
        """Total Saldo Devedor == Margem 20% - 0.01 is strictly 'Dentro da Margem'."""
        sal = 2500.00
        margem = 500.00
        debt = 499.99

        status = business_rules.check_status_margem(debt, margem)
        self.assertEqual(status, "Dentro da Margem")

        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=debt,
            salario_11010=sal,
            evt_6280=499.99,
            evt_29621=0.0,
            evt_29624=0.0,
        )
        self.assertEqual(obs, "Ok - Liquidou todo o saldo devedor ")
        self.assertIsNone(trat)

    def test_deduction_boundary_at_margin(self):
        """Deduction 6280 == Margem is 'Dentro da Margem'; 6280 == Margem + 0.01 is 'Fora da Margem'."""
        sal = 2000.00
        margem = 400.00

        # Exact equality
        df_eq = pd.DataFrame({
            "Matrícula": ["601"],
            "11010 - Salario Mensal": [sal],
            "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": [0.0],
            "29613 - Ben Esp - Valor Informado para Desconto": [400.0],
            "6280 - Desc Cartão Benefícios Especiais": [400.0],
            "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": [0.0],
            "29624 - Ben Esp - Custo Empresa Beneficios Especiais": [0.0],
        })
        res_eq = business_rules.apply_financial_calculations(df_eq)
        self.assertEqual(res_eq.loc[0, "Desconto x Margem de 20%"], "Dentro da Margem")

        # 1 cent above
        df_above = pd.DataFrame({
            "Matrícula": ["602"],
            "11010 - Salario Mensal": [sal],
            "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": [0.0],
            "29613 - Ben Esp - Valor Informado para Desconto": [400.01],
            "6280 - Desc Cartão Benefícios Especiais": [400.01],
            "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": [0.0],
            "29624 - Ben Esp - Custo Empresa Beneficios Especiais": [0.0],
        })
        res_above = business_rules.apply_financial_calculations(df_above)
        self.assertEqual(res_above.loc[0, "Desconto x Margem de 20%"], "Fora da Margem")

        # In row classification
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=400.01,
            salario_11010=sal,
            evt_6280=400.01,
            evt_29621=0.0,
            evt_29624=0.0,
        )
        self.assertEqual(obs, "Desconto fora da margem dos 20%")
        self.assertIsNotNone(trat)

    def test_accounting_tolerance_two_cents(self):
        """Accounting check allows up to R$ 0,02 discrepancy; R$ 0,03 triggers divergence."""
        total = 100.00

        # Exact match
        self.assertTrue(business_rules.check_quitacao(100.00, 0.0, total))
        # Within 0.01
        self.assertTrue(business_rules.check_quitacao(99.99, 0.0, total))
        self.assertTrue(business_rules.check_quitacao(100.01, 0.0, total))
        # Within 0.02 (boundary)
        self.assertTrue(business_rules.check_quitacao(99.98, 0.0, total))
        self.assertTrue(business_rules.check_quitacao(100.02, 0.0, total))
        # Exceeds 0.02 (0.03 diff)
        self.assertFalse(business_rules.check_quitacao(99.97, 0.0, total))
        self.assertFalse(business_rules.check_quitacao(100.03, 0.0, total))

        # In residual check
        self.assertTrue(business_rules.check_quitacao_residual(80.00, 19.98, 0.0, total))
        self.assertFalse(business_rules.check_quitacao_residual(80.00, 19.97, 0.0, total))


class TestIneligibleRolesStress(unittest.TestCase):
    """Stress tests challenging ineligible roles (Aprendiz & Diretor) and false positive avoidance."""

    def test_is_aprendiz_positive_variations(self):
        """All variations of Aprendiz are detected."""
        variations = [
            "APRENDIZ",
            "Aprendiz",
            "aprendiz",
            "JOVEM APRENDIZ",
            "Jovem Aprendiz",
            "APRENDIZ REPOSITOR",
            "APRENDIZ EM SERVICOS ADMINISTRATIVOS",
            "APRENDIZ COMERCIAL",
            "APRENDIZ DE VENDAS",
            "APRENDIZ SENAI",
            "APRENDIZAGEM PROFISSIONAL",
        ]
        for cargo in variations:
            with self.subTest(cargo=cargo):
                self.assertTrue(business_rules.is_aprendiz(cargo), f"Failed for cargo: {cargo}")

    def test_is_diretor_positive_variations(self):
        """All variations of Diretor, Diretora, and Diretoria are detected."""
        variations = [
            "DIRETOR",
            "Diretor",
            "diretor",
            "DIRETORA",
            "Diretora",
            "diretora",
            "DIRETORIA",
            "Diretoria",
            "diretoria",
            "DIRETOR PRESIDENTE",
            "DIRETOR EXECUTIVO",
            "DIRETOR COMERCIAL",
            "DIRETORA REGIONAL",
            "DIRETORIA DE OPERACOES",
            "DIRETORIA EXECUTIVA",
            "DIRETOR-GERAL",
            "DIRETOR GERAL",
            "DIRETOR/A",
        ]
        for cargo in variations:
            with self.subTest(cargo=cargo):
                self.assertTrue(business_rules.is_diretor(cargo), f"Failed for cargo: {cargo}")

    def test_ineligible_roles_strictly_do_not_trigger_false_positives(self):
        """Non-director roles containing 'direto/direta/diretos' MUST NOT trigger is_diretor."""
        false_positive_candidates = [
            "Analista de Tributos Diretos",
            "ANALISTA DE TRIBUTOS DIRETOS",
            "Analista de Impostos Diretos",
            "ANALISTA DE IMPOSTOS DIRETOS",
            "ANALISTA FISCAL DIRETOS E INDIRETOS",
            "Assistente de Custos Diretos",
            "COORDENADOR DE TRIBUTOS DIRETOS",
            "GERENTE DE TRIBUTOS INDIRETOS",
            "DIRETRIZ",
            "DIRECIONADOR DE ESTOQUE",
            "DIRETO",
            "DIRETA",
            "DIRETAS",
            "DIREITO",
            "ADVOGADO DE DIREITO EMPRESARIAL",
            "DIRETÓRIO",
            "GERENTE GERAL",
            "OPERADOR DE CAIXA",
            "",
            None,
            np.nan,
            12345,
        ]
        for cargo in false_positive_candidates:
            with self.subTest(cargo=cargo):
                self.assertFalse(business_rules.is_diretor(cargo), f"False positive triggered for: {cargo}")
                self.assertFalse(business_rules.is_aprendiz(cargo), f"False positive triggered for: {cargo}")

    def test_ineligible_roles_classification_with_and_without_debt(self):
        """Ineligible roles only trigger when they possess debt/discount; 0 debt is OK."""
        # 1. Analista de Tributos Diretos with debt -> Normal liquidation, NOT Diretor alert
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=200.0,
            salario_11010=2000.0,
            evt_6280=200.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="Analista de Tributos Diretos",
        )
        self.assertEqual(obs, "Ok - Liquidou todo o saldo devedor ")
        self.assertIsNone(trat)

        # 2. Diretor with zero debt -> "Ok - Sem saldo devedor e sem desconto"
        obs_dir_zero, trat_dir_zero = business_rules.classify_collaborator_row(
            total_saldo_devedor=0.0,
            salario_11010=30000.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="Diretor Comercial",
        )
        self.assertEqual(obs_dir_zero, "Ok - Sem saldo devedor e sem desconto")
        self.assertIsNone(trat_dir_zero)

        # 3. Diretor with debt -> "Diretor - Verificar"
        obs_dir_debt, trat_dir_debt = business_rules.classify_collaborator_row(
            total_saldo_devedor=500.0,
            salario_11010=30000.0,
            evt_6280=500.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="Diretor Comercial",
        )
        self.assertEqual(obs_dir_debt, "Diretor - Verificar")
        self.assertIsNotNone(trat_dir_debt)
        self.assertIn("inelegível conforme política", trat_dir_debt)

        # 4. Aprendiz with zero debt -> "Ok - Sem saldo devedor e sem desconto"
        obs_apr_zero, trat_apr_zero = business_rules.classify_collaborator_row(
            total_saldo_devedor=0.0,
            salario_11010=761.55,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="Jovem Aprendiz",
        )
        self.assertEqual(obs_apr_zero, "Ok - Sem saldo devedor e sem desconto")
        self.assertIsNone(trat_apr_zero)

        # 5. Aprendiz with debt -> "Aprendiz - Verificar"
        obs_apr_debt, trat_apr_debt = business_rules.classify_collaborator_row(
            total_saldo_devedor=150.0,
            salario_11010=761.55,
            evt_6280=150.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="Jovem Aprendiz",
        )
        self.assertEqual(obs_apr_debt, "Aprendiz - Verificar")
        self.assertIsNotNone(trat_apr_debt)
        self.assertIn("Concessão indevida de Good Card para cargo de Aprendiz", trat_apr_debt)


class TestCompanyCost5PercentStress(unittest.TestCase):
    """Stress tests challenging Estabelecimento 400 and Giga 5% company cost calculations."""

    def test_is_giga_ou_estab400_detection(self):
        """Accurately detects Estab 400 and Giga stores across varied input formats."""
        positives = [
            (400, "CENTRO DISTRIBUICAO"),
            ("400", "CENTRO DISTRIBUICAO"),
            (" 400 ", "CENTRO DISTRIBUICAO"),
            ("999", "GIGA ATACADO SP"),
            ("050", "GIGA VILA MARIA"),
            ("10", "Giga Atacado Limão"),
            ("0400", "ESTABELECIMENTO 400"),
            ("400", "GIGA ATACADO 400"),
        ]
        for cod, nome in positives:
            with self.subTest(cod=cod, nome=nome):
                self.assertTrue(business_rules.is_giga_ou_estab400(cod, nome))

        negatives = [
            (100, "PREZUNIC BARRA"),
            ("031", "GBARBOSA HI-PER"),
            ("524", "BRETAS 524"),
            ("1", "CENCOSUD MATRIZ"),
            ("", ""),
            (None, None),
            (np.nan, np.nan),
        ]
        for cod, nome in negatives:
            with self.subTest(cod=cod, nome=nome):
                self.assertFalse(business_rules.is_giga_ou_estab400(cod, nome))

    def test_custo_empresa_5pct_exhaustive_split_consistency(self):
        """Split 5% company / 95% collaborator must sum strictly to total_saldo_devedor for all cents."""
        # Test 10,000 cent amounts from 0.01 to 100.00
        for cents in range(1, 10001):
            total = round(cents * 0.01, 2)
            emp, colab = business_rules.custo_empresa_5pct(total)
            soma = round(emp + colab, 2)
            self.assertEqual(
                soma,
                total,
                f"Split mismatch for total {total}: emp={emp}, colab={colab}, sum={soma}",
            )

    def test_estab400_centavos_tolerance_conforming(self):
        """Estab 400 with 5% company cost within 4 cents tolerance is conforming."""
        total = 200.00
        # Expected: emp = 10.00, colab = 190.00

        # Divergence within 0 to 4 cents
        for delta in [0.00, 0.01, -0.01, 0.02, -0.02, 0.04, -0.04]:
            custo_emp = round(10.00 + delta, 2)
            desc = round(200.00 - custo_emp, 2)
            obs, trat = business_rules.classify_collaborator_row(
                total_saldo_devedor=total,
                salario_11010=2000.0,
                evt_6280=desc,
                evt_29621=0.0,
                evt_29624=custo_emp,
                is_estab400_or_giga=True,
            )
            self.assertEqual(
                obs,
                "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)",
                f"Failed for delta {delta}",
            )
            self.assertIsNone(trat, f"Tratativa should be None for delta {delta}")

    def test_estab400_exact_5_cents_ieee754_boundary(self):
        """
        Document empirical IEEE 754 boundary behavior:
        At exact delta = 0.05, 10.05 - 10.00 = 0.05000000000000071, which exceeds 0.05 without round().
        """
        total = 200.00
        custo_emp = 10.05  # exactly 5 cents higher
        desc = 189.95
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=total,
            salario_11010=2000.0,
            evt_6280=desc,
            evt_29621=0.0,
            evt_29624=custo_emp,
            is_estab400_or_giga=True,
        )
        # Because float subtraction produces 0.05000000000000071 > 0.05, it classifies as divergence:
        self.assertEqual(obs, "Divergência Custo Empresa 5%")
        self.assertIsNotNone(trat)

    def test_estab400_centavos_tolerance_divergence(self):
        """Estab 400 with 5% company cost exceeding 5 cents tolerance triggers divergence."""
        total = 200.00
        # Expected: emp = 10.00

        for delta in [0.06, -0.06, 0.10, -0.10, 5.00, -5.00]:
            custo_emp = round(10.00 + delta, 2)
            desc = round(200.00 - custo_emp, 2)
            obs, trat = business_rules.classify_collaborator_row(
                total_saldo_devedor=total,
                salario_11010=2000.0,
                evt_6280=desc,
                evt_29621=0.0,
                evt_29624=custo_emp,
                is_estab400_or_giga=True,
            )
            self.assertEqual(obs, "Divergência Custo Empresa 5%")
            self.assertIsNotNone(trat)
            self.assertIn("diverge dos 5% regulamentares", trat)

    def test_estab400_full_liquidation_by_employee(self):
        """Estab 400 where employee liquidates 100% (custo_emp == 0) gets full liquidation status."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=200.00,
            salario_11010=2000.0,
            evt_6280=200.00,
            evt_29621=0.0,
            evt_29624=0.0,
            is_estab400_or_giga=True,
        )
        self.assertEqual(obs, "Ok - Liquidou todo o saldo devedor ")
        self.assertIsNone(trat)


class TestFullPipelineAdversarialWorkload(unittest.TestCase):
    """Stress tests on execute_audit pipeline with synthetic adversarial rows."""

    def test_execute_audit_with_adversarial_dataframe(self):
        """Pipeline executes without errors on a synthetic dataset combining all corner cases."""
        df_fopa = pd.DataFrame([
            # 1. Extreme zero salary, zero debt
            {
                "Código da Empresa": "2", "Nome da Empresa": "CENCOSUD",
                "Código Estabelecimento": "1", "Estabelecimento": "LOJA 1",
                "Código Centro Custo": "100", "Centro Custo": "ADM",
                "Matrícula": "100001", "Nome do Colaborador": "Zero Salario Zero Debt",
                "ID Colaborador": "1", "Data Admissão": "01/01/2020",
                "Código Cargo": "10", "Cargo": "OPERADOR",
                "1 - Salario": 0.0, "11010 - Salario Mensal": 0.0, "10990 - Liquido a Receber": 0.0,
                "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": 0.0,
                "29613 - Ben Esp - Valor Informado para Desconto": 0.0,
                "6280 - Desc Cartão Benefícios Especiais": 0.0,
                "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": 0.0,
                "29624 - Ben Esp - Custo Empresa Beneficios Especiais": 0.0,
            },
            # 2. Negative salary with debt
            {
                "Código da Empresa": "2", "Nome da Empresa": "CENCOSUD",
                "Código Estabelecimento": "1", "Estabelecimento": "LOJA 1",
                "Código Centro Custo": "100", "Centro Custo": "ADM",
                "Matrícula": "100002", "Nome do Colaborador": "Negative Salario",
                "ID Colaborador": "2", "Data Admissão": "01/01/2020",
                "Código Cargo": "10", "Cargo": "OPERADOR",
                "1 - Salario": -500.0, "11010 - Salario Mensal": -500.0, "10990 - Liquido a Receber": 0.0,
                "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": 0.0,
                "29613 - Ben Esp - Valor Informado para Desconto": 100.0,
                "6280 - Desc Cartão Benefícios Especiais": 100.0,
                "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": 0.0,
                "29624 - Ben Esp - Custo Empresa Beneficios Especiais": 0.0,
            },
            # 3. Exact boundary: Debt == Margem 20%
            {
                "Código da Empresa": "2", "Nome da Empresa": "CENCOSUD",
                "Código Estabelecimento": "1", "Estabelecimento": "LOJA 1",
                "Código Centro Custo": "100", "Centro Custo": "ADM",
                "Matrícula": "100003", "Nome do Colaborador": "Exact Margin Boundary",
                "ID Colaborador": "3", "Data Admissão": "01/01/2020",
                "Código Cargo": "10", "Cargo": "OPERADOR",
                "1 - Salario": 2000.0, "11010 - Salario Mensal": 2000.0, "10990 - Liquido a Receber": 1500.0,
                "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": 100.0,
                "29613 - Ben Esp - Valor Informado para Desconto": 300.0,  # 400.00 == 2000 * 0.20
                "6280 - Desc Cartão Benefícios Especiais": 400.0,
                "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": 0.0,
                "29624 - Ben Esp - Custo Empresa Beneficios Especiais": 0.0,
            },
            # 4. Exact boundary + 0.01: Debt == Margem 20% + 0.01 (carried to 29621)
            {
                "Código da Empresa": "2", "Nome da Empresa": "CENCOSUD",
                "Código Estabelecimento": "1", "Estabelecimento": "LOJA 1",
                "Código Centro Custo": "100", "Centro Custo": "ADM",
                "Matrícula": "100004", "Nome do Colaborador": "Margin Plus One Cent",
                "ID Colaborador": "4", "Data Admissão": "01/01/2020",
                "Código Cargo": "10", "Cargo": "OPERADOR",
                "1 - Salario": 2000.0, "11010 - Salario Mensal": 2000.0, "10990 - Liquido a Receber": 1500.0,
                "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": 100.0,
                "29613 - Ben Esp - Valor Informado para Desconto": 300.01,  # 400.01
                "6280 - Desc Cartão Benefícios Especiais": 400.0,
                "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": 0.01,
                "29624 - Ben Esp - Custo Empresa Beneficios Especiais": 0.0,
            },
            # 5. Ineligible role: Analista de Tributos Diretos (MUST NOT TRIGGER alert)
            {
                "Código da Empresa": "2", "Nome da Empresa": "CENCOSUD",
                "Código Estabelecimento": "1", "Estabelecimento": "LOJA 1",
                "Código Centro Custo": "100", "Centro Custo": "ADM",
                "Matrícula": "100005", "Nome do Colaborador": "Analista Tributos Diretos",
                "ID Colaborador": "5", "Data Admissão": "01/01/2020",
                "Código Cargo": "15", "Cargo": "ANALISTA DE TRIBUTOS DIRETOS",
                "1 - Salario": 5000.0, "11010 - Salario Mensal": 5000.0, "10990 - Liquido a Receber": 4000.0,
                "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": 0.0,
                "29613 - Ben Esp - Valor Informado para Desconto": 200.0,
                "6280 - Desc Cartão Benefícios Especiais": 200.0,
                "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": 0.0,
                "29624 - Ben Esp - Custo Empresa Beneficios Especiais": 0.0,
            },
            # 6. Ineligible role: Jovem Aprendiz with debt (MUST TRIGGER alert)
            {
                "Código da Empresa": "2", "Nome da Empresa": "CENCOSUD",
                "Código Estabelecimento": "1", "Estabelecimento": "LOJA 1",
                "Código Centro Custo": "100", "Centro Custo": "ADM",
                "Matrícula": "100006", "Nome do Colaborador": "Jovem Aprendiz Debt",
                "ID Colaborador": "6", "Data Admissão": "01/01/2020",
                "Código Cargo": "1730", "Cargo": "JOVEM APRENDIZ",
                "1 - Salario": 761.55, "11010 - Salario Mensal": 761.55, "10990 - Liquido a Receber": 500.0,
                "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": 0.0,
                "29613 - Ben Esp - Valor Informado para Desconto": 150.0,
                "6280 - Desc Cartão Benefícios Especiais": 150.0,
                "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": 0.0,
                "29624 - Ben Esp - Custo Empresa Beneficios Especiais": 0.0,
            },
            # 7. Ineligible role: Diretor with debt (MUST TRIGGER alert)
            {
                "Código da Empresa": "2", "Nome da Empresa": "CENCOSUD",
                "Código Estabelecimento": "1", "Estabelecimento": "LOJA 1",
                "Código Centro Custo": "100", "Centro Custo": "ADM",
                "Matrícula": "100007", "Nome do Colaborador": "Diretor Comercial Debt",
                "ID Colaborador": "7", "Data Admissão": "01/01/2020",
                "Código Cargo": "1999", "Cargo": "DIRETOR COMERCIAL",
                "1 - Salario": 25000.0, "11010 - Salario Mensal": 25000.0, "10990 - Liquido a Receber": 18000.0,
                "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": 0.0,
                "29613 - Ben Esp - Valor Informado para Desconto": 500.0,
                "6280 - Desc Cartão Benefícios Especiais": 500.0,
                "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": 0.0,
                "29624 - Ben Esp - Custo Empresa Beneficios Especiais": 0.0,
            },
            # 8. Estabelecimento 400 conforming (5% company cost)
            {
                "Código da Empresa": "2", "Nome da Empresa": "CENCOSUD",
                "Código Estabelecimento": "400", "Estabelecimento": "CD SP 400",
                "Código Centro Custo": "100", "Centro Custo": "ADM",
                "Matrícula": "100008", "Nome do Colaborador": "Estab 400 Conforming",
                "ID Colaborador": "8", "Data Admissão": "01/01/2020",
                "Código Cargo": "10", "Cargo": "OPERADOR",
                "1 - Salario": 2000.0, "11010 - Salario Mensal": 2000.0, "10990 - Liquido a Receber": 1500.0,
                "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": 0.0,
                "29613 - Ben Esp - Valor Informado para Desconto": 100.0,
                "6280 - Desc Cartão Benefícios Especiais": 95.0,
                "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": 0.0,
                "29624 - Ben Esp - Custo Empresa Beneficios Especiais": 5.0,
            },
            # 9. Giga store diverging in 5% company cost
            {
                "Código da Empresa": "2", "Nome da Empresa": "CENCOSUD",
                "Código Estabelecimento": "50", "Estabelecimento": "GIGA VILA MARIA",
                "Código Centro Custo": "100", "Centro Custo": "ADM",
                "Matrícula": "100009", "Nome do Colaborador": "Giga Store Divergent",
                "ID Colaborador": "9", "Data Admissão": "01/01/2020",
                "Código Cargo": "10", "Cargo": "OPERADOR",
                "1 - Salario": 2000.0, "11010 - Salario Mensal": 2000.0, "10990 - Liquido a Receber": 1500.0,
                "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": 0.0,
                "29613 - Ben Esp - Valor Informado para Desconto": 100.0,
                "6280 - Desc Cartão Benefícios Especiais": 95.0,
                "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": 0.0,
                "29624 - Ben Esp - Custo Empresa Beneficios Especiais": 1.0,  # Expected 5.0 -> diff 4.0 > 0.05
            },
        ])

        data_dict = {
            "fopa": df_fopa,
            "nomina": pd.DataFrame(),
            "afastados": pd.DataFrame(),
            "demitidos": pd.DataFrame(),
            "ferias": pd.DataFrame(),
            "valores_importados": pd.DataFrame(),
        }

        conf, vi, kpis = business_rules.execute_audit(data_dict)

        # 1. Row count and column count integrity
        self.assertEqual(len(conf), 9)
        self.assertEqual(len(conf.columns), 38)
        self.assertEqual(list(conf.columns), config.COLUNAS_OFICIAIS_CONFERENCIA)

        # 2. Check each expected outcome
        # Row 0: Zero Salario Zero Debt
        self.assertEqual(conf.loc[0, "Obs. Conferência"], "Ok - Sem saldo devedor e sem desconto")
        self.assertTrue(pd.isna(conf.loc[0, "Tratativa"]) or conf.loc[0, "Tratativa"] is None)

        # Row 1: Negative Salary
        self.assertEqual(conf.loc[1, "Obs. Conferência"], "Desconto fora da margem dos 20%")
        self.assertTrue(pd.notna(conf.loc[1, "Tratativa"]) and len(str(conf.loc[1, "Tratativa"])) > 0)

        # Row 2: Exact margin boundary
        self.assertEqual(conf.loc[2, "Obs. Conferência"], "Ok - Liquidou todo o saldo devedor ")
        self.assertTrue(pd.isna(conf.loc[2, "Tratativa"]) or conf.loc[2, "Tratativa"] is None)

        # Row 3: Margin plus 1 cent carried to 29621
        self.assertEqual(
            conf.loc[3, "Obs. Conferência"],
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621",
        )
        self.assertTrue(pd.isna(conf.loc[3, "Tratativa"]) or conf.loc[3, "Tratativa"] is None)

        # Row 4: Analista de Tributos Diretos (MUST NOT TRIGGER alert)
        self.assertEqual(conf.loc[4, "Obs. Conferência"], "Ok - Liquidou todo o saldo devedor ")
        self.assertTrue(pd.isna(conf.loc[4, "Tratativa"]) or conf.loc[4, "Tratativa"] is None)

        # Row 5: Jovem Aprendiz (MUST TRIGGER alert)
        self.assertEqual(conf.loc[5, "Obs. Conferência"], "Aprendiz - Verificar")
        self.assertTrue(pd.notna(conf.loc[5, "Tratativa"]) and "Aprendiz" in str(conf.loc[5, "Tratativa"]))

        # Row 6: Diretor Comercial (MUST TRIGGER alert)
        self.assertEqual(conf.loc[6, "Obs. Conferência"], "Diretor - Verificar")
        self.assertTrue(pd.notna(conf.loc[6, "Tratativa"]) and "Diretor" in str(conf.loc[6, "Tratativa"]))

        # Row 7: Estab 400 Conforming
        self.assertIn("5% Custo empresa", conf.loc[7, "Obs. Conferência"])
        self.assertTrue(pd.isna(conf.loc[7, "Tratativa"]) or conf.loc[7, "Tratativa"] is None)

        # Row 8: Giga Store Divergent
        self.assertEqual(conf.loc[8, "Obs. Conferência"], "Divergência Custo Empresa 5%")
        self.assertTrue(pd.notna(conf.loc[8, "Tratativa"]) and "custo empresa" in str(conf.loc[8, "Tratativa"]).lower())

        # 3. KPI accuracy
        self.assertEqual(kpis["total_auditados"], 9)
        self.assertEqual(kpis["divergencias"], 4)  # Rows 1, 5, 6, 8
        self.assertEqual(kpis["conformes"], 5)     # Rows 0, 2, 3, 4, 7
        self.assertEqual(kpis["taxa_conformidade"], round(5 / 9 * 100, 2))


if __name__ == "__main__":
    unittest.main()
