"""
Adversarial Stress Test Suite - Milestone 2 Empirical Challenge
Agent: challenger_m2_3 (EMPIRICAL CHALLENGER)
"""

import sys
import unittest
from pathlib import Path
import numpy as np
import pandas as pd

PROJECT_ROOT = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\04 - GOOD CARD")
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import config
import business_rules


class TestAdversarialBoundaryMargins(unittest.TestCase):
    """Scenario 1: Boundary Margins (total == margem vs total == margem + 0.01)."""

    def test_exact_margin_equality(self):
        """total == margem must be classified 'Dentro da Margem' and liquidates cleanly."""
        sal = 1500.00
        margem = 300.00  # 20%
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=margem,
            salario_11010=sal,
            evt_6280=margem,
            evt_29621=0.0,
            evt_29624=0.0,
        )
        self.assertEqual(obs, "Ok - Liquidou todo o saldo devedor ")
        self.assertIsNone(trat)

        status = business_rules.check_status_margem(margem, margem)
        self.assertEqual(status, "Dentro da Margem")

    def test_one_cent_above_margin(self):
        """total == margem + 0.01 must be classified 'Fora da Margem'."""
        sal = 1500.00
        margem = 300.00
        total = 300.01

        status = business_rules.check_status_margem(total, margem)
        self.assertEqual(status, "Fora da Margem")

        # When collaborator discounts up to margin (300.00) and carries 0.01 in 29621
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=total,
            salario_11010=sal,
            evt_6280=300.00,
            evt_29621=0.01,
            evt_29624=0.0,
        )
        self.assertEqual(
            obs,
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621",
        )
        self.assertIsNone(trat)

    def test_one_cent_below_margin(self):
        """total == margem - 0.01 must be classified 'Dentro da Margem'."""
        sal = 1500.00
        margem = 300.00
        total = 299.99

        status = business_rules.check_status_margem(total, margem)
        self.assertEqual(status, "Dentro da Margem")

        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=total,
            salario_11010=sal,
            evt_6280=total,
            evt_29621=0.0,
            evt_29624=0.0,
        )
        self.assertEqual(obs, "Ok - Liquidou todo o saldo devedor ")
        self.assertIsNone(trat)

    def test_deduction_boundary_violations(self):
        """Deduction exceeding margin by 1 cent must be flagged 'Desconto fora da margem dos 20%'."""
        sal = 2000.00
        margem = 400.00
        total = 500.00

        # Exactly at margin is fine
        obs_eq, trat_eq = business_rules.classify_collaborator_row(
            total_saldo_devedor=total,
            salario_11010=sal,
            evt_6280=400.00,
            evt_29621=100.00,
            evt_29624=0.0,
        )
        self.assertEqual(
            obs_eq,
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621",
        )
        self.assertIsNone(trat_eq)

        # 1 cent above margin triggers violation
        obs_violation, trat_violation = business_rules.classify_collaborator_row(
            total_saldo_devedor=total,
            salario_11010=sal,
            evt_6280=400.01,
            evt_29621=99.99,
            evt_29624=0.0,
        )
        self.assertEqual(obs_violation, "Desconto fora da margem dos 20%")
        self.assertIsNotNone(trat_violation)
        self.assertIn("6280 superior à margem", trat_violation)

    def test_floating_point_rounding_across_fractional_salaries(self):
        """Verify margin calculation consistency across difficult float fractions."""
        salaries = [
            (1234.55, 246.91),
            (1234.56, 246.91),
            (1234.57, 246.91),
            (1234.58, 246.92),
            (1999.99, 400.00),
            (3333.33, 666.67),
            (5582.44, 1116.49),
            (761.55, 152.31),
        ]
        for sal, expected_margem in salaries:
            with self.subTest(sal=sal):
                calc_margem = business_rules.calculate_margem_20(sal)
                self.assertEqual(calc_margem, expected_margem)


class TestAdversarialIneligibleRoles(unittest.TestCase):
    """Scenario 2: Ineligible roles (Aprendiz, Diretor vs false positives)."""

    def test_positive_aprendiz_variations(self):
        """Verify regex/keyword recognition of Aprendiz variants."""
        cargos = [
            "APRENDIZ",
            "Aprendiz",
            "aprendiz",
            "JOVEM APRENDIZ",
            "Jovem Aprendiz",
            "APRENDIZ COMERCIAL",
            "APRENDIZ REPOSITOR",
            "APRENDIZ DE SERVIÇOS ADMINISTRATIVOS",
            "APRENDIZ SENAI",
        ]
        for c in cargos:
            with self.subTest(cargo=c):
                self.assertTrue(business_rules.is_aprendiz(c))
                obs, trat = business_rules.classify_collaborator_row(
                    total_saldo_devedor=50.0,
                    salario_11010=761.55,
                    evt_6280=50.0,
                    evt_29621=0.0,
                    evt_29624=0.0,
                    cargo=c,
                )
                self.assertEqual(obs, "Aprendiz - Verificar")
                self.assertIsNotNone(trat)

    def test_positive_diretor_variations(self):
        """Verify regex recognition of Diretor, Diretora, and Diretoria."""
        cargos = [
            "DIRETOR",
            "Diretor",
            "diretor",
            "DIRETORA",
            "Diretora",
            "diretora",
            "DIRETORIA",
            "Diretoria",
            "DIRETOR PRESIDENTE",
            "DIRETOR-PRESIDENTE",
            "DIRETOR EXECUTIVO",
            "DIRETOR COMERCIAL",
            "DIRETORA REGIONAL",
            "DIRETORIA DE OPERACOES",
            "DIRETOR GERAL",
            "DIRETOR/A",
        ]
        for c in cargos:
            with self.subTest(cargo=c):
                self.assertTrue(business_rules.is_diretor(c))
                obs, trat = business_rules.classify_collaborator_row(
                    total_saldo_devedor=100.0,
                    salario_11010=20000.0,
                    evt_6280=100.0,
                    evt_29621=0.0,
                    evt_29624=0.0,
                    cargo=c,
                )
                self.assertEqual(obs, "Diretor - Verificar")
                self.assertIsNotNone(trat)

    def test_negative_false_positives_must_not_trigger(self):
        """Verify compound titles like 'Analista de Tributos Diretos' do NOT trigger is_diretor."""
        cargos = [
            "Analista de Tributos Diretos",
            "ANALISTA DE TRIBUTOS DIRETOS",
            "Analista de Impostos Diretos",
            "ANALISTA DE IMPOSTOS DIRETOS",
            "Coordenador de Canais Diretos",
            "Assistente de Custos Diretos",
            "ANALISTA FISCAL DIRETOS E INDIRETOS",
            "DIRETRIZ",
            "DIRECIONADOR",
            "DIREITO",
            "DIRETO",
            "DIRETA",
            "DIRETÓRIO",
            "SUBDIRETOR",
            "GERENTE GERAL",
            "OPERADOR DE CAIXA",
        ]
        for c in cargos:
            with self.subTest(cargo=c):
                self.assertFalse(business_rules.is_diretor(c), f"False positive is_diretor on {c}")
                self.assertFalse(business_rules.is_aprendiz(c), f"False positive is_aprendiz on {c}")

                # With debt, compliant non-director must liquidate cleanly
                obs, trat = business_rules.classify_collaborator_row(
                    total_saldo_devedor=100.0,
                    salario_11010=2500.0,
                    evt_6280=100.0,
                    evt_29621=0.0,
                    evt_29624=0.0,
                    cargo=c,
                )
                self.assertEqual(obs, "Ok - Liquidou todo o saldo devedor ")
                self.assertIsNone(trat)

    def test_ineligible_role_zero_debt_precedence(self):
        """Ineligible role with 0 debt and 0 discount must NOT trigger alert."""
        obs_dir, trat_dir = business_rules.classify_collaborator_row(
            total_saldo_devedor=0.0,
            salario_11010=30000.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="DIRETOR PRESIDENTE",
        )
        self.assertEqual(obs_dir, "Ok - Sem saldo devedor e sem desconto")
        self.assertIsNone(trat_dir)

        obs_apr, trat_apr = business_rules.classify_collaborator_row(
            total_saldo_devedor=0.0,
            salario_11010=761.55,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="JOVEM APRENDIZ",
        )
        self.assertEqual(obs_apr, "Ok - Sem saldo devedor e sem desconto")
        self.assertIsNone(trat_apr)


class TestAdversarialGigaEstab400(unittest.TestCase):
    """Scenario 3: Giga / Estabelecimento 400 (100% liquidation vs 5% subsidy vs divergence)."""

    def test_giga_detection(self):
        """Accurate detection of Giga and Estabelecimento 400 across formats."""
        self.assertTrue(business_rules.is_giga_ou_estab400(400, "CENTRO DISTRIBUICAO"))
        self.assertTrue(business_rules.is_giga_ou_estab400("400", "LOJA"))
        self.assertTrue(business_rules.is_giga_ou_estab400(" 400 ", "LOJA"))
        self.assertTrue(business_rules.is_giga_ou_estab400("950", "GIGA ATACADO 950"))
        self.assertTrue(business_rules.is_giga_ou_estab400(50, "CENCOSUD BRASIL - GIGA VILA MARIA"))
        self.assertTrue(business_rules.is_giga_ou_estab400("0400", "ESTABELECIMENTO 400"))

        self.assertFalse(business_rules.is_giga_ou_estab400(100, "PREZUNIC"))
        self.assertFalse(business_rules.is_giga_ou_estab400("031", "GBARBOSA"))
        self.assertFalse(business_rules.is_giga_ou_estab400("524", "BRETAS 524"))

    def test_giga_full_employee_liquidation_100_percent(self):
        """Giga employee who pays 100% (custo_emp == 0) must be 'Ok - Liquidou todo o saldo devedor '."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=419.29,
            salario_11010=2312.45,
            evt_6280=419.29,
            evt_29621=0.0,
            evt_29624=0.0,
            is_estab400_or_giga=True,
        )
        self.assertEqual(obs, "Ok - Liquidou todo o saldo devedor ")
        self.assertIsNone(trat)

    def test_giga_conforming_5_percent_subsidy(self):
        """Giga employee with 5% company cost and 95% deduction within tolerance."""
        total = 500.00
        # 5% = 25.00, 95% = 475.00
        for delta in [0.00, 0.01, -0.01, 0.02, -0.02, 0.04, -0.04]:
            custo = round(25.00 + delta, 2)
            desc = round(500.00 - custo, 2)
            obs, trat = business_rules.classify_collaborator_row(
                total_saldo_devedor=total,
                salario_11010=3000.00,
                evt_6280=desc,
                evt_29621=0.0,
                evt_29624=custo,
                is_estab400_or_giga=True,
            )
            self.assertEqual(
                obs,
                "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)",
                f"Failed for delta {delta}",
            )
            self.assertIsNone(trat)

    def test_giga_genuine_5_percent_divergence(self):
        """Giga employee with 29624 diverging > 0.05 from 5% when saldo_rem == 0."""
        total = 500.00
        # Expected: 25.00
        for delta in [0.06, -0.06, 0.50, -0.50, 10.00, -10.00]:
            custo = round(25.00 + delta, 2)
            desc = round(500.00 - custo, 2)
            obs, trat = business_rules.classify_collaborator_row(
                total_saldo_devedor=total,
                salario_11010=3000.00,
                evt_6280=desc,
                evt_29621=0.0,
                evt_29624=custo,
                is_estab400_or_giga=True,
            )
            self.assertEqual(obs, "Divergência Custo Empresa 5%", f"Failed for delta {delta}")
            self.assertIsNotNone(trat)
            self.assertIn("diverge dos 5% regulamentares", trat)

    def test_giga_residual_carryover_with_debt_above_margin(self):
        """
        Giga employee with debt exceeding margin who discounts within margin
        and carries the rest to 29621 (Matrículas 320923 / 321226 pattern).
        """
        # Emulating Matrícula 320923: Total=3569.08, Margem=796.28, Desc=756.47, 29621=2772.80, 29624=39.81
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=3569.08,
            salario_11010=3981.38,
            evt_6280=756.47,
            evt_29621=2772.80,
            evt_29624=39.81,
            is_estab400_or_giga=True,
        )
        self.assertEqual(
            obs,
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621",
        )
        self.assertIsNone(trat)


class TestAdversarialResidualBalances29621(unittest.TestCase):
    """Scenario 4: Residual balances in evento 29621."""

    def test_afastado_with_full_residual_balance(self):
        """Afastado > 15 days with 100% debt transferred to 29621 and desc == 0."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=300.00,
            salario_11010=1800.00,
            evt_6280=0.0,
            evt_29621=300.00,
            evt_29624=0.0,
            has_afastamento=True,
            dias_afastamento=30,
            situacao="Afast Doenca > 15 Dias",
        )
        self.assertEqual(
            obs,
            "Afastamento - Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ",
        )
        self.assertIsNone(trat)

    def test_afastado_unprovisioned_debt(self):
        """Afastado > 15 days with debt but no discount and 29621 == 0 triggers alert."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=300.00,
            salario_11010=1800.00,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            has_afastamento=True,
            dias_afastamento=30,
            situacao="Afast Doenca > 15 Dias",
        )
        self.assertEqual(obs, "Afastado - Saldo devedor não provisionado")
        self.assertIsNotNone(trat)
        self.assertIn("não transferido para o evento 29621", trat)

    def test_afastado_improper_deduction_zero_salary(self):
        """Afastado > 15 days with deduction and zero salary triggers alert."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=200.00,
            salario_11010=0.0,
            evt_6280=200.00,
            evt_29621=0.0,
            evt_29624=0.0,
            has_afastamento=True,
            dias_afastamento=30,
            situacao="Afast Doenca > 15 Dias",
        )
        self.assertEqual(obs, "Afastado - Desconto indevido em folha")
        self.assertIsNotNone(trat)
        self.assertIn("desconto em folha", trat.lower())

    def test_active_employee_residual_outside_margin(self):
        """Active employee with debt > margin, discounting within margin and carrying remainder in 29621."""
        total = 800.00
        sal = 2500.00  # Margem = 500.00
        desc = 500.00
        saldo_rem = 300.00

        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=total,
            salario_11010=sal,
            evt_6280=desc,
            evt_29621=saldo_rem,
            evt_29624=0.0,
        )
        self.assertEqual(
            obs,
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621",
        )
        self.assertIsNone(trat)

    def test_broken_accounting_equation_divergence(self):
        """Accounting equation 6280 + 29621 + 29624 != Total Saldo Devedor by > 0.02 triggers alert."""
        total = 500.00
        sal = 2500.00
        desc = 300.00
        saldo_rem = 150.00  # Sum = 450.00, gap = 50.00 > 0.02

        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=total,
            salario_11010=sal,
            evt_6280=desc,
            evt_29621=saldo_rem,
            evt_29624=0.0,
        )
        self.assertEqual(obs, "Divergência no saldo devedor")
        self.assertIsNotNone(trat)
        self.assertIn("diverge do Total Saldo Devedor", trat)

    def test_active_employee_with_debt_and_no_discount_triggers_quality_alert(self):
        """Active employee without afastamento or ferias with positive debt and zero discount."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=250.00,
            salario_11010=2000.00,
            evt_6280=0.0,
            evt_29621=250.00,
            evt_29624=0.0,
            has_afastamento=False,
            is_ferias=False,
            situacao="Atividade Normal",
        )
        self.assertEqual(
            obs,
            "Verificar com quality - Não estava em situação de afastamento e não houve o desconto",
        )
        self.assertIsNotNone(trat)
        self.assertIn("não estando afastado nem em férias", trat)

    def test_demitido_with_residual_debt_triggers_alert(self):
        """Dismissed collaborator with residual debt triggers Demitido - Saldo devedor pendente."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=400.00,
            salario_11010=2000.00,
            evt_6280=200.00,
            evt_29621=200.00,
            evt_29624=0.0,
            is_demitido=True,
            tipo_rescisao="DISPENSA SEM JUSTA CAUSA",
        )
        self.assertEqual(obs, "Demitido - Saldo devedor pendente")
        self.assertIsNotNone(trat)
        self.assertIn("TRCT e solicitar cancelamento", trat)


if __name__ == "__main__":
    unittest.main()
