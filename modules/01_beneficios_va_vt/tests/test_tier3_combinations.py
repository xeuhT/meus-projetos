"""
Tier 3: Multi-Rule Interaction and Combinatorial Edge Cases Test Suite for "04 - GOOD CARD"
Covers complex real-world interactions between multiple overlapping business rules:
- Terminated (Demitido) + Leave (Afastado)
- Vacation (Férias) + Margin Burst (> 20%)
- Ineligible Roles (Aprendiz/Diretor) + Balances/Leaves
- Estabelecimento 400 / Giga + Leave/Vacation
- Transferência sem Ônus + Balances
- Active Non-Leave/Non-Vacation + Zero Deduction (Quality Alert)
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
    oracle_classify_audit_case,
    oracle_custo_empresa_5pct,
    oracle_calculate_margem_20,
    oracle_calculate_total_saldo_devedor
)


class TestDemitidoCombinations(unittest.TestCase):
    """Interaction between Dismissal (Rule 1) and other rules."""

    def test_demitido_and_afastado_with_pending_balance(self):
        """Collaborator both terminated and on medical leave with pending balance.
        Dismissal rule takes precedence to ensure TRCT deduction and card cancellation."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=500.0,
            salario_11010=0.0,
            evt_6280=0.0,
            evt_29621=500.0,
            evt_29624=0.0,
            is_demitido=True,
            tipo_rescisao="2 - Sem justa causa",
            has_afastamento=True,
            dias_afastamento=30
        )
        self.assertEqual(obs, "Demitido - Saldo devedor pendente")
        self.assertIsNotNone(trat)
        self.assertIn("TRCT", trat)

    def test_demitido_and_afastado_with_liquidated_balance(self):
        """Terminated and on leave with balance fully liquidated in termination TRCT."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=300.0,
            salario_11010=0.0,
            evt_6280=300.0,
            evt_29621=0.0,
            evt_29624=0.0,
            is_demitido=True,
            tipo_rescisao="2 - Sem justa causa",
            has_afastamento=True,
            dias_afastamento=20
        )
        self.assertEqual(obs, "Ok - Liquidou todo o saldo devedor ")
        self.assertIsNone(trat)

    def test_demitido_transferencia_sem_onus_with_afastamento(self):
        """Dismissed with 'Transferência sem ônus' and on medical leave."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=200.0,
            salario_11010=0.0,
            evt_6280=0.0,
            evt_29621=200.0,
            evt_29624=0.0,
            is_demitido=True,
            tipo_rescisao="5 - Transferência sem ônus",
            has_afastamento=True,
            dias_afastamento=25
        )
        self.assertEqual(obs, "Transferência sem ônus")


class TestFeriasCombinations(unittest.TestCase):
    """Interaction between Vacation rule and Margin Burst rule."""

    def test_ferias_with_margin_burst_carried_to_29621(self):
        """Employee on vacation whose balance exceeds 20% salary margin.
        On vacation, 0 is deducted and 100% of balance is carried to Event 29621."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=800.0, # Exceeds 20% of 2000.0 (400.0)
            salario_11010=2000.0,
            evt_6280=0.0,
            evt_29621=800.0,
            evt_29624=0.0,
            is_ferias=True
        )
        self.assertEqual(obs, "Férias- Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ")
        self.assertIsNone(trat)

    def test_ferias_with_improper_deduction(self):
        """Employee on vacation who suffered a deduction: accounting check evaluates normally."""
        # When 6280 is deducted during vacation, if it's within margin, it evaluates accounting
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=400.0,
            salario_11010=2000.0,
            evt_6280=400.0,
            evt_29621=0.0,
            evt_29624=0.0,
            is_ferias=True
        )
        # Liquidated within margin
        self.assertEqual(obs, "Ok - Liquidou todo o saldo devedor ")


class TestInelegiveisCombinations(unittest.TestCase):
    """Interactions between Ineligible Roles (Rule 3) and financial events."""

    def test_aprendiz_with_saldo_devedor_and_deduction(self):
        """Aprendiz with both balance and deduction.
        Rule 3 strict ineligibility overrides standard Ok status."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=100.0,
            salario_11010=1000.0,
            evt_6280=100.0, # Even though liquidated!
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="APRENDIZ DE VENDAS"
        )
        self.assertEqual(obs, "Aprendiz - Verificar")
        self.assertIsNotNone(trat)

    def test_diretor_with_leave_and_balance(self):
        """Diretor on leave with balance. Ineligibility takes precedence."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=600.0,
            salario_11010=30000.0,
            evt_6280=0.0,
            evt_29621=600.0,
            evt_29624=0.0,
            cargo="DIRETOR DE OPERAÇÕES",
            has_afastamento=True,
            dias_afastamento=20
        )
        self.assertEqual(obs, "Diretor - Verificar")

    def test_aprendiz_in_estab400_with_5_percent_rule(self):
        """Aprendiz in Estab 400. Ineligibility strictly overrides the 5% rule."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=100.0,
            salario_11010=1200.0,
            evt_6280=95.0,
            evt_29621=0.0,
            evt_29624=5.0,
            cargo="JOVEM APRENDIZ",
            is_estab400_or_giga=True
        )
        self.assertEqual(obs, "Aprendiz - Verificar")


class TestEstab400Combinations(unittest.TestCase):
    """Interactions with Estabelecimento 400 and Giga stores."""

    def test_estab400_with_afastamento_saldo_carried(self):
        """Estabelecimento 400 collaborator on leave > 15 days.
        Full balance carried to Event 29621 without 5% split in current folha."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=300.0,
            salario_11010=0.0,
            evt_6280=0.0,
            evt_29621=300.0,
            evt_29624=0.0,
            is_estab400_or_giga=True,
            has_afastamento=True,
            dias_afastamento=25
        )
        self.assertEqual(obs, "Afastamento - Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ")

    def test_estab400_active_normal_with_exact_split(self):
        """Estab 400 active collaborator with exact 95% / 5% split."""
        emp, colab = oracle_custo_empresa_5pct(250.0)
        self.assertEqual(emp, 12.50)
        self.assertEqual(colab, 237.50)

        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=250.0,
            salario_11010=4000.0,
            evt_6280=colab,
            evt_29621=0.0,
            evt_29624=emp,
            is_estab400_or_giga=True
        )
        self.assertEqual(obs, "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)")
        self.assertIsNone(trat)


class TestQualityDivergenceCombinations(unittest.TestCase):
    """Scenarios where an active employee without reason was not deducted."""

    def test_active_no_leave_no_vacation_zero_deduction(self):
        """Active collaborator, not on leave, not on vacation, with debt and NO deduction.
        Must trigger 'Verificar com quality - Não estava em situação de afastamento e não houve o desconto'."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=350.0,
            salario_11010=2500.0,
            evt_6280=0.0, # Zero deduction!
            evt_29621=350.0,
            evt_29624=0.0,
            has_afastamento=False,
            is_ferias=False,
            is_demitido=False,
            situacao="Atividade Normal"
        )
        self.assertEqual(obs, "Verificar com quality - Não estava em situação de afastamento e não houve o desconto")
        self.assertIsNotNone(trat)
        self.assertIn("Quality", trat)

    def test_broken_accounting_equation_triggers_divergence(self):
        """Sum of deductions does not match Total Saldo Devedor."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=500.0,
            salario_11010=3000.0,
            evt_6280=300.0,
            evt_29621=100.0, # Sum is 400.0 vs 500.0 (100.0 missing!)
            evt_29624=0.0
        )
        self.assertEqual(obs, "Divergência no saldo devedor")
        self.assertIn("Auditar folha", trat)


if __name__ == "__main__":
    unittest.main(verbosity=2)
