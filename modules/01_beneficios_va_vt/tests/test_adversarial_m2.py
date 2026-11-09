"""
Adversarial Verification Suite for Milestone 2 (business_rules.py)
Auditoria "04 - GOOD CARD" - Cencosud Brasil

Adversarially verifies:
1. build_conferencia_df: exact 38 columns matching COLUNAS_OFICIAIS_CONFERENCIA in strict sequence,
   handling missing columns, extra columns, permuted columns, and empty inputs.
2. reconcile_valores_importados: exact 11 columns matching COLUNAS_OFICIAIS_VALORES_IMPORTADOS in
   strict sequence, penny tolerance, zero division, missing inputs, and status mappings.
3. compute_kpis: exact mathematical accuracy, zero-division resilience, consistency equations,
   and metric preservation.
4. Boundary values: 15 vs 16 days, exact 20% margin +/- 0.01, Estab 400 5% subsidy,
   ineligible roles (Aprendiz/Diretor) with false positive resistance.
5. Large-scale adversarial workload stress test (10,000 rows).
"""

import sys
import unittest
from pathlib import Path
import numpy as np
import pandas as pd

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import config
import business_rules


class TestBuildConferenciaDfContract(unittest.TestCase):
    """Adversarial stress-testing of build_conferencia_df contract and columns."""

    def test_01_exact_38_columns_count_and_sequence(self):
        """Verifies output has exactly 38 columns matching COLUNAS_OFICIAIS_CONFERENCIA in exact order."""
        # Create a mock dataframe with all 38 columns
        df_in = pd.DataFrame(columns=config.COLUNAS_OFICIAIS_CONFERENCIA)
        df_out = business_rules.build_conferencia_df(df_in)

        self.assertEqual(len(df_out.columns), 38, f"Expected 38 columns, got {len(df_out.columns)}")
        self.assertListEqual(list(df_out.columns), config.COLUNAS_OFICIAIS_CONFERENCIA)

    def test_02_col_16_trailing_space_preservation(self):
        """Column 16 (index 15) must preserve trailing space: '...Mes Ant '."""
        df_in = pd.DataFrame({
            "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": [150.00]
        })
        df_out = business_rules.build_conferencia_df(df_in)
        col_16 = df_out.columns[15]
        self.assertEqual(col_16, "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ")
        self.assertTrue(col_16.endswith(" "))
        self.assertEqual(df_out[col_16].iloc[0], 150.00)

    def test_03_col_20_exact_quote_structure(self):
        """Column 20 (index 19) must match exact quote formatting."""
        expected_col_20 = 'Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"'
        df_in = pd.DataFrame({expected_col_20: ["Dentro da Margem"]})
        df_out = business_rules.build_conferencia_df(df_in)
        self.assertEqual(df_out.columns[19], expected_col_20)
        self.assertEqual(df_out[expected_col_20].iloc[0], "Dentro da Margem")

    def test_04_empty_input_produces_38_columns(self):
        """Completely empty DataFrame input produces empty DataFrame with strictly 38 columns."""
        df_empty = pd.DataFrame()
        df_out = business_rules.build_conferencia_df(df_empty)
        self.assertEqual(len(df_out), 0)
        self.assertEqual(len(df_out.columns), 38)
        self.assertListEqual(list(df_out.columns), config.COLUNAS_OFICIAIS_CONFERENCIA)

    def test_05_missing_columns_filled_with_correct_defaults(self):
        """When input lacks columns, monetary gets 0.0, boolean gets True, text gets ''."""
        df_minimal = pd.DataFrame({"Matrícula": ["12345"]})
        df_out = business_rules.build_conferencia_df(df_minimal)

        self.assertEqual(len(df_out), 1)
        self.assertListEqual(list(df_out.columns), config.COLUNAS_OFICIAIS_CONFERENCIA)
        self.assertEqual(df_out["Matrícula"].iloc[0], "12345")
        # Check monetary default
        self.assertEqual(df_out["11010 - Salario Mensal"].iloc[0], 0.0)
        self.assertEqual(df_out["Total Saldo Devedor"].iloc[0], 0.0)
        # Check boolean default
        self.assertTrue(df_out["Saldo Devedor x Evt 6280"].iloc[0])
        # Check text default
        self.assertEqual(df_out["Nome do Colaborador"].iloc[0], "")

    def test_06_extra_unwanted_columns_are_stripped(self):
        """Any extra columns outside the official 38 must be purged."""
        df_extra = pd.DataFrame({
            "Matrícula": ["999"],
            "COLUNA_INTRUSA_1": ["x"],
            "COLUNA_INTRUSA_2": [123],
            "11010 - Salario Mensal": [3000.0],
        })
        df_out = business_rules.build_conferencia_df(df_extra)
        self.assertEqual(len(df_out.columns), 38)
        self.assertNotIn("COLUNA_INTRUSA_1", df_out.columns)
        self.assertNotIn("COLUNA_INTRUSA_2", df_out.columns)

    def test_07_shuffled_input_columns_reordered_strictly(self):
        """Input with completely permuted columns must be re-ordered strictly to A..AL."""
        shuffled_cols = list(reversed(config.COLUNAS_OFICIAIS_CONFERENCIA))
        df_shuffled = pd.DataFrame(columns=shuffled_cols)
        df_out = business_rules.build_conferencia_df(df_shuffled)
        self.assertListEqual(list(df_out.columns), config.COLUNAS_OFICIAIS_CONFERENCIA)

    def test_08_monetary_columns_round_to_two_decimals(self):
        """Monetary values must be rounded to float 2 decimal places."""
        df_in = pd.DataFrame({
            "Total Saldo Devedor": [123.4567],
            "Margem 20% do Salário": [200.009],
        })
        df_out = business_rules.build_conferencia_df(df_in)
        self.assertEqual(df_out["Total Saldo Devedor"].iloc[0], 123.46)
        self.assertEqual(df_out["Margem 20% do Salário"].iloc[0], 200.01)

    def test_09_tratativa_empty_converted_to_none(self):
        """Tratativa must be null/None/NaN when there is no divergence."""
        df_in = pd.DataFrame({
            "Tratativa": ["", np.nan, "Divergência detectada"]
        })
        df_out = business_rules.build_conferencia_df(df_in)
        self.assertTrue(pd.isna(df_out["Tratativa"].iloc[0]) or df_out["Tratativa"].iloc[0] is None)
        self.assertTrue(pd.isna(df_out["Tratativa"].iloc[1]) or df_out["Tratativa"].iloc[1] is None)
        self.assertEqual(df_out["Tratativa"].iloc[2], "Divergência detectada")


class TestReconcileValoresImportadosContract(unittest.TestCase):
    """Adversarial stress-testing of reconcile_valores_importados contract and columns."""

    def test_01_exact_11_columns_count_and_sequence(self):
        """Output must have strictly 11 columns matching COLUNAS_OFICIAIS_VALORES_IMPORTADOS."""
        df_vi = pd.DataFrame({
            "Emp": ["1"],
            "Matrícula": ["100"],
            "Evento": ["29613"],
            "Mês": ["09/2026"],
            "Valor": [150.0],
        })
        df_fopa = pd.DataFrame({
            "Matrícula": ["100"],
            "29613 - Ben Esp - Valor Informado para Desconto": [150.0],
        })
        df_out = business_rules.reconcile_valores_importados(df_vi, df_fopa)
        self.assertEqual(len(df_out.columns), 11)
        self.assertListEqual(list(df_out.columns), config.COLUNAS_OFICIAIS_VALORES_IMPORTADOS)

    def test_02_empty_vi_returns_11_columns(self):
        """Empty input df_vi produces empty dataframe with exactly 11 columns."""
        df_out = business_rules.reconcile_valores_importados(pd.DataFrame(), pd.DataFrame())
        self.assertEqual(len(df_out), 0)
        self.assertEqual(len(df_out.columns), 11)
        self.assertListEqual(list(df_out.columns), config.COLUNAS_OFICIAIS_VALORES_IMPORTADOS)

    def test_03_numeric_equality_importado_x_folha(self):
        """Importado x Folha is True when Valor == Folha and False otherwise."""
        df_vi = pd.DataFrame({
            "Matrícula": ["10", "20", "30"],
            "Valor": [100.00, 200.00, 300.00],
        })
        df_fopa = pd.DataFrame({
            "Matrícula": ["10", "20", "30"],
            "29613 - Ben Esp - Valor Informado para Desconto": [100.00, 200.01, 0.0],
        })
        df_out = business_rules.reconcile_valores_importados(df_vi, df_fopa)
        self.assertTrue(df_out.loc[df_out["Matrícula"] == "10", "Importado x Folha"].iloc[0])
        self.assertFalse(df_out.loc[df_out["Matrícula"] == "20", "Importado x Folha"].iloc[0])
        self.assertFalse(df_out.loc[df_out["Matrícula"] == "30", "Importado x Folha"].iloc[0])

    def test_04_penny_difference_divergence_flagging(self):
        """A single cent difference (e.g. 50.00 vs 50.01) must trigger divergence and tratativa."""
        df_vi = pd.DataFrame({
            "Matrícula": ["99"],
            "Valor": [50.00],
        })
        df_fopa = pd.DataFrame({
            "Matrícula": ["99"],
            "29613 - Ben Esp - Valor Informado para Desconto": [50.01],
        })
        df_out = business_rules.reconcile_valores_importados(df_vi, df_fopa)
        self.assertFalse(df_out["Importado x Folha"].iloc[0])
        self.assertIn("Divergência entre valor importado e valor lançado em folha", df_out["Tratativa"].iloc[0])

    def test_05_unmatched_matricula_missing_in_fopa(self):
        """Matrícula in VI not found in FOPA results in Folha=0.0 and specific tratativa."""
        df_vi = pd.DataFrame({
            "Matrícula": ["8888"],
            "Valor": [120.00],
        })
        df_fopa = pd.DataFrame(columns=["Matrícula", "29613 - Ben Esp - Valor Informado para Desconto"])
        df_out = business_rules.reconcile_valores_importados(df_vi, df_fopa)
        self.assertEqual(df_out["Folha"].iloc[0], 0.0)
        self.assertFalse(df_out["Importado x Folha"].iloc[0])
        self.assertIn("não localizado na folha", df_out["Tratativa"].iloc[0])

    def test_06_desligado_attachment_and_priority_tratativa(self):
        """Terminated collaborator with positive value triggers desligado tratativa."""
        df_vi = pd.DataFrame({
            "Matrícula": ["777"],
            "Valor": [85.00],
        })
        df_fopa = pd.DataFrame({
            "Matrícula": ["777"],
            "29613 - Ben Esp - Valor Informado para Desconto": [85.00],
        })
        df_demit = pd.DataFrame({
            "Matrícula": ["777"],
            "Data Rescisão": ["10/09/2026"],
            "Tipo de Rescisão": ["2 - Sem justa causa"],
        })
        df_out = business_rules.reconcile_valores_importados(df_vi, df_fopa, df_demit=df_demit)
        self.assertEqual(df_out["Desligado"].iloc[0], "10/09/2026")
        self.assertEqual(df_out["Tipo de Rescisão"].iloc[0], "2 - Sem justa causa")
        self.assertIn("Colaborador desligado com valor informado para desconto", df_out["Tratativa"].iloc[0])

    def test_07_matricula_string_sanitization(self):
        """Leading zeros and whitespace in Matrícula are properly handled."""
        df_vi = pd.DataFrame({
            "Matrícula": [" 00123 "],
            "Valor": [90.00],
        })
        df_fopa = pd.DataFrame({
            "Matrícula": ["00123"],
            "29613 - Ben Esp - Valor Informado para Desconto": [90.00],
        })
        df_out = business_rules.reconcile_valores_importados(df_vi, df_fopa)
        self.assertEqual(df_out["Matrícula"].iloc[0], "00123")
        self.assertTrue(df_out["Importado x Folha"].iloc[0])


class TestComputeKpisArithmetic(unittest.TestCase):
    """Adversarial stress-testing of compute_kpis arithmetic and consistency."""

    def setUp(self):
        # Build a robust mock conferencia df
        self.df_conf = pd.DataFrame({
            "Matrícula": ["1", "2", "3", "4", "5"],
            "Total Saldo Devedor": [100.0, 200.0, 300.0, 0.0, 50.0],
            "6280 - Desc Cartão Benefícios Especiais": [100.0, 150.0, 0.0, 0.0, 47.50],
            "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": [0.0, 50.0, 300.0, 0.0, 0.0],
            "29624 - Ben Esp - Custo Empresa Beneficios Especiais": [0.0, 0.0, 0.0, 0.0, 2.50],
            "29613 - Ben Esp - Valor Informado para Desconto": [100.0, 200.0, 300.0, 0.0, 50.0],
            'Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"': [
                "Dentro da Margem", "Fora da Margem", "Fora da Margem", "Dentro da Margem", "Dentro da Margem"
            ],
            "Obs. Conferência": [
                "Ok - Liquidou todo o saldo devedor ",
                "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621",
                "Verificar com quality - Não estava em situação de afastamento e não houve o desconto",
                "Ok - Sem saldo devedor e sem desconto",
                "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)",
            ],
            "Tratativa": [
                None,
                None,
                "Verificar com Quality: sem desconto",
                None,
                None,
            ],
            "Desligamento": ["", "", "", "10/09/2026", ""],
            "Início Afastamento": ["", "", "", "", ""],
        })
        self.df_vi = pd.DataFrame({
            "Valor": [100.0, 200.0, 300.0, 0.0, 50.0]
        })

    def test_01_all_required_keys_present(self):
        """All expected KPI keys must be present in output dictionary."""
        kpis = business_rules.compute_kpis(self.df_conf, self.df_vi)
        expected_keys = {
            "total_auditados", "com_saldo_devedor", "com_desconto_6280",
            "dentro_margem", "fora_margem", "saldo_liquidado",
            "saldo_residual_29621", "regra_5pct_estab400", "desligados",
            "afastados", "conformes", "divergencias", "taxa_conformidade",
            "tot_saldo_devedor", "tot_desc_6280", "tot_saldo_29621",
            "tot_custo_29624", "tot_folha_29613", "tot_vi",
            "div_financeira_vi", "distribuicao_status"
        }
        self.assertTrue(expected_keys.issubset(set(kpis.keys())))

    def test_02_total_auditados_equals_conformes_plus_divergencias(self):
        """total_auditados must strictly equal conformes + divergencias."""
        kpis = business_rules.compute_kpis(self.df_conf, self.df_vi)
        self.assertEqual(kpis["total_auditados"], kpis["conformes"] + kpis["divergencias"])
        self.assertEqual(kpis["total_auditados"], 5)
        self.assertEqual(kpis["divergencias"], 1)
        self.assertEqual(kpis["conformes"], 4)

    def test_03_taxa_conformidade_arithmetic(self):
        """taxa_conformidade must be exactly (conformes / total) * 100 rounded to 2 places."""
        kpis = business_rules.compute_kpis(self.df_conf, self.df_vi)
        expected_taxa = round((4 / 5) * 100, 2)  # 80.0
        self.assertEqual(kpis["taxa_conformidade"], expected_taxa)

    def test_04_monetary_sums_accuracy(self):
        """All monetary totals must match direct sums."""
        kpis = business_rules.compute_kpis(self.df_conf, self.df_vi)
        self.assertAlmostEqual(kpis["tot_saldo_devedor"], 650.0, places=2)
        self.assertAlmostEqual(kpis["tot_desc_6280"], 297.50, places=2)
        self.assertAlmostEqual(kpis["tot_saldo_29621"], 350.0, places=2)
        self.assertAlmostEqual(kpis["tot_custo_29624"], 2.50, places=2)
        self.assertAlmostEqual(kpis["tot_folha_29613"], 650.0, places=2)
        self.assertAlmostEqual(kpis["tot_vi"], 650.0, places=2)
        self.assertAlmostEqual(kpis["div_financeira_vi"], 0.0, places=2)

    def test_05_zero_division_protection(self):
        """Empty dataframes must return taxa_conformidade = 0.0 without ZeroDivisionError."""
        df_empty_conf = pd.DataFrame(columns=config.COLUNAS_OFICIAIS_CONFERENCIA)
        df_empty_vi = pd.DataFrame(columns=config.COLUNAS_OFICIAIS_VALORES_IMPORTADOS)
        kpis = business_rules.compute_kpis(df_empty_conf, df_empty_vi)
        self.assertEqual(kpis["total_auditados"], 0)
        self.assertEqual(kpis["taxa_conformidade"], 0.0)
        self.assertEqual(kpis["div_financeira_vi"], 0.0)

    def test_06_margin_counts_partition(self):
        """dentro_margem + fora_margem must equal total_auditados."""
        kpis = business_rules.compute_kpis(self.df_conf, self.df_vi)
        self.assertEqual(kpis["dentro_margem"] + kpis["fora_margem"], kpis["total_auditados"])


class TestBoundaryValuesAndEdgeCases(unittest.TestCase):
    """Stress tests for business rules thresholds and boundary conditions."""

    def test_01_leave_15_vs_16_days(self):
        """
        15 days of leave does not qualify for Rule 2 (Afastados > 15 dias).
        16 days of leave activates Rule 2.
        """
        # Exactly 15 days with unliquidated balance and zero deduction
        # Does NOT trigger Rule 2 (Afastados > 15 dias); breaks accounting equation
        obs15, trat15 = business_rules.classify_collaborator_row(
            total_saldo_devedor=100.0,
            salario_11010=2000.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            dias_afastamento=15,
            has_afastamento=True,
            situacao="Atividade Normal",
        )
        self.assertEqual(obs15, "Divergência no saldo devedor")
        self.assertIn("Auditar folha", trat15)

        # Active normal without leave with debt and no deduction triggers quality check
        obs_active, trat_active = business_rules.classify_collaborator_row(
            total_saldo_devedor=100.0,
            salario_11010=2000.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            dias_afastamento=0,
            has_afastamento=False,
            situacao="Atividade Normal",
        )
        self.assertIn("Verificar com quality", obs_active)

        # Exactly 16 days with balance transferred to 29621
        obs16, trat16 = business_rules.classify_collaborator_row(
            total_saldo_devedor=100.0,
            salario_11010=0.0,
            evt_6280=0.0,
            evt_29621=100.0,
            evt_29624=0.0,
            dias_afastamento=16,
            has_afastamento=True,
            situacao="Afastado Previdência",
        )
        self.assertIn("Afastamento - Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621", obs16)
        self.assertIsNone(trat16)

    def test_02_margin_exact_thresholds(self):
        """Margin 20% boundary: exact 20%, 20% - 0.01, 20% + 0.01."""
        salario = 1500.00
        margem = round(salario * 0.20, 2)  # 300.00

        # Exact margin
        self.assertEqual(business_rules.check_status_margem(300.00, margem), "Dentro da Margem")
        # 1 cent below
        self.assertEqual(business_rules.check_status_margem(299.99, margem), "Dentro da Margem")
        # 1 cent above
        self.assertEqual(business_rules.check_status_margem(300.01, margem), "Fora da Margem")

    def test_03_ineligible_roles_positive_and_negative_matching(self):
        """
        is_aprendiz and is_diretor must correctly identify ineligibility
        without triggering false positives on composite job titles.
        """
        # True positives
        self.assertTrue(business_rules.is_aprendiz("JOVEM APRENDIZ"))
        self.assertTrue(business_rules.is_aprendiz("APRENDIZ OPERACIONAL"))
        self.assertTrue(business_rules.is_diretor("DIRETOR EXECUTIVO"))
        self.assertTrue(business_rules.is_diretor("DIRETORA DE OPERACOES"))
        self.assertTrue(business_rules.is_diretor("DIRETORIA REGIONAL"))

        # False positive resistance
        self.assertFalse(business_rules.is_diretor("ANALISTA DE TRIBUTOS DIRETOS"))
        self.assertFalse(business_rules.is_diretor("ESPECIALISTA EM IMPOSTOS DIRETOS"))
        self.assertFalse(business_rules.is_diretor("ASSISTENTE ADMINISTRATIVO"))
        self.assertFalse(business_rules.is_aprendiz("OPERADOR DE CAIXA"))

    def test_04_ineligible_inactive_collaborator_no_warning(self):
        """Aprendiz or Diretor with 0 debt and 0 deduction does not generate warning."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=0.0,
            salario_11010=1000.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="JOVEM APRENDIZ",
        )
        self.assertEqual(obs, "Ok - Sem saldo devedor e sem desconto")
        self.assertIsNone(trat)

    def test_05_estab400_split_arithmetic(self):
        """custo_empresa_5pct preserves exact sum down to pennies."""
        for val in [100.00, 33.33, 99.99, 1234.56, 0.01, 0.05, 10.03]:
            emp, colab = business_rules.custo_empresa_5pct(val)
            self.assertAlmostEqual(emp + colab, val, places=2)
            self.assertAlmostEqual(emp, round(val * 0.05, 2), delta=0.02)


class TestLargeScaleAdversarialWorkload(unittest.TestCase):
    """Stress test with 10,000 synthetic rows across all combinations."""

    def test_execute_audit_with_10k_rows(self):
        """Runs 10,000 rows through execute_audit to verify stability, speed, and column integrity."""
        n = 10000
        np.random.seed(42)

        matriculas = [str(i).zfill(6) for i in range(1, n + 1)]
        salarios = np.random.choice([1500.0, 2000.0, 3500.0, 5000.0], size=n)
        saldos_ant = np.random.choice([0.0, 50.0, 150.0, 500.0], size=n)
        valores_inf = np.random.choice([0.0, 100.0, 250.0, 800.0], size=n)
        cargos = np.random.choice(["Operador", "Caixa", "Aprendiz", "Diretor Geral", "Gerente"], size=n)
        estab_cods = np.random.choice(["101", "205", "400", "501"], size=n)

        df_fopa = pd.DataFrame({
            "Código da Empresa": ["1"] * n,
            "Nome da Empresa": ["Cencosud"] * n,
            "Código Estabelecimento": estab_cods,
            "Estabelecimento": ["Loja " + str(c) for c in estab_cods],
            "Código Centro Custo": ["1000"] * n,
            "Centro Custo": ["Operações"] * n,
            "Matrícula": matriculas,
            "Nome do Colaborador": ["Colaborador " + m for m in matriculas],
            "ID Colaborador": range(1, n + 1),
            "Data Admissão": ["01/01/2020"] * n,
            "Código Cargo": ["10"] * n,
            "Cargo": cargos,
            "1 - Salario": salarios,
            "11010 - Salario Mensal": salarios,
            "10990 - Liquido a Receber": salarios * 0.8,
            "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": saldos_ant,
            "29613 - Ben Esp - Valor Informado para Desconto": valores_inf,
            "6280 - Desc Cartão Benefícios Especiais": [0.0] * n,
            "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": [0.0] * n,
            "29624 - Ben Esp - Custo Empresa Beneficios Especiais": [0.0] * n,
        })

        # Synthetic VI
        df_vi = pd.DataFrame({
            "Emp": ["1"] * 500,
            "Matrícula": matriculas[:500],
            "Evento": ["29613"] * 500,
            "Mês": ["09/2026"] * 500,
            "Valor": valores_inf[:500],
        })

        data_dict = {
            "fopa": df_fopa,
            "nomina": pd.DataFrame({"Matrícula": matriculas, "Situação": ["Atividade Normal"] * n}),
            "valores_importados": df_vi,
        }

        df_conf, df_vi_res, kpis = business_rules.execute_audit(data_dict)

        # 1. Row count preserved
        self.assertEqual(len(df_conf), n)
        # 2. Columns strictly 38
        self.assertEqual(len(df_conf.columns), 38)
        self.assertListEqual(list(df_conf.columns), config.COLUNAS_OFICIAIS_CONFERENCIA)
        # 3. VI columns strictly 11
        self.assertEqual(len(df_vi_res.columns), 11)
        self.assertListEqual(list(df_vi_res.columns), config.COLUNAS_OFICIAIS_VALORES_IMPORTADOS)
        # 4. KPIs math consistency
        self.assertEqual(kpis["total_auditados"], n)
        self.assertEqual(kpis["conformes"] + kpis["divergencias"], n)


if __name__ == "__main__":
    unittest.main(verbosity=2)
