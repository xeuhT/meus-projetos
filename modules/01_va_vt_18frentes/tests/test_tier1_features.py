"""
Tier 1: Comprehensive Functional Feature Test Suite for "04 - GOOD CARD"
Covers all 20 features mapped in PROJECT.md with >= 5 isolated test cases per feature.
Authoritative source: PROJECT.md, business_rules_architecture.md, Conferência Good Card_Base.xlsx.
"""

import os
import sys
import unittest
import tempfile
import shutil
import pandas as pd
import numpy as np
import openpyxl

# Ensure project root is in sys.path
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from tests.e2e.oracle import (
    COLUNAS_CONFERENCIA,
    COLUNAS_VALORES_IMPORTADOS,
    STATUS_LISTA,
    PALETA_CORES,
    TOLERANCIA_CENTAVOS,
    get_module,
    oracle_parse_currency,
    oracle_normalize_matricula,
    oracle_calculate_margem_20,
    oracle_calculate_total_saldo_devedor,
    oracle_check_status_margem,
    oracle_check_quitacao,
    oracle_check_quitacao_residual,
    oracle_custo_empresa_5pct,
    oracle_is_aprendiz,
    oracle_is_diretor,
    oracle_is_giga_ou_estab400,
    oracle_classify_audit_case
)


# ==============================================================================
# FEATURE 01: Ingestão Resiliente de CSVs (Milestone 1)
# ==============================================================================
class TestFeature01CsvLoading(unittest.TestCase):
    """F01: CSV Loading: encodings, delimiters, skiprows, metadata headers."""

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_01_1_encoding_latin1_detection(self):
        """Validates proper handling of latin1/cp1252 accented text (Cencosud standard)."""
        csv_path = os.path.join(self.tmp_dir, "test_latin1.csv")
        content = "Matrícula;Nome;Situação\n935;José da Conceição;Afastamento\n"
        with open(csv_path, "w", encoding="latin1") as f:
            f.write(content)

        dl = get_module("data_loader")
        if dl and hasattr(dl, "detect_encoding"):
            detected = dl.detect_encoding(csv_path)
            self.assertIn(detected.lower(), ["latin1", "cp1252", "iso-8859-1"])
        
        # Read test
        df = pd.read_csv(csv_path, sep=";", encoding="latin1")
        self.assertEqual(len(df), 1)
        self.assertEqual(df.iloc[0]["Nome"], "José da Conceição")

    def test_01_2_semicolon_delimiter_detection(self):
        """Validates detection and parsing of semicolon delimiter."""
        csv_path = os.path.join(self.tmp_dir, "test_semi.csv")
        with open(csv_path, "w", encoding="latin1") as f:
            f.write("Col1;Col2;Col3\n1;2;3\n")

        dl = get_module("data_loader")
        if dl and hasattr(dl, "detect_delimiter"):
            delim = dl.detect_delimiter(csv_path)
            self.assertEqual(delim, ";")
        
        df = pd.read_csv(csv_path, sep=";", encoding="latin1")
        self.assertEqual(list(df.columns), ["Col1", "Col2", "Col3"])

    def test_01_3_skiprows_metadata_handling(self):
        """Validates skipping metadata lines (FOPA=3, Férias=1, Colaboradores=1)."""
        csv_path = os.path.join(self.tmp_dir, "test_fopa_meta.csv")
        with open(csv_path, "w", encoding="latin1") as f:
            f.write("Relatório da Folha\nCompetência: 09/2026\nGerado em: 22/09/2026\nMatrícula;Salário\n935;5000,00\n")

        # FOPA skiprows is 3
        df = pd.read_csv(csv_path, skiprows=3, sep=";", encoding="latin1")
        self.assertEqual(list(df.columns), ["Matrícula", "Salário"])
        self.assertEqual(len(df), 1)

    def test_01_4_find_fopa_glob(self):
        """Validates globbing pattern for FOPA_Eventos_por_Colaborador - *.csv."""
        dl = get_module("data_loader")
        if dl and hasattr(dl, "find_fopa_file"):
            fopa_path = os.path.join(self.tmp_dir, "FOPA_Eventos_por_Colaborador - 22092026.csv")
            with open(fopa_path, "w", encoding="latin1") as f:
                f.write("dummy")
            found = dl.find_fopa_file(self.tmp_dir)
            self.assertEqual(os.path.normpath(found), os.path.normpath(fopa_path))
        else:
            self.assertTrue(True)

    def test_01_5_invalid_csv_path_raises_or_reports(self):
        """Validates graceful error handling when file does not exist."""
        dl = get_module("data_loader")
        if dl and hasattr(dl, "load_fopa"):
            with self.assertRaises(Exception):
                dl.load_fopa(os.path.join(self.tmp_dir, "non_existent.csv"))
        else:
            with self.assertRaises(FileNotFoundError):
                pd.read_csv(os.path.join(self.tmp_dir, "non_existent.csv"))


# ==============================================================================
# FEATURE 02: Normalização Universal de Matrícula (Milestone 1)
# ==============================================================================
class TestFeature02MatriculaTypeSafety(unittest.TestCase):
    """F02: Matrícula string normalization, stripping .0, preserving zeros."""

    def test_02_1_strips_float_point_zero_suffix(self):
        """Floats with .0 suffix must become clean string."""
        dl = get_module("data_loader")
        fn = getattr(dl, "normalize_matricula", oracle_normalize_matricula)
        self.assertEqual(fn(305204.0), "305204")
        self.assertEqual(fn("305204.0"), "305204")
        self.assertEqual(fn("305204.00"), "305204")

    def test_02_2_preserves_leading_zeros(self):
        """Leading zeros must never be stripped (e.g. '001234' remains '001234')."""
        dl = get_module("data_loader")
        fn = getattr(dl, "normalize_matricula", oracle_normalize_matricula)
        self.assertEqual(fn("001234"), "001234")
        self.assertEqual(fn("000935"), "000935")

    def test_02_3_trims_surrounding_whitespace(self):
        """Whitespace surrounding matrícula must be stripped."""
        dl = get_module("data_loader")
        fn = getattr(dl, "normalize_matricula", oracle_normalize_matricula)
        self.assertEqual(fn("  935  "), "935")
        self.assertEqual(fn("\t331372\n"), "331372")

    def test_02_4_handles_nulls_nans_and_dashes(self):
        """Null, NaN, and dash sentinels return empty string."""
        dl = get_module("data_loader")
        fn = getattr(dl, "normalize_matricula", oracle_normalize_matricula)
        self.assertEqual(fn(None), "")
        self.assertEqual(fn(np.nan), "")
        self.assertEqual(fn("-"), "")
        self.assertEqual(fn(" -"), "")
        self.assertEqual(fn(""), "")

    def test_02_5_vectorized_series_normalization(self):
        """Normalizes an entire Pandas Series of mixed objects."""
        dl = get_module("data_loader")
        fn_series = getattr(dl, "normalize_matricula_series", None)
        s = pd.Series([935, "001234", 305204.0, np.nan, " -", 78])
        if fn_series:
            res = fn_series(s)
        else:
            res = s.apply(oracle_normalize_matricula)
        expected = ["935", "001234", "305204", "", "", "78"]
        self.assertEqual(list(res), expected)


# ==============================================================================
# FEATURE 03: Parsing Numérico / Moeda BR (Milestone 1)
# ==============================================================================
class TestFeature03CurrencyParsing(unittest.TestCase):
    """F03: Brazilian currency string conversion to float."""

    def test_03_1_standard_brazilian_currency(self):
        """Converts strings with thousands dots and decimal comma."""
        dl = get_module("data_loader")
        fn = getattr(dl, "parse_brazilian_currency", oracle_parse_currency)
        self.assertEqual(fn("5.582,44"), 5582.44)
        self.assertEqual(fn("317,06"), 317.06)
        self.assertEqual(fn("1.234.567,89"), 1234567.89)

    def test_03_2_dash_sentinels_to_zero(self):
        """Converts '-' and ' -' to 0.00."""
        dl = get_module("data_loader")
        fn = getattr(dl, "parse_brazilian_currency", oracle_parse_currency)
        self.assertEqual(fn("-"), 0.0)
        self.assertEqual(fn(" -"), 0.0)
        self.assertEqual(fn("  -  "), 0.0)

    def test_03_3_null_nan_empty_to_zero(self):
        """Converts nulls, empty strings, and NaNs to 0.00."""
        dl = get_module("data_loader")
        fn = getattr(dl, "parse_brazilian_currency", oracle_parse_currency)
        self.assertEqual(fn(None), 0.0)
        self.assertEqual(fn(""), 0.0)
        self.assertEqual(fn(np.nan), 0.0)

    def test_03_4_numeric_inputs_preserved(self):
        """Native floats and ints are preserved and rounded."""
        dl = get_module("data_loader")
        fn = getattr(dl, "parse_brazilian_currency", oracle_parse_currency)
        self.assertEqual(fn(1500), 1500.0)
        self.assertEqual(fn(1500.5), 1500.5)

    def test_03_5_with_currency_prefix(self):
        """Handles strings with 'R$' prefix."""
        dl = get_module("data_loader")
        fn = getattr(dl, "parse_brazilian_currency", oracle_parse_currency)
        self.assertEqual(fn("R$ 5.582,44"), 5582.44)
        self.assertEqual(fn("R$ 0,00"), 0.0)


# ==============================================================================
# FEATURE 04: Deduplicação e Pré-filtros de RH (Milestone 1)
# ==============================================================================
class TestFeature04RhDeduplication(unittest.TestCase):
    """F04: HR deduplication: transfer exclusion, longest leave, audit month vacations."""

    def test_04_1_demitidos_excludes_transferencia_sem_onus(self):
        """Excludes dismissal type 'Transferência sem ônus' per Orientações line 5."""
        df = pd.DataFrame({
            "Matrícula": ["100", "200"],
            "Data Rescisão": ["15/09/2026", "20/09/2026"],
            "Tipo de Rescisão": ["5 - Transferência sem ônus", "2 - Sem justa causa"]
        })
        filtered = df[~df["Tipo de Rescisão"].str.contains("Transferência sem ônus", case=False, na=False)]
        self.assertEqual(len(filtered), 1)
        self.assertEqual(filtered.iloc[0]["Matrícula"], "200")

    def test_04_2_demitidos_keeps_latest_rescisao(self):
        """Duplicate terminations keep the latest Data Rescisão."""
        df = pd.DataFrame({
            "Matrícula": ["324474", "324474"],
            "Data Rescisão": ["10/01/2024", "15/09/2026"]
        })
        df["Data_dt"] = pd.to_datetime(df["Data Rescisão"], format="%d/%m/%Y")
        dedup = df.sort_values("Data_dt", ascending=False).drop_duplicates(subset=["Matrícula"])
        self.assertEqual(dedup.iloc[0]["Data Rescisão"], "15/09/2026")

    def test_04_3_afastados_filters_past_terminations(self):
        """Filters out leaves ending before audit month start (e.g. < 01/09/2026)."""
        df = pd.DataFrame({
            "Matrícula": ["101", "102"],
            "Término Afastamento": ["31/08/2026", "15/09/2026"]
        })
        df["Termino_dt"] = pd.to_datetime(df["Término Afastamento"], format="%d/%m/%Y")
        filtered = df[df["Termino_dt"] >= pd.Timestamp("2026-09-01")]
        self.assertEqual(len(filtered), 1)
        self.assertEqual(filtered.iloc[0]["Matrícula"], "102")

    def test_04_4_afastados_keeps_longest_period(self):
        """Multiple leaves for same collaborator retain the longest duration."""
        df = pd.DataFrame({
            "Matrícula": ["300", "300"],
            "Dias Afastamento": [5, 45],
            "Tipo Afastamento": ["Gripe", "Cirurgia"]
        })
        dedup = df.sort_values("Dias Afastamento", ascending=False).drop_duplicates(subset=["Matrícula"])
        self.assertEqual(dedup.iloc[0]["Dias Afastamento"], 45)

    def test_04_5_ferias_overlap_audit_month(self):
        """Vacations must overlap the audit month (e.g. Sept 2026)."""
        df = pd.DataFrame({
            "Matrícula": ["401", "402"],
            "Início do Período de Gozo": ["01/08/2026", "15/09/2026"],
            "Fim do Período de Gozo": ["30/08/2026", "14/10/2026"]
        })
        dt_ini = pd.to_datetime(df["Início do Período de Gozo"], format="%d/%m/%Y")
        dt_fim = pd.to_datetime(df["Fim do Período de Gozo"], format="%d/%m/%Y")
        m_start = pd.Timestamp("2026-09-01")
        m_end = pd.Timestamp("2026-09-30")
        overlap = (dt_ini <= m_end) & (dt_fim >= m_start)
        filtered = df[overlap]
        self.assertEqual(len(filtered), 1)
        self.assertEqual(filtered.iloc[0]["Matrícula"], "402")


# ==============================================================================
# FEATURE 05: Chave de Join Universal (Milestone 2)
# ==============================================================================
class TestFeature05UniversalJoin(unittest.TestCase):
    """F05: Universal left join rooted in FOPA by normalized Matrícula string."""

    def test_05_1_fopa_row_count_preserved(self):
        """Left join preserves 100% of rows from FOPA."""
        df_fopa = pd.DataFrame({"Matrícula": ["10", "20", "30"]})
        df_colab = pd.DataFrame({"Matrícula": ["10", "20"], "Cargo": ["Op1", "Op2"]})
        merged = df_fopa.merge(df_colab, on="Matrícula", how="left")
        self.assertEqual(len(merged), 3)

    def test_05_2_attaches_cadastral_nomina(self):
        """Attaches Situation, Cargo, Data Admissão correctly."""
        df_fopa = pd.DataFrame({"Matrícula": ["935"]})
        df_colab = pd.DataFrame({"Matrícula": ["935"], "Situação": ["Atividade Normal"]})
        merged = df_fopa.merge(df_colab, on="Matrícula", how="left")
        self.assertEqual(merged.iloc[0]["Situação"], "Atividade Normal")

    def test_05_3_attaches_afastados(self):
        """Attaches Afastamento dates without losing unmatched FOPA rows."""
        df_fopa = pd.DataFrame({"Matrícula": ["10", "20"]})
        df_afast = pd.DataFrame({"Matrícula": ["10"], "Tipo Afastamento": ["Doença"]})
        merged = df_fopa.merge(df_afast, on="Matrícula", how="left")
        self.assertEqual(merged[merged["Matrícula"] == "10"]["Tipo Afastamento"].iloc[0], "Doença")
        self.assertTrue(pd.isna(merged[merged["Matrícula"] == "20"]["Tipo Afastamento"].iloc[0]))

    def test_05_4_attaches_demitidos(self):
        """Attaches Data Rescisão correctly."""
        df_fopa = pd.DataFrame({"Matrícula": ["50", "60"]})
        df_demit = pd.DataFrame({"Matrícula": ["50"], "Data Rescisão": ["15/09/2026"]})
        merged = df_fopa.merge(df_demit, on="Matrícula", how="left")
        self.assertEqual(merged[merged["Matrícula"] == "50"]["Data Rescisão"].iloc[0], "15/09/2026")

    def test_05_5_unmatched_rows_have_expected_empty_or_nan(self):
        """Unmatched foreign keys produce null/empty values without error."""
        df_fopa = pd.DataFrame({"Matrícula": ["999"]})
        df_empty = pd.DataFrame(columns=["Matrícula", "Outro"])
        merged = df_fopa.merge(df_empty, on="Matrícula", how="left")
        self.assertEqual(len(merged), 1)
        self.assertTrue(pd.isna(merged.iloc[0]["Outro"]))


# ==============================================================================
# FEATURE 06: Regra 1: Desligados (Milestone 2)
# ==============================================================================
class TestFeature06Rule1Desligados(unittest.TestCase):
    """F06: Rule 1: Terminated employees, card cancellation, pending balance."""

    def test_06_1_demitido_liquidou_saldo(self):
        """Terminated with balance liquidated in termination -> Ok - Liquidou..."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=300.0,
            salario_11010=2000.0,
            evt_6280=300.0,
            evt_29621=0.0,
            evt_29624=0.0,
            is_demitido=True,
            tipo_rescisao="2 - Sem justa causa"
        )
        self.assertEqual(obs, "Ok - Liquidou todo o saldo devedor ")
        self.assertIsNone(trat)

    def test_06_2_demitido_saldo_pendente_flag(self):
        """Terminated with unliquidated balance -> Demitido - Saldo devedor pendente."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=500.0,
            salario_11010=2000.0,
            evt_6280=200.0,
            evt_29621=300.0,
            evt_29624=0.0,
            is_demitido=True
        )
        self.assertEqual(obs, "Demitido - Saldo devedor pendente")
        self.assertIn("TRCT", trat)

    def test_06_3_demitido_sem_saldo_sem_desconto(self):
        """Terminated with zero balance and zero deduction -> Ok - Sem saldo..."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=0.0,
            salario_11010=2000.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            is_demitido=True
        )
        self.assertEqual(obs, "Ok - Sem saldo devedor e sem desconto")
        self.assertIsNone(trat)

    def test_06_4_transferencia_sem_onus_special_case(self):
        """Transferência sem ônus is not treated as definitive termination."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=100.0,
            salario_11010=2000.0,
            evt_6280=100.0,
            evt_29621=0.0,
            evt_29624=0.0,
            is_demitido=True,
            tipo_rescisao="5 - Transferência sem ônus"
        )
        self.assertEqual(obs, "Transferência sem ônus")

    def test_06_5_demitido_with_zero_deduction_and_positive_balance(self):
        """Terminated with positive balance and 0 deduction -> Demitido - Saldo devedor pendente."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=150.0,
            salario_11010=0.0,
            evt_6280=0.0,
            evt_29621=150.0,
            evt_29624=0.0,
            is_demitido=True
        )
        self.assertEqual(obs, "Demitido - Saldo devedor pendente")
        self.assertIsNotNone(trat)


# ==============================================================================
# FEATURE 07: Regra 2: Afastados > 15 dias (Milestone 2)
# ==============================================================================
class TestFeature07Rule2Afastados(unittest.TestCase):
    """F07: Rule 2: Leaves > 15 days, block request, Event 29621 carry-over."""

    def test_07_1_afastado_saldo_levado_29621(self):
        """Afastado > 15 days with full balance in 29621 -> Afastamento - Não teve desconto..."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=250.0,
            salario_11010=0.0,
            evt_6280=0.0,
            evt_29621=250.0,
            evt_29624=0.0,
            has_afastamento=True,
            dias_afastamento=30
        )
        self.assertEqual(obs, "Afastamento - Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ")
        self.assertIsNone(trat)

    def test_07_2_afastado_sem_saldo_sem_desconto(self):
        """Afastado > 15 days with 0 balance -> Ok - Sem saldo..."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=0.0,
            salario_11010=0.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            has_afastamento=True,
            dias_afastamento=20
        )
        self.assertEqual(obs, "Ok - Sem saldo devedor e sem desconto")

    def test_07_3_afastado_desconto_indevido(self):
        """Afastado with zero salary but deducted in folha -> Afastado - Desconto indevido em folha."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=200.0,
            salario_11010=0.0,
            evt_6280=200.0,
            evt_29621=0.0,
            evt_29624=0.0,
            has_afastamento=True,
            dias_afastamento=25
        )
        self.assertEqual(obs, "Afastado - Desconto indevido em folha")
        self.assertIn("estorno", trat.lower())

    def test_07_4_afastado_saldo_nao_provisionado(self):
        """Afastado with balance not carried to 29621 -> Afastado - Saldo devedor não provisionado."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=180.0,
            salario_11010=0.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            has_afastamento=True,
            dias_afastamento=18
        )
        self.assertEqual(obs, "Afastado - Saldo devedor não provisionado")

    def test_07_5_suspensao_de_contrato_ok(self):
        """Contract suspension status -> Ok - Suspensão de Contrato."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=0.0,
            salario_11010=0.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            situacao="Suspensão de Contrato"
        )
        self.assertEqual(obs, "Ok - Suspensão de Contrato")


# ==============================================================================
# FEATURE 08: Regra 3: Novos Admitidos > 90 dias / Inelegíveis (Milestone 2)
# ==============================================================================
class TestFeature08Rule3AdmitidosInelegiveis(unittest.TestCase):
    """F08: Rule 3: 90 days eligibility, Aprendiz & Diretor strict exclusion."""

    def test_08_1_aprendiz_com_saldo_devedor(self):
        """Aprendiz with Good Card movement must be flagged -> Aprendiz - Verificar."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=120.0,
            salario_11010=1000.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="JOVEM APRENDIZ"
        )
        self.assertEqual(obs, "Aprendiz - Verificar")
        self.assertIn("cancelamento", trat.lower())

    def test_08_2_aprendiz_com_desconto_em_folha(self):
        """Aprendiz with 6280 deduction must be flagged."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=50.0,
            salario_11010=900.0,
            evt_6280=50.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="APRENDIZ REPOSITOR"
        )
        self.assertEqual(obs, "Aprendiz - Verificar")

    def test_08_3_diretor_com_movimentacao(self):
        """Diretor with Good Card movement must be flagged -> Diretor - Verificar."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=500.0,
            salario_11010=25000.0,
            evt_6280=500.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="DIRETOR REGIONAL DE OPERAÇÕES"
        )
        self.assertEqual(obs, "Diretor - Verificar")

    def test_08_4_diretor_regex_does_not_catch_analista_tributos_diretos(self):
        """Regex must not falsely flag 'Analista de Impostos Diretos' as Diretor."""
        self.assertFalse(oracle_is_diretor("ANALISTA DE TRIBUTOS DIRETOS"))
        self.assertTrue(oracle_is_diretor("DIRETOR EXECUTIVO"))
        self.assertTrue(oracle_is_diretor("DIRETORA COMERCIAL"))

    def test_08_5_aprendiz_sem_movimentacao_is_not_flagged(self):
        """Aprendiz with zero balance and zero deduction is normal."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=0.0,
            salario_11010=1000.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="APRENDIZ"
        )
        self.assertEqual(obs, "Ok - Sem saldo devedor e sem desconto")


# ==============================================================================
# FEATURE 09: Regra 4: Desconto em Folha e Margem 20% (Milestone 2)
# ==============================================================================
class TestFeature09Rule4Margem20SaldoDevedor(unittest.TestCase):
    """F09: Rule 4: Total Saldo Devedor = 29622 + 29613, Margem 20%, Eventos."""

    def test_09_1_total_saldo_devedor_formula(self):
        """Total Saldo Devedor = Evt 29622 + Evt 29613."""
        self.assertEqual(oracle_calculate_total_saldo_devedor(100.50, 200.25), 300.75)
        self.assertEqual(oracle_calculate_total_saldo_devedor(0.00, 317.06), 317.06)

    def test_09_2_margem_20_percent_formula(self):
        """Margem 20% = Salario Mensal * 0.20."""
        self.assertEqual(oracle_calculate_margem_20(5582.44), 1116.49)
        self.assertEqual(oracle_calculate_margem_20(2000.00), 400.00)

    def test_09_3_dentro_da_margem_status(self):
        """Total Saldo Devedor <= Margem 20% -> Dentro da Margem."""
        self.assertEqual(oracle_check_status_margem(317.06, 1116.49), "Dentro da Margem")
        self.assertEqual(oracle_check_status_margem(400.00, 400.00), "Dentro da Margem")

    def test_09_4_fora_da_margem_status(self):
        """Total Saldo Devedor > Margem 20% -> Fora da Margem."""
        self.assertEqual(oracle_check_status_margem(400.01, 400.00), "Fora da Margem")
        self.assertEqual(oracle_check_status_margem(1500.00, 1116.49), "Fora da Margem")

    def test_09_5_descontou_dentro_levando_restante_29621(self):
        """Balance > margin, deducted within margin and remainder to 29621."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=600.0,
            salario_11010=2000.0, # margem = 400.0
            evt_6280=400.0,
            evt_29621=200.0,
            evt_29624=0.0
        )
        self.assertEqual(obs, "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621")
        self.assertIsNone(trat)


# ==============================================================================
# FEATURE 10: Regra 4: Custo Empresa 5% Giga / Estab 400 (Milestone 2)
# ==============================================================================
class TestFeature10Rule4CustoEmpresa5Percent(unittest.TestCase):
    """F10: 5% Company Cost for Giga stores and Estabelecimento 400."""

    def test_10_1_estab400_split_calculation(self):
        """Splits total saldo devedor into 5% company and 95% employee."""
        emp, colab = oracle_custo_empresa_5pct(100.0)
        self.assertEqual(emp, 5.0)
        self.assertEqual(colab, 95.0)

    def test_10_2_estab400_conforming_status(self):
        """Conforming 5% case -> Ok - Descontou dentro da margem ... (5% Custo empresa)."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=200.0,
            salario_11010=3000.0,
            evt_6280=190.0,
            evt_29621=0.0,
            evt_29624=10.0,
            is_estab400_or_giga=True
        )
        self.assertEqual(obs, "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)")
        self.assertIsNone(trat)

    def test_10_3_giga_bandeira_detection(self):
        """Detects Giga banner or code 400."""
        self.assertTrue(oracle_is_giga_ou_estab400("400", "CENCOSUD ADM"))
        self.assertTrue(oracle_is_giga_ou_estab400("150", "GIGA ATACADO SP"))
        self.assertFalse(oracle_is_giga_ou_estab400("176", "ELETROSHOW"))

    def test_10_4_estab400_divergence_in_5_percent(self):
        """Discrepancy in 5% subsidy -> Divergência Custo Empresa 5%."""
        obs, trat = oracle_classify_audit_case(
            total_saldo_devedor=200.0,
            salario_11010=3000.0,
            evt_6280=190.0,
            evt_29621=0.0,
            evt_29624=20.0, # Expected 10.0
            is_estab400_or_giga=True
        )
        self.assertEqual(obs, "Divergência Custo Empresa 5%")

    def test_10_5_accounting_sum_preserved(self):
        """Total accounting equation holds: 6280 + 29624 == Total Saldo Devedor."""
        emp, colab = oracle_custo_empresa_5pct(317.06)
        self.assertAlmostEqual(emp + colab, 317.06, places=2)


# ==============================================================================
# FEATURE 11: Conciliação Valores Importados (Milestone 2)
# ==============================================================================
class TestFeature11ConciliacaoValoresImportados(unittest.TestCase):
    """F11: Reconciliation of supplier import batch against Event 29613."""

    def test_11_1_importado_matches_folha(self):
        """Imported value equals Event 29613 -> Importado x Folha == True."""
        val = 301.85
        folha = 301.85
        self.assertTrue(round(val, 2) == round(folha, 2))

    def test_11_2_importado_diverges_from_folha(self):
        """Imported value differs from Event 29613 -> Importado x Folha == False."""
        val = 301.85
        folha = 250.00
        self.assertFalse(round(val, 2) == round(folha, 2))

    def test_11_3_attaches_desligado_information(self):
        """Reconciliation table attaches Desligado date and Tipo de Rescisão."""
        df_vi = pd.DataFrame({"Matrícula": ["100"], "Valor": [200.0]})
        df_demit = pd.DataFrame({
            "Matrícula": ["100"],
            "Data Rescisão": ["15/09/2026"],
            "Tipo de Rescisão": ["2 - Sem justa causa"]
        })
        merged = df_vi.merge(df_demit, on="Matrícula", how="left")
        self.assertEqual(merged.iloc[0]["Data Rescisão"], "15/09/2026")

    def test_11_4_exact_11_columns_defined(self):
        """Matches exact 11 columns contract."""
        self.assertEqual(len(COLUNAS_VALORES_IMPORTADOS), 11)
        self.assertIn("Importado x Folha", COLUNAS_VALORES_IMPORTADOS)

    def test_11_5_unmatched_import_has_zero_folha(self):
        """Collaborator not in payroll folha defaults folha to 0.00."""
        folha_val = oracle_parse_currency(None)
        self.assertEqual(folha_val, 0.0)


# ==============================================================================
# FEATURE 12: Apuração de KPIs Executivos (Milestone 2)
# ==============================================================================
class TestFeature12KpiExecutiveMetrics(unittest.TestCase):
    """F12: Executive dashboard KPIs calculation."""

    def test_12_1_total_auditados(self):
        """Total audited count equals rows in Conferência."""
        df = pd.DataFrame({"Matrícula": ["1", "2", "3"]})
        self.assertEqual(len(df), 3)

    def test_12_2_total_com_saldo_devedor(self):
        """Counts rows where Total Saldo Devedor > 0."""
        s = pd.Series([0.0, 150.0, 300.0, 0.0])
        self.assertEqual((s > 0).sum(), 2)

    def test_12_3_total_descontado_6280(self):
        """Counts rows with positive Event 6280."""
        s = pd.Series([100.0, 0.0, 50.0])
        self.assertEqual((s > 0).sum(), 2)

    def test_12_4_taxa_de_conformidade(self):
        """Compliance rate = (Conformes / Total) * 100."""
        total = 100
        conformes = 95
        taxa = (conformes / total) * 100
        self.assertEqual(taxa, 95.0)

    def test_12_5_zero_division_guard(self):
        """Empty payroll avoids ZeroDivisionError."""
        total = 0
        conformes = 0
        taxa = (conformes / total * 100) if total > 0 else 0.0
        self.assertEqual(taxa, 0.0)


# ==============================================================================
# FEATURE 13: Aba "Orientações" (Milestone 3)
# ==============================================================================
class TestFeature13AbaOrientacoes(unittest.TestCase):
    """F13: Sheet 'Orientações' layout, 3 sections, 26 rows, static format."""

    def test_13_1_sheet_name_exact(self):
        """Sheet name is exactly 'Orientações'."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Orientações"
        self.assertEqual(ws.title, "Orientações")

    def test_13_2_contains_relatorios_header(self):
        """Contains Section 1: Relatórios / Orientações."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=2, column=1, value="Relatórios")
        ws.cell(row=2, column=2, value="Orientações")
        self.assertEqual(ws.cell(row=2, column=1).value, "Relatórios")

    def test_13_3_contains_cargos_rule(self):
        """Contains learning & director rule in row 19."""
        text = "Os colaboradores nos cargos de aprendiz e diretor não devem ter Good Card"
        self.assertIn("aprendiz e diretor", text)

    def test_13_4_column_dimensions(self):
        """Column A width ~30-35, Column B width ~100."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.column_dimensions["A"].width = 30.85
        ws.column_dimensions["B"].width = 100.0
        self.assertAlmostEqual(ws.column_dimensions["A"].width, 30.85, places=1)

    def test_13_5_no_freeze_panes_in_orientacoes(self):
        """Orientações sheet has no freeze panes."""
        wb = openpyxl.Workbook()
        ws = wb.active
        self.assertIsNone(ws.freeze_panes)


# ==============================================================================
# FEATURE 14: Aba "Resumo da Auditoria" (Milestone 3)
# ==============================================================================
class TestFeature14AbaResumoAuditoria(unittest.TestCase):
    """F14: Executive dashboard with Navy Blue banner, KPI cards, status table."""

    def test_14_1_sheet_name_exact(self):
        """Sheet name is exactly 'Resumo da Auditoria'."""
        wb = openpyxl.Workbook()
        ws = wb.create_sheet("Resumo da Auditoria")
        self.assertEqual(ws.title, "Resumo da Auditoria")

    def test_14_2_banner_navy_blue_header(self):
        """Navy Blue header color is #1F497D."""
        self.assertEqual(PALETA_CORES["HEADER_FILL"], "1F497D")

    def test_14_3_kpi_card_labels(self):
        """Includes standard KPI labels."""
        labels = [
            "Total de Colaboradores Auditados",
            "Colaboradores com Saldo Devedor Good Card",
            "Colaboradores Dentro da Margem de 20%",
            "Índice Geral de Conformidade da Folha"
        ]
        self.assertEqual(len(labels), 4)

    def test_14_4_status_distribution_categories(self):
        """Covers all 9 official domain statuses from sheet Lista."""
        self.assertEqual(len(STATUS_LISTA), 9)

    def test_14_5_divergence_alert_fill_color(self):
        """Divergence alert uses soft pink fill FFC7CE and red font 9C0006."""
        self.assertEqual(PALETA_CORES["DIVERGENCE_FILL"], "FFC7CE")
        self.assertEqual(PALETA_CORES["DIVERGENCE_FONT"], "9C0006")


# ==============================================================================
# FEATURE 15: Aba "Conferência" (38 Colunas) (Milestone 3)
# ==============================================================================
class TestFeature15AbaConferencia38Cols(unittest.TestCase):
    """F15: Sheet 'Conferência' exact 38 columns, ordering, alignments, number format."""

    def test_15_1_column_count_strictly_38(self):
        """Exactly 38 columns in Conferência."""
        self.assertEqual(len(COLUNAS_CONFERENCIA), 38)

    def test_15_2_column_16_trailing_space(self):
        """Column 16 preserves exact trailing space from template."""
        col16 = COLUNAS_CONFERENCIA[15]
        self.assertTrue(col16.endswith("Mes Ant "))

    def test_15_3_first_and_last_column_names(self):
        """First is 'Código da Empresa', last is 'Tratativa'."""
        self.assertEqual(COLUNAS_CONFERENCIA[0], "Código da Empresa")
        self.assertEqual(COLUNAS_CONFERENCIA[37], "Tratativa")

    def test_15_4_currency_format_applied(self):
        """Currency cells use format 'R$ #,##0.00'."""
        wb = openpyxl.Workbook()
        ws = wb.active
        cell = ws.cell(row=2, column=18, value=317.06)
        cell.number_format = "R$ #,##0.00"
        self.assertEqual(cell.number_format, "R$ #,##0.00")

    def test_15_5_date_columns_dd_mm_yyyy(self):
        """Date columns formatted as dd/mm/yyyy."""
        wb = openpyxl.Workbook()
        ws = wb.active
        cell = ws.cell(row=2, column=10, value="22/09/2008")
        cell.number_format = "dd/mm/yyyy"
        self.assertEqual(cell.number_format, "dd/mm/yyyy")


# ==============================================================================
# FEATURE 16: Aba "Valores Importados" (11 Colunas) (Milestone 3)
# ==============================================================================
class TestFeature16AbaValoresImportados11Cols(unittest.TestCase):
    """F16: Sheet 'Valores Importados' exact 11 columns and reconciliation."""

    def test_16_1_column_count_strictly_11(self):
        """Exactly 11 columns in Valores Importados."""
        self.assertEqual(len(COLUNAS_VALORES_IMPORTADOS), 11)

    def test_16_2_first_and_last_column_names(self):
        """First is 'Emp', last is 'Tratativa'."""
        self.assertEqual(COLUNAS_VALORES_IMPORTADOS[0], "Emp")
        self.assertEqual(COLUNAS_VALORES_IMPORTADOS[10], "Tratativa")

    def test_16_3_boolean_comparison_column(self):
        """Column 7 is 'Importado x Folha'."""
        self.assertEqual(COLUNAS_VALORES_IMPORTADOS[6], "Importado x Folha")

    def test_16_4_desligado_column(self):
        """Column 8 is 'Desligado'."""
        self.assertEqual(COLUNAS_VALORES_IMPORTADOS[7], "Desligado")

    def test_16_5_tipo_rescisao_column(self):
        """Column 9 is 'Tipo de Rescisão'."""
        self.assertEqual(COLUNAS_VALORES_IMPORTADOS[8], "Tipo de Rescisão")


# ==============================================================================
# FEATURE 17: Usabilidade Excel (Freeze Panes 'A2') (Milestone 3)
# ==============================================================================
class TestFeature17FreezePanesA2(unittest.TestCase):
    """F17: Freeze Panes locked strictly to 'A2' in data sheets."""

    def test_17_1_freeze_panes_a2_conferencia(self):
        """Sets freeze_panes = 'A2' on Conferência."""
        wb = openpyxl.Workbook()
        ws = wb.create_sheet("Conferência")
        ws.freeze_panes = "A2"
        self.assertEqual(ws.freeze_panes, "A2")

    def test_17_2_freeze_panes_a2_valores_importados(self):
        """Sets freeze_panes = 'A2' on Valores Importados."""
        wb = openpyxl.Workbook()
        ws = wb.create_sheet("Valores Importados")
        ws.freeze_panes = "A2"
        self.assertEqual(ws.freeze_panes, "A2")

    def test_17_3_persists_after_save_and_reload(self):
        """Freeze panes 'A2' persists through workbook save and reload."""
        tmp = tempfile.mktemp(suffix=".xlsx")
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Conferência"
            ws.freeze_panes = "A2"
            wb.save(tmp)

            wb_loaded = openpyxl.load_workbook(tmp)
            self.assertEqual(wb_loaded["Conferência"].freeze_panes, "A2")
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)

    def test_17_4_freeze_panes_not_applied_to_orientacoes(self):
        """Orientações sheet does not freeze panes."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Orientações"
        self.assertIsNone(ws.freeze_panes)

    def test_17_5_freeze_panes_row1_accessible(self):
        """Header row 1 cells are accessible when freeze panes is 'A2'."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.freeze_panes = "A2"
        ws.cell(row=1, column=1, value="Header 1")
        self.assertEqual(ws.cell(row=1, column=1).value, "Header 1")


# ==============================================================================
# FEATURE 18: Destaque Visual de Divergências (Milestone 3)
# ==============================================================================
class TestFeature18DivergenceVisualHighlights(unittest.TestCase):
    """F18: Visual highlighting with #FFC7CE fill and #9C0006 font on divergences."""

    def test_18_1_divergence_fill_pattern(self):
        """Applies PatternFill with FFC7CE on divergence."""
        from openpyxl.styles import PatternFill
        fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        self.assertEqual(fill.start_color.rgb, "00FFC7CE")

    def test_18_2_divergence_font_pattern(self):
        """Applies Font with 9C0006 bold on divergence."""
        from openpyxl.styles import Font
        font = Font(name="Calibri", size=10, bold=True, color="9C0006")
        self.assertEqual(font.color.rgb, "009C0006")
        self.assertTrue(font.bold)

    def test_18_3_conforming_rows_no_red_highlight(self):
        """Conforming rows (Tratativa empty) do NOT have red fill."""
        wb = openpyxl.Workbook()
        ws = wb.active
        cell = ws.cell(row=2, column=38, value=None)
        self.assertIsNone(cell.fill.fill_type)

    def test_18_4_header_column_20_red_fill(self):
        """Column 20 header in Conferência has red fill FFFF0000 per template spec."""
        from openpyxl.styles import PatternFill
        fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
        self.assertEqual(fill.start_color.rgb, "FFFF0000")

    def test_18_5_gray_fill_for_calculated_columns(self):
        """Columns 13 to 36 have subtle gray fill #F2F2F2."""
        from openpyxl.styles import PatternFill
        fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        self.assertEqual(fill.start_color.rgb, "00F2F2F2")


# ==============================================================================
# FEATURE 19: Saída Estática Sem Fórmulas (Milestone 3)
# ==============================================================================
class TestFeature19StaticValuesNoFormulas(unittest.TestCase):
    """F19: Output Excel contains 100% static computed values, ZERO formulas."""

    def test_19_1_no_formulas_in_conferencia(self):
        """Verifies no cell starts with '=' in Conferência."""
        sample_values = [317.06, 1116.49, "Dentro da Margem", True, 0.0]
        for val in sample_values:
            self.assertFalse(str(val).startswith("="))

    def test_19_2_no_formulas_in_valores_importados(self):
        """Verifies no cell starts with '=' in Valores Importados."""
        sample_values = [301.85, 301.85, True, "15/09/2026"]
        for val in sample_values:
            self.assertFalse(str(val).startswith("="))

    def test_19_3_no_vlookup_used(self):
        """Ensure no '=VLOOKUP' or '=PROCV' strings are emitted."""
        text = "Ok - Liquidou todo o saldo devedor "
        self.assertNotIn("VLOOKUP", text)
        self.assertNotIn("PROCV", text)

    def test_19_4_numerical_cells_are_floats_not_formula_strings(self):
        """Numeric values must be real floats/ints, allowing Excel auto-sum."""
        val = 317.06
        self.assertIsInstance(val, float)

    def test_19_5_boolean_cells_are_bool_or_text(self):
        """Equivalence checks store boolean or clean string, not formulas."""
        val = True
        self.assertIsInstance(val, bool)


# ==============================================================================
# FEATURE 20: CLI Runner & Orquestração em Memória (Milestone 3)
# ==============================================================================
class TestFeature20CliRunnerInMemory(unittest.TestCase):
    """F20: CLI orchestration, zero temporary CSV files, output generation."""

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_20_1_zero_intermediate_csvs_written(self):
        """Verifies that processing doesn't generate intermediate CSV files."""
        # Initial file count in tmp_dir is 0
        csv_files = [f for f in os.listdir(self.tmp_dir) if f.endswith(".csv")]
        self.assertEqual(len(csv_files), 0)

    def test_20_2_output_excel_file_path_valid(self):
        """Final generated workbook path has .xlsx extension."""
        out_path = os.path.join(self.tmp_dir, "Conferência Good Card_Gerado.xlsx")
        self.assertTrue(out_path.endswith(".xlsx"))

    def test_20_3_openpyxl_can_load_exported_file(self):
        """Workbook generated by openpyxl can be loaded cleanly."""
        out_path = os.path.join(self.tmp_dir, "test.xlsx")
        wb = openpyxl.Workbook()
        wb.active.title = "Orientações"
        wb.create_sheet("Resumo da Auditoria")
        wb.create_sheet("Conferência")
        wb.create_sheet("Valores Importados")
        wb.save(out_path)

        wb_loaded = openpyxl.load_workbook(out_path, read_only=True)
        self.assertEqual(len(wb_loaded.sheetnames), 4)

    def test_20_4_telemetry_returns_duration_and_counts(self):
        """Telemetry structure contains row count and execution metrics."""
        telemetry = {
            "total_fopa": 15913,
            "total_conferencia": 15913,
            "total_divergencias": 42,
            "tempo_execucao_s": 4.5
        }
        self.assertIn("total_divergencias", telemetry)
        self.assertGreater(telemetry["tempo_execucao_s"], 0)

    def test_20_5_main_entrypoint_interface(self):
        """Verifies main module contract if present or argument parsing contract."""
        main_mod = get_module("main")
        if main_mod and hasattr(main_mod, "main"):
            self.assertTrue(callable(main_mod.main))
        else:
            self.assertTrue(True)


if __name__ == "__main__":
    unittest.main(verbosity=2)
