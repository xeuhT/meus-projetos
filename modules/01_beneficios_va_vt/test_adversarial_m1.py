"""
Adversarial Verification Suite for Milestone 1 (config.py & data_loader.py)
Empirically tests boundary conditions, exact contract compliance, and edge cases.
"""

import os
import sys
import unittest
from pathlib import Path
import tempfile
import pandas as pd
import numpy as np
import openpyxl

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import config
import data_loader


class TestMilestone1Adversarial(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.excel_path = PROJECT_ROOT / "Conferência Good Card_Base.xlsx"
        cls.wb = openpyxl.load_workbook(cls.excel_path, read_only=True, data_only=True)

    @classmethod
    def tearDownClass(cls):
        cls.wb.close()

    # =========================================================================
    # MISSION 1: COLUNAS_OFICIAIS_CONFERENCIA vs Conferência Good Card_Base.xlsx
    # =========================================================================

    def test_01_conferencia_columns_exact_match(self):
        """Verify COLUNAS_OFICIAIS_CONFERENCIA exactly matches all 38 columns of sheet 'Conferência'."""
        ws = self.wb["Conferência"]
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        actual_headers = [c for c in header_row if c is not None]

        # 1. Total count must be exactly 38
        self.assertEqual(len(actual_headers), 38, f"Expected 38 columns in Excel, found {len(actual_headers)}")
        self.assertEqual(len(config.COLUNAS_OFICIAIS_CONFERENCIA), 38, f"Expected 38 columns in config, found {len(config.COLUNAS_OFICIAIS_CONFERENCIA)}")

        # 2. Exact match per column (index, name, whitespace, quotes)
        mismatches = []
        for idx, (expected_conf, actual_excel) in enumerate(zip(config.COLUNAS_OFICIAIS_CONFERENCIA, actual_headers)):
            if expected_conf != actual_excel:
                mismatches.append((idx + 1, expected_conf, actual_excel))

        self.assertEqual(mismatches, [], f"Column mismatches found (idx, config, excel): {mismatches}")

        # 3. Specifically verify tricky columns:
        # Col 16 (trailing space)
        self.assertTrue(config.COLUNAS_OFICIAIS_CONFERENCIA[15].endswith(" "), "Col 16 must have trailing space")
        self.assertEqual(config.COLUNAS_OFICIAIS_CONFERENCIA[15], actual_headers[15])

        # Col 20 (quotes)
        self.assertIn('" Dentro da Margem" ou "Fora da Margem"', config.COLUNAS_OFICIAIS_CONFERENCIA[19])
        self.assertEqual(config.COLUNAS_OFICIAIS_CONFERENCIA[19], actual_headers[19])

    # =========================================================================
    # MISSION 2: COLUNAS_OFICIAIS_VALORES_IMPORTADOS vs Excel & Requirements
    # =========================================================================

    def test_02_valores_importados_columns_exact_match(self):
        """Verify COLUNAS_OFICIAIS_VALORES_IMPORTADOS exactly matches all 11 columns."""
        ws = self.wb["Valores Importados"]
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        actual_headers = [c for c in header_row if c is not None]

        # 1. Total count must be exactly 11
        self.assertEqual(len(actual_headers), 11, f"Expected 11 columns in Excel, found {len(actual_headers)}")
        self.assertEqual(len(config.COLUNAS_OFICIAIS_VALORES_IMPORTADOS), 11, f"Expected 11 columns in config, found {len(config.COLUNAS_OFICIAIS_VALORES_IMPORTADOS)}")

        # 2. Exact match per column
        mismatches = []
        for idx, (expected_conf, actual_excel) in enumerate(zip(config.COLUNAS_OFICIAIS_VALORES_IMPORTADOS, actual_headers)):
            if expected_conf != actual_excel:
                mismatches.append((idx + 1, expected_conf, actual_excel))

        self.assertEqual(mismatches, [], f"Valores Importados mismatches found (idx, config, excel): {mismatches}")

    # =========================================================================
    # MISSION 3: ALL 5 CSV LOADERS PRE-FILTERS & DEDUPLICATIONS
    # =========================================================================

    def test_03_fopa_loader_adversarial(self):
        """Test load_fopa with real data and adversarial synthetic data."""
        # Real file test
        fopa_file = data_loader.find_fopa_file(config.DIR_RELATORIO_ORIGINAL)
        df_real = data_loader.load_fopa(fopa_file)

        self.assertEqual(len(df_real), 15913, f"FOPA row count mismatch: {len(df_real)}")
        self.assertTrue("Matrícula" in df_real.columns)
        self.assertFalse((df_real["Matrícula"] == "").any(), "Empty Matrícula found in loaded FOPA")
        self.assertFalse(df_real["Matrícula"].str.contains(r"\.0$", regex=True).any(), "Matrícula ends with .0")

        # Monetary columns must be float64
        for col in config.COLUNAS_MONETARIAS_FOPA:
            if col in df_real.columns:
                self.assertEqual(df_real[col].dtype, np.float64, f"Column {col} is not float64")

        # Col 29622 exact naming
        col_29622 = "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant "
        self.assertIn(col_29622, df_real.columns, "Col 29622 with trailing space missing in df_fopa")

        # Synthetic adversarial test: skiprows, empty lines, missing 29622, invalid currency
        with tempfile.NamedTemporaryFile("w", delete=False, suffix=".csv", encoding="latin1") as tmp:
            tmp.write("HEADER LINE 1\nHEADER LINE 2\nHEADER LINE 3\n")
            tmp.write("Matrícula;Nome;11010 - Salario Mensal;29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant\n")
            tmp.write("001234;João;1.500,50;250,00\n")
            tmp.write(";Sem Matrícula;500,00;0,00\n")
            tmp.write("5678.0;Maria;-; -\n")
            tmp_path = tmp.name

        try:
            df_synth = data_loader.load_fopa(tmp_path)
            self.assertEqual(len(df_synth), 2, "Failed to drop empty matrícula row in synthetic FOPA")
            self.assertEqual(df_synth.iloc[0]["Matrícula"], "001234", "Failed to preserve leading zero")
            self.assertEqual(df_synth.iloc[1]["Matrícula"], "5678", "Failed to strip .0")
            self.assertEqual(df_synth.iloc[0]["11010 - Salario Mensal"], 1500.50)
            self.assertEqual(df_synth.iloc[1]["11010 - Salario Mensal"], 0.0)
            self.assertEqual(df_synth.iloc[1][col_29622], 0.0)
        finally:
            os.remove(tmp_path)

    def test_04_colaboradores_loader_adversarial(self):
        """Test load_colaboradores (Nomina) with real and synthetic data."""
        df_real = data_loader.load_colaboradores(config.DIR_RELATORIO_ORIGINAL / config.ARQUIVO_CSV_COLABORADORES)
        self.assertEqual(len(df_real), 16038)
        self.assertIn("Matrícula", df_real.columns)
        self.assertFalse((df_real["Matrícula"] == "").any())

        # Check required columns per PROJECT.md interface contract
        for req_col in ["Matrícula", "Cargo", "Situação", "Data Admissão"]:
            self.assertIn(req_col, df_real.columns, f"Required column {req_col} missing in nomina")

    def test_05_demitidos_loader_prefilter_and_dedup(self):
        """Test load_demitidos pre-filtering of 'Transferência sem ônus' and latest termination dedup."""
        df_real = data_loader.load_demitidos(config.DIR_RELATORIO_ORIGINAL / config.ARQUIVO_CSV_DEMITIDOS)
        self.assertEqual(len(df_real), 12901)
        self.assertIn("Matrícula", df_real.columns)

        # Pre-filter check: Zero 'Transferência sem ônus'
        transf_count = df_real["Tipo de Rescisão"].fillna("").str.contains("Transferência sem ônus", case=False).sum()
        self.assertEqual(transf_count, 0, "Found 'Transferência sem ônus' in demitidos output")

        # Dedup check: Zero duplicate Matrículas
        dup_count = df_real["Matrícula"].duplicated().sum()
        self.assertEqual(dup_count, 0, "Found duplicate Matrículas in demitidos output")

        # Synthetic adversarial test: multiple terminations, transfer vs dismissal, tie breaking
        with tempfile.NamedTemporaryFile("w", delete=False, suffix=".csv", encoding="latin1") as tmp:
            tmp.write("Matrícula;Data Rescisão;Tipo de Rescisão\n")
            # Employee 1: older valid dismissal vs newer transfer -> transfer filtered out, keep older valid
            tmp.write("1001;10/01/2026;Demissão sem justa causa\n")
            tmp.write("1001;15/05/2026;Transferência sem ônus\n")
            # Employee 2: two dismissals -> keep newer date (20/08/2026 over 10/02/2025)
            tmp.write("2002;10/02/2025;Demissão com justa causa\n")
            tmp.write("2002;20/08/2026;Pedido de demissão\n")
            # Employee 3: only transfer -> should be completely filtered out
            tmp.write("3003;01/01/2026;Transferência sem ônus\n")
            # Employee 4: missing date vs valid date -> keep valid date
            tmp.write("4004;;Demissão sem justa causa\n")
            tmp.write("4004;05/04/2026;Demissão sem justa causa\n")
            tmp_path = tmp.name

        try:
            df_synth = data_loader.load_demitidos(tmp_path)
            self.assertEqual(len(df_synth), 3, "Wrong count after pre-filter and dedup")
            self.assertNotIn("3003", df_synth["Matrícula"].values, "Employee with only transfer should be excluded")
            
            row_1001 = df_synth[df_synth["Matrícula"] == "1001"].iloc[0]
            self.assertEqual(row_1001["Data Rescisão"], "10/01/2026")

            row_2002 = df_synth[df_synth["Matrícula"] == "2002"].iloc[0]
            self.assertEqual(row_2002["Data Rescisão"], "20/08/2026")

            row_4004 = df_synth[df_synth["Matrícula"] == "4004"].iloc[0]
            self.assertEqual(row_4004["Data Rescisão"], "05/04/2026")
        finally:
            os.remove(tmp_path)

    def test_06_afastados_loader_prefilter_and_dedup(self):
        """Test load_afastados pre-filtering of ended leaves and dedup of multiple leaves."""
        df_real = data_loader.load_afastados(config.DIR_RELATORIO_ORIGINAL / config.ARQUIVO_CSV_AFASTADOS, ref_date="2026-09-01")
        self.assertEqual(len(df_real), 1867)
        self.assertEqual(df_real["Matrícula"].duplicated().sum(), 0, "Duplicate Matrícula in afastados")

        # Check required columns per PROJECT.md interface contract
        for req_col in ["Matrícula", "Início Afastamento", "Término Afastamento", "Tipo Afastamento"]:
            self.assertIn(req_col, df_real.columns, f"Required column {req_col} missing in afastados")

        # Synthetic test: past leave (ended before 2026-09-01) vs ongoing leave vs open leave (NaT)
        with tempfile.NamedTemporaryFile("w", delete=False, suffix=".csv", encoding="latin1") as tmp:
            tmp.write("Matrícula;Início Afastamento;Término Afastamento;Dias Afastamento;Tipo Afastamento\n")
            # Expired leave: ended 31/08/2026 -> must be excluded
            tmp.write("111;01/08/2026;31/08/2026;31;Doença\n")
            # Active leave ending in September: ended 15/09/2026 -> must be kept
            tmp.write("222;01/09/2026;15/09/2026;15;Acidente\n")
            # Open leave: Término empty -> must be kept
            tmp.write("333;10/08/2026;;;Licença\n")
            # Employee with 2 leaves: one ended in August, one open -> keep open one
            tmp.write("444;01/01/2026;10/08/2026;222;Doença\n")
            tmp.write("444;15/08/2026;;;Doença\n")
            # Employee with 2 active leaves: one 15 days, one 60 days -> keep longer one
            tmp.write("555;01/09/2026;15/09/2026;15;Doença\n")
            tmp.write("555;01/08/2026;30/09/2026;60;Doença\n")
            tmp_path = tmp.name

        try:
            df_synth = data_loader.load_afastados(tmp_path, ref_date="2026-09-01")
            self.assertNotIn("111", df_synth["Matrícula"].values, "Past leave ended before ref_date should be excluded")
            self.assertIn("222", df_synth["Matrícula"].values, "Active leave in September must be included")
            self.assertIn("333", df_synth["Matrícula"].values, "Open leave must be included")
            self.assertIn("444", df_synth["Matrícula"].values)
            self.assertTrue(pd.isna(df_synth[df_synth["Matrícula"] == "444"]["Término Afastamento"].iloc[0]))
            
            row_555 = df_synth[df_synth["Matrícula"] == "555"].iloc[0]
            self.assertEqual(row_555["Término Afastamento"], "30/09/2026", "Should keep longer leave")
        finally:
            os.remove(tmp_path)

    def test_07_ferias_loader_prefilter_and_dedup(self):
        """Test load_ferias month intersection pre-filtering and dedup."""
        df_real = data_loader.load_ferias(config.DIR_RELATORIO_ORIGINAL / config.ARQUIVO_CSV_FERIAS, year=2026, month=9)
        self.assertEqual(len(df_real), 2005)
        self.assertEqual(df_real["Matrícula"].duplicated().sum(), 0, "Duplicate Matrícula in ferias")

        # Required columns per PROJECT.md interface contract
        # PROJECT.md line 83: Matrícula, Início Férias, Final Férias, Qtd de dias
        # Notice in raw CSV, headers are 'Início do Período de Gozo', 'Fim do Período de Gozo', 'Dias Gozo'
        self.assertIn("Matrícula", df_real.columns)

        # Synthetic test: strictly August vacation vs overlapping vacation vs October vacation
        with tempfile.NamedTemporaryFile("w", delete=False, suffix=".csv", encoding="latin1") as tmp:
            tmp.write("TITLE LINE TO SKIP\n")
            tmp.write("Matrícula;Início do Período de Gozo;Fim do Período de Gozo;Dias Gozo;Status\n")
            # August vacation: 01/08/2026 to 20/08/2026 -> excluded
            tmp.write("1;01/08/2026;20/08/2026;20;Finalizado\n")
            # Overlapping start: 25/08/2026 to 05/09/2026 -> included
            tmp.write("2;25/08/2026;05/09/2026;12;Em férias\n")
            # Entirely in September: 10/09/2026 to 20/09/2026 -> included
            tmp.write("3;10/09/2026;20/09/2026;11;Em férias\n")
            # Overlapping end: 28/09/2026 to 10/10/2026 -> included
            tmp.write("4;28/09/2026;10/10/2026;13;Em férias\n")
            # October vacation: 05/10/2026 to 20/10/2026 -> excluded
            tmp.write("5;05/10/2026;20/10/2026;16;Programado\n")
            tmp_path = tmp.name

        try:
            df_synth = data_loader.load_ferias(tmp_path, year=2026, month=9)
            self.assertNotIn("1", df_synth["Matrícula"].values, "August vacation should be excluded")
            self.assertIn("2", df_synth["Matrícula"].values, "Vacation overlapping Sept start should be included")
            self.assertIn("3", df_synth["Matrícula"].values, "September vacation should be included")
            self.assertIn("4", df_synth["Matrícula"].values, "Vacation overlapping Sept end should be included")
            self.assertNotIn("5", df_synth["Matrícula"].values, "October vacation should be excluded")
        finally:
            os.remove(tmp_path)

    def test_08_valores_importados_loader(self):
        """Test load_valores_importados from Excel template and fallback CSV."""
        df_real = data_loader.load_valores_importados(config.TEMPLATE_BASE_EXCEL)
        self.assertEqual(len(df_real), 1334)
        self.assertAlmostEqual(df_real["Valor"].sum(), 482265.63, places=2)
        
        # Verify required interface contract columns (PROJECT.md line 84):
        # 'valores_importados': DataFrame com Emp, Matrícula, Evento, Mês, Valor.
        for col in ["Emp", "Matrícula", "Evento", "Mês", "Valor"]:
            self.assertIn(col, df_real.columns, f"Required column {col} missing in valores_importados")

        self.assertTrue(pd.api.types.is_integer_dtype(df_real["Emp"]))
        self.assertTrue(pd.api.types.is_string_dtype(df_real["Matrícula"]) or df_real["Matrícula"].dtype == object)
        self.assertTrue(pd.api.types.is_integer_dtype(df_real["Evento"]))
        self.assertTrue(pd.api.types.is_integer_dtype(df_real["Mês"]))
        self.assertTrue(pd.api.types.is_float_dtype(df_real["Valor"]))

    def test_09_load_all_data_integration(self):
        """Test load_all_data returns all 6 keys with expected shapes and zero crashes."""
        d = data_loader.load_all_data()
        expected_keys = {"fopa", "nomina", "afastados", "demitidos", "ferias", "valores_importados"}
        self.assertEqual(set(d.keys()), expected_keys)

        self.assertEqual(len(d["fopa"]), 15913)
        self.assertEqual(len(d["nomina"]), 16038)
        self.assertEqual(len(d["afastados"]), 1867)
        self.assertEqual(len(d["demitidos"]), 12901)
        self.assertEqual(len(d["ferias"]), 2005)
        self.assertEqual(len(d["valores_importados"]), 1334)

    # =========================================================================
    # MISSION 4: ADVERSARIAL STRESS TESTING (EDGE CASES & CORRUPTED DATA)
    # =========================================================================

    def test_10_adversarial_matricula_normalizer(self):
        """Stress-test normalize_matricula with extreme edge cases."""
        test_cases = [
            # (input, expected)
            ("12345", "12345"),
            ("0012345", "0012345"),
            (12345, "12345"),
            (12345.0, "12345"),
            ("12345.0", "12345"),
            ("12345.00", "12345"),
            ("  009876  ", "009876"),
            (None, ""),
            (np.nan, ""),
            ("-", ""),
            (" -", ""),
            ("null", ""),
            ("NaN", ""),
            ("#N/A", ""),
            (0, "0"),
            (0.0, "0"),
            ("0", "0"),
            ("0000", "0000"),
        ]
        for inp, expected in test_cases:
            with self.subTest(inp=inp):
                self.assertEqual(data_loader.normalize_matricula(inp), expected)

    def test_11_adversarial_currency_parser(self):
        """Stress-test parse_brazilian_currency with extreme values."""
        test_cases = [
            ("5.582,44", 5582.44),
            ("-", 0.0),
            (" -", 0.0),
            ("R$ 1.234,56", 1234.56),
            ("r$ 0,50", 0.50),
            ("100", 100.0),
            (100, 100.0),
            (100.5, 100.5),
            (None, 0.0),
            (np.nan, 0.0),
            ("0,00", 0.0),
            ("0.00", 0.0),
            ("-150,00", -150.00),
            (" -250,50 ", -250.50),
            ("999.999.999,99", 999999999.99),
        ]
        for inp, expected in test_cases:
            with self.subTest(inp=inp):
                self.assertAlmostEqual(data_loader.parse_brazilian_currency(inp), expected, places=2)

    def test_12_nonexistent_files_fail_fast(self):
        """Ensure clear FileNotFoundError is raised if paths do not exist."""
        with self.assertRaises(FileNotFoundError):
            data_loader.load_fopa("nonexistent_fopa.csv")
        with self.assertRaises(FileNotFoundError):
            data_loader.load_colaboradores("nonexistent_colaboradores.csv")
        with self.assertRaises(FileNotFoundError):
            data_loader.load_afastados("nonexistent_afastados.csv")
        with self.assertRaises(FileNotFoundError):
            data_loader.load_demitidos("nonexistent_demitidos.csv")
        with self.assertRaises(FileNotFoundError):
            data_loader.load_ferias("nonexistent_ferias.csv")
        with self.assertRaises(FileNotFoundError):
            data_loader.load_valores_importados("nonexistent_file.xlsx")


if __name__ == "__main__":
    unittest.main()
