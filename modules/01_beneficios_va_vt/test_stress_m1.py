"""
Adversarial Stress Test Suite - Milestone 1 (config.py & data_loader.py)
Auditoria "04 - GOOD CARD" - Cencosud Brasil

Coverage:
1. normalize_matricula with extreme inputs (spaces, non-numeric, sentinels, high precision, large numbers, series)
2. parse_brazilian_currency with extreme inputs (dirty strings, malformed numbers, negative numbers, spaces, non-numeric, NaN, nulls, high precision)
3. Loaders stress-testing (missing columns, empty files, corrupted date formats, duplicate records, filtering)
4. load_all_data integrity, multi-cycle execution, and memory sanity check
"""

import gc
import os
import shutil
import tempfile
import tracemalloc
import unittest
from pathlib import Path

import numpy as np
import pandas as pd

import config
from data_loader import (
    _parse_currency_scalar,
    _parse_date_series,
    detect_delimiter,
    detect_encoding,
    find_fopa_file,
    load_afastados,
    load_all_data,
    load_colaboradores,
    load_demitidos,
    load_ferias,
    load_fopa,
    load_valores_importados,
    normalize_matricula,
    normalize_matricula_series,
    parse_brazilian_currency,
)


class StressTestNormalizeMatricula(unittest.TestCase):
    """Stress testing normalize_matricula with adversarial and extreme inputs."""

    def test_dirty_and_whitespace_inputs(self):
        self.assertEqual(normalize_matricula("  001234  "), "001234")
        self.assertEqual(normalize_matricula("\t001234\n"), "001234")
        self.assertEqual(normalize_matricula("\xa0001234\xa0"), "001234")
        self.assertEqual(normalize_matricula(""), "")
        self.assertEqual(normalize_matricula("   "), "")
        self.assertEqual(normalize_matricula("\n\t  \r"), "")

    def test_float_and_integer_variations(self):
        self.assertEqual(normalize_matricula(305204), "305204")
        self.assertEqual(normalize_matricula(305204.0), "305204")
        self.assertEqual(normalize_matricula("305204.0"), "305204")
        self.assertEqual(normalize_matricula("305204.00"), "305204")
        self.assertEqual(normalize_matricula("305204.000"), "305204")
        self.assertEqual(normalize_matricula("001234.0"), "001234")
        self.assertEqual(normalize_matricula("001234.00"), "001234")
        self.assertEqual(normalize_matricula("001234.0000"), "001234")
        self.assertEqual(normalize_matricula(0), "0")
        self.assertEqual(normalize_matricula(0.0), "0")
        self.assertEqual(normalize_matricula("0"), "0")
        self.assertEqual(normalize_matricula("000"), "000")

    def test_sentinels_and_nulls(self):
        sentinels = [
            None,
            np.nan,
            pd.NA,
            "",
            "-",
            " -",
            "- ",
            "--",
            "nan",
            "NaN",
            "<NA>",
            "None",
            "null",
            "NULL",
            "#N/A",
        ]
        for val in sentinels:
            self.assertEqual(normalize_matricula(val), "", f"Failed for sentinel: {val}")

    def test_non_numeric_and_alphanumeric(self):
        self.assertEqual(normalize_matricula("M12345"), "M12345")
        self.assertEqual(normalize_matricula("ABC-999"), "ABC-999")
        self.assertEqual(normalize_matricula("007A"), "007A")
        self.assertEqual(normalize_matricula("123/45"), "123/45")
        self.assertEqual(normalize_matricula("123 45"), "123 45")

    def test_extreme_numbers(self):
        self.assertEqual(normalize_matricula(999999999999999), "999999999999999")
        self.assertEqual(normalize_matricula(-100), "-100")
        self.assertEqual(normalize_matricula(-100.0), "-100")
        self.assertEqual(normalize_matricula("-100.0"), "-100")

    def test_series_stress_vectorized(self):
        data = [
            "  001234  ",
            305204.0,
            "305204.00",
            None,
            np.nan,
            pd.NA,
            "-",
            "078",
            "M999",
            999999999,
        ] * 1000
        s = pd.Series(data)
        res = normalize_matricula_series(s)
        self.assertEqual(len(res), 10000)
        self.assertEqual(res.iloc[0], "001234")
        self.assertEqual(res.iloc[1], "305204")
        self.assertEqual(res.iloc[2], "305204")
        self.assertEqual(res.iloc[3], "")
        self.assertEqual(res.iloc[4], "")
        self.assertEqual(res.iloc[5], "")
        self.assertEqual(res.iloc[6], "")
        self.assertEqual(res.iloc[7], "078")
        self.assertEqual(res.iloc[8], "M999")
        self.assertEqual(res.iloc[9], "999999999")


class StressTestParseBrazilianCurrency(unittest.TestCase):
    """Stress testing parse_brazilian_currency with adversarial and extreme inputs."""

    def test_standard_and_prefixed_brazilian_currency(self):
        self.assertEqual(parse_brazilian_currency("5.582,44"), 5582.44)
        self.assertEqual(parse_brazilian_currency("317,06"), 317.06)
        self.assertEqual(parse_brazilian_currency("0,00"), 0.0)
        self.assertEqual(parse_brazilian_currency("1.234.567,89"), 1234567.89)
        self.assertEqual(parse_brazilian_currency("R$ 5.582,44"), 5582.44)
        self.assertEqual(parse_brazilian_currency("r$ 50,00"), 50.0)
        self.assertEqual(parse_brazilian_currency("R$  100,50 "), 100.50)
        self.assertEqual(parse_brazilian_currency("  R$ 0,05  "), 0.05)

    def test_numeric_natives(self):
        self.assertEqual(parse_brazilian_currency(301.85), 301.85)
        self.assertEqual(parse_brazilian_currency(0.0), 0.0)
        self.assertEqual(parse_brazilian_currency(0), 0.0)
        self.assertEqual(parse_brazilian_currency(100), 100.0)
        self.assertEqual(parse_brazilian_currency(5582.44), 5582.44)

    def test_sentinels_and_nulls(self):
        sentinels = [
            "",
            "   ",
            "-",
            " -",
            "- ",
            "--",
            None,
            np.nan,
            pd.NA,
            "#N/A",
            "null",
            "NULL",
            "None",
            "<NA>",
            "nan",
            "NaN",
            "0,00",
            "0.00",
        ]
        for val in sentinels:
            self.assertEqual(parse_brazilian_currency(val), 0.0, f"Failed for: {val}")

    def test_negative_values(self):
        self.assertEqual(parse_brazilian_currency("-50,25"), -50.25)
        self.assertEqual(parse_brazilian_currency("-1.500,00"), -1500.0)
        self.assertEqual(parse_brazilian_currency(-50.25), -50.25)
        self.assertEqual(parse_brazilian_currency(-100), -100.0)
        # Note on spacing: "- 50,25" has a space between minus and number
        self.assertEqual(parse_brazilian_currency("- 50,25"), 0.0)

    def test_extreme_precision_and_large_numbers(self):
        self.assertEqual(parse_brazilian_currency(0.00001), 0.0)
        self.assertEqual(parse_brazilian_currency(0.005), 0.01)
        self.assertEqual(parse_brazilian_currency(999999999.99), 999999999.99)
        self.assertEqual(parse_brazilian_currency("999.999.999,99"), 999999999.99)
        self.assertEqual(parse_brazilian_currency("0,0001"), 0.0)

    def test_dirty_and_malformed_strings(self):
        self.assertEqual(parse_brazilian_currency("N/D"), 0.0)
        self.assertEqual(parse_brazilian_currency("PENDENTE"), 0.0)
        self.assertEqual(parse_brazilian_currency("TOTAL: R$ 500,00"), 0.0)
        self.assertEqual(parse_brazilian_currency("abc"), 0.0)
        self.assertEqual(parse_brazilian_currency("12.34.56,78"), 123456.78)

    def test_series_stress_performance(self):
        data = [
            "5.582,44",
            "-",
            " -",
            "301.85",
            None,
            np.nan,
            "R$ 1.000,00",
            "-50,25",
            "0,00",
            "PENDENTE",
        ] * 1000
        s = pd.Series(data)
        res = parse_brazilian_currency(s)
        self.assertEqual(len(res), 10000)
        self.assertTrue(pd.api.types.is_float_dtype(res))
        self.assertEqual(res.iloc[0], 5582.44)
        self.assertEqual(res.iloc[1], 0.0)
        self.assertEqual(res.iloc[2], 0.0)
        self.assertEqual(res.iloc[3], 301.85)
        self.assertEqual(res.iloc[4], 0.0)
        self.assertEqual(res.iloc[5], 0.0)
        self.assertEqual(res.iloc[6], 1000.0)
        self.assertEqual(res.iloc[7], -50.25)
        self.assertEqual(res.iloc[8], 0.0)
        self.assertEqual(res.iloc[9], 0.0)


class StressTestDataLoadersResilience(unittest.TestCase):
    """Stress testing file loaders under edge cases and corrupted structures."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_load_fopa_missing_monetary_columns(self):
        csv_path = Path(self.test_dir) / "fopa_sparse.csv"
        content = (
            "Meta 1\nMeta 2\nMeta 3\n"
            "Código da Empresa;Nome da Empresa;Matrícula;Nome do Colaborador;11010 - Salario Mensal\n"
            "1;Cencosud;001234.0;João Silva;5.000,00\n"
            "1;Cencosud;005678;Maria Santos;-\n"
        )
        csv_path.write_text(content, encoding="latin1")

        df = load_fopa(csv_path)
        self.assertEqual(len(df), 2)
        self.assertIn("Matrícula", df.columns)
        self.assertEqual(df.iloc[0]["Matrícula"], "001234")
        self.assertEqual(df.iloc[0]["11010 - Salario Mensal"], 5000.0)
        self.assertEqual(df.iloc[1]["Matrícula"], "005678")
        self.assertEqual(df.iloc[1]["11010 - Salario Mensal"], 0.0)

    def test_load_fopa_filters_empty_matricula(self):
        csv_path = Path(self.test_dir) / "fopa_empty_mat.csv"
        content = (
            "Meta 1\nMeta 2\nMeta 3\n"
            "Matrícula;Nome do Colaborador;11010 - Salario Mensal\n"
            "1001;Valido;1.000,00\n"
            ";Sem Matricula;2.000,00\n"
            "-;Traco;3.000,00\n"
            "NaN;Nulo;4.000,00\n"
        )
        csv_path.write_text(content, encoding="latin1")

        df = load_fopa(csv_path)
        self.assertEqual(len(df), 1)
        self.assertEqual(df.iloc[0]["Matrícula"], "1001")

    def test_load_afastados_corrupted_dates(self):
        csv_path = Path(self.test_dir) / "afastados_corrupt_date.csv"
        content = (
            "Matrícula;Início Afastamento;Término Afastamento;Dias Afastamento;Tipo Afastamento\n"
            "101;01/09/2026;15/09/2026;15;Doença\n"
            "102;INVALID;CORRUPT;30;Acidente\n"
            "103;01/09/2026;;90;Maternidade\n"
        )
        csv_path.write_text(content, encoding="latin1")

        df = load_afastados(csv_path, ref_date="2026-09-01")
        # 101 is valid, 103 is open-ended (Término vacío -> NaT -> mask_vigente=True)
        # 102 Término is CORRUPT -> NaT -> mask_vigente=True (treated as open)
        self.assertIn("101", df["Matrícula"].values)
        self.assertIn("103", df["Matrícula"].values)

    def test_load_demitidos_duplicate_rescisao_dates(self):
        csv_path = Path(self.test_dir) / "demitidos_duplicates.csv"
        content = (
            "Matrícula;Data Rescisão;Tipo de Rescisão\n"
            "201;15/05/2025;2 - Sem justa causa\n"
            "201;20/09/2026;2 - Sem justa causa\n"
            "201;10/01/2024;2 - Sem justa causa\n"
            "202;15/09/2026;5 - Transferência sem ônus\n"
        )
        csv_path.write_text(content, encoding="latin1")

        df = load_demitidos(csv_path)
        self.assertEqual(len(df), 1)
        self.assertEqual(df.iloc[0]["Matrícula"], "201")
        self.assertEqual(df.iloc[0]["Data Rescisão"], "20/09/2026")

    def test_load_ferias_status_and_overlap_priority(self):
        csv_path = Path(self.test_dir) / "ferias_stress.csv"
        content = (
            "Header linha 1\n"
            "Matrícula;Início Férias;Final Férias;Qtd de dias;Status\n"
            "301;01/07/2026;30/07/2026;30;Gozada\n"  # Outside Sep 2026 -> filtered
            "302;20/08/2026;05/09/2026;15;Gozada\n"  # Overlaps Sep 2026 -> kept
            "302;15/09/2026;30/09/2026;15;Em férias\n"  # Overlaps Sep 2026, Em férias -> priority!
        )
        csv_path.write_text(content, encoding="latin1")

        df = load_ferias(csv_path, year=2026, month=9)
        self.assertEqual(len(df), 1)
        self.assertEqual(df.iloc[0]["Matrícula"], "302")
        self.assertEqual(df.iloc[0]["Status"], "Em férias")
        self.assertEqual(df.iloc[0]["Início Férias"], "15/09/2026")

    def test_load_valores_importados_csv_resilience(self):
        csv_path = Path(self.test_dir) / "val_imp_valid.csv"
        content = (
            "Emp;Matrícula;Evento;Mês;Valor\n"
            "10;004455.0;29613;92026;1.250,50\n"
            "10;007788;29613;92026;-\n"
            "10;;29613;92026;300,00\n"  # Empty matricula -> filtered
        )
        csv_path.write_text(content, encoding="latin1")

        df = load_valores_importados(csv_path)
        self.assertEqual(len(df), 2)
        self.assertEqual(df.iloc[0]["Matrícula"], "004455")
        self.assertEqual(df.iloc[0]["Valor"], 1250.50)
        self.assertEqual(df.iloc[1]["Matrícula"], "007788")
        self.assertEqual(df.iloc[1]["Valor"], 0.0)

    def test_missing_matricula_column_raises_keyerror(self):
        """
        Adversarial test confirming finding:
        When 'Matrícula' column is absent, deduplication functions raise KeyError.
        """
        f = Path(self.test_dir) / "no_mat.csv"
        f.write_text("Data Rescisão;Tipo de Rescisão\n01/09/2026;2\n", encoding="latin1")

        with self.assertRaises(KeyError):
            load_demitidos(f)


class StressTestRealFilesAndMemorySanity(unittest.TestCase):
    """Integrity and memory sanity check on the production dataset."""

    def test_real_files_load_and_data_integrity(self):
        data_dir = config.DIR_RELATORIO_ORIGINAL
        template = config.TEMPLATE_BASE_EXCEL

        if not (data_dir.exists() and template.exists()):
            self.skipTest("Production data files not found.")

        data = load_all_data(data_dir=data_dir, template_path=template)

        self.assertIn("fopa", data)
        self.assertIn("nomina", data)
        self.assertIn("afastados", data)
        self.assertIn("demitidos", data)
        self.assertIn("ferias", data)
        self.assertIn("valores_importados", data)

        # Exact expected record counts
        self.assertEqual(len(data["fopa"]), 15913)
        self.assertEqual(len(data["nomina"]), 16038)
        self.assertEqual(len(data["afastados"]), 1867)
        self.assertEqual(len(data["demitidos"]), 12901)
        self.assertEqual(len(data["ferias"]), 2005)
        self.assertEqual(len(data["valores_importados"]), 1334)

        # Integrity of Matrícula in all datasets
        for name, df in data.items():
            self.assertIn("Matrícula", df.columns, f"Matrícula missing in {name}")
            self.assertFalse((df["Matrícula"] == "").any(), f"Empty Matrícula found in {name}")
            self.assertFalse(df["Matrícula"].isna().any(), f"NaN Matrícula found in {name}")

        # Ensure FOPA financial columns are strictly float64
        df_fopa = data["fopa"]
        for col in config.COLUNAS_MONETARIAS_FOPA:
            self.assertIn(col, df_fopa.columns)
            self.assertTrue(pd.api.types.is_float_dtype(df_fopa[col]))

    def test_multi_cycle_memory_leak_sanity(self):
        """
        Stress test: Execute load_all_data multiple times to ensure
        memory usage stabilizes and there are no circular reference leaks.
        """
        data_dir = config.DIR_RELATORIO_ORIGINAL
        template = config.TEMPLATE_BASE_EXCEL

        if not (data_dir.exists() and template.exists()):
            self.skipTest("Production data files not found.")

        tracemalloc.start()
        gc.collect()

        mem_snapshots = []
        for cycle in range(3):
            data = load_all_data(data_dir=data_dir, template_path=template)
            self.assertEqual(len(data["fopa"]), 15913)
            del data
            gc.collect()
            current, peak = tracemalloc.get_traced_memory()
            mem_snapshots.append(current)

        tracemalloc.stop()

        # Check memory growth between cycle 1 and cycle 3
        # In a leak-free environment with gc.collect(), current memory growth should be bounded (< 25 MB)
        growth_mb = (mem_snapshots[-1] - mem_snapshots[0]) / (1024 * 1024)
        self.assertLess(
            growth_mb,
            25.0,
            f"Excessive memory growth detected across load cycles: {growth_mb:.2f} MB",
        )


if __name__ == "__main__":
    unittest.main()
