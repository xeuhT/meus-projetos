"""
Testes Unitários do Módulo de Ingestão e Normalização (data_loader.py)
Auditoria "04 - GOOD CARD" - Cencosud Brasil

Valida:
- Normalização universal de Matrícula (zeros à esquerda, remoção de .0, nulos).
- Parsing de valores monetários no formato brasileiro (strings, floats, traços).
- Detecção de encoding e delimitador.
- Localização resiliente de arquivos FOPA.
- Pré-filtros e deduplicação de Demitidos (exclusão de Transferência sem ônus, data mais recente).
- Pré-filtros e deduplicação de Afastados (exclusão de termos anteriores ao mês, maior duração/aberto).
- Pré-filtros e deduplicação de Férias (sobreposição com mês auditado, prioridade de status ativo).
- Ingestão de Valores Importados (Excel e CSV).
- Função integradora load_all_data() contra bases reais do projeto.
- Constantes e conformidade do config.py.
"""

import os
import shutil
import tempfile
import unittest
from pathlib import Path

import numpy as np
import pandas as pd

import config
from data_loader import (
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


class TestNormalizeMatricula(unittest.TestCase):
    """Testes para a função de normalização universal de Matrícula."""

    def test_normalize_matricula_numeric(self):
        self.assertEqual(normalize_matricula(305204), "305204")
        self.assertEqual(normalize_matricula(305204.0), "305204")
        self.assertEqual(normalize_matricula("305204"), "305204")
        self.assertEqual(normalize_matricula("  305204  "), "305204")
        self.assertEqual(normalize_matricula("305204.0"), "305204")
        self.assertEqual(normalize_matricula("305204.00"), "305204")

    def test_normalize_matricula_leading_zeros(self):
        # Garante preservação de zeros à esquerda (essencial para evitar perda de dados)
        self.assertEqual(normalize_matricula("001234"), "001234")
        self.assertEqual(normalize_matricula("  001234  "), "001234")
        self.assertEqual(normalize_matricula("001234.0"), "001234")
        self.assertEqual(normalize_matricula("078"), "078")

    def test_normalize_matricula_nulls_and_hyphens(self):
        self.assertEqual(normalize_matricula(None), "")
        self.assertEqual(normalize_matricula(np.nan), "")
        self.assertEqual(normalize_matricula(pd.NA), "")
        self.assertEqual(normalize_matricula(""), "")
        self.assertEqual(normalize_matricula("-"), "")
        self.assertEqual(normalize_matricula(" -"), "")
        self.assertEqual(normalize_matricula("- "), "")
        self.assertEqual(normalize_matricula("--"), "")
        self.assertEqual(normalize_matricula("nan"), "")
        self.assertEqual(normalize_matricula("NaN"), "")
        self.assertEqual(normalize_matricula("#N/A"), "")

    def test_normalize_matricula_short_values(self):
        self.assertEqual(normalize_matricula("78"), "78")
        self.assertEqual(normalize_matricula(78), "78")
        self.assertEqual(normalize_matricula(78.0), "78")

    def test_normalize_matricula_series_vectorized(self):
        input_series = pd.Series([
            305204,
            305204.0,
            "001234",
            "001234.0",
            "  935  ",
            None,
            np.nan,
            "-",
            " -",
        ])
        expected = pd.Series([
            "305204",
            "305204",
            "001234",
            "001234",
            "935",
            "",
            "",
            "",
            "",
        ])
        result = normalize_matricula_series(input_series)
        pd.testing.assert_series_equal(result, expected)

    def test_normalize_matricula_handles_series_directly(self):
        # Se normalize_matricula for chamado passando uma Series inteira
        input_series = pd.Series([100, "0200.0", None])
        result = normalize_matricula(input_series)
        expected = pd.Series(["100", "0200", ""])
        pd.testing.assert_series_equal(result, expected)


class TestParseBrazilianCurrency(unittest.TestCase):
    """Testes para o conversor de moedas brasileiras para float64."""

    def test_parse_brazilian_currency_strings(self):
        self.assertEqual(parse_brazilian_currency("5.582,44"), 5582.44)
        self.assertEqual(parse_brazilian_currency("317,06"), 317.06)
        self.assertEqual(parse_brazilian_currency("0,00"), 0.0)
        self.assertEqual(parse_brazilian_currency("1.234.567,89"), 1234567.89)
        self.assertEqual(parse_brazilian_currency("  50,25  "), 50.25)

    def test_parse_brazilian_currency_hyphens_and_nulls(self):
        self.assertEqual(parse_brazilian_currency("-"), 0.0)
        self.assertEqual(parse_brazilian_currency(" -"), 0.0)
        self.assertEqual(parse_brazilian_currency("- "), 0.0)
        self.assertEqual(parse_brazilian_currency("--"), 0.0)
        self.assertEqual(parse_brazilian_currency(""), 0.0)
        self.assertEqual(parse_brazilian_currency(None), 0.0)
        self.assertEqual(parse_brazilian_currency(np.nan), 0.0)
        self.assertEqual(parse_brazilian_currency(pd.NA), 0.0)
        self.assertEqual(parse_brazilian_currency("#N/A"), 0.0)

    def test_parse_brazilian_currency_floats_and_integers(self):
        # Crucial: não pode multiplicar por 100 ao receber float nativo (ex: 301.85 não vira 30185.0)
        self.assertEqual(parse_brazilian_currency(301.85), 301.85)
        self.assertEqual(parse_brazilian_currency(0.0), 0.0)
        self.assertEqual(parse_brazilian_currency(5582.44), 5582.44)
        self.assertEqual(parse_brazilian_currency(100), 100.0)

    def test_parse_brazilian_currency_prefix(self):
        self.assertEqual(parse_brazilian_currency("R$ 5.582,44"), 5582.44)
        self.assertEqual(parse_brazilian_currency("r$ 300,00"), 300.0)

    def test_parse_brazilian_currency_series(self):
        input_series = pd.Series(["5.582,44", "-", 301.85, None, "0,00"])
        expected = pd.Series([5582.44, 0.0, 301.85, 0.0, 0.0], dtype=float)
        result = parse_brazilian_currency(input_series)
        pd.testing.assert_series_equal(result, expected)


class TestEncodingAndDelimiterDetection(unittest.TestCase):
    """Testes para detecção de encoding e delimitador de CSVs."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def test_detect_encoding_utf8(self):
        path = Path(self.test_dir) / "test_utf8.csv"
        with open(path, "w", encoding="utf-8") as f:
            f.write("Matrícula;Nome;Situação\n123;José;Ativo\n")
        self.assertEqual(detect_encoding(path), "utf-8")

    def test_detect_encoding_latin1(self):
        path = Path(self.test_dir) / "test_latin1.csv"
        # Bytes incompatíveis com UTF-8 estrito
        with open(path, "wb") as f:
            f.write("Matrícula;Situação\n123;Férias\n".encode("latin1"))
        self.assertEqual(detect_encoding(path), "latin1")

    def test_detect_encoding_file_not_found(self):
        with self.assertRaises(FileNotFoundError):
            detect_encoding(Path(self.test_dir) / "inexistente.csv")

    def test_detect_delimiter_semicolon(self):
        path = Path(self.test_dir) / "test_semi.csv"
        with open(path, "w", encoding="latin1") as f:
            f.write("Col1;Col2;Col3\n1;2;3\n4;5;6\n")
        self.assertEqual(detect_delimiter(path), ";")

    def test_detect_delimiter_comma(self):
        path = Path(self.test_dir) / "test_comma.csv"
        with open(path, "w", encoding="latin1") as f:
            f.write("Col1,Col2,Col3\n1,2,3\n4,5,6\n")
        self.assertEqual(detect_delimiter(path), ",")

    def test_detect_delimiter_file_not_found(self):
        with self.assertRaises(FileNotFoundError):
            detect_delimiter(Path(self.test_dir) / "inexistente.csv")


class TestFindFopaFile(unittest.TestCase):
    """Testes para localização do arquivo FOPA via padrão glob."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def test_find_fopa_file_found(self):
        fopa_file = Path(self.test_dir) / "FOPA_Eventos_por_Colaborador - 22092026_091703_796705.csv"
        fopa_file.write_text("dummy", encoding="utf-8")
        found = find_fopa_file(self.test_dir)
        self.assertEqual(Path(found).resolve(), fopa_file.resolve())

    def test_find_fopa_file_not_found_raises(self):
        with self.assertRaises(FileNotFoundError):
            find_fopa_file(self.test_dir)

    def test_find_fopa_file_real_directory(self):
        real_dir = Path("Relatorio Original")
        if real_dir.exists():
            found = find_fopa_file(real_dir)
            self.assertTrue(os.path.exists(found))
            self.assertIn("FOPA_Eventos_por_", found)


class TestLoadDemitidos(unittest.TestCase):
    """Testes para ingestão e deduplicação de Demitidos."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def test_deduplicate_demitidos_transferencia_sem_onus(self):
        csv_path = Path(self.test_dir) / "demitidos_test.csv"
        # 1 registro normal e 1 de transferência sem ônus
        content = (
            "Matrícula;Data Rescisão;Tipo de Rescisão\n"
            "1001;15/09/2026;2 - Sem justa causa do empregador\n"
            "1002;10/09/2026;5 - Transferência sem ônus\n"
        )
        csv_path.write_text(content, encoding="latin1")

        df = load_demitidos(csv_path)
        self.assertEqual(len(df), 1)
        self.assertEqual(df.iloc[0]["Matrícula"], "1001")

    def test_deduplicate_demitidos_most_recent_date(self):
        csv_path = Path(self.test_dir) / "demitidos_dedup.csv"
        # Colaborador com duas rescisões: uma antiga e uma recente
        content = (
            "Matrícula;Data Rescisão;Tipo de Rescisão\n"
            "2001;01/07/2025;2 - Sem justa causa do empregador\n"
            "2001;01/09/2026;2 - Sem justa causa do empregador\n"
        )
        csv_path.write_text(content, encoding="latin1")

        df = load_demitidos(csv_path)
        self.assertEqual(len(df), 1)
        self.assertEqual(df.iloc[0]["Matrícula"], "2001")
        self.assertEqual(df.iloc[0]["Data Rescisão"], "01/09/2026")


class TestLoadAfastados(unittest.TestCase):
    """Testes para ingestão e deduplicação de Afastados."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def test_deduplicate_afastados_past_termination_filtered(self):
        csv_path = Path(self.test_dir) / "afastados_test.csv"
        # Colaborador encerrado em 15/08/2026 (antes de 01/09/2026) -> deve ser descartado
        # Colaborador com término em 10/09/2026 -> deve ser mantido
        content = (
            "Matrícula;Início Afastamento;Término Afastamento;Dias Afastamento;Tipo Afastamento\n"
            "3001;01/08/2026;15/08/2026;15;5- Motivo de Doença\n"
            "3002;01/08/2026;10/09/2026;40;5- Motivo de Doença\n"
        )
        csv_path.write_text(content, encoding="latin1")

        df = load_afastados(csv_path, ref_date="2026-09-01")
        self.assertEqual(len(df), 1)
        self.assertEqual(df.iloc[0]["Matrícula"], "3002")

    def test_deduplicate_afastados_open_ended_priority(self):
        csv_path = Path(self.test_dir) / "afastados_open.csv"
        # Mesmo colaborador com 1 período encerrado e 1 em aberto (Término vazio)
        content = (
            "Matrícula;Início Afastamento;Término Afastamento;Dias Afastamento;Tipo Afastamento\n"
            "4001;01/05/2026;15/09/2026;137;5- Motivo de Doença\n"
            "4001;16/09/2026;;300;5- Motivo de Doença\n"
        )
        csv_path.write_text(content, encoding="latin1")

        df = load_afastados(csv_path, ref_date="2026-09-01")
        self.assertEqual(len(df), 1)
        self.assertEqual(df.iloc[0]["Matrícula"], "4001")
        self.assertEqual(df.iloc[0]["Início Afastamento"], "16/09/2026")
        self.assertTrue(pd.isna(df.iloc[0]["Término Afastamento"]) or df.iloc[0]["Término Afastamento"] == "")


class TestLoadFerias(unittest.TestCase):
    """Testes para ingestão e deduplicação de Férias."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def test_deduplicate_ferias_outside_audit_month(self):
        csv_path = Path(self.test_dir) / "ferias_test.csv"
        # skiprows=1 no formato padrão (linha 1: título, linha 2: colunas)
        content = (
            "Férias de Colaboradores\n"
            "Matrícula;Início do Período de Gozo;Fim do Período de Gozo;Dias Gozo;Status\n"
            "5001;01/07/2026;30/07/2026;30;Gozada\n"
            "5002;15/09/2026;24/09/2026;10;Em férias\n"
        )
        csv_path.write_text(content, encoding="latin1")

        df = load_ferias(csv_path, year=2026, month=9)
        self.assertEqual(len(df), 1)
        self.assertEqual(df.iloc[0]["Matrícula"], "5002")

    def test_deduplicate_ferias_active_status_priority(self):
        csv_path = Path(self.test_dir) / "ferias_overlap.csv"
        # Mesmo colaborador com 2 períodos fracionados em setembro: um 'Gozada' e um 'Em férias'
        content = (
            "Férias de Colaboradores\n"
            "Matrícula;Início do Período de Gozo;Fim do Período de Gozo;Dias Gozo;Status\n"
            "6001;19/08/2026;02/09/2026;15;Gozada\n"
            "6001;16/09/2026;30/09/2026;15;Em férias\n"
        )
        csv_path.write_text(content, encoding="latin1")

        df = load_ferias(csv_path, year=2026, month=9)
        self.assertEqual(len(df), 1)
        self.assertEqual(df.iloc[0]["Matrícula"], "6001")
        self.assertEqual(df.iloc[0]["Status"], "Em férias")
        self.assertEqual(df.iloc[0]["Início do Período de Gozo"], "16/09/2026")


class TestLoadValoresImportados(unittest.TestCase):
    """Testes para carga de Valores Importados via CSV e Excel."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def test_load_valores_importados_from_csv(self):
        csv_path = Path(self.test_dir) / "valores_test.csv"
        content = (
            "Emp;Matrícula;Evento;Mês;Valor\n"
            "7;001234.0;29613;92026;5.582,44\n"
            "7;305204;29613;92026;301,85\n"
        )
        csv_path.write_text(content, encoding="latin1")

        df = load_valores_importados(csv_path)
        self.assertEqual(len(df), 2)
        self.assertIn("Matrícula", df.columns)
        self.assertIn("Valor", df.columns)
        self.assertEqual(df.iloc[0]["Matrícula"], "001234")
        self.assertEqual(df.iloc[0]["Valor"], 5582.44)
        self.assertEqual(df.iloc[1]["Matrícula"], "305204")
        self.assertEqual(df.iloc[1]["Valor"], 301.85)

    def test_load_valores_importados_from_excel_real(self):
        excel_path = Path("Conferência Good Card_Base.xlsx")
        if excel_path.exists():
            df = load_valores_importados(excel_path)
            self.assertEqual(len(df), 1334)
            self.assertAlmostEqual(df["Valor"].sum(), 482265.63, places=2)
            self.assertEqual(df["Evento"].iloc[0], 29613)


class TestRealFilesDataLoading(unittest.TestCase):
    """Testes de conformidade e integridade contra as bases de dados reais do projeto."""

    def test_load_fopa_real_file(self):
        data_dir = Path("Relatorio Original")
        if not data_dir.exists():
            self.skipTest("Diretório Relatorio Original não encontrado.")

        fopa_file = find_fopa_file(data_dir)
        df_fopa = load_fopa(fopa_file)

        self.assertEqual(len(df_fopa), 15913)
        self.assertEqual(len(df_fopa.columns), 23)
        self.assertIn("Matrícula", df_fopa.columns)

        # Validação do nome exato da coluna 29622 com espaço no final
        col_29622 = "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant "
        self.assertIn(col_29622, df_fopa.columns)

        # Validação de conversão monetária para float64
        self.assertTrue(pd.api.types.is_float_dtype(df_fopa["11010 - Salario Mensal"]))
        self.assertTrue(pd.api.types.is_float_dtype(df_fopa["6280 - Desc Cartão Benefícios Especiais"]))
        self.assertTrue(pd.api.types.is_float_dtype(df_fopa["29613 - Ben Esp - Valor Informado para Desconto"]))

    def test_load_colaboradores_real_file(self):
        data_dir = Path("Relatorio Original")
        if not data_dir.exists():
            self.skipTest("Diretório Relatorio Original não encontrado.")

        cola_file = data_dir / "COLA_Colaboradores.CSV"
        df_cola = load_colaboradores(cola_file)

        self.assertEqual(len(df_cola), 16038)
        self.assertEqual(len(df_cola.columns), 109)
        self.assertIn("Matrícula", df_cola.columns)
        self.assertTrue(pd.api.types.is_string_dtype(df_cola["Matrícula"]))

    def test_load_all_data_integration(self):
        data_dir = Path("Relatorio Original")
        template_file = Path("Conferência Good Card_Base.xlsx")
        if not (data_dir.exists() and template_file.exists()):
            self.skipTest("Bases reais não encontradas.")

        all_data = load_all_data(data_dir=data_dir, template_path=template_file)

        self.assertIn("fopa", all_data)
        self.assertIn("nomina", all_data)
        self.assertIn("afastados", all_data)
        self.assertIn("demitidos", all_data)
        self.assertIn("ferias", all_data)
        self.assertIn("valores_importados", all_data)

        self.assertEqual(len(all_data["fopa"]), 15913)
        self.assertEqual(len(all_data["nomina"]), 16038)
        self.assertEqual(len(all_data["afastados"]), 1867)
        self.assertEqual(len(all_data["demitidos"]), 12901)
        self.assertEqual(len(all_data["ferias"]), 2005)
        self.assertEqual(len(all_data["valores_importados"]), 1334)


class TestConfigConstants(unittest.TestCase):
    """Testes de conformidade para o módulo config.py."""

    def test_config_conferencia_columns(self):
        self.assertEqual(len(config.COLUNAS_OFICIAIS_CONFERENCIA), 38)
        self.assertEqual(config.COLUNAS_OFICIAIS_CONFERENCIA[0], "Código da Empresa")
        self.assertEqual(config.COLUNAS_OFICIAIS_CONFERENCIA[6], "Matrícula")
        self.assertEqual(
            config.COLUNAS_OFICIAIS_CONFERENCIA[15],
            "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ",
        )
        self.assertEqual(config.COLUNAS_OFICIAIS_CONFERENCIA[-1], "Tratativa")

    def test_config_valores_importados_columns(self):
        self.assertEqual(len(config.COLUNAS_OFICIAIS_VALORES_IMPORTADOS), 11)
        self.assertEqual(config.COLUNAS_OFICIAIS_VALORES_IMPORTADOS[0], "Emp")
        self.assertEqual(config.COLUNAS_OFICIAIS_VALORES_IMPORTADOS[1], "Matrícula")
        self.assertEqual(config.COLUNAS_OFICIAIS_VALORES_IMPORTADOS[-1], "Tratativa")

    def test_config_status_lista(self):
        self.assertEqual(len(config.STATUS_LISTA), 9)
        self.assertEqual(config.STATUS_LISTA[0], "Ok - Sem saldo devedor e sem desconto")
        self.assertEqual(config.STATUS_LISTA[7], "Aprendiz - Verificar")

    def test_config_tolerancia_centavos(self):
        self.assertEqual(config.TOLERANCIA_CENTAVOS, 0.02)

    def test_config_paleta_cores(self):
        self.assertIn("azul_cencosud", config.PALETA_CORES)
        self.assertEqual(config.PALETA_CORES["azul_cencosud"], "1F497D")
        self.assertIn("divergencia_fill", config.PALETA_CORES)
        self.assertEqual(config.PALETA_CORES["divergencia_fill"], "FFC7CE")
        self.assertIn("divergencia_font", config.PALETA_CORES)
        self.assertEqual(config.PALETA_CORES["divergencia_font"], "9C0006")


if __name__ == "__main__":
    unittest.main()
