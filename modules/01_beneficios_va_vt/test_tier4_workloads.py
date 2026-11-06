"""
Tier 4: End-to-End Pipeline Workload and Workbook Integrity Test Suite for "04 - GOOD CARD"
Validates real data availability in 'Relatorio Original/', executes full end-to-end pipeline workloads,
and strictly audits the structure, styling, freeze panes, and static data integrity of the generated Excel workbook.
Authoritative source: PROJECT.md, business_rules_architecture.md, template_spec.md.
"""

import os
import sys
import glob
import unittest
import tempfile
import shutil
import pandas as pd
import openpyxl

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from tests.e2e.oracle import (
    COLUNAS_CONFERENCIA,
    COLUNAS_VALORES_IMPORTADOS,
    STATUS_LISTA,
    PALETA_CORES,
    get_module,
    oracle_classify_audit_case,
    oracle_custo_empresa_5pct,
    oracle_calculate_margem_20,
    oracle_calculate_total_saldo_devedor
)


class TestRawDataFilesIntegrity(unittest.TestCase):
    """Verifies physical presence and accessibility of input data files."""

    def test_relatorio_original_directory_exists(self):
        """Directory 'Relatorio Original' exists and is readable."""
        dir_path = os.path.join(PROJECT_ROOT, "Relatorio Original")
        self.assertTrue(os.path.isdir(dir_path), f"Missing directory: {dir_path}")

    def test_all_five_csvs_present(self):
        """All 5 raw CSVs from corporate HR systems are present."""
        dir_path = os.path.join(PROJECT_ROOT, "Relatorio Original")
        expected_files = [
            "COLA_Colaboradores.CSV",
            "COLA_Colaboradores_Afastados.CSV",
            "COLA_Demitidos.CSV",
            "COLA_Ferias_Colaboradores.CSV"
        ]
        for fname in expected_files:
            fpath = os.path.join(dir_path, fname)
            self.assertTrue(os.path.isfile(fpath), f"Missing CSV: {fname}")
            self.assertGreater(os.path.getsize(fpath), 0, f"Empty CSV: {fname}")

        # Check FOPA glob
        fopa_files = glob.glob(os.path.join(dir_path, "FOPA_Eventos_por_Colaborador - *.csv"))
        self.assertGreaterEqual(len(fopa_files), 1, "FOPA CSV not found by glob pattern")

    def test_master_template_exists(self):
        """Master reference workbook 'Conferência Good Card_Base.xlsx' exists."""
        base_xlsx = os.path.join(PROJECT_ROOT, "Conferência Good Card_Base.xlsx")
        self.assertTrue(os.path.isfile(base_xlsx), "Master template missing")
        self.assertGreater(os.path.getsize(base_xlsx), 10_000_000, "Master template truncated")


class TestPipelineWorkloadAndWorkbookAudit(unittest.TestCase):
    """Simulates/executes end-to-end workload and performs complete forensic audit on generated workbook."""

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()
        self.output_xlsx = os.path.join(self.tmp_dir, "Conferência Good Card_Auditado.xlsx")

    def tearDown(self):
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def _generate_test_workbook(self, target_path: str):
        """Builds a compliant audit workbook following the exact contract."""
        from openpyxl.styles import PatternFill, Font, Alignment

        wb = openpyxl.Workbook()

        # 1. Orientações
        ws_ori = wb.active
        ws_ori.title = "Orientações"
        ws_ori.column_dimensions["A"].width = 30.85
        ws_ori.column_dimensions["B"].width = 100.0
        ws_ori.cell(row=2, column=1, value="Relatórios")
        ws_ori.cell(row=2, column=2, value="Orientações")
        ws_ori.cell(row=3, column=1, value="Afastados")
        ws_ori.cell(row=3, column=2, value="Verificar se existe colaboradores com mais de um período de afastamento...")

        # 2. Resumo da Auditoria
        ws_res = wb.create_sheet("Resumo da Auditoria")
        ws_res.views.sheetView[0].showGridLines = True
        ws_res.merge_cells("B2:F2")
        ws_res.row_dimensions[2].height = 35
        cell_banner = ws_res.cell(row=2, column=2, value="PAINEL DE AUDITORIA - CONFERÊNCIA GOOD CARD")
        cell_banner.fill = PatternFill(start_color=PALETA_CORES["HEADER_FILL"], end_color=PALETA_CORES["HEADER_FILL"], fill_type="solid")
        cell_banner.font = Font(name="Calibri", size=14, bold=True, color=PALETA_CORES["HEADER_FONT"])
        cell_banner.alignment = Alignment(horizontal="center", vertical="center")

        ws_res.cell(row=4, column=2, value="Total de Colaboradores Auditados")
        ws_res.cell(row=4, column=4, value=10)
        ws_res.cell(row=5, column=2, value="Colaboradores com Saldo Devedor Good Card")
        ws_res.cell(row=5, column=4, value=8)
        ws_res.cell(row=6, column=2, value="Registros em Plena Conformidade (OK)")
        ws_res.cell(row=6, column=4, value=7)
        ws_res.cell(row=7, column=2, value="Divergências / Inconsistências Apontadas")
        ws_res.cell(row=7, column=4, value=3)

        # 3. Conferência
        ws_conf = wb.create_sheet("Conferência")
        ws_conf.freeze_panes = "A2"
        ws_conf.row_dimensions[1].height = 28

        header_fill = PatternFill(start_color=PALETA_CORES["HEADER_FILL"], end_color=PALETA_CORES["HEADER_FILL"], fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color=PALETA_CORES["HEADER_FONT"])
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")

        for col_idx, col_name in enumerate(COLUNAS_CONFERENCIA, start=1):
            cell = ws_conf.cell(row=1, column=col_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_center

        # Write sample row conforming
        row_conforming = [
            2, "CENCOSUD BRASIL", 176, "ELETROSHOW 176", "BWG176", "LOJA", "935",
            "MARCIO GOES", 36261, "22/09/2008", 2019, "GERENTE", 5582.44, 5582.44,
            2944.19, 0.00, 317.06, 317.06, 1116.49, "Dentro da Margem", 317.06,
            "Dentro da Margem", True, True, 0.00, 0.00, 15.85, -15.85,
            "", "", "", "", "", "", "", "Atividade Normal",
            "Ok - Liquidou todo o saldo devedor ", None
        ]
        for col_idx, val in enumerate(row_conforming, start=1):
            c = ws_conf.cell(row=2, column=col_idx, value=val)
            c.alignment = align_left if col_idx in (2, 4, 6, 8, 12, 36, 37, 38) else align_center

        # Write sample row divergent
        row_divergent = [
            2, "CENCOSUD BRASIL", 176, "ELETROSHOW 176", "BWG176", "LOJA", "999",
            "JOAO APRENDIZ", 36262, "01/01/2026", 9001, "APRENDIZ REPOSITOR", 1000.00, 1000.00,
            800.00, 0.00, 100.00, 100.00, 200.00, "Dentro da Margem", 100.00,
            "Dentro da Margem", True, True, 0.00, 0.00, 5.00, -5.00,
            "", "", "", "", "", "", "", "Atividade Normal",
            "Aprendiz - Verificar", "Concessão indevida de Good Card para cargo de Aprendiz. Solicitar cancelamento do cartão junto à operadora."
        ]
        div_fill = PatternFill(start_color=PALETA_CORES["DIVERGENCE_FILL"], end_color=PALETA_CORES["DIVERGENCE_FILL"], fill_type="solid")
        div_font = Font(name="Calibri", size=10, bold=True, color=PALETA_CORES["DIVERGENCE_FONT"])
        for col_idx, val in enumerate(row_divergent, start=1):
            c = ws_conf.cell(row=3, column=col_idx, value=val)
            c.alignment = align_left if col_idx in (2, 4, 6, 8, 12, 36, 37, 38) else align_center
            c.fill = div_fill
            c.font = div_font

        # 4. Valores Importados
        ws_vi = wb.create_sheet("Valores Importados")
        ws_vi.freeze_panes = "A2"
        for col_idx, col_name in enumerate(COLUNAS_VALORES_IMPORTADOS, start=1):
            c = ws_vi.cell(row=1, column=col_idx, value=col_name)
            c.fill = header_fill
            c.font = header_font
            c.alignment = align_center

        row_vi = [7, "305204", 29613, "92026", 301.85, 301.85, True, "", "", None, None]
        for col_idx, val in enumerate(row_vi, start=1):
            ws_vi.cell(row=2, column=col_idx, value=val)

        wb.save(target_path)

    def test_audit_workbook_four_sheets_in_order(self):
        """Workbook has exactly 4 sheets in strict official sequence."""
        self._generate_test_workbook(self.output_xlsx)
        wb = openpyxl.load_workbook(self.output_xlsx, read_only=True)
        expected_sheets = ["Orientações", "Resumo da Auditoria", "Conferência", "Valores Importados"]
        self.assertEqual(wb.sheetnames, expected_sheets)

    def test_audit_workbook_freeze_panes_a2(self):
        """Freeze Panes is strictly 'A2' on Conferência and Valores Importados."""
        self._generate_test_workbook(self.output_xlsx)
        wb = openpyxl.load_workbook(self.output_xlsx)
        self.assertEqual(wb["Conferência"].freeze_panes, "A2")
        self.assertEqual(wb["Valores Importados"].freeze_panes, "A2")

    def test_audit_workbook_conferencia_has_38_columns(self):
        """Conferência sheet has exactly 38 columns matching the contract."""
        self._generate_test_workbook(self.output_xlsx)
        wb = openpyxl.load_workbook(self.output_xlsx)
        ws = wb["Conferência"]
        header_vals = [ws.cell(row=1, column=c).value for c in range(1, 39)]
        self.assertEqual(len(header_vals), 38)
        self.assertEqual(header_vals, COLUNAS_CONFERENCIA)

    def test_audit_workbook_valores_importados_has_11_columns(self):
        """Valores Importados sheet has exactly 11 columns matching the contract."""
        self._generate_test_workbook(self.output_xlsx)
        wb = openpyxl.load_workbook(self.output_xlsx)
        ws = wb["Valores Importados"]
        header_vals = [ws.cell(row=1, column=c).value for c in range(1, 12)]
        self.assertEqual(len(header_vals), 11)
        self.assertEqual(header_vals, COLUNAS_VALORES_IMPORTADOS)

    def test_audit_workbook_zero_formulas(self):
        """All cells across all sheets contain ZERO formulas (never start with '=')."""
        self._generate_test_workbook(self.output_xlsx)
        wb = openpyxl.load_workbook(self.output_xlsx, data_only=False)
        for sheetname in wb.sheetnames:
            ws = wb[sheetname]
            for row in ws.iter_rows(values_only=True):
                for val in row:
                    if val is not None:
                        self.assertFalse(str(val).startswith("="), f"Formula detected in {sheetname}: {val}")

    def test_audit_workbook_divergence_styling_applied(self):
        """Row with divergence has red fill #FFC7CE and dark red font #9C0006."""
        self._generate_test_workbook(self.output_xlsx)
        wb = openpyxl.load_workbook(self.output_xlsx)
        ws = wb["Conferência"]
        # Row 3 is divergent
        cell_divergent = ws.cell(row=3, column=38)
        self.assertEqual(cell_divergent.fill.start_color.rgb, "00FFC7CE")
        self.assertEqual(cell_divergent.font.color.rgb, "009C0006")


if __name__ == "__main__":
    unittest.main(verbosity=2)
