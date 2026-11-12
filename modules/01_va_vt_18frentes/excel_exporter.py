"""
Exportador Excel Corporativo Cencosud — Valores Pré-Folha VT (Eventos 5750, 31935 e 6292).
Estrutura oficial de 4 abas:
1. Orientações
2. Resumo da Auditoria
3. Conferência Geral
4. Divergências e Inconsistências
"""

import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Set

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
if str(SCRIPT_DIR.parent) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR.parent))

try:
    from config import (
        COLOR_ALERT_BG,
        COLOR_BLUE,
        COLOR_BORDER,
        COLOR_LIGHT_BLUE,
        COLOR_NAVY,
        COLOR_OK_BG,
        COLOR_ZEBRA,
        COMPETENCIA,
        OUTPUT_FILE,
    )
except ImportError:
    from scripts.config import (
        COLOR_ALERT_BG,
        COLOR_BLUE,
        COLOR_BORDER,
        COLOR_LIGHT_BLUE,
        COLOR_NAVY,
        COLOR_OK_BG,
        COLOR_ZEBRA,
        COMPETENCIA,
        OUTPUT_FILE,
    )

CURRENCY_COLS: Set[str] = {
    "118 - Credito Vale Transporte",
    "147 - Credito Vale Transporte Dev Operadora FL 40",
    "148 - Credito Vale Transporte Dev Operadora FL 41",
    "159 - Deslocamento Atividades Externa",
    "490 - Credito Vale Transporte",
    "5731 - Saldo Vale Transporte Rescisao",
    "5750 - Adiantamento VT Atraves Alcadas",
    "6292 - Desconto Cartao VTE",
    "9290 - Dev Vale Transporte",
    "9421 - Reembolso VT Sem Dedução - Folha 40",
    "9422 - Reembolso VT Sem Dedução - Folha 41",
    "9423 - Reembolso VT Sem Dedução - Folha 42",
    "9424 - Credito Vale Transporte - Folha 41",
    "9425 - Credito Vale Transporte - Folha 42",
    "9426 - Dev Vale Transporte - Folha 41",
    "9427 - Dev Vale Transporte - Folha 42",
    "9428 - Deslocamento Atividades Externa - Folha 41",
    "9429 - Deslocamento Atividades Externa - Folha 42",
    "31935 - Valor Vale Transporte Alternativo Mes Seguinte",
    "Soma Eventos Pagamento",
    "Diferença (Soma x 5750)",
    "Importado 31935",
    "Importado 6292",
}

DATE_COLS: Set[str] = {
    "Data Admissão",
    "Data Rescisão",
}

LEFT_COLS: Set[str] = {
    "Nome da Empresa",
    "Estabelecimento",
    "Centro Custo",
    "Nome do Colaborador",
    "Cargo",
    "Situação",
    "Tipo Situação",
    "Status Auditoria",
    "Obs.",
}


class ExcelExporterPrefolha:
    """Gera o relatório executivo em 4 abas da Conferência de Valores Pré-Folha VT."""

    def __init__(self, output_path: Path = OUTPUT_FILE) -> None:
        self.output_path = Path(output_path)
        self.font_title = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
        self.font_section = Font(name="Calibri", size=11, bold=True, color=COLOR_NAVY)
        self.font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        self.font_header_audit = Font(name="Calibri", size=10, bold=True, color="9C0006")
        self.font_subhdr = Font(name="Calibri", size=10, bold=True, color=COLOR_NAVY)
        self.font_data = Font(name="Calibri", size=10, bold=False, color="111111")
        self.font_bold = Font(name="Calibri", size=10, bold=True, color="111111")
        self.font_alert = Font(name="Calibri", size=10, bold=True, color="9C0006")
        self.font_ok = Font(name="Calibri", size=10, bold=True, color="274E13")

        self.fill_navy = PatternFill(start_color=COLOR_NAVY, end_color=COLOR_NAVY, fill_type="solid")
        self.fill_yellow_audit = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        self.fill_blue = PatternFill(start_color=COLOR_BLUE, end_color=COLOR_BLUE, fill_type="solid")
        self.fill_light = PatternFill(
            start_color=COLOR_LIGHT_BLUE, end_color=COLOR_LIGHT_BLUE, fill_type="solid"
        )
        self.fill_zebra = PatternFill(
            start_color=COLOR_ZEBRA, end_color=COLOR_ZEBRA, fill_type="solid"
        )
        self.fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        self.fill_alert = PatternFill(
            start_color=COLOR_ALERT_BG, end_color=COLOR_ALERT_BG, fill_type="solid"
        )
        self.fill_ok = PatternFill(start_color=COLOR_OK_BG, end_color=COLOR_OK_BG, fill_type="solid")

        thin = Side(border_style="thin", color=COLOR_BORDER)
        self.border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)

        self.align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
        self.align_left = Alignment(horizontal="left", vertical="center", wrap_text=False)
        self.align_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def export(self, result: Dict[str, Any]) -> Path:
        """Constrói e salva o arquivo Excel completo de 4 abas."""
        wb = Workbook()
        try:
            ws_ori = wb.active
            ws_ori.title = "Orientações"
            ws_res = wb.create_sheet(title="Resumo da Auditoria")
            ws_ger = wb.create_sheet(title="Conferência Geral")
            ws_div = wb.create_sheet(title="Divergências e Inconsistências")

            self._build_orientacoes(ws_ori)
            self._build_resumo(ws_res, result["kpis"])
            self._build_data_sheet(ws_ger, result["df_geral"], is_divergencia=False)
            self._build_data_sheet(ws_div, result["df_divergencias"], is_divergencia=True)

            self.output_path.parent.mkdir(parents=True, exist_ok=True)
            wb.save(self.output_path)
            return self.output_path
        finally:
            wb.close()

    def _build_orientacoes(self, ws: Any) -> None:
        ws.views.sheetView[0].showGridLines = False
        ws.merge_cells("B2:E2")
        ws["B2"] = f"GUIA TÉCNICO E REGRAS DE NEGÓCIO — VALORES PRÉ-FOLHA VT ({COMPETENCIA})"
        ws["B2"].font = self.font_title
        ws["B2"].fill = self.fill_navy
        ws["B2"].alignment = self.align_center
        ws.row_dimensions[2].height = 34

        rows_info = [
            (
                "1. Objetivo do Relatório",
                "Auditar a integridade do espelho de Valores Pré-Folha de Vale Transporte (FOPA) na competência "
                "09/2026, assegurando que a soma dos 10 eventos de pagamento/reembolso corresponda exatamente ao "
                "evento de saída '5750 - Adiantamento VT Através Alçadas' e que os lotes de importação '31935' e "
                "'6292' estejam conciliados.",
            ),
            (
                "2. Regra de Equilíbrio do Evento 5750 (Alçadas)",
                "Para cada colaborador na pré-folha (1.183 colaboradores auditados), o sistema soma as rubricas:\n"
                "• 147 (Crédito VT Dev Operadora FL 40) + 148 (Crédito VT Dev Operadora FL 41)\n"
                "• 159 (Deslocamento Atividade Externa) + 490 (Crédito Vale Transporte)\n"
                "• 9421, 9422 e 9423 (Reembolso VT Sem Dedução - Folhas 40, 41 e 42)\n"
                "• 9428 e 9429 (Deslocamento Atividades Externa - Folhas 41 e 42)\n"
                "• 31935 (Valor Vale Transporte Alternativo Mês Seguinte)\n"
                "O resultado deve ser 100% idêntico ao evento '5750 - Adiantamento VT Atraves Alcadas'. Qualquer "
                "diferença sinaliza falha de parametrização junto à Quality.",
            ),
            (
                "3. Validação de Importação — Evento 31935",
                "Cruza os 387 lançamentos do lote de importação de VT Alternativo Mês Seguinte contra o espelho "
                "da FOPA Mensal. Colaboradores desligados cujo valor esteja transitando em rescisão são "
                "automaticamente segregados na aba 'Divergências e Inconsistências'.",
            ),
            (
                "4. Validação de Importação — Evento 6292",
                "Cruza os 30 lançamentos do lote de importação de Desconto Cartão VTE (Evento 6292) contra a "
                "FOPA Mensal, apontando colaboradores demitidos antes do processamento da folha.",
            ),
            (
                "5. Estrutura das Abas",
                "• Orientações: Documentação executiva das regras de validação Pré-Folha.\n"
                "• Resumo da Auditoria: KPIs consolidados, composição analítica por evento FOPA e por situação.\n"
                "• Conferência Geral: Espelho completo dos 1.183 colaboradores ativos/afastados na FOPA.\n"
                "• Divergências e Inconsistências: Painel focado exclusivamente nos 4 casos críticos (1 falha "
                "de contrapartida 5750 na FOPA + 2 rescisões do lote 31935 + 1 rescisão do lote 6292).",
            ),
        ]

        r_idx = 4
        for topic, desc in rows_info:
            ws.cell(row=r_idx, column=2, value=topic).font = self.font_subhdr
            ws.cell(row=r_idx, column=2).fill = self.fill_light
            ws.cell(row=r_idx, column=2).alignment = self.align_wrap
            ws.cell(row=r_idx, column=2).border = self.border_cell

            ws.merge_cells(start_row=r_idx, start_column=3, end_row=r_idx, end_column=5)
            for c_i in range(3, 6):
                cell = ws.cell(row=r_idx, column=c_i)
                cell.border = self.border_cell
                cell.fill = self.fill_white
            ws.cell(row=r_idx, column=3, value=desc).font = self.font_data
            ws.cell(row=r_idx, column=3).alignment = self.align_wrap
            ws.row_dimensions[r_idx].height = 64 if "2." in topic or "5." in topic else 45
            r_idx += 2

        ws.column_dimensions["A"].width = 4
        ws.column_dimensions["B"].width = 36
        ws.column_dimensions["C"].width = 45
        ws.column_dimensions["D"].width = 30
        ws.column_dimensions["E"].width = 30

    def _build_resumo(self, ws: Any, kpis: Dict[str, Any]) -> None:
        ws.views.sheetView[0].showGridLines = False
        ws.merge_cells("B2:F2")
        ws["B2"] = f"RESUMO EXECUTIVO DA AUDITORIA — VALORES PRÉ-FOLHA VT ({COMPETENCIA})"
        ws["B2"].font = self.font_title
        ws["B2"].fill = self.fill_navy
        ws["B2"].alignment = self.align_center
        ws.row_dimensions[2].height = 34

        # Bloco 1: Indicadores Gerais
        ws["B4"] = "1. INDICADORES GERAIS DE CONFORMIDADE (FOPA x IMPORTAÇÕES)"
        ws["B4"].font = self.font_section

        headers_kpi = ["Indicador Auditado", "Resultado / Volume", "Status / Referência"]
        for idx, h_text in enumerate(headers_kpi, start=2):
            c = ws.cell(row=5, column=idx, value=h_text)
            c.font = self.font_header
            c.fill = self.fill_blue
            c.alignment = self.align_center
            c.border = self.border_cell
        ws.row_dimensions[5].height = 24

        kpi_rows = [
            ("Total de Colaboradores na FOPA Pré-Folha", kpis["total_fopa"], "Base FOPA 22/09/2026", "int"),
            ("Colaboradores Conciliados na FOPA (Soma = 5750)", kpis["conciliados_fopa"], f"{kpis['taxa_conformidade_fopa']:.2f}% de Conformidade", "int"),
            ("Divergências de Contrapartida 5750 na FOPA", kpis["divergentes_fopa"], "Sinalizar à Quality (Evt 147 sem 5750)", "int"),
            ("Soma Total Eventos de Pagamento / Reembolso (R$)", kpis["soma_pgtos"], "10 Rubricas de Crédito Pré-Folha", "curr"),
            ("Total Evento 5750 - Adiantamento VT Alçadas (R$)", kpis["soma_5750"], "Rubrica de Saída Pré-Folha", "curr"),
            ("Diferença Líquida (Soma Pagamentos - Evento 5750)", kpis["diff_5750"], "Divergência Matrícula c/ Evt 147", "curr"),
            ("Lote Importação 31935 — Qtd Total / Conciliados", f"{kpis['ok_31935']} de {kpis['qtd_31935_imp']}", f"{kpis['div_31935']} em Rescisão", "str"),
            ("Lote Importação 31935 — Valor Importado vs FOPA", kpis["tot_31935_fopa"], f"Lote Original: R$ {kpis['tot_31935_imp']:,.2f}", "curr"),
            ("Lote Importação 6292 — Qtd Total / Conciliados", f"{kpis['ok_6292']} de {kpis['qtd_6292_imp']}", f"{kpis['div_6292']} Demitido (29/08)", "str"),
            ("Lote Importação 6292 — Valor Importado vs FOPA", kpis["tot_6292_fopa"], f"Lote Original: R$ {kpis['tot_6292_imp']:,.2f}", "curr"),
            ("Total de Inconsistências Consolidadas p/ Tratativa", kpis["total_divergencias_consolidadas"], "Ver Aba Divergências", "int"),
        ]

        curr_row = 6
        for label, val, ref, vtype in kpi_rows:
            fill_r = self.fill_zebra if curr_row % 2 == 0 else self.fill_white
            c_lbl = ws.cell(row=curr_row, column=2, value=label)
            c_val = ws.cell(row=curr_row, column=3, value=val)
            c_ref = ws.cell(row=curr_row, column=4, value=ref)

            c_lbl.font = self.font_data
            c_lbl.alignment = self.align_left
            c_val.font = self.font_bold
            c_val.alignment = self.align_center
            c_ref.font = self.font_data
            c_ref.alignment = self.align_center

            if vtype == "curr":
                c_val.number_format = "R$ #,##0.00"
            elif vtype == "int":
                c_val.number_format = "#,##0"

            for cell in (c_lbl, c_val, c_ref):
                cell.fill = fill_r
                cell.border = self.border_cell
            ws.row_dimensions[curr_row].height = 20
            curr_row += 1

        # Bloco 2: Composição Analítica por Evento FOPA
        curr_row += 2
        ws.cell(row=curr_row, column=2, value="2. COMPOSIÇÃO FINANCEIRA DOS EVENTOS QUE INTEGRAM O 5750").font = self.font_section
        curr_row += 1

        df_ev = kpis["df_resumo_eventos"]
        ev_headers = list(df_ev.columns)
        for idx, h_text in enumerate(ev_headers, start=2):
            c = ws.cell(row=curr_row, column=idx, value=h_text)
            c.font = self.font_header
            c.fill = self.fill_navy
            c.alignment = self.align_center
            c.border = self.border_cell
        ws.row_dimensions[curr_row].height = 24
        curr_row += 1

        for _, r_ev in df_ev.iterrows():
            is_5750 = str(r_ev["Evento FOPA"]).startswith("5750")
            fill_r = self.fill_light if is_5750 else (self.fill_zebra if curr_row % 2 == 0 else self.fill_white)
            font_r = self.font_bold if is_5750 else self.font_data

            c1 = ws.cell(row=curr_row, column=2, value=r_ev["Evento FOPA"])
            c2 = ws.cell(row=curr_row, column=3, value=r_ev["Tipo"])
            c3 = ws.cell(row=curr_row, column=4, value=int(r_ev["Qtd Colaboradores (> R$ 0)"]))
            c4 = ws.cell(row=curr_row, column=5, value=float(r_ev["Valor Total FOPA (R$)"]))
            c5 = ws.cell(row=curr_row, column=6, value=float(r_ev["% sobre Soma Pagamentos"]) / 100.0)

            c1.alignment = self.align_left
            c2.alignment = self.align_left
            c3.alignment = self.align_center
            c3.number_format = "#,##0"
            c4.alignment = self.align_center
            c4.number_format = "R$ #,##0.00"
            c5.alignment = self.align_center
            c5.number_format = "0.00%"

            for cell in (c1, c2, c3, c4, c5):
                cell.font = font_r
                cell.fill = fill_r
                cell.border = self.border_cell
            ws.row_dimensions[curr_row].height = 20
            curr_row += 1

        # Bloco 3: Distribuição por Situação na FOPA
        curr_row += 2
        ws.cell(row=curr_row, column=2, value="3. DISTRIBUIÇÃO POR SITUAÇÃO FUNCIONAL NA PRÉ-FOLHA").font = self.font_section
        curr_row += 1

        sit_headers = [
            "Tipo Situação",
            "Situação Detalhada",
            "Qtd Colaboradores",
            "Soma Pagamentos (R$)",
            "Total Evento 5750 (R$)",
        ]
        for idx, h_text in enumerate(sit_headers, start=2):
            c = ws.cell(row=curr_row, column=idx, value=h_text)
            c.font = self.font_header
            c.fill = self.fill_blue
            c.alignment = self.align_center
            c.border = self.border_cell
        ws.row_dimensions[curr_row].height = 24
        curr_row += 1

        for _, r_sit in kpis["df_resumo_situacao"].iterrows():
            fill_r = self.fill_zebra if curr_row % 2 == 0 else self.fill_white
            c1 = ws.cell(row=curr_row, column=2, value=r_sit["Tipo Situação"])
            c2 = ws.cell(row=curr_row, column=3, value=r_sit["Situação"])
            c3 = ws.cell(row=curr_row, column=4, value=int(r_sit["Colaboradores"]))
            c4 = ws.cell(row=curr_row, column=5, value=float(r_sit["Soma_Pagamentos"]))
            c5 = ws.cell(row=curr_row, column=6, value=float(r_sit["Total_5750"]))

            c1.alignment = self.align_left
            c2.alignment = self.align_left
            c3.alignment = self.align_center
            c3.number_format = "#,##0"
            c4.alignment = self.align_center
            c4.number_format = "R$ #,##0.00"
            c5.alignment = self.align_center
            c5.number_format = "R$ #,##0.00"

            for cell in (c1, c2, c3, c4, c5):
                cell.font = self.font_data
                cell.fill = fill_r
                cell.border = self.border_cell
            ws.row_dimensions[curr_row].height = 20
            curr_row += 1

        ws.column_dimensions["A"].width = 4
        ws.column_dimensions["B"].width = 50
        ws.column_dimensions["C"].width = 42
        ws.column_dimensions["D"].width = 28
        ws.column_dimensions["E"].width = 26
        ws.column_dimensions["F"].width = 26

    def _build_data_sheet(self, ws: Any, df: pd.DataFrame, is_divergencia: bool) -> None:
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = "A2"

        cols = list(df.columns)
        cols_audit = {
            "Soma Eventos Pagamento",
            "Diferença (Soma x 5750)",
            "Evts Pgtos x Evt 5750",
            "Importado 31935",
            "Importado 6292",
            "Status Auditoria",
            "Obs.",
        }
        for c_idx, col_name in enumerate(cols, start=1):
            cell = ws.cell(row=1, column=c_idx, value=col_name)
            if col_name in cols_audit:
                cell.font = self.font_header_audit
                cell.fill = self.fill_yellow_audit
            else:
                cell.font = self.font_header
                cell.fill = self.fill_navy
            cell.alignment = self.align_center
            cell.border = self.border_cell
        ws.row_dimensions[1].height = 28

        max_lens = {c_idx: len(str(col_name)) for c_idx, col_name in enumerate(cols, start=1)}

        for r_idx, row_vals in enumerate(df.itertuples(index=False, name=None), start=2):
            row_fill = (
                self.fill_alert
                if is_divergencia
                else (self.fill_zebra if r_idx % 2 == 0 else self.fill_white)
            )
            ws.row_dimensions[r_idx].height = 20

            for c_idx, (col_name, val) in enumerate(zip(cols, row_vals), start=1):
                if pd.isna(val) and not isinstance(val, (date, datetime)):
                    val = None
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.font = self.font_data
                cell.fill = row_fill
                cell.border = self.border_cell

                if col_name in CURRENCY_COLS:
                    cell.number_format = "R$ #,##0.00"
                    cell.alignment = self.align_center
                    s_len = 14
                elif col_name in DATE_COLS:
                    if isinstance(val, (date, datetime)):
                        cell.number_format = "dd/mm/yyyy"
                    cell.alignment = self.align_center
                    s_len = 12
                elif col_name in LEFT_COLS:
                    cell.alignment = self.align_left
                    if col_name in ("Status Auditoria", "Obs.") and val not in ("Conciliado (Ok)", "Ok"):
                        cell.font = self.font_alert
                        cell.fill = self.fill_alert
                    elif col_name == "Status Auditoria" and val == "Conciliado (Ok)":
                        cell.font = self.font_ok
                    s_len = len(str(val or ""))
                else:
                    cell.alignment = self.align_center
                    s_len = len(str(val or ""))

                if s_len > max_lens[c_idx]:
                    max_lens[c_idx] = s_len

        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(2, len(df) + 1)}"

        for c_idx, max_l in max_lens.items():
            col_letter = get_column_letter(c_idx)
            ws.column_dimensions[col_letter].width = min(max(max_l + 5, 14), 52)
