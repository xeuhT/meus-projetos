"""
Módulo de Exportação Excel para a Conferência de Seguro de Vida (Cencosud Brasil).
Gera:
1. Arquivos individuais de carga por bandeira no padrão operacional exato:
   - Gbarbosa e Jaguaquara, Prezunic: cabeçalho 'CPF/CNPJ'.
   - Bretas Lojas, Bretas Postos, Giga e Mercantil, Perini: cabeçalho 'CPF'.
2. Relatório Master de Conferência e Auditoria com estilização visual institucional rigorosa:
   - Colunas Cadastrais: Azul Corporativo (#1F497D) com fonte branca negrito.
   - Colunas de Conferência/Validação: Amarelo (#FFFF00) com fonte vermelha negrito (#9C0006).
   - Abas: Resumo Executivo, Base Consolidada, Lista de Exclusões, Divergências Folha,
           Conciliação Fatura Seguradora e Seguro Complementar.
   - Fechamento estrito de conexões/workbooks para evitar vazamento de memória.
"""

from pathlib import Path
from typing import Dict, Any, List, Optional, Callable
import datetime
import json
import zipfile
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config import (
    OUTPUT_DIR,
    AUDIT_REPORT_DIR,
    CADASTRO_DIRETORES_JSON,
    COLUNAS_CARGA_PADRAO_CPF,
    COLUNAS_CARGA_PADRAO_CPF_CNPJ,
    COLUNAS_CARGA_COMPLEMENTAR,
    COR_CABECALHO_CADASTRAL,
    COR_FONTE_CADASTRAL,
    COR_CABECALHO_CONFERENCIA,
    COR_FONTE_CONFERENCIA,
    COR_ZEBRA_PAR,
    COR_BORDA,
    COR_KPI_HEADER,
    COR_KPI_BOX
)


class ExcelExporter:
    """Classe responsável pela formatação e geração de todos os arquivos Excel."""

    def __init__(
        self,
        competencia: str = "102026",
        output_dir: Optional[Path] = None,
        audit_dir: Optional[Path] = None,
        zip_dir: Optional[Path] = None,
        log_cb: Optional[Callable[[str], None]] = None,
        progress_cb: Optional[Callable[[float, str], None]] = None,
        prompt_lock_cb: Optional[Callable[[str], str]] = None,
    ):
        self.competencia = competencia
        self.output_dir = Path(output_dir) if output_dir else OUTPUT_DIR
        self.audit_dir = Path(audit_dir) if audit_dir else AUDIT_REPORT_DIR
        self.zip_dir = Path(zip_dir) if zip_dir else self.output_dir.parent
        self.log_cb = log_cb
        self.progress_cb = progress_cb
        self.prompt_lock_cb = prompt_lock_cb

        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.audit_dir.mkdir(parents=True, exist_ok=True)
        self.zip_dir.mkdir(parents=True, exist_ok=True)

        self.thin_border = Border(
            left=Side(style="thin", color=COR_BORDA),
            right=Side(style="thin", color=COR_BORDA),
            top=Side(style="thin", color=COR_BORDA),
            bottom=Side(style="thin", color=COR_BORDA)
        )
        self.zebra_fill = PatternFill(start_color=COR_ZEBRA_PAR, end_color=COR_ZEBRA_PAR, fill_type="solid")
        self.white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        self.diretores = []
        self.mats_diretores = set()
        if CADASTRO_DIRETORES_JSON.exists():
            try:
                with open(CADASTRO_DIRETORES_JSON, "r", encoding="utf-8") as f:
                    self.diretores = json.load(f)
                    self.mats_diretores = {int(d["matricula"]) for d in self.diretores}
            except Exception:
                pass

    def _log(self, msg: str) -> None:
        print(msg)
        if self.log_cb:
            self.log_cb(msg)

    def _progress(self, pct: float, msg: str) -> None:
        if self.progress_cb:
            self.progress_cb(pct, msg)

    def _save_workbook_safe(self, wb: openpyxl.Workbook, target_path: Path) -> Path:
        """
        Salva o workbook garantindo fechamento (wb.close()) em bloco try/finally
        e tratando bloqueios de arquivo aberto no Excel via FileLockModal se disponível.
        """
        actual_path = target_path
        try:
            while True:
                try:
                    wb.save(actual_path)
                    return actual_path
                except PermissionError:
                    if not self.prompt_lock_cb:
                        raise
                    decision = self.prompt_lock_cb(str(actual_path))
                    if decision == "retry":
                        continue
                    elif decision == "copy":
                        stamp = datetime.datetime.now().strftime("%H%M%S")
                        actual_path = actual_path.with_name(f"{actual_path.stem}_Copia_{stamp}{actual_path.suffix}")
                        self._log(f"[!] Arquivo original em uso; salvando como cópia: {actual_path.name}")
                        continue
                    else:
                        raise RuntimeError(f"Gravação cancelada pelo usuário (arquivo aberto no Excel: {actual_path.name}).")
        finally:
            try:
                wb.close()
            except Exception:
                pass

    def export_all(self, results: Dict[str, Any]) -> Dict[str, Path]:
        """Orquestra a geração de todos os arquivos de carga e do relatório de auditoria."""
        self._log("[*] Iniciando exportação dos arquivos Excel...")
        self._progress(0.45, "Gerando cargas operacionais por bandeira...")

        arquivos_gerados = {}

        # 1. Gera os Arquivos de Carga do Seguro Normal por Bandeira
        cargas_normal = self._export_cargas_normal(results["seguro_normal"])
        arquivos_gerados.update(cargas_normal)

        # 2. Gera os Arquivos de Carga do Seguro Complementar (12x e 24x)
        self._progress(0.68, "Gerando cargas do Seguro Complementar (12x e 24x)...")
        cargas_comp = self._export_cargas_complementar(results["comp_12x"], results["comp_24x"])
        arquivos_gerados.update(cargas_comp)

        # 3. Gera o Relatório Master de Conferência e Auditoria Trilateral
        self._progress(0.75, "Formatando Relatório Master de Conferência (6 abas, padrão institucional)...")
        audit_file = self._export_audit_report(results)
        arquivos_gerados["relatorio_auditoria"] = audit_file

        # 4. Compacta o pacote ZIP com todas as cargas geradas na competência
        self._progress(0.95, "Compactando pacote ZIP das cargas de Seguro de Vida...")
        zip_file = self._export_zip_package(list(cargas_normal.values()) + list(cargas_comp.values()))
        if zip_file:
            arquivos_gerados["pacote_zip"] = zip_file

        self._log(f"[+] Exportação concluída com sucesso ({len(cargas_normal) + len(cargas_comp)} cargas + Relatório Master + ZIP).")
        return arquivos_gerados

    def _export_zip_package(self, load_files: List[Path]) -> Optional[Path]:
        """Compacta os 8 arquivos operacionais de carga em um único .zip pronto para envio."""
        zip_path = self.zip_dir / "Carga Seguro de Vida Gerada.zip"
        try:
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for lf in load_files:
                    if lf.exists():
                        zf.write(lf, arcname=lf.name)
            self._log(f"    -> Pacote compactado atualizado: {zip_path.name}")
            return zip_path
        except PermissionError:
            alt_zip = self.zip_dir / f"Carga Seguro de Vida Gerada_{self.competencia}.zip"
            try:
                with zipfile.ZipFile(alt_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                    for lf in load_files:
                        if lf.exists():
                            zf.write(lf, arcname=lf.name)
                return alt_zip
            except Exception:
                return None

    def _export_cargas_normal(self, df_normal: pd.DataFrame) -> Dict[str, Path]:
        """Gera os arquivos operacionais de carga para o Seguro Normal."""
        gerados = {}

        # Mapeamento do nome do arquivo para cada agrupamento e parametrização do cabeçalho CPF vs CPF/CNPJ
        agrupamentos = {
            "Bretas Lojas": (df_normal[df_normal["arquivo_destino"] == "Bretas Lojas"], COLUNAS_CARGA_PADRAO_CPF),
            "Bretas Postos": (df_normal[df_normal["arquivo_destino"] == "Bretas Postos"], COLUNAS_CARGA_PADRAO_CPF),
            "Gbarbosa e Jaguaquara": (df_normal[df_normal["arquivo_destino"] == "Gbarbosa e Jaguaquara"], COLUNAS_CARGA_PADRAO_CPF_CNPJ),
            "Giga e Mercantil": (df_normal[df_normal["arquivo_destino"] == "Giga e Mercantil"], COLUNAS_CARGA_PADRAO_CPF),
            "Perini": (df_normal[df_normal["arquivo_destino"] == "Perini"], COLUNAS_CARGA_PADRAO_CPF),
            "Prezunic": (df_normal[df_normal["arquivo_destino"] == "Prezunic"], COLUNAS_CARGA_PADRAO_CPF_CNPJ)
        }

        total_agrup = len(agrupamentos)
        for idx_ag, (nome_base, (df_subset, headers)) in enumerate(agrupamentos.items(), start=1):
            file_name = f"{nome_base} {self.competencia}.xlsx"
            file_path = self.output_dir / file_name
            self._progress(0.45 + (0.20 * idx_ag / total_agrup), f"Gerando carga: {file_name}...")

            # Ordenação inteligente:
            # Em Gbarbosa e Jaguaquara, todos os Diretores extraídos da aba 'Com Gestor' ficam fixados
            # no topo absoluto (em ordem alfabética) para visibilidade imediata pela equipe de Benefícios/RH,
            # e os demais colaboradores ficam ordenados alfabeticamente logo em seguida.
            if nome_base == "Gbarbosa e Jaguaquara":
                if "is_diretor" in df_subset.columns:
                    mask_dir = df_subset["is_diretor"].fillna(False).astype(bool)
                else:
                    mask_dir = df_subset["matricula"].isin(self.mats_diretores)
                df_dir = df_subset[mask_dir].sort_values(by="colaborador")
                df_outros = df_subset[~mask_dir].sort_values(by="colaborador")
                df_subset = pd.concat([df_dir, df_outros], ignore_index=True)
            else:
                df_subset = df_subset.sort_values(by="colaborador")

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Planilha1"

            ws.append(headers)

            header_fill = PatternFill(start_color=COR_CABECALHO_CADASTRAL, end_color=COR_CABECALHO_CADASTRAL, fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color=COR_FONTE_CADASTRAL)

            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = self.thin_border

            # Preenchimento das Linhas
            for row_idx, r in enumerate(df_subset.itertuples(index=False), start=2):
                mat = int(r.matricula) if pd.notna(r.matricula) else ""
                nome = str(r.colaborador).strip().upper()
                dt_nasc = r.data_nascimento.strftime("%d/%m/%Y") if pd.notna(r.data_nascimento) else ""
                cpf = str(r.cpf).strip()
                plano = str(r.plano_atribuido).strip()

                ws.append([mat, nome, dt_nasc, cpf, plano])

                fill = self.white_fill if row_idx % 2 == 0 else self.zebra_fill
                ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal="center")
                ws.cell(row=row_idx, column=2).alignment = Alignment(horizontal="left")
                ws.cell(row=row_idx, column=3).alignment = Alignment(horizontal="center")
                ws.cell(row=row_idx, column=4).alignment = Alignment(horizontal="center")
                ws.cell(row=row_idx, column=5).alignment = Alignment(horizontal="center")

                for c_idx in range(1, 6):
                    c = ws.cell(row=row_idx, column=c_idx)
                    c.fill = fill
                    c.border = self.thin_border
                    c.font = Font(name="Calibri", size=10)

            self._auto_adjust_columns(ws)
            saved_path = self._save_workbook_safe(wb, file_path)

            self._log(f"    -> Carga gerada: {saved_path.name} ({len(df_subset)} registros) [Cabeçalho: {headers[3]}]")
            gerados[nome_base] = saved_path

        return gerados

    def _export_cargas_complementar(self, df_12x: pd.DataFrame, df_24x: pd.DataFrame) -> Dict[str, Path]:
        """Gera os arquivos operacionais do Seguro Complementar (12x e 24x)."""
        gerados = {}

        file_12x = self.output_dir / f"Seguro Complementar 12x_{self.competencia}.xlsx"
        gerados["Comp 12x"] = self._write_complementar_file(file_12x, df_12x, sheet_name="12x")

        file_24x = self.output_dir / f"Seguro Complementar 24x_{self.competencia}.xlsx"
        gerados["Comp 24x"] = self._write_complementar_file(file_24x, df_24x, sheet_name="24x")

        return gerados

    def _write_complementar_file(self, path: Path, df: pd.DataFrame, sheet_name: str) -> Path:
        """Grava arquivo de Seguro Complementar com colunas e formatações exatas."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name

        headers = COLUNAS_CARGA_COMPLEMENTAR
        ws.append(headers)

        header_fill = PatternFill(start_color=COR_CABECALHO_CADASTRAL, end_color=COR_CABECALHO_CADASTRAL, fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color=COR_FONTE_CADASTRAL)

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = self.thin_border

        for row_idx, r in enumerate(df.itertuples(index=False), start=2):
            mat = int(r.matricula) if pd.notna(r.matricula) else ""
            nome = str(r.nome).strip().upper()
            dt_nasc = r.nascimento.strftime("%d/%m/%Y") if pd.notna(r.nascimento) else ""
            cpf = str(r.cpf).strip()
            salario = float(r.salario) if pd.notna(r.salario) else ""
            plano = str(r.plano).strip()

            ws.append([mat, nome, dt_nasc, cpf, salario, plano])

            fill = self.white_fill if row_idx % 2 == 0 else self.zebra_fill
            ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=2).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=3).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=4).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=5).alignment = Alignment(horizontal="right")
            ws.cell(row=row_idx, column=6).alignment = Alignment(horizontal="center")

            if salario != "":
                ws.cell(row=row_idx, column=5).number_format = "#,##0.00"

            for c_idx in range(1, 7):
                c = ws.cell(row=row_idx, column=c_idx)
                c.fill = fill
                c.border = self.thin_border
                c.font = Font(name="Calibri", size=10)

        self._auto_adjust_columns(ws)
        saved_path = self._save_workbook_safe(wb, path)
        self._log(f"    -> Carga Complementar gerada: {saved_path.name} ({len(df)} registros)")
        return saved_path

    def _export_audit_report(self, results: Dict[str, Any]) -> Path:
        """
        Gera o Relatório Master de Conferência e Auditoria Técnica.
        Aplica estritamente o padrão visual institucional:
        - Colunas Cadastrais: Azul Corporativo (#1F497D) com fonte branca negrito.
        - Colunas de Conferência / Validação: Amarelo (#FFFF00) com fonte vermelha negrito (#9C0006).
        """
        report_path = self.audit_dir / f"Relatorio_Conferencia_Seguro_{self.competencia}.xlsx"
        wb = openpyxl.Workbook()
        try:
            # 1. Aba Resumo Executivo
            ws_resumo = wb.active
            ws_resumo.title = "Resumo Executivo"
            self._build_executive_summary_sheet(ws_resumo, results["kpis"])

            # 2. Aba Base Consolidada (Com formatação bicolor institucional rigorosa)
            self._progress(0.80, "Preenchendo aba Base Consolidada (15.900+ vidas)...")
            ws_base = wb.create_sheet(title="Base Consolidada")
            self._build_consolidated_base_sheet(ws_base, results["base_consolidada"])

            # 3. Aba Lista de Exclusões (Com cabeçalhos bicolor e justificativas)
            self._progress(0.90, "Gerando abas de Exclusões, Divergências e Conciliação...")
            ws_excl = wb.create_sheet(title="Lista de Exclusões")
            self._build_exclusions_sheet(ws_excl, results["exclusoes"])

            # 4. Aba Divergências Folha x Nomina
            ws_div = wb.create_sheet(title="Divergências Folha")
            self._build_divergences_sheet(ws_div, results["divergencias"])

            # 5. Aba Conciliação Fatura Seguradora (Critério 3 da Auditoria)
            ws_fat = wb.create_sheet(title="Conciliação Seguradora")
            self._build_fatura_reconciliation_sheet(ws_fat, results["conciliacao_fatura"])

            # 6. Aba Seguro Complementar
            ws_comp = wb.create_sheet(title="Seguro Complementar")
            self._build_complementar_audit_sheet(ws_comp, results["comp_12x"], results["comp_24x"])
        except Exception:
            try:
                wb.close()
            except Exception:
                pass
            raise

        saved_report = self._save_workbook_safe(wb, report_path)
        self._log(f"[*] Relatório de Auditoria gravado com sucesso: {saved_report.name}")
        return saved_report

    def _build_executive_summary_sheet(self, ws, kpis: Dict[str, Any]) -> None:
        """Constrói o painel executivo com cards e tabelas consolidadas, unificando paleta #1F497D."""
        ws.views.sheetView[0].showGridLines = True

        ws.merge_cells("A1:G1")
        title_cell = ws["A1"]
        title_cell.value = f"RELATÓRIO GERENCIAL DE CONFERÊNCIA DE SEGURO DE VIDA - CENCOSUD BRASIL ({self.competencia})"
        title_cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill(start_color=COR_KPI_HEADER, end_color=COR_KPI_HEADER, fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 40

        cards = [
            ("TOTAL NA NOMINA (+ COM GESTOR)", kpis["total_nomina"], "1F497D"),
            ("ELEGÍVEIS SEGURO NORMAL", kpis["total_elegiveis"], "1F497D"),
            ("DIRETORES NO BÁSICO (GB)", kpis.get("total_diretores_elegiveis", len(self.diretores)), "1F497D"),
            ("TOTAL EXCLUÍDOS POR REGRA", kpis["total_excluidos"], "9C0006"),
            ("ESTAGIÁRIOS MANTIDOS", kpis["total_estagiarios_mantidos"], "1F497D"),
            ("COMPLEMENTAR 12X", kpis["total_comp_12x"], "1F497D"),
            ("COMPLEMENTAR 24X", kpis["total_comp_24x"], "1F497D")
        ]

        col_start = 1
        for title, val, color in cards:
            c_letter = get_column_letter(col_start)

            cell_h = ws[f"{c_letter}3"]
            cell_h.value = title
            cell_h.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
            cell_h.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            cell_h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            cell_v = ws[f"{c_letter}4"]
            cell_v.value = val
            cell_v.font = Font(name="Calibri", size=16, bold=True, color=color)
            cell_v.fill = PatternFill(start_color=COR_KPI_BOX, end_color=COR_KPI_BOX, fill_type="solid")
            cell_v.alignment = Alignment(horizontal="center", vertical="center")
            cell_v.border = self.thin_border

            ws.column_dimensions[c_letter].width = 24
            col_start += 1

        ws.row_dimensions[3].height = 25
        ws.row_dimensions[4].height = 35

        # Tabela 1: Distribuição de Elegíveis por Bandeira (Padronizada com #1F497D)
        ws["A7"].value = "DISTRIBUIÇÃO DE ELEGÍVEIS POR BANDEIRA / PLANO"
        ws["A7"].font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        ws["A7"].fill = PatternFill(start_color=COR_CABECALHO_CADASTRAL, end_color=COR_CABECALHO_CADASTRAL, fill_type="solid")
        ws.merge_cells("A7:C7")

        ws.append(["Bandeira", "Tipo de Plano", "Vidas Elegíveis"])
        header_row = ws.max_row
        for col_idx in range(1, 4):
            c = ws.cell(row=header_row, column=col_idx)
            c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill(start_color=COR_CABECALHO_CADASTRAL, end_color=COR_CABECALHO_CADASTRAL, fill_type="solid")
            c.border = self.thin_border
            c.alignment = Alignment(horizontal="center")

        mapa_planos = {
            "DIRETORES (GBARBOSA)": "PLANO 01 (Aba Com Gestor)",
            "GBARBOSA": "PLANO 01",
            "PREZUNIC": "PLANO 01",
            "GIGA E MERCANTIL": "PLANO 01",
            "MERCANTIL": "PLANO 01",
            "GIGA": "PLANO 01",
            "PERINI": "PLANO 01",
            "BRETAS LOJAS": "PLANO 03",
            "BRETAS POSTOS": "PLANO 04",
            "JAGUAQUARA": "PLANO 07"
        }

        for bandeira, count in kpis["contagem_por_bandeira"].items():
            ws.append([bandeira, mapa_planos.get(bandeira, "Conforme Regra"), count])
            r = ws.max_row
            ws.cell(row=r, column=1).alignment = Alignment(horizontal="left")
            ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")
            ws.cell(row=r, column=3).alignment = Alignment(horizontal="right")
            ws.cell(row=r, column=3).number_format = "#,##0"
            for col_idx in range(1, 4):
                ws.cell(row=r, column=col_idx).border = self.thin_border

        # Tabela 2: Detalhamento das Exclusões por Motivo
        ws["E7"].value = "DETALHAMENTO DE EXCLUSÕES DA CARGA"
        ws["E7"].font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        ws["E7"].fill = PatternFill(start_color=COR_CABECALHO_CADASTRAL, end_color=COR_CABECALHO_CADASTRAL, fill_type="solid")
        ws.merge_cells("E7:F7")

        ws.cell(row=8, column=5, value="Motivo da Exclusão")
        ws.cell(row=8, column=6, value="Quantidade de Vidas")
        for c_idx in [5, 6]:
            c = ws.cell(row=8, column=c_idx)
            c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill(start_color=COR_CABECALHO_CADASTRAL, end_color=COR_CABECALHO_CADASTRAL, fill_type="solid")
            c.border = self.thin_border
            c.alignment = Alignment(horizontal="center")

        r_excl = 9
        for motivo, count in kpis["exclusoes_por_motivo"].items():
            ws.cell(row=r_excl, column=5, value=motivo).alignment = Alignment(horizontal="left")
            c_val = ws.cell(row=r_excl, column=6, value=count)
            c_val.alignment = Alignment(horizontal="right")
            c_val.number_format = "#,##0"
            for c_idx in [5, 6]:
                ws.cell(row=r_excl, column=c_idx).border = self.thin_border
            r_excl += 1

        self._auto_adjust_columns(ws)

    def _build_consolidated_base_sheet(self, ws, df: pd.DataFrame) -> None:
        """
        Constrói a aba Base Consolidada aplicando rigorosamente a regra bicolor:
        - Colunas Cadastrais (1 a 13): Azul Corporativo (#1F497D) / Fonte Branca.
        - Colunas de Conferência / Validação (14 a 20): Amarelo (#FFFF00) / Fonte Vermelha Negrito (#9C0006).
        """
        ws.views.sheetView[0].showGridLines = True

        colunas_cadastrais = [
            "Matrícula", "Colaborador", "CPF", "Data Nascimento", "Idade",
            "Data Admissão", "Cargo", "Estabelecimento", "Natureza Profissional",
            "Sindicato", "Categoria", "Tipo Salário", "Situação Especial"
        ]

        colunas_conferencia = [
            "Oposição SGV", "Desconto Folha 6266", "Desconto Comp 6267",
            "Capital Segurado Est. (R$)", "Custeio Empresa Est. (R$)",
            "Conformidade Desconto", "Status Conferência", "Motivo Detalhado"
        ]

        todas_colunas = colunas_cadastrais + colunas_conferencia
        ws.append(todas_colunas)

        fill_cadastral = PatternFill(start_color=COR_CABECALHO_CADASTRAL, end_color=COR_CABECALHO_CADASTRAL, fill_type="solid")
        font_cadastral = Font(name="Calibri", size=11, bold=True, color=COR_FONTE_CADASTRAL)

        fill_conferencia = PatternFill(start_color=COR_CABECALHO_CONFERENCIA, end_color=COR_CABECALHO_CONFERENCIA, fill_type="solid")
        font_conferencia = Font(name="Calibri", size=11, bold=True, color=COR_FONTE_CONFERENCIA)

        for col_idx in range(1, len(todas_colunas) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = self.thin_border

            if col_idx <= len(colunas_cadastrais):
                cell.fill = fill_cadastral
                cell.font = font_cadastral
            else:
                cell.fill = fill_conferencia
                cell.font = font_conferencia

        ws.row_dimensions[1].height = 28

        for row_idx, r in enumerate(df.itertuples(index=False), start=2):
            dt_nasc = r.data_nascimento.strftime("%d/%m/%Y") if pd.notna(r.data_nascimento) else ""
            dt_adm = r.data_admissao.strftime("%d/%m/%Y") if pd.notna(r.data_admissao) else ""

            row_data = [
                int(r.matricula),
                str(r.colaborador).strip().upper(),
                str(r.cpf).strip(),
                dt_nasc,
                int(r.idade),
                dt_adm,
                str(r.cargo).strip(),
                str(r.estabelecimento).strip(),
                str(r.natureza_profissional).strip(),
                str(r.sindicato).strip(),
                str(r.categoria).strip(),
                str(r.tipo_salario).strip(),
                str(r.situacao_especial).strip(),
                # Colunas de Conferência:
                "SIM" if r.tem_oposicao else "NÃO",
                float(r.desconto_6266),
                float(r.desconto_6267),
                float(r.capital_segurado_normal),
                float(r.custeio_empresa_estimado),
                str(r.conformidade_desconto).strip(),
                str(r.status_conferencia).strip(),
                str(r.motivo_detalhado).strip()
            ]
            ws.append(row_data)

            fill_row = self.white_fill if row_idx % 2 == 0 else self.zebra_fill

            for c_idx in range(1, len(row_data) + 1):
                c = ws.cell(row=row_idx, column=c_idx)
                c.fill = fill_row
                c.border = self.thin_border
                c.font = Font(name="Calibri", size=10)

                if c_idx in [1, 3, 4, 5, 6, 14, 19, 20]:
                    c.alignment = Alignment(horizontal="center")
                elif c_idx in [15, 16, 17, 18]:
                    c.alignment = Alignment(horizontal="right")
                    c.number_format = "#,##0.00"
                else:
                    c.alignment = Alignment(horizontal="left")

                if c_idx == 20 and r.status_conferencia == "EXCLUIDO":
                    c.font = Font(name="Calibri", size=10, bold=True, color="9C0006")

        self._auto_adjust_columns(ws)

    def _build_exclusions_sheet(self, ws, df_excl: pd.DataFrame) -> None:
        """
        Constrói a aba Lista de Exclusões aplicando rigorosamente o padrão visual:
        - Colunas Cadastrais (1 a 8): Azul Corporativo (#1F497D) / Fonte Branca.
        - Colunas de Conferência / Validação (9 a 11): Amarelo (#FFFF00) / Fonte Vermelha Negrito (#9C0006).
        """
        ws.views.sheetView[0].showGridLines = True

        colunas_cadastrais = [
            "Matrícula", "Colaborador", "CPF", "Idade", "Data Admissão",
            "Bandeira", "Cargo", "Estabelecimento"
        ]
        colunas_conferencia = [
            "Oposição SGV", "Status", "Motivo Detalhado da Exclusão"
        ]
        headers = colunas_cadastrais + colunas_conferencia
        ws.append(headers)

        fill_cadastral = PatternFill(start_color=COR_CABECALHO_CADASTRAL, end_color=COR_CABECALHO_CADASTRAL, fill_type="solid")
        font_cadastral = Font(name="Calibri", size=11, bold=True, color=COR_FONTE_CADASTRAL)

        fill_conferencia = PatternFill(start_color=COR_CABECALHO_CONFERENCIA, end_color=COR_CABECALHO_CONFERENCIA, fill_type="solid")
        font_conferencia = Font(name="Calibri", size=11, bold=True, color=COR_FONTE_CONFERENCIA)

        for c_idx in range(1, len(headers) + 1):
            c = ws.cell(row=1, column=c_idx)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = self.thin_border

            if c_idx <= len(colunas_cadastrais):
                c.fill = fill_cadastral
                c.font = font_cadastral
            else:
                c.fill = fill_conferencia
                c.font = font_conferencia

        for row_idx, r in enumerate(df_excl.itertuples(index=False), start=2):
            dt_adm = r.data_admissao.strftime("%d/%m/%Y") if pd.notna(r.data_admissao) else ""

            row_data = [
                int(r.matricula),
                str(r.colaborador).strip().upper(),
                str(r.cpf).strip(),
                int(r.idade),
                dt_adm,
                str(r.bandeira).strip(),
                str(r.cargo).strip(),
                str(r.estabelecimento).strip(),
                "SIM" if r.tem_oposicao else "NÃO",
                str(r.status_conferencia).strip(),
                str(r.motivo_detalhado).strip()
            ]
            ws.append(row_data)

            fill_row = self.white_fill if row_idx % 2 == 0 else self.zebra_fill
            for c_idx in range(1, len(row_data) + 1):
                c = ws.cell(row=row_idx, column=c_idx)
                c.fill = fill_row
                c.border = self.thin_border
                c.font = Font(name="Calibri", size=10)
                if c_idx in [1, 3, 4, 5, 9, 10]:
                    c.alignment = Alignment(horizontal="center")
                else:
                    c.alignment = Alignment(horizontal="left")

        self._auto_adjust_columns(ws)

    def _build_divergences_sheet(self, ws, divergencias: Dict[str, pd.DataFrame]) -> None:
        """
        Constrói a aba de divergências da Folha aplicando rigorosamente o padrão visual:
        - Colunas Cadastrais (1 a 5): Azul Corporativo (#1F497D) / Fonte Branca.
        - Colunas de Conferência / Auditoria (6 a 9): Amarelo (#FFFF00) / Fonte Vermelha Negrito (#9C0006).
        """
        ws.views.sheetView[0].showGridLines = True

        colunas_cadastrais = ["Matrícula", "Colaborador", "CPF", "Bandeira", "Cargo"]
        colunas_conferencia = ["Desconto 6266 (R$)", "Oposição SGV", "Tipo de Divergência", "Ação Recomendada RH"]
        headers = colunas_cadastrais + colunas_conferencia
        ws.append(headers)

        fill_cadastral = PatternFill(start_color=COR_CABECALHO_CADASTRAL, end_color=COR_CABECALHO_CADASTRAL, fill_type="solid")
        font_cadastral = Font(name="Calibri", size=11, bold=True, color=COR_FONTE_CADASTRAL)

        fill_conferencia = PatternFill(start_color=COR_CABECALHO_CONFERENCIA, end_color=COR_CABECALHO_CONFERENCIA, fill_type="solid")
        font_conferencia = Font(name="Calibri", size=11, bold=True, color=COR_FONTE_CONFERENCIA)

        for c_idx in range(1, len(headers) + 1):
            c = ws.cell(row=1, column=c_idx)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = self.thin_border

            if c_idx <= len(colunas_cadastrais):
                c.fill = fill_cadastral
                c.font = font_cadastral
            else:
                c.fill = fill_conferencia
                c.font = font_conferencia

        dfs_to_concat = []
        if not divergencias["desconto_indevido"].empty:
            df1 = divergencias["desconto_indevido"].copy()
            df1["acao"] = "Cessar desconto na folha (colaborador inelegível para o seguro)"
            dfs_to_concat.append(df1)

        if not divergencias["oposicao_descontada"].empty:
            df2 = divergencias["oposicao_descontada"].copy()
            df2["acao"] = "Estornar desconto em folha e validar vigência da oposição no Gen.te"
            dfs_to_concat.append(df2)

        if dfs_to_concat:
            df_all_div = pd.concat(dfs_to_concat, ignore_index=True).drop_duplicates(subset=["matricula"])
            for row_idx, r in enumerate(df_all_div.itertuples(index=False), start=2):
                row_data = [
                    int(r.matricula),
                    str(r.colaborador).strip().upper(),
                    str(r.cpf).strip(),
                    str(r.bandeira).strip(),
                    str(r.cargo).strip(),
                    float(r.desconto_6266),
                    "SIM" if r.tem_oposicao else "NÃO",
                    str(r.tipo_divergencia).strip(),
                    str(r.acao).strip()
                ]
                ws.append(row_data)

                fill_row = self.white_fill if row_idx % 2 == 0 else self.zebra_fill
                for c_idx in range(1, len(row_data) + 1):
                    c = ws.cell(row=row_idx, column=c_idx)
                    c.fill = fill_row
                    c.border = self.thin_border
                    c.font = Font(name="Calibri", size=10)
                    if c_idx in [1, 3, 7]:
                        c.alignment = Alignment(horizontal="center")
                    elif c_idx == 6:
                        c.alignment = Alignment(horizontal="right")
                        c.number_format = "#,##0.00"
                    else:
                        c.alignment = Alignment(horizontal="left")

        self._auto_adjust_columns(ws)

    def _build_fatura_reconciliation_sheet(self, ws, conciliacao_fatura: Dict[str, pd.DataFrame]) -> None:
        """
        Constrói a aba de Conciliação com a Fatura da Seguradora (Critério 3 da Auditoria):
        - Colunas Cadastrais (1 a 5): Azul Corporativo (#1F497D) / Fonte Branca.
        - Colunas de Conferência / Validação (6 a 8): Amarelo (#FFFF00) / Fonte Vermelha Negrito (#9C0006).
        """
        ws.views.sheetView[0].showGridLines = True

        colunas_cadastrais = ["Matrícula", "Colaborador", "CPF", "Bandeira", "Cargo"]
        colunas_conferencia = ["Prêmio Faturado (R$)", "Diagnóstico da Conciliação", "Ação Corretiva Solicitada"]
        headers = colunas_cadastrais + colunas_conferencia
        ws.append(headers)

        fill_cadastral = PatternFill(start_color=COR_CABECALHO_CADASTRAL, end_color=COR_CABECALHO_CADASTRAL, fill_type="solid")
        font_cadastral = Font(name="Calibri", size=11, bold=True, color=COR_FONTE_CADASTRAL)

        fill_conferencia = PatternFill(start_color=COR_CABECALHO_CONFERENCIA, end_color=COR_CABECALHO_CONFERENCIA, fill_type="solid")
        font_conferencia = Font(name="Calibri", size=11, bold=True, color=COR_FONTE_CONFERENCIA)

        for c_idx in range(1, len(headers) + 1):
            c = ws.cell(row=1, column=c_idx)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = self.thin_border

            if c_idx <= len(colunas_cadastrais):
                c.fill = fill_cadastral
                c.font = font_cadastral
            else:
                c.fill = fill_conferencia
                c.font = font_conferencia

        # Unifica cobranças indevidas e inclusões pendentes
        concil_dfs = []
        if not conciliacao_fatura["cobranca_indevida"].empty:
            concil_dfs.append(conciliacao_fatura["cobranca_indevida"])
        if not conciliacao_fatura["inclusoes_pendentes"].empty:
            concil_dfs.append(conciliacao_fatura["inclusoes_pendentes"])

        if concil_dfs:
            df_full_concil = pd.concat(concil_dfs, ignore_index=True)
            for row_idx, r in enumerate(df_full_concil.itertuples(index=False), start=2):
                premio = float(getattr(r, "premio_faturado", 0.0))
                row_data = [
                    int(r.matricula),
                    str(getattr(r, "colaborador", "")).strip().upper(),
                    str(getattr(r, "cpf", "")).strip(),
                    str(getattr(r, "bandeira", "")).strip(),
                    str(getattr(r, "cargo", "")).strip(),
                    premio,
                    str(getattr(r, "tipo_conciliacao", "")).strip(),
                    str(getattr(r, "acao_sugerida", "")).strip()
                ]
                ws.append(row_data)

                fill_row = self.white_fill if row_idx % 2 == 0 else self.zebra_fill
                for c_idx in range(1, len(row_data) + 1):
                    c = ws.cell(row=row_idx, column=c_idx)
                    c.fill = fill_row
                    c.border = self.thin_border
                    c.font = Font(name="Calibri", size=10)
                    if c_idx in [1, 3]:
                        c.alignment = Alignment(horizontal="center")
                    elif c_idx == 6:
                        c.alignment = Alignment(horizontal="right")
                        c.number_format = "#,##0.00"
                    else:
                        c.alignment = Alignment(horizontal="left")

        self._auto_adjust_columns(ws)

    def _build_complementar_audit_sheet(self, ws, df_12x: pd.DataFrame, df_24x: pd.DataFrame) -> None:
        """Constrói a aba de Seguro Complementar com a conferência de 12x e 24x."""
        ws.views.sheetView[0].showGridLines = True

        colunas_cadastrais = ["Modalidade", "Matrícula", "Nome do Colaborador", "Nascimento", "CPF"]
        colunas_conferencia = ["Salário Atual (R$)", "Capital Segurado Est. (R$)", "Plano Atribuído", "Origem"]
        headers = colunas_cadastrais + colunas_conferencia
        ws.append(headers)

        fill_cadastral = PatternFill(start_color=COR_CABECALHO_CADASTRAL, end_color=COR_CABECALHO_CADASTRAL, fill_type="solid")
        font_cadastral = Font(name="Calibri", size=11, bold=True, color=COR_FONTE_CADASTRAL)

        fill_conferencia = PatternFill(start_color=COR_CABECALHO_CONFERENCIA, end_color=COR_CABECALHO_CONFERENCIA, fill_type="solid")
        font_conferencia = Font(name="Calibri", size=11, bold=True, color=COR_FONTE_CONFERENCIA)

        for c_idx in range(1, len(headers) + 1):
            c = ws.cell(row=1, column=c_idx)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = self.thin_border

            if c_idx <= len(colunas_cadastrais):
                c.fill = fill_cadastral
                c.font = font_cadastral
            else:
                c.fill = fill_conferencia
                c.font = font_conferencia

        row_idx = 2
        for mod, df in [("Complementar 12x", df_12x), ("Complementar 24x", df_24x)]:
            for r in df.itertuples(index=False):
                dt_nasc = r.nascimento.strftime("%d/%m/%Y") if pd.notna(r.nascimento) else ""
                sal = float(r.salario) if pd.notna(r.salario) else ""
                cap = float(getattr(r, "capital_segurado_estimado", 0.0)) if pd.notna(getattr(r, "capital_segurado_estimado", np.nan)) else ""
                origem = getattr(r, "origem", "Regular")

                row_data = [
                    mod,
                    int(r.matricula),
                    str(r.nome).strip().upper(),
                    dt_nasc,
                    str(r.cpf).strip(),
                    sal,
                    cap,
                    str(r.plano).strip(),
                    origem
                ]
                ws.append(row_data)

                fill_row = self.white_fill if row_idx % 2 == 0 else self.zebra_fill
                for c_idx in range(1, len(row_data) + 1):
                    c = ws.cell(row=row_idx, column=c_idx)
                    c.fill = fill_row
                    c.border = self.thin_border
                    c.font = Font(name="Calibri", size=10)
                    if c_idx in [1, 2, 4, 5, 8, 9]:
                        c.alignment = Alignment(horizontal="center")
                    elif c_idx in [6, 7]:
                        c.alignment = Alignment(horizontal="right")
                        if c.value != "":
                            c.number_format = "#,##0.00"
                    else:
                        c.alignment = Alignment(horizontal="left")
                row_idx += 1

        self._auto_adjust_columns(ws)

    def _auto_adjust_columns(self, ws, max_rows: int = 150) -> None:
        """Ajusta automaticamente a largura das colunas amostrando até max_rows para alta performance."""
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col[:max_rows]:
                val = cell.value
                if val is not None:
                    s_len = len(str(val))
                    if s_len > max_len:
                        max_len = s_len
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
