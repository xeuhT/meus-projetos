"""
Adaptador de Execução Assíncrona do Motor Compra Mensal VT para a Interface Gráfica Cencosud.

Integra os módulos homologados (data_loader, business_rules, excel_exporter) com:
- Barra de progresso real etapa a etapa (progress_cb).
- Console de logs em tempo real (log_cb).
- Proteção contra bloqueio de arquivo aberto no Excel (prompt_lock_cb / FileLockModal).
"""

import os
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Optional

import openpyxl

from engines.compra_mensal_vt.data_loader import DataLoader
from engines.compra_mensal_vt.business_rules import AuditEngine
from engines.compra_mensal_vt.excel_exporter import ExcelExporter


def exec_auditoria_compra_mensal_vt(
    raw_map: Dict[str, Any],
    output_dir: str,
    month: int = 10,
    year: int = 2026,
    progress_cb: Optional[Callable[[float, str], None]] = None,
    log_cb: Optional[Callable[[str], None]] = None,
    prompt_lock_cb: Optional[Callable[[str], str]] = None,
) -> str:
    """
    Executa a Auditoria de Compra Mensal de Vale Transporte com telemetria para a GUI.
    Retorna o caminho completo do arquivo Excel gerado.
    """
    def _prog(pct: float, msg: str) -> None:
        if progress_cb:
            progress_cb(pct, msg)

    def _log(msg: str) -> None:
        if log_cb:
            log_cb(msg)

    t0 = time.time()
    competencia_str = f"{month:02d}/{year}"
    competencia_arq = f"{month:02d}.{year}"

    _log("=" * 68)
    _log(f"INICIANDO AUDITORIA COMPRA MENSAL VT | COMPETÊNCIA: {competencia_str}")
    _log("=" * 68)

    data_dir = raw_map.get("data_dir", "")
    if not data_dir or not os.path.exists(data_dir):
        raise FileNotFoundError(f"Diretório de relatórios brutos não encontrado: {data_dir}")

    loader = DataLoader(input_dir=data_dir)

    # 1. Carga do Censo
    _prog(0.10, "Carregando Censo Geral de Colaboradores e Nomina...")
    df_censo = loader.load_censo()
    _log(f"[1/6] Censo Funcional carregado: {len(df_censo):,} colaboradores cadastrados.")

    # 2. Carga de Optantes
    _prog(0.25, "Carregando base de Optantes VT e parâmetros de linha...")
    df_optantes = loader.load_optantes()
    _log(f"[2/6] Optantes VT vigentes: {len(df_optantes):,} cadastros de linha ativos.")

    # 3. Carga das Compras
    _prog(0.40, "Carregando Relação de Compra Regular e Complementar...")
    df_compras = loader.load_compras()
    _log(f"[3/6] Compras de VT identificadas: {len(df_compras):,} registros de recarga.")

    # 4. Carga de Afastados
    _prog(0.55, "Processando colaboradores afastados (Ativos e Cessados)...")
    df_afastados = loader.load_afastados()
    _log(f"[4/6] Afastados processados: {len(df_afastados):,} registros consolidados.")

    # 5. Carga de Férias
    _prog(0.68, f"Carregando programação de férias da competência {competencia_str}...")
    df_ferias = loader.load_ferias()
    _log(f"[5/6] Férias vigentes no período: {len(df_ferias):,} colaboradores.")

    # 6. Carga de Desligados
    _prog(0.78, "Processando colaboradores desligados/rescindidos...")
    df_desligados = loader.load_desligados()
    _log(f"[6/6] Desligados identificados: {len(df_desligados):,} registros.")

    raw_data = {
        "censo": df_censo,
        "optantes": df_optantes,
        "compras": df_compras,
        "afastados": df_afastados,
        "ferias": df_ferias,
        "desligados": df_desligados,
    }

    # 7. Motor de Auditoria e Regras de Negócio
    _prog(0.85, "Cruzando dados e validando elegibilidade de recarga de VT...")
    engine = AuditEngine(raw_data)
    audit_results = engine.run_audit()

    kpis = audit_results["kpis"]
    _log("-" * 68)
    _log("RESULTADO DA AUDITORIA DA COMPRA MENSAL VT:")
    _log(f"  • Optantes Ativos Vigentes:     {kpis.get('total_optantes_ativos', 0):,} colaboradores")
    _log(f"  • Com Compra Gerada no Mês:     {kpis.get('total_com_compra', 0):,} colaboradores")
    _log(f"  • Taxa de Cobertura da Compra:  {kpis.get('cobertura_pct', 0.0):.1f}%")
    _log(f"  • Optantes sem Compra Gerada:   {kpis.get('total_sem_compra', 0):,}")
    _log(f"  • Desligados com Compra:        {kpis.get('desligados_com_compra', 0)}")
    _log(f"  • Custo Consolidado da Carga:   R$ {kpis.get('valor_total_consolidado', 0.0):,.2f}")
    _log(f"  • Quantidade Total de Vales:    {kpis.get('qtd_vales_total', 0):,} passagens")
    _log("-" * 68)

    # 8. Construção e Exportação do Relatório Excel
    _prog(0.92, "Construindo relatório Excel formatado (4 abas)...")
    out_dir_path = Path(output_dir)
    out_dir_path.mkdir(parents=True, exist_ok=True)
    target_file = out_dir_path / f"Conferência Compra Mensal VT_{competencia_arq}.xlsx"

    wb = openpyxl.Workbook()
    try:
        ws_orient = wb.active
        ws_orient.title = "Orientações"
        exporter = ExcelExporter(str(target_file))
        exporter._build_orientacoes_sheet(ws_orient)

        ws_resumo = wb.create_sheet(title="Resumo da Auditoria")
        exporter._build_resumo_sheet(ws_resumo, kpis, audit_results["distribuicao"])

        ws_conf = wb.create_sheet(title="Conferência Geral")
        exporter._build_data_sheet(ws_conf, audit_results["df_conferencia"], is_divergence_only=False)

        ws_div = wb.create_sheet(title="Divergências e Inconsistências")
        exporter._build_data_sheet(ws_div, audit_results["df_divergencias"], is_divergence_only=True)

        wb.active = 0

        _prog(0.96, "Gravando arquivo Excel com proteção de bloqueio...")
        while True:
            try:
                wb.save(str(target_file))
                break
            except PermissionError:
                if prompt_lock_cb is not None:
                    choice = prompt_lock_cb(str(target_file))
                    if choice == "retry":
                        continue
                    elif choice == "copy":
                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        target_file = out_dir_path / f"Conferência Compra Mensal VT_{competencia_arq}_{ts}.xlsx"
                        wb.save(str(target_file))
                        break
                    else:
                        raise RuntimeError("Gravação cancelada pelo usuário (arquivo estava aberto no Excel).")
                else:
                    raise
    finally:
        try:
            wb.close()
        except Exception:
            pass

    elapsed = time.time() - t0
    size_mb = target_file.stat().st_size / (1024 * 1024)
    _log(f"Planilha gravada com sucesso ({size_mb:.2f} MB) em {elapsed:.2f}s:")
    _log(f"  -> {target_file}")
    _prog(1.0, "Auditoria da Compra Mensal VT concluída com sucesso!")

    return str(target_file)
