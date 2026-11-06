"""
Adaptador de Execução do Motor Compra Mensal Ticket Bretas para a Central de Benefícios Cencosud.
"""

import os
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from engines.compra_ticket_bretas.business_rules import BusinessRulesTicketBretas
from engines.compra_ticket_bretas.data_loader import DataLoaderTicketBretas
from engines.compra_ticket_bretas.excel_exporter import ExcelExporterTicketBretas


def exec_auditoria_compra_ticket_bretas(
    raw_map: Dict[str, Any],
    output_dir: str,
    month: int = 9,
    year: int = 2026,
    progress_cb: Optional[Callable[[float, str], None]] = None,
    log_cb: Optional[Callable[[str], None]] = None,
    prompt_lock_cb: Optional[Callable[[str], str]] = None,
) -> str:
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
    _log(f"INICIANDO AUDITORIA COMPRA TICKET BRETAS | COMPETÊNCIA: {competencia_str}")
    _log("=" * 68)

    data_dir = raw_map.get("data_dir", "")
    input_path = Path(data_dir) if data_dir and os.path.exists(data_dir) else Path.cwd()

    loader = DataLoaderTicketBretas(input_dir=input_path)

    _prog(0.15, "Carregando base de conferência e remessa Ticket Bretas...")
    raw_data = loader.load_all()
    _log(f"[1/2] Base Ticket Bretas carregada: {len(raw_data['df_conf']):,} colaboradores.")

    _prog(0.50, "Auditando elegibilidade sindical, jornada, férias e compra manual...")
    rules = BusinessRulesTicketBretas()
    result = rules.process(raw_data)
    kpis = result["kpis"]

    _log("-" * 68)
    _log("RESULTADO DA AUDITORIA COMPRA TICKET BRETAS:")
    _log(f"  • Total Colaboradores na Base:  {kpis['total_colaboradores']:,}")
    _log(f"  • Elegíveis Compra Manual:      {kpis['total_elegiveis_compra']:,}")
    _log(f"  • Carga de Valor Fixo:          {kpis['compra_valor_fixo']:,}")
    _log(f"  • Total com Compra Aprovada:    {kpis['total_compra_ok']:,}")
    _log(f"  • Afastados no Período:         {kpis['afastados']:,}")
    _log(f"  • Férias Integrais Zeradas:     {kpis['ferias_zeradas']:,}")
    _log(f"  • Valor Total Compra Prevista:  R$ {kpis['total_valor_compra']:,.2f}")
    _log(f"  • Total Divergências:           {kpis['total_divergencias']:,}")
    _log("-" * 68)

    _prog(0.75, "Construindo relatório Excel com formatação institucional...")
    out_dir_path = Path(output_dir)
    out_dir_path.mkdir(parents=True, exist_ok=True)
    target_file = out_dir_path / f"Conferência Compra Ticket Bretas_Auditada_{competencia_arq}.xlsx"

    _prog(0.90, "Gravando arquivo Excel com proteção de bloqueio...")
    while True:
        try:
            exporter = ExcelExporterTicketBretas(output_path=target_file)
            saved_path = exporter.export(result)
            target_file = Path(saved_path)
            break
        except PermissionError:
            if prompt_lock_cb is not None:
                choice = prompt_lock_cb(str(target_file))
                if choice == "retry":
                    continue
                elif choice == "copy":
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    target_file = out_dir_path / f"Conferência Compra Ticket Bretas_Auditada_{competencia_arq}_{ts}.xlsx"
                    exporter = ExcelExporterTicketBretas(output_path=target_file)
                    saved_path = exporter.export(result)
                    target_file = Path(saved_path)
                    break
                else:
                    raise RuntimeError("Gravação cancelada pelo usuário.")
            else:
                raise

    elapsed = time.time() - t0
    size_mb = target_file.stat().st_size / (1024 * 1024)
    _log(f"Planilha gravada com sucesso ({size_mb:.2f} MB) em {elapsed:.2f}s:")
    _log(f"  -> {target_file}")
    _prog(1.0, "Auditoria Compra Ticket Bretas concluída com sucesso!")

    return str(target_file)
