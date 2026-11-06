"""
Adaptador de Execução do Motor Cesta Prêmio Bretas para a Central de Benefícios Cencosud.
"""

import os
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from engines.cesta_premio_bretas.business_rules import BusinessRulesCestaPremio
from engines.cesta_premio_bretas.data_loader import DataLoaderCestaPremio
from engines.cesta_premio_bretas.excel_exporter import ExcelExporterCestaPremio


def exec_auditoria_cesta_premio_bretas(
    raw_map: Dict[str, Any],
    output_dir: str,
    month: int = 10,
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
    _log(f"INICIANDO AUDITORIA CESTA PRÊMIO BRETAS | COMPETÊNCIA: {competencia_str}")
    _log("=" * 68)

    data_dir = raw_map.get("data_dir", "")
    input_path = Path(data_dir) if data_dir and os.path.exists(data_dir) else Path.cwd()

    loader = DataLoaderCestaPremio(input_dir=input_path)

    _prog(0.15, "Carregando base de elegibilidade e premiação Bretas...")
    raw_data = loader.load_all()
    _log(f"[1/2] Base Cesta Prêmio carregada: {len(raw_data['df_conf']):,} colaboradores.")

    _prog(0.50, "Auditando regras de admissão, cargo, CC e frequência/ponto...")
    rules = BusinessRulesCestaPremio()
    result = rules.process(raw_data)
    kpis = result["kpis"]

    _log("-" * 68)
    _log("RESULTADO DA AUDITORIA CESTA PRÊMIO BRETAS:")
    _log(f"  • Total Colaboradores na Base:  {kpis['total_colaboradores']:,}")
    _log(f"  • Colaboradores Premiados:      {kpis['total_elegiveis_premio']:,}")
    _log(f"  • Valor Cheio (R$ 250):         {kpis['compra_valor_cheio']:,}")
    _log(f"  • Carga Proporcional:           {kpis['compra_proporcional']:,}")
    _log(f"  • Férias no Período (Zerado):   {kpis['ferias_zeradas']:,}")
    _log(f"  • Afastamentos:                 {kpis['afastados']:,}")
    _log(f"  • Cargos Inelegíveis:           {kpis['cargo_inelegivel']:,}")
    _log(f"  • Centros de Custo Inelegíveis: {kpis['cc_inelegivel']:,}")
    _log(f"  • Valor Total Premiação:        R$ {kpis['total_valor_premio']:,.2f}")
    _log(f"  • Total Divergências:           {kpis['total_divergencias']:,}")
    _log("-" * 68)

    _prog(0.75, "Construindo relatório Excel com formatação institucional...")
    out_dir_path = Path(output_dir)
    out_dir_path.mkdir(parents=True, exist_ok=True)
    target_file = out_dir_path / f"Conferência Cesta Prêmio Bretas_Auditada_{competencia_arq}.xlsx"

    _prog(0.90, "Gravando arquivo Excel com proteção de bloqueio...")
    while True:
        try:
            exporter = ExcelExporterCestaPremio(output_path=target_file)
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
                    target_file = out_dir_path / f"Conferência Cesta Prêmio Bretas_Auditada_{competencia_arq}_{ts}.xlsx"
                    exporter = ExcelExporterCestaPremio(output_path=target_file)
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
    _prog(1.0, "Auditoria Cesta Prêmio Bretas concluída com sucesso!")

    return str(target_file)
