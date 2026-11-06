"""
Adaptador de Execução do Motor Cencopay para a Central de Benefícios Cencosud.
"""

import os
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from engines.cencopay.business_rules import BusinessRulesCencopay
from engines.cencopay.data_loader import DataLoaderCencopay
from engines.cencopay.excel_exporter import ExcelExporterCencopay


def exec_auditoria_cencopay(
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
    _log(f"INICIANDO AUDITORIA BENEFÍCIOS CENCOPAY | COMPETÊNCIA: {competencia_str}")
    _log("=" * 68)

    data_dir = raw_map.get("data_dir", "")
    base_dir = Path(data_dir) if data_dir and os.path.exists(data_dir) else Path.cwd()
    input_dir = base_dir / "Relatório Original" if (base_dir / "Relatório Original").exists() else base_dir

    loader = DataLoaderCencopay(input_dir=input_dir, base_dir=base_dir)

    _prog(0.15, "Carregando bases cadastrais e histórico Cencopay...")
    raw_data = loader.load_all()
    _log(f"[1/3] Bases Cencopay carregadas: {len(raw_data['df_conf_base']):,} colaboradores na base.")

    _prog(0.50, "Processando motor de regras (Margem 30%, Eventos 5081, 14003 e 14004)...")
    rules = BusinessRulesCencopay()
    result = rules.process_conferencia(raw_data)
    kpis = result["kpis"]

    _log("-" * 68)
    _log("RESULTADO DA AUDITORIA CENCOPAY:")
    _log(f"  • Total Colaboradores:          {kpis['total_colaboradores']:,}")
    _log(f"  • Total Divergências:           {kpis['total_divergencias']:,}")
    _log(f"  • Taxa de Conformidade:         {kpis['taxa_conformidade']:.2f}%")
    _log(f"  • Saldo Devedor Apurado:        R$ {kpis['total_saldo_devedor']:,.2f}")
    _log(f"  • Desconto em Folha (5081):     R$ {kpis['total_desconto_folha']:,.2f}")
    _log(f"  • Saldo Mês Seguinte (14003):   R$ {kpis['total_saldo_mes_seguinte']:,.2f}")
    _log("-" * 68)

    _prog(0.75, "Construindo relatório Excel com formatação institucional...")
    out_dir_path = Path(output_dir)
    out_dir_path.mkdir(parents=True, exist_ok=True)
    target_file = out_dir_path / f"Conferência Cencopay_Auditada_{competencia_arq}.xlsx"

    _prog(0.90, "Gravando arquivo Excel com proteção de bloqueio...")
    while True:
        try:
            exporter = ExcelExporterCencopay(output_path=target_file)
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
                    target_file = out_dir_path / f"Conferência Cencopay_Auditada_{competencia_arq}_{ts}.xlsx"
                    exporter = ExcelExporterCencopay(output_path=target_file)
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
    _prog(1.0, "Auditoria de Benefícios Cencopay concluída com sucesso!")

    return str(target_file)
