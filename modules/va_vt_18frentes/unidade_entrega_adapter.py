"""
Adaptador de Execução do Motor Unidade de Entrega VT para a Central de Benefícios Cencosud.
"""

import os
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from engines.unidade_entrega.business_rules import BusinessRulesUnidadeEntrega
from engines.unidade_entrega.data_loader import DataLoaderUnidadeEntrega
from engines.unidade_entrega.excel_exporter import ExcelExporterUnidadeEntrega


def exec_auditoria_unidade_entrega(
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
    _log(f"INICIANDO AUDITORIA UNIDADE DE ENTREGA VT | COMPETÊNCIA: {competencia_str}")
    _log("=" * 68)

    data_dir = raw_map.get("data_dir", "")
    input_path = Path(data_dir) if data_dir and os.path.exists(data_dir) else Path.cwd()

    loader = DataLoaderUnidadeEntrega(input_dir=input_path)

    _prog(0.15, "Carregando Relação de Compra de Vale Transporte (RELAOD)...")
    df_compra = loader.load_compra_vt()
    _log(f"[1/2] Compra VT carregada: {len(df_compra):,} registros.")

    _prog(0.35, "Carregando parametrização de Optantes...")
    df_optantes = loader.load_optantes()
    _log(f"[2/2] Optantes carregados: {len(df_optantes):,} registros.")

    raw_data = {
        "df_compra": df_compra,
        "df_optantes": df_optantes,
    }

    _prog(0.55, "Auditando unidades de entrega x lojas de lotação...")
    rules = BusinessRulesUnidadeEntrega()
    result = rules.process(raw_data)
    kpis = result["kpis"]

    _log("-" * 68)
    _log("RESULTADO DA AUDITORIA DE UNIDADES DE ENTREGA:")
    _log(f"  • Total Linhas Auditadas:       {kpis['total_linhas']:,}")
    _log(f"  • Utiliza UE Padrão (Loja):     {kpis['padrao_ok']:,}")
    _log(f"  • Utiliza UE de Optantes:       {kpis['nao_padrao_ok']:,}")
    _log(f"  • Numeração Diferente Homol.:   {kpis['diff_num_ok']:,}")
    _log(f"  • Operações Spid:               {kpis['spid_ok']:,}")
    _log(f"  • Divergências Reais:           {kpis['total_divergencias']:,}")
    _log(f"  • Taxa de Conformidade:         {kpis['taxa_conformidade']:.2f}%")
    _log("-" * 68)

    _prog(0.75, "Construindo relatório Excel com formatação institucional...")
    out_dir_path = Path(output_dir)
    out_dir_path.mkdir(parents=True, exist_ok=True)
    target_file = out_dir_path / f"Conferência Unidade de Entrega_{competencia_arq}.xlsx"

    _prog(0.90, "Gravando arquivo Excel com proteção de arquivo aberto...")
    while True:
        try:
            exporter = ExcelExporterUnidadeEntrega(output_path=target_file)
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
                    target_file = out_dir_path / f"Conferência Unidade de Entrega_{competencia_arq}_{ts}.xlsx"
                    exporter = ExcelExporterUnidadeEntrega(output_path=target_file)
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
    _prog(1.0, "Auditoria de Unidade de Entrega concluída com sucesso!")

    return str(target_file)
