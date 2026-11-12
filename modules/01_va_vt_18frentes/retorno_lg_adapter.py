"""
Adaptador de Execução do Motor Arquivo de Retorno LG para a Central de Benefícios Cencosud.
"""

import os
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from engines.retorno_lg.business_rules import BusinessRulesRetornoLG
from engines.retorno_lg.data_loader import DataLoaderRetornoLG
from engines.retorno_lg.excel_exporter import ExcelExporterRetornoLG


def exec_auditoria_retorno_lg(
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
    _log(f"INICIANDO AUDITORIA ARQUIVO DE RETORNO LG | COMPETÊNCIA: {competencia_str}")
    _log("=" * 68)

    data_dir = raw_map.get("data_dir", "")
    input_path = Path(data_dir) if data_dir and os.path.exists(data_dir) else Path.cwd()

    loader = DataLoaderRetornoLG(input_dir=input_path)

    _prog(0.15, "Carregando arquivos de retorno Alelo (Gbarbosa e Mercantil)...")
    raw_data = loader.load_all()
    _log(f"[1/2] Dados brutos carregados: {len(raw_data['df_alelo']):,} pedidos processados.")

    _prog(0.45, "Processando conciliação Sem RC vs Com RC e gerando layouts LG...")
    rules = BusinessRulesRetornoLG()
    result = rules.process(raw_data)
    kpis = result["kpis"]

    _log("-" * 68)
    _log("RESULTADO DA AUDITORIA ARQUIVO DE RETORNO LG:")
    _log(f"  • Total Colaboradores:          {kpis['total_colaboradores']:,}")
    _log(f"  • Com Gestão de Saldo:          {kpis['com_gestao']:,}")
    _log(f"  • Sem Gestão de Saldo:          {kpis['sem_gestao']:,}")
    _log(f"  • Total Sem RC:                 R$ {kpis['total_sem_rc']:,.2f}")
    _log(f"  • Total Com RC:                 R$ {kpis['total_com_rc']:,.2f}")
    _log(f"  • Economia Gestão de Saldo:     R$ {kpis['total_economia_gestao']:,.2f}")
    _log(f"  • Linhas a Importar (31903):    {kpis['total_linhas_importar']:,}")
    _log(f"  • Linhas a Excluir (Zerados):   {kpis['total_linhas_excluir']:,}")
    _log("-" * 68)

    _prog(0.75, "Construindo relatório Excel com formatação institucional...")
    out_dir_path = Path(output_dir)
    out_dir_path.mkdir(parents=True, exist_ok=True)
    target_file = out_dir_path / f"Conferência Arquivo Retorno LG_Auditada_{competencia_arq}.xlsx"

    _prog(0.90, "Gravando arquivo Excel com proteção de bloqueio...")
    while True:
        try:
            exporter = ExcelExporterRetornoLG(output_path=target_file)
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
                    target_file = out_dir_path / f"Conferência Arquivo Retorno LG_Auditada_{competencia_arq}_{ts}.xlsx"
                    exporter = ExcelExporterRetornoLG(output_path=target_file)
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
    _prog(1.0, "Auditoria do Arquivo de Retorno LG concluída com sucesso!")

    return str(target_file)
