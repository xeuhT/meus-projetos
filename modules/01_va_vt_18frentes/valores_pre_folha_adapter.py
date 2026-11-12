"""
Adaptador de Execução Assíncrona do Motor Valores Pré-Folha para a Interface Gráfica Cencosud.

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

from engines.valores_pre_folha.data_loader import DataLoader
from engines.valores_pre_folha.business_rules import BusinessRulesPrefolha
from engines.valores_pre_folha.excel_exporter import ExcelExporterPrefolha


def exec_auditoria_valores_pre_folha(
    raw_map: Dict[str, Any],
    output_dir: str,
    month: int = 9,
    year: int = 2026,
    progress_cb: Optional[Callable[[float, str], None]] = None,
    log_cb: Optional[Callable[[str], None]] = None,
    prompt_lock_cb: Optional[Callable[[str], str]] = None,
) -> str:
    """
    Executa a Auditoria de Valores Pré-Folha VT com telemetria em tempo real para a GUI.
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
    _log(f"INICIANDO AUDITORIA VALORES PRÉ-FOLHA VT | COMPETÊNCIA: {competencia_str}")
    _log("=" * 68)

    base_file = raw_map.get("conferencia_base")
    data_dir = raw_map.get("data_dir", "")
    if not base_file or not os.path.exists(base_file):
        if data_dir and os.path.exists(data_dir):
            import glob
            matches = glob.glob(os.path.join(data_dir, "*onfer*.xlsx"))
            if matches:
                base_file = matches[0]
            else:
                raise FileNotFoundError(f"Planilha de conferência base (*onfer*.xlsx) não encontrada em {data_dir}")
        else:
            raise FileNotFoundError("Arquivo base de Conferência Pré-Folha não localizado.")

    conf_path = Path(base_file)
    loader = DataLoader(input_dir=conf_path.parent)

    # 1. Carregamento da FOPA Pré-Folha
    _prog(0.15, f"Carregando FOPA Pré-Folha de {conf_path.name}...")
    df_fopa = loader.load_fopa_prefolha(conf_path)
    _log(f"[1/3] FOPA Pré-Folha carregada: {len(df_fopa):,} colaboradores na base.")

    # 2. Carregamento do Lote Importado 31935
    _prog(0.35, "Carregando lote de importação do Evento 31935 (VT Alternativo)...")
    df_31935 = loader.load_import_31935(conf_path)
    _log(f"[2/3] Lote Evento 31935 carregado: {len(df_31935):,} colaboradores.")

    # 3. Carregamento do Lote Importado 6292
    _prog(0.55, "Carregando lote de importação do Evento 6292 (Desconto VTE)...")
    df_6292 = loader.load_import_6292(conf_path)
    _log(f"[3/3] Lote Evento 6292 carregado: {len(df_6292):,} registros.")

    raw_data = {
        "df_fopa": df_fopa,
        "df_31935": df_31935,
        "df_6292": df_6292,
    }

    # 4. Motor de Regras e Auditoria
    _prog(0.72, "Auditando soma de pagamentos vs evento 5750 e lotes 31935 / 6292...")
    rules = BusinessRulesPrefolha()
    result = rules.process_conferencia(raw_data)
    kpis = result["kpis"]

    _log("-" * 68)
    _log("RESULTADO DA CONCILIAÇÃO VALORES PRÉ-FOLHA VT:")
    _log(f"  • Total FOPA Auditada:          {kpis.get('total_fopa', len(df_fopa)):,} colaboradores")
    _log(f"  • Conciliados em Plena Conf.:   {kpis.get('conciliados_fopa', 0):,} ({kpis.get('taxa_conformidade_fopa', 0.0):.2f}%)")
    _log(f"  • Soma Eventos de Pagamento:    R$ {kpis.get('soma_pgtos', 0.0):,.2f}")
    _log(f"  • Evento 5750 em Folha:         R$ {kpis.get('soma_5750', 0.0):,.2f}")
    _log(f"  • Total Divergências Apuradas:  {kpis.get('total_divergencias_consolidadas', 0):,}")
    _log("-" * 68)

    # 5. Construção e Exportação do Excel
    _prog(0.88, "Construindo relatório Excel formatado (4 abas estáticas)...")
    out_dir_path = Path(output_dir)
    out_dir_path.mkdir(parents=True, exist_ok=True)
    target_file = out_dir_path / f"Conferência Valores Pré-Folha VT_{competencia_arq}.xlsx"

    _prog(0.95, "Gravando arquivo Excel com proteção de bloqueio...")
    while True:
        try:
            exporter = ExcelExporterPrefolha(output_path=target_file)
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
                    target_file = out_dir_path / f"Conferência Valores Pré-Folha VT_{competencia_arq}_{ts}.xlsx"
                    exporter = ExcelExporterPrefolha(output_path=target_file)
                    saved_path = exporter.export(result)
                    target_file = Path(saved_path)
                    break
                else:
                    raise RuntimeError("Gravação cancelada pelo usuário (arquivo estava aberto no Excel).")
            else:
                raise

    elapsed = time.time() - t0
    size_mb = target_file.stat().st_size / (1024 * 1024)
    _log(f"Planilha gravada com sucesso ({size_mb:.2f} MB) em {elapsed:.2f}s:")
    _log(f"  -> {target_file}")
    _prog(1.0, "Auditoria de Valores Pré-Folha VT concluída com sucesso!")

    return str(target_file)
