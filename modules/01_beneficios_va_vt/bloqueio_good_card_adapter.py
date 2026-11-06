"""
Adaptador de Execução Assíncrona do Motor Bloqueio Good Card para a Interface Gráfica Cencosud.

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

from engines.bloqueio_good_card.data_loader import DataLoader
from engines.bloqueio_good_card.business_rules import BusinessRulesBloqueioGoodCard
from engines.bloqueio_good_card.excel_exporter import ExcelExporterBloqueioGoodCard


def exec_auditoria_bloqueio_good_card(
    raw_map: Dict[str, Any],
    output_dir: str,
    month: int = 9,
    year: int = 2026,
    progress_cb: Optional[Callable[[float, str], None]] = None,
    log_cb: Optional[Callable[[str], None]] = None,
    prompt_lock_cb: Optional[Callable[[str], str]] = None,
) -> str:
    """
    Executa a Auditoria de Bloqueio Good Card com telemetria em tempo real para a GUI.
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
    _log(f"INICIANDO AUDITORIA BLOQUEIO GOOD CARD | COMPETÊNCIA: {competencia_str}")
    _log("=" * 68)

    base_file = raw_map.get("bloqueio_base")
    data_dir = raw_map.get("data_dir", "")
    if not base_file or not os.path.exists(base_file):
        if data_dir and os.path.exists(data_dir):
            import glob
            matches = glob.glob(os.path.join(data_dir, "*loqueio*.xlsx"))
            if matches:
                base_file = matches[0]
            else:
                raise FileNotFoundError(f"Planilha de bloqueio (*loqueio*.xlsx) não encontrada em {data_dir}")
        else:
            raise FileNotFoundError("Arquivo base de Bloqueio Good Card não localizado.")

    xlsx_path = Path(base_file)
    loader = DataLoader(input_dir=xlsx_path.parent)

    # 1. Carregamento do Mapeamento por Bandeira
    _prog(0.15, f"Carregando abas de bandeiras em {xlsx_path.name}...")
    card_bandeira = loader.load_bandeira_map(xlsx_path)
    _log(f"[1/3] Mapeamento de bandeiras concluído: {len(card_bandeira):,} referências de cartões.")

    # 2. Carregamento dos Afastados
    _prog(0.35, "Carregando base de afastamentos vigentes e períodos...")
    afast_map = loader.load_afastados(xlsx_path)
    _log(f"[2/3] Colaboradores afastados mapeados: {len(afast_map):,} colaboradores.")

    # 3. Carregamento e Cruzamento da Aba Geral
    _prog(0.55, "Cruzando cartões da aba Geral com regras de afastamento...")
    df_cartoes = loader.load_cartoes_geral(xlsx_path, card_bandeira, afast_map)
    _log(f"[3/3] Base geral consolidada: {len(df_cartoes):,} cartões analisados.")

    raw_data = {
        "df_cartoes": df_cartoes,
        "total_afastados_base": len(afast_map),
    }

    # 4. Motor de Regras e Auditoria
    _prog(0.72, "Auditando limiar de 15 dias de afastamento e status de bloqueio...")
    rules = BusinessRulesBloqueioGoodCard()
    result = rules.process_conferencia(raw_data)
    kpis = result["kpis"]

    _log("-" * 68)
    _log("RESULTADO DA AUDITORIA BLOQUEIO GOOD CARD:")
    _log(f"  • Total de Cartões Auditados:     {kpis.get('total_cartoes', len(df_cartoes)):,}")
    _log(f"  • Afastados com Cartão Ativo:    {kpis.get('total_afast_com_cartao', 0):,}")
    _log(f"  • Cartões com Bloqueio Recom.:   {kpis.get('bloqueio_recomendado', 0):,}")
    _log(f"  • Cartões Regularmente Ativos:   {kpis.get('regular_ativo', 0):,}")
    _log("-" * 68)

    # 5. Construção e Exportação do Excel
    _prog(0.88, "Construindo relatório Excel formatado (4 abas estáticas)...")
    out_dir_path = Path(output_dir)
    out_dir_path.mkdir(parents=True, exist_ok=True)
    target_file = out_dir_path / f"Conferência Bloqueio Good Card_{competencia_arq}.xlsx"

    _prog(0.95, "Gravando arquivo Excel com proteção de bloqueio...")
    while True:
        try:
            exporter = ExcelExporterBloqueioGoodCard(output_path=target_file)
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
                    target_file = out_dir_path / f"Conferência Bloqueio Good Card_{competencia_arq}_{ts}.xlsx"
                    exporter = ExcelExporterBloqueioGoodCard(output_path=target_file)
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
    _prog(1.0, "Auditoria de Bloqueio Good Card concluída com sucesso!")

    return str(target_file)
