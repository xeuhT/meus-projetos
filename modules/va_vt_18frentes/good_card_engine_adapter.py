"""
Adaptador de Execução Assíncrona do Motor Good Card para a Interface Gráfica Cencosud.

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

from data_loader import (
    load_afastados,
    load_colaboradores,
    load_demitidos,
    load_ferias,
    load_fopa,
    load_valores_importados,
)
from business_rules import execute_audit
from excel_exporter import create_audit_workbook


def exec_auditoria_good_card(
    raw_map: Dict[str, Any],
    output_dir: str,
    month: int = 9,
    year: int = 2026,
    progress_cb: Optional[Callable[[float, str], None]] = None,
    log_cb: Optional[Callable[[str], None]] = None,
    prompt_lock_cb: Optional[Callable[[str], str]] = None,
) -> str:
    """
    Executa a Auditoria 04 - Good Card com telemetria em tempo real para a GUI.
    Retorna o caminho completo do arquivo Excel gerado.
    """
    def _prog(pct: float, msg: str) -> None:
        if progress_cb:
            progress_cb(pct, msg)

    def _log(msg: str) -> None:
        if log_cb:
            log_cb(msg)

    t0 = time.time()
    ref_date = f"{year:04d}-{month:02d}-01"

    _log("=" * 68)
    _log(f"INICIANDO AUDITORIA 04 - GOOD CARD | COMPETÊNCIA: {month:02d}/{year}")
    _log("=" * 68)

    # 1. Carregamento de FOPA
    _prog(0.08, "Carregando Folha de Pagamento (FOPA)...")
    fopa_path = raw_map.get("fopa")
    if not fopa_path or not os.path.exists(fopa_path):
        raise FileNotFoundError("Arquivo FOPA_Eventos_por_Colaborador.CSV não encontrado.")
    df_fopa = load_fopa(fopa_path)
    _log(f"[1/6] FOPA carregado: {len(df_fopa):,} colaboradores ativos na folha.")

    # 2. Carregamento de Nomina
    _prog(0.20, "Carregando Cadastro Geral Ativo (Nomina)...")
    nom_path = raw_map.get("nomina")
    if not nom_path or not os.path.exists(nom_path):
        raise FileNotFoundError("Arquivo COLA_Colaboradores.CSV não encontrado.")
    df_nomina = load_colaboradores(nom_path)
    _log(f"[2/6] Nomina carregada: {len(df_nomina):,} registros cadastrais.")

    # 3. Carregamento de Afastados
    _prog(0.32, "Processando e filtrando Colaboradores Afastados...")
    afast_path = raw_map.get("afastados")
    if not afast_path or not os.path.exists(afast_path):
        raise FileNotFoundError("Arquivo COLA_Colaboradores_Afastados.CSV não encontrado.")
    df_afastados = load_afastados(afast_path, ref_date=ref_date)
    _log(f"[3/6] Afastados vigentes (pós {ref_date}): {len(df_afastados):,} registros deduplicados.")

    # 4. Carregamento de Demitidos
    _prog(0.44, "Processando Colaboradores Demitidos...")
    dem_path = raw_map.get("demitidos")
    if not dem_path or not os.path.exists(dem_path):
        raise FileNotFoundError("Arquivo COLA_Demitidos.CSV não encontrado.")
    df_demitidos = load_demitidos(dem_path)
    _log(f"[4/6] Demitidos válidos (sem Transferência s/ ônus): {len(df_demitidos):,} registros.")

    # 5. Carregamento de Férias
    _prog(0.55, f"Filtrando períodos de Férias da competência {month:02d}/{year}...")
    fer_path = raw_map.get("ferias")
    if not fer_path or not os.path.exists(fer_path):
        raise FileNotFoundError("Arquivo COLA_Ferias_Colaboradores.CSV não encontrado.")
    df_ferias = load_ferias(fer_path, year=year, month=month)
    _log(f"[5/6] Férias com gozo em {month:02d}/{year}: {len(df_ferias):,} colaboradores.")

    # 6. Carregamento de Valores Importados (Evento 29613)
    _prog(0.65, "Carregando lote de Valores Importados (Evento 29613)...")
    vi_path = raw_map.get("valores_importados")
    if not vi_path or not os.path.exists(vi_path):
        raise FileNotFoundError("Base de Valores Importados (Evento 29613) não localizada.")
    df_vi_raw = load_valores_importados(vi_path)
    _log(f"[6/6] Valores Importados (Evt 29613): {len(df_vi_raw):,} lançamentos carregados.")

    # 7. Execução do Motor de Regras de Negócio
    _prog(0.76, "Aplicando Regras de Negócio Good Card (Margem 20%, 6280, 29621, 5% Giga)...")
    data_dict = {
        "fopa": df_fopa,
        "nomina": df_nomina,
        "afastados": df_afastados,
        "demitidos": df_demitidos,
        "ferias": df_ferias,
        "valores_importados": df_vi_raw,
    }
    df_conferencia, df_valores_importados, kpis = execute_audit(data_dict)

    total_auditados = kpis.get("total_auditados", len(df_conferencia))
    com_saldo = kpis.get("com_saldo_devedor", 0)
    conformes = kpis.get("conformes", 0)
    divergencias = kpis.get("divergencias", 0)
    taxa_conf = kpis.get("taxa_conformidade", 100.0)

    _log("-" * 68)
    _log("RESULTADO DA CONCILIAÇÃO GOOD CARD:")
    _log(f"  • Total Auditado na Folha:     {total_auditados:,} colaboradores")
    _log(f"  • Com Saldo Devedor Apurado:   {com_saldo:,} colaboradores")
    _log(f"  • Registros Conformes (OK):    {conformes:,} ({taxa_conf:.2f}%)")
    _log(f"  • Divergências para Análise:   {divergencias:,}")
    _log("-" * 68)

    # 8. Construção e Gravação do Workbook Excel Estático
    _prog(0.88, "Construindo planilha Excel formatada (4 abas estáticas)...")
    wb = create_audit_workbook(
        df_conferencia=df_conferencia,
        df_vi=df_valores_importados,
        kpis=kpis,
    )

    out_dir_path = Path(output_dir)
    out_dir_path.mkdir(parents=True, exist_ok=True)
    target_file = out_dir_path / "Conferência Good Card_Gerado.xlsx"

    _prog(0.95, "Gravando arquivo Excel no diretório de destino...")
    try:
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
                        target_file = out_dir_path / f"Conferência Good Card_Gerado_{ts}.xlsx"
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
    _prog(1.0, "Auditoria Good Card concluída com sucesso!")

    return str(target_file)
