"""
Orquestrador Principal da Conferência de Seguro de Vida - Cencosud Brasil.
Suporta execução tanto via Interface Gráfica Corporativa (GUI) quanto via linha de comando (CLI):
1. Ingestão e higienização dos dados (Nomina + Aba Com Gestor, Oposição, FOPA, Optantes Complementares e Fatura).
2. Processamento das regras de elegibilidade, cortes etários/admissão, situações especiais (Afastados, Férias, Aprendizes, Rescisões).
3. Segregação e montagem auditada de Cargas Seguro Normal e Complementar (12x e 24x).
4. Conciliação Trilateral (Nomina x FOPA x Fatura Seguradora).
5. Geração dos arquivos operacionais de carga por bandeira com layout específico + pacote ZIP.
6. Geração do Relatório Master de Conferência e Auditoria Técnica no padrão institucional.
"""

import sys
import os
import time
from pathlib import Path
from typing import Optional, Dict, Any, Callable

CURRENT_DIR = Path(__file__).resolve().parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))
if str(CURRENT_DIR.parent) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR.parent))

try:
    from config import BASE_DIR, RAW_DIR, OUTPUT_DIR, AUDIT_REPORT_DIR
    from data_loader import DataLoader
    from business_rules import BusinessRulesEngine
    from excel_exporter import ExcelExporter
except ImportError:
    from conferencia_seguro_vida.config import BASE_DIR, RAW_DIR, OUTPUT_DIR, AUDIT_REPORT_DIR
    from conferencia_seguro_vida.data_loader import DataLoader
    from conferencia_seguro_vida.business_rules import BusinessRulesEngine
    from conferencia_seguro_vida.excel_exporter import ExcelExporter


def run_conferencia(
    competencia: str = "102026",
    raw_dir: Optional[Path] = None,
    dest_dir: Optional[Path] = None,
    raw_map: Optional[Dict[str, Optional[str]]] = None,
    progress_cb: Optional[Callable[[float, str], None]] = None,
    log_cb: Optional[Callable[[str], None]] = None,
    prompt_lock_cb: Optional[Callable[[str], str]] = None,
) -> Dict[str, Any]:
    """Executa o pipeline completo de conferência e geração de arquivos de seguro de vida."""
    start_time = time.time()

    def _emit_log(msg: str) -> None:
        print(msg)
        if log_cb:
            log_cb(msg)

    def _emit_prog(pct: float, msg: str) -> None:
        if progress_cb:
            progress_cb(pct, msg)

    _emit_log("=" * 80)
    _emit_log(f"[*] INICIANDO CONFERÊNCIA DE SEGURO DE VIDA — COMPETÊNCIA {competencia}")
    _emit_log("=" * 80)

    # Resolve diretórios de saída (personalizados pela GUI ou padrão de config)
    if dest_dir is not None:
        dest_root = Path(dest_dir)
        out_cargas_dir = dest_root / "Carga Seguro de Vida Gerada"
        out_audit_dir = dest_root / "Relatórios de Conferência"
        out_zip_dir = dest_root
    else:
        out_cargas_dir = OUTPUT_DIR
        out_audit_dir = AUDIT_REPORT_DIR
        out_zip_dir = BASE_DIR

    # 1. Carga de Dados
    _emit_log("\n[Etapa 1/4] Carregando bases de dados oficiais (Nomina + Com Gestor, Oposição, FOPA e Fatura)...")
    loader = DataLoader(
        raw_dir=Path(raw_dir) if raw_dir else None,
        raw_map=raw_map,
        log_cb=log_cb,
        progress_cb=progress_cb,
    )
    raw_data = loader.load_all()

    # 2. Motor de Regras, Situações Especiais e Conciliação
    _emit_log("\n[Etapa 2/4] Aplicando regras de elegibilidade, cortes etários, Diretores e conciliação trilateral...")
    _emit_prog(0.36, "Aplicando regras de elegibilidade, Diretores e situações especiais...")
    engine = BusinessRulesEngine(raw_data)
    results = engine.run_all()
    _emit_prog(0.43, "Regras de negócio e conciliação trilateral concluídas.")

    # 3. Exportação dos Arquivos Excel e Pacote ZIP
    _emit_log("\n[Etapa 3/4] Exportando cargas operacionais por bandeira e relatório master de auditoria...")
    exporter = ExcelExporter(
        competencia=competencia,
        output_dir=out_cargas_dir,
        audit_dir=out_audit_dir,
        zip_dir=out_zip_dir,
        log_cb=log_cb,
        progress_cb=progress_cb,
        prompt_lock_cb=prompt_lock_cb,
    )
    arquivos_gerados = exporter.export_all(results)

    # 4. Apresentação do Resumo Executivo
    elapsed = time.time() - start_time
    kpis = results["kpis"]
    _emit_prog(1.0, "Conferência e cargas finalizadas com sucesso!")

    _emit_log("\n" + "=" * 80)
    _emit_log(f"[+] PIPELINE CONCLUÍDO COM SUCESSO EM {elapsed:.1f} SEGUNDOS")
    _emit_log("=" * 80)
    _emit_log(f"[*] Total de Colaboradores na Nomina (+ Aba Com Gestor): {kpis['total_nomina']:,}")
    _emit_log(f"[*] Total de Vidas Elegíveis Seguro Normal: {kpis['total_elegiveis']:,}")
    _emit_log(f"[*] Total de Diretores Elegíveis (Extraídos da Aba Com Gestor): {kpis['total_diretores_elegiveis']:,}")
    _emit_log(f"[*] Total de Exclusões por Regra: {kpis['total_excluidos']:,}")
    for mot, qtd in kpis["exclusoes_por_motivo"].items():
        if qtd > 0:
            _emit_log(f"    - {mot}: {qtd:,}")
    _emit_log("-" * 80)
    _emit_log("[*] Seguro Complementar:")
    _emit_log(f"    - Complementar 12x (Plano 05 / Dir Plano 08): {kpis['total_comp_12x']} vidas")
    _emit_log(f"    - Complementar 24x (Plano 06 / Dir Plano 09): {kpis['total_comp_24x']} vidas")
    _emit_log("-" * 80)
    _emit_log(f"[*] Arquivos de Carga gerados em: {out_cargas_dir}")
    _emit_log(f"[*] Relatório de Auditoria gerado em: {out_audit_dir}")
    _emit_log("=" * 80)

    return {
        "status": "SUCESSO",
        "tempo_execucao": elapsed,
        "kpis": kpis,
        "arquivos_gerados": arquivos_gerados,
        "output_dir": out_cargas_dir,
        "audit_dir": out_audit_dir,
    }


if __name__ == "__main__":
    run_conferencia()
