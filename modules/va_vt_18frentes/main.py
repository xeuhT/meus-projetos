"""
Orquestrador Principal — Conferência de Valores Pré-Folha VT (Eventos 5750, 31935 e 6292).
Competência: 09/2026.
"""

import sys
import time
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
if str(SCRIPT_DIR.parent) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR.parent))

try:
    from business_rules import BusinessRulesPrefolha
    from config import COMPETENCIA, OUTPUT_FILE
    from data_loader import DataLoader
    from excel_exporter import ExcelExporterPrefolha
except ImportError:
    from scripts.business_rules import BusinessRulesPrefolha
    from scripts.config import COMPETENCIA, OUTPUT_FILE
    from scripts.data_loader import DataLoader
    from scripts.excel_exporter import ExcelExporterPrefolha


def main() -> None:
    start_ts = time.perf_counter()
    print("=" * 80)
    print(f" INICIANDO AUDITORIA DE VALORES PRÉ-FOLHA VT (5750 / 31935 / 6292) - {COMPETENCIA}")
    print("=" * 80)

    loader = DataLoader()
    raw_data = loader.load_all()

    rules = BusinessRulesPrefolha()
    result = rules.process_conferencia(raw_data)

    exporter = ExcelExporterPrefolha(OUTPUT_FILE)
    saved_path = exporter.export(result)

    elapsed = time.perf_counter() - start_ts
    kpis = result["kpis"]

    print("-" * 80)
    print(f" [SUCESSO] Relatório salvo em: {saved_path}")
    print(f" [TEMPO] Execução concluída em {elapsed:.2f}s")
    print(f" [KPI] Total FOPA Pré-Folha: {kpis['total_fopa']:,} | Conciliados: {kpis['conciliados_fopa']:,}")
    print(f" [KPI] Soma Eventos Pagamento: R$ {kpis['soma_pgtos']:,.2f} | Evento 5750: R$ {kpis['soma_5750']:,.2f}")
    print(f" [KPI] Divergências Consolidadas (FOPA + 31935 + 6292): {kpis['total_divergencias_consolidadas']}")
    print("=" * 80)


if __name__ == "__main__":
    main()
