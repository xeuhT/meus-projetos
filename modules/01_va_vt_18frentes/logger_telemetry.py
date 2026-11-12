"""
Módulo de Telemetria e Registro Técnico da Auditoria.
"""

import os
import time
from datetime import datetime

class AuditLogger:
    def __init__(self, log_dir=None):
        if log_dir is None:
            log_dir = os.path.join(os.path.dirname(__file__), "logs")
        os.makedirs(log_dir, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = os.path.join(log_dir, f"auditoria_vt_{timestamp}.log")
        self.start_time = time.time()
        self.step_times = {}
        self.counts = {}
        
        self.log("="*70)
        self.log("INÍCIO DO PROCESSAMENTO - AUDITORIA DE VT (6%)")
        self.log(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        self.log("="*70)

    def log(self, message):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")

    def record_step(self, step_name):
        elapsed = time.time() - self.start_time
        self.step_times[step_name] = elapsed
        self.log(f"ETAPA CONCLUÍDA: {step_name} (Tempo acumulado: {elapsed:.2f}s)")

    def record_file_stats(self, file_name, row_count, col_count):
        self.counts[file_name] = {"rows": row_count, "cols": col_count}
        self.log(f"CARGA: {file_name} -> {row_count:,} linhas | {col_count} colunas")

    def finalize(self, summary_stats):
        total_time = time.time() - self.start_time
        self.log("\n" + "="*70)
        self.log("SUMÁRIO TÉCNICO E TELEMETRIA DA AUDITORIA")
        self.log("="*70)
        self.log(f"Tempo Total de Execução: {total_time:.2f} segundos")
        self.log("\nBases de Dados Processadas:")
        for fname, stat in self.counts.items():
            self.log(f"  - {fname}: {stat['rows']:,} linhas")
            
        self.log("\nResultados da Auditoria:")
        for k, v in summary_stats.items():
            self.log(f"  - {k}: {v:,}")
            
        self.log("="*70)
        self.log(f"Log salvo em: {self.log_file}")
        self.log("="*70)
