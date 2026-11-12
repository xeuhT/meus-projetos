"""
Módulo Integrador da Auditoria de Férias e Folha Cencosud
Delega as execuções diretamente para os motores oficiais homologados:
- Aba 1 (Formatada): gerar_conferencia_ferias_formatado_v4.py
- Aba 2 (Cálculo): atualizar_todas_auditorias.py (run_auditoria_calculo)
- Aba 3 (Folha): atualizar_todas_auditorias.py (run_auditoria_folha)
"""

import os
import sys

from gerar_conferencia_ferias_formatado_v4 import run_conferencia_formatada
from atualizar_todas_auditorias import run_auditoria_folha, run_auditoria_calculo

def exec_conferencia_formatada(raw_map, output_dir, progress_cb=None, log_cb=None, prompt_lock_cb=None):
    """Executa a geração da Conferência Formatada utilizando o motor oficial v4 homologado."""
    return run_conferencia_formatada(
        output_dir=output_dir,
        file_map=raw_map,
        progress_cb=progress_cb,
        log_cb=log_cb,
        prompt_lock_cb=prompt_lock_cb
    )

def exec_auditoria_calculo(raw_map, output_dir, progress_cb=None, log_cb=None, prompt_lock_cb=None):
    """Executa a Auditoria do Cálculo de Férias utilizando o motor oficial homologado."""
    return run_auditoria_calculo(
        output_dir=output_dir,
        file_map=raw_map,
        progress_cb=progress_cb,
        log_cb=log_cb,
        prompt_lock_cb=prompt_lock_cb
    )

def exec_auditoria_folha(raw_map, output_dir, progress_cb=None, log_cb=None, prompt_lock_cb=None):
    """Executa a Auditoria da Folha Mensal utilizando o motor oficial homologado."""
    return run_auditoria_folha(
        output_dir=output_dir,
        file_map=raw_map,
        progress_cb=progress_cb,
        log_cb=log_cb,
        prompt_lock_cb=prompt_lock_cb
    )
