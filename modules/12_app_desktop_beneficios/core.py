"""
core.py - Ponto de entrada da automação V5.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import logging
from pathlib import Path
from datetime import datetime

import leitores
import motor_auditoria
import exportador

def configurar_log(pasta_log: Path):
    pasta_log.mkdir(parents=True, exist_ok=True)
    arquivo_log = pasta_log / f"auditoria_{datetime.now():%Y%m%d_%H%M%S}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[
            logging.FileHandler(arquivo_log, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return arquivo_log

import unicodedata
import re

def clean_name(s):
    s = s.replace('�', ' ')
    s = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('utf-8').lower()
    return re.sub(r'[^a-z0-9]+', ' ', s)

PADROES_ATIVOS = {
    "bradesco": "BRADESCO_ODONTO",
    "gndi": "GNDI",
    "hapvida": "HAPVIDA",
    "klini": "KLINI",
    "sul america odonto": "SULAMERICA_ODONTO",
    "sulamerica odonto": "SULAMERICA_ODONTO",
    "sul america sa": "SULAMERICA_SAUDE",
    "sulamerica sa": "SULAMERICA_SAUDE",
    "sas saude": "SULAMERICA_SAUDE",
    "unimed": "UNIMED",
}

def _detectar_operadora(nome_arquivo: str) -> str | None:
    nome = clean_name(nome_arquivo)
    for padrao in sorted(PADROES_ATIVOS, key=len, reverse=True):
        if padrao in nome:
            return PADROES_ATIVOS[padrao]
    return None

def descobrir_arquivos(pasta):

    achados = {"ativos": {}, "fopag": {}, "colaboradores": None, "demitidos": None, "dependentes": None}
    for caminho in pasta.iterdir():
        if not caminho.is_file():
            continue
        nome = clean_name(caminho.name)

        # 🚨 BLINDAGEM: Ignora os arquivos financeiros na leitura padrão para não travar
        if "evento" in nome or "fopa eventos" in nome or "fopa_eventos" in nome:
            continue

        if "demitidos" in nome or "desligados" in nome:
            achados["demitidos"] = caminho
        elif "dependentes marsh" in nome or "dependentes" in nome:
            achados["dependentes"] = caminho
        elif "cola colaboradores" in nome or "cola_colaboradores" in nome or "colaboradores" in nome or "nomina ativos" in nome or ("nomina" in nome and not _detectar_operadora(nome)):
            achados["colaboradores"] = caminho
        elif "movimenta" in nome or "fopag" in nome:
            op = _detectar_operadora(nome)
            if op:
                achados["fopag"][op] = caminho
        elif "listagem de ativos" in nome or "ativos" in nome:
            op = _detectar_operadora(nome)
            if op:
                achados["ativos"][op] = caminho
    return achados

def executar(pasta_entrada: Path, caminho_saida: Path):
    log = logging.getLogger(__name__)
    log.info("Pasta de entrada: %s", pasta_entrada)
    arquivos = descobrir_arquivos(pasta_entrada)

    if not arquivos["ativos"]:
        log.error("Nenhuma base de ativos encontrada. Abortando.")
        return

    ativos_por_op, fopag_por_op = {}, {}
    for op, caminho in arquivos["ativos"].items():
        ativos_por_op[op] = leitores.ler_ativos(caminho, op)
    for op, caminho in arquivos["fopag"].items():
        fopag_por_op[op] = leitores.ler_fopag(caminho, op)

    import pandas as pd
    colaboradores = leitores.ler_colaboradores(arquivos["colaboradores"]) if arquivos["colaboradores"] else pd.DataFrame()
    demitidos = leitores.ler_demitidos(arquivos["demitidos"]) if arquivos["demitidos"] else pd.DataFrame()
    dependentes = leitores.ler_dependentes(arquivos["dependentes"]) if arquivos["dependentes"] else pd.DataFrame()

    # 1. Roda o Motor de Cruzamento Geral
    consolidado = motor_auditoria.auditar_todas(
        ativos_por_op, fopag_por_op, colaboradores, demitidos, dependentes
    )

    # 2. Ingestão das Bases Financeiras
    log.info("Iniciando varredura financeira e faturas...")

    # Dicionário exige o arquivo EXATO
    caminho_dicio = pasta_entrada / "Eventos 1.xlsx"
    df_eventos_dicionario = leitores.ler_dicionario_eventos(caminho_dicio)

    # FOPAG Financeira exige APENAS A PASTA
    df_fopag_financeira = leitores.ler_financeiro_fopag(pasta_entrada, prefixo="FOPA_Eventos_por_Colaborador")

    # === CONEXÃO NOVA: LEITURA DAS FATURAS ===
    df_faturas = leitores.ler_faturas_operadoras(pasta_entrada)

    # 3. Exportação ÚNICA final com faturas incluídas (Variável df_faturas adicionada)
    exportador.exportar(consolidado, caminho_saida, df_fopag_financeira, df_eventos_dicionario, df_faturas)

    log.info("✅ Concluído. Relatório gerado com sucesso!")

if __name__ == "__main__":
    PASTA_ENTRADA = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\planilhas beneficios")
    PASTA_SAIDA = PASTA_ENTRADA / f"Auditoria_Consolidada_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    configurar_log(PASTA_ENTRADA / "logs")
    executar(PASTA_ENTRADA, PASTA_SAIDA)