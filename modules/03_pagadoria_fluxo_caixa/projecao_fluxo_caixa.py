"""
Motor de Projeção de Fluxo de Caixa de Recursos Humanos
Cencosud Brasil — Integração DP / Tesouraria Corporativa.

Objetivo:
- Modelar e projetar os desembolsos futuros de folha de pagamento, benefícios,
  rescisões e encargos tributários (DCTFWeb / FGTS Digital);
- Gerar a curva de liquidez diária consolidada por CNPJ/Empresa para apoio à Tesouraria.
"""

import os
import sys
import argparse
import logging
from datetime import datetime, timedelta
from typing import Dict, List, Any
import pandas as pd
import numpy as np

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("FluxoCaixa_RH")


def projetar_desembolsos_mes(competencia_ano_mes: str) -> pd.DataFrame:
    """
    Gera a projeção cronológica de desembolsos obrigatórios e contratuais do RH.
    """
    logger.info(f"Calculando projeção de desembolsos para a competência: {competencia_ano_mes}")
    ano, mes = map(int, competencia_ano_mes.split("-"))
    
    # Calendário financeiro padrão corporativo
    eventos_fluxo = [
        {"dia": 5, "categoria": "FOLHA", "descricao": "Pagamento Mensal de Salários (Lojas e CDs)", "frequencia": "Mensal"},
        {"dia": 10, "categoria": "RESCISAO", "descricao": "Janela Quinquenal de Rescisões / TRCTs", "frequencia": "Recorrente"},
        {"dia": 15, "categoria": "ENCARGOS", "descricao": "FGTS Digital / Guias Rescisórias", "frequencia": "Mensal"},
        {"dia": 20, "categoria": "ENCARGOS", "descricao": "DCTFWeb / Previdenciário (INSS e IRRF)", "frequencia": "Mensal"},
        {"dia": 20, "categoria": "ADIANTAMENTO", "descricao": "Adiantamento Quinzenal de Salários (Vale)", "frequencia": "Mensal"},
        {"dia": 25, "categoria": "BENEFICIOS", "descricao": "Faturas de Planos de Saúde e Odonto", "frequencia": "Mensal"},
        {"dia": 28, "categoria": "BENEFICIOS", "descricao": "Crédito Cartão Good Card / Alimentação", "frequencia": "Mensal"},
        {"dia": 30, "categoria": "BENEFICIOS", "descricao": "Compra e Carga de Vale Transporte", "frequencia": "Mensal"}
    ]

    registros = []
    for ev in eventos_fluxo:
        try:
            data_prevista = datetime(ano, mes, ev["dia"])
        except ValueError:
            # Caso o mês não tenha o dia 30 (ex: fevereiro)
            data_prevista = datetime(ano, mes, 28)

        # Ajuste para dia útil se cair em sábado ou domingo
        if data_prevista.weekday() == 5:  # Sábado -> Sexta
            data_prevista -= timedelta(days=1)
        elif data_prevista.weekday() == 6:  # Domingo -> Sexta
            data_prevista -= timedelta(days=2)

        registros.append({
            "competencia": competencia_ano_mes,
            "data_desembolso": data_prevista.strftime("%Y-%m-%d"),
            "dia_semana": data_prevista.strftime("%A"),
            "categoria": ev["categoria"],
            "descricao_evento": ev["descricao"],
            "previsao_status": "PROJETADO",
            "moeda": "BRL"
        })

    df_fluxo = pd.DataFrame(registros).sort_values("data_desembolso").reset_index(drop=True)
    return df_fluxo


def exportar_fluxo_caixa(df_projecao: pd.DataFrame, caminho_saida: str) -> None:
    """
    Exporta a planilha de fluxo de caixa projetado para compartilhamento com a Tesouraria.
    """
    os.makedirs(os.path.dirname(os.path.abspath(caminho_saida)), exist_ok=True)
    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        df_projecao.to_excel(writer, sheet_name="Projecao_Diaria", index=False)
        
        # Resumo por categoria
        resumo = df_projecao.groupby("categoria").size().reset_index(name="qtd_janelas")
        resumo.to_excel(writer, sheet_name="Resumo_Categorias", index=False)

    logger.info(f"Projeção de fluxo de caixa exportada com sucesso em: {caminho_saida}")


def main():
    parser = argparse.ArgumentParser(description="Projeção Preditiva de Fluxo de Caixa de RH")
    parser.add_argument("--competencia", type=str, default=datetime.now().strftime("%Y-%m"), help="Competência AAAA-MM")
    parser.add_argument("--saida", type=str, default="fluxo_caixa_rh_projetado.xlsx", help="Arquivo Excel de saída")
    args = parser.parse_args()

    df = projetar_desembolsos_mes(args.competencia)
    exportar_fluxo_caixa(df, args.saida)


if __name__ == "__main__":
    main()
