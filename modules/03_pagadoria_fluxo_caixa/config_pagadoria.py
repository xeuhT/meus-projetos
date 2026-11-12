"""
Configurações Corporativas - Módulo Pagadoria & Fluxo de Caixa RH (Cencosud Brasil)
Mapeamento de empresas SAP, centros de custo, alçadas de aprovação e layouts.
"""

import os
from typing import Dict, Any

# Variáveis de ambiente corporativas
SAP_MANDANTE = os.getenv("SAP_MANDANTE", "100")
SAP_USUARIO_CARGA = os.getenv("SAP_USUARIO_CARGA", "S_AUTOMACAO_RH")

# Tabela de Alçadas de Aprovação de Pagamentos de RH (Governança Corporativa)
ALCADAS_APROVACAO = {
    "NIVEL_1_COORDENACAO": {"limite_maximo": 50000.00, "cargo": "Coordenador de RH"},
    "NIVEL_2_GERENCIA": {"limite_maximo": 250000.00, "cargo": "Gerente de RH / DP"},
    "NIVEL_3_DIRETORIA": {"limite_maximo": 1000000.00, "cargo": "Diretor de Recursos Humanos"},
    "NIVEL_4_VICE_PRESIDENCIA": {"limite_maximo": float("inf"), "cargo": "Vice-Presidente / CFO"}
}

# Códigos de Empresas Cencosud no SAP
EMPRESAS_SAP = {
    "BRETAS": "BR01",
    "G_BARBOSA": "BR02",
    "PREZUNIC": "BR03",
    "PERINI": "BR04",
    "MERCANTIL": "BR05",
    "SPID": "BR06"
}

# Layout de Colunas Esperadas na Planilha de Pagamentos
COLUNAS_OBRIGATORIAS_APROVACAO = [
    "empresa",
    "documento",
    "tipo_pagamento",
    "favorecido",
    "cnpj_cpf_favorecido",
    "valor_bruto",
    "data_vencimento",
    "centro_custo",
    "conta_contabil",
    "aprovador_1",
    "status_aprovacao"
]
