"""
config.py - Mapeamento de colunas das operadoras + arquivos RH + Movimentações.
"""

# ---------------------------------------------------------------------------
# BASES DE ATIVOS (Faturamento / Listagem das Operadoras)
# ---------------------------------------------------------------------------
ATIVOS = {
    "BRADESCO_ODONTO": {
        "aba": "BASE DE ATIVOS - ODONTO", "header": 0,
        "colunas": {
            "plano": "PLANO", "matricula": "MATRICULA", "cpf": "CPF",
            "nascimento": "NASCIMENTO", "titular": "TITULAR",
            "beneficiario": "BENEFICIÁRIO", "carteira": "CARTEIRA",
            "relacao_dependencia": "RELAÇÃO DE DEPENDÊNCIA",
            "posicao_cadastral": "POSIÇÃO CADASTRAL",
            "inicio_vigencia": "INICIO DA VIGENCIA",
            "cancelamento": "CANCELAMENTO", "loja": "LOJA",
        },
    },
    "GNDI": {
        "aba": "BASE DE ATIVOS", "header": 0,
        "colunas": {
            "plano": "PLANO", "matricula": "MATRICULA", "cpf": "CPF",
            "nascimento": "NASCIMENTO", "titular": "TITULAR",
            "beneficiario": "BENEFICIÁRIO", "carteira": "CARTEIRA",
            "relacao_dependencia": "RELAÇÃO DE DEPENDÊNCIA",
            "posicao_cadastral": "POSIÇÃO CADASTRAL",
            "inicio_vigencia": "INICIO DA VIGENCIA",
            "cancelamento": "CANCELAMENTO", "loja": "LOJA ORGANICO",
        },
    },
    "HAPVIDA": {
        "aba": "BASE DE ATIVOS", "header": 0,
        "colunas": {
            "plano": "PLANO", "matricula": "MATRICULA ", "cpf": "CPF",
            "nascimento": "NASCIMENTO", "titular": "TITULAR",
            "beneficiario": "BENEFICIÁRIO", "carteira": "CARTEIRA",
            "relacao_dependencia": "RELAÇÃO DE DEPENDÊNCIA",
            "posicao_cadastral": "POSIÇÃO CADASTRAL",
            "inicio_vigencia": "INICIO DA VIGENCIA",
            "cancelamento": "CANCELAMENTO", "loja": " ORGANICO",
        },
    },
    "KLINI": {
        "aba": "BASE DE ATIVOS", "header": 0,
        "colunas": {
            "plano": "PLANO", "matricula": "MATRICULA", "cpf": "CPF ",
            "nascimento": "NASCIMENTO", "titular": "TITULAR",
            "beneficiario": "BENEFICIÁRIO", "carteira": "CARTEIRA",
            "relacao_dependencia": "RELAÇÃO DE DEPENDÊNCIA",
            "posicao_cadastral": "POSIÇÃO CADASTRAL",
            "inicio_vigencia": "INICIO DA VIGENCIA",
            "cancelamento": "CANCELAMENTO", "loja": "LOJA",
        },
    },
    "SULAMERICA_ODONTO": {
        "aba": "BASE DE ATIVOS", "header": 0,
        "colunas": {
            "plano": "PLANO", "matricula": "MATRICULA", "cpf": "CPF",
            "nascimento": "NASCIMENTO", "titular": "TITULAR",
            "beneficiario": "BENEFICIÁRIO", "carteira": "CARTEIRA",
            "relacao_dependencia": "RELAÇÃO DE DEPENDÊNCIA",
            "posicao_cadastral": "POSIÇÃO CADASTRAL",
            "inicio_vigencia": "INICIO DA VIGENCIA",
            "cancelamento": "CANCELAMENTO", "loja": "LOJA",
        },
    },
    "SULAMERICA_SAUDE": {
        "aba": "BASE DE ATIVOS", "header": 0,
        "colunas": {
            "plano": "PLANO", "matricula": "MATRICULA", "cpf": "CPF",
            "nascimento": "DATA DE NASCIMENTO", "titular": "TITULAR",
            "beneficiario": "BENEFICIÁRIO", "carteira": "CARTEIRA",
            "relacao_dependencia": "RELAÇÃO DE DEPENDÊNCIA",
            "posicao_cadastral": "POSIÇÃO CADASTRAL",
            "inicio_vigencia": "INICIO DA VIGENCIA",
            "cancelamento": "CANCELAMENTO", "loja": "LOJA",
        },
    },
    "UNIMED": {
        "aba": "BASE DE ATIVOS", "header": 0,
        "colunas": {
            "plano": "PLANO", "matricula": "MATRICULA", "cpf": "CPF ",
            "nascimento": "NASCIMENTO", "titular": "TITULAR",
            "beneficiario": "BENEFICIÁRIO", "carteira": "CARTEIRA",
            "relacao_dependencia": "RELAÇÃO DE DEPENDÊNCIA",
            "posicao_cadastral": "POSIÇÃO CADASTRAL",
            "inicio_vigencia": "INICIO DA VIGENCIA",
            "cancelamento": "CANCELAMENTO", "loja": "LOJA ORGANICO",
        },
    },
}

# ---------------------------------------------------------------------------
# FOPAG (Movimentações e Descontos - Abas corretas por operadora)
# ---------------------------------------------------------------------------
FOPAG = {
    "BRADESCO_ODONTO": {
        "aba": "Base Odonto", "header": 0,
        "colunas": {
            "categoria_plano": "Categoria do Plano", "bandeira": "Bandeira",
            "matricula": "Matricula", "cpf": "CPF", "titular": "Titular",
            "beneficiario": "Beneficiário", "carteira": "Carteira",
            "relacao_dependencia": "Relação de Dependência",
            "posicao_cadastral": "Posição Cadastral",
            "inicio_vigencia": "Inicio de Vigência", "cancelamento": " Cancelamento",
            "loja": "Loja",
        },
    },
    "GNDI": {
        "aba": "Base Saúde", "header": 0,
        "colunas": {
            "categoria_plano": "Categoria do Plano", "bandeira": "Bandeira",
            "matricula": "Matricula", "cpf": "CPF", "titular": "Titular",
            "beneficiario": "Beneficiário", "carteira": "Carteira",
            "relacao_dependencia": "Relação de Dependência",
            "posicao_cadastral": "Posição Cadastral",
            "inicio_vigencia": "Inicio de Vigência", "cancelamento": " Cancelamento",
            "loja": "Loja",
        },
    },
    "HAPVIDA": {
        "aba": "Base Saúde", "header": 0,
        "colunas": {
            "categoria_plano": "Categoria do Plano", "bandeira": "Bandeira",
            "cpf": "CPF", "beneficiario": "Beneficiário",
            "inicio_vigencia": "Inicio de Vigência", "cancelamento": " Cancelamento",
            "loja": "Loja",
        },
    },
    "KLINI": {
        "aba": "BASE  FOPAG", "header": 0,
        "colunas": {
            "categoria_plano": "Categoria do Plano", "bandeira": "Bandeira",
            "matricula": "Matricula", "cpf": "CPF", "titular": "Titular",
            "beneficiario": "Beneficiário", "carteira": "Carteirinha",
            "relacao_dependencia": "Relação de Dependência",
            "posicao_cadastral": "Posição cadastral",
            "inicio_vigencia": "Inicio de Vigência", "cancelamento": " Cancelamento",
            "loja": "Loja",
        },
    },
    "SULAMERICA_ODONTO": {
        "aba": "BASE ODONTO", "header": 0,
        "colunas": {
            "categoria_plano": "Categoria do Plano", "bandeira": "Bandeira",
            "matricula": "Matricula", "cpf": "CPF", "titular": "Tittular",
            "beneficiario": "Beneficiário", "carteira": "Carteira",
            "relacao_dependencia": "Relação de Dependência",
            "posicao_cadastral": "Posição Cadastral",
            "inicio_vigencia": "Inicio de Vigência", "cancelamento": "Fim de vigência",
            "loja": "Loja",
        },
    },
    "SULAMERICA_SAUDE": {
        "aba": "Base Saúde", "header": 0,
        "colunas": {
            "categoria_plano": "Categoria do Plano", "bandeira": "Bandeira",
            "matricula": "Matricula", "cpf": "CPF", "titular": "Titular",
            "beneficiario": "Beneficiário", "carteira": "Carteira",
            "relacao_dependencia": "Relação de Dependência",
            "posicao_cadastral": "Posição Cadastral",
            "inicio_vigencia": "Inicio de Vigência", "cancelamento": " Cancelamento",
            "loja": "Loja",
        },
    },
    "UNIMED": {
        "aba": "Base Saúde", "header": 0,
        "colunas": {
            "categoria_plano": "Categoria do Plano", "bandeira": "Bandeira",
            "matricula": "Matricula", "cpf": "CPF", "titular": "Titular",
            "beneficiario": "Beneficiário", "carteira": "Carteira",
            "relacao_dependencia": "Relação de Dependência",
            "posicao_cadastral": "Posição Cadastral",
            "inicio_vigencia": "Inicio de Vigência", "cancelamento": " Cancelamento",
            "loja": "Loja",
        },
    },
}

# ---------------------------------------------------------------------------
# COLABORADORES (cadastro mestre do RH)
# ---------------------------------------------------------------------------
COLABORADORES = {
    "aba": 0,
    "header": 1,
    "colunas": {
        "codigo_empresa": "Código da Empresa",
        "matricula": "Matrícula",
        "nome": "Colaborador",
        "cargo": "Cargo",
        "estabelecimento": "Estabelecimento",
        "admissao": "Data Admissão",
        "cpf": "CPF/CNPJ",
        "nascimento": "Data Nascimento",
        "sindicato": "Sindicato",
        "codigo_sindicato": "Código Sindicato",
        "situacao": "Situação",
        "tipo_situacao": "Tipo de Situação",
        "razao_social": "Razão Social",
        "natureza_profissional": "Natureza Profissional",
        "codigo_estabelecimento": "Código Estabelecimento",
        "data_rescisao": "Data de Rescisão",
    },
}

# ---------------------------------------------------------------------------
# DEMITIDOS
# ---------------------------------------------------------------------------
DEMITIDOS = {
    "header": 0,
    "encoding": "latin-1",
    "separador": ";",
    "colunas": {
        "matricula": "Matrícula",
        "nome": "Colaborador",
        "cpf": "CPF",
        "admissao": "Data Admissão",
        "data_rescisao": "Data Rescisão",
        "tipo_rescisao": "Tipo de Rescisão",
        "motivo": "Motivo do Desligamento",
        "cargo": "Cargo",
    },
}

# ---------------------------------------------------------------------------
# DEPENDENTES (Base Marsh - CSV)
# ---------------------------------------------------------------------------
DEPENDENTES = {
    "header": 0,
    "encoding": "utf-8",
    "separador": ";",
    "colunas_indices": {
        6: "matricula",             # MATRÍCULA COLABORADOR
        7: "titular_marsh",         # COLABORADOR (Nome do Titular)
        9: "inicio_vigencia_marsh", # INICIO DA VIGÊNCIA ATUAL
        12: "beneficiario",         # NOME DEPENDENTE
        13: "cpf_dependente",       # CPF DEPENDENTE
        14: "nascimento_marsh",     # DATA NASCIMENTO
        17: "parentesco_marsh",     # PARENTESCO
        18: "dependencia_marsh",    # DEPENDÊNCIA
        24: "grau_instrucao_marsh", # CODIGO GRAU DE INSTRUÇÃO
        25: "desc_instrucao_marsh", # DESCRIÇÃO GRAU DE INSTRUÇÃO
        26: "invalido_marsh",       # DEPENDENTE INVÁLIDO
    }
}