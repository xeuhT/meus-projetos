"""
Módulo de Configuração Central para a Conferência e Automação de Seguro de Vida - Cencosud Brasil.
Define diretórios dinâmicos (compatíveis com modo desenvolvimento e executável standalone .exe),
parâmetros de corte, eventos de folha, regras de planos e identidade visual institucional Cencosud.
"""

import os
import sys
from pathlib import Path


def resource_path(relative_path: str) -> Path:
    """
    Obtém o caminho absoluto para recursos embutidos (logos, ícones, JSONs),
    funcionando tanto em desenvolvimento quanto em executável PyInstaller (--onefile).
    """
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_path / relative_path


def get_app_root() -> Path:
    """Retorna o diretório onde o executável (.exe) ou os scripts estão localizados."""
    env_root = os.environ.get("CENCOSUD_APP_ROOT")
    if env_root:
        return Path(env_root).resolve()
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resolve_catalog_json(filename: str) -> Path:
    """
    Prioriza um arquivo JSON externo ao lado do executável (caso o RH deseje atualizar
    o catálogo sem recompilar) e faz fallback transparente para o JSON embutido no app.
    """
    external_candidates = [
        get_app_root() / filename,
        get_app_root() / "config" / filename,
    ]
    for cand in external_candidates:
        if cand.exists():
            return cand
    return resource_path(filename)


# Diretório Corporativo Padrão (com fallback inteligente para máquinas da equipe de RH)
_CORP_DEFAULT_DIR = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\seguro de vida")
BASE_DIR = _CORP_DEFAULT_DIR if _CORP_DEFAULT_DIR.exists() else get_app_root()

RAW_DIR = BASE_DIR / "Relatórios brutos"
CARGA_DIR = BASE_DIR / "Carga Seguro de Vida 092026"  # Base histórica de referência
OUTPUT_DIR = BASE_DIR / "Carga Seguro de Vida Gerada"
AUDIT_REPORT_DIR = BASE_DIR / "Relatórios de Conferência"

# Arquivos de Cadastro e Catálogo de Optantes
CADASTRO_DIRETORES_JSON = resolve_catalog_json("diretores_cadastro.json")
OPTANTES_COMPLEMENTAR_JSON = resolve_catalog_json("optantes_complementar.json")

# Eventos Oficiais da Folha de Pagamento (FOPA)
EVENTO_SEGURO_NORMAL = "6266"  # Seguro de Vida - Normal
EVENTO_SEGURO_COMP = "6267"    # Seguro Complementar

# Regras de Negócio de Elegibilidade
DATA_CORTE_ADESAO_IDADE = "2015-12-31"  # Admitidos a partir de 2016 com > 65 anos ficam fora
IDADE_LIMITE_ADESAO = 65

# Relação de Planos por Bandeira / Estabelecimento / Sindicato
PLANO_GBARBOSA = "PLANO 01"
PLANO_PREZUNIC = "PLANO 01"
PLANO_MERCANTIL = "PLANO 01"
PLANO_GIGA = "PLANO 01"
PLANO_PERINI = "PLANO 01"
PLANO_BRETAS_LOJAS = "PLANO 03"
PLANO_BRETAS_POSTOS = "PLANO 04"
PLANO_COMP_12X = "PLANO 05"
PLANO_COMP_24X = "PLANO 06"
PLANO_JAGUAQUARA = "PLANO 07"
PLANO_DIRETOR_12X = "PLANO 08"
PLANO_DIRETOR_24X = "PLANO 09"

# Identificadores Especiais
SINDICATO_JAGUAQUARA = "SIND DOS EMP NO COM DE JAGUAQUARA"

# Estruturas Oficiais dos Arquivos de Carga (Parametrizadas por Bandeira)
COLUNAS_CARGA_PADRAO_CPF = [
    "Matrícula",
    "Colaborador",
    "Data Nascimento",
    "CPF",
    "PLANO"
]

COLUNAS_CARGA_PADRAO_CPF_CNPJ = [
    "Matrícula",
    "Colaborador",
    "Data Nascimento",
    "CPF/CNPJ",
    "PLANO"
]

COLUNAS_CARGA_COMPLEMENTAR = [
    "Matrícula",
    "Nome",
    "NASCIMENTO",
    "CPF",
    "SALARIO",
    "Plano"
]

# Paleta Oficial da Interface Gráfica Cencosud (CustomTkinter)
CENCOSUD_BLUE = "#095AA5"
CENCOSUD_BLUE_HOVER = "#0B6FCB"
CENCOSUD_DARK_BLUE = "#063A6B"
CENCOSUD_ORANGE = "#FF690A"
CENCOSUD_ORANGE_HOVER = "#E05500"
SUCCESS_GREEN = "#1E7B34"
SUCCESS_GREEN_HOVER = "#28A745"

# Padrão Visual Institucional Cencosud nos Relatórios Excel
# 1. Colunas Cadastrais: Azul Corporativo (#1F497D) com fonte branca negrito
COR_CABECALHO_CADASTRAL = "1F497D"
COR_FONTE_CADASTRAL = "FFFFFF"

# 2. Colunas de Conferência / Validação / Auditoria: Amarelo (#FFFF00) com fonte vermelha negrito (#9C0006)
COR_CABECALHO_CONFERENCIA = "FFFF00"
COR_FONTE_CONFERENCIA = "9C0006"

# 3. Formatações e Cores Complementares Excel
COR_ZEBRA_PAR = "F2F5F9"
COR_BORDA = "D9D9D9"
COR_KPI_HEADER = "1F497D"
COR_KPI_BOX = "E9EFF7"
