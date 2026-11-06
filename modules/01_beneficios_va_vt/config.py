"""
Configurações e Caminhos Resilientes para a Conferência de Valores Pré-Folha VT.
Eventos Auditados: 5750 (Adiantamento VT Através Alçadas), 31935 (VT Alternativo) e 6292 (Desconto VTE).
Competência: 09/2026.
"""

from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
BASE_DIR = SCRIPT_DIR.parent

if not BASE_DIR.exists():
    va_vt_root = SCRIPT_DIR.parent.parent
    matches = list(va_vt_root.glob("Valores Pr*-Folha"))
    if matches:
        BASE_DIR = matches[0]

INPUT_DIR = BASE_DIR
OUTPUT_DIR = BASE_DIR / "Relatório Tratado"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

COMPETENCIA = "09/2026"
COMPETENCIA_ARQUIVO = "09.2026"
OUTPUT_FILE = OUTPUT_DIR / f"Conferência Valores Pré-Folha VT_{COMPETENCIA_ARQUIVO}.xlsx"

# Padrões de arquivos de entrada
GLOB_CONFERENCIA = "Confer*09.2026*.xlsx"
GLOB_FOPA_BRUTO = "FOPA_Eventos_por_Colaborador*.xlsx"

# Paleta Institucional Cencosud
COLOR_NAVY = "003366"
COLOR_BLUE = "1F4E78"
COLOR_LIGHT_BLUE = "D9E1F2"
COLOR_ZEBRA = "F9FAFC"
COLOR_ALERT_BG = "FCE4D6"
COLOR_OK_BG = "E2EFDA"
COLOR_BORDER = "D9D9D9"

# Colunas dos 10 eventos que compõem a base do evento 5750 (Adiantamento VT Através Alçadas)
EVENTOS_SOMA_5750 = [
    "147 - Credito Vale Transporte Dev Operadora FL 40",
    "148 - Credito Vale Transporte Dev Operadora FL 41",
    "159 - Deslocamento Atividades Externa",
    "490 - Credito Vale Transporte",
    "9421 - Reembolso VT Sem Dedução - Folha 40",
    "9422 - Reembolso VT Sem Dedução - Folha 41",
    "9423 - Reembolso VT Sem Dedução - Folha 42",
    "9428 - Deslocamento Atividades Externa - Folha 41",
    "9429 - Deslocamento Atividades Externa - Folha 42",
    "31935 - Valor Vale Transporte Alternativo Mes Seguinte",
]

# Colunas oficiais da Conferência Geral (41 colunas)
CONFERENCIA_COLUMNS = [
    "Código da Empresa",
    "Nome da Empresa",
    "Código Estabelecimento",
    "Estabelecimento",
    "Código Centro Custo",
    "Centro Custo",
    "Matrícula",
    "Nome do Colaborador",
    "ID Colaborador",
    "Data Admissão",
    "Código Cargo",
    "Cargo",
    "Situação",
    "Tipo Situação",
    "Data Rescisão",
    "118 - Credito Vale Transporte",
    "147 - Credito Vale Transporte Dev Operadora FL 40",
    "148 - Credito Vale Transporte Dev Operadora FL 41",
    "159 - Deslocamento Atividades Externa",
    "490 - Credito Vale Transporte",
    "5731 - Saldo Vale Transporte Rescisao",
    "5750 - Adiantamento VT Atraves Alcadas",
    "6292 - Desconto Cartao VTE",
    "9290 - Dev Vale Transporte",
    "9421 - Reembolso VT Sem Dedução - Folha 40",
    "9422 - Reembolso VT Sem Dedução - Folha 41",
    "9423 - Reembolso VT Sem Dedução - Folha 42",
    "9424 - Credito Vale Transporte - Folha 41",
    "9425 - Credito Vale Transporte - Folha 42",
    "9426 - Dev Vale Transporte - Folha 41",
    "9427 - Dev Vale Transporte - Folha 42",
    "9428 - Deslocamento Atividades Externa - Folha 41",
    "9429 - Deslocamento Atividades Externa - Folha 42",
    "31935 - Valor Vale Transporte Alternativo Mes Seguinte",
    "Soma Eventos Pagamento",
    "Diferença (Soma x 5750)",
    "Evts Pgtos x Evt 5750",
    "Importado 31935",
    "Importado 6292",
    "Status Auditoria",
    "Obs.",
]
