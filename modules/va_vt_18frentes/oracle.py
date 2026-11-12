"""
Authoritative Reference Oracle & Contract Specification for "04 - GOOD CARD"
Derived strictly from PROJECT.md, business_rules_architecture.md, and Conferência Good Card_Base.xlsx.
"""

import os
import re
import math
import importlib
from typing import Any, Optional, Dict, Tuple, List
import pandas as pd
import numpy as np

# --- OFFICIAL SPECIFICATIONS & CONSTANTS ---

COLUNAS_CONFERENCIA: List[str] = [
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
    "1 - Salario",
    "11010 - Salario Mensal",
    "10990 - Liquido a Receber",
    "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ",
    "29613 - Ben Esp - Valor Informado para Desconto",
    "Total Saldo Devedor",
    "Margem 20% do Salário",
    "Total Saldo Devedor \" Dentro da Margem\" ou \"Fora da Margem\"",
    "6280 - Desc Cartão Benefícios Especiais",
    "Desconto x Margem de 20%",
    "Saldo Devedor x Evt 6280",
    "Total Saldo Devedor - Evt 6280 x Evt 29621",
    "29621 - Ben Esp - Saldo Devedor Beneficios Especiais",
    "29624 - Ben Esp - Custo Empresa Beneficios Especiais",
    "5% Saldo devedor",
    "Cálculo 5% x 29624",
    "Início Afastamento",
    "Final Afastamento",
    "Tipo do Afastamento",
    "Início Férias",
    "Final Férias",
    "Qtd de dias",
    "Desligamento",
    "Situação",
    "Obs. Conferência",
    "Tratativa"
]

COLUNAS_VALORES_IMPORTADOS: List[str] = [
    "Emp",
    "Matrícula",
    "Evento",
    "Mês",
    "Valor",
    "Folha",
    "Importado x Folha",
    "Desligado",
    "Tipo de Rescisão",
    "Obs. Conferência",
    "Tratativa"
]

STATUS_LISTA: List[str] = [
    "Ok - Sem saldo devedor e sem desconto",
    "Ok - Liquidou todo o saldo devedor ",
    "Ok - Suspensão de Contrato",
    "Afastamento - Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ",
    "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621",
    "Férias- Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ",
    "Verificar com quality - Não estava em situação de afastamento e não houve o desconto",
    "Aprendiz - Verificar",
    "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)"
]

PALETA_CORES = {
    "HEADER_FILL": "1F497D",
    "HEADER_FONT": "FFFFFF",
    "DIVERGENCE_FILL": "FFC7CE",
    "DIVERGENCE_FONT": "9C0006",
    "OK_FILL": "C6EFCE",
    "OK_FONT": "006100",
    "SUB_HEADER_FILL": "DCE6F1"
}

TOLERANCIA_CENTAVOS: float = 0.02


# --- MODULE RESOLVER ---

def get_module(module_name: str):
    """Dynamically resolves and imports project modules if available."""
    try:
        return importlib.import_module(module_name)
    except (ImportError, ModuleNotFoundError):
        return None


# --- ORACLE IMPLEMENTATIONS ---

def oracle_parse_currency(val: Any) -> float:
    """Authoritative parsing for Brazilian currency values."""
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        if math.isnan(val):
            return 0.0
        return round(float(val), 2)
    s = str(val).strip()
    if s in ("", "-", " -", "nan", "NaN", "None", "<NA>", "null"):
        return 0.0
    # Clean currency symbols and thousands dots
    s = s.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
    try:
        return round(float(s), 2)
    except ValueError:
        return 0.0


def oracle_normalize_matricula(val: Any) -> str:
    """Authoritative normalization of Matrícula string."""
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, float):
        if np.isnan(val):
            return ""
        if val.is_integer():
            return str(int(val))
        return str(val)
    if isinstance(val, int):
        return str(val)
    s = str(val).strip()
    if s in ("", "-", " -", "nan", "NaN", "None", "<NA>", "null"):
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    elif re.search(r"\.0+$", s):
        s = re.sub(r"\.0+$", "", s)
    return s.strip()


def oracle_calculate_margem_20(salario_mensal: float) -> float:
    """Calculates 20% margin of salary rounded to 2 decimal places."""
    return round(float(salario_mensal) * 0.20, 2)


def oracle_calculate_total_saldo_devedor(evt_29622: float, evt_29613: float) -> float:
    """Calculates Total Saldo Devedor = Event 29622 + Event 29613."""
    return round(float(evt_29622) + float(evt_29613), 2)


def oracle_check_status_margem(valor: float, margem_20: float) -> str:
    """Returns 'Dentro da Margem' or 'Fora da Margem'."""
    if round(float(valor), 2) <= round(float(margem_20), 2):
        return "Dentro da Margem"
    return "Fora da Margem"


def oracle_check_quitacao(evt_6280: float, evt_29624: float, total_saldo_devedor: float) -> bool:
    """Checks whether 6280 + 29624 == Total Saldo Devedor (within 0.02 tolerance)."""
    soma = round(float(evt_6280) + float(evt_29624), 2)
    total = round(float(total_saldo_devedor), 2)
    return abs(soma - total) <= TOLERANCIA_CENTAVOS


def oracle_check_quitacao_residual(evt_6280: float, evt_29621: float, evt_29624: float, total_saldo_devedor: float) -> bool:
    """Checks complete accounting equation 6280 + 29621 + 29624 == Total Saldo Devedor."""
    soma = round(float(evt_6280) + float(evt_29621) + float(evt_29624), 2)
    total = round(float(total_saldo_devedor), 2)
    return abs(soma - total) <= TOLERANCIA_CENTAVOS


def oracle_custo_empresa_5pct(total_saldo_devedor: float) -> Tuple[float, float]:
    """Returns (empresa_5pct, colaborador_95pct)."""
    empresa = round(float(total_saldo_devedor) * 0.05, 2)
    colaborador = round(float(total_saldo_devedor) * 0.95, 2)
    # Adjustment for penny rounding
    if round(empresa + colaborador, 2) != round(total_saldo_devedor, 2):
        colaborador = round(total_saldo_devedor - empresa, 2)
    return empresa, colaborador


def oracle_is_aprendiz(cargo: str) -> bool:
    """Detects whether cargo indicates Aprendiz."""
    if not cargo:
        return False
    return "APRENDIZ" in str(cargo).upper()


def oracle_is_diretor(cargo: str) -> bool:
    """Detects whether cargo indicates Diretor/Diretoria (excluding Analista de Impostos Diretos)."""
    if not cargo:
        return False
    c = str(cargo).upper()
    return bool(re.search(r"\bDIRETOR\b|\bDIRETORA\b|\bDIRETORIA\b", c))


def oracle_is_giga_ou_estab400(estabelecimento_codigo: Any, estabelecimento_nome: str) -> bool:
    """Checks if establishment is 400 or belongs to Giga."""
    estab_cod_str = str(estabelecimento_codigo).strip()
    if estab_cod_str == "400":
        return True
    nome = str(estabelecimento_nome).upper()
    return "GIGA" in nome or "400" in nome


def oracle_classify_audit_case(
    total_saldo_devedor: float,
    salario_11010: float,
    evt_6280: float,
    evt_29621: float,
    evt_29624: float,
    cargo: str = "",
    dias_afastamento: int = 0,
    has_afastamento: bool = False,
    is_ferias: bool = False,
    is_demitido: bool = False,
    tipo_rescisao: str = "",
    is_estab400_or_giga: bool = False,
    situacao: str = "Atividade Normal"
) -> Tuple[str, Optional[str]]:
    """
    Authoritative classification of a collaborator record returning:
    (Obs. Conferência, Tratativa)
    """
    total = round(total_saldo_devedor, 2)
    margem = round(salario_11010 * 0.20, 2)
    desc = round(evt_6280, 2)
    saldo_rem = round(evt_29621, 2)
    custo_emp = round(evt_29624, 2)

    # 1. Cargos Inelegíveis (Regra 3)
    if oracle_is_aprendiz(cargo) and (total > 0 or desc > 0):
        return "Aprendiz - Verificar", "Concessão indevida de Good Card para cargo de Aprendiz. Solicitar cancelamento do cartão junto à operadora."

    if oracle_is_diretor(cargo) and (total > 0 or desc > 0):
        return "Diretor - Verificar", "Concessão indevida de Good Card para cargo de Diretor. Cargo inelegível conforme política de benefícios."

    # 2. Desligados (Regra 1)
    if is_demitido:
        if "TRANSFERÊNCIA SEM ÔNUS" in str(tipo_rescisao).upper():
            return "Transferência sem ônus", "Transferência de empresa: verificar migração do cartão e acerto entre filiais."
        if total == 0 and desc == 0:
            return "Ok - Sem saldo devedor e sem desconto", None
        if total > 0 and desc == total and saldo_rem == 0:
            return "Ok - Liquidou todo o saldo devedor ", None
        if saldo_rem > 0 or (total > desc):
            return "Demitido - Saldo devedor pendente", "Colaborador desligado com saldo devedor não liquidado. Verificar inclusão no TRCT e solicitar cancelamento do cartão."

    # 3. Suspensão de Contrato
    if "SUSPENS" in str(situacao).upper():
        return "Ok - Suspensão de Contrato", None

    # 4. Sem Saldo e Sem Desconto
    if total == 0 and desc == 0:
        return "Ok - Sem saldo devedor e sem desconto", None

    # 5. Afastados > 15 dias (Regra 2)
    if has_afastamento and (dias_afastamento > 15 or "AFAST" in str(situacao).upper()):
        if desc == 0 and saldo_rem == total and total > 0:
            return "Afastamento - Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ", None
        if desc > 0 and salario_11010 == 0:
            return "Afastado - Desconto indevido em folha", "Colaborador afastado > 15 dias com desconto em folha. Validar margem líquida e solicitar estorno/regularização."
        if total > 0 and desc == 0 and saldo_rem == 0:
            return "Afastado - Saldo devedor não provisionado", "Saldo devedor não transferido para o evento 29621 na folha. Acionar equipe de folha para provisionamento."

    # 6. Férias no mês
    if is_ferias and desc == 0 and saldo_rem == total and total > 0:
        return "Férias- Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ", None

    # 7. Quitação Integral Universal (Válido para todas as filiais, incluindo Giga/400 quando o colaborador liquida 100%)
    if total <= margem and abs(desc - total) <= TOLERANCIA_CENTAVOS and saldo_rem == 0 and custo_emp == 0:
        return "Ok - Liquidou todo o saldo devedor ", None

    # 8. Regra Especial 5% Custo Empresa (Giga / Estab 400)
    if is_estab400_or_giga and total > 0:
        calc_emp = round(total * 0.05, 2)
        if abs(custo_emp - calc_emp) <= 0.05:
            # Equação contábil completa incluindo eventual saldo_rem
            if abs(desc + saldo_rem + custo_emp - total) <= TOLERANCIA_CENTAVOS:
                return "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)", None
        elif custo_emp > 0 and saldo_rem == 0:
            return "Divergência Custo Empresa 5%", "Valor levado para o custo empresa (29624) diverge dos 5% regulamentares para o estabelecimento 400/Giga."

    # 9. Regra 4: Desconto e Quitação Padrão
    # 9.1 Desconto superior à margem de 20%
    if desc > margem:
        return "Desconto fora da margem dos 20%", "Desconto no evento 6280 superior à margem consignável de 20% do salário mensal (11010). Ajustar valor descontado."

    # 9.2 Total Fora da Margem, descontou dentro da margem e levou o restante para 29621 (inclui Giga com saldo remanescente)
    if total > margem and desc <= margem and abs(desc + saldo_rem + custo_emp - total) <= TOLERANCIA_CENTAVOS and saldo_rem > 0:
        return "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621", None

    # 9.3 Ativo, sem afastamento, sem férias, com saldo e SEM desconto
    if not has_afastamento and not is_ferias and total > 0 and desc == 0:
        return "Verificar com quality - Não estava em situação de afastamento e não houve o desconto", "Verificar com Quality/Folha: Colaborador com saldo devedor ativo e sem desconto em folha, não estando afastado nem em férias."

    # 9.4 Quebra da equação contábil
    if abs(desc + saldo_rem + custo_emp - total) > TOLERANCIA_CENTAVOS:
        return "Divergência no saldo devedor", "Soma do desconto (6280) + saldo residual (29621) + custo empresa (29624) diverge do Total Saldo Devedor. Auditar folha."

    return "Ok - Liquidou todo o saldo devedor ", None
