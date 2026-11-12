"""
utils.py - Funções utilitárias compartilhadas.

Aqui ficam as funções "tijolo": limpar CPF, normalizar nome, validar data.
Tudo que é usado por mais de um módulo mora aqui.
"""
import re
import unicodedata
from datetime import datetime
import pandas as pd


# ---------------------------------------------------------------------------
# CPF
# ---------------------------------------------------------------------------

_CPF_INVALIDOS = {c * 11 for c in "0123456789"}  # 00000000000, 11111111111, ...


def limpar_cpf(valor) -> str:
    """
    Retorna o CPF como string de 11 dígitos (com zeros à esquerda)
    ou string vazia se for inválido.

    Aceita: '123.456.789-00', '12345678900', 12345678900 (int), '1.23E+10' etc.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""

    # Se veio como float (Excel adora transformar tudo em float), formata sem ',0'
    if isinstance(valor, float):
        valor = f"{valor:.0f}"

    # Tira tudo que não é dígito
    apenas_digitos = re.sub(r"\D", "", str(valor))

    if not apenas_digitos:
        return ""

    # Zero-pad até 11 (CPFs que começam com 0 perdem o zero quando viram número)
    apenas_digitos = apenas_digitos.zfill(11)

    # Se tem mais de 11 dígitos, é lixo
    if len(apenas_digitos) != 11:
        return ""

    return apenas_digitos


def validar_cpf(cpf: str) -> bool:
    """
    Valida dígitos verificadores do CPF.
    Retorna False para CPFs com todos os dígitos iguais (000..., 111...).
    """
    if not cpf or len(cpf) != 11 or not cpf.isdigit():
        return False

    if cpf in _CPF_INVALIDOS:
        return False

    # Cálculo do 1º dígito
    soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
    dig1 = (soma * 10 % 11) % 10
    if dig1 != int(cpf[9]):
        return False

    # Cálculo do 2º dígito
    soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
    dig2 = (soma * 10 % 11) % 10
    if dig2 != int(cpf[10]):
        return False

    return True


def formatar_cpf(cpf: str) -> str:
    """Formata CPF como 000.000.000-00 (só para exibição final)."""
    cpf = limpar_cpf(cpf)
    if len(cpf) != 11:
        return ""
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"


# ---------------------------------------------------------------------------
# Nomes
# ---------------------------------------------------------------------------

def normalizar_nome(valor) -> str:
    """
    Padroniza nomes para servir como chave de merge:
    - tira acentos
    - maiúsculo
    - tira espaços extras
    - tira pontuação
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""

    texto = str(valor).strip().upper()

    # Remove acentos
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))

    # Remove pontuação e múltiplos espaços
    texto = re.sub(r"[^A-Z0-9 ]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()

    return texto


# ---------------------------------------------------------------------------
# Matrícula
# ---------------------------------------------------------------------------

def limpar_matricula(valor) -> str:
    """
    Padroniza matrícula como string sem zeros à esquerda redundantes
    nem '.0' do float. Aceita matrícula alfanumérica.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""

    if isinstance(valor, float):
        valor = f"{valor:.0f}"

    return str(valor).strip().upper()


# ---------------------------------------------------------------------------
# Datas
# ---------------------------------------------------------------------------

def parse_data(valor):
    """
    Tenta converter para datetime. Devolve pd.NaT se não conseguir.
    Aceita: datetime, string DD/MM/AAAA, string AAAA-MM-DD, número Excel.
    """
    if valor is None or valor == "" or pd.isna(valor):
        return pd.NaT

    if isinstance(valor, datetime):
        return pd.Timestamp(valor)

    # Se já vier em formato ISO (AAAA-MM-DD do Excel), não inverter dia/mês
    if isinstance(valor, str) and re.match(r"^\d{4}-\d{2}-\d{2}", valor.strip()):
        try:
            return pd.to_datetime(valor.strip(), dayfirst=False, errors="coerce")
        except Exception:
            return pd.NaT

    # pandas já tenta vários formatos; dayfirst=True garante DD/MM/AAAA
    try:
        return pd.to_datetime(valor, format="mixed", dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT


def formatar_data_br(valor) -> str:
    """Formata datetime como DD/MM/AAAA. Vazio se NaT."""
    if valor is None or pd.isna(valor):
        return ""
    if isinstance(valor, str):
        valor = parse_data(valor)
        if pd.isna(valor):
            return ""
    return valor.strftime("%d/%m/%Y")
