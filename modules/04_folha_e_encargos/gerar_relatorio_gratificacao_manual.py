"""
=====================================================================================
CENCOSUD BRASIL - ENGENHARIA DE DADOS & AUDITORIA DE RH
MÓDULO 7: AUDITORIA E APURAÇÃO DE GRATIFICAÇÃO MANUAL — DOMINGOS E FERIADOS (GB & MR)
=====================================================================================
Esteira automatizada de apuração e auditoria das Gratificações Sindicais de Domingos e
Feriados (Eventos 227 e 5751 da Folha) para as lojas GBarbosa, Eletroshow e Mercantil
da Bahia (Sindicatos 6, 9, 36, 40, 60, 71, 74 e 82).

Fontes de Dados Integradas:
  1. Relatorio_HorasTrabalhadas_16102026134959.xlsx (Batidas e Horas Trabalhadas)
  2. Regras das Gratificações GB e Mercantil.xlsx (Tabelas CCT Domingos, Feriados, Check-outs)
  3. Feriados Folha - Padrão.xlsb & Feriados_Cadastrados15102026.csv (Calendário Municipal/Nacional)
  4. ORGA_Estabelecimento_CSV - 16072026_150121_960183.CSV (Cadastro de Filiais e Municípios)
  5. COLA_Colaboradores_CSV_Gestor - 15102026_100130_412050.CSV (Base Cadastral Ativa)
  6. GRATIFICAÇÃO 10.2026.xlsx (Planilha Manual da Equipe - Espelho de Conciliação)

Estrutura Oficial do Workbook (5 Abas):
  - Aba 1: Orientações (Documentação Institucional, Regras Sindicais CCT e Metodologia)
  - Aba 2: Resumo da Auditoria (KPIs Executivos, Resumo por Loja/Sindicato, Quadro Autorização EVT 227/5751)
  - Aba 3: Conferência Geral (300 Ocorrências Auditadas Dia a Dia)
  - Aba 4: Divergências (Inconsistências entre Marcações Reais x Horas Trabalhadas / Batidas Ímpares)
  - Aba 5: Consolidado Folha (Soma por Matrícula para os 230 Colaboradores - Eventos 227 e 5751)
=====================================================================================
"""

import os
import re
import csv
import time
import shutil
from datetime import datetime
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pyxlsb

# ======================================================================================
# 1. DIRETÓRIOS E ARQUIVOS DE ENTRADA / SAÍDA
# ======================================================================================

BASE_DIR = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\ponto"
GRAT_DIR = os.path.join(BASE_DIR, "Gratificação Manual")
OUTPUT_DIR = os.path.join(BASE_DIR, "conferencias geradas pelo script")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Relatorio_Conferencia_Gratificacao_Manual_09.2026.xlsx")

FILE_HT = os.path.join(GRAT_DIR, "Relatorio_HorasTrabalhadas_16102026134959.xlsx")
FILE_REGRAS = os.path.join(GRAT_DIR, "Regras das Gratificações GB e Mercantil.xlsx")
FILE_FER_XLSB = os.path.join(GRAT_DIR, "Feriados Folha - Padrão.xlsb")
FILE_FER_CSV = os.path.join(GRAT_DIR, "Feriados_Cadastrados15102026.csv")
FILE_ORGA_CSV = os.path.join(GRAT_DIR, "ORGA_Estabelecimento_CSV - 16072026_150121_960183.CSV")
FILE_COLA_CSV = os.path.join(GRAT_DIR, "COLA_Colaboradores_CSV_Gestor - 15102026_100130_412050.CSV")
FILE_GRAT_MANUAL = os.path.join(GRAT_DIR, "GRATIFICAÇÃO 10.2026.xlsx")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ======================================================================================
# 2. PALETA DE CORES E ESTILOS CORPORATIVOS CENCOSUD
# ======================================================================================

C_AZUL_MARINHO = "1B365D"
C_AZUL_MEDIO   = "2C4A7A"
C_AZUL_ESCURO  = "0F172A"
C_CINZA_CAB    = "F1F5F9"
C_CINZA_ZEBRA  = "F8FAFC"
C_BRANCO       = "FFFFFF"
C_BORDA        = "CBD5E1"

C_CRIT_BG      = "FFC7CE"
C_CRIT_TXT     = "9C0006"
C_ALERT_BG     = "FFEB9C"
C_ALERT_TXT    = "9C6500"
C_REG_BG       = "C6EFCE"
C_REG_TXT      = "006100"

FONT_TITULO = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
FONT_SUBTITULO = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_CABECALHO = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_DADOS = Font(name="Calibri", size=10, color="000000")
FONT_DADOS_BOLD = Font(name="Calibri", size=10, bold=True, color="000000")

FILL_HEADER = PatternFill(start_color=C_AZUL_MARINHO, end_color=C_AZUL_MARINHO, fill_type="solid")
FILL_HEADER_SEC = PatternFill(start_color=C_AZUL_MEDIO, end_color=C_AZUL_MEDIO, fill_type="solid")
FILL_ZEBRA = PatternFill(start_color=C_CINZA_ZEBRA, end_color=C_CINZA_ZEBRA, fill_type="solid")
FILL_BRANCO = PatternFill(start_color=C_BRANCO, end_color=C_BRANCO, fill_type="solid")
FILL_TOTAL = PatternFill(start_color=C_CINZA_CAB, end_color=C_CINZA_CAB, fill_type="solid")

FILL_CRIT = PatternFill(start_color=C_CRIT_BG, end_color=C_CRIT_BG, fill_type="solid")
FONT_CRIT = Font(name="Calibri", size=10, bold=True, color=C_CRIT_TXT)

FILL_ALERT = PatternFill(start_color=C_ALERT_BG, end_color=C_ALERT_BG, fill_type="solid")
FONT_ALERT = Font(name="Calibri", size=10, bold=True, color=C_ALERT_TXT)

FILL_REG = PatternFill(start_color=C_REG_BG, end_color=C_REG_BG, fill_type="solid")
FONT_REG = Font(name="Calibri", size=10, bold=True, color=C_REG_TXT)

BORDER_THIN = Border(
    left=Side(style="thin", color=C_BORDA),
    right=Side(style="thin", color=C_BORDA),
    top=Side(style="thin", color=C_BORDA),
    bottom=Side(style="thin", color=C_BORDA)
)

BORDER_TOTAL = Border(
    top=Side(style="thin", color=C_AZUL_MARINHO),
    bottom=Side(style="double", color=C_AZUL_MARINHO),
    left=Side(style="thin", color=C_BORDA),
    right=Side(style="thin", color=C_BORDA)
)

ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

DIAS_SEMANA_PT = {
    0: "seg", 1: "ter", 2: "qua", 3: "qui", 4: "sex", 5: "sáb", 6: "dom"
}

# ======================================================================================
# 3. FUNÇÕES AUXILIARES E MOTOR DE CÁLCULO SINDICAL (CCT BAHIA)
# ======================================================================================

def parse_minutes(hhmm_str):
    """Converte string HH:MM em minutos inteiros."""
    if not hhmm_str:
        return 0
    s = str(hhmm_str).strip()
    if ":" not in s:
        return 0
    try:
        parts = s.split(":")
        return int(parts[0]) * 60 + int(parts[1])
    except Exception:
        return 0

def format_minutes(mins):
    """Formata minutos inteiros em HH:MM."""
    if mins <= 0:
        return "00:00"
    return f"{mins // 60:02d}:{mins % 60:02d}"

def format_data_br(dt_val):
    """Converte data YYYY-MM-DD ou datetime para DD/MM/AAAA."""
    if isinstance(dt_val, datetime):
        return dt_val.strftime("%d/%m/%Y")
    s = str(dt_val).strip()[:10]
    if len(s) == 10 and s[4] == "-" and s[7] == "-":
        return f"{s[8:10]}/{s[5:7]}/{s[0:4]}"
    return s

def analisar_batidas(marc_str):
    """
    Analisa a string de marcações:
    Retorna (qtd_batidas, is_impar, mins_calculados, str_calculado, ultima_batida_mins).
    """
    if not marc_str:
        return 0, False, 0, "00:00", 0
    batidas = [b.strip() for b in str(marc_str).strip().split() if ":" in b]
    qtd = len(batidas)
    if qtd == 0:
        return 0, False, 0, "00:00", 0
    is_impar = (qtd % 2 != 0)
    mins_list = [parse_minutes(b) for b in batidas]
    ultima_batida = mins_list[-1]
    if not is_impar and qtd >= 2:
        calc_mins = 0
        for i in range(0, qtd, 2):
            diff = mins_list[i + 1] - mins_list[i]
            if diff < 0:
                diff += 1440
            calc_mins += diff
        return qtd, False, calc_mins, format_minutes(calc_mins), ultima_batida
    else:
        # Se ímpar, calcula pares completos disponíveis
        calc_mins = 0
        for i in range(0, qtd - 1, 2):
            diff = mins_list[i + 1] - mins_list[i]
            if diff < 0:
                diff += 1440
            calc_mins += diff
        return qtd, True, calc_mins, format_minutes(calc_mins), ultima_batida

# Tabela Padrão Oficial dos 8 Sindicatos (Regras das Gratificações GB e Mercantil.xlsx)
REGRAS_SINDICATOS_PADRAO = [
    {
        "cod": 6,
        "nome": "SIN EMPREG COMER ALAGOINHAS",
        "municipios": "Alagoinhas, Cipó, Entre Rios, Olindina, Rio Real, Monte Santo, Cansanção, Queimadas, Santaluz, Crisópolis",
        "regra_dom": "Check-outs: Até 4 Check-outs = R$ 73,50 | Acima de 4 Check-outs = R$ 79,80",
        "regra_fer": "Horas Trabalhadas: Até 6h = R$ 87,19 | Até 07:20 = R$ 94,68 | Jornada 8h = R$ 100,53",
        "feriados_proibidos": "01/01 (Confraternização Universal), 01/05 (Dia do Trabalhador), 25/12 (Natal) e Dia do Comerciário",
        "atualizacao": "ATUALIZADO 15/05/2026"
    },
    {
        "cod": 9,
        "nome": "SIN EMPR ESTAB COMER SERV PAULO AFONSO E REGIÃO",
        "municipios": "Paulo Afonso, Tucano, Euclides da Cunha, Jeremoabo, Paripiranga",
        "regra_dom": "Horas Trabalhadas: 5h = R$ 57,00 | 6h = R$ 64,00 | 8h = R$ 90,00",
        "regra_fer": "Horas Trabalhadas: Até 5h = R$ 57,00 | Até 6h = R$ 64,00 | Jornada 8h = R$ 90,00",
        "feriados_proibidos": "01/01, Sexta-feira da Paixão, 01/05, 20/11 (Consciência Negra), 25/12 e Dia do Comerciário",
        "atualizacao": "ATUALIZADO 18/05/2026"
    },
    {
        "cod": 36,
        "nome": "SIN EMP EST COM AFONSO E REG R. DO POMBAL",
        "municipios": "Ribeira do Pombal, Cícero Dantas",
        "regra_dom": "Horas Trabalhadas: 6h = R$ 50,00 | 8h = R$ 60,00",
        "regra_fer": "Horas Trabalhadas: Até 6h = R$ 52,00 | Jornada 8h = R$ 71,00",
        "feriados_proibidos": "01/01, Sexta-feira da Paixão, 01/05, 20/11, 25/12 e Dia do Comerciário ( Carnaval )",
        "atualizacao": "ATUALIZADO 16/07/2026"
    },
    {
        "cod": 40,
        "nome": "SIN EMP COM DA CIDADE DE JUAZEIRO E REGIÃO",
        "municipios": "Juazeiro, Campo Formoso",
        "regra_dom": "Domingo Normal = R$ 70,44 | Após 18:30 (até 6h) = R$ 79,03 | Após 18:30 (acima 6h) = R$ 96,01",
        "regra_fer": "Horas Trabalhadas: Até 6h = R$ 80,34 | Jornada 8h = R$ 101,46",
        "feriados_proibidos": "01/01, Segunda de Carnaval, 01/05 e 25/12",
        "atualizacao": "ATUALIZADO 18/05/2026"
    },
    {
        "cod": 60,
        "nome": "SIN EMP COMER VITÓRIA DA CONQUISTA",
        "municipios": "Vitória da Conquista",
        "regra_dom": "Horas Trabalhadas: Até 6h = R$ 33,65 | Até 07:20 = R$ 45,60",
        "regra_fer": "Porte Nacional: Até 1.000 func. = R$ 60,85 | Acima de 1.000 func. nacional = R$ 68,40",
        "feriados_proibidos": "01/01 (Confraternização Universal) e 25/12 (Natal)",
        "atualizacao": "CCT VIGENTE 2026"
    },
    {
        "cod": 71,
        "nome": "SIN DOS EMP NO COMER ALAGOINHAS E REGIÃO",
        "municipios": "Alagoinhas e Região (Base Sindicato 71)",
        "regra_dom": "Check-outs: Até 4 Check-outs = R$ 70,00 | Acima de 4 Check-outs = R$ 76,00",
        "regra_fer": "Horas Trabalhadas: Até 6h = R$ 83,04 | Até 07:20 = R$ 90,18 | Jornada 8h = R$ 95,75",
        "feriados_proibidos": "01/01, 01/05, 25/12 e Dia do Comerciário",
        "atualizacao": "CCT VIGENTE 2026"
    },
    {
        "cod": 74,
        "nome": "SIND EMPREG NO COM DE CATU",
        "municipios": "Catu",
        "regra_dom": "Check-outs: Até 4 Check-outs = R$ 70,35 | Acima de 4 Check-outs = R$ 80,00",
        "regra_fer": "Horas Trabalhadas: Até 6h = R$ 84,73 | Até 07:20 = R$ 93,45 | Jornada 8h = R$ 103,40",
        "feriados_proibidos": "01/01, 01/05 e Dia do Comerciário",
        "atualizacao": "CCT VIGENTE 2026"
    },
    {
        "cod": 82,
        "nome": "SINDICATO EMP COMER CONCEIÇÃO DO JACUÍPE",
        "municipios": "Conceição do Jacuípe",
        "regra_dom": "Porte Loja: Até 50 empregados = R$ 27,45 | Mais de 50 empregados = R$ 36,70",
        "regra_fer": "Até 14:00 (<=50 func: R$ 35,00 | >=51 func: R$ 53,70) | Após 14:00 (<=50 func: R$ 51,00 | >=51 func: R$ 81,40)",
        "feriados_proibidos": "01/01 (Confraternização Universal)",
        "atualizacao": "CCT VIGENTE 2026"
    }
]

def calcular_valor_cct(sindicato, is_domingo, mins_trabalhados, checkouts="Acima de 4", ultima_batida_mins=0, qtd_empregados_loja=60):
    """
    Calcula o valor exato da Gratificação conforme a Convenção Coletiva (CCT)
    da planilha 'Regras das Gratificações GB e Mercantil.xlsx' para todos os 8 sindicatos padrão:
    6, 9, 36, 40, 60, 71, 74 e 82.
    Retorna (valor, descricao_faixa).
    """
    sind = int(sindicato)
    is_acima_4_co = "ACIMA" in str(checkouts).upper() or "MAIS" in str(checkouts).upper()

    if sind == 6:  # SIN EMPREG COMER ALAGOINHAS
        if is_domingo:
            val = 79.80 if is_acima_4_co else 73.50
            faixa = "Domingo (> 4 Check-outs = R$ 79,80)" if is_acima_4_co else "Domingo (Até 4 Check-outs = R$ 73,50)"
            return val, faixa
        else:
            if mins_trabalhados <= 360:
                return 87.19, "Feriado - Jornada até 6h (R$ 87,19)"
            elif mins_trabalhados <= 440:
                return 94.68, "Feriado - Jornada até 07:20 (R$ 94,68)"
            else:
                return 100.53, "Feriado - Jornada 8h (R$ 100,53)"

    elif sind == 9:  # SIN EMPR ESTAB COMER SERV PAULO AFONSO E REGIÃO
        tipo_str = "Domingo" if is_domingo else "Feriado"
        if mins_trabalhados <= 300:
            return 57.00, f"{tipo_str} - Jornada até 5h (R$ 57,00)"
        elif mins_trabalhados <= 360:
            return 64.00, f"{tipo_str} - Jornada até 6h (R$ 64,00)"
        else:
            return 90.00, f"{tipo_str} - Jornada 8h / > 6h (R$ 90,00)"

    elif sind == 36:  # SIN EMP EST COM AFONSO E REG R. DO POMBAL
        if is_domingo:
            if mins_trabalhados <= 360:
                return 50.00, "Domingo - Jornada até 6h (R$ 50,00)"
            else:
                return 60.00, "Domingo - Jornada 8h / > 6h (R$ 60,00)"
        else:
            if mins_trabalhados <= 360:
                return 52.00, "Feriado - Jornada até 6h (R$ 52,00)"
            else:
                return 71.00, "Feriado - Jornada 8h / > 6h (R$ 71,00)"

    elif sind == 40:  # SIN EMP COM DA CIDADE DE JUAZEIRO E REGIÃO
        if is_domingo:
            if ultima_batida_mins > (18 * 60 + 30):
                if mins_trabalhados <= 360:
                    return 79.03, "Domingo - Até 6h após 18:30 (R$ 79,03)"
                else:
                    return 96.01, "Domingo - Acima 6h após 18:30 (R$ 96,01)"
            return 70.44, "Domingo Regular (R$ 70,44)"
        else:
            if mins_trabalhados <= 360:
                return 80.34, "Feriado - Turno até 6h (R$ 80,34)"
            else:
                return 101.46, "Feriado - Turno 8h / > 6h (R$ 101,46)"

    elif sind == 60:  # SIN EMP COMER VITÓRIA DA CONQUISTA
        if is_domingo:
            if mins_trabalhados <= 360:
                return 33.65, "Domingo - Até 6h (R$ 33,65)"
            else:
                return 45.60, "Domingo - Até 07:20 (R$ 45,60)"
        else:
            return 68.40, "Feriado - Acima 1.000 func. nacional (R$ 68,40)"

    elif sind == 71:  # SIN DOS EMP NO COMER ALAGOINHAS E REGIÃO
        if is_domingo:
            val = 76.00 if is_acima_4_co else 70.00
            faixa = "Domingo (> 4 Check-outs = R$ 76,00)" if is_acima_4_co else "Domingo (Até 4 Check-outs = R$ 70,00)"
            return val, faixa
        else:
            if mins_trabalhados <= 360:
                return 83.04, "Feriado - Até 6h (R$ 83,04)"
            elif mins_trabalhados <= 440:
                return 90.18, "Feriado - Até 07:20 (R$ 90,18)"
            else:
                return 95.75, "Feriado - Jornada 8h (R$ 95,75)"

    elif sind == 74:  # SIND EMPREG NO COM DE CATU
        if is_domingo:
            val = 80.00 if is_acima_4_co else 70.35
            faixa = "Domingo (> 4 Check-outs = R$ 80,00)" if is_acima_4_co else "Domingo (Até 4 Check-outs = R$ 70,35)"
            return val, faixa
        else:
            if mins_trabalhados <= 360:
                return 84.73, "Feriado - Até 6h (R$ 84,73)"
            elif mins_trabalhados <= 440:
                return 93.45, "Feriado - Até 07:20 (R$ 93,45)"
            else:
                return 103.40, "Feriado - Jornada 8h (R$ 103,40)"

    elif sind == 82:  # SINDICATO EMP COMER CONCEIÇÃO DO JACUÍPE
        mais_50 = (qtd_empregados_loja > 50)
        if is_domingo:
            if mais_50:
                return 36.70, "Domingo - Mais de 50 empregados (R$ 36,70)"
            else:
                return 27.45, "Domingo - Até 50 empregados (R$ 27,45)"
        else:
            apos_14h = (ultima_batida_mins > 14 * 60)
            if not apos_14h:
                if mais_50:
                    return 53.70, "Feriado - 51+ func. até 14:00 (R$ 53,70)"
                else:
                    return 35.00, "Feriado - Até 50 func. até 14:00 (R$ 35,00)"
            else:
                if mais_50:
                    return 81.40, "Feriado - 51+ func. após 14:00 (R$ 81,40)"
                else:
                    return 51.00, "Feriado - Até 50 func. após 14:00 (R$ 51,00)"

    return 0.0, "Sindicato Sem Regra Manual"

# ======================================================================================
# 4. CARREGAMENTO DAS BASES DE DADOS E PARAMETRIZAÇÃO
# ======================================================================================

def carregar_bases():
    print("-> [1/4] Carregando tabelas de Feriados, Estabelecimentos, Check-outs e Sindicatos...")

    # 1. Check-outs por Loja (Regras das Gratificações GB e Mercantil.xlsx)
    checkouts_map = {}
    wb_reg = openpyxl.load_workbook(FILE_REGRAS, read_only=True, data_only=True)
    ws_co = wb_reg["Check Out"]
    for idx, r in enumerate(ws_co.iter_rows(min_row=2, values_only=True)):
        if r[0] is not None and r[2] is not None:
            try:
                checkouts_map[int(float(r[0]))] = str(r[2]).strip()
            except Exception:
                pass
    wb_reg.close()

    # 2. Feriados Cadastrados por Localidade (Feriados_Cadastrados15102026.csv)
    feriados_desc_map = {}  # (cod_localidade, YYYY-MM-DD) -> Descricao do feriado
    with open(FILE_FER_CSV, mode="r", encoding="utf-8-sig") as f_fer:
        reader = csv.DictReader(f_fer, delimiter=";")
        for row in reader:
            try:
                cod_loc = int(float(row["Código"]))
                dt_str = str(row["Data"]).strip()[:10]
                desc = str(row["Descrição do feriado"]).strip()
                feriados_desc_map[(cod_loc, dt_str)] = desc
            except Exception:
                pass

    # 3. Mapeamento de Lojas x Sindicato x Localidade x Feriados (Feriados Folha - Padrão.xlsb)
    lojas_info = {}
    with pyxlsb.open_workbook(FILE_FER_XLSB) as wb_xlsb:
        with wb_xlsb.get_sheet("VALIDAÇÃO") as ws_val:
            for i, r in enumerate(ws_val.rows()):
                if i == 0:
                    continue
                vals = [c.v for c in r[:7]]
                if vals[2] is None or vals[0] in (None, "Não tem Funcionários"):
                    continue
                try:
                    loja = int(float(vals[2]))
                    sind = int(float(vals[0]))
                    sind_desc = str(vals[1]).strip()
                    loc_lg = int(float(vals[3])) if vals[3] is not None else 0
                    cidade = str(vals[4]).strip()
                    uf = str(vals[5]).strip()
                    fer_raw = str(vals[6]).strip()

                    datas_feriado = set()
                    if fer_raw == "46272.0":
                        datas_feriado.add("2026-09-07")
                    elif fer_raw != "Não Tem":
                        for part in re.findall(r"\d{2}/\d{2}", fer_raw):
                            dd, mm = part.split("/")
                            ano = "2026"
                            datas_feriado.add(f"{ano}-{mm}-{dd}")

                    lojas_info[loja] = {
                        "sindicato": sind,
                        "sindicato_nome": sind_desc,
                        "localidade": loc_lg,
                        "municipio": cidade,
                        "uf": uf,
                        "feriados_str": "07/09" if fer_raw == "46272.0" else fer_raw,
                        "datas_feriado": datas_feriado,
                        "checkouts": checkouts_map.get(loja, "Acima de 4")
                    }
                except Exception:
                    pass

    # 4. Base Cadastral de Colaboradores (COLA_Colaboradores_CSV_Gestor)
    print("-> [2/4] Carregando base cadastral de Colaboradores (CSV Gestor)...")
    cola_map = {}
    with open(FILE_COLA_CSV, mode="r", encoding="latin1") as f_cola:
        next(f_cola, None)
        reader = csv.DictReader(f_cola, delimiter=";")
        for row in reader:
            m = str(row.get("Matrícula", "")).strip()
            if m:
                try:
                    cola_map[int(m)] = {
                        "nome": str(row.get("Colaborador", "")).strip(),
                        "cargo": str(row.get("Cargo", "")).strip(),
                        "situacao": str(row.get("Situação", "Atividade Normal")).strip(),
                        "sindicato_cad": str(row.get("Código Sindicato", "")).strip(),
                        "estab_nome": str(row.get("Estabelecimento", "")).strip()
                    }
                except Exception:
                    pass

    # 5. Planilha Manual da Equipe (GRATIFICAÇÃO 10.2026.xlsx)
    print("-> [3/4] Carregando planilha manual GRATIFICAÇÃO 10.2026.xlsx...")
    wb_g = openpyxl.load_workbook(FILE_GRAT_MANUAL, data_only=True)
    ws_g = wb_g["Gratificação 10.2026"]
    grat_manual_map = {}
    grat_manual_list = []
    for idx, r in enumerate(ws_g.iter_rows(min_row=2, values_only=True), start=2):
        if r[0] is None:
            continue
        emp, loja, sind, mat, nome, cargo, dt, dia, marc, ht, regra, co, val = r[:13]
        dt_str = str(dt)[:10]
        chave = (int(loja), int(mat), dt_str)
        item = {
            "row_idx": idx,
            "empresa": int(emp),
            "loja": int(loja),
            "sindicato": int(sind),
            "matricula": int(mat),
            "nome": str(nome).strip(),
            "cargo": str(cargo).strip(),
            "data_iso": dt_str,
            "dia_sem": str(dia).strip(),
            "marcacoes": str(marc).strip() if marc else "",
            "ht": str(ht).strip() if ht else "00:00",
            "regra": str(regra).strip() if regra else "Horas",
            "checkouts": str(co).strip() if co else "",
            "valor": float(val) if val is not None else 0.0
        }
        grat_manual_map[chave] = item
        grat_manual_list.append(item)
    wb_g.close()

    return lojas_info, feriados_desc_map, cola_map, grat_manual_map, grat_manual_list

# ======================================================================================
# 5. PROCESSAMENTO E AUDITORIA DO PONTO VS. GRATIFICAÇÃO MANUAL
# ======================================================================================

def processar_auditoria(lojas_info, feriados_desc_map, cola_map, grat_manual_map, grat_manual_list):
    print("-> [4/4] Auditando Relatorio_HorasTrabalhadas contra Regras CCT e Planilha Manual...")

    wb_ht = openpyxl.load_workbook(FILE_HT, read_only=True, data_only=True)
    ws_ht = wb_ht.active

    ht_index = {}
    lojas_no_ht = defaultdict( lambda: {"total_linhas": 0, "com_batida": 0, "estab_nome": ""} )

    for r in ws_ht.iter_rows(min_row=2, values_only=True):
        if r[0] is None or str(r[0]).strip() == "Colaborador":
            continue
        try:
            mat = int(str(r[0]).strip())
        except Exception:
            continue
        nome_ht = str(r[1]).strip() if r[1] else ""
        cargo_ht = str(r[2]).strip() if r[2] else ""
        estab_full = str(r[4]).strip() if r[4] else ""
        m_loja = re.search(r"(\d+)\s*$", estab_full)
        loja = int(m_loja.group(1)) if m_loja else 0
        dt_str = str(r[7])[:10]
        marc = str(r[8]).strip() if r[8] else ""
        ht_val = str(r[9]).strip() if r[9] else "00:00"
        he_val = str(r[10]).strip() if r[10] else "00:00"

        lojas_no_ht[loja]["total_linhas"] += 1
        lojas_no_ht[loja]["estab_nome"] = estab_full
        if marc:
            lojas_no_ht[loja]["com_batida"] += 1

        ht_index[(loja, mat, dt_str)] = {
            "nome": nome_ht,
            "cargo": cargo_ht,
            "estab_full": estab_full,
            "marcacoes": marc,
            "ht": ht_val,
            "he": he_val
        }
    wb_ht.close()

    linhas_geral = []
    linhas_divergencias = []
    consolidado_colab = {}
    stats_loja = defaultdict(lambda: {
        "qtd_dom": 0, "val_dom": 0.0,
        "qtd_fer": 0, "val_fer": 0.0,
        "qtd_tot": 0, "val_planilha": 0.0, "val_auditado": 0.0,
        "alertas": 0
    })

    kpis = {
        "total_ocorrencias": 0,
        "total_colaboradores": 0,
        "valor_planilha": 0.0,
        "valor_auditado": 0.0,
        "conformes": 0,
        "alertas": 0,
        "qtd_feriados": 0,
        "valor_feriados": 0.0,
        "qtd_domingos": 0,
        "valor_domingos": 0.0
    }

    # Ordenar registros por Loja, Nome, Data
    grat_sorted = sorted(grat_manual_list, key=lambda x: (x["loja"], x["nome"], x["data_iso"]))

    for item in grat_sorted:
        loja = item["loja"]
        mat = item["matricula"]
        dt_iso = item["data_iso"]
        sind = item["sindicato"]
        chave = (loja, mat, dt_iso)

        l_info = lojas_info.get(loja, {})
        municipio = l_info.get("municipio", "NÃO IDENTIFICADO")
        sind_nome = l_info.get("sindicato_nome", f"SINDICATO {sind}")
        loc_lg = l_info.get("localidade", 0)
        co_porte = l_info.get("checkouts", "Acima de 4")

        c_info = cola_map.get(mat, {})
        situacao_cad = c_info.get("situacao", "Atividade Normal")
        estab_nome = lojas_no_ht[loja]["estab_nome"] or c_info.get("estab_nome", f"LOJA {loja:03d}")

        dt_obj = datetime.strptime(dt_iso, "%Y-%m-%d")
        dia_sem = DIAS_SEMANA_PT[dt_obj.weekday()]
        is_domingo = (dt_obj.weekday() == 6)

        if is_domingo:
            tipo_dia = "Domingo Trabalhado"
        else:
            desc_fer = feriados_desc_map.get((loc_lg, dt_iso), "Feriado Municipal / Nacional")
            tipo_dia = f"Feriado ({desc_fer})"

        marc_str = item["marcacoes"]
        ht_str = item["ht"]
        mins_ht = parse_minutes(ht_str)

        qtd_bat, is_impar, mins_calc, str_calc, ult_bat = analisar_batidas(marc_str)

        # Valor pela Hora Trabalhada do Sistema vs Valor pelas Marcações Reais Calculadas
        val_planilha = item["valor"]
        val_cct_ht, faixa_ht = calcular_valor_cct(sind, is_domingo, mins_ht, co_porte, ult_bat)
        val_cct_calc, faixa_calc = calcular_valor_cct(sind, is_domingo, mins_calc if not is_impar else mins_ht, co_porte, ult_bat)

        diff_mins = abs(mins_calc - mins_ht)
        # Se a diferença entre batidas e HT estiver dentro da tolerância CLT (<= 15 min), prevalece o HT apurado no ponto
        val_cct_final = val_cct_calc if (not is_impar and diff_mins > 15) else val_cct_ht
        diff_val = round(val_cct_final - val_planilha, 2)

        # Regras de Auditoria
        if is_impar:
            status_aud = "Marcação Ímpar (Falta Batida)"
            severidade = "Alerta"
            parecer = (
                f"Colaborador com {qtd_bat} batidas ('{marc_str}'). O sistema apurou apenas {ht_str} "
                f"(R$ {val_planilha:.2f}). Verificar batida faltante pois jornada real pode alterar a faixa para R$ 94,68 ou R$ 100,53."
            )
        elif diff_mins > 15 and abs(diff_val) > 0.01:
            if val_cct_final > val_planilha:
                status_aud = "Divergência de Faixa (HT < Marcações)"
                severidade = "Alerta"
                parecer = (
                    f"As marcações ('{marc_str}') totalizam {str_calc} trabalhadas ({faixa_calc} = R$ {val_cct_final:.2f}), "
                    f"porém HorasTrabalhadas registrou {ht_str} (pago R$ {val_planilha:.2f}; diferença de +R$ {diff_val:.2f} ao colaborador)."
                )
            else:
                status_aud = "Divergência de Faixa (HT > Marcações)"
                severidade = "Alerta"
                parecer = (
                    f"As marcações ('{marc_str}') totalizam apenas {str_calc} ({faixa_calc} = R$ {val_cct_final:.2f}), "
                    f"porém HorasTrabalhadas consta {ht_str} (pago R$ {val_planilha:.2f}; diferença de R$ {diff_val:.2f}). Validar se houve abono parcial."
                )
        elif abs(val_cct_ht - val_planilha) > 0.01:
            status_aud = "Valor Divergente da Tabela CCT"
            severidade = "Crítico"
            parecer = f"Valor na planilha (R$ {val_planilha:.2f}) diverge da tabela sindical (R$ {val_cct_ht:.2f} - {faixa_ht})."
        else:
            status_aud = "Conforme"
            severidade = "Regular"
            parecer = f"Valor validado conforme CCT ({faixa_ht}). Marcações e carga horária consistentes."

        # Acumular KPIs
        kpis["total_ocorrencias"] += 1
        kpis["valor_planilha"] += val_planilha
        kpis["valor_auditado"] += val_cct_final

        st = stats_loja[loja]
        st["qtd_tot"] += 1
        st["val_planilha"] += val_planilha
        st["val_auditado"] += val_cct_final

        if is_domingo:
            kpis["qtd_domingos"] += 1
            kpis["valor_domingos"] += val_planilha
            st["qtd_dom"] += 1
            st["val_dom"] += val_planilha
        else:
            kpis["qtd_feriados"] += 1
            kpis["valor_feriados"] += val_planilha
            st["qtd_fer"] += 1
            st["val_fer"] += val_planilha

        if severidade == "Regular":
            kpis["conformes"] += 1
        else:
            kpis["alertas"] += 1
            st["alertas"] += 1

        linha = [
            item["empresa"],
            loja,
            municipio,
            sind,
            sind_nome,
            mat,
            item["nome"],
            item["cargo"],
            situacao_cad,
            format_data_br(dt_iso),
            dia_sem,
            tipo_dia,
            marc_str,
            ht_str,
            str_calc,
            faixa_ht,
            co_porte,
            val_planilha,
            val_cct_final,
            diff_val,
            status_aud,
            severidade,
            parecer
        ]
        linhas_geral.append(linha)
        if severidade != "Regular":
            linhas_divergencias.append(linha)

        # Acumular no Consolidado por Colaborador
        if mat not in consolidado_colab:
            consolidado_colab[mat] = {
                "empresa": item["empresa"],
                "loja": loja,
                "municipio": municipio,
                "sindicato": sind,
                "matricula": mat,
                "nome": item["nome"],
                "cargo": item["cargo"],
                "qtd_dom": 0,
                "qtd_fer": 0,
                "total_dias": 0,
                "val_planilha": 0.0,
                "val_auditado": 0.0,
                "tem_alerta": False,
                "obs": []
            }
        cc = consolidado_colab[mat]
        if is_domingo:
            cc["qtd_dom"] += 1
        else:
            cc["qtd_fer"] += 1
        cc["total_dias"] += 1
        cc["val_planilha"] += val_planilha
        cc["val_auditado"] += val_cct_final
        if severidade != "Regular":
            cc["tem_alerta"] = True
            cc["obs"].append(f"{format_data_br(dt_iso)}: {status_aud}")

    kpis["total_colaboradores"] = len(consolidado_colab)
    return linhas_geral, linhas_divergencias, consolidado_colab, stats_loja, lojas_no_ht, kpis

# ======================================================================================
# 6. CONSTRUÇÃO DAS 5 ABAS DO WORKBOOK OFICIAL CENCOSUD
# ======================================================================================

def construir_aba_orientacoes(ws):
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:K1")
    c1 = ws["A1"]
    c1.value = "CENCOSUD BRASIL — AUDITORIA E APURAÇÃO DE GRATIFICAÇÃO MANUAL (DOMINGOS E FERIADOS)"
    c1.fill = FILL_HEADER
    c1.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    c1.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 38

    ws.merge_cells("A2:K2")
    c2 = ws["A2"]
    c2.value = "GBarbosa, Eletroshow & Mercantil — Eventos 227 (Provento Alimentação) e 5751 (Desconto Antecipação) | Competência Folha 10/2026 (Ponto 09/2026)"
    c2.fill = FILL_HEADER_SEC
    c2.font = Font(name="Calibri", size=11, italic=True, color="FFFFFF")
    c2.alignment = ALIGN_CENTER
    ws.row_dimensions[2].height = 24

    secoes = [
        ("1. OBJETIVO DO MÓDULO DE GRATIFICAÇÃO MANUAL",
         "Automatizar a apuração e auditar 100% dos lançamentos de Gratificação de Domingos e Feriados (Eventos 227 e 5751) das regionais GBarbosa e Mercantil (Bahia), "
         "cruzando as batidas reais do Relatório de Horas Trabalhadas com o Calendário Oficial de Feriados Municipais/Nacionais e as Convenções Coletivas de Trabalho (CCT) de todos os 8 sindicatos padrão da planilha de regras (6, 9, 36, 40, 60, 71, 74 e 82)."),

        ("2. TABELA DE REGRAS SINDICAIS PADRÃO (TODOS OS 8 SINDICATOS: 6, 9, 36, 40, 60, 71, 74 E 82)",
         "• SINDICATO 6 (SIN EMPREG COMER ALAGOINHAS — Alagoinhas, Cipó, Rio Real, Cansanção, Entre Rios, Olindina, Monte Santo, Queimadas, Santaluz, Crisópolis):\n"
         "   - Domingos (Regra Check-outs): Até 4 Check-outs = R$ 73,50 | Com mais de 4 Check-outs = R$ 79,80.\n"
         "   - Feriados (Regra Horas Trabalhadas): Até 6h = R$ 87,19 | Até 07:20 = R$ 94,68 | Jornada 8h = R$ 100,53.\n"
         "• SINDICATO 9 (SIN EMPR ESTAB COMER SERV PAULO AFONSO E REGIÃO — Paulo Afonso, Tucano, Euclides da Cunha, Jeremoabo, Paripiranga):\n"
         "   - Domingos (Regra Horas Trabalhadas): 5h = R$ 57,00 | 6h = R$ 64,00 | 8h = R$ 90,00.\n"
         "   - Feriados (Regra Horas Trabalhadas): Até 5h = R$ 57,00 | Até 6h = R$ 64,00 | Jornada 8h = R$ 90,00.\n"
         "• SINDICATO 36 (SIN EMP EST COM AFONSO E REG R. DO POMBAL — Ribeira do Pombal e Cícero Dantas):\n"
         "   - Domingos (Regra Horas Trabalhadas): 6h = R$ 50,00 | 8h = R$ 60,00.\n"
         "   - Feriados (Regra Horas Trabalhadas): Até 6h = R$ 52,00 | Jornada 8h = R$ 71,00.\n"
         "• SINDICATO 40 (SIN EMP COM DA CIDADE DE JUAZEIRO E REGIÃO — Juazeiro e Campo Formoso):\n"
         "   - Domingos (Regra Horário Saída + Horas): Domingo normal = R$ 70,44 | Até 6h (Após 18:30) = R$ 79,03 | Acima 6h (Após 18:30) = R$ 96,01.\n"
         "   - Feriados (Regra Horas Trabalhadas): Até 6h = R$ 80,34 | Jornada 8h = R$ 101,46.\n"
         "• SINDICATO 60 (SIN EMP COMER VITÓRIA DA CONQUISTA — Vitória da Conquista):\n"
         "   - Domingos (Regra Horas Trabalhadas): 6h = R$ 33,65 | 07:20 = R$ 45,60.\n"
         "   - Feriados (Regra Porte Nacional): Até 1.000 func. nacional = R$ 60,85 | Acima de 1.000 func. nacional = R$ 68,40.\n"
         "• SINDICATO 71 (SIN DOS EMP NO COMER ALAGOINHAS E REGIÃO):\n"
         "   - Domingos (Regra Check-outs): Até 4 Check-outs = R$ 70,00 | Com mais de 4 Check-outs = R$ 76,00.\n"
         "   - Feriados (Regra Horas Trabalhadas): Até 6h = R$ 83,04 | Até 07:20 = R$ 90,18 | Jornada 8h = R$ 95,75.\n"
         "• SINDICATO 74 (SIND EMPREG NO COM DE CATU — Catu):\n"
         "   - Domingos (Regra Check-outs): Até 4 Check-outs = R$ 70,35 | Com mais de 4 Check-outs = R$ 80,00.\n"
         "   - Feriados (Regra Horas Trabalhadas): Até 6h = R$ 84,73 | Até 07:20 = R$ 93,45 | Jornada 8h = R$ 103,40.\n"
         "• SINDICATO 82 (SINDICATO EMP COMER CONCEIÇÃO DO JACUÍPE — Conceição do Jacuípe):\n"
         "   - Domingos (Regra Porte Loja): Até 50 empregados = R$ 27,45 | Mais de 50 empregados = R$ 36,70.\n"
         "   - Feriados (Regra Porte + Horário 14h): Até 50 func até 14:00 = R$ 35,00 | 51+ até 14:00 = R$ 53,70 | Até 50 após 14:00 = R$ 51,00 | 51+ após 14:00 = R$ 81,40."),

        ("3. METODOLOGIA DE AUDITORIA CRUZADA (PONTO X FOLHA)",
         "• Parametrização Completa (8 Sindicatos Padrão): O motor mantém parametrizados todos os 8 sindicatos da planilha 'Regras das Gratificações GB e Mercantil.xlsx' (6, 9, 36, 40, 60, 71, 74 e 82), aplicando automaticamente a regra de cada sindicato independentemente de ter havido ocorrência no mês de exemplo.\n"
         "• Recálculo Matemático das Marcações: Confronta o somatório real dos pares de batidas contra a coluna 'HorasTrabalhadas' do sistema, sinalizando casos em que batidas ímpares ou parciais alteram a faixa de pagamento.\n"
         "• Consolidação Pronta para Folha: Gera automaticamente a soma individual por matrícula (Aba 5) e o Quadro de Autorização para os Eventos 227 e 5751.")
    ]

    cur_row = 4
    for titulo, texto in secoes:
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=11)
        ct = ws.cell(row=cur_row, column=1, value=titulo)
        ct.fill = FILL_HEADER_SEC
        ct.font = FONT_SUBTITULO
        ct.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[cur_row].height = 25
        cur_row += 1

        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=11)
        cx = ws.cell(row=cur_row, column=1, value=texto)
        cx.fill = FILL_ZEBRA
        cx.font = FONT_DADOS
        cx.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        linhas_txt = texto.count("\n") + len(texto) // 115 + 1
        ws.row_dimensions[cur_row].height = max(48, linhas_txt * 15)
        cur_row += 2

    # TABELA ESTRUTURADA DOS 8 SINDICATOS PADRÃO NA ABA ORIENTAÇÕES
    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=11)
    ct_tab = ws.cell(row=cur_row, column=1, value="4. MATRIZ PADRÃO OFICIAL DE REGRAS SINDICAIS (SINDICATOS 6, 9, 36, 40, 60, 71, 74 E 82)")
    ct_tab.fill = FILL_HEADER_SEC
    ct_tab.font = FONT_SUBTITULO
    ct_tab.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[cur_row].height = 25
    cur_row += 1

    # Colunas da Matriz: A=Cód, B-C=Entidade Sindical, D-E=Municípios, F-G=Regra Domingos, H-I=Regra Feriados, J-K=Feriados Proibidos
    cols_m = [(1, 1, "Cód. Sind."), (2, 3, "Entidade Sindical"), (4, 5, "Municípios Abrangidos"), (6, 7, "Regra e Valores — Domingos"), (8, 9, "Regra e Valores — Feriados"), (10, 11, "Feriados Proibidos (CCT)")]
    ws.row_dimensions[cur_row].height = 24
    for ci, cf, h_txt in cols_m:
        if ci != cf:
            ws.merge_cells(start_row=cur_row, start_column=ci, end_row=cur_row, end_column=cf)
        c_h = ws.cell(row=cur_row, column=ci, value=h_txt)
        c_h.alignment = ALIGN_CENTER
        for c_i in range(ci, cf + 1):
            cell_tmp = ws.cell(row=cur_row, column=c_i)
            cell_tmp.fill = FILL_HEADER
            cell_tmp.font = FONT_CABECALHO
            cell_tmp.border = BORDER_THIN
    cur_row += 1

    for idx_s, s_info in enumerate(REGRAS_SINDICATOS_PADRAO):
        fill_r = FILL_ZEBRA if idx_s % 2 == 1 else FILL_BRANCO
        ws.row_dimensions[cur_row].height = 32
        vals_m = [
            (1, 1, s_info["cod"], ALIGN_CENTER, FONT_DADOS_BOLD),
            (2, 3, s_info["nome"], ALIGN_LEFT, FONT_DADOS_BOLD),
            (4, 5, s_info["municipios"], ALIGN_LEFT, FONT_DADOS),
            (6, 7, s_info["regra_dom"], ALIGN_LEFT, FONT_DADOS),
            (8, 9, s_info["regra_fer"], ALIGN_LEFT, FONT_DADOS),
            (10, 11, s_info["feriados_proibidos"], ALIGN_LEFT, FONT_DADOS)
        ]
        for ci, cf, val_m, al_m, fn_m in vals_m:
            if ci != cf:
                ws.merge_cells(start_row=cur_row, start_column=ci, end_row=cur_row, end_column=cf)
            c_v = ws.cell(row=cur_row, column=ci, value=val_m)
            c_v.alignment = Alignment(horizontal=al_m.horizontal, vertical="center", wrap_text=True)
            c_v.font = fn_m
            for c_i in range(ci, cf + 1):
                cell_tmp = ws.cell(row=cur_row, column=c_i)
                cell_tmp.fill = fill_r
                cell_tmp.border = BORDER_THIN
        cur_row += 1

    for c in range(1, 12):
        ws.column_dimensions[get_column_letter(c)].width = 16

def construir_aba_resumo(ws, kpis, stats_loja, lojas_info, lojas_no_ht):
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:L1")
    ct = ws["A1"]
    ct.value = "CENCOSUD BRASIL — DASHBOARD EXECUTIVO DE GRATIFICAÇÃO MANUAL (DOMINGOS E FERIADOS)"
    ct.fill = FILL_HEADER
    ct.font = FONT_TITULO
    ct.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 36

    taxa_conf = (kpis["conformes"] / kpis["total_ocorrencias"]) if kpis["total_ocorrencias"] > 0 else 1.0

    cards_l1 = [
        ("OCORRÊNCIAS ELEGÍVEIS", f"{kpis['total_ocorrencias']:,}", f"{kpis['total_colaboradores']} Colaboradores Únicos", C_AZUL_MARINHO, 1, 3),
        ("VALOR TOTAL PLANILHA (EVT 227)", f"R$ {kpis['valor_planilha']:,.2f}", "Gratificação Domingos e Feriados", "166534", 4, 6),
        ("VALOR RECALCULADO (AUDITORIA)", f"R$ {kpis['valor_auditado']:,.2f}", f"Diferença Líquida: R$ {kpis['valor_auditado'] - kpis['valor_planilha']:+.2f}", C_AZUL_MEDIO, 7, 9),
        ("TAXA DE CONFORMIDADE", f"{taxa_conf:.1%}", f"{kpis['conformes']} Regulares | {kpis['alertas']} Alertas", C_ALERT_TXT if kpis["alertas"] > 0 else "166534", 10, 12)
    ]

    cards_l2 = [
        ("FERIADOS TRABALHADOS (NAC./MUN.)", f"{kpis['qtd_feriados']:,} ocorr.", f"Total: R$ {kpis['valor_feriados']:,.2f}", C_AZUL_MARINHO, 1, 4),
        ("DOMINGOS APURAÇÃO MANUAL", f"{kpis['qtd_domingos']:,} ocorr.", f"Total: R$ {kpis['valor_domingos']:,.2f} (Lojas 32 e 133)", C_AZUL_MARINHO, 5, 8),
        ("COBERTURA DE LOJAS NA BASE", f"{len(stats_loja)} / {len(lojas_no_ht)} Lojas", "16 Lojas sem expediente nos feriados", C_ALERT_TXT, 9, 12)
    ]

    def render_card(titulo, valor, sub, cor_txt, ci, cf, r_base):
        ws.merge_cells(start_row=r_base, start_column=ci, end_row=r_base, end_column=cf)
        c_t = ws.cell(row=r_base, column=ci, value=titulo)
        c_t.fill = FILL_TOTAL
        c_t.font = Font(name="Calibri", size=9, bold=True, color="475569")
        c_t.alignment = ALIGN_CENTER

        ws.merge_cells(start_row=r_base + 1, start_column=ci, end_row=r_base + 1, end_column=cf)
        c_v = ws.cell(row=r_base + 1, column=ci, value=valor)
        c_v.fill = FILL_BRANCO
        c_v.font = Font(name="Calibri", size=16, bold=True, color=cor_txt)
        c_v.alignment = ALIGN_CENTER

        ws.merge_cells(start_row=r_base + 2, start_column=ci, end_row=r_base + 2, end_column=cf)
        c_s = ws.cell(row=r_base + 2, column=ci, value=sub)
        c_s.fill = FILL_BRANCO
        c_s.font = Font(name="Calibri", size=8, italic=True, color="64748B")
        c_s.alignment = ALIGN_CENTER

        for r in range(r_base, r_base + 3):
            for c in range(ci, cf + 1):
                ws.cell(row=r, column=c).border = BORDER_THIN

    for t, v, s, cor, ci, cf in cards_l1:
        render_card(t, v, s, cor, ci, cf, 3)
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 28
    ws.row_dimensions[5].height = 16

    for t, v, s, cor, ci, cf in cards_l2:
        render_card(t, v, s, cor, ci, cf, 7)
    ws.row_dimensions[7].height = 18
    ws.row_dimensions[8].height = 28
    ws.row_dimensions[9].height = 16

    # TABELA 1: QUADRO DE AUTORIZAÇÃO GERENCIAL (ESPELHO DA ABA AUTORIZAÇÃO)
    cur_r = 11
    ws.merge_cells(f"A{cur_r}:L{cur_r}")
    t_aut = ws[f"A{cur_r}"]
    t_aut.value = "1. QUADRO DE AUTORIZAÇÃO GERENCIAL PARA FOLHA DE PAGAMENTO (EVENTOS 227 E 5751)"
    t_aut.fill = FILL_HEADER_SEC
    t_aut.font = FONT_SUBTITULO
    t_aut.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[cur_r].height = 24
    cur_r += 1

    ws.merge_cells(f"A{cur_r}:E{cur_r}")
    ws.cell(row=cur_r, column=1, value="Evento da Folha / Descrição").alignment = ALIGN_CENTER
    ws.merge_cells(f"F{cur_r}:G{cur_r}")
    ws.cell(row=cur_r, column=6, value="Qtd. Colaboradores").alignment = ALIGN_CENTER
    ws.merge_cells(f"H{cur_r}:I{cur_r}")
    ws.cell(row=cur_r, column=8, value="Valor Planilha (R$)").alignment = ALIGN_CENTER
    ws.merge_cells(f"J{cur_r}:K{cur_r}")
    ws.cell(row=cur_r, column=10, value="Valor Auditado (R$)").alignment = ALIGN_CENTER
    ws.cell(row=cur_r, column=12, value="Qtd. Ocorrências").alignment = ALIGN_CENTER
    for c in range(1, 13):
        ws.cell(row=cur_r, column=c).fill = FILL_HEADER
        ws.cell(row=cur_r, column=c).font = FONT_CABECALHO
        ws.cell(row=cur_r, column=c).border = BORDER_THIN
    ws.row_dimensions[cur_r].height = 24
    cur_r += 1

    eventos_aut = [
        ("Gratificação GB e MR - Domingos e Feriados (EVT 227 - Provento)", kpis["total_colaboradores"], kpis["valor_planilha"], kpis["valor_auditado"], kpis["total_ocorrencias"]),
        ("Gratificação GB e MR - Domingos e Feriados (EVT 5751 - Desconto Antecipação)", kpis["total_colaboradores"], kpis["valor_planilha"], kpis["valor_auditado"], kpis["total_ocorrencias"])
    ]
    for desc_ev, q_col, v_plan, v_aud, q_oc in eventos_aut:
        ws.row_dimensions[cur_r].height = 21
        ws.merge_cells(f"A{cur_r}:E{cur_r}")
        c_d = ws.cell(row=cur_r, column=1, value=desc_ev)
        c_d.font = FONT_DADOS_BOLD
        c_d.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws.merge_cells(f"F{cur_r}:G{cur_r}")
        c_q = ws.cell(row=cur_r, column=6, value=q_col)
        c_q.font = FONT_DADOS_BOLD
        c_q.alignment = ALIGN_CENTER
        c_q.number_format = "#,##0"

        ws.merge_cells(f"H{cur_r}:I{cur_r}")
        c_vp = ws.cell(row=cur_r, column=8, value=v_plan)
        c_vp.font = FONT_DADOS_BOLD
        c_vp.alignment = ALIGN_RIGHT
        c_vp.number_format = "R$ #,##0.00"

        ws.merge_cells(f"J{cur_r}:K{cur_r}")
        c_va = ws.cell(row=cur_r, column=10, value=v_aud)
        c_va.font = FONT_DADOS_BOLD
        c_va.alignment = ALIGN_RIGHT
        c_va.number_format = "R$ #,##0.00"

        c_qo = ws.cell(row=cur_r, column=12, value=q_oc)
        c_qo.font = FONT_DADOS
        c_qo.alignment = ALIGN_CENTER
        c_qo.number_format = "#,##0"

        for c in range(1, 13):
            ws.cell(row=cur_r, column=c).fill = FILL_BRANCO
            ws.cell(row=cur_r, column=c).border = BORDER_THIN
        cur_r += 1

    cur_r += 2

    # TABELA 2: TABELA PADRÃO DE REGRAS SINDICAIS E CONSOLIDAÇÃO POR SINDICATO (TODOS OS 8 SINDICATOS: 6, 9, 36, 40, 60, 71, 74 E 82)
    ws.merge_cells(f"A{cur_r}:L{cur_r}")
    t_s = ws[f"A{cur_r}"]
    t_s.value = "2. TABELA PADRÃO DE REGRAS SINDICAIS & CONSOLIDAÇÃO POR SINDICATO (SINDICATOS 6, 9, 36, 40, 60, 71, 74 E 82)"
    t_s.fill = FILL_HEADER_SEC
    t_s.font = FONT_SUBTITULO
    t_s.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[cur_r].height = 24
    cur_r += 1

    headers_sind = [
        "Cód. Sind.", "Entidade Sindical", "Regra Padrão — Domingos", "Regra Padrão — Feriados",
        "Qtd. Dom.", "Valor Domingos", "Qtd. Fer.", "Valor Feriados",
        "Total Ocorr.", "Valor Planilha (R$)", "Valor Auditado (R$)", "Status no Mês"
    ]
    ws.row_dimensions[cur_r].height = 25
    for idx, h in enumerate(headers_sind, start=1):
        c = ws.cell(row=cur_r, column=idx, value=h)
        c.fill = FILL_HEADER
        c.font = FONT_CABECALHO
        c.alignment = ALIGN_CENTER
        c.border = BORDER_THIN
    cur_r += 1

    # Agregar estatísticas por sindicato a partir das lojas
    stats_sind = defaultdict(lambda: {"qtd_dom": 0, "val_dom": 0.0, "qtd_fer": 0, "val_fer": 0.0, "qtd_tot": 0, "val_planilha": 0.0, "val_auditado": 0.0, "alertas": 0})
    for loja_k, st_l in stats_loja.items():
        s_cod = lojas_info.get(loja_k, {}).get("sindicato", 0)
        ss = stats_sind[s_cod]
        ss["qtd_dom"] += st_l["qtd_dom"]
        ss["val_dom"] += st_l["val_dom"]
        ss["qtd_fer"] += st_l["qtd_fer"]
        ss["val_fer"] += st_l["val_fer"]
        ss["qtd_tot"] += st_l["qtd_tot"]
        ss["val_planilha"] += st_l["val_planilha"]
        ss["val_auditado"] += st_l["val_auditado"]
        ss["alertas"] += st_l["alertas"]

    for s_info in REGRAS_SINDICATOS_PADRAO:
        scod = s_info["cod"]
        ss = stats_sind[scod]
        fill_row = FILL_ZEBRA if (cur_r % 2 == 0) else FILL_BRANCO
        if ss["qtd_tot"] == 0:
            st_sind_str = "Padrão Ativo (Sem Ocorr. no Mês)"
        elif ss["alertas"] > 0:
            st_sind_str = f"Atenção ({ss['alertas']} Alerta)"
        else:
            st_sind_str = "100% Conforme"

        vals_s = [
            scod,
            s_info["nome"],
            s_info["regra_dom"],
            s_info["regra_fer"],
            ss["qtd_dom"],
            ss["val_dom"],
            ss["qtd_fer"],
            ss["val_fer"],
            ss["qtd_tot"],
            ss["val_planilha"],
            ss["val_auditado"],
            st_sind_str
        ]
        ws.row_dimensions[cur_r].height = 28
        for idx, val in enumerate(vals_s, start=1):
            c = ws.cell(row=cur_r, column=idx, value=val)
            c.fill = fill_row
            c.border = BORDER_THIN
            c.font = FONT_DADOS_BOLD if idx in (1, 2, 9, 10, 11) else FONT_DADOS
            if idx in (1, 5, 7, 9):
                c.alignment = ALIGN_CENTER
            elif idx in (6, 8, 10, 11):
                c.alignment = ALIGN_RIGHT
                c.number_format = "R$ #,##0.00"
            elif idx == 12:
                c.alignment = ALIGN_CENTER
                if "100% Conforme" in st_sind_str:
                    c.fill = FILL_REG
                    c.font = FONT_REG
                elif "Atenção" in st_sind_str:
                    c.fill = FILL_ALERT
                    c.font = FONT_ALERT
            else:
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cur_r += 1

    # Linha Total dos 8 Sindicatos Padrão
    tot_sind_vals = [
        "TOTAL (8 SIND.)", "TODOS OS SINDICATOS DA REGRA (6, 9, 36, 40, 60, 71, 74 E 82)", "-", "-",
        kpis["qtd_domingos"], kpis["valor_domingos"],
        kpis["qtd_feriados"], kpis["valor_feriados"],
        kpis["total_ocorrencias"], kpis["valor_planilha"], kpis["valor_auditado"],
        f"{kpis['alertas']} Alertas" if kpis["alertas"] > 0 else "100% Conforme"
    ]
    ws.row_dimensions[cur_r].height = 24
    for idx, val in enumerate(tot_sind_vals, start=1):
        c = ws.cell(row=cur_r, column=idx, value=val)
        c.fill = FILL_TOTAL
        c.font = Font(name="Calibri", size=10, bold=True, color=C_AZUL_MARINHO)
        c.border = BORDER_TOTAL
        if idx in (1, 3, 4, 5, 7, 9, 12):
            c.alignment = ALIGN_CENTER
        elif idx in (6, 8, 10, 11):
            c.alignment = ALIGN_RIGHT
            c.number_format = "R$ #,##0.00"
        else:
            c.alignment = ALIGN_LEFT
    cur_r += 3

    # TABELA 3: CONSOLIDAÇÃO POR LOJA, MUNICÍPIO E SINDICATO
    ws.merge_cells(f"A{cur_r}:L{cur_r}")
    t_l = ws[f"A{cur_r}"]
    t_l.value = "3. CONSOLIDAÇÃO DE GRATIFICAÇÕES POR LOJA, MUNICÍPIO E SINDICATO (28 LOJAS DA BASE)"
    t_l.fill = FILL_HEADER_SEC
    t_l.font = FONT_SUBTITULO
    t_l.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[cur_r].height = 24
    cur_r += 1

    headers_t2 = [
        "Loja", "Município (BA)", "Cód. Sind.", "Feriados no Mês", "Porte Check-outs",
        "Qtd. Domingos", "Valor Domingos", "Qtd. Feriados", "Valor Feriados",
        "Total Ocorr.", "Valor Planilha (R$)", "Status Loja"
    ]
    ws.row_dimensions[cur_r].height = 25
    for idx, h in enumerate(headers_t2, start=1):
        c = ws.cell(row=cur_r, column=idx, value=h)
        c.fill = FILL_HEADER
        c.font = FONT_CABECALHO
        c.alignment = ALIGN_CENTER
        c.border = BORDER_THIN
    cur_r += 1

    todas_lojas_ordenadas = sorted(lojas_no_ht.keys(), key=lambda l: (0 if l in stats_loja else 1, l))
    for loja in todas_lojas_ordenadas:
        li = lojas_info.get(loja, {})
        st = stats_loja.get(loja, {"qtd_dom": 0, "val_dom": 0.0, "qtd_fer": 0, "val_fer": 0.0, "qtd_tot": 0, "val_planilha": 0.0, "alertas": 0})
        fill_row = FILL_ZEBRA if (cur_r % 2 == 0) else FILL_BRANCO

        if st["qtd_tot"] == 0:
            status_l = "Sem Expediente no Feriado"
        elif st["alertas"] > 0:
            status_l = f"Atenção ({st['alertas']} Alerta)"
        else:
            status_l = "100% Conforme"

        vals = [
            f"Loja {loja:03d}",
            li.get("municipio", "-"),
            li.get("sindicato", "-"),
            li.get("feriados_str", "07/09"),
            li.get("checkouts", "Acima de 4"),
            st["qtd_dom"],
            st["val_dom"],
            st["qtd_fer"],
            st["val_fer"],
            st["qtd_tot"],
            st["val_planilha"],
            status_l
        ]
        ws.row_dimensions[cur_r].height = 20
        for idx, val in enumerate(vals, start=1):
            c = ws.cell(row=cur_r, column=idx, value=val)
            c.fill = fill_row
            c.border = BORDER_THIN
            c.font = FONT_DADOS_BOLD if idx in (1, 10, 11) else FONT_DADOS
            if idx in (1, 3, 4, 5, 6, 8, 10):
                c.alignment = ALIGN_CENTER
            elif idx in (7, 9, 11):
                c.alignment = ALIGN_RIGHT
                c.number_format = "R$ #,##0.00"
            elif idx == 12:
                c.alignment = ALIGN_CENTER
                if "100% Conforme" in status_l:
                    c.fill = FILL_REG
                    c.font = FONT_REG
                elif "Atenção" in status_l:
                    c.fill = FILL_ALERT
                    c.font = FONT_ALERT
            else:
                c.alignment = ALIGN_LEFT
        cur_r += 1

    # Linha Total Geral
    tot_vals = [
        "TOTAL GERAL", f"{len(stats_loja)} Lojas c/ Movimento", "-", "-", "-",
        kpis["qtd_domingos"], kpis["valor_domingos"],
        kpis["qtd_feriados"], kpis["valor_feriados"],
        kpis["total_ocorrencias"], kpis["valor_planilha"],
        f"{kpis['alertas']} Alertas" if kpis["alertas"] > 0 else "100% Conforme"
    ]
    ws.row_dimensions[cur_r].height = 24
    for idx, val in enumerate(tot_vals, start=1):
        c = ws.cell(row=cur_r, column=idx, value=val)
        c.fill = FILL_TOTAL
        c.font = Font(name="Calibri", size=10, bold=True, color=C_AZUL_MARINHO)
        c.border = BORDER_TOTAL
        if idx in (1, 3, 4, 5, 6, 8, 10, 12):
            c.alignment = ALIGN_CENTER
        elif idx in (7, 9, 11):
            c.alignment = ALIGN_RIGHT
            c.number_format = "R$ #,##0.00"
        else:
            c.alignment = ALIGN_LEFT

    larguras = {
        "A": 14, "B": 24, "C": 12, "D": 18, "E": 16, "F": 14,
        "G": 16, "H": 14, "I": 16, "J": 14, "K": 18, "L": 24
    }
    for col_l, w in larguras.items():
        ws.column_dimensions[col_l].width = w

def gravar_aba_analitica(ws, linhas, titulo_banner):
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A3"

    ws.merge_cells("A1:W1")
    ct = ws["A1"]
    ct.value = titulo_banner
    ct.fill = FILL_HEADER
    ct.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    ct.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 30

    headers = [
        "Empresa", "Loja", "Município", "Cód. Sind.", "Entidade Sindical",
        "Matrícula", "Nome", "Cargo", "Situação Cadastral", "Data",
        "Dia", "Tipo de Dia / Feriado", "Marcações", "Horas Trab. (Sistema)", "Horas Calc. (Batidas)",
        "Enquadramento CCT", "Porte Check-outs", "Valor Planilha (R$)", "Valor Auditado (R$)", "Diferença (R$)",
        "Status da Auditoria", "Severidade", "Parecer / Ação Recomendada RH"
    ]
    ws.append(headers)
    ws.row_dimensions[2].height = 25
    for col in range(1, 24):
        c = ws.cell(row=2, column=col)
        c.fill = FILL_HEADER_SEC
        c.font = FONT_CABECALHO
        c.alignment = ALIGN_CENTER
        c.border = BORDER_THIN

    for r_idx, row_data in enumerate(linhas, start=3):
        ws.append(row_data)
        ws.row_dimensions[r_idx].height = 20
        fill_z = FILL_ZEBRA if (r_idx % 2 == 0) else FILL_BRANCO
        sev = row_data[21]

        for col in range(1, 24):
            c = ws.cell(row=r_idx, column=col)
            c.border = BORDER_THIN
            c.fill = fill_z
            c.font = FONT_DADOS

            if col in (1, 2, 4, 6, 10, 11, 14, 15, 17):
                c.alignment = ALIGN_CENTER
                if col == 6:
                    c.font = FONT_DADOS_BOLD
            elif col in (18, 19, 20):
                c.alignment = ALIGN_RIGHT
                c.number_format = "R$ #,##0.00"
                if col in (18, 19):
                    c.font = FONT_DADOS_BOLD
            elif col in (21, 22):
                c.alignment = ALIGN_CENTER
                if sev == "Crítico":
                    c.fill = FILL_CRIT
                    c.font = FONT_CRIT
                elif sev == "Alerta":
                    c.fill = FILL_ALERT
                    c.font = FONT_ALERT
                else:
                    c.fill = FILL_REG
                    c.font = FONT_REG
            else:
                c.alignment = ALIGN_LEFT

    ws.auto_filter.ref = f"A2:W{max(2, len(linhas) + 2)}"

    larguras = {
        "A": 10, "B": 10, "C": 22, "D": 11, "E": 36, "F": 12, "G": 32, "H": 28,
        "I": 20, "J": 13, "K": 8,  "L": 28, "M": 26, "N": 18, "O": 18, "P": 32,
        "Q": 16, "R": 18, "S": 18, "T": 15, "U": 32, "V": 14, "W": 65
    }
    for col_l, w in larguras.items():
        ws.column_dimensions[col_l].width = w

def gravar_aba_consolidado(ws, consolidado_colab):
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A3"

    ws.merge_cells("A1:N1")
    ct = ws["A1"]
    ct.value = "CENCOSUD BRASIL — CONSOLIDADO POR COLABORADOR PARA IMPORTAÇÃO NA FOLHA (EVENTOS 227 E 5751)"
    ct.fill = FILL_HEADER
    ct.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    ct.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 30

    headers = [
        "Empresa", "Loja", "Município", "Cód. Sind.", "Matrícula",
        "Colaborador (A-Z)", "Cargo", "Qtd. Domingos", "Qtd. Feriados", "Total Dias",
        "Soma Valor Planilha (R$)", "Soma Valor Auditado (R$)", "Diferença (R$)", "Status Consolidado"
    ]
    ws.append(headers)
    ws.row_dimensions[2].height = 25
    for col in range(1, 15):
        c = ws.cell(row=2, column=col)
        c.fill = FILL_HEADER_SEC
        c.font = FONT_CABECALHO
        c.alignment = ALIGN_CENTER
        c.border = BORDER_THIN

    # Ordenar colaboradores de A a Z pelo Nome
    lista_cons = sorted(consolidado_colab.values(), key=lambda x: x["nome"])

    tot_dom = 0
    tot_fer = 0
    tot_dias = 0
    tot_plan = 0.0
    tot_aud = 0.0

    for r_idx, cc in enumerate(lista_cons, start=3):
        diff = round(cc["val_auditado"] - cc["val_planilha"], 2)
        status_c = "Revisar Apuração Ponto" if cc["tem_alerta"] else "Consolidado OK"
        row_vals = [
            cc["empresa"],
            cc["loja"],
            cc["municipio"],
            cc["sindicato"],
            cc["matricula"],
            cc["nome"],
            cc["cargo"],
            cc["qtd_dom"],
            cc["qtd_fer"],
            cc["total_dias"],
            cc["val_planilha"],
            cc["val_auditado"],
            diff,
            status_c
        ]
        ws.append(row_vals)
        ws.row_dimensions[r_idx].height = 20
        fill_z = FILL_ZEBRA if (r_idx % 2 == 0) else FILL_BRANCO

        tot_dom += cc["qtd_dom"]
        tot_fer += cc["qtd_fer"]
        tot_dias += cc["total_dias"]
        tot_plan += cc["val_planilha"]
        tot_aud += cc["val_auditado"]

        for col in range(1, 15):
            c = ws.cell(row=r_idx, column=col)
            c.border = BORDER_THIN
            c.fill = fill_z
            c.font = FONT_DADOS
            if col in (1, 2, 4, 5, 8, 9, 10):
                c.alignment = ALIGN_CENTER
                if col == 5:
                    c.font = FONT_DADOS_BOLD
            elif col in (11, 12, 13):
                c.alignment = ALIGN_RIGHT
                c.number_format = "R$ #,##0.00"
                if col in (11, 12):
                    c.font = FONT_DADOS_BOLD
            elif col == 14:
                c.alignment = ALIGN_CENTER
                if cc["tem_alerta"]:
                    c.fill = FILL_ALERT
                    c.font = FONT_ALERT
                else:
                    c.fill = FILL_REG
                    c.font = FONT_REG
            else:
                c.alignment = ALIGN_LEFT

    # Linha de Total
    r_tot = len(lista_cons) + 3
    ws.row_dimensions[r_tot].height = 24
    tot_row = [
        "TOTAL", "-", f"{len(lista_cons)} Colaboradores", "-", "-",
        "TOTAL GERAL PARA FOLHA (EVT 227 / 5751)", "-",
        tot_dom, tot_fer, tot_dias,
        tot_plan, tot_aud, round(tot_aud - tot_plan, 2), "-"
    ]
    ws.append(tot_row)
    for col in range(1, 15):
        c = ws.cell(row=r_tot, column=col)
        c.fill = FILL_TOTAL
        c.font = Font(name="Calibri", size=10, bold=True, color=C_AZUL_MARINHO)
        c.border = BORDER_TOTAL
        if col in (1, 2, 4, 5, 8, 9, 10, 14):
            c.alignment = ALIGN_CENTER
        elif col in (11, 12, 13):
            c.alignment = ALIGN_RIGHT
            c.number_format = "R$ #,##0.00"
        else:
            c.alignment = ALIGN_LEFT

    ws.auto_filter.ref = f"A2:N{r_tot - 1}"

    larguras = {
        "A": 10, "B": 10, "C": 22, "D": 11, "E": 14, "F": 34, "G": 28,
        "H": 14, "I": 14, "J": 12, "K": 22, "L": 22, "M": 16, "N": 24
    }
    for col_l, w in larguras.items():
        ws.column_dimensions[col_l].width = w

# ======================================================================================
# 7. EXECUÇÃO PRINCIPAL
# ======================================================================================

def main():
    t0 = time.time()
    print("=" * 85)
    print("CENCOSUD BRASIL - MÓDULO 7: AUDITORIA DE GRATIFICAÇÃO MANUAL (GB & MERCANTIL)")
    print("=" * 85)

    lojas_info, feriados_desc_map, cola_map, grat_manual_map, grat_manual_list = carregar_bases()
    linhas_geral, linhas_divergencias, consolidado_colab, stats_loja, lojas_no_ht, kpis = processar_auditoria(
        lojas_info, feriados_desc_map, cola_map, grat_manual_map, grat_manual_list
    )

    print("\n-> Construindo Workbook Excel Oficial de 5 Abas...")
    wb = openpyxl.Workbook()
    default_ws = wb.active

    ws_ori = wb.create_sheet(title="Orientações")
    construir_aba_orientacoes(ws_ori)

    ws_res = wb.create_sheet(title="Resumo da Auditoria")
    construir_aba_resumo(ws_res, kpis, stats_loja, lojas_info, lojas_no_ht)

    ws_conf = wb.create_sheet(title="Conferência Geral")
    gravar_aba_analitica(
        ws_conf,
        linhas_geral,
        f"CENCOSUD BRASIL — BASE ANALÍTICA DE GRATIFICAÇÃO MANUAL DOMINGOS E FERIADOS ({len(linhas_geral)} OCORRÊNCIAS)"
    )

    ws_div = wb.create_sheet(title="Divergências")
    gravar_aba_analitica(
        ws_div,
        linhas_divergencias,
        f"CENCOSUD BRASIL — PAINEL DE ALERTAS E DIVERGÊNCIAS DE APURAÇÃO ({len(linhas_divergencias)} CASOS)"
    )

    ws_cons = wb.create_sheet(title="Consolidado Folha")
    gravar_aba_consolidado(ws_cons, consolidado_colab)

    wb.remove(default_ws)

    print(f"-> Salvando relatório em:\n   {OUTPUT_FILE}")
    wb.save(OUTPUT_FILE)

    print("=" * 85)
    print(f"CONCLUÍDO COM SUCESSO EM {time.time() - t0:.2f} SEGUNDOS!")
    print(f" - Ocorrências Auditadas:   {kpis['total_ocorrencias']} ({kpis['total_colaboradores']} colaboradores)")
    print(f" - Valor Total Planilha:    R$ {kpis['valor_planilha']:,.2f}")
    print(f" - Valor Total Auditado:    R$ {kpis['valor_auditado']:,.2f}")
    print(f" - Conformes / Regulares:   {kpis['conformes']} ({kpis['conformes']/kpis['total_ocorrencias']:.1%})")
    print(f" - Alertas de Apuração:     {kpis['alertas']}")
    print("=" * 85)

if __name__ == "__main__":
    main()
