"""
=====================================================================================
CENCOSUD BRASIL - ENGENHARIA DE DADOS & AUDITORIA DE RH
SUPER MÓDULO DE JORNADA CRÍTICA: NOTURNO & HORAS EXTRAS (PRIORIDADES 3 & 7)
=====================================================================================
Esteira de conciliação e auditoria matemática de limites legais de jornada, horas extras
e adicional noturno para a competência de Setembro/2026.

Auditorias Implementadas:
  1. Aprendiz com Hora Extra (Art. 432 CLT) - Infração Legal / Crítico
  2. Hora Extra Abusiva (> 4h diárias) - Risco Passivo / Crítico
  3. Hora Extra Acima do Limite CLT (> 2h diárias - Art. 59 CLT) - Crítico
  4. Marcação Incompleta (Batida Ímpar - Portaria 671) - Alerta
  5. Marcação Invertida / Descendo de Dia (Virada Noturna) - Alerta
  6. Adicional Noturno Elevado (> 7h diárias - Súmula 60 TST) - Alerta
  7. Jornada Conforme / Regular

Padrão Cencosud de 4 Abas:
  - Orientações
  - Resumo da Auditoria
  - Conferência Geral
  - Divergências
=====================================================================================
"""

import os
import glob
import time
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ======================================================================================
# 1. CONSTANTES E CONFIGURAÇÕES DE DIRETÓRIOS
# ======================================================================================

BASE_DIR = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\ponto"
INPUT_DIR = os.path.join(BASE_DIR, "Marcações Noturnas")
OUTPUT_DIR = os.path.join(BASE_DIR, "conferencias geradas pelo script")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Relatorio_Conferencia_Jornada_Critica_09.2026.xlsx")

DIR_CARTAO = os.path.join(BASE_DIR, "Justificativas de Ponto", "Cartão de Ponto")
DIR_OCORRENCIAS = os.path.join(BASE_DIR, "Justificativas de Ponto", "Ocorrências")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# Paleta Corporativa Cencosud
C_AZUL_MARINHO = "1B365D"
C_AZUL_MEDIO   = "2C4A7A"
C_AZUL_CLARO   = "EBF3FA"
C_CINZA_CAB    = "F1F5F9"
C_CINZA_ZEBRA  = "F8FAFC"
C_BRANCO       = "FFFFFF"
C_BORDA        = "CBD5E1"

# Badges de Severidade
C_CRIT_BG      = "FFC7CE"
C_CRIT_TXT     = "9C0006"
C_ALERT_BG     = "FFEB9C"
C_ALERT_TXT    = "9C6500"
C_REG_BG       = "C6EFCE"
C_REG_TXT      = "006100"

FONT_TITULO_TABELA = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_CABECALHO = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_DADOS = Font(name="Calibri", size=10, color="000000")
FONT_DADOS_BOLD = Font(name="Calibri", size=10, bold=True, color="000000")

FILL_HEADER = PatternFill(start_color=C_AZUL_MARINHO, end_color=C_AZUL_MARINHO, fill_type="solid")
FILL_HEADER_SEC = PatternFill(start_color=C_AZUL_MEDIO, end_color=C_AZUL_MEDIO, fill_type="solid")
FILL_ZEBRA = PatternFill(start_color=C_CINZA_ZEBRA, end_color=C_CINZA_ZEBRA, fill_type="solid")
FILL_BRANCO = PatternFill(start_color=C_BRANCO, end_color=C_BRANCO, fill_type="solid")

FILL_CRIT = PatternFill(start_color=C_CRIT_BG, end_color=C_CRIT_BG, fill_type="solid")
FONT_CRIT = Font(name="Calibri", size=10, bold=True, color=C_CRIT_TXT)

FILL_ALERT = PatternFill(start_color=C_ALERT_BG, end_color=C_ALERT_BG, fill_type="solid")
FONT_ALERT = Font(name="Calibri", size=10, bold=True, color=C_ALERT_TXT)

FILL_REG = PatternFill(start_color=C_REG_BG, end_color=C_REG_BG, fill_type="solid")
FONT_REG = Font(name="Calibri", size=10, bold=True, color=C_REG_TXT)

BORDER_THIN = Border(
    left=Side(style='thin', color=C_BORDA),
    right=Side(style='thin', color=C_BORDA),
    top=Side(style='thin', color=C_BORDA),
    bottom=Side(style='thin', color=C_BORDA)
)

ALIGN_LEFT = Alignment(horizontal='left', vertical='center')
ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
ALIGN_RIGHT = Alignment(horizontal='right', vertical='center')

# ======================================================================================
# 2. FUNÇÕES AUXILIARES DE PARSING, NORMALIZAÇÃO E CLASSIFICAÇÃO
# ======================================================================================

def parse_minutes(val):
    """Converte horários em formato 'HH:MM' ou decimal para minutos inteiros."""
    if val is None:
        return 0
    s = str(val).strip()
    if not s or s in ['', '0', '0.0', '00:00', 'None', 'nan']:
        return 0
    try:
        if ':' in s:
            parts = s.split(':')
            h = int(parts[0])
            m = int(parts[1])
            return h * 60 + m if h >= 0 else h * 60 - m
        return int(float(s) * 60)
    except:
        return 0

def check_batidas(marc_str):
    """
    Analisa marcações diárias: detecta batidas ímpares e batidas invertidas.
    - Jornada Noturna Regular (Conforme): inicia à tarde/noite (>= 12:00), cruza a
      meia-noite uma única vez e encerra pela manhã (<= 13:00).
    - Marcação Invertida:
      1) Inicia na madrugada (<= 07:30) e possui entrada noturna (>= 18:00) no mesmo dia
         com hiato diurno >= 8h entre batidas (ex: '00:00 01:00 06:00 21:45');
      2) Possui queda de horário que não configura jornada noturna regular.
    """
    if not marc_str:
        return False, False, []
    batidas = [b.strip() for b in str(marc_str).strip().split() if ':' in b]
    qtd = len(batidas)
    is_impar = (qtd % 2 != 0) and (qtd > 0)
    is_invertida = False
    if qtd >= 2:
        try:
            mins = [parse_minutes(b) for b in batidas]
            wraps = sum(1 for i in range(qtd - 1) if mins[i+1] < mins[i])
            max_gap = max(mins[i+1] - mins[i] for i in range(qtd - 1))
            
            # Caso 1: Batidas da madrugada/manhã seguidas de entrada noturna no mesmo dia (virada não associada)
            if mins[0] <= 450 and max(mins) >= 1080 and max_gap >= 480:
                is_invertida = True
            # Caso 2: Queda de horário fora do padrão de turno noturno regular
            elif wraps > 0:
                is_regular_night = (wraps == 1 and mins[0] >= 720 and mins[-1] <= 780)
                if not is_regular_night:
                    is_invertida = True
        except:
            pass
    return is_impar, is_invertida, batidas

def map_regional(estab, dept=''):
    """Mapeia a regional/bandeira a partir do Estabelecimento e Departamento."""
    text = (str(estab) + ' ' + str(dept)).upper()
    if 'BRETAS' in text or 'POSTO 5' in text or 'POSTO 6' in text or 'CD 552' in text:
        return 'Bretas'
    if 'GIGA' in text:
        return 'Giga'
    if 'MERCANTIL' in text:
        return 'Mercantil'
    if 'PERINI' in text:
        return 'Perini'
    if 'PREZUNIC' in text or 'CD 703' in text:
        return 'Prezunic'
    if 'SPID' in text:
        return 'Spid'
    if 'GBARBOSA' in text or 'CD 289' in text or 'GB' in text or 'ELETROSHOW' in text or 'FARMACIA' in text or 'CD 099' in text:
        return 'GBarbosa'
    return 'Corporativo / Outros'

def format_data(dt_val):
    """Padroniza data para string DD/MM/AAAA."""
    if dt_val is None:
        return ""
    if isinstance(dt_val, datetime):
        return dt_val.strftime("%d/%m/%Y")
    s = str(dt_val).strip()
    if len(s) >= 10 and s[4] == '-' and s[7] == '-':
        return f"{s[8:10]}/{s[5:7]}/{s[0:4]}"
    return s[:10]

def auditar_jornada(cargo, he_str, an_str, marc_str):
    """
    Executa a auditoria de limites de jornada seguindo a escada de precedência:
    Aprendiz com HE > HE Abusiva (> 4h) > HE CLT (> 2h) > Marcação Ímpar > Batida Invertida > Adicional Noturno Elevado.
    """
    he_mins = parse_minutes(he_str)
    an_mins = parse_minutes(an_str)
    is_impar, is_invertida, _ = check_batidas(marc_str)
    
    c_upper = str(cargo).upper()
    is_aprendiz = 'APRENDIZ' in c_upper
    
    if is_aprendiz and he_mins > 0:
        return (
            'Aprendiz com Hora Extra (Ilegal Art. 432 CLT)',
            'Crítico',
            'Ajustar imediatamente escala do jovem aprendiz. A legislação trabalhista veda qualquer prorrogação de jornada para aprendizes.'
        )
    elif he_mins > 240:
        return (
            'Hora Extra Abusiva (> 4h diárias)',
            'Crítico',
            'Auditar justificativa da sobrejornada extrema com o gestor da loja. Risco grave de passivo trabalhista.'
        )
    elif he_mins > 120:
        return (
            'Hora Extra Acima do Limite CLT (> 2h diárias)',
            'Crítico',
            'Notificar liderança de loja para readequar escala. O Art. 59 da CLT limita o trabalho extraordinário a no máximo 2 horas por dia.'
        )
    elif is_impar:
        return (
            'Marcação Incompleta (Batida Ímpar)',
            'Alerta',
            'Cobrar gestor de loja para tratar a marcação em aberto e regularizar o espelho de ponto.'
        )
    elif is_invertida:
        return (
            'Marcação Invertida / Descendo de Dia',
            'Alerta',
            'Verificar se a virada noturna está devidamente configurada no horário do colaborador para evitar inconsistência de apuração.'
        )
    elif an_mins > 420:
        return (
            'Adicional Noturno Elevado (> 7h)',
            'Alerta',
            'Validar se houve prorrogação de jornada noturna (Súmula 60 TST) ou erro de cálculo sistêmico.'
        )
    else:
        return (
            'Conforme',
            'Regular',
            'Conforme'
        )

# ======================================================================================
# 3. MAPEAMENTO AUXILIAR DE CARGOS A PARTIR DAS BASES DE APOIO
# ======================================================================================

def carregar_mapa_cargos():
    """Carrega mapeamento de matrícula -> cargo a partir da base cadastral CSV e bases de apoio."""
    import csv
    print("-> Carregando base cadastral auxiliar de Cargos...")
    cargo_map = {}
    
    # 1. Base Cadastral Geral (CSV Gestor)
    csv_cola = os.path.join(BASE_DIR, "COLA_Colaboradores_CSV_Gestor - 17092026_112137_057280.csv")
    if os.path.exists(csv_cola):
        try:
            with open(csv_cola, mode='r', encoding='latin1') as f_csv:
                next(f_csv, None) # Pular linha de título
                reader = csv.DictReader(f_csv, delimiter=';')
                for row in reader:
                    m = str(row.get('Matrícula', '')).strip().split('.')[0]
                    c = str(row.get('Cargo', '')).strip()
                    if m and c and c not in ['', 'Cargo', 'None']:
                        cargo_map[m] = c
        except Exception as e:
            print(f"   Aviso ao ler CSV Colaboradores: {e}")

    # 2. Arquivos de Ocorrências e Cartão Bretas (complemento)
    arquivos_apoio = glob.glob(os.path.join(DIR_OCORRENCIAS, "*.xlsx")) + glob.glob(os.path.join(DIR_CARTAO, "*Bretas*.xlsx"))
    for f in arquivos_apoio:
        try:
            wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
            ws = wb.active
            header = [str(c).strip() if c is not None else '' for c in next(ws.iter_rows(max_row=1, values_only=True))]
            if 'Matricula' in header and 'Cargo' in header:
                idx_m = header.index('Matricula')
                idx_c = header.index('Cargo')
                for r in ws.iter_rows(values_only=True):
                    if r[idx_m] and r[idx_c]:
                        m = str(r[idx_m]).strip().split(' - ')[0].split('.')[0]
                        c = str(r[idx_c]).strip()
                        if m and c and c not in ['', 'Cargo', 'None']:
                            cargo_map[m] = c
            wb.close()
        except Exception as e:
            print(f"   Aviso ao ler {os.path.basename(f)}: {e}")
                
    print(f"   Mapeamento de cargos concluído: {len(cargo_map):,} matrículas indexadas.")
    return cargo_map

# ======================================================================================
# 4. CARREGAMENTO E AUDITORIA DAS JORNADAS COM DEDUPLICAÇÃO
# ======================================================================================

def processar_arquivos_jornada(cargo_map):
    """
    Lê os 9 arquivos de Horas Trabalhadas em streaming, aplica deduplicação 100% matemática
    e executa a auditoria de regras de jornada.
    """
    print("\n-> Processando arquivos de Marcações Noturnas / Horas Trabalhadas...")
    arquivos = sorted(glob.glob(os.path.join(INPUT_DIR, "*.xlsx")))
    print(f"   Total de arquivos encontrados: {len(arquivos)}")
    
    seen_keys = set()
    linhas_geral = []
    linhas_divergencias = []
    
    kpis = {
        'total': 0,
        'conforme': 0,
        'aprendiz_he': 0,
        'he_abusiva': 0,
        'he_clt': 0,
        'impar': 0,
        'invertida': 0,
        'noturno_elevado': 0
    }
    
    regionais_stats = {}
    tipos_stats = {
        'Aprendiz com Hora Extra (Ilegal Art. 432 CLT)': {'sev': 'Crítico', 'qtd': 0, 'acao': 'Ajustar imediatamente escala do jovem aprendiz. A legislação trabalhista veda qualquer prorrogação de jornada para aprendizes.'},
        'Hora Extra Abusiva (> 4h diárias)': {'sev': 'Crítico', 'qtd': 0, 'acao': 'Auditar justificativa da sobrejornada extrema com o gestor da loja. Risco grave de passivo trabalhista.'},
        'Hora Extra Acima do Limite CLT (> 2h diárias)': {'sev': 'Crítico', 'qtd': 0, 'acao': 'Notificar liderança de loja para readequar escala. O Art. 59 da CLT limita o trabalho extraordinário a no máximo 2 horas por dia.'},
        'Marcação Incompleta (Batida Ímpar)': {'sev': 'Alerta', 'qtd': 0, 'acao': 'Cobrar gestor de loja para tratar a marcação em aberto e regularizar o espelho de ponto.'},
        'Marcação Invertida / Descendo de Dia': {'sev': 'Alerta', 'qtd': 0, 'acao': 'Verificar se a virada noturna está devidamente configurada no horário do colaborador para evitar inconsistência de apuração.'},
        'Adicional Noturno Elevado (> 7h)': {'sev': 'Alerta', 'qtd': 0, 'acao': 'Validar se houve prorrogação de jornada noturna (Súmula 60 TST) ou erro de cálculo sistêmico.'},
        'Conforme': {'sev': 'Regular', 'qtd': 0, 'acao': 'Conforme'}
    }
    
    total_linhas_lidas = 0
    total_duplicadas = 0
    
    t_start = time.time()
    
    for f in arquivos:
        fname = os.path.basename(f)
        t_f0 = time.time()
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        header = [str(c).strip() if c is not None else '' for c in next(rows)]
        
        # Mapeamento dinâmico de colunas
        idx_m = header.index('Matricula') if 'Matricula' in header else header.index('Colaborador')
        idx_col1 = header.index('Coluna1') if 'Coluna1' in header else -1
        idx_nome = header.index('Nome') if 'Nome' in header else -1
        idx_cargo = header.index('Cargo') if 'Cargo' in header else -1
        idx_estab = header.index('Estabelecimento')
        idx_dept = header.index('Departamento') if 'Departamento' in header else -1
        idx_cc = header.index('CentroDeCustos') if 'CentroDeCustos' in header else -1
        idx_data = header.index('Data')
        idx_marc = header.index('Marcacoes')
        idx_ht = header.index('HorasTrabalhadas')
        idx_he = header.index('HorasExtras')
        idx_an = header.index('AdicionalNoturno')
        
        linhas_arquivo = 0
        
        for r in rows:
            total_linhas_lidas += 1
            raw_m = r[idx_m]
            raw_d = r[idx_data]
            raw_marc = r[idx_marc]
            
            if raw_m is None or str(raw_m).strip() in ['', 'Matricula', 'Colaborador', 'None']:
                continue
                
            # Tratamento da Matrícula e Nome
            raw_m_str = str(raw_m).strip()
            if ' - ' in raw_m_str:
                parts = raw_m_str.split(' - ', 1)
                m_str = parts[0].strip().split('.')[0]
                nome = parts[1].strip()
            else:
                m_str = raw_m_str.split('.')[0]
                if idx_col1 != -1 and r[idx_col1]:
                    nome = str(r[idx_col1]).strip()
                elif idx_nome != -1 and r[idx_nome]:
                    nome = str(r[idx_nome]).strip()
                else:
                    nome = "COLABORADOR"
                    
            try:
                mat_int = int(m_str)
            except:
                continue
                
            d_str = str(raw_d)[:10] if raw_d else ''
            marc_str = str(raw_marc).strip() if raw_marc else ''
            
            # Deduplicação estrita via chave única (Matrícula + Data + Marcações)
            chave = (mat_int, d_str, marc_str)
            if chave in seen_keys:
                total_duplicadas += 1
                continue
            seen_keys.add(chave)
            
            # Cargo
            if idx_cargo != -1 and r[idx_cargo] and str(r[idx_cargo]).strip() not in ['', 'Cargo', 'None']:
                cargo = str(r[idx_cargo]).strip()
                cargo_map[str(mat_int)] = cargo
            else:
                cargo = cargo_map.get(str(mat_int), "NÃO INFORMADO")
                
            estab = str(r[idx_estab]).strip() if r[idx_estab] else ""
            if estab == "Estabelecimento":
                continue
            dept = str(r[idx_dept]).strip() if idx_dept != -1 and r[idx_dept] else ""
            cc = str(r[idx_cc]).strip() if idx_cc != -1 and r[idx_cc] else ""
            
            reg = map_regional(estab, dept)
            
            ht_str = str(r[idx_ht]).strip() if r[idx_ht] is not None else "00:00"
            he_str = str(r[idx_he]).strip() if r[idx_he] is not None else "00:00"
            an_str = str(r[idx_an]).strip() if r[idx_an] is not None else "00:00"
            
            tipo_div, severidade, acao = auditar_jornada(cargo, he_str, an_str, marc_str)
            
            # Atualização de KPIs e Estatísticas
            kpis['total'] += 1
            tipos_stats[tipo_div]['qtd'] += 1
            
            if tipo_div == 'Aprendiz com Hora Extra (Ilegal Art. 432 CLT)':
                kpis['aprendiz_he'] += 1
            elif tipo_div == 'Hora Extra Abusiva (> 4h diárias)':
                kpis['he_abusiva'] += 1
            elif tipo_div == 'Hora Extra Acima do Limite CLT (> 2h diárias)':
                kpis['he_clt'] += 1
            elif tipo_div == 'Marcação Incompleta (Batida Ímpar)':
                kpis['impar'] += 1
            elif tipo_div == 'Marcação Invertida / Descendo de Dia':
                kpis['invertida'] += 1
            elif tipo_div == 'Adicional Noturno Elevado (> 7h)':
                kpis['noturno_elevado'] += 1
            else:
                kpis['conforme'] += 1
                
            if reg not in regionais_stats:
                regionais_stats[reg] = {
                    'total': 0, 'conforme': 0, 'critico': 0, 'alerta': 0,
                    'aprendiz_he': 0, 'he_abusiva': 0, 'he_clt': 0,
                    'impar': 0, 'invertida': 0, 'noturno_elevado': 0
                }
            rs = regionais_stats[reg]
            rs['total'] += 1
            if severidade == 'Regular':
                rs['conforme'] += 1
            elif severidade == 'Crítico':
                rs['critico'] += 1
            else:
                rs['alerta'] += 1
                
            if tipo_div == 'Aprendiz com Hora Extra (Ilegal Art. 432 CLT)':
                rs['aprendiz_he'] += 1
            elif tipo_div == 'Hora Extra Abusiva (> 4h diárias)':
                rs['he_abusiva'] += 1
            elif tipo_div == 'Hora Extra Acima do Limite CLT (> 2h diárias)':
                rs['he_clt'] += 1
            elif tipo_div == 'Marcação Incompleta (Batida Ímpar)':
                rs['impar'] += 1
            elif tipo_div == 'Marcação Invertida / Descendo de Dia':
                rs['invertida'] += 1
            elif tipo_div == 'Adicional Noturno Elevado (> 7h)':
                rs['noturno_elevado'] += 1
                
            # Formatação de Data
            dt_fmt = format_data(raw_d)
            
            # Linha analítica de 12 colunas (excluídas colunas B, C e D: Estabelecimento, Departamento, Centro de Custo)
            linha = (
                reg, mat_int, nome, cargo, dt_fmt,
                marc_str, ht_str, he_str, an_str, tipo_div, severidade, acao
            )
            
            linhas_geral.append(linha)
            if severidade in ('Crítico', 'Alerta'):
                linhas_divergencias.append(linha)
                
            linhas_arquivo += 1
            
        wb.close()
        t_f1 = time.time()
        print(f"   [{fname}] Lido em {t_f1 - t_f0:.2f}s | {linhas_arquivo:,} registros válidos.")
        
    t_end = time.time()
    print(f"\n-> Processamento concluído em {t_end - t_start:.2f}s!")
    print(f"   Total de Linhas Lidas: {total_linhas_lidas:,}")
    print(f"   Duplicatas Eliminadas: {total_duplicadas:,}")
    print(f"   Jornadas Únicas Auditadas: {len(linhas_geral):,}")
    print(f"   Divergências + Alertas: {len(linhas_divergencias):,}")
    
    return linhas_geral, linhas_divergencias, kpis, regionais_stats, tipos_stats

# ======================================================================================
# 5. CONSTRUÇÃO DO WORKBOOK OFICIAL NO PADRÃO CENCOSUD DE 4 ABAS
# ======================================================================================

def construir_aba_orientacoes(ws):
    """Constrói a Aba 1: Orientações Institucionais e Embasamento Legal."""
    ws.views.sheetView[0].showGridLines = True
    
    # Bloco de Título
    ws.merge_cells("A1:K1")
    cell_t = ws["A1"]
    cell_t.value = "CENCOSUD BRASIL — AUDITORIA DE JORNADA CRÍTICA, NOTURNO & HORAS EXTRAS"
    cell_t.fill = FILL_HEADER
    cell_t.font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    cell_t.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 40
    
    # Subtítulo institucional
    ws.merge_cells("A2:K2")
    cell_sub = ws["A2"]
    cell_sub.value = "Área de Benefícios e Ponto — Módulo Automatizado de Conformidade Trabalhista e Gestão de Riscos (Competência: 09/2026)"
    cell_sub.fill = FILL_HEADER_SEC
    cell_sub.font = Font(name="Calibri", size=11, italic=True, color="FFFFFF")
    cell_sub.alignment = ALIGN_CENTER
    ws.row_dimensions[2].height = 24
    
    # Conteúdo de Orientações em Tabela Estruturada
    secoes = [
        ("1. OBJETIVO DO SUPER MÓDULO",
         "Consolidar e auditar 100% das jornadas trabalhadas em todas as 7 regionais da Cencosud Brasil (Bretas, GBarbosa, Giga, Mercantil, Perini, Prezunic e Spid), "
         "confrontando os registros eletrônicos de ponto contra os limites da legislação trabalhista brasileira (CLT), súmulas dos tribunais superiores e acordos coletivos."),
        
        ("2. EMBASAMENTO LEGAL CLT & JURISPRUDÊNCIA",
         "• Art. 432 da CLT: É vedada a prorrogação da jornada de trabalho do aprendiz. A prestação de horas extras é infração administrativa gravíssima e gera nulidade do contrato de aprendizagem.\n"
         "• Art. 59 da CLT: A duração diária do trabalho poderá ser acrescida de horas suplementares, em número não excedente de 2 (duas), mediante acordo individual, convenção coletiva ou acordo coletivo de trabalho.\n"
         "• Art. 73 da CLT & Súmula 60 TST: O trabalho noturno (22h às 05h) é remunerado com adicional de no mínimo 20% e tem hora ficta de 52 minutos e 30 segundos. Cumprida integralmente a jornada no período noturno e prorrogada esta, devido é também o adicional quanto às horas prorrogadas.\n"
         "• Art. 74 da CLT & Portaria MTP 671/2021: O registro de ponto deve refletir a real jornada praticada, sendo proibidas marcações ímpares, incompletas ou fraudulentas que descumpram a assiduidade e a fidelidade dos registros."),
        
        ("3. MATRIZ DE SEVERIDADE & CLASSIFICAÇÃO DE RISCO",
         "• CRÍTICO (Vermelho): Violações diretas de preceitos proibitivos da CLT e riscos de autuação do MTE ou passivo trabalhista iminente. Exige saneamento imediato e readequação de escalas pela liderança da loja.\n"
         "• ALERTA (Amarelo): Inconsistências operacionais de preenchimento, batidas ímpares pendentes de tratamento e viradas de turno que podem distorcer o cálculo de horas e gerar pagamentos incorretos.\n"
         "• REGULAR (Verde): Jornadas regulares que cumpriram fielmente as disposições legais e regulamentares."),
        
        ("4. DIRETRIZES DE AUDITORIA & METODOLOGIA",
         "• Deduplicação Matemática: Descarte rigoroso de registros repetidos oriundos de múltiplos downloads de relatórios regionais pela chave [Matrícula + Data + Marcações].\n"
         "• Enriquecimento Cadastral: Recuperação automática de cargos operacionais e técnicos em bases espelho quando o relatório de origem possui campos nulos.\n"
         "• Escada de Precedência: Aplicação da severidade máxima em jornadas com múltiplos apontamentos (Aprendiz com HE > HE Abusiva > HE CLT > Marcação Ímpar > Batida Invertida > Adicional Noturno).")
    ]
    
    cur_row = 4
    for titulo, texto in secoes:
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=11)
        c_tit = ws.cell(row=cur_row, column=1)
        c_tit.value = titulo
        c_tit.fill = PatternFill(start_color=C_AZUL_MEDIO, end_color=C_AZUL_MEDIO, fill_type="solid")
        c_tit.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c_tit.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[cur_row].height = 26
        cur_row += 1
        
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=11)
        c_txt = ws.cell(row=cur_row, column=1)
        c_txt.value = texto
        c_txt.fill = PatternFill(start_color=C_CINZA_ZEBRA, end_color=C_CINZA_ZEBRA, fill_type="solid")
        c_txt.font = Font(name="Calibri", size=10, color="000000")
        c_txt.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True, indent=1)
        
        # Estimar altura da linha
        linhas_texto = texto.count('\n') + len(texto) // 110 + 1
        ws.row_dimensions[cur_row].height = max(45, linhas_texto * 18)
        cur_row += 2

    # Ajuste de larguras para colunas A a K
    for col in range(1, 12):
        ws.column_dimensions[get_column_letter(col)].width = 14

def construir_aba_resumo(ws, kpis, regionais_stats, tipos_stats):
    """Constrói a Aba 2: Resumo da Auditoria com 7 Cards de KPIs e Tabelas Consolidadas."""
    ws.views.sheetView[0].showGridLines = True
    
    # Título do Dashboard
    ws.merge_cells("A1:L1")
    c_t = ws["A1"]
    c_t.value = "CENCOSUD BRASIL — DASHBOARD EXECUTIVO DE JORNADA CRÍTICA & NOTURNO"
    c_t.fill = FILL_HEADER
    c_t.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    c_t.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 36
    
    # 7 Cards Executivos de KPIs (Distribuídos em duas linhas elegantes)
    # Linha 1 de Cards: Total, Taxa Conformidade, Aprendiz com HE, HE Abusiva (> 4h)
    cards_l1 = [
        ("TOTAL JORNADAS AUDITADAS", f"{kpis['total']:,}", "Competência 09/2026", C_AZUL_MARINHO, 1, 3),
        ("TAXA DE CONFORMIDADE", f"{(kpis['conforme'] / kpis['total']):.1%}", "Jornadas Regulares", "166534", 4, 6),
        ("APRENDIZ COM HORA EXTRA", f"{kpis['aprendiz_he']:,}", "Infração Legal (Art. 432)", C_CRIT_TXT, 7, 9),
        ("HORA EXTRA ABUSIVA (> 4H)", f"{kpis['he_abusiva']:,}", "Sobrejornada Extrema", C_CRIT_TXT, 10, 12)
    ]
    
    # Linha 2 de Cards: HE Acima da CLT (> 2h), Marcações Incompletas, Marcações Invertidas
    cards_l2 = [
        ("HE ACIMA DA CLT (> 2H)", f"{kpis['he_clt']:,}", "Violação do Art. 59 CLT", C_CRIT_TXT, 1, 4),
        ("MARCAÇÕES INCOMPLETAS (ÍMPARES)", f"{kpis['impar']:,}", "Batidas em Aberto", C_ALERT_TXT, 5, 8),
        ("MARCAÇÕES INVERTIDAS / VIRADAS", f"{kpis['invertida']:,}", "Virada de Turno Não Ajustada", C_ALERT_TXT, 9, 12)
    ]
    
    def render_card(titulo, valor, subtitulo, cor_txt, col_ini, col_fim, row_base):
        # Linha Título
        ws.merge_cells(start_row=row_base, start_column=col_ini, end_row=row_base, end_column=col_fim)
        ct = ws.cell(row=row_base, column=col_ini)
        ct.value = titulo
        ct.fill = PatternFill(start_color=C_CINZA_CAB, end_color=C_CINZA_CAB, fill_type="solid")
        ct.font = Font(name="Calibri", size=9, bold=True, color="475569")
        ct.alignment = ALIGN_CENTER
        
        # Linha Valor
        ws.merge_cells(start_row=row_base+1, start_column=col_ini, end_row=row_base+1, end_column=col_fim)
        cv = ws.cell(row=row_base+1, column=col_ini)
        cv.value = valor
        cv.fill = FILL_BRANCO
        cv.font = Font(name="Calibri", size=18, bold=True, color=cor_txt)
        cv.alignment = ALIGN_CENTER
        
        # Linha Subtítulo
        ws.merge_cells(start_row=row_base+2, start_column=col_ini, end_row=row_base+2, end_column=col_fim)
        cs = ws.cell(row=row_base+2, column=col_ini)
        cs.value = subtitulo
        cs.fill = FILL_BRANCO
        cs.font = Font(name="Calibri", size=8, italic=True, color="64748B")
        cs.alignment = ALIGN_CENTER
        
        # Bordas
        for r in range(row_base, row_base+3):
            for c in range(col_ini, col_fim+1):
                ws.cell(row=r, column=c).border = BORDER_THIN
                
    # Renderizar Linha 1 de Cards (Linhas 3, 4, 5)
    for t, v, s, cor, ci, cf in cards_l1:
        render_card(t, v, s, cor, ci, cf, 3)
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 30
    ws.row_dimensions[5].height = 16
    
    # Renderizar Linha 2 de Cards (Linhas 7, 8, 9)
    for t, v, s, cor, ci, cf in cards_l2:
        render_card(t, v, s, cor, ci, cf, 7)
    ws.row_dimensions[7].height = 18
    ws.row_dimensions[8].height = 30
    ws.row_dimensions[9].height = 16
    
    # ----------------------------------------------------------------------------------
    # TABELA 1: CONSOLIDAÇÃO EXECUTIVA POR REGIONAL / BANDEIRA
    # ----------------------------------------------------------------------------------
    row_t1_tit = 11
    ws.merge_cells(f"A{row_t1_tit}:L{row_t1_tit}")
    ct1 = ws[f"A{row_t1_tit}"]
    ct1.value = "1. CONSOLIDAÇÃO DE JORNADAS POR REGIONAL / BANDEIRA"
    ct1.fill = FILL_HEADER_SEC
    ct1.font = FONT_TITULO_TABELA
    ct1.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[row_t1_tit].height = 24
    
    headers_t1 = [
        "Regional / Bandeira", "Total Jornadas", "Conformes", "Alertas Operac.", "Diverg. Críticas",
        "Taxa Conform. (%)", "Aprendiz c/ HE", "HE Abusiva (> 4h)", "HE Acima CLT (> 2h)",
        "Batida Ímpar", "Batida Invertida", "Noturno > 7h"
    ]
    
    row_t1_h = 12
    ws.row_dimensions[row_t1_h].height = 26
    for c_idx, h in enumerate(headers_t1, start=1):
        cell = ws.cell(row=row_t1_h, column=c_idx, value=h)
        cell.fill = FILL_HEADER
        cell.font = FONT_CABECALHO
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_THIN
        
    ordem_regionais = ["Bretas", "GBarbosa", "Giga", "Mercantil", "Perini", "Prezunic", "Spid", "Corporativo / Outros"]
    cur_r = 13
    
    tot_geral = {
        'total': 0, 'conforme': 0, 'alerta': 0, 'critico': 0,
        'aprendiz_he': 0, 'he_abusiva': 0, 'he_clt': 0,
        'impar': 0, 'invertida': 0, 'noturno_elevado': 0
    }
    
    for reg in ordem_regionais:
        if reg not in regionais_stats:
            continue
        st = regionais_stats[reg]
        for k in tot_geral:
            tot_geral[k] += st[k]
            
        taxa = (st['conforme'] / st['total']) if st['total'] > 0 else 0.0
        fill_zebra = FILL_ZEBRA if (cur_r % 2 == 0) else FILL_BRANCO
        
        vals = [
            reg, st['total'], st['conforme'], st['alerta'], st['critico'],
            taxa, st['aprendiz_he'], st['he_abusiva'], st['he_clt'],
            st['impar'], st['invertida'], st['noturno_elevado']
        ]
        
        ws.row_dimensions[cur_r].height = 20
        for c_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=cur_r, column=c_idx, value=val)
            cell.fill = fill_zebra
            cell.border = BORDER_THIN
            if c_idx == 1:
                cell.font = FONT_DADOS_BOLD
                cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            elif c_idx == 6:
                cell.font = FONT_DADOS_BOLD
                cell.alignment = ALIGN_RIGHT
                cell.number_format = '0.0%'
            else:
                cell.font = FONT_DADOS
                cell.alignment = ALIGN_RIGHT
                cell.number_format = '#,##0'
                if c_idx in (5, 7, 8, 9) and val > 0:
                    cell.font = Font(name="Calibri", size=10, bold=True, color=C_CRIT_TXT)
                elif c_idx in (4, 10, 11, 12) and val > 0:
                    cell.font = Font(name="Calibri", size=10, bold=True, color=C_ALERT_TXT)
        cur_r += 1
        
    # Linha Total Geral Tabela 1
    taxa_tot = (tot_geral['conforme'] / tot_geral['total']) if tot_geral['total'] > 0 else 0.0
    tot_vals = [
        "TOTAL GERAL (7 BANDEIRAS)", tot_geral['total'], tot_geral['conforme'], tot_geral['alerta'], tot_geral['critico'],
        taxa_tot, tot_geral['aprendiz_he'], tot_geral['he_abusiva'], tot_geral['he_clt'],
        tot_geral['impar'], tot_geral['invertida'], tot_geral['noturno_elevado']
    ]
    ws.row_dimensions[cur_r].height = 24
    for c_idx, val in enumerate(tot_vals, start=1):
        cell = ws.cell(row=cur_r, column=c_idx, value=val)
        cell.fill = PatternFill(start_color=C_CINZA_CAB, end_color=C_CINZA_CAB, fill_type="solid")
        cell.font = Font(name="Calibri", size=10, bold=True, color=C_AZUL_MARINHO)
        cell.border = Border(top=Side(style='thin', color=C_AZUL_MARINHO), bottom=Side(style='double', color=C_AZUL_MARINHO))
        if c_idx == 1:
            cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        elif c_idx == 6:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = '0.0%'
        else:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = '#,##0'
    cur_r += 3
    
    # ----------------------------------------------------------------------------------
    # TABELA 2: CONSOLIDAÇÃO POR TIPO DE DIVERGÊNCIA & AÇÃO RH
    # ----------------------------------------------------------------------------------
    ws.merge_cells(f"A{cur_r}:L{cur_r}")
    ct2 = ws[f"A{cur_r}"]
    ct2.value = "2. CONSOLIDAÇÃO POR TIPO DE DIVERGÊNCIA, SEVERIDADE E AÇÃO RECOMENDADA RH"
    ct2.fill = FILL_HEADER_SEC
    ct2.font = FONT_TITULO_TABELA
    ct2.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[cur_r].height = 24
    cur_r += 1
    
    headers_t2 = ["Tipo de Divergência", "Severidade", "Qtd. Ocorrências", "% do Total", "Ação Recomendada RH"]
    ws.row_dimensions[cur_r].height = 24
    
    # Mesclagem elegante das colunas para Tabela 2
    # A-E: Tipo, F: Severidade, G-H: Qtd, I: %, J-L: Ação
    ws.merge_cells(f"A{cur_r}:E{cur_r}")
    ws.cell(row=cur_r, column=1, value="Tipo de Divergência").alignment = ALIGN_CENTER
    ws.cell(row=cur_r, column=6, value="Severidade").alignment = ALIGN_CENTER
    ws.merge_cells(f"G{cur_r}:H{cur_r}")
    ws.cell(row=cur_r, column=7, value="Qtd. Jornadas").alignment = ALIGN_CENTER
    ws.cell(row=cur_r, column=9, value="% do Total").alignment = ALIGN_CENTER
    ws.merge_cells(f"J{cur_r}:L{cur_r}")
    ws.cell(row=cur_r, column=10, value="Ação Recomendada RH").alignment = ALIGN_CENTER
    
    for c in range(1, 13):
        ws.cell(row=cur_r, column=c).fill = FILL_HEADER
        ws.cell(row=cur_r, column=c).font = FONT_CABECALHO
        ws.cell(row=cur_r, column=c).border = BORDER_THIN
    cur_r += 1
    
    ordem_tipos = [
        'Aprendiz com Hora Extra (Ilegal Art. 432 CLT)',
        'Hora Extra Abusiva (> 4h diárias)',
        'Hora Extra Acima do Limite CLT (> 2h diárias)',
        'Marcação Incompleta (Batida Ímpar)',
        'Marcação Invertida / Descendo de Dia',
        'Adicional Noturno Elevado (> 7h)',
        'Conforme'
    ]
    
    for t_nome in ordem_tipos:
        info = tipos_stats[t_nome]
        qtd = info['qtd']
        pct = (qtd / kpis['total']) if kpis['total'] > 0 else 0.0
        sev = info['sev']
        acao = info['acao']
        
        fill_zebra = FILL_ZEBRA if (cur_r % 2 == 0) else FILL_BRANCO
        ws.row_dimensions[cur_r].height = 22
        
        ws.merge_cells(f"A{cur_r}:E{cur_r}")
        c_tipo = ws.cell(row=cur_r, column=1, value=t_nome)
        c_tipo.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        c_tipo.font = FONT_DADOS_BOLD if sev == 'Crítico' else FONT_DADOS
        
        c_sev = ws.cell(row=cur_r, column=6, value=sev)
        c_sev.alignment = ALIGN_CENTER
        if sev == 'Crítico':
            c_sev.fill = FILL_CRIT
            c_sev.font = FONT_CRIT
        elif sev == 'Alerta':
            c_sev.fill = FILL_ALERT
            c_sev.font = FONT_ALERT
        else:
            c_sev.fill = FILL_REG
            c_sev.font = FONT_REG
            
        ws.merge_cells(f"G{cur_r}:H{cur_r}")
        c_qtd = ws.cell(row=cur_r, column=7, value=qtd)
        c_qtd.alignment = ALIGN_RIGHT
        c_qtd.number_format = '#,##0'
        c_qtd.font = FONT_DADOS_BOLD
        
        c_pct = ws.cell(row=cur_r, column=9, value=pct)
        c_pct.alignment = ALIGN_RIGHT
        c_pct.number_format = '0.0%'
        c_pct.font = FONT_DADOS
        
        ws.merge_cells(f"J{cur_r}:L{cur_r}")
        c_acao = ws.cell(row=cur_r, column=10, value=acao)
        c_acao.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        c_acao.font = FONT_DADOS
        
        for c in range(1, 13):
            if c != 6:
                ws.cell(row=cur_r, column=c).fill = fill_zebra
            ws.cell(row=cur_r, column=c).border = BORDER_THIN
        cur_r += 1
        
    # Ajuste de Larguras de Colunas na Aba Resumo
    col_widths = {
        'A': 22, 'B': 14, 'C': 14, 'D': 15, 'E': 16,
        'F': 16, 'G': 14, 'H': 16, 'I': 16, 'J': 15,
        'K': 16, 'L': 24
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

def gravar_aba_conferencia(ws, linhas):
    """Grava os 275k registros na Aba 3: Conferência Geral em streaming ultrarrápido."""
    print("   [Aba 3/4] Gravando Conferência Geral (Streaming de ~275k linhas)...")
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A3"
    
    # Bloco de Título
    ws.merge_cells("A1:L1")
    c_t = ws["A1"]
    c_t.value = "CENCOSUD BRASIL — BASE ANALÍTICA COMPLETA DE AUDITORIA DE JORNADAS (09/2026)"
    c_t.fill = FILL_HEADER
    c_t.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    c_t.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 30
    
    headers = [
        "Regional", "Matrícula", "Nome", "Cargo", "Data",
        "Marcações", "Horas Trabalhadas", "Horas Extras", "Adicional Noturno",
        "Tipo de Divergência", "Severidade", "Ação Recomendada RH"
    ]
    
    ws.append(headers)
    ws.row_dimensions[2].height = 24
    for col in range(1, 13):
        cell = ws.cell(row=2, column=col)
        cell.fill = FILL_HEADER_SEC
        cell.font = FONT_CABECALHO
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_THIN
        
    t0 = time.time()
    for row_data in linhas:
        ws.append(row_data)
    t1 = time.time()
    print(f"   Conferência Geral gravada com sucesso em {t1 - t0:.2f}s ({len(linhas):,} linhas)!")
    
    # Larguras otimizadas (12 colunas: A a L)
    larguras = {
        'A': 14, 'B': 12, 'C': 32, 'D': 28, 'E': 12,
        'F': 26, 'G': 16, 'H': 14, 'I': 16, 'J': 34,
        'K': 14, 'L': 52
    }
    for col_l, w in larguras.items():
        ws.column_dimensions[col_l].width = w

def gravar_aba_divergencias(ws, linhas_div):
    """Grava as divergências na Aba 4 com formatação executiva, zebrado e badges de severidade."""
    print(f"   [Aba 4/4] Gravando Aba Divergências ({len(linhas_div):,} linhas com estilização executiva)...")
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A3"
    
    # Bloco de Título
    ws.merge_cells("A1:L1")
    c_t = ws["A1"]
    c_t.value = "CENCOSUD BRASIL — PAINEL OPERACIONAL DE DIVERGÊNCIAS E ALERTAS DE PONTO (09/2026)"
    c_t.fill = FILL_HEADER
    c_t.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    c_t.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 30
    
    headers = [
        "Regional", "Matrícula", "Nome", "Cargo", "Data",
        "Marcações", "Horas Trabalhadas", "Horas Extras", "Adicional Noturno",
        "Tipo de Divergência", "Severidade", "Ação Recomendada RH"
    ]
    
    ws.append(headers)
    ws.row_dimensions[2].height = 24
    for col in range(1, 13):
        cell = ws.cell(row=2, column=col)
        cell.fill = FILL_HEADER_SEC
        cell.font = FONT_CABECALHO
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_THIN
        
    t0 = time.time()
    for row_idx, r in enumerate(linhas_div, start=3):
        ws.append(r)
        ws.row_dimensions[row_idx].height = 20
        fill_zebra = FILL_ZEBRA if (row_idx % 2 == 0) else FILL_BRANCO
        sev = r[10]
        
        # Estilização das células
        for c in range(1, 13):
            cell = ws.cell(row=row_idx, column=c)
            cell.border = BORDER_THIN
            if c != 11:
                cell.fill = fill_zebra
                
            if c == 2: # Matrícula
                cell.alignment = ALIGN_CENTER
                cell.font = FONT_DADOS_BOLD
            elif c in (5, 7, 8, 9): # Data e Horas
                cell.alignment = ALIGN_CENTER
                cell.font = FONT_DADOS
            elif c == 11: # Severidade (Badge)
                cell.alignment = ALIGN_CENTER
                if sev == 'Crítico':
                    cell.fill = FILL_CRIT
                    cell.font = FONT_CRIT
                else:
                    cell.fill = FILL_ALERT
                    cell.font = FONT_ALERT
            else:
                cell.alignment = ALIGN_LEFT
                cell.font = FONT_DADOS
                
    t1 = time.time()
    print(f"   Divergências gravadas e estilizadas em {t1 - t0:.2f}s!")
    
    # Larguras otimizadas (12 colunas: A a L)
    larguras = {
        'A': 14, 'B': 12, 'C': 32, 'D': 28, 'E': 12,
        'F': 26, 'G': 16, 'H': 14, 'I': 16, 'J': 34,
        'K': 14, 'L': 52
    }
    for col_l, w in larguras.items():
        ws.column_dimensions[col_l].width = w

# ======================================================================================
# 6. EXECUÇÃO PRINCIPAL DO PIPELINE
# ======================================================================================

def main():
    print("=" * 85)
    print("CENCOSUD BRASIL - ENGENHARIA DE DADOS & AUDITORIA DE RH")
    print("SUPER MÓDULO DE JORNADA CRÍTICA: NOTURNO & HORAS EXTRAS (COMPETÊNCIA 09/2026)")
    print("=" * 85)
    
    t_global_0 = time.time()
    
    # 1. Carregar base cadastral auxiliar de Cargos
    cargo_map = carregar_mapa_cargos()
    
    # 2. Processar e auditar todas as jornadas com deduplicação
    linhas_geral, linhas_divergencias, kpis, regionais_stats, tipos_stats = processar_arquivos_jornada(cargo_map)
    
    # 3. Criar Workbook oficial
    print("\n-> Gerando Workbook Excel Oficial no Padrão Cencosud de 4 Abas...")
    wb = openpyxl.Workbook()
    
    # Configurar Abas
    ws_ori = wb.active
    ws_ori.title = "Orientações"
    
    ws_res = wb.create_sheet(title="Resumo da Auditoria")
    ws_ger = wb.create_sheet(title="Conferência Geral")
    ws_div = wb.create_sheet(title="Divergências")
    
    print("   [Aba 1/4] Construindo Orientações...")
    construir_aba_orientacoes(ws_ori)
    
    print("   [Aba 2/4] Construindo Resumo da Auditoria...")
    construir_aba_resumo(ws_res, kpis, regionais_stats, tipos_stats)
    
    gravar_aba_conferencia(ws_ger, linhas_geral)
    gravar_aba_divergencias(ws_div, linhas_divergencias)
    
    print("\n-> Salvando arquivo final em:")
    print(f"   {OUTPUT_FILE}")
    t_save_0 = time.time()
    wb.save(OUTPUT_FILE)
    t_save_1 = time.time()
    print(f"-> Arquivo salvo com sucesso em {t_save_1 - t_save_0:.2f}s!")
    
    t_global_1 = time.time()
    print("\n" + "=" * 85)
    print(f"PROCESSAMENTO GERAL CONCLUÍDO EM {t_global_1 - t_global_0:.1f} SEGUNDOS!")
    print(f"Taxa de Conformidade: {(kpis['conforme']/kpis['total']):.2%}")
    print(f"Total Auditado: {kpis['total']:,} | Conformes: {kpis['conforme']:,}")
    print(f"Divergências Críticas: {kpis['aprendiz_he'] + kpis['he_abusiva'] + kpis['he_clt']:,}")
    print(f"Alertas Operacionais: {kpis['impar'] + kpis['invertida'] + kpis['noturno_elevado']:,}")
    print("=" * 85)

if __name__ == "__main__":
    main()
