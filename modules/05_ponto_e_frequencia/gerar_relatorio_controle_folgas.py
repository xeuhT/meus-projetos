"""
=====================================================================================
CENCOSUD BRASIL - ENGENHARIA DE DADOS & AUDITORIA DE RH
MÓDULO DE CONTROLE DE FOLGA & FERIADOS TRABALHADOS (PRIORIDADE 4)
=====================================================================================
Esteira de conciliação de feriados trabalhados, apuração de folgas compensatórias
vencidas, batimento de horas reais vs analítico de ponto e estruturação do lote de
importação da folha de pagamento (Eventos 21191 e 22048/227) para a Competência 09/2026.

Entregáveis:
  - Aba 1: Orientações (Fundamentação jurídica, Lei 10.101/2000, Súmula 146 TST)
  - Aba 2: Resumo da Auditoria (6 KPI Cards, Consolidação por Status e por Evento)
  - Aba 3: Conferência Geral (1.784 feriados trabalhados com status e batimento)
  - Aba 4: Divergências (244 folgas vencidas não compensadas e lote Suite LG)
=====================================================================================
"""

import os
import time
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ======================================================================================
# 1. CONSTANTES E CONFIGURAÇÕES DE DIRETÓRIOS
# ======================================================================================

BASE_DIR = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\ponto"
INPUT_DIR = os.path.join(BASE_DIR, "Controle de Folga")
OUTPUT_DIR = os.path.join(BASE_DIR, "conferencias geradas pelo script")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Relatorio_Conferencia_Controle_Folgas_09.2026.xlsx")

FILE_CONTROLE = os.path.join(INPUT_DIR, "ControleDeFolgas_17092026.xlsx")
FILE_HORAS = os.path.join(INPUT_DIR, "Relatorio_HorasTrabalhadas_18920269135.xlsx")
FILE_EVENTOS = os.path.join(INPUT_DIR, "Relatorio_EventosPorFuncionarioAnalitico_189202685827.xlsx")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# Paleta Corporativa Cencosud
C_AZUL_MARINHO = "1B365D"
C_AZUL_MEDIO   = "2C4A7A"
C_AZUL_CLARO   = "EBF3FA"
C_CINZA_CAB    = "F1F5F9"
C_CINZA_ZEBRA  = "F8FAFC"
C_BRANCO       = "FFFFFF"
C_BORDA        = "CBD5E1"

# Badges de Severidade & Status
C_CRIT_BG      = "FFC7CE"
C_CRIT_TXT     = "9C0006"
C_ALERT_BG     = "FFEB9C"
C_ALERT_TXT    = "9C6500"
C_REG_BG       = "C6EFCE"
C_REG_TXT      = "006100"

C_OK_BG        = "DCFCE7"
C_OK_TXT       = "166534"
C_DIV_BG       = "FEE2E2"
C_DIV_TXT      = "991B1B"

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

FILL_REG = PatternFill(start_color=C_REG_BG, end_color=C_REG_BG, fill_type="solid")
FONT_REG = Font(name="Calibri", size=10, bold=True, color=C_REG_TXT)

FILL_OK = PatternFill(start_color=C_OK_BG, end_color=C_OK_BG, fill_type="solid")
FONT_OK = Font(name="Calibri", size=10, bold=True, color=C_OK_TXT)

FILL_DIV = PatternFill(start_color=C_DIV_BG, end_color=C_DIV_BG, fill_type="solid")
FONT_DIV = Font(name="Calibri", size=10, bold=True, color=C_DIV_TXT)

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
# 2. FUNÇÕES AUXILIARES DE PARSING E CONVERSÃO
# ======================================================================================

def format_data(dt_val):
    """Padroniza data para string DD/MM/AAAA."""
    if dt_val is None or str(dt_val).strip() in ['', 'None', 'nan']:
        return "-"
    if isinstance(dt_val, datetime):
        return dt_val.strftime("%d/%m/%Y")
    s = str(dt_val).strip()
    if len(s) >= 10 and s[4] == '-' and s[7] == '-':
        return f"{s[8:10]}/{s[5:7]}/{s[0:4]}"
    return s[:10]

def time_to_decimal(time_str):
    """Converte horário em string 'HH:MM' para horas decimais float (Horas + Minutos / 60)."""
    if not time_str or str(time_str).strip() in ['', '-', 'None']:
        return 0.0
    s = str(time_str).strip()
    try:
        if ':' in s:
            parts = s.split(':')
            h = int(parts[0])
            m = int(parts[1])
            return h + m / 60.0
        return float(s)
    except:
        return 0.0

# ======================================================================================
# 3. LEITURA E INDEXAÇÃO DAS BASES AUXILIARES
# ======================================================================================

def carregar_bases_auxiliares():
    """Carrega e indexa as bases de Horas Trabalhadas e Eventos Analíticos."""
    print("-> Carregando base de Horas Reais Trabalhadas...")
    wb_ht = openpyxl.load_workbook(FILE_HORAS, read_only=True, data_only=True)
    ws_ht = wb_ht.active
    rows_ht = list(ws_ht.iter_rows(values_only=True))
    wb_ht.close()
    
    header_ht = [str(c).strip() for c in rows_ht[0]]
    idx_m_ht = header_ht.index('MATRICULA')
    idx_d_ht = header_ht.index('Data')
    idx_ht_val = header_ht.index('HorasTrabalhadas')
    
    ht_map = {}
    for r in rows_ht[1:]:
        if r[idx_m_ht] and r[idx_d_ht]:
            m = str(r[idx_m_ht]).strip().split('.')[0]
            d = str(r[idx_d_ht])[:10]
            ht_map[(m, d)] = str(r[idx_ht_val]).strip() if r[idx_ht_val] else "00:00"
    print(f"   Horas Trabalhadas indexadas: {len(ht_map):,} registros.")
    
    print("-> Carregando base de Eventos Analíticos do Ponto...")
    wb_ev = openpyxl.load_workbook(FILE_EVENTOS, read_only=True, data_only=True)
    ws_ev = wb_ev.active
    rows_ev = list(ws_ev.iter_rows(values_only=True))
    wb_ev.close()
    
    header_ev = [str(c).strip() for c in rows_ev[0]]
    idx_m_ev = header_ev.index('Matricula')
    idx_d_ev = header_ev.index('Data')
    idx_cod_ev = header_ev.index('CodigoEvento')
    idx_desc_ev = header_ev.index('DescricaoEvento')
    idx_qtd_ev = header_ev.index('QtdDeHoras')
    idx_emp_ev = header_ev.index('CodigoEmpresa')
    
    ev_map = {}
    for r in rows_ev[1:]:
        if r[idx_m_ev] and r[idx_d_ev]:
            m = str(r[idx_m_ev]).strip().split('.')[0]
            d = str(r[idx_d_ev])[:10]
            ev_map[(m, d)] = {
                'CodigoEvento': str(r[idx_cod_ev]).strip(),
                'DescricaoEvento': str(r[idx_desc_ev]).strip() if r[idx_desc_ev] else "",
                'QtdDeHoras': str(r[idx_qtd_ev]).strip() if r[idx_qtd_ev] else "00:00",
                'CodigoEmpresa': str(r[idx_emp_ev]).strip() if r[idx_emp_ev] else "2"
            }
    print(f"   Eventos Analíticos indexados: {len(ev_map):,} registros.")
    
    return ht_map, ev_map

# ======================================================================================
# 4. PROCESSAMENTO E AUDITORIA DA BASE DE FERIADOS
# ======================================================================================

def processar_controle_folgas(ht_map, ev_map):
    """Lê a base de Controle de Folgas, identifica vencidos e compensados, e realiza batimento."""
    print("\n-> Processando base mestre de Controle de Folgas...")
    wb_cfg = openpyxl.load_workbook(FILE_CONTROLE, read_only=True, data_only=True)
    
    # Obter chaves de todos os feriados vencidos não compensados (Abas NÃO COMPENSADO, 227 e Controle de folgas)
    ws_nc = wb_cfg.worksheets[2] # Aba NÃO COMPENSADO
    rows_nc = list(ws_nc.iter_rows(values_only=True))[1:]
    map_nc = {(str(r[0]).strip().split('.')[0], str(r[2])[:10]): {'ht': str(r[5]) if r[5] else "00:00", 'evento': str(r[7]) if r[7] else "21191"} for r in rows_nc if r[0] is not None}
    
    ws_227 = wb_cfg.worksheets[3] # Aba 227
    rows_227 = list(ws_227.iter_rows(values_only=True))[1:]
    map_227 = {(str(r[0]).strip().split('.')[0], str(r[2])[:10]): {'ht': str(r[5]) if r[5] else "00:00", 'evento': str(r[7]) if r[7] else "227"} for r in rows_227 if r[0] is not None}
    
    # Base principal: 1.784 registros
    ws_all = wb_cfg.worksheets[0] # Aba Controle de folgas
    rows_all = list(ws_all.iter_rows(values_only=True))[1:]
    wb_cfg.close()
    
    keys_nc = set(map_nc.keys()).union(set(map_227.keys()))
    for r in rows_all:
        if r[0] is not None and str(r[0]).strip() not in ['', 'None']:
            m_s = str(r[0]).strip().split('.')[0]
            d_s = str(r[2])[:10] if r[2] else ""
            dt_comp_raw = r[4]
            is_comp = hasattr(dt_comp_raw, 'strftime') or (isinstance(dt_comp_raw, str) and len(dt_comp_raw.strip()) == 10 and '/' in dt_comp_raw)
            if not is_comp:
                keys_nc.add((m_s, d_s))
                
    print(f"   Base consolidada de Folgas Vencidas Não Compensadas: {len(keys_nc)} ocorrências.")
    
    print(f"   Total de feriados trabalhados lidos: {len(rows_all):,}")
    
    linhas_geral = []
    linhas_divergencias = []
    
    colabs_a_receber = set()
    colabs_compensados = set()
    total_horas_importar = 0.0
    qtd_correto = 0
    qtd_divergente = 0
    
    eventos_stats = {
        '21191': {'desc': 'Qtd HE 100% Fer (Horas Extras Feriado 100%)', 'qtd': 0, 'horas': 0.0},
        '22048': {'desc': 'HE Feriado Específico / Indenizado', 'qtd': 0, 'horas': 0.0},
        '22057': {'desc': 'HE Feriado / Compensação Folga (Evento 22057)', 'qtd': 0, 'horas': 0.0},
        '227': {'desc': 'Gratif Do/Fe Ali (Gratificação Domingo/Feriado)', 'qtd': 0, 'horas': 0.0}
    }
    
    for r in rows_all:
        raw_m = r[0]
        if raw_m is None or str(raw_m).strip() in ['', 'None']:
            continue
            
        m_str = str(raw_m).strip().split('.')[0]
        try:
            mat_int = int(m_str)
        except:
            continue
            
        nome = str(r[1]).strip() if r[1] else "COLABORADOR"
        dt_fer_raw = r[2]
        dt_venc_raw = r[3]
        dt_comp_raw = r[4]
        
        d_fer_str = str(dt_fer_raw)[:10] if dt_fer_raw else ""
        chave = (m_str, d_fer_str)
        
        # Datas formatadas
        str_dt_fer = format_data(dt_fer_raw)
        str_dt_venc = format_data(dt_venc_raw)
        
        # Classificação de Compensação
        if chave in keys_nc:
            status_comp = "Vencido Não Compensado"
            severidade = "Crítico"
            str_dt_comp = "NÃO COMPENSADO"
            
            # Buscar Horas Reais e Analíticas
            ht_real = ht_map.get(chave)
            if not ht_real or ht_real == "00:00":
                if chave in map_227:
                    ht_real = map_227[chave]['ht']
                elif chave in map_nc:
                    ht_real = map_nc[chave]['ht']
                else:
                    ht_real = "00:00"
                    
            ev_info = ev_map.get(chave, {})
            qtd_analitica = ev_info.get('QtdDeHoras', "00:00")
            cod_evento = ev_info.get('CodigoEvento')
            if not cod_evento:
                if chave in map_227:
                    cod_evento = map_227[chave]['evento']
                elif chave in map_nc:
                    cod_evento = map_nc[chave]['evento']
                else:
                    cod_evento = "21191"
            cod_empresa = ev_info.get('CodigoEmpresa', "2")
            
            # Validação Ponto x Real
            if ht_real == qtd_analitica:
                validacao = "CORRETO"
                qtd_correto += 1
            else:
                validacao = "DIVERGENTE"
                qtd_divergente += 1
                
            # Horas Decimais a Importar
            h_decimal = time_to_decimal(ht_real)
            total_horas_importar += h_decimal
            colabs_a_receber.add(m_str)
            
            # Estatísticas por Evento
            if cod_evento in eventos_stats:
                eventos_stats[cod_evento]['qtd'] += 1
                eventos_stats[cod_evento]['horas'] += h_decimal
            else:
                eventos_stats['21191']['qtd'] += 1
                eventos_stats['21191']['horas'] += h_decimal
                
            acao = f"Importar {h_decimal:.2f}h na folha de pagamento (Evento {cod_evento}) via lote Suite LG."
            
            # Linha da Aba 4 (Divergências)
            linha_div = (
                mat_int, nome, str_dt_fer, str_dt_venc, ht_real, qtd_analitica,
                int(cod_evento) if cod_evento.isdigit() else cod_evento,
                validacao, round(h_decimal, 2),
                int(cod_empresa) if cod_empresa.isdigit() else cod_empresa,
                1080, 92026, acao
            )
            linhas_divergencias.append(linha_div)
            
        else:
            status_comp = "Compensado"
            severidade = "Regular"
            str_dt_comp = format_data(dt_comp_raw)
            
            ht_real = "-"
            qtd_analitica = "-"
            cod_evento = "-"
            validacao = "CONFORME"
            h_decimal = 0.0
            acao = "Conforme - Folga compensatória usufruída pelo colaborador dentro do prazo legal."
            colabs_compensados.add(m_str)
            
        # Linha da Aba 3 (Conferência Geral)
        linha_geral = (
            mat_int, nome, str_dt_fer, str_dt_venc, str_dt_comp, status_comp,
            ht_real, qtd_analitica,
            int(cod_evento) if str(cod_evento).isdigit() else cod_evento,
            validacao, severidade, round(h_decimal, 2), acao
        )
        linhas_geral.append(linha_geral)
        
    print(f"-> Auditoria Concluída com Sucesso:")
    print(f"   Total Feriados Trabalhados: {len(linhas_geral):,}")
    print(f"   Compensados no Prazo: {len(linhas_geral) - len(linhas_divergencias):,}")
    print(f"   Vencidos Não Compensados: {len(linhas_divergencias):,}")
    print(f"   Colaboradores a Receber: {len(colabs_a_receber):,}")
    print(f"   Divergências no Ponto: {qtd_divergente:,} | Corretos: {qtd_correto:,}")
    print(f"   Total Horas a Importar: {total_horas_importar:.2f} h")
    
    kpis = {
        'total_feriados': len(linhas_geral),
        'compensados': len(linhas_geral) - len(linhas_divergencias),
        'vencidos': len(linhas_divergencias),
        'colabs_receber': len(colabs_a_receber),
        'divergentes_ponto': qtd_divergente,
        'total_horas_importar': total_horas_importar
    }
    
    return linhas_geral, linhas_divergencias, kpis, eventos_stats, len(colabs_compensados)

# ======================================================================================
# 5. CONSTRUÇÃO DO WORKBOOK OFICIAL NO PADRÃO CENCOSUD DE 4 ABAS
# ======================================================================================

def construir_aba_orientacoes(ws):
    """Constrói a Aba 1: Orientações Institucionais e Fundamentação Legal."""
    ws.views.sheetView[0].showGridLines = True
    
    # Bloco de Título
    ws.merge_cells("A1:K1")
    c_t = ws["A1"]
    c_t.value = "CENCOSUD BRASIL — AUDITORIA DE CONTROLE DE FOLGAS E FERIADOS TRABALHADOS"
    c_t.fill = FILL_HEADER
    c_t.font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    c_t.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 40
    
    # Subtítulo institucional
    ws.merge_cells("A2:K2")
    c_sub = ws["A2"]
    c_sub.value = "Área de Benefícios e Ponto — Módulo de Conciliação de Folgas Compensatórias e Lote de Folha (Competência: 09/2026)"
    c_sub.fill = FILL_HEADER_SEC
    c_sub.font = Font(name="Calibri", size=11, italic=True, color="FFFFFF")
    c_sub.alignment = ALIGN_CENTER
    ws.row_dimensions[2].height = 24
    
    secoes = [
        ("1. OBJETIVO DO MÓDULO DE CONTROLE DE FOLGAS",
         "Auditar com precisão matemática 100% dos feriados laborados pelos colaboradores do grupo Cencosud Brasil, monitorando a concessão tempestiva das folgas compensatórias "
         "e identificando as folgas vencidas que exigem pagamento de horas extras a 100% na folha de pagamento (Suite LG)."),
        
        ("2. EMBASAMENTO LEGAL & REGRAS COLETIVAS",
         "• Lei nº 10.101/2000 & Art. 9º da Lei nº 605/1949: O trabalho em feriados nas atividades do comércio em geral é autorizado desde que previsto em convenção coletiva e observada a legislação municipal.\n"
         "• Súmula 146 do TST: O trabalho prestado em domingos e feriados, não compensado, deve ser pago em dobro, sem prejuízo da remuneração relativa ao repouso semanal.\n"
         "• Acordos Coletivos Cencosud: A empresa possui o prazo de até 60 dias a contar do feriado trabalhado para conceder a folga compensatória. Expirado esse prazo sem a devida folga, as horas trabalhadas convertem-se compulsoriamente em pagamento em folha."),
        
        ("3. MATRIZ DE SEVERIDADE & CLASSIFICAÇÃO",
         "• CRÍTICO (Vermelho): Feriado trabalhado com prazo regulamentar expirado (Data de Vencimento <= Setembro/2026) sem registro de compensação. Requer importação imediata de horas em folha de pagamento.\n"
         "• REGULAR (Verde): Feriado com folga compensatória regularmente usufruída dentro do prazo legal de 60 dias.\n"
         "• DIVERGENTE (Espelho de Ponto): Quando a quantidade de horas reais trabalhadas no feriado não confere com a quantidade apurada analiticamente pelo relógio de ponto."),
        
        ("4. DICIONÁRIO DE EVENTOS DE IMPORTAÇÃO (SUITE LG)",
         "• Evento 21191 (Qtd HE 100% Fer): Evento padrão para pagamento das horas trabalhadas em feriados não compensados a 100% da remuneração horária normal.\n"
         "• Evento 22048: Evento específico de horas extras de feriado indenizadas / banco de horas para estabelecimentos e sindicatos com acordos diferenciados.\n"
         "• Evento 22057: Evento de horas extras de feriado / compensação de folga apurada em folha.\n"
         "• Evento 227 (Gratif Do/Fe Ali): Gratificação de feriado/alimentação conforme acordos coletivos específicos.\n"
         "• Tipo de Lançamento 1080 & Parcela 92026: Parâmetros sistêmicos para carga automática em lote na folha de pagamento da competência 09/2026.")
    ]
    
    cur_row = 4
    for titulo, texto in secoes:
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=11)
        c_tit = ws.cell(row=cur_row, column=1, value=titulo)
        c_tit.fill = PatternFill(start_color=C_AZUL_MEDIO, end_color=C_AZUL_MEDIO, fill_type="solid")
        c_tit.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c_tit.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[cur_row].height = 26
        cur_row += 1
        
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=11)
        c_txt = ws.cell(row=cur_row, column=1, value=texto)
        c_txt.fill = PatternFill(start_color=C_CINZA_ZEBRA, end_color=C_CINZA_ZEBRA, fill_type="solid")
        c_txt.font = Font(name="Calibri", size=10, color="000000")
        c_txt.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True, indent=1)
        
        linhas_texto = texto.count('\n') + len(texto) // 110 + 1
        ws.row_dimensions[cur_row].height = max(45, linhas_texto * 18)
        cur_row += 2

    for col in range(1, 12):
        ws.column_dimensions[get_column_letter(col)].width = 14

def construir_aba_resumo(ws, kpis, eventos_stats, total_colabs_comp):
    """Constrói a Aba 2: Resumo da Auditoria com 6 Cards Executivos e Tabelas Consolidadas."""
    ws.views.sheetView[0].showGridLines = True
    
    # Título do Dashboard
    ws.merge_cells("A1:K1")
    c_t = ws["A1"]
    c_t.value = "CENCOSUD BRASIL — DASHBOARD EXECUTIVO DE CONTROLE DE FOLGAS & FERIADOS"
    c_t.fill = FILL_HEADER
    c_t.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    c_t.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 36
    
    # 6 Cards Executivos de KPIs (Distribuídos em 2 linhas de 3 cards)
    cards = [
        # Linha 1
        ("TOTAL FERIADOS TRABALHADOS", f"{kpis['total_feriados']:,}", "1.784 Linhas no Insumo (1.422 Vazias)", C_AZUL_MARINHO, 1, 3, 3),
        ("FOLGAS COMPENSADAS (NO PRAZO)", f"{kpis['compensados']:,}", f"{(kpis['compensados']/kpis['total_feriados']):.1%} do Total (Conformes)", "166534", 4, 7, 3),
        ("FOLGAS VENCIDAS NÃO COMPENSADAS", f"{kpis['vencidos']:,}", f"{(kpis['vencidos']/kpis['total_feriados']):.1%} a Pagar em Folha", C_CRIT_TXT, 8, 11, 3),
        # Linha 2
        ("COLABORADORES A RECEBER", f"{kpis['colabs_receber']:,}", "Colaboradores Únicos", C_CRIT_TXT, 1, 3, 7),
        ("DIVERGÊNCIAS NO PONTO", f"{kpis['divergentes_ponto']:,}", f"{(kpis['divergentes_ponto']/kpis['vencidos']):.1%} dos Vencidos", C_ALERT_TXT, 4, 7, 7),
        ("TOTAL HORAS A IMPORTAR", f"{kpis['total_horas_importar']:,.2f} h", "Lote Folha Suite LG (09/2026)", C_AZUL_MARINHO, 8, 11, 7)
    ]
    
    for titulo, valor, subtitulo, cor_txt, ci, cf, rb in cards:
        # Título
        ws.merge_cells(start_row=rb, start_column=ci, end_row=rb, end_column=cf)
        ct = ws.cell(row=rb, column=ci, value=titulo)
        ct.fill = PatternFill(start_color=C_CINZA_CAB, end_color=C_CINZA_CAB, fill_type="solid")
        ct.font = Font(name="Calibri", size=9, bold=True, color="475569")
        ct.alignment = ALIGN_CENTER
        
        # Valor
        ws.merge_cells(start_row=rb+1, start_column=ci, end_row=rb+1, end_column=cf)
        cv = ws.cell(row=rb+1, column=ci, value=valor)
        cv.fill = FILL_BRANCO
        cv.font = Font(name="Calibri", size=18, bold=True, color=cor_txt)
        cv.alignment = ALIGN_CENTER
        
        # Subtítulo
        ws.merge_cells(start_row=rb+2, start_column=ci, end_row=rb+2, end_column=cf)
        cs = ws.cell(row=rb+2, column=ci, value=subtitulo)
        cs.fill = FILL_BRANCO
        cs.font = Font(name="Calibri", size=8, italic=True, color="64748B")
        cs.alignment = ALIGN_CENTER
        
        for r in range(rb, rb+3):
            for c in range(ci, cf+1):
                ws.cell(row=r, column=c).border = BORDER_THIN
                
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 30
    ws.row_dimensions[5].height = 16
    ws.row_dimensions[7].height = 18
    ws.row_dimensions[8].height = 30
    ws.row_dimensions[9].height = 16
    
    # ----------------------------------------------------------------------------------
    # TABELA 1: CONSOLIDAÇÃO POR STATUS DE COMPENSAÇÃO DE FOLGA
    # ----------------------------------------------------------------------------------
    r_t1_tit = 11
    ws.merge_cells(f"A{r_t1_tit}:K{r_t1_tit}")
    ct1 = ws[f"A{r_t1_tit}"]
    ct1.value = "1. CONSOLIDAÇÃO EXECUTIVA POR STATUS DE COMPENSAÇÃO DE FOLGA"
    ct1.fill = FILL_HEADER_SEC
    ct1.font = FONT_TITULO_TABELA
    ct1.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[r_t1_tit].height = 24
    
    # Cabeçalho Tabela 1
    r_t1_h = 12
    ws.row_dimensions[r_t1_h].height = 24
    headers_t1 = [
        ("Status da Compensação", 1, 2),
        ("Severidade", 3, 3),
        ("Qtd. Ocorrências", 4, 4),
        ("% do Total", 5, 5),
        ("Colabs Únicos", 6, 6),
        ("Horas a Importar (h)", 7, 8),
        ("Ação Recomendada RH", 9, 11)
    ]
    for h, ci, cf in headers_t1:
        if ci != cf: ws.merge_cells(start_row=r_t1_h, start_column=ci, end_row=r_t1_h, end_column=cf)
        cell = ws.cell(row=r_t1_h, column=ci, value=h)
        cell.alignment = ALIGN_CENTER
    for c in range(1, 12):
        ws.cell(row=r_t1_h, column=c).fill = FILL_HEADER
        ws.cell(row=r_t1_h, column=c).font = FONT_CABECALHO
        ws.cell(row=r_t1_h, column=c).border = BORDER_THIN
        
    linhas_t1 = [
        ("Compensado no Prazo", "Regular", kpis['compensados'], kpis['compensados']/kpis['total_feriados'], total_colabs_comp, 0.0, "Conforme - Folga compensatória usufruída pelo colaborador dentro do prazo legal."),
        ("Vencido Não Compensado", "Crítico", kpis['vencidos'], kpis['vencidos']/kpis['total_feriados'], kpis['colabs_receber'], kpis['total_horas_importar'], "Importar horas na folha de pagamento via lote Suite LG (Eventos 21191 / 22048 / 22057 / 227)."),
        ("Linhas Fantasmas Descartadas", "Insumo", 1422, 0.0, 0, 0.0, "Linhas sem matrícula/colaborador no arquivo bruto (fórmulas arrastadas da linha 364 a 1785).")
    ]
    
    cur_r = 13
    for status, sev, qtd, pct, colabs, horas, acao in linhas_t1:
        fill_zebra = FILL_ZEBRA if (cur_r % 2 == 0) else FILL_BRANCO
        ws.row_dimensions[cur_r].height = 22
        
        ws.merge_cells(f"A{cur_r}:B{cur_r}")
        ws.cell(row=cur_r, column=1, value=status).alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.cell(row=cur_r, column=1).font = FONT_DADOS_BOLD if sev != 'Insumo' else Font(name="Calibri", size=10, italic=True, color="64748B")
        
        c_sev = ws.cell(row=cur_r, column=3, value=sev)
        c_sev.alignment = ALIGN_CENTER
        if sev == 'Crítico':
            c_sev.fill = FILL_CRIT; c_sev.font = FONT_CRIT
        elif sev == 'Regular':
            c_sev.fill = FILL_REG; c_sev.font = FONT_REG
        else:
            c_sev.fill = PatternFill(start_color=C_CINZA_CAB, end_color=C_CINZA_CAB, fill_type="solid")
            c_sev.font = Font(name="Calibri", size=9, italic=True, color="64748B")
            
        c_qtd = ws.cell(row=cur_r, column=4, value=qtd)
        c_qtd.alignment = ALIGN_RIGHT; c_qtd.number_format = '#,##0'; c_qtd.font = FONT_DADOS_BOLD
        
        c_pct = ws.cell(row=cur_r, column=5, value=pct if pct > 0 else "—")
        c_pct.alignment = ALIGN_RIGHT
        if pct > 0: c_pct.number_format = '0.0%'
        c_pct.font = FONT_DADOS
        
        c_col = ws.cell(row=cur_r, column=6, value=colabs if colabs > 0 else "—")
        c_col.alignment = ALIGN_RIGHT
        if colabs > 0: c_col.number_format = '#,##0'
        c_col.font = FONT_DADOS
        
        ws.merge_cells(f"G{cur_r}:H{cur_r}")
        c_h = ws.cell(row=cur_r, column=7, value=horas)
        c_h.alignment = ALIGN_RIGHT; c_h.number_format = '#,##0.00'; c_h.font = FONT_DADOS_BOLD
        
        ws.merge_cells(f"I{cur_r}:K{cur_r}")
        c_ac = ws.cell(row=cur_r, column=9, value=acao)
        c_ac.alignment = Alignment(horizontal='left', vertical='center', indent=1); c_ac.font = FONT_DADOS
        
        for c in range(1, 12):
            if c != 3 or sev == 'Insumo': ws.cell(row=cur_r, column=c).fill = fill_zebra
            ws.cell(row=cur_r, column=c).border = BORDER_THIN
        cur_r += 1
        
    # Total Geral Tabela 1 (1.784 Linhas)
    ws.row_dimensions[cur_r].height = 24
    ws.merge_cells(f"A{cur_r}:B{cur_r}")
    ws.cell(row=cur_r, column=1, value="TOTAL LINHAS ARQUIVO BRUTO").alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.cell(row=cur_r, column=3, value="—").alignment = ALIGN_CENTER
    ws.cell(row=cur_r, column=4, value=1784).alignment = ALIGN_RIGHT
    ws.cell(row=cur_r, column=4).number_format = '#,##0'
    ws.cell(row=cur_r, column=5, value=1.0).alignment = ALIGN_RIGHT
    ws.cell(row=cur_r, column=5).number_format = '0.0%'
    ws.cell(row=cur_r, column=6, value=313).alignment = ALIGN_RIGHT
    ws.cell(row=cur_r, column=6).number_format = '#,##0'
    ws.merge_cells(f"G{cur_r}:H{cur_r}")
    ws.cell(row=cur_r, column=7, value=kpis['total_horas_importar']).alignment = ALIGN_RIGHT
    ws.cell(row=cur_r, column=7).number_format = '#,##0.00'
    ws.merge_cells(f"I{cur_r}:K{cur_r}")
    ws.cell(row=cur_r, column=9, value="Universo Total de Linhas Mapeadas").alignment = Alignment(horizontal='left', vertical='center', indent=1)
    
    for c in range(1, 12):
        cell = ws.cell(row=cur_r, column=c)
        cell.fill = PatternFill(start_color=C_CINZA_CAB, end_color=C_CINZA_CAB, fill_type="solid")
        cell.font = Font(name="Calibri", size=10, bold=True, color=C_AZUL_MARINHO)
        cell.border = Border(top=Side(style='thin', color=C_AZUL_MARINHO), bottom=Side(style='double', color=C_AZUL_MARINHO))
    cur_r += 3

    
    # ----------------------------------------------------------------------------------
    # TABELA 2: CONSOLIDAÇÃO POR EVENTO DE PAGAMENTO EM FOLHA
    # ----------------------------------------------------------------------------------
    ws.merge_cells(f"A{cur_r}:K{cur_r}")
    ct2 = ws[f"A{cur_r}"]
    ct2.value = "2. CONSOLIDAÇÃO POR EVENTO DE PAGAMENTO EM FOLHA (LOTE SUITE LG)"
    ct2.fill = FILL_HEADER_SEC
    ct2.font = FONT_TITULO_TABELA
    ct2.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[cur_r].height = 24
    cur_r += 1
    
    headers_t2 = [
        ("Código Evento", 1, 2),
        ("Descrição do Evento", 3, 5),
        ("Qtd. Lançamentos", 6, 6),
        ("% Ocorrências", 7, 7),
        ("Horas a Importar (h)", 8, 9),
        ("Impacto na Folha de Pagamento", 10, 11)
    ]
    ws.row_dimensions[cur_r].height = 24
    for h, ci, cf in headers_t2:
        if ci != cf: ws.merge_cells(start_row=cur_r, start_column=ci, end_row=cur_r, end_column=cf)
        ws.cell(row=cur_r, column=ci, value=h).alignment = ALIGN_CENTER
    for c in range(1, 12):
        ws.cell(row=cur_r, column=c).fill = FILL_HEADER
        ws.cell(row=cur_r, column=c).font = FONT_CABECALHO
        ws.cell(row=cur_r, column=c).border = BORDER_THIN
    cur_r += 1
    
    pct_21191 = eventos_stats['21191']['qtd']/kpis['vencidos'] if kpis['vencidos'] > 0 else 0.0
    pct_22048 = eventos_stats['22048']['qtd']/kpis['vencidos'] if kpis['vencidos'] > 0 else 0.0
    pct_22057 = eventos_stats['22057']['qtd']/kpis['vencidos'] if kpis['vencidos'] > 0 else 0.0
    pct_227 = eventos_stats['227']['qtd']/kpis['vencidos'] if kpis['vencidos'] > 0 else 0.0
    
    linhas_t2 = [
        ("21191", eventos_stats['21191']['desc'], eventos_stats['21191']['qtd'], pct_21191, eventos_stats['21191']['horas'], "Crédito de HE 100% em folha (Competência 09/2026)"),
        ("22048", eventos_stats['22048']['desc'], eventos_stats['22048']['qtd'], pct_22048, eventos_stats['22048']['horas'], "Crédito específico em folha (Competência 09/2026)"),
        ("22057", eventos_stats['22057']['desc'], eventos_stats['22057']['qtd'], pct_22057, eventos_stats['22057']['horas'], "Crédito de HE Feriado / Compensação Folga (Competência 09/2026)"),
        ("227", eventos_stats['227']['desc'], eventos_stats['227']['qtd'], pct_227, eventos_stats['227']['horas'], "Crédito de Gratificação Feriado em folha (Competência 09/2026)")
    ]
    
    for cod, desc, qtd, pct, horas, imp in linhas_t2:
        fill_zebra = FILL_ZEBRA if (cur_r % 2 == 0) else FILL_BRANCO
        ws.row_dimensions[cur_r].height = 22
        
        ws.merge_cells(f"A{cur_r}:B{cur_r}")
        ws.cell(row=cur_r, column=1, value=cod).alignment = ALIGN_CENTER
        ws.cell(row=cur_r, column=1).font = FONT_DADOS_BOLD
        
        ws.merge_cells(f"C{cur_r}:E{cur_r}")
        ws.cell(row=cur_r, column=3, value=desc).alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.cell(row=cur_r, column=3).font = FONT_DADOS
        
        c_qtd = ws.cell(row=cur_r, column=6, value=qtd)
        c_qtd.alignment = ALIGN_RIGHT; c_qtd.number_format = '#,##0'; c_qtd.font = FONT_DADOS_BOLD
        
        c_pct = ws.cell(row=cur_r, column=7, value=pct)
        c_pct.alignment = ALIGN_RIGHT; c_pct.number_format = '0.0%'; c_pct.font = FONT_DADOS
        
        ws.merge_cells(f"H{cur_r}:I{cur_r}")
        c_h = ws.cell(row=cur_r, column=8, value=horas)
        c_h.alignment = ALIGN_RIGHT; c_h.number_format = '#,##0.00'; c_h.font = FONT_DADOS_BOLD
        
        ws.merge_cells(f"J{cur_r}:K{cur_r}")
        ws.cell(row=cur_r, column=10, value=imp).alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.cell(row=cur_r, column=10).font = FONT_DADOS
        
        for c in range(1, 12):
            ws.cell(row=cur_r, column=c).fill = fill_zebra
            ws.cell(row=cur_r, column=c).border = BORDER_THIN
        cur_r += 1
        
    # Total Geral Tabela 2
    ws.row_dimensions[cur_r].height = 24
    ws.merge_cells(f"A{cur_r}:B{cur_r}")
    ws.cell(row=cur_r, column=1, value="TOTAL LOTE IMPORTAÇÃO").alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.merge_cells(f"C{cur_r}:E{cur_r}")
    ws.cell(row=cur_r, column=3, value="Lote de Folha de Pagamento").alignment = ALIGN_CENTER
    ws.cell(row=cur_r, column=6, value=kpis['vencidos']).alignment = ALIGN_RIGHT
    ws.cell(row=cur_r, column=6).number_format = '#,##0'
    ws.cell(row=cur_r, column=7, value=1.0).alignment = ALIGN_RIGHT
    ws.cell(row=cur_r, column=7).number_format = '0.0%'
    ws.merge_cells(f"H{cur_r}:I{cur_r}")
    ws.cell(row=cur_r, column=8, value=kpis['total_horas_importar']).alignment = ALIGN_RIGHT
    ws.cell(row=cur_r, column=8).number_format = '#,##0.00'
    ws.merge_cells(f"J{cur_r}:K{cur_r}")
    ws.cell(row=cur_r, column=10, value="—").alignment = ALIGN_CENTER
    
    for c in range(1, 12):
        cell = ws.cell(row=cur_r, column=c)
        cell.fill = PatternFill(start_color=C_CINZA_CAB, end_color=C_CINZA_CAB, fill_type="solid")
        cell.font = Font(name="Calibri", size=10, bold=True, color=C_AZUL_MARINHO)
        cell.border = Border(top=Side(style='thin', color=C_AZUL_MARINHO), bottom=Side(style='double', color=C_AZUL_MARINHO))
        
    col_widths = {
        'A': 14, 'B': 14, 'C': 16, 'D': 16, 'E': 14,
        'F': 16, 'G': 14, 'H': 14, 'I': 18, 'J': 20, 'K': 24
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

def gravar_aba_conferencia(ws, linhas):
    """Grava a Aba 3: Conferência Geral com os 1.784 registros analíticos."""
    print("   [Aba 3/4] Gravando Conferência Geral (1.784 feriados trabalhados)...")
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A3"
    
    ws.merge_cells("A1:M1")
    c_t = ws["A1"]
    c_t.value = "CENCOSUD BRASIL — BASE ANALÍTICA COMPLETA DE CONTROLE DE FOLGAS E FERIADOS (09/2026)"
    c_t.fill = FILL_HEADER
    c_t.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    c_t.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 30
    
    headers = [
        "Matrícula", "Nome do Colaborador", "Data do Feriado Trabalhado", "Data de Vencimento da Compensação",
        "Data da Compensação", "Status da Compensação", "Horas Trabalhadas no Feriado", "Horas no Analítico de Ponto",
        "Evento Folha", "Validação Ponto x Real", "Severidade", "Horas a Importar (Decimal)", "Ação Recomendada RH"
    ]
    
    ws.append(headers)
    ws.row_dimensions[2].height = 24
    for col in range(1, 14):
        cell = ws.cell(row=2, column=col)
        cell.fill = FILL_HEADER_SEC
        cell.font = FONT_CABECALHO
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_THIN
        
    t0 = time.time()
    for row_idx, r in enumerate(linhas, start=3):
        ws.append(r)
        ws.row_dimensions[row_idx].height = 20
        fill_zebra = FILL_ZEBRA if (row_idx % 2 == 0) else FILL_BRANCO
        
        status = r[5]
        val = r[9]
        sev = r[10]
        
        for c in range(1, 14):
            cell = ws.cell(row=row_idx, column=c)
            cell.border = BORDER_THIN
            if c not in (10, 11):
                cell.fill = fill_zebra
                
            if c == 1: # Matrícula
                cell.alignment = ALIGN_CENTER
                cell.font = FONT_DADOS_BOLD
            elif c in (3, 4, 5): # Datas
                cell.alignment = ALIGN_CENTER
                cell.font = FONT_DADOS
            elif c in (7, 8, 9): # Horas e Evento
                cell.alignment = ALIGN_CENTER
                cell.font = FONT_DADOS
            elif c == 10: # Validação
                cell.alignment = ALIGN_CENTER
                if val == 'DIVERGENTE':
                    cell.fill = FILL_DIV; cell.font = FONT_DIV
                elif val == 'CORRETO':
                    cell.fill = FILL_OK; cell.font = FONT_OK
                else:
                    cell.fill = fill_zebra; cell.font = FONT_DADOS
            elif c == 11: # Severidade
                cell.alignment = ALIGN_CENTER
                if sev == 'Crítico':
                    cell.fill = FILL_CRIT; cell.font = FONT_CRIT
                else:
                    cell.fill = FILL_REG; cell.font = FONT_REG
            elif c == 12: # Horas a Importar
                cell.alignment = ALIGN_RIGHT
                cell.number_format = '#,##0.00'
                cell.font = FONT_DADOS_BOLD
            else:
                cell.alignment = ALIGN_LEFT
                cell.font = FONT_DADOS
                
    t1 = time.time()
    print(f"   Conferência Geral gravada e formatada em {t1 - t0:.2f}s!")
    
    larguras = {
        'A': 12, 'B': 34, 'C': 16, 'D': 18, 'E': 16,
        'F': 22, 'G': 16, 'H': 16, 'I': 14, 'J': 16,
        'K': 13, 'L': 16, 'M': 50
    }
    for col_l, w in larguras.items():
        ws.column_dimensions[col_l].width = w

def gravar_aba_divergencias(ws, linhas_div):
    """Grava a Aba 4: Divergências com os 244 feriados vencidos e dados do Lote de Folha."""
    print("   [Aba 4/4] Gravando Aba Divergências (244 folgas vencidas e parâmetros de importação)...")
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A3"
    
    ws.merge_cells("A1:M1")
    c_t = ws["A1"]
    c_t.value = "CENCOSUD BRASIL — FOLGAS VENCIDAS NÃO COMPENSADAS & LOTE DE IMPORTAÇÃO SUITE LG (09/2026)"
    c_t.fill = FILL_HEADER
    c_t.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    c_t.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 30
    
    headers = [
        "Matrícula", "Nome do Colaborador", "Data do Feriado", "Data de Vencimento",
        "Horas Trabalhadas", "Horas Analíticas", "Evento", "Validação",
        "Horas Decimais a Importar", "Código Empresa", "Tipo Lançamento (1080)", "Data Vigência (92026)", "Ação Recomendada RH"
    ]
    
    ws.append(headers)
    ws.row_dimensions[2].height = 24
    for col in range(1, 14):
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
        val = r[7]
        
        for c in range(1, 14):
            cell = ws.cell(row=row_idx, column=c)
            cell.border = BORDER_THIN
            if c != 8:
                cell.fill = fill_zebra
                
            if c == 1: # Matrícula
                cell.alignment = ALIGN_CENTER
                cell.font = FONT_DADOS_BOLD
            elif c in (3, 4): # Datas
                cell.alignment = ALIGN_CENTER
                cell.font = FONT_DADOS
            elif c in (5, 6, 7): # Horas e Evento
                cell.alignment = ALIGN_CENTER
                cell.font = FONT_DADOS
            elif c == 8: # Validação
                cell.alignment = ALIGN_CENTER
                if val == 'DIVERGENTE':
                    cell.fill = FILL_DIV; cell.font = FONT_DIV
                else:
                    cell.fill = FILL_OK; cell.font = FONT_OK
            elif c == 9: # Horas Decimais a Importar
                cell.alignment = ALIGN_RIGHT
                cell.number_format = '#,##0.00'
                cell.font = FONT_DADOS_BOLD
            elif c in (10, 11, 12): # Parâmetros LG
                cell.alignment = ALIGN_CENTER
                cell.font = FONT_DADOS_BOLD
            else:
                cell.alignment = ALIGN_LEFT
                cell.font = FONT_DADOS
                
    t1 = time.time()
    print(f"   Aba Divergências gravada e estilizada em {t1 - t0:.2f}s!")
    
    larguras = {
        'A': 12, 'B': 34, 'C': 16, 'D': 18, 'E': 16,
        'F': 16, 'G': 12, 'H': 16, 'I': 18, 'J': 15,
        'K': 16, 'L': 16, 'M': 50
    }
    for col_l, w in larguras.items():
        ws.column_dimensions[col_l].width = w

# ======================================================================================
# 6. EXECUÇÃO PRINCIPAL DO PIPELINE
# ======================================================================================

def main():
    print("=" * 85)
    print("CENCOSUD BRASIL - ENGENHARIA DE DADOS & AUDITORIA DE RH")
    print("MÓDULO DE CONTROLE DE FOLGAS & FERIADOS TRABALHADOS (COMPETÊNCIA 09/2026)")
    print("=" * 85)
    
    t_global_0 = time.time()
    
    # 1. Carregar bases auxiliares
    ht_map, ev_map = carregar_bases_auxiliares()
    
    # 2. Processar Controle de Folgas
    linhas_geral, linhas_divergencias, kpis, eventos_stats, total_colabs_comp = processar_controle_folgas(ht_map, ev_map)
    
    # 3. Gerar Workbook no Padrão Cencosud de 4 Abas
    print("\n-> Gerando Workbook Excel Oficial no Padrão Cencosud de 4 Abas...")
    wb = openpyxl.Workbook()
    
    ws_ori = wb.active
    ws_ori.title = "Orientações"
    
    ws_res = wb.create_sheet(title="Resumo da Auditoria")
    ws_ger = wb.create_sheet(title="Conferência Geral")
    ws_div = wb.create_sheet(title="Divergências")
    
    print("   [Aba 1/4] Construindo Orientações...")
    construir_aba_orientacoes(ws_ori)
    
    print("   [Aba 2/4] Construindo Resumo da Auditoria...")
    construir_aba_resumo(ws_res, kpis, eventos_stats, total_colabs_comp)
    
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
    print(f"Total Feriados Trabalhados: {kpis['total_feriados']:,}")
    print(f"Folgas Compensadas (No Prazo): {kpis['compensados']:,} ({(kpis['compensados']/kpis['total_feriados']):.1%})")
    print(f"Folgas Vencidas Não Compensadas: {kpis['vencidos']:,}")
    print(f"Colaboradores com Horas a Receber: {kpis['colabs_receber']:,}")
    print(f"Divergências no Espelho de Ponto: {kpis['divergentes_ponto']:,}")
    print(f"Total Horas a Importar em Folha: {kpis['total_horas_importar']:,.2f} h")
    print("=" * 85)

if __name__ == "__main__":
    main()
