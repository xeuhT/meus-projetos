"""
========================================================================================
CENCOSUD BRASIL - ENGENHARIA DE DADOS & AUDITORIA DE RH
SISTEMA DE AUTOMAÇÃO DE CONFERÊNCIAS DE PONTO E FOLHA
MÓDULO: CONFERÊNCIA DE INTERMITENTES (COMPETÊNCIA 09/2026)
========================================================================================
Autor: Agente Desenvolvedor (Pair Engineering com Arquiteto de Dados & Validação Crítica)
Data: Outubro/2026
Objetivo: Processar e auditar 100% dos colaboradores intermitentes do fechamento mensal,
cruzando dados do SAV, Ponto Analítico (LG), Folha de Pagamento (FOPA) e Perfil de Frequência.
Gera relatório executivo em Excel no Padrão Cencosud de 4 Abas:
  1. Orientações
  2. Resumo da Auditoria (KPIs e Visão Loja)
  3. Conferência Geral (População Completa - 55 Colaboradores)
  4. Divergências (Filtro Exclusivo de Pendências / Alertas)
========================================================================================
"""

import os
import sys
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ======================================================================================
# 1. CONFIGURAÇÕES E CAMINHOS DO ECOSSISTEMA
# ======================================================================================
BASE_DIR = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\ponto"
INTER_DIR = os.path.join(BASE_DIR, "Intermitentes")

HIST_FILE = os.path.join(BASE_DIR, "1 - BASE DE CONFERENCIA INTERMITENTE 08.2026.xlsb")
SAV_FILE = os.path.join(INTER_DIR, "RelatorioIntermitente_01.10.2026.xlsx")
PONTO_FILE = os.path.join(INTER_DIR, "Relatorio_EventosPorFuncionarioAnalitico_1102026142056.xlsx")
FOPA_FILE = os.path.join(INTER_DIR, "FOPA_Eventos_Calculados _Linha - 01102026_153524_464658.CSV")
PERFIL_FILE = os.path.join(INTER_DIR, "Escala_e_Perfil_de_Frequência_Atual_do_Colaborador30092026.csv")

OUTPUT_DIR = os.path.join(BASE_DIR, "conferencias geradas pelo script")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Relatorio_Conferencia_Intermitentes_09.2026.xlsx")

# ======================================================================================
# 2. FUNÇÕES DE INGESTÃO E LIMPEZA DE DADOS BRUTOS
# ======================================================================================

def carregar_parametros_sav(caminho_hist):
    """Carrega tabela paramétrica de sindicatos e valores de SAV por loja."""
    print("-> Carregando Tabela Paramétrica VALOR SAV...")
    df = pd.read_excel(caminho_hist, sheet_name='VALOR SAV', engine='pyxlsb')
    df.columns = [c.strip().replace('\n', ' ') for c in df.columns]
    # Tratar a duplicidade da loja 905 (possui 2 sindicatos com valores idênticos)
    df = df.drop_duplicates(subset=['Loja']).copy()
    return df

def carregar_sav(caminho_sav):
    """Ingere e processa as solicitações de diárias de intermitentes no SAV."""
    print("-> Carregando e processando SAV (Diárias de Intermitentes)...")
    df_raw = pd.read_excel(caminho_sav)
    # Filtrar estritamente 'Diaria Intermitente'
    df_diaria = df_raw[df_raw['Tipo'] == 'Diaria Intermitente'].copy()
    
    # Identificar solicitações duplicadas na mesma data
    df_dups = df_diaria[df_diaria.duplicated(subset=['Matrícula', 'Data Solicitação'], keep=False)]
    mats_dups = set(df_dups['Matrícula'].unique())
    
    # Detalhar datas duplicadas por matrícula
    dup_details = {}
    if len(df_dups) > 0:
        for mat, grp in df_dups.groupby('Matrícula'):
            datas_str = ", ".join(grp['Data Solicitação'].dt.strftime('%d/%m').unique())
            dup_details[mat] = f"DIÁRIA DUPLICADA NO SAV (Data: {datas_str})"
    
    # Totalizar por matrícula
    sav_agg = df_diaria.groupby('Matrícula').agg(
        Loja_SAV=('Loja', 'first'),
        Nome_SAV=('Funcionário', 'first'),
        Dias_SAV=('Valor Detalhe', 'count'),
        Valor_SAV=('Valor Detalhe', 'sum')
    ).reset_index()
    
    sav_agg['Alerta_Duplicidade'] = sav_agg['Matrícula'].apply(
        lambda m: 'DIÁRIA DUPLICADA NO SAV' if m in mats_dups else 'NENHUMA'
    )
    sav_agg['Detalhe_Duplicidade'] = sav_agg['Matrícula'].map(dup_details).fillna('')
    
    return sav_agg, df_diaria

def to_centesimal(val):
    """Converte formato de hora HH:MM ou string para horas centesimais (float)."""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip()
    if ':' in val_str:
        parts = val_str.split(':')
        h = float(parts[0])
        m = float(parts[1]) if len(parts) > 1 else 0.0
        return round(h + (m / 60.0), 6)
    try:
        return float(val_str.replace(',', '.'))
    except:
        return 0.0

def carregar_ponto(caminho_ponto):
    """Ingere e totaliza o espelho analítico de ponto (Evento 20650)."""
    print("-> Carregando e processando Ponto Analítico (Evento 20650)...")
    df_raw = pd.read_excel(caminho_ponto)
    # Filtrar apenas CodigoEvento 20650 (Qtd Horas Trab inter)
    df_inter = df_raw[df_raw['CodigoEvento'] == 20650].copy()
    df_inter['Centesimal'] = df_inter['QtdDeHoras'].apply(to_centesimal)
    
    ponto_agg = df_inter.groupby('Matricula').agg(
        Loja_Ponto=('CodigoEstabelecimento', 'first'),
        Estabelecimento_Ponto=('Estabelecimento', 'first'),
        Nome_Ponto=('NomeDoColaborador', 'first'),
        Cargo_Ponto=('Cargo', 'first'),
        Dias_Ponto=('Data', 'nunique'),
        Centesimal_Total=('Centesimal', 'sum')
    ).reset_index()
    
    # Mapa de códigos para nomes de estabelecimento
    lojas_nomes = df_raw[['CodigoEstabelecimento', 'Estabelecimento']].drop_duplicates()
    lojas_map = dict(zip(lojas_nomes['CodigoEstabelecimento'], lojas_nomes['Estabelecimento']))
    
    return ponto_agg, df_inter, lojas_map

def clean_valor_br(val):
    """Converte valor monetário em formato brasileiro string para float nativo."""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace('.', '').replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0.0

def carregar_folha(caminho_fopa):
    """Ingere eventos da folha de pagamento e realiza pivotagem dos eventos-chave."""
    print("-> Carregando e processando Folha de Pagamento (FOPA)...")
    df_raw = pd.read_csv(caminho_fopa, sep=';', encoding='latin1')
    df_raw['Valor_Num'] = df_raw['Valor'].apply(clean_valor_br)
    df_raw['Matricula'] = pd.to_numeric(df_raw['Matrícula Colaborador'], errors='coerce')
    
    eventos_interesse = [900, 7921, 11220, 20650, 20651, 20653]
    df_inter = df_raw[df_raw['Código Evento'].isin(eventos_interesse)].copy()
    
    fopa_pivot = df_inter.pivot_table(
        index='Matricula',
        columns='Código Evento',
        values='Valor_Num',
        aggfunc='sum',
        fill_value=0.0
    ).reset_index()
    
    fopa_pivot.columns = [f"Ev_{c}" if c != 'Matricula' else 'Matricula' for c in fopa_pivot.columns]
    
    # Metadados cadastrais da FOPA
    fopa_meta = df_raw.groupby('Matricula').agg(
        Loja_FOPA=('Código Estabelecimento', 'first'),
        Nome_FOPA=('Nome Colaborador', 'first'),
        Data_Admissao=('Data Admissão', 'first'),
        Tipo_Situacao=('Tipo Situação', 'first')
    ).reset_index()
    
    fopa_agg = pd.merge(fopa_meta, fopa_pivot, on='Matricula', how='left').fillna(0.0)
    
    # Identificar matrículas com eventos intermitentes ativos (> 0)
    eventos_operacionais = [900, 7921, 20650, 20651, 20653]
    fopa_ativos = df_raw[(df_raw['Código Evento'].isin(eventos_operacionais)) & (df_raw['Valor_Num'] > 0)]
    mats_fopa_ativas = set(fopa_ativos['Matricula'].unique())
    
    return fopa_agg, mats_fopa_ativas, df_raw

def carregar_perfil(caminho_perfil):
    """Carrega cadastro de escala e perfil de frequência do sistema LG."""
    print("-> Carregando Perfil de Frequência e Escala...")
    df = pd.read_csv(caminho_perfil, sep=';', encoding='latin1')
    df.columns = [c.strip().replace('\ufeff', '') for c in df.columns]
    df['MATRICULA'] = pd.to_numeric(df['MATRICULA'], errors='coerce')
    df = df.drop_duplicates(subset=['MATRICULA']).copy()
    return df

# ======================================================================================
# 3. CONSOLIDAÇÃO DA AUDITORIA E REGRAS DE NEGÓCIO
# ======================================================================================

def processar_auditoria(df_valor_sav, sav_agg, ponto_agg, fopa_agg, mats_fopa_ativas, df_perfil, lojas_map):
    """Executa o cruzamento completo (Outer Join) e aplica todas as fórmulas e regras."""
    print("-> Executando auditoria cruzada e aplicando regras de negócio...")
    
    mats_sav = set(sav_agg['Matrícula'])
    mats_ponto = set(ponto_agg['Matricula'])
    
    # Universo da auditoria: união externa de todas as matrículas ativas
    universo_mats = sorted(list(mats_sav.union(mats_ponto).union(mats_fopa_ativas)))
    print(f"   [Universo da Auditoria] Total de colaboradores ativos no fechamento: {len(universo_mats)}")
    
    registros = []
    
    for mat in universo_mats:
        s_row = sav_agg[sav_agg['Matrícula'] == mat]
        p_row = ponto_agg[ponto_agg['Matricula'] == mat]
        f_row = fopa_agg[fopa_agg['Matricula'] == mat]
        perf_row = df_perfil[df_perfil['MATRICULA'] == mat]
        
        # Identificação de Loja
        loja = (int(s_row['Loja_SAV'].iloc[0]) if len(s_row) > 0 
                else (int(p_row['Loja_Ponto'].iloc[0]) if len(p_row) > 0 
                      else (int(f_row['Loja_FOPA'].iloc[0]) if len(f_row) > 0 
                            else int(perf_row['ESTABELECIMENTO'].iloc[0]))))
        
        # Nome do Colaborador
        nome = (str(s_row['Nome_SAV'].iloc[0]) if len(s_row) > 0 
                else (str(p_row['Nome_Ponto'].iloc[0]) if len(p_row) > 0 
                      else (str(f_row['Nome_FOPA'].iloc[0]) if len(f_row) > 0 
                            else str(perf_row['NOME'].iloc[0]))))
        
        # Cargo
        cargo = str(p_row['Cargo_Ponto'].iloc[0]) if len(p_row) > 0 else "OPERADOR DE LOJA"
        
        # Nome Estabelecimento / Bandeira
        nome_estab = lojas_map.get(loja, f"LOJA {loja}")
        
        # Parâmetros SAV da Loja
        v_row = df_valor_sav[df_valor_sav['Loja'] == loja]
        jornada_padrao = float(v_row['Jornada Praticada Em Horas'].iloc[0]) if len(v_row) > 0 else 6.0
        val_sav_dia = float(v_row['Valor SAV dia'].iloc[0]) if len(v_row) > 0 else 0.0
        val_sav_hora = float(v_row['Valor SAV Hora'].iloc[0]) if len(v_row) > 0 else 0.0
        
        # SAV
        dias_sav = int(s_row['Dias_SAV'].iloc[0]) if len(s_row) > 0 else 0
        valor_sav = float(s_row['Valor_SAV'].iloc[0]) if len(s_row) > 0 else 0.0
        alerta_dup = str(s_row['Alerta_Duplicidade'].iloc[0]) if len(s_row) > 0 else 'NENHUMA'
        detalhe_dup = str(s_row['Detalhe_Duplicidade'].iloc[0]) if len(s_row) > 0 else ''
        
        # Ponto Analítico
        dias_ponto = int(p_row['Dias_Ponto'].iloc[0]) if len(p_row) > 0 else 0
        centesimal = float(p_row['Centesimal_Total'].iloc[0]) if len(p_row) > 0 else 0.0
        
        # Folha de Pagamento (FOPA)
        ev_900 = float(f_row['Ev_900'].iloc[0]) if len(f_row) > 0 and 'Ev_900' in f_row.columns else 0.0
        ev_7921 = float(f_row['Ev_7921'].iloc[0]) if len(f_row) > 0 and 'Ev_7921' in f_row.columns else 0.0
        ev_11220 = float(f_row['Ev_11220'].iloc[0]) if len(f_row) > 0 and 'Ev_11220' in f_row.columns else 0.0
        ev_20650 = float(f_row['Ev_20650'].iloc[0]) if len(f_row) > 0 and 'Ev_20650' in f_row.columns else 0.0
        ev_20651 = float(f_row['Ev_20651'].iloc[0]) if len(f_row) > 0 and 'Ev_20651' in f_row.columns else 0.0
        ev_20653 = float(f_row['Ev_20653'].iloc[0]) if len(f_row) > 0 and 'Ev_20653' in f_row.columns else 0.0
        
        # Perfil de Frequência LG
        perfil_freq = str(perf_row['PERFIL DE FRQUENCIA'].iloc[0]) if len(perf_row) > 0 else 'NÃO LOCALIZADO'
        
        # Cálculos Matemáticos Oficiais
        # 1. Jornada Prevista = Dias SAV * Carga Padrão da Loja
        jornada_prevista = dias_sav * jornada_padrao
        # 2. Diferença de Horas = Horas Centesimais Realizadas no Ponto - Jornada Prevista
        dif_horas = centesimal - jornada_prevista
        # 3. Valor Horas Trabalhadas = Centesimal * Valor SAV Hora
        val_horas_trab = centesimal * val_sav_hora
        # 4. Valor Carga Prevista = Jornada Prevista * Valor SAV Hora
        val_carga_prev = jornada_prevista * val_sav_hora
        # 5. Estouro de Folha = Valor Carga Prevista - Valor Horas Trabalhadas
        estouro_folha = val_carga_prev - val_horas_trab
        
        # Validações de Status Unitárias
        # Status Horas (Ponto x Folha)
        if centesimal == 0.0 and ev_20650 == 0.0:
            status_horas = 'CORRETO'
        elif round(round(centesimal, 2), 0) == round(round(ev_20650, 2), 0):
            status_horas = 'CORRETO'
        else:
            status_horas = 'DIVERGENTE'
            
        # Status Dias (SAV x Ponto)
        status_dias = 'CONFORME' if dias_sav == dias_ponto else 'DIVERGENTE'
        
        # Status Adiantamento (SAV x Folha 7921)
        status_adiantamento = 'CONFORME' if round(valor_sav, 2) == round(ev_7921, 2) else 'DIVERGENTE'
        
        # Status Evento 900 (Remuneração)
        if (dias_sav > 0 or dias_ponto > 0 or ev_20650 > 0):
            status_ev_900 = 'CONFORME' if ev_900 > 0 else 'SEM EVENTO 900'
        else:
            status_ev_900 = 'CONFORME' if ev_900 > 0 else 'NÃO CONVOCADO'
            
        # Status Perfil LG
        status_perfil = 'CONFORME' if 'HORISTAS INTERMITENTE' in perfil_freq.upper() else 'PERFIL INCORRETO'
        
        # Avaliação de Severidade Global (Status Geral) e Parecer
        motivos_criticos = []
        motivos_alerta = []
        
        # Checagens Críticas
        if status_adiantamento != 'CONFORME':
            motivos_criticos.append(f"Divergência Adiantamento SAV vs Ev.7921 (Dif: R$ {abs(valor_sav - ev_7921):.2f})")
        if status_ev_900 == 'SEM EVENTO 900':
            motivos_criticos.append("Ausência do Evento 900 com trabalho prestado")
        if (dias_sav == 0 or valor_sav == 0) and (ev_900 > 0 or ev_20650 > 0):
            motivos_criticos.append("Pagamento/Horas em Folha sem convocação registrada no SAV ou Ponto")
        if status_perfil != 'CONFORME':
            motivos_criticos.append(f"Perfil LG indevido ({perfil_freq})")
        if status_horas != 'CORRETO':
            motivos_criticos.append(f"Divergência Horas Ponto ({centesimal:.2f}h) x Folha Ev.20650 ({ev_20650:.2f}h)")
            
        # Checagens Operacionais
        if status_dias != 'CONFORME':
            motivos_alerta.append(f"Divergência Dias (SAV: {dias_sav} vs Ponto: {dias_ponto})")
        if abs(estouro_folha) > 0.05:
            motivos_alerta.append(f"Diferença Jornada vs Horas Reais (Estouro: R$ {estouro_folha:.2f})")
        if alerta_dup != 'NENHUMA':
            motivos_alerta.append(detalhe_dup if detalhe_dup else "Diária em duplicidade no SAV")
            
        if len(motivos_criticos) > 0:
            status_geral = 'DIVERGÊNCIA CRÍTICA'
            parecer = " | ".join(motivos_criticos + motivos_alerta)
        elif len(motivos_alerta) > 0:
            status_geral = 'ALERTA OPERACIONAL'
            parecer = " | ".join(motivos_alerta)
        else:
            status_geral = 'CONFORME'
            parecer = "Colaborador regular em todas as validações (SAV, Ponto, Folha e Perfil)."
            
        registros.append({
            'Loja': loja,
            'Estabelecimento': nome_estab,
            'Matrícula': mat,
            'Colaborador': nome,
            'Cargo': cargo,
            'Perfil_Frequencia': perfil_freq,
            # Controle de Dias e Horas
            'Dias_SAV': dias_sav,
            'Dias_Ponto': dias_ponto,
            'Carga_Padrao': jornada_padrao,
            'Jornada_Prevista': jornada_prevista,
            'Horas_Ponto': centesimal,
            'Dif_Horas': dif_horas,
            # Parâmetros Financeiros & Estouro
            'Valor_SAV_Dia': val_sav_dia,
            'Valor_SAV_Hora': val_sav_hora,
            'Valor_Carga_Prev': val_carga_prev,
            'Valor_Horas_Trab': val_horas_trab,
            'Estouro_Folha': estouro_folha,
            # Folha de Pagamento
            'Valor_SAV_Total': valor_sav,
            'Ev_7921': ev_7921,
            'Ev_900': ev_900,
            'Ev_11220': ev_11220,
            'Ev_20650': ev_20650,
            'Status_Consolidado': 'SIM' if ev_20650 > 0 else 'NÃO',
            'Ev_20651': ev_20651,
            'Ev_20653': ev_20653,
            # Validações de Auditoria
            'Status_Horas': status_horas,
            'Status_Dias': status_dias,
            'Status_Adiantamento': status_adiantamento,
            'Status_Ev_900': status_ev_900,
            'Status_Perfil': status_perfil,
            'Alerta_Duplicidade': alerta_dup,
            # Conclusão
            'Status_Geral': status_geral,
            'Parecer_Auditoria': parecer
        })
        
    df_audit = pd.DataFrame(registros)
    
    # Ordenar por Colaborador de A a Z
    df_audit = df_audit.sort_values(by=['Colaborador'], ascending=True).reset_index(drop=True)
    
    # Consolidação por Loja
    df_loja = df_audit.groupby(['Loja', 'Estabelecimento']).agg(
        Total_Colabs=('Matrícula', 'count'),
        Conformes=('Status_Geral', lambda x: (x == 'CONFORME').sum()),
        Alertas=('Status_Geral', lambda x: (x == 'ALERTA OPERACIONAL').sum()),
        Criticos=('Status_Geral', lambda x: (x == 'DIVERGÊNCIA CRÍTICA').sum()),
        Jornada_Prevista_Tot=('Jornada_Prevista', 'sum'),
        Horas_Ponto_Tot=('Horas_Ponto', 'sum'),
        Dif_Horas_Tot=('Dif_Horas', 'sum'),
        Total_SAV=('Valor_SAV_Total', 'sum'),
        Total_7921=('Ev_7921', 'sum'),
        Total_900=('Ev_900', 'sum'),
        Estouro_Liq=('Estouro_Folha', 'sum')
    ).reset_index()
    
    def define_status_loja(r):
        if r['Criticos'] > 0:
            return 'DIVERGÊNCIA CRÍTICA'
        if r['Alertas'] > 0:
            return 'ALERTA OPERACIONAL'
        return 'CONFORME'
        
    df_loja['Status_Loja'] = df_loja.apply(define_status_loja, axis=1)
    df_loja = df_loja.sort_values(by='Loja').reset_index(drop=True)
    
    # Totais consolidados globais
    totais = {
        'total_auditado': len(df_audit),
        'conformes': (df_audit['Status_Geral'] == 'CONFORME').sum(),
        'alertas': (df_audit['Status_Geral'] == 'ALERTA OPERACIONAL').sum(),
        'criticos': (df_audit['Status_Geral'] == 'DIVERGÊNCIA CRÍTICA').sum(),
        'taxa_conformidade': ((df_audit['Status_Geral'] == 'CONFORME').sum() / len(df_audit)) * 100,
        'total_sav': df_audit['Valor_SAV_Total'].sum(),
        'total_7921': df_audit['Ev_7921'].sum(),
        'total_900': df_audit['Ev_900'].sum(),
        'total_estouro': df_audit['Estouro_Folha'].sum(),
        'total_prevista_h': df_audit['Jornada_Prevista'].sum(),
        'total_ponto_h': df_audit['Horas_Ponto'].sum(),
        'total_dif_h': df_audit['Dif_Horas'].sum()
    }
    
    return df_audit, df_loja, totais

# ======================================================================================
# 4. FORMATAÇÃO E DESIGN EXECUTIVO CENCOSUD (OPENPYXL)
# ======================================================================================

# Paleta Corporativa Cencosud
C_AZUL_MARINHO = "1B365D"
C_AZUL_MEDIO   = "2C4A7A"
C_AZUL_CLARO   = "EBF3FA"
C_CINZA_CAB    = "F1F5F9"
C_CINZA_ZEBRA  = "F8FAFC"
C_BRANCO       = "FFFFFF"
C_BORDA        = "CBD5E1"

# Badges
C_VERDE_BG     = "DCFCE7"
C_VERDE_TXT    = "166534"
C_VERM_BG      = "FEE2E2"
C_VERM_TXT     = "991B1B"
C_AMAR_BG      = "FEF3C7"
C_AMAR_TXT     = "92400E"

# Estilos reutilizáveis
font_base = Font(name="Calibri", size=10, color="000000")
font_bold = Font(name="Calibri", size=10, bold=True, color="000000")
font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
font_super_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

border_thin = Border(
    left=Side(style='thin', color=C_BORDA),
    right=Side(style='thin', color=C_BORDA),
    top=Side(style='thin', color=C_BORDA),
    bottom=Side(style='thin', color=C_BORDA)
)
border_total = Border(
    left=Side(style='thin', color=C_BORDA),
    right=Side(style='thin', color=C_BORDA),
    top=Side(style='thin', color="000000"),
    bottom=Side(style='double', color="000000")
)

align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")

def aplicar_badge(cell, valor):
    """Aplica cores suaves e tipografia em badges de status."""
    v_upper = str(valor).upper().strip()
    if v_upper in ['CONFORME', 'CORRETO', 'NENHUMA', 'SIM', 'CONSOLIDADO']:
        cell.fill = PatternFill(start_color=C_VERDE_BG, end_color=C_VERDE_BG, fill_type="solid")
        cell.font = Font(name="Calibri", size=10, bold=True, color=C_VERDE_TXT)
    elif any(x in v_upper for x in ['DIVERGÊNCIA CRÍTICA', 'DIVERGENTE', 'SEM EVENTO 900', 'PERFIL INCORRETO', 'NÃO', 'PENDENTE']):
        cell.fill = PatternFill(start_color=C_VERM_BG, end_color=C_VERM_BG, fill_type="solid")
        cell.font = Font(name="Calibri", size=10, bold=True, color=C_VERM_TXT)
    elif any(x in v_upper for x in ['ALERTA OPERACIONAL', 'DUPLICADA']):
        cell.fill = PatternFill(start_color=C_AMAR_BG, end_color=C_AMAR_BG, fill_type="solid")
        cell.font = Font(name="Calibri", size=10, bold=True, color=C_AMAR_TXT)

def auto_ajuste_colunas(ws, max_limit=55):
    """Ajusta a largura das colunas dinamicamente com margem de segurança."""
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            # Ignora células mescladas na linha de título para não distorcer a largura
            if cell.row in [1, 2, 3] and ws.title == 'Orientações':
                continue
            val_str = str(cell.value or '')
            if '\n' in val_str:
                val_str = max(val_str.split('\n'), key=len)
            max_len = max(max_len, len(val_str))
        largura = min(max(max_len + 3, 12), max_limit)
        ws.column_dimensions[col_letter].width = largura

# ======================================================================================
# 5. CONSTRUÇÃO DA ABA 1: ORIENTAÇÕES
# ======================================================================================

def construir_aba_orientacoes(ws):
    """Gera a documentação executiva institucional, glossário e matriz de regras."""
    ws.title = "Orientações"
    ws.views.sheetView[0].showGridLines = True
    
    # Banner Principal
    ws.merge_cells('B2:K2')
    b1 = ws['B2']
    b1.value = "CENCOSUD BRASIL - DIRETORIA DE RECURSOS HUMANOS"
    b1.font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    b1.fill = PatternFill(start_color=C_AZUL_MARINHO, end_color=C_AZUL_MARINHO, fill_type="solid")
    b1.alignment = align_center
    ws.row_dimensions[2].height = 35
    
    ws.merge_cells('B3:K3')
    b2 = ws['B3']
    b2.value = "SISTEMA INTEGRADO DE AUDITORIA DE FOLHA E PONTO | CONFERÊNCIA DE INTERMITENTES (COMPETÊNCIA 09/2026)"
    b2.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    b2.fill = PatternFill(start_color=C_AZUL_MEDIO, end_color=C_AZUL_MEDIO, fill_type="solid")
    b2.alignment = align_center
    ws.row_dimensions[3].height = 24
    
    curr_row = 5
    
    def add_section_header(row, title):
        ws.merge_cells(f'B{row}:K{row}')
        cell = ws[f'B{row}']
        cell.value = title
        cell.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=C_AZUL_MARINHO, end_color=C_AZUL_MARINHO, fill_type="solid")
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 25
        return row + 1

    # 1. OBJETIVO & FUNDAMENTAÇÃO LEGAL
    curr_row = add_section_header(curr_row, "1. OBJETIVO DO RELATÓRIO E EMBASAMENTO LEGAL")
    
    textos_obj = [
        "• O contrato de trabalho intermitente é regido pelo art. 452-A da CLT, caracterizado pela subordinação não contínua, com alternância de períodos de prestação de serviços e inatividade.",
        "• Na Cencosud Brasil, a convocação formal é operacionalizada pelo adiantamento de diárias via SAV (Sistema de Adiantamento de Viagens/Despesas), garantindo remuneração e transporte imediatos.",
        "• As horas trabalhadas são obrigatoriamente registradas no Ponto Eletrônico (LG) e totalizadas no Evento 20650.",
        "• No fechamento mensal da Folha de Pagamento (FOPA), as diárias adiantadas via SAV são compensadas no Evento 7921, e a contraprestação integral das horas trabalhadas é paga no Evento 900.",
        "• Esta conferência automatizada audita 100% dos colaboradores intermitentes do fechamento, garantindo acurácia matemática, conformidade sindical e zero passivo trabalhista."
    ]
    for t in textos_obj:
        ws.merge_cells(f'B{curr_row}:K{curr_row}')
        c = ws[f'B{curr_row}']
        c.value = t
        c.font = font_base
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[curr_row].height = 20
        curr_row += 1
        
    curr_row += 1
    
    # 2. GLOSSÁRIO DE EVENTOS DA FOLHA
    curr_row = add_section_header(curr_row, "2. GLOSSÁRIO OFICIAL DOS EVENTOS DA FOLHA DE PAGAMENTO (FOPA)")
    
    headers_glo = ["Código", "Nome do Evento", "Tipo / Função", "Regra de Negócio e Validação de Auditoria"]
    cols_glo = [('B', 'C'), ('D', 'E'), ('F', 'G'), ('H', 'K')]
    
    for idx, (c1, c2) in enumerate(cols_glo):
        ws.merge_cells(f'{c1}{curr_row}:{c2}{curr_row}')
        cell = ws[f'{c1}{curr_row}']
        cell.value = headers_glo[idx]
        cell.font = font_header
        cell.fill = PatternFill(start_color=C_AZUL_MEDIO, end_color=C_AZUL_MEDIO, fill_type="solid")
        cell.alignment = align_center
    ws.row_dimensions[curr_row].height = 22
    curr_row += 1
    
    eventos_data = [
        ("900", "Remuneração Horas Trabalhadas Intermitente", "Provento / Salário", "Obrigatório para todo colaborador que prestou serviço. Quita as horas reais apuradas."),
        ("7921", "Valor Diário Adiant Sal Intermitente", "Desconto / Compensação", "Desconto em folha que compensa centavo a centavo o adiantamento realizado pelo SAV."),
        ("11220", "Remuneração Fixa Mensal", "Informativo Cadastral", "Salário/hora cadastral acordado no contrato de trabalho do colaborador intermitente."),
        ("20650", "Qtd Horas Trabalhadas Intermitente", "Referência / Horas", "Total de horas apuradas no ponto eletrônico que sobe para cálculo da folha. Deve bater com o centesimal."),
        ("20651", "Num Dias Trabalhadas Intermitente", "Referência / Dias", "Quantidade de dias de convocação com trabalho efetivo prestado no mês de referência."),
        ("20653", "Carga de Horas Intermitente", "Informativo / Carga", "Carga horária teórica máxima ou contratada de referência no cadastro da empresa.")
    ]
    
    for row_data in eventos_data:
        for idx, (c1, c2) in enumerate(cols_glo):
            ws.merge_cells(f'{c1}{curr_row}:{c2}{curr_row}')
            cell = ws[f'{c1}{curr_row}']
            cell.value = row_data[idx]
            cell.font = font_bold if idx == 0 else font_base
            cell.alignment = align_center if idx in [0, 2] else align_left
            cell.border = border_thin
        ws.row_dimensions[curr_row].height = 20
        curr_row += 1
        
    curr_row += 1
    
    # 3. MATRIZ DE CRITÉRIOS DE AUDITORIA E STATUS
    curr_row = add_section_header(curr_row, "3. MATRIZ DE CRITÉRIOS DE AUDITORIA E SEVERIDADE DE STATUS")
    
    headers_mat = ["Status Global", "Severidade", "Critério de Enquadramento Técnico", "Ação Operacional Exigida da Equipe de RH"]
    cols_mat = [('B', 'C'), ('D', 'D'), ('E', 'G'), ('H', 'K')]
    
    for idx, (c1, c2) in enumerate(cols_mat):
        ws.merge_cells(f'{c1}{curr_row}:{c2}{curr_row}')
        cell = ws[f'{c1}{curr_row}']
        cell.value = headers_mat[idx]
        cell.font = font_header
        cell.fill = PatternFill(start_color=C_AZUL_MEDIO, end_color=C_AZUL_MEDIO, fill_type="solid")
        cell.alignment = align_center
    ws.row_dimensions[curr_row].height = 22
    curr_row += 1
    
    matriz_data = [
        ("CONFORME", "Verde / Regular", "Todas as validações (SAV, Ponto, Folha e Perfil) 100% aderentes, sem divergência ou duplicidade.", "Nenhuma ação requerida. Processo liberado para homologação e fechamento."),
        ("ALERTA OPERACIONAL", "Amarelo / Atenção", "Diferenças entre horas teóricas convocadas e horas reais (estouro de jornada) ou diária duplicada no SAV.", "Verificar espelho de ponto da loja, solicitar justificativa operacional e ajustar SAV se aplicável."),
        ("DIVERGÊNCIA CRÍTICA", "Vermelho / Imediato", "Adiantamento SAV diverge do Ev. 7921, falta de Ev. 900, pagamento sem SAV ou perfil LG regular indevido.", "BLOQUEIO DE FECHAMENTO: Retificar imediatamente cadastro no LG, lançar evento ou ajustar diária.")
    ]
    
    for row_data in matriz_data:
        for idx, (c1, c2) in enumerate(cols_mat):
            ws.merge_cells(f'{c1}{curr_row}:{c2}{curr_row}')
            cell = ws[f'{c1}{curr_row}']
            cell.value = row_data[idx]
            cell.alignment = align_center if idx in [0, 1] else align_left
            cell.border = border_thin
            if idx == 0:
                aplicar_badge(cell, row_data[0])
            else:
                cell.font = font_base
        ws.row_dimensions[curr_row].height = 24
        curr_row += 1
        
    curr_row += 1
    
    # 4. METADADOS E CONTROLE DE QUALIDADE
    ws.merge_cells(f'B{curr_row}:K{curr_row}')
    ws[f'B{curr_row}'].value = "Relatório processado via Automação de Dados de RH Cencosud | Base Histórica Calibrada: 08/2026 | Competência Ativa: 09/2026."
    ws[f'B{curr_row}'].font = Font(name="Calibri", size=9, italic=True, color="64748B")
    ws[f'B{curr_row}'].alignment = align_left

# ======================================================================================
# 6. CONSTRUÇÃO DA ABA 2: RESUMO DA AUDITORIA (DASHBOARD & KPIs)
# ======================================================================================

def construir_aba_resumo(ws, df_loja, totais):
    """Gera dashboard executivo com cards estilizados e consolidação por loja."""
    ws.title = "Resumo da Auditoria"
    ws.views.sheetView[0].showGridLines = True
    
    # Banner
    ws.merge_cells('B2:M2')
    b = ws['B2']
    b.value = "CENCOSUD BRASIL - DASHBOARD EXECUTIVO DA CONFERÊNCIA DE INTERMITENTES (09/2026)"
    b.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    b.fill = PatternFill(start_color=C_AZUL_MARINHO, end_color=C_AZUL_MARINHO, fill_type="solid")
    b.alignment = align_center
    ws.row_dimensions[2].height = 32
    
    # Cards de KPIs
    # Linha 4 (Cabeçalho do Card), Linha 5 (Valor do Card), Linha 6 (Subtítulo)
    cards = [
        ("TOTAL AUDITADO", f"{totais['total_auditado']}", "Colaboradores Ativos", C_AZUL_MARINHO, "B", "C"),
        ("EM CONFORMIDADE", f"{totais['conformes']}", f"Taxa: {totais['taxa_conformidade']:.1f}%", C_VERDE_TXT, "D", "E"),
        ("ALERTAS OPERACIONAIS", f"{totais['alertas']}", "Ajustes de Jornada/SAV", C_AMAR_TXT, "F", "G"),
        ("DIVERGÊNCIAS CRÍTICAS", f"{totais['criticos']}", "Ação Imediata RH", C_VERM_TXT, "H", "I"),
        ("TOTAL SAV (ADIANTADO)", f"R$ {totais['total_sav']:,.2f}", "Solicitações no SAV", C_AZUL_MARINHO, "J", "K"),
        ("TOTAL FOLHA (EV. 7921)", f"R$ {totais['total_7921']:,.2f}", f"Dif: R$ {abs(totais['total_sav'] - totais['total_7921']):.2f}", C_AZUL_MARINHO, "L", "M")
    ]
    
    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 28
    ws.row_dimensions[6].height = 16
    
    for title, val_str, sub_str, cor_txt, c1, c2 in cards:
        # Cabeçalho do Card
        ws.merge_cells(f'{c1}4:{c2}4')
        cell_t = ws[f'{c1}4']
        cell_t.value = title
        cell_t.font = Font(name="Calibri", size=9, bold=True, color="64748B")
        cell_t.alignment = align_center
        cell_t.fill = PatternFill(start_color=C_CINZA_CAB, end_color=C_CINZA_CAB, fill_type="solid")
        cell_t.border = Border(top=Side(style='thin', color=C_BORDA), left=Side(style='thin', color=C_BORDA), right=Side(style='thin', color=C_BORDA))
        
        # Valor do Card
        ws.merge_cells(f'{c1}5:{c2}5')
        cell_v = ws[f'{c1}5']
        cell_v.value = val_str
        cell_v.font = Font(name="Calibri", size=14, bold=True, color=cor_txt)
        cell_v.alignment = align_center
        cell_v.fill = PatternFill(start_color=C_BRANCO, end_color=C_BRANCO, fill_type="solid")
        cell_v.border = Border(left=Side(style='thin', color=C_BORDA), right=Side(style='thin', color=C_BORDA))
        
        # Subtítulo do Card
        ws.merge_cells(f'{c1}6:{c2}6')
        cell_s = ws[f'{c1}6']
        cell_s.value = sub_str
        cell_s.font = Font(name="Calibri", size=8, italic=True, color="64748B")
        cell_s.alignment = align_center
        cell_s.fill = PatternFill(start_color=C_CINZA_ZEBRA, end_color=C_CINZA_ZEBRA, fill_type="solid")
        cell_s.border = Border(bottom=Side(style='thin', color=C_BORDA), left=Side(style='thin', color=C_BORDA), right=Side(style='thin', color=C_BORDA))
        
    # Título da Tabela Consolidada
    ws.merge_cells('B8:M8')
    t_loja = ws['B8']
    t_loja.value = "CONSOLIDAÇÃO POR LOJA / ESTABELECIMENTO"
    t_loja.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    t_loja.fill = PatternFill(start_color=C_AZUL_MEDIO, end_color=C_AZUL_MEDIO, fill_type="solid")
    t_loja.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[8].height = 24
    
    headers_loja = [
        "Loja", "Estabelecimento / Filial", "Colabs", "Conformes", "Alertas", "Críticos",
        "Jornada Prev. (h)", "Horas Ponto (h)", "Total SAV (R$)", "Folha 7921 (R$)",
        "Folha 900 (R$)", "Estouro Líq. (R$)", "Status Loja"
    ]
    cols_loja = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
    
    # Cabeçalho da Tabela
    ws.row_dimensions[9].height = 22
    for idx, h_text in enumerate(headers_loja):
        col_letter = get_column_letter(idx + 2)
        cell = ws[f'{col_letter}9']
        cell.value = h_text
        cell.font = font_header
        cell.fill = PatternFill(start_color=C_AZUL_MARINHO, end_color=C_AZUL_MARINHO, fill_type="solid")
        cell.alignment = align_center
        cell.border = border_thin
        
    curr_row = 10
    for _, row in df_loja.iterrows():
        ws.row_dimensions[curr_row].height = 20
        bg_color = C_CINZA_ZEBRA if curr_row % 2 == 1 else C_BRANCO
        fill_row = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        
        c_loja = ws[f'B{curr_row}']
        c_loja.value = int(row['Loja'])
        c_loja.alignment = align_center
        
        c_estab = ws[f'C{curr_row}']
        c_estab.value = str(row['Estabelecimento'])
        c_estab.alignment = align_left
        
        c_tot = ws[f'D{curr_row}']
        c_tot.value = int(row['Total_Colabs'])
        c_tot.alignment = align_center
        c_tot.number_format = '#,##0'
        
        c_conf = ws[f'E{curr_row}']
        c_conf.value = int(row['Conformes'])
        c_conf.alignment = align_center
        c_conf.number_format = '#,##0'
        
        c_alert = ws[f'F{curr_row}']
        c_alert.value = int(row['Alertas'])
        c_alert.alignment = align_center
        c_alert.number_format = '#,##0'
        
        c_crit = ws[f'G{curr_row}']
        c_crit.value = int(row['Criticos'])
        c_crit.alignment = align_center
        c_crit.number_format = '#,##0'
        
        c_prev = ws[f'H{curr_row}']
        c_prev.value = float(row['Jornada_Prevista_Tot'])
        c_prev.alignment = align_center
        c_prev.number_format = '#,##0.00'
        
        c_ponto = ws[f'I{curr_row}']
        c_ponto.value = float(row['Horas_Ponto_Tot'])
        c_ponto.alignment = align_center
        c_ponto.number_format = '#,##0.00'
        
        c_sav = ws[f'J{curr_row}']
        c_sav.value = float(row['Total_SAV'])
        c_sav.alignment = align_center
        c_sav.number_format = 'R$ #,##0.00'
        
        c_7921 = ws[f'K{curr_row}']
        c_7921.value = float(row['Total_7921'])
        c_7921.alignment = align_center
        c_7921.number_format = 'R$ #,##0.00'
        
        c_900 = ws[f'L{curr_row}']
        c_900.value = float(row['Total_900'])
        c_900.alignment = align_center
        c_900.number_format = 'R$ #,##0.00'
        
        c_est = ws[f'M{curr_row}']
        c_est.value = float(row['Estouro_Liq'])
        c_est.alignment = align_center
        c_est.number_format = 'R$ #,##0.00'
        
        c_st = ws[f'N{curr_row}']
        c_st.value = str(row['Status_Loja'])
        c_st.alignment = align_left
        aplicar_badge(c_st, row['Status_Loja'])
        
        for idx in range(len(headers_loja)):
            col_letter = get_column_letter(idx + 2)
            c = ws[f'{col_letter}{curr_row}']
            c.border = border_thin
            if col_letter != 'N':
                c.fill = fill_row
                c.font = font_base
                
        curr_row += 1
        
    # Linha de Totais da Tabela de Lojas
    ws.row_dimensions[curr_row].height = 22
    fill_tot = PatternFill(start_color=C_CINZA_CAB, end_color=C_CINZA_CAB, fill_type="solid")
    
    ws[f'B{curr_row}'].value = "TOTAL"
    ws[f'C{curr_row}'].value = "CONSOLIDAÇÃO GERAL (16 LOJAS AUDITADAS)"
    ws[f'D{curr_row}'].value = totais['total_auditado']
    ws[f'E{curr_row}'].value = totais['conformes']
    ws[f'F{curr_row}'].value = totais['alertas']
    ws[f'G{curr_row}'].value = totais['criticos']
    ws[f'H{curr_row}'].value = totais['total_prevista_h']
    ws[f'I{curr_row}'].value = totais['total_ponto_h']
    ws[f'J{curr_row}'].value = totais['total_sav']
    ws[f'K{curr_row}'].value = totais['total_7921']
    ws[f'L{curr_row}'].value = totais['total_900']
    ws[f'M{curr_row}'].value = totais['total_estouro']
    ws[f'N{curr_row}'].value = "DIVERGÊNCIA CRÍTICA" if totais['criticos'] > 0 else "CONFORME"
    aplicar_badge(ws[f'N{curr_row}'], ws[f'N{curr_row}'].value)
    
    for idx in range(len(headers_loja)):
        col_letter = get_column_letter(idx + 2)
        c = ws[f'{col_letter}{curr_row}']
        c.border = border_total
        if col_letter != 'N':
            c.fill = fill_tot
            c.font = font_bold
            c.alignment = align_center if col_letter != 'C' else align_left
            if col_letter in ['D', 'E', 'F', 'G']:
                c.number_format = '#,##0'
            elif col_letter in ['H', 'I']:
                c.number_format = '#,##0.00'
            elif col_letter in ['J', 'K', 'L', 'M']:
                c.number_format = 'R$ #,##0.00'

# ======================================================================================
# 7. CONSTRUÇÃO DAS ABAS ANALÍTICAS (CONFERÊNCIA GERAL E DIVERGÊNCIAS)
# ======================================================================================

SUPER_HEADERS = [
    ("DADOS CADASTRAIS & ENQUADRAMENTO", 1, 5, C_AZUL_MARINHO),
    ("CONTROLE DE JORNADA E PONTO (SAV X PONTO)", 6, 11, C_AZUL_MEDIO),
    ("PARÂMETROS FINANCEIROS & ESTOURO DE FOLHA", 12, 16, "1E3A8A"),
    ("FOLHA DE PAGAMENTO (FOPA)", 17, 24, "1E293B"),
    ("AUDITORIA E VALIDAÇÃO DE REGRAS", 25, 30, "334155"),
    ("DIAGNÓSTICO E PARECER FINAL", 31, 32, "0F172A")
]

COLUNAS_ANALITICAS = [
    # 1 a 5: Cadastrais
    ("Loja", "center", "@"),
    ("Matrícula", "center", "@"),
    ("Colaborador", "left", "@"),
    ("Cargo", "left", "@"),
    ("Perfil de Frequência", "left", "@"),
    # 6 a 11: Jornada e Ponto
    ("Dias SAV", "center", "#,##0"),
    ("Dias Ponto", "center", "#,##0"),
    ("Carga Padrão (h)", "center", "#,##0"),
    ("Jornada Prevista (h)", "center", "#,##0.00"),
    ("Horas Ponto (h)", "center", "#,##0.00"),
    ("Diferença Horas (h)", "center", "#,##0.00"),
    # 12 a 16: Financeiro & Estouro
    ("Valor SAV Dia (R$)", "center", "R$ #,##0.00"),
    ("Valor SAV Hora (R$)", "center", "R$ #,##0.000"),
    ("Valor Carga Prev. (R$)", "center", "R$ #,##0.00"),
    ("Valor Horas Trab. (R$)", "center", "R$ #,##0.00"),
    ("Estouro Folha (R$)", "center", "R$ #,##0.00"),
    # 17 a 24: FOPA
    ("Total SAV (R$)", "center", "R$ #,##0.00"),
    ("Ev. 7921 Adiant. (R$)", "center", "R$ #,##0.00"),
    ("Ev. 900 Remun. (R$)", "center", "R$ #,##0.00"),
    ("Ev. 11220 Sal/h (R$)", "center", "R$ #,##0.00"),
    ("Ev. 20650 Horas (h)", "center", "#,##0.00"),
    ("Consolidado (Ev. 20650)", "center", "@"),
    ("Ev. 20651 Dias", "center", "#,##0"),
    ("Ev. 20653 Carga (h)", "center", "#,##0.00"),
    # 25 a 30: Validações
    ("Status Horas", "left", "@"),
    ("Status Dias", "left", "@"),
    ("Status Adiantamento", "left", "@"),
    ("Status Ev. 900", "left", "@"),
    ("Status Perfil", "left", "@"),
    ("Alerta Duplicidade", "left", "@"),
    # 31 a 32: Conclusão
    ("Status Geral", "left", "@"),
    ("Parecer da Auditoria", "left", "@")
]

def construir_aba_analitica(ws, sheet_name, df_data):
    """Constrói a tabela analítica com duplo cabeçalho, tipagem nativa e badges."""
    ws.title = sheet_name
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = 'A3'
    
    # 1. Super-cabeçalho (Linha 1)
    ws.row_dimensions[1].height = 24
    for title, col_start, col_end, bg_color in SUPER_HEADERS:
        start_letter = get_column_letter(col_start)
        end_letter = get_column_letter(col_end)
        ws.merge_cells(f'{start_letter}1:{end_letter}1')
        cell = ws[f'{start_letter}1']
        cell.value = title
        cell.font = font_super_header
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        cell.alignment = align_center
        
        # Bordas
        for c_idx in range(col_start, col_end + 1):
            c_let = get_column_letter(c_idx)
            ws[f'{c_let}1'].border = border_thin
            
    # 2. Cabeçalho de Colunas (Linha 2)
    ws.row_dimensions[2].height = 24
    for col_idx, (col_name, align_str, _) in enumerate(COLUNAS_ANALITICAS, start=1):
        col_letter = get_column_letter(col_idx)
        cell = ws[f'{col_letter}2']
        cell.value = col_name
        cell.font = font_header
        cell.fill = PatternFill(start_color=C_AZUL_MARINHO, end_color=C_AZUL_MARINHO, fill_type="solid")
        cell.alignment = align_center
        cell.border = border_thin
        
    # 3. Linhas de Dados
    curr_row = 3
    for _, r in df_data.iterrows():
        ws.row_dimensions[curr_row].height = 20
        bg_color = C_CINZA_ZEBRA if curr_row % 2 == 1 else C_BRANCO
        fill_data = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        
        valores_linha = [
            r['Loja'],
            r['Matrícula'],
            r['Colaborador'],
            r['Cargo'],
            r['Perfil_Frequencia'],
            r['Dias_SAV'],
            r['Dias_Ponto'],
            r['Carga_Padrao'],
            r['Jornada_Prevista'],
            r['Horas_Ponto'],
            r['Dif_Horas'],
            r['Valor_SAV_Dia'],
            r['Valor_SAV_Hora'],
            r['Valor_Carga_Prev'],
            r['Valor_Horas_Trab'],
            r['Estouro_Folha'],
            r['Valor_SAV_Total'],
            r['Ev_7921'],
            r['Ev_900'],
            r['Ev_11220'],
            r['Ev_20650'],
            r['Status_Consolidado'],
            r['Ev_20651'],
            r['Ev_20653'],
            r['Status_Horas'],
            r['Status_Dias'],
            r['Status_Adiantamento'],
            r['Status_Ev_900'],
            r['Status_Perfil'],
            r['Alerta_Duplicidade'],
            r['Status_Geral'],
            r['Parecer_Auditoria']
        ]
        
        for col_idx, (col_name, align_str, num_fmt) in enumerate(COLUNAS_ANALITICAS, start=1):
            col_letter = get_column_letter(col_idx)
            cell = ws[f'{col_letter}{curr_row}']
            v = valores_linha[col_idx - 1]
            cell.value = v
            cell.number_format = num_fmt
            cell.border = border_thin
            cell.alignment = align_center if align_str == 'center' else align_left
            
            # Formatação especial de status (coluna 22 e colunas 25 a 31)
            if col_idx == 22 or col_idx in range(25, 32):
                aplicar_badge(cell, v)
            else:
                cell.fill = fill_data
                cell.font = font_base
                
        curr_row += 1
        
    # 4. Linha de Total Geral
    ws.row_dimensions[curr_row].height = 22
    fill_tot = PatternFill(start_color=C_CINZA_CAB, end_color=C_CINZA_CAB, fill_type="solid")
    
    # Preencher totais calculados
    totais_map = {
        1: "TOTAL",
        2: f"{len(df_data)} Colabs",
        3: "TOTALIZAÇÃO GERAL DA AUDITORIA",
        6: df_data['Dias_SAV'].sum(),
        7: df_data['Dias_Ponto'].sum(),
        9: df_data['Jornada_Prevista'].sum(),
        10: df_data['Horas_Ponto'].sum(),
        11: df_data['Dif_Horas'].sum(),
        14: df_data['Valor_Carga_Prev'].sum(),
        15: df_data['Valor_Horas_Trab'].sum(),
        16: df_data['Estouro_Folha'].sum(),
        17: df_data['Valor_SAV_Total'].sum(),
        18: df_data['Ev_7921'].sum(),
        19: df_data['Ev_900'].sum(),
        21: df_data['Ev_20650'].sum(),
        23: df_data['Ev_20651'].sum(),
        24: df_data['Ev_20653'].sum()
    }
    
    for col_idx, (_, align_str, num_fmt) in enumerate(COLUNAS_ANALITICAS, start=1):
        col_letter = get_column_letter(col_idx)
        cell = ws[f'{col_letter}{curr_row}']
        cell.border = border_total
        cell.fill = fill_tot
        cell.font = font_bold
        cell.alignment = align_center if align_str == 'center' else align_left
        
        if col_idx in totais_map:
            cell.value = totais_map[col_idx]
            cell.number_format = num_fmt
        else:
            cell.value = ""

# ======================================================================================
# 8. WORKFLOW PRINCIPAL DE EXECUÇÃO
# ======================================================================================

def main():
    print("=" * 80)
    print("CENCOSUD BRASIL - ENGENHARIA DE DADOS & AUDITORIA DE RH")
    print("AUTOMAÇÃO DA CONFERÊNCIA DE INTERMITENTES (COMPETÊNCIA 09/2026)")
    print("=" * 80)
    
    # 1. Carregar dados brutos
    df_valor_sav = carregar_parametros_sav(HIST_FILE)
    sav_agg, _ = carregar_sav(SAV_FILE)
    ponto_agg, _, lojas_map = carregar_ponto(PONTO_FILE)
    fopa_agg, mats_fopa_ativas, _ = carregar_folha(FOPA_FILE)
    df_perfil = carregar_perfil(PERFIL_FILE)
    
    # 2. Processar auditoria cruzada
    df_audit, df_loja, totais = processar_auditoria(
        df_valor_sav, sav_agg, ponto_agg, fopa_agg, mats_fopa_ativas, df_perfil, lojas_map
    )
    
    # Filtrar divergências para a Aba 4
    df_diverg = df_audit[df_audit['Status_Geral'] != 'CONFORME'].copy()
    
    # 3. Gerar Workbook Excel
    print("-> Criando Workbook Excel com Padrão Visual Executivo Cencosud...")
    wb = openpyxl.Workbook()
    # Remove sheet padrão inicial
    default_sheet = wb.active
    
    # Aba 1: Orientações
    ws_ori = wb.create_sheet(title="Orientações")
    construir_aba_orientacoes(ws_ori)
    
    # Aba 2: Resumo da Auditoria
    ws_res = wb.create_sheet(title="Resumo da Auditoria")
    construir_aba_resumo(ws_res, df_loja, totais)
    
    # Aba 3: Conferência Geral
    ws_conf = wb.create_sheet(title="Conferência Geral")
    construir_aba_analitica(ws_conf, "Conferência Geral", df_audit)
    
    # Aba 4: Divergências
    ws_div = wb.create_sheet(title="Divergências")
    construir_aba_analitica(ws_div, "Divergências", df_diverg)
    
    # Remover aba padrão vazia criada pelo openpyxl
    wb.remove(default_sheet)
    
    # 4. Auto-ajuste de colunas em todas as abas
    print("-> Aplicando auto-ajuste de larguras de coluna e formatação visual...")
    auto_ajuste_colunas(ws_ori, max_limit=80)
    auto_ajuste_colunas(ws_res, max_limit=50)
    auto_ajuste_colunas(ws_conf, max_limit=60)
    auto_ajuste_colunas(ws_div, max_limit=60)
    
    # Salvar arquivo oficial na pasta de conferencias e espelho na raiz
    print(f"-> Gravando arquivo consolidado em:\n   {OUTPUT_FILE}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    wb.save(OUTPUT_FILE)
    wb.save(os.path.join(BASE_DIR, "Relatorio_Conferencia_Intermitentes_09.2026.xlsx"))
    
    print("\n" + "=" * 80)
    print("PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
    print(f"Total Colaboradores Auditados: {totais['total_auditado']}")
    print(f" - Em Conformidade:            {totais['conformes']}")
    print(f" - Alertas Operacionais:       {totais['alertas']}")
    print(f" - Divergências Críticas:      {totais['criticos']}")
    print(f"Total SAV Adiantado:           R$ {totais['total_sav']:,.2f}")
    print(f"Total Folha (Ev. 7921):        R$ {totais['total_7921']:,.2f}")
    print(f"Total Folha (Ev. 900):         R$ {totais['total_900']:,.2f}")
    print(f"Estouro de Folha Líquido:      R$ {totais['total_estouro']:,.2f}")
    print("=" * 80)

if __name__ == "__main__":
    main()
