"""
========================================================================================
CENCOSUD BRASIL - ENGENHARIA DE DADOS & AUDITORIA DE RH
SUPER MÓDULO UNIFICADO DE JUSTIFICATIVAS E OCORRÊNCIAS DE PONTO
COMPETÊNCIA: 09/2026
========================================================================================
Autor: Agente Desenvolvedor (Pair Engineering com Arquiteto de Dados & Validação Crítica)
Data: Outubro/2026

Objetivo:
Processar e auditar 100% das ocorrências de ponto das 7 regionais Cencosud Brasil
(Bretas, GBarbosa, Giga, Mercantil, Perini, Prezunic, Spid), consolidando em uma
única esteira automatizada de alta performance as seguintes conferências:
  - Prioridade 2: Validação da Justificativa da Ocorrência vs. Cartão de Ponto
  - Prioridade 6: Aderência de Justificativas a Códigos Sindicais
  - Prioridade 8: Auditoria de Horas Excessivas de Interjornada (> 2h de supressão)
  - Prioridade 9: Auditoria de Cargos Elegíveis em Sobreaviso (Técnico vs. Operacional)

Gera relatório executivo em Excel no Padrão Cencosud de 4 Abas:
  1. Orientações (Embasamento CLT, Regras de Conciliação e Matriz de Status)
  2. Resumo da Auditoria (Cards Executivos de KPIs e Consolidação por Regional)
  3. Conferência Geral (Base Analítica Completa de Todas as Regionais)
  4. Divergências (Filtro Exclusivo de Divergências Críticas e Alertas)
========================================================================================
"""

import os
import sys
import time
import re
import unicodedata
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ======================================================================================
# 1. MAPEAMENTO DE DIRETÓRIOS E CONSTANTES
# ======================================================================================
BASE_DIR = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\ponto\Justificativas de Ponto"
CARTAO_DIR = os.path.join(BASE_DIR, "Cartão de Ponto")
OCORRENCIAS_DIR = os.path.join(BASE_DIR, "Ocorrências")
OUTPUT_DIR = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\ponto\conferencias geradas pelo script"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Relatorio_Conferencia_Justificativas_Ponto_09.2026.xlsx")
OUTPUT_FILE_ROOT = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\ponto\Relatorio_Conferencia_Justificativas_Ponto_09.2026.xlsx"

REGIONAIS = ["BRETAS", "GBARBOSA", "GIGA", "MERCANTIL", "PERINI", "PREZUNIC", "SPID"]

# Cargos autorizados/técnicos em sobreaviso
CARGOS_TECNICOS_SOBREAVISO = [
    'HELP SUPPORT', 'FIELD SUPPORT', 'ANALISTA', 'ANALYST', 'COORDENADOR', 'COORD', 'GERENTE',
    'ESPECIALISTA', 'SPECIALIST', 'CONSULTOR', 'CONSULTANT', 'DESENVOLVEDOR', 'DEVELOPER',
    'ENGENHEIRO', 'ENGINEER', 'ARQUITETO', 'ARCHITECT',
    'TECNICO TI', 'TECNICO DE TI', 'SUPORTE', 'SUPPORT', 'NETWORK', 'INFRA',
    'ADMINISTRADOR', 'SUPERVISOR TI', 'TECH LEAD', 'ANALISTA DE SUPORTE',
    'ANALISTA DE SISTEMAS', 'SEGURANCA DA INFORMACAO', 'SECURITY'
]

# Paleta Corporativa Cencosud
C_AZUL_MARINHO = "1B365D"
C_AZUL_MEDIO   = "2C4A7A"
C_AZUL_CLARO   = "EBF3FA"
C_CINZA_CAB    = "F1F5F9"
C_CINZA_ZEBRA  = "F8FAFC"
C_BRANCO       = "FFFFFF"
C_BORDA        = "CBD5E1"

# Badges de Status
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

# ======================================================================================
# 2. MOTOR DE NORMALIZAÇÃO E REGRAS DE AUDITORIA
# ======================================================================================

def normalizar_texto(texto):
    """Normaliza texto removendo acentos, prefixos operacionais e pontuação."""
    if not texto:
        return ""
    t = str(texto).upper().strip()
    t = ''.join(c for c in unicodedata.normalize('NFD', t) if unicodedata.category(c) != 'Mn')
    for prefix in ['APROVADO -', 'APROVADO-', 'APROVADO', 'PENDENTE -', 'PENDENTE-', 'PENDENTE', 'ABONO -', 'ABONO-']:
        if t.startswith(prefix):
            t = t[len(prefix):].strip()
    t = re.sub(r'[_]+', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t

def extrair_codigo_sindicato(texto):
    """Extrai código numérico de sindicato a partir de descrição de ocorrência/justificativa/perfil."""
    if not texto:
        return None
    # Padrão SIND 108, SIND. 108, SINDICATO 108
    match = re.search(r'\bSIND(?:ICATO)?\.?\s*(\d+)\b', str(texto).upper())
    if match:
        return match.group(1).lstrip('0')
    # Padrão inicial como '127 SIN EMP...'
    match_inicio = re.match(r'^\s*(\d+)\s+SIN', str(texto).upper())
    if match_inicio:
        return match_inicio.group(1).lstrip('0')
    return None

def auditar_registro(oc_just, just_cartao, ocorrencia, qtd_horas, cargo, perfil_frequencia):
    """Executa auditoria em 4 dimensões e emite parecer e classificação de status."""
    # 1. Dimensão A: Validação da Justificativa (Prioridade 2)
    norm_oc = normalizar_texto(oc_just)
    norm_cp = normalizar_texto(just_cartao)
    
    if not norm_cp or norm_cp == "" or norm_cp == "______" or norm_cp == "NAN" or norm_cp == "NONE":
        dim_a = "PENDENTE"
    elif norm_oc == norm_cp or norm_oc in norm_cp or norm_cp in norm_oc:
        dim_a = "CORRETO"
    else:
        dim_a = "DIVERGENTE"
        
    # 2. Dimensão B: Auditoria de Interjornada (Prioridade 8 - Acima de 07:20)
    texto_inter = (str(ocorrencia) + " " + str(oc_just)).upper()
    h_inter_float = 0.0
    inter_mins = 0
    if "INTERJORNADA" in texto_inter:
        try:
            val_str = str(qtd_horas).strip()
            if ':' in val_str:
                parts = val_str.split(':')
                inter_mins = int(parts[0]) * 60 + int(parts[1])
                h_inter_float = inter_mins / 60.0
            else:
                h_inter_float = float(val_str.replace(',', '.'))
                inter_mins = int(round(h_inter_float * 60))
        except:
            h_inter_float = 0.0
            inter_mins = 0
        dim_b = "ATENÇÃO / CORRIGIR" if inter_mins > 440 else "CONFORME"
    else:
        dim_b = "CONFORME"
        
    # 3. Dimensão C: Auditoria de Sobreaviso (Prioridade 9)
    texto_sobre = (str(ocorrencia) + " " + str(oc_just)).upper()
    if "SOBREAVISO" in texto_sobre:
        c_upper = str(cargo).upper()
        if any(tecnico in c_upper for tecnico in CARGOS_TECNICOS_SOBREAVISO):
            dim_c = "ELEGIVEL"
        else:
            dim_c = "DIVERGENTE"
    else:
        dim_c = ""
        
    # 4. Dimensão D: Auditoria de Sindicato (Prioridade 6)
    sind_just = extrair_codigo_sindicato(oc_just)
    if sind_just:
        if not perfil_frequencia or str(perfil_frequencia).strip() in ['', 'NÃO INFORMADO', 'NONE', 'NAN']:
            dim_d = "PENDENTE (SEM PERFIL LG)"
        else:
            sind_perf = extrair_codigo_sindicato(perfil_frequencia)
            if sind_perf:
                dim_d = "CONFORME" if sind_just == sind_perf else "SINDICATO DIVERGENTE"
            else:
                dim_d = "CONFORME" if re.search(r'\b0*' + sind_just + r'\b', str(perfil_frequencia)) else "SINDICATO DIVERGENTE"
    else:
        dim_d = "CONFORME"
        
    # Consolidação do Status Geral e Parecer Técnico
    criticos = []
    alertas = []
    
    if dim_a == "DIVERGENTE":
        criticos.append(f"Justificativa divergente no espelho (Ocorrência: '{oc_just}' vs Cartão: '{just_cartao}')")
    elif dim_a == "PENDENTE":
        alertas.append("Justificativa pendente de espelhamento no Cartão de Ponto")
        
    if dim_c == "DIVERGENTE":
        criticos.append(f"Sobreaviso lançado para cargo operacional não autorizado ({cargo})")
        
    if dim_d == "SINDICATO DIVERGENTE":
        criticos.append(f"Código sindical da justificativa divergente do perfil LG ({perfil_frequencia})")
    elif dim_d == "PENDENTE (SEM PERFIL LG)" and dim_a != "PENDENTE":
        alertas.append("Código sindical não auditado (espelho de ponto/perfil LG pendente)")
        
    if dim_b == "ATENÇÃO / CORRIGIR":
        alertas.append(f"Infração de interjornada acima de 07:20 ({qtd_horas} registradas)")
        
    if criticos:
        status_geral = "DIVERGÊNCIA CRÍTICA"
        parecer = " | ".join(criticos + alertas)
    elif alertas:
        status_geral = "ALERTA OPERACIONAL"
        parecer = " | ".join(alertas)
    else:
        status_geral = "CONFORME"
        parecer = "Justificativa validada no espelho e em conformidade sindical, interjornada e sobreaviso."
        
    return dim_a, dim_b, dim_c, dim_d, status_geral, parecer

# ======================================================================================
# 3. LEITURA DE STREAMING E MATCHING DE DADOS POR REGIONAL
# ======================================================================================

def indexar_cartao_ponto(arquivos_cartao):
    """Carrega arquivos de Cartão de Ponto via streaming e indexa por Matrícula + Data."""
    cp_dict = {}
    total_linhas = 0
    
    for f_path in arquivos_cartao:
        wb = openpyxl.load_workbook(f_path, read_only=True, data_only=True)
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        header = [str(c).strip() if c is not None else '' for c in next(rows)]
        
        idx_mat = header.index('Matricula')
        idx_data = header.index('Data')
        idx_oc = header.index('Ocorrencia') if 'Ocorrencia' in header else -1
        idx_just = header.index('Justificativa') if 'Justificativa' in header else -1
        idx_perf = header.index('PerfilDeFrequencia') if 'PerfilDeFrequencia' in header else -1
        idx_esc = header.index('Escala') if 'Escala' in header else -1
        
        for r in rows:
            total_linhas += 1
            mat = r[idx_mat]
            dt = r[idx_data]
            if mat is None or dt is None:
                continue
            try:
                mat_int = int(float(mat))
                dt_str = str(dt)[:10]
                chave = f"{mat_int}_{dt_str}"
                if chave not in cp_dict:
                    cp_dict[chave] = []
                cp_dict[chave].append({
                    'Ocorrencia': r[idx_oc] if idx_oc != -1 else None,
                    'Justificativa': r[idx_just] if idx_just != -1 else None,
                    'PerfilDeFrequencia': r[idx_perf] if idx_perf != -1 else None,
                    'Escala': r[idx_esc] if idx_esc != -1 else None
                })
            except:
                continue
        wb.close()
        
    return cp_dict, total_linhas

def processar_regional(reg_nome):
    """Executa o processamento e auditoria completa de uma regional."""
    t0 = time.time()
    print(f"\n-> Processando Regional: {reg_nome}...")
    
    # Mapear arquivos do Cartão de Ponto
    if reg_nome == "PREZUNIC":
        cartao_files = [
            os.path.join(CARTAO_DIR, "PREZUNIC 700 a 710.xlsx"),
            os.path.join(CARTAO_DIR, "PREZUNIC 711 a 730.xlsx"),
            os.path.join(CARTAO_DIR, "PREZUNIC 731 a 750.xlsx")
        ]
    else:
        cartao_files = [os.path.join(CARTAO_DIR, f"{reg_nome}.xlsx")]
        
    # Indexar Cartão de Ponto
    cp_dict, total_cp_linhas = indexar_cartao_ponto(cartao_files)
    t_cp = time.time()
    print(f"   Cartão de Ponto indexado em {t_cp - t0:.2f}s ({total_cp_linhas:,} linhas / {len(cp_dict):,} chaves)")
    
    # Processar Ocorrências
    f_oc = os.path.join(OCORRENCIAS_DIR, f"{reg_nome}.xlsx")
    wb_oc = openpyxl.load_workbook(f_oc, read_only=True, data_only=True)
    ws_oc = wb_oc.active
    rows_oc = ws_oc.iter_rows(values_only=True)
    header_oc = [str(c).strip() if c is not None else '' for c in next(rows_oc)]
    
    cols_needed = ['Matricula', 'Nome', 'Cargo', 'Empresa', 'Estabelecimento', 'CodigoEstabelecimento', 'Data', 'MarcacoesDoDia', 'QtdDeHoras', 'Ocorrencia', 'Justificativa']
    col_map = {c: header_oc.index(c) for c in cols_needed if c in header_oc}
    
    linhas_auditadas = []
    
    for r in rows_oc:
        mat = r[col_map.get('Matricula')]
        dt = r[col_map.get('Data')]
        if mat is None or dt is None:
            continue
        try:
            mat_int = int(float(mat))
            dt_str = str(dt)[:10]
            chave = f"{mat_int}_{dt_str}"
            dt_br = f"{dt_str[8:10]}/{dt_str[5:7]}/{dt_str[0:4]}" if (len(dt_str) == 10 and dt_str[4] == '-' and dt_str[7] == '-') else dt_str
        except:
            continue
            
        nome = str(r[col_map['Nome']]) if r[col_map['Nome']] is not None else ''
        cargo = str(r[col_map['Cargo']]) if r[col_map['Cargo']] is not None else ''
        empresa = str(r[col_map['Empresa']]) if r[col_map['Empresa']] is not None else 'CENCOSUD'
        estab = str(r[col_map['Estabelecimento']]) if r[col_map['Estabelecimento']] is not None else ''
        cod_estab = str(r[col_map['CodigoEstabelecimento']]) if r[col_map['CodigoEstabelecimento']] is not None else ''
        marcacoes = str(r[col_map['MarcacoesDoDia']]) if r[col_map['MarcacoesDoDia']] is not None else ''
        qtd_h = str(r[col_map['QtdDeHoras']]) if r[col_map['QtdDeHoras']] is not None else ''
        oc_nome_raw = str(r[col_map['Ocorrencia']]) if r[col_map['Ocorrencia']] is not None else ''
        oc_just = str(r[col_map['Justificativa']]) if r[col_map['Justificativa']] is not None else ''
        
        oc_nome_norm = oc_nome_raw.strip().upper()
        norm_oc_just = normalizar_texto(oc_just)
        
        # Matching inteligente 1:1 com o Cartão de Ponto
        cand = cp_dict.get(chave, [])
        just_cartao = None
        perfil = None
        escala = None
        
        if cand:
            # 1. Match exato: mesma ocorrência e mesma justificativa
            match_perf = [c for c in cand if str(c.get('Ocorrencia', '')).strip().upper() == oc_nome_norm and normalizar_texto(c.get('Justificativa')) == norm_oc_just]
            if match_perf:
                escolhido = match_perf[0]
            else:
                # 2. Match de mesma justificativa
                match_just = [c for c in cand if normalizar_texto(c.get('Justificativa')) == norm_oc_just]
                if match_just:
                    escolhido = match_just[0]
                else:
                    # 3. Match de mesma ocorrência
                    match_oc = [c for c in cand if str(c.get('Ocorrencia', '')).strip().upper() == oc_nome_norm]
                    if match_oc:
                        com_just = [c for c in match_oc if c.get('Justificativa') is not None and str(c.get('Justificativa')).strip() != '']
                        escolhido = com_just[0] if com_just else match_oc[0]
                    else:
                        # 4. Fallback no dia
                        com_just = [c for c in cand if c.get('Justificativa') is not None and str(c.get('Justificativa')).strip() != '']
                        escolhido = com_just[0] if com_just else cand[0]
                        
            just_cartao = escolhido.get('Justificativa')
            perfil = escolhido.get('PerfilDeFrequencia')
            escala = escolhido.get('Escala')
            
        dim_a, dim_b, dim_c, dim_d, status_geral, parecer = auditar_registro(
            oc_just, just_cartao, oc_nome_raw, qtd_h, cargo, perfil
        )
        
        # Estrutura completa da linha analítica (20 colunas - retirada coluna D Estabelecimento)
        linhas_auditadas.append([
            reg_nome,
            empresa,
            cod_estab,
            mat_int,
            nome,
            cargo,
            dt_br,
            marcacoes,
            qtd_h,
            oc_nome_raw,
            oc_just,
            just_cartao if just_cartao is not None else "NÃO REFLETIDO",
            perfil if perfil is not None else "NÃO INFORMADO",
            escala if escala is not None else "NÃO INFORMADA",
            dim_a,
            dim_b,
            dim_c,
            dim_d,
            status_geral,
            parecer
        ])
        
    wb_oc.close()
    t_end = time.time()
    print(f"   Ocorrências auditadas em {t_end - t_cp:.2f}s | Total Registros: {len(linhas_auditadas):,}")
    
    return linhas_auditadas

# ======================================================================================
# 4. GERAÇÃO DAS 4 ABAS NO WORKBOOK EXCEL (OPENPYXL)
# ======================================================================================

def construir_aba_orientacoes(ws):
    """Constrói documentação institucional executiva e matriz de regras na Aba 1."""
    ws.title = "Orientações"
    ws.views.sheetView[0].showGridLines = True
    
    # Banner Principal
    ws.merge_cells('B2:K2')
    b1 = ws['B2']
    b1.value = "CENCOSUD BRASIL - DIRETORIA DE RECURSOS HUMANOS"
    b1.font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    b1.fill = PatternFill(start_color=C_AZUL_MARINHO, end_color=C_AZUL_MARINHO, fill_type="solid")
    b1.alignment = align_center
    ws.row_dimensions[2].height = 36
    
    ws.merge_cells('B3:K3')
    b2 = ws['B3']
    b2.value = "SUPER MÓDULO UNIFICADO DE JUSTIFICATIVAS E OCORRÊNCIAS DE PONTO | COMPETÊNCIA 09/2026"
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

    # 1. OBJETIVO & EMBASAMENTO LEGAL
    curr_row = add_section_header(curr_row, "1. OBJETIVO DO SUPER MÓDULO E EMBASAMENTO LEGAL")
    textos_obj = [
        "• O Super Módulo Unificado consolida 4 auditorias estratégicas de RH (Prioridades 2, 6, 8 e 9) em uma esteira única automatizada de alta performance.",
        "• Art. 74 da CLT (Registro de Ponto): Garante que toda justificativa lançada pelo gestor da filial seja fidedignamente espelhada no Cartão de Ponto sem distorções.",
        "• Art. 66 da CLT (Intervalo Interjornada): Estabelece repouso obrigatório mínimo de 11 horas consecutivas entre duas jornadas diárias de trabalho.",
        "• Art. 244, § 2º da CLT e ACTs (Sobreaviso): Restringe a escala de sobreaviso exclusivamente a cargos técnicos de suporte/TI com necessidade comprovada.",
        "• Enquadramento Sindical (LG): Assegura que códigos de horas extras e compensações sindicais pertençam estritamente ao sindicato do perfil do colaborador."
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
    
    # 2. AS 4 DIMENSÕES DE AUDITORIA
    curr_row = add_section_header(curr_row, "2. AS 4 DIMENSÕES ANALÍTICAS DO MOTOR DE AUDITORIA")
    headers_dim = ["Dimensão", "Prioridade", "Objeto Auditado", "Regra de Negócio e Validação Técnica"]
    cols_dim = [('B', 'C'), ('D', 'D'), ('E', 'F'), ('G', 'K')]
    
    for idx, (c1, c2) in enumerate(cols_dim):
        ws.merge_cells(f'{c1}{curr_row}:{c2}{curr_row}')
        cell = ws[f'{c1}{curr_row}']
        cell.value = headers_dim[idx]
        cell.font = font_header
        cell.fill = PatternFill(start_color=C_AZUL_MEDIO, end_color=C_AZUL_MEDIO, fill_type="solid")
        cell.alignment = align_center
    ws.row_dimensions[curr_row].height = 22
    curr_row += 1
    
    dim_data = [
        ("Dimensão A", "Prioridade 2", "Justificativa da Ocorrência vs. Cartão", "Normaliza textos e compara ocorrência com espelho. Classifica em CORRETO, PENDENTE ou DIVERGENTE."),
        ("Dimensão B", "Prioridade 8", "Descanso Interjornada (Art. 66 CLT)", "Identifica ocorrências de interjornada. Se acima de 07:20, sinaliza ATENÇÃO / CORRIGIR."),
        ("Dimensão C", "Prioridade 9", "Elegibilidade de Cargos em Sobreaviso", "Quando há Sobreaviso, retorna ELEGIVEL (cargos técnicos TI) ou DIVERGENTE (inelegíveis); sem Sobreaviso retorna em BRANCO."),
        ("Dimensão D", "Prioridade 6", "Aderência a Códigos Sindicais", "Extrai código numérico sindical da justificativa e valida contra o perfil cadastral do colaborador.")
    ]
    for r_data in dim_data:
        for idx, (c1, c2) in enumerate(cols_dim):
            ws.merge_cells(f'{c1}{curr_row}:{c2}{curr_row}')
            cell = ws[f'{c1}{curr_row}']
            cell.value = r_data[idx]
            cell.font = font_bold if idx in [0, 1] else font_base
            cell.alignment = align_center if idx in [0, 1] else align_left
            cell.border = border_thin
        ws.row_dimensions[curr_row].height = 20
        curr_row += 1
        
    curr_row += 1
    
    # 3. MATRIZ DE STATUS E AÇÕES OPERACIONAIS
    curr_row = add_section_header(curr_row, "3. MATRIZ DE SEVERIDADE DE STATUS E GUIA DE ATUAÇÃO DO RH")
    headers_st = ["Status Global", "Severidade", "Critério Técnico", "Ação Recomendada da Equipe de RH"]
    cols_st = [('B', 'C'), ('D', 'D'), ('E', 'G'), ('H', 'K')]
    
    for idx, (c1, c2) in enumerate(cols_st):
        ws.merge_cells(f'{c1}{curr_row}:{c2}{curr_row}')
        cell = ws[f'{c1}{curr_row}']
        cell.value = headers_st[idx]
        cell.font = font_header
        cell.fill = PatternFill(start_color=C_AZUL_MEDIO, end_color=C_AZUL_MEDIO, fill_type="solid")
        cell.alignment = align_center
    ws.row_dimensions[curr_row].height = 22
    curr_row += 1
    
    matriz_data = [
        ("CONFORME", "Verde / Regular", "Justificativa validada no cartão e aderente em interjornada, sobreaviso e sindicato.", "Nenhuma ação requerida. Espelho de ponto homologado para fechamento."),
        ("ALERTA OPERACIONAL", "Amarelo / Atenção", "Justificativa pendente no espelho de ponto ou interjornada > 07:20.", "Solicitar ao gestor da filial o recálculo/espelhamento no LG e verificar escala de folga."),
        ("DIVERGÊNCIA CRÍTICA", "Vermelho / Imediato", "Justificativa conflitante, sobreaviso em cargo operacional ou sindicato incompatível.", "BLOQUEIO DE FECHAMENTO: Retificar justificativa no ponto ou remover escala de sobreaviso.")
    ]
    for r_data in matriz_data:
        for idx, (c1, c2) in enumerate(cols_st):
            ws.merge_cells(f'{c1}{curr_row}:{c2}{curr_row}')
            cell = ws[f'{c1}{curr_row}']
            cell.value = r_data[idx]
            cell.alignment = align_center if idx in [0, 1] else align_left
            cell.border = border_thin
            if idx == 0:
                aplicar_badge(cell, r_data[0])
            else:
                cell.font = font_base
        ws.row_dimensions[curr_row].height = 24
        curr_row += 1

def aplicar_badge(cell, valor):
    """Aplica preenchimento e cor de texto em badges de status."""
    if valor is None:
        return
    v_upper = str(valor).upper().strip()
    if not v_upper:
        return
    if any(x in v_upper for x in ['CONFORME', 'CORRETO', 'ELEGIVEL', 'ELEGÍVEL', 'VALIDADO']):
        cell.fill = PatternFill(start_color=C_VERDE_BG, end_color=C_VERDE_BG, fill_type="solid")
        cell.font = Font(name="Calibri", size=10, bold=True, color=C_VERDE_TXT)
    elif any(x in v_upper for x in ['DIVERGÊNCIA CRÍTICA', 'DIVERGENTE', 'INELÉGÍVEL', 'SINDICATO DIVERGENTE']):
        cell.fill = PatternFill(start_color=C_VERM_BG, end_color=C_VERM_BG, fill_type="solid")
        cell.font = Font(name="Calibri", size=10, bold=True, color=C_VERM_TXT)
    elif any(x in v_upper for x in ['ALERTA OPERACIONAL', 'PENDENTE', 'ATENÇÃO']):
        cell.fill = PatternFill(start_color=C_AMAR_BG, end_color=C_AMAR_BG, fill_type="solid")
        cell.font = Font(name="Calibri", size=10, bold=True, color=C_AMAR_TXT)

def construir_aba_resumo(ws, resumo_regional, kpis_totais):
    """Constrói o dashboard executivo com cards e consolidação por regional na Aba 2."""
    ws.title = "Resumo da Auditoria"
    ws.views.sheetView[0].showGridLines = True
    
    # Banner
    ws.merge_cells('B2:N2')
    b = ws['B2']
    b.value = "CENCOSUD BRASIL - DASHBOARD EXECUTIVO DO SUPER MÓDULO DE JUSTIFICATIVAS E PONTO"
    b.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    b.fill = PatternFill(start_color=C_AZUL_MARINHO, end_color=C_AZUL_MARINHO, fill_type="solid")
    b.alignment = align_center
    ws.row_dimensions[2].height = 32
    
    # Cards Executivos de KPIs
    cards = [
        ("TOTAL OCORRÊNCIAS", f"{kpis_totais['total']:,}", "Auditadas (7 Regionais)", C_AZUL_MARINHO, "B", "C"),
        ("CONFORME / VALIDADO", f"{kpis_totais['conforme']:,}", f"Taxa: {kpis_totais['taxa_conf']:.1f}%", C_VERDE_TXT, "D", "E"),
        ("JUSTIFICATIVA PENDENTE", f"{kpis_totais['pendente']:,}", "Sem Espelho no Cartão", C_AMAR_TXT, "F", "G"),
        ("JUSTIF. DIVERGENTE", f"{kpis_totais['divergente_just']:,}", "Conflito de Código", C_VERM_TXT, "H", "I"),
        ("INTERJORNADA > 07:20", f"{kpis_totais['interjornada_alta']:,}", "Acima de 07:20", C_AMAR_TXT, "J", "K"),
        ("DIVERGÊNCIAS CRÍTICAS", f"{kpis_totais['criticos']:,}", "Bloqueio de Fechamento", C_VERM_TXT, "L", "N")
    ]
    
    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 28
    ws.row_dimensions[6].height = 16
    
    for title, val_str, sub_str, cor_txt, c1, c2 in cards:
        ws.merge_cells(f'{c1}4:{c2}4')
        cell_t = ws[f'{c1}4']
        cell_t.value = title
        cell_t.font = Font(name="Calibri", size=9, bold=True, color="64748B")
        cell_t.alignment = align_center
        cell_t.fill = PatternFill(start_color=C_CINZA_CAB, end_color=C_CINZA_CAB, fill_type="solid")
        cell_t.border = Border(top=Side(style='thin', color=C_BORDA), left=Side(style='thin', color=C_BORDA), right=Side(style='thin', color=C_BORDA))
        
        ws.merge_cells(f'{c1}5:{c2}5')
        cell_v = ws[f'{c1}5']
        cell_v.value = val_str
        cell_v.font = Font(name="Calibri", size=13, bold=True, color=cor_txt)
        cell_v.alignment = align_center
        cell_v.fill = PatternFill(start_color=C_BRANCO, end_color=C_BRANCO, fill_type="solid")
        cell_v.border = Border(left=Side(style='thin', color=C_BORDA), right=Side(style='thin', color=C_BORDA))
        
        ws.merge_cells(f'{c1}6:{c2}6')
        cell_s = ws[f'{c1}6']
        cell_s.value = sub_str
        cell_s.font = Font(name="Calibri", size=8, italic=True, color="64748B")
        cell_s.alignment = align_center
        cell_s.fill = PatternFill(start_color=C_CINZA_ZEBRA, end_color=C_CINZA_ZEBRA, fill_type="solid")
        cell_s.border = Border(bottom=Side(style='thin', color=C_BORDA), left=Side(style='thin', color=C_BORDA), right=Side(style='thin', color=C_BORDA))
        
    # Tabela Consolidada por Regional
    ws.merge_cells('B8:N8')
    t_reg = ws['B8']
    t_reg.value = "CONSOLIDAÇÃO EXECUTIVA POR REGIONAL / BANDEIRA (COMPETÊNCIA 09/2026)"
    t_reg.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    t_reg.fill = PatternFill(start_color=C_AZUL_MEDIO, end_color=C_AZUL_MEDIO, fill_type="solid")
    t_reg.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[8].height = 24
    
    headers_reg = [
        "Regional / Bandeira", "Total Ocorrências", "Conformes", "Alertas Operacionais",
        "Divergências Críticas", "Taxa Conform. (%)", "Dim A: Corretas", "Dim A: Pendentes",
        "Dim A: Divergentes", "Dim B: Interjornada > 07:20", "Dim C: Sobreaviso Diverg.",
        "Dim D: Sindicato Diverg.", "Status Geral Regional"
    ]
    
    ws.row_dimensions[9].height = 22
    for idx, h_text in enumerate(headers_reg, start=2):
        col_letter = get_column_letter(idx)
        cell = ws[f'{col_letter}9']
        cell.value = h_text
        cell.font = font_header
        cell.fill = PatternFill(start_color=C_AZUL_MARINHO, end_color=C_AZUL_MARINHO, fill_type="solid")
        cell.alignment = align_center
        cell.border = border_thin
        
    curr_row = 10
    for r in resumo_regional:
        ws.row_dimensions[curr_row].height = 20
        bg_color = C_CINZA_ZEBRA if curr_row % 2 == 1 else C_BRANCO
        fill_row = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        
        valores = [
            r['Regional'],
            r['Total'],
            r['Conforme'],
            r['Alerta'],
            r['Critico'],
            r['Taxa_Conf'] / 100.0,
            r['Dim_A_Correto'],
            r['Dim_A_Pendente'],
            r['Dim_A_Divergente'],
            r['Dim_B_Alerta'],
            r['Dim_C_Inelegivel'],
            r['Dim_D_Divergente'],
            r['Status_Regional']
        ]
        
        for idx, val in enumerate(valores, start=2):
            col_letter = get_column_letter(idx)
            cell = ws[f'{col_letter}{curr_row}']
            cell.value = val
            cell.border = border_thin
            if idx == 2:
                cell.alignment = align_left
                cell.font = font_bold
                cell.fill = fill_row
            elif idx == 7:
                cell.alignment = align_center
                cell.number_format = '0.0%'
                cell.font = font_base
                cell.fill = fill_row
            elif idx == 14:
                cell.alignment = align_left
                aplicar_badge(cell, val)
            else:
                cell.alignment = align_center
                cell.number_format = '#,##0'
                cell.font = font_base
                cell.fill = fill_row
                
        curr_row += 1
        
    # Linha de Totais da Tabela
    ws.row_dimensions[curr_row].height = 22
    fill_tot = PatternFill(start_color=C_CINZA_CAB, end_color=C_CINZA_CAB, fill_type="solid")
    
    totais_linha = [
        "TOTAL GERAL (7 REGIONAIS)",
        kpis_totais['total'],
        kpis_totais['conforme'],
        kpis_totais['alerta'],
        kpis_totais['criticos'],
        kpis_totais['taxa_conf'] / 100.0,
        kpis_totais['dim_a_correto'],
        kpis_totais['pendente'],
        kpis_totais['divergente_just'],
        kpis_totais['interjornada_alta'],
        kpis_totais['sobreaviso_inelegivel'],
        kpis_totais['sindicato_divergente'],
        "DIVERGÊNCIA CRÍTICA" if kpis_totais['criticos'] > 0 else "CONFORME"
    ]
    
    for idx, val in enumerate(totais_linha, start=2):
        col_letter = get_column_letter(idx)
        cell = ws[f'{col_letter}{curr_row}']
        cell.value = val
        cell.border = border_total
        if idx == 14:
            cell.alignment = align_left
            aplicar_badge(cell, val)
        else:
            cell.fill = fill_tot
            cell.font = font_bold
            if idx == 2:
                cell.alignment = align_left
            elif idx == 7:
                cell.alignment = align_center
                cell.number_format = '0.0%'
            else:
                cell.alignment = align_center
                cell.number_format = '#,##0'

# Super-cabeçalhos e Colunas da Base Analítica (Abas 3 e 4 - 20 Colunas)
SUPER_HEADERS_ANALITICA = [
    ("DADOS CADASTRAIS & LOCALIZAÇÃO", 1, 6, C_AZUL_MARINHO),
    ("OCORRÊNCIA LANÇADA PELO GESTOR", 7, 11, C_AZUL_MEDIO),
    ("ESPELHO DE PONTO (CARTÃO)", 12, 14, "1E3A8A"),
    ("MOTOR DE AUDITORIA (4 DIMENSÕES)", 15, 18, "334155"),
    ("DIAGNÓSTICO & PARECER FINAL", 19, 20, "0F172A")
]

COLUNAS_ANALITICA = [
    # 1 a 6: Cadastrais (Retirada coluna D Estabelecimento)
    ("Regional", "center", 14),
    ("Empresa", "center", 14),
    ("Cód. Estab.", "center", 12),
    ("Matrícula", "center", 14),
    ("Colaborador", "left", 36),
    ("Cargo", "left", 28),
    # 7 a 11: Ocorrência
    ("Data", "center", 13),
    ("Marcações do Dia", "center", 24),
    ("Qtd. Horas", "center", 12),
    ("Ocorrência", "left", 26),
    ("Justificativa Lançada", "left", 32),
    # 12 a 14: Cartão
    ("Justificativa Cartão", "left", 32),
    ("Perfil de Frequência", "left", 35),
    ("Escala", "left", 38),
    # 15 a 18: 4 Dimensões
    ("Dim A: Justificativa", "left", 22),
    ("Dim B: Interjornada", "left", 22),
    ("Dim C: Sobreaviso", "left", 22),
    ("Dim D: Sindicato", "left", 24),
    # 19 a 20: Diagnóstico
    ("Status Geral", "left", 22),
    ("Parecer da Auditoria", "left", 50)
]

def preparar_cabecalhos_analitica(ws, sheet_title):
    """Prepara o duplo cabeçalho estilizado para as abas analíticas."""
    ws.title = sheet_title
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = 'A3'
    
    # 1. Super-cabeçalho (Linha 1)
    ws.row_dimensions[1].height = 24
    for title, c_start, c_end, bg_color in SUPER_HEADERS_ANALITICA:
        start_letter = get_column_letter(c_start)
        end_letter = get_column_letter(c_end)
        ws.merge_cells(f'{start_letter}1:{end_letter}1')
        cell = ws[f'{start_letter}1']
        cell.value = title
        cell.font = font_super_header
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        cell.alignment = align_center
        for c_idx in range(c_start, c_end + 1):
            ws[f'{get_column_letter(c_idx)}1'].border = border_thin
            
    # 2. Cabeçalho de Colunas (Linha 2)
    ws.row_dimensions[2].height = 24
    for idx, (col_name, align_str, width) in enumerate(COLUNAS_ANALITICA, start=1):
        col_letter = get_column_letter(idx)
        cell = ws[f'{col_letter}2']
        cell.value = col_name
        cell.font = font_header
        cell.fill = PatternFill(start_color=C_AZUL_MARINHO, end_color=C_AZUL_MARINHO, fill_type="solid")
        cell.alignment = align_center
        cell.border = border_thin
        ws.column_dimensions[col_letter].width = width

# ======================================================================================
# 5. PIPELINE PRINCIPAL DE EXECUÇÃO
# ======================================================================================

def main():
    print("=" * 85)
    print("CENCOSUD BRASIL - ENGENHARIA DE DADOS & AUDITORIA DE RH")
    print("SUPER MÓDULO UNIFICADO DE JUSTIFICATIVAS E OCORRÊNCIAS DE PONTO (09/2026)")
    print("=" * 85)
    
    t_global_inicio = time.time()
    
    # Estruturas para consolidação
    todas_linhas_geral = []
    todas_linhas_diverg = []
    resumo_regional = []
    
    totais_gerais = {
        'total': 0,
        'conforme': 0,
        'alerta': 0,
        'criticos': 0,
        'dim_a_correto': 0,
        'pendente': 0,
        'divergente_just': 0,
        'interjornada_alta': 0,
        'sobreaviso_inelegivel': 0,
        'sindicato_divergente': 0
    }
    
    # Processar cada uma das 7 regionais
    for reg in REGIONAIS:
        linhas_reg = processar_regional(reg)
        todas_linhas_geral.extend(linhas_reg)
        
        # Estatísticas da regional (20 colunas: 14=DimA, 15=DimB, 16=DimC, 17=DimD, 18=Status, 19=Parecer)
        tot_reg = len(linhas_reg)
        conf_reg = sum(1 for r in linhas_reg if r[18] == 'CONFORME')
        alert_reg = sum(1 for r in linhas_reg if r[18] == 'ALERTA OPERACIONAL')
        crit_reg = sum(1 for r in linhas_reg if r[18] == 'DIVERGÊNCIA CRÍTICA')
        
        dim_a_corr = sum(1 for r in linhas_reg if r[14] == 'CORRETO')
        dim_a_pend = sum(1 for r in linhas_reg if r[14] == 'PENDENTE')
        dim_a_div = sum(1 for r in linhas_reg if r[14] == 'DIVERGENTE')
        
        dim_b_alert = sum(1 for r in linhas_reg if r[15] == 'ATENÇÃO / CORRIGIR')
        dim_c_ineleg = sum(1 for r in linhas_reg if r[16] == 'DIVERGENTE')
        dim_d_div = sum(1 for r in linhas_reg if r[17] == 'SINDICATO DIVERGENTE')
        
        # Filtrar divergências para Aba 4
        diverg_reg = [r for r in linhas_reg if r[18] != 'CONFORME']
        todas_linhas_diverg.extend(diverg_reg)
        
        status_reg = 'DIVERGÊNCIA CRÍTICA' if crit_reg > 0 else ('ALERTA OPERACIONAL' if alert_reg > 0 else 'CONFORME')
        taxa_conf_reg = (conf_reg / tot_reg * 100.0) if tot_reg > 0 else 0.0
        
        resumo_regional.append({
            'Regional': reg,
            'Total': tot_reg,
            'Conforme': conf_reg,
            'Alerta': alert_reg,
            'Critico': crit_reg,
            'Taxa_Conf': taxa_conf_reg,
            'Dim_A_Correto': dim_a_corr,
            'Dim_A_Pendente': dim_a_pend,
            'Dim_A_Divergente': dim_a_div,
            'Dim_B_Alerta': dim_b_alert,
            'Dim_C_Inelegivel': dim_c_ineleg,
            'Dim_D_Divergente': dim_d_div,
            'Status_Regional': status_reg
        })
        
        # Acumular totais globais
        totais_gerais['total'] += tot_reg
        totais_gerais['conforme'] += conf_reg
        totais_gerais['alerta'] += alert_reg
        totais_gerais['criticos'] += crit_reg
        totais_gerais['dim_a_correto'] += dim_a_corr
        totais_gerais['pendente'] += dim_a_pend
        totais_gerais['divergente_just'] += dim_a_div
        totais_gerais['interjornada_alta'] += dim_b_alert
        totais_gerais['sobreaviso_inelegivel'] += dim_c_ineleg
        totais_gerais['sindicato_divergente'] += dim_d_div

    totais_gerais['taxa_conf'] = (totais_gerais['conforme'] / totais_gerais['total'] * 100.0) if totais_gerais['total'] > 0 else 0.0
    
    print("\n" + "=" * 85)
    print(f"AUDITORIA CONCLUÍDA: {totais_gerais['total']:,} OCORRÊNCIAS PROCESSADAS")
    print(f" - Em Conformidade:            {totais_gerais['conforme']:,} ({totais_gerais['taxa_conf']:.1f}%)")
    print(f" - Alertas Operacionais:       {totais_gerais['alerta']:,}")
    print(f" - Divergências Críticas:      {totais_gerais['criticos']:,}")
    print(f" - Total Divergências + Alertas (Aba 4): {len(todas_linhas_diverg):,}")
    print("=" * 85)
    
    # Criação do Workbook Excel Oficial
    print("\n-> Gerando Workbook Excel Oficial no Padrão Cencosud de 4 Abas...")
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    
    # 1. Aba Orientações
    print("   [Aba 1/4] Construindo Orientações...")
    ws_ori = wb.create_sheet(title="Orientações")
    construir_aba_orientacoes(ws_ori)
    for c_idx in range(1, 12):
        ws_ori.column_dimensions[get_column_letter(c_idx)].width = 25
    ws_ori.column_dimensions['B'].width = 16
    ws_ori.column_dimensions['C'].width = 28
    ws_ori.column_dimensions['E'].width = 30
    ws_ori.column_dimensions['H'].width = 45
    
    # 2. Aba Resumo da Auditoria
    print("   [Aba 2/4] Construindo Resumo da Auditoria...")
    ws_res = wb.create_sheet(title="Resumo da Auditoria")
    construir_aba_resumo(ws_res, resumo_regional, totais_gerais)
    for c_idx in range(2, 15):
        col_letter = get_column_letter(c_idx)
        ws_res.column_dimensions[col_letter].width = 20
    ws_res.column_dimensions['B'].width = 22
    ws_res.column_dimensions['N'].width = 24
    
    # 3. Aba Conferência Geral
    print(f"   [Aba 3/4] Gravando Conferência Geral ({len(todas_linhas_geral):,} linhas)...")
    ws_conf = wb.create_sheet(title="Conferência Geral")
    preparar_cabecalhos_analitica(ws_conf, "Conferência Geral")
    
    t_w0 = time.time()
    for row_val in todas_linhas_geral:
        ws_conf.append(row_val)
    print(f"   Conferência Geral gravada via streaming em {time.time() - t_w0:.2f}s")
    
    # 4. Aba Divergências
    print(f"   [Aba 4/4] Gravando Aba Divergências ({len(todas_linhas_diverg):,} linhas com formatação executiva)...")
    ws_div = wb.create_sheet(title="Divergências")
    preparar_cabecalhos_analitica(ws_div, "Divergências")
    
    t_wd = time.time()
    curr_row = 3
    for row_val in todas_linhas_diverg:
        ws_div.append(row_val)
        ws_div.row_dimensions[curr_row].height = 20
        
        # Zebrado suave
        if curr_row % 2 == 1:
            bg_fill = PatternFill(start_color=C_CINZA_ZEBRA, end_color=C_CINZA_ZEBRA, fill_type="solid")
            for c_i in range(1, 21):
                ws_div[f'{get_column_letter(c_i)}{curr_row}'].fill = bg_fill
                
        # Aplicar badges nas colunas O, P, Q, R, S (índices 14 a 18)
        aplicar_badge(ws_div[f'O{curr_row}'], row_val[14]) # Dim A
        aplicar_badge(ws_div[f'P{curr_row}'], row_val[15]) # Dim B
        aplicar_badge(ws_div[f'Q{curr_row}'], row_val[16]) # Dim C
        aplicar_badge(ws_div[f'R{curr_row}'], row_val[17]) # Dim D
        aplicar_badge(ws_div[f'S{curr_row}'], row_val[18]) # Status Geral
            
        curr_row += 1
    # 5. Aba Painel de Interjornada (Art. 66 CLT)
    linhas_interjornada = []
    for r in todas_linhas_geral:
        txt_check = f"{r[9]} {r[10]} {r[11]}".upper()
        if "INTERJORNADA" in txt_check:
            oc_just_val = str(r[10]).strip() if r[10] else str(r[9]).strip()
            cp_just_val = str(r[11]).strip() if r[11] else "NÃO REFLETIDO"
            norm_oc = normalizar_texto(oc_just_val)
            norm_cp = normalizar_texto(cp_just_val)

            if (
                cp_just_val != "NÃO REFLETIDO"
                and norm_oc != ""
                and norm_cp != ""
                and (norm_oc == norm_cp or norm_oc in norm_cp or norm_cp in norm_oc)
            ):
                val_inter = "Validado"
                if r[15] == "ATENÇÃO / CORRIGIR":
                    acao_inter = "Validado no espelho, porém auditar interjornada acima de 07:20 (Art. 66 CLT)."
                else:
                    acao_inter = "Conforme — Justificativa de interjornada devidamente espelhada no Cartão de Ponto."
            else:
                val_inter = "Divergente"
                if cp_just_val == "NÃO REFLETIDO" or norm_cp == "":
                    acao_inter = "Regularizar espelhamento da justificativa de interjornada no Cartão de Ponto (Art. 66 CLT)."
                else:
                    acao_inter = f"Corrigir divergência no Cartão de Ponto: Lançado ('{oc_just_val}') vs Cartão ('{cp_just_val}')."

            linhas_interjornada.append([
                r[0],          # Regional
                r[1],          # Empresa
                r[2],          # Cód. Estab.
                r[3],          # Matrícula
                r[4],          # Colaborador
                r[5],          # Cargo
                r[6],          # Data (dd/mm/aaaa)
                r[7],          # Marcações do Dia
                r[8],          # Qtd. Horas
                oc_just_val,   # Justificativa Lançada
                cp_just_val,   # Justificativa no Cartão
                val_inter,     # Validação Interjornada
                acao_inter,    # Ação Recomendada RH
            ])

    # Ordenar Painel de Interjornada: Divergente primeiro, depois por Regional, Matrícula, Data
    linhas_interjornada.sort(key=lambda x: (0 if x[11] == "Divergente" else 1, x[0], x[3], f"{x[6][6:10]}-{x[6][3:5]}-{x[6][0:2]}" if len(x[6]) == 10 else x[6]))

    print(f"   [Aba 5/5] Gravando Aba Painel de Interjornada ({len(linhas_interjornada):,} ocorrências de interjornada)...")
    ws_inter = wb.create_sheet(title="Painel de Interjornada")
    ws_inter.views.sheetView[0].showGridLines = True
    ws_inter.freeze_panes = "A3"

    ws_inter.merge_cells("A1:M1")
    b_inter = ws_inter["A1"]
    b_inter.value = f"PAINEL DE AUDITORIA DE INTERJORNADA (ART. 66 CLT) — COMPETÊNCIA 09/2026 | {len(linhas_interjornada):,} OCORRÊNCIAS"
    b_inter.font = font_super_header
    b_inter.fill = PatternFill(start_color=C_AZUL_MARINHO, end_color=C_AZUL_MARINHO, fill_type="solid")
    b_inter.alignment = Alignment(horizontal="left", vertical="center")
    ws_inter.row_dimensions[1].height = 28

    cols_inter = [
        ("Regional", "center", 14),
        ("Empresa", "center", 14),
        ("Cód. Estab.", "center", 12),
        ("Matrícula", "center", 14),
        ("Colaborador", "left", 36),
        ("Cargo", "left", 28),
        ("Data", "center", 13),
        ("Marcações do Dia", "center", 26),
        ("Qtd. Horas", "center", 13),
        ("Justificativa Lançada", "left", 32),
        ("Justificativa no Cartão", "left", 32),
        ("Validação Interjornada", "center", 22),
        ("Ação Recomendada RH", "left", 56),
    ]

    ws_inter.row_dimensions[2].height = 24
    for idx, (col_name, _, width) in enumerate(cols_inter, start=1):
        col_letter = get_column_letter(idx)
        cell = ws_inter[f"{col_letter}2"]
        cell.value = col_name
        cell.font = font_header
        cell.fill = PatternFill(start_color=C_AZUL_MARINHO, end_color=C_AZUL_MARINHO, fill_type="solid")
        cell.alignment = align_center
        cell.border = border_thin
        ws_inter.column_dimensions[col_letter].width = width

    fill_div_inter = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_div_inter = Font(name="Calibri", size=10, bold=True, color="9C0006")
    fill_val_inter = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    font_val_inter = Font(name="Calibri", size=10, bold=True, color="006100")

    t_wi = time.time()
    for idx_i, row_inter in enumerate(linhas_interjornada):
        r_num = idx_i + 3
        ws_inter.append(row_inter)
        ws_inter.row_dimensions[r_num].height = 20
        fill_zebra = PatternFill(
            start_color=(C_CINZA_ZEBRA if idx_i % 2 == 1 else C_BRANCO),
            end_color=(C_CINZA_ZEBRA if idx_i % 2 == 1 else C_BRANCO),
            fill_type="solid",
        )
        for c_i, (_, align_str, _) in enumerate(cols_inter, start=1):
            c_cell = ws_inter.cell(r_num, c_i)
            c_cell.font = font_base
            c_cell.fill = fill_zebra
            c_cell.alignment = align_center if align_str == "center" else align_left
            c_cell.border = border_thin

        c_status_inter = ws_inter.cell(r_num, 12)
        if row_inter[11] == "Divergente":
            c_status_inter.fill = fill_div_inter
            c_status_inter.font = font_div_inter
        else:
            c_status_inter.fill = fill_val_inter
            c_status_inter.font = font_val_inter

    ws_inter.auto_filter.ref = f"A2:M{len(linhas_interjornada) + 2}"
    print(f"   Painel de Interjornada gravado e formatado em {time.time() - t_wi:.2f}s")

    # Remover aba padrão vazia
    wb.remove(default_sheet)

    # Salvar arquivo
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"\n-> Salvando arquivo final em:\n   {OUTPUT_FILE}")
    t_save0 = time.time()
    wb.save(OUTPUT_FILE)
    import shutil
    try:
        shutil.copy2(OUTPUT_FILE, OUTPUT_FILE_ROOT)
    except Exception:
        pass
    print(f"-> Arquivo salvo com sucesso em {time.time() - t_save0:.2f}s!")

    print("\n" + "=" * 85)
    print(f"PROCESSAMENTO GERAL CONCLUÍDO EM {time.time() - t_global_inicio:.1f} SEGUNDOS!")
    print("=" * 85)

if __name__ == "__main__":
    main()
