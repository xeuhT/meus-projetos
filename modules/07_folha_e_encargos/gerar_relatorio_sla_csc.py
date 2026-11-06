"""
=====================================================================================
CENCOSUD BRASIL - ENGENHARIA DE DADOS & AUDITORIA DE RH
MÓDULO 8: CONTROLE E AUDITORIA DE SLA — CHAMADOS DO PORTAL CSC (PONTO & BENEFÍCIOS)
=====================================================================================
Esteira automatizada de consolidação e monitoramento preventivo de SLA dos chamados do
Portal CSC Brasil (Prioridade 12 do Levantamento de Conferências da Área).

Fontes de Dados Integradas (SLA CSC / 00- CONTROLE SLA):
  1. Afastamento de Colaborador-CSC-Brasil.xlsx (SLA: 8 dias)
  2. Cadastro Feriados Estaduais e Municipais-CSC-Brasil.xlsx (SLA: 8 dias)
  3. Controle de Frequência-CSC-Brasil.xlsx (SLA: 15 dias)
  4. CONTROLE SLA - CSC.xlsx (Planilha Modelo da Equipe)

Regras de Negócio Oficiais:
  - Filtro de Chamados Ativos: Considera chamados com Status '1 - Pendente de Atendimento'
    ou '2 - Em Atendimento' (descartando chamados já finalizados ou cancelados).
  - Cálculo de Vencimento: Data de Vencimento = Data de Criação + Prazo de Atendimento.
  - Margem de Segurança Preventiva (3 Dias): Conforme diretriz da área, chamados cuja Data
    de Vencimento esteja a até 3 dias da data-base de corte já recebem Situação = 'Vencido'
    para atuação preventiva antes do estouro real do SLA.

Padrão Cencosud de 4 Abas:
  - Aba 1: Orientações (Regras de SLA, Margem de Segurança de 3 Dias e Fontes CSC)
  - Aba 2: Resumo da Auditoria (KPIs Executivos, Resumo por Origem, Bandeira e Tipo de Solicitação)
  - Aba 3: Conferência Geral (Base Consolidada Completa de 257 Chamados Ativos no layout CSC)
  - Aba 4: Divergências (Painel Prioritário dos 95 Chamados Vencidos / Na Margem de Segurança)
=====================================================================================
"""

import os
import time
from datetime import datetime, timedelta
from collections import Counter, defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ======================================================================================
# 1. DIRETÓRIOS E ARQUIVOS
# ======================================================================================

BASE_DIR = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\ponto"
SLA_DIR = os.path.join(BASE_DIR, "SLA CSC")
INPUT_SUBDIR = os.path.join(SLA_DIR, "00- CONTROLE SLA")
OUTPUT_DIR = os.path.join(BASE_DIR, "conferencias geradas pelo script")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Relatorio_Conferencia_SLA_CSC_09.2026.xlsx")

FILE_AFASTAMENTO = os.path.join(INPUT_SUBDIR, "Afastamento de Colaborador-CSC-Brasil.xlsx")
FILE_FERIADOS = os.path.join(INPUT_SUBDIR, "Cadastro Feriados Estaduais e Municipais-CSC-Brasil.xlsx")
FILE_FREQUENCIA = os.path.join(INPUT_SUBDIR, "Controle de Frequência-CSC-Brasil.xlsx")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ======================================================================================
# 2. PALETA CORPORATIVA CENCOSUD E ESTILOS
# ======================================================================================

C_AZUL_MARINHO = "1B365D"
C_AZUL_MEDIO   = "2C4A7A"
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

# ======================================================================================
# 3. FUNÇÕES AUXILIARES DE PARSING DE DATAS E REGRAS DE SLA
# ======================================================================================

def parse_date(val):
    if val is None:
        return None
    if isinstance(val, datetime):
        return datetime(val.year, val.month, val.day)
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            dt = datetime.strptime(s[:len(fmt) if "H" not in fmt else len(s)], fmt)
            return datetime(dt.year, dt.month, dt.day)
        except Exception:
            pass
    try:
        dt = datetime.strptime(s[:10], "%d/%m/%Y")
        return datetime(dt.year, dt.month, dt.day)
    except Exception:
        return None

def format_dt_br(dt_obj):
    if dt_obj is None:
        return ""
    if isinstance(dt_obj, datetime):
        return dt_obj.strftime("%d/%m/%Y")
    return str(dt_obj)

def is_chamado_ativo(status_str):
    """Retorna True se o chamado estiver Pendente de Atendimento (1) ou Em Atendimento (2)."""
    if not status_str:
        return False
    s = str(status_str).strip().upper()
    return s.startswith("1") or s.startswith("2") or "PENDENTE" in s or "EM ATENDIMENTO" in s

# ======================================================================================
# 4. LEITURA E CONSOLIDAÇÃO DOS 3 RELATÓRIOS DO PORTAL CSC
# ======================================================================================

def processar_relatorios_csc():
    print("-> Lendo e consolidando os 3 relatórios do Portal CSC Brasil...")

    chamados_brutos = []
    descartados_finalizados = 0

    # 1. Controle de Frequência-CSC-Brasil.xlsx
    wb_cf = openpyxl.load_workbook(FILE_FREQUENCIA, data_only=True)
    ws_cf = wb_cf.active
    rows_cf = list(ws_cf.iter_rows(min_row=2, values_only=True))
    wb_cf.close()
    for r in rows_cf:
        if not any(r):
            continue
        status = str(r[3]).strip() if r[3] else ""
        if not is_chamado_ativo(status):
            descartados_finalizados += 1
            continue
        dt_criacao = parse_date(r[0])
        prazo = int(float(r[1])) if r[1] is not None else 15
        chamados_brutos.append({
            "origem": "Controle de Frequência",
            "dt_criacao": dt_criacao,
            "prazo": prazo,
            "identificador": str(r[2]).strip() if r[2] else "",
            "status": status,
            "ult_tramitacao": str(r[4]).strip() if r[4] else "",
            "tipo_solicitacao": str(r[6]).strip() if r[6] else "",
            "bandeira": str(r[7]).strip() if r[7] else "",
            "loja": str(r[8]).strip() if r[8] else "",
            "matricula": str(r[9]).strip() if r[9] else "",
            "nome_colab": str(r[10]).strip() if r[10] else ""
        })

    # 2. Afastamento de Colaborador-CSC-Brasil.xlsx
    wb_af = openpyxl.load_workbook(FILE_AFASTAMENTO, data_only=True)
    ws_af = wb_af.active
    rows_af = list(ws_af.iter_rows(min_row=2, values_only=True))
    wb_af.close()
    for r in rows_af:
        if not any(r):
            continue
        status = str(r[3]).strip() if r[3] else ""
        if not is_chamado_ativo(status):
            descartados_finalizados += 1
            continue
        dt_criacao = parse_date(r[0])
        prazo = int(float(r[1])) if r[1] is not None else 8
        motivo_af = str(r[11]).strip() if len(r) > 11 and r[11] else "Afastamento de Colaborador"
        mat_af = str(r[10]).strip() if len(r) > 10 and r[10] else ""
        nome_af = str(r[9]).strip() if len(r) > 9 and r[9] else ""
        chamados_brutos.append({
            "origem": "Afastamento",
            "dt_criacao": dt_criacao,
            "prazo": prazo,
            "identificador": str(r[2]).strip() if r[2] else "",
            "status": status,
            "ult_tramitacao": str(r[4]).strip() if r[4] else "",
            "tipo_solicitacao": motivo_af,
            "bandeira": str(r[7]).strip() if r[7] else "",
            "loja": str(r[8]).strip() if r[8] else "",
            "matricula": mat_af,
            "nome_colab": nome_af
        })

    # 3. Cadastro Feriados Estaduais e Municipais-CSC-Brasil.xlsx
    wb_fe = openpyxl.load_workbook(FILE_FERIADOS, data_only=True)
    ws_fe = wb_fe.active
    rows_fe = list(ws_fe.iter_rows(min_row=2, values_only=True))
    wb_fe.close()
    for r in rows_fe:
        if not any(r):
            continue
        status = str(r[3]).strip() if r[3] else ""
        if not is_chamado_ativo(status):
            descartados_finalizados += 1
            continue
        dt_criacao = parse_date(r[0])
        prazo = int(float(r[1])) if r[1] is not None else 8
        mun_uf = str(r[9]).strip() if len(r) > 9 and r[9] else ""
        chamados_brutos.append({
            "origem": "Cadastro Feriados",
            "dt_criacao": dt_criacao,
            "prazo": prazo,
            "identificador": str(r[2]).strip() if r[2] else "",
            "status": status,
            "ult_tramitacao": str(r[4]).strip() if r[4] else "",
            "tipo_solicitacao": str(r[6]).strip() if r[6] else "Cadastro Feriado Municipal",
            "bandeira": str(r[7]).strip() if r[7] else "",
            "loja": str(r[8]).strip() if r[8] else "",
            "matricula": "-",
            "nome_colab": f"Município: {mun_uf}" if mun_uf else "-"
        })

    # Data de referência da extração (25/10/2026 -> com margem de segurança de 3 dias = 28/10/2026)
    # Reproduzindo exatamente a regra da planilha CONTROLE SLA - CSC.xlsx
    dt_referencia = datetime(2026, 10, 25)
    dt_corte_margem = dt_referencia + timedelta(days=3)  # 28/10/2026

    for c in chamados_brutos:
        dt_venc = c["dt_criacao"] + timedelta(days=c["prazo"])
        c["dt_vencimento"] = dt_venc
        dias_restantes = (dt_venc - dt_referencia).days
        c["dias_restantes"] = dias_restantes

        if dt_venc <= dt_corte_margem:
            c["situacao"] = "Vencido"
            if dt_venc < dt_referencia:
                c["sub_status"] = "Vencido (Prazo Expirado)"
                c["severidade"] = "Crítico"
                c["acao_rh"] = f"SLA expirado há {abs(dias_restantes)} dia(s). Prioridade máxima de atendimento no Portal CSC."
            else:
                c["sub_status"] = "Vencido (Margem Segurança 3d)"
                c["severidade"] = "Alerta"
                c["acao_rh"] = f"Vence em {dias_restantes} dia(s) ({format_dt_br(dt_venc)}). Sinalizado como Vencido pela margem preventiva de 3 dias."
        else:
            c["situacao"] = "Em dia"
            c["sub_status"] = "Em Dia (Dentro do Prazo)"
            c["severidade"] = "Regular"
            c["acao_rh"] = f"Dentro do prazo SLA ({dias_restantes} dias restantes até o vencimento em {format_dt_br(dt_venc)})."

    # Ordenar por Data de Vencimento crescente e Identificador
    chamados_ordenados = sorted(chamados_brutos, key=lambda x: (x["dt_vencimento"], x["dt_criacao"], x["identificador"]))

    linhas_geral = []
    linhas_divergencias = []

    kpis = {
        "total_ativos": len(chamados_ordenados),
        "descartados": descartados_finalizados,
        "vencidos_total": 0,
        "vencidos_expirados": 0,
        "vencidos_margem_3d": 0,
        "em_dia": 0
    }

    stats_origem = defaultdict(lambda: {"total": 0, "expirado": 0, "margem": 0, "vencido_tot": 0, "em_dia": 0, "prazo": 0})
    stats_bandeira = defaultdict(lambda: {"total": 0, "expirado": 0, "margem": 0, "vencido_tot": 0, "em_dia": 0})
    stats_tipo = defaultdict(lambda: {"origem": "", "total": 0, "vencido_tot": 0, "em_dia": 0})

    for c in chamados_ordenados:
        orig = c["origem"]
        band = c["bandeira"] if c["bandeira"] else "Não Informada"
        tipo = c["tipo_solicitacao"] if c["tipo_solicitacao"] else "Não Especificado"

        so = stats_origem[orig]
        so["total"] += 1
        so["prazo"] = c["prazo"]

        sb = stats_bandeira[band]
        sb["total"] += 1

        st = stats_tipo[tipo]
        st["origem"] = orig
        st["total"] += 1

        if c["situacao"] == "Vencido":
            kpis["vencidos_total"] += 1
            so["vencido_tot"] += 1
            sb["vencido_tot"] += 1
            st["vencido_tot"] += 1
            if c["severidade"] == "Crítico":
                kpis["vencidos_expirados"] += 1
                so["expirado"] += 1
                sb["expirado"] += 1
            else:
                kpis["vencidos_margem_3d"] += 1
                so["margem"] += 1
                sb["margem"] += 1
        else:
            kpis["em_dia"] += 1
            so["em_dia"] += 1
            sb["em_dia"] += 1
            st["em_dia"] += 1

        # As 13 primeiras colunas seguem fielmente o modelo CONTROLE SLA - CSC.xlsx + 2 colunas de diagnóstico
        linha = [
            c["origem"],
            format_dt_br(c["dt_criacao"]),
            c["prazo"],
            format_dt_br(c["dt_vencimento"]),
            c["situacao"],
            c["identificador"],
            c["status"],
            c["ult_tramitacao"],
            c["tipo_solicitacao"],
            c["bandeira"],
            c["loja"],
            c["matricula"],
            c["nome_colab"],
            c["sub_status"],
            c["acao_rh"]
        ]
        linhas_geral.append(linha)
        if c["situacao"] == "Vencido":
            linhas_divergencias.append(linha)

    return linhas_geral, linhas_divergencias, kpis, stats_origem, stats_bandeira, stats_tipo

# ======================================================================================
# 5. CONSTRUÇÃO DO WORKBOOK EXCEL OFICIAL NO PADRÃO CENCOSUD DE 4 ABAS
# ======================================================================================

def construir_aba_orientacoes(ws):
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:K1")
    c1 = ws["A1"]
    c1.value = "CENCOSUD BRASIL — CONTROLE E AUDITORIA DE SLA DO PORTAL CSC (PONTO & BENEFÍCIOS)"
    c1.fill = FILL_HEADER
    c1.font = FONT_TITULO
    c1.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 38

    ws.merge_cells("A2:K2")
    c2 = ws["A2"]
    c2.value = "Monitoramento Preventivo de Chamados Pendentes com Margem de Segurança de 3 Dias — Competência 10/2026"
    c2.fill = FILL_HEADER_SEC
    c2.font = Font(name="Calibri", size=11, italic=True, color="FFFFFF")
    c2.alignment = ALIGN_CENTER
    ws.row_dimensions[2].height = 24

    secoes = [
        ("1. OBJETIVO DO MÓDULO DE CONTROLE DE SLA CSC",
         "Unificar e monitorar diariamente todos os chamados abertos no Portal CSC Brasil referentes à área de Ponto e Benefícios, "
         "garantindo o tratamento tempestivo antes do estouro do prazo regulamentar (SLA)."),

        ("2. FONTES DE DADOS & PRAZOS DE ATENDIMENTO POR FILA",
         "• Controle de Frequência-CSC-Brasil.xlsx (Origem: Controle de Frequência) — Prazo de Atendimento (SLA): 15 dias corridos.\n"
         "• Afastamento de Colaborador-CSC-Brasil.xlsx (Origem: Afastamento) — Prazo de Atendimento (SLA): 8 dias corridos.\n"
         "• Cadastro Feriados Estaduais e Municipais-CSC-Brasil.xlsx (Origem: Cadastro Feriados) — Prazo de Atendimento (SLA): 8 dias corridos.\n"
         "• Filtro de Status Ativo: São consolidados exclusivamente os chamados com Status '1 - Pendente de Atendimento' ou '2 - Em Atendimento' (descartando chamados já finalizados ou cancelados)."),

        ("3. REGRA DA MARGEM DE SEGURANÇA DE 3 DIAS (STATUS VENCIDO PREVENTIVO)",
         "• Cálculo da Data de Vencimento: Data de Criação + Prazo de Atendimento (8 ou 15 dias).\n"
         "• Margem de Segurança de 3 Dias: Para evitar o estouro real do SLA, o módulo aplica uma trava preventiva de 3 dias de antecedência. "
         "Assim, qualquer chamado cuja Data de Vencimento esteja a 3 dias ou menos da data de referência já recebe Situação = 'Vencido', "
         "entrando imediatamente na fila prioritária de atuação (Aba Divergências).\n"
         "• Enriquecimento Automático: Recupera automaticamente Matrícula, Nome do Colaborador e Motivo de Afastamento nos chamados de Afastamento e o Município/UF nos chamados de Cadastro de Feriados.")
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
        ws.row_dimensions[cur_row].height = max(48, linhas_txt * 16)
        cur_row += 2

    for c in range(1, 12):
        ws.column_dimensions[get_column_letter(c)].width = 16

def construir_aba_resumo(ws, kpis, stats_origem, stats_bandeira, stats_tipo):
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:L1")
    ct = ws["A1"]
    ct.value = "CENCOSUD BRASIL — DASHBOARD EXECUTIVO DE CONTROLE DE SLA (PORTAL CSC)"
    ct.fill = FILL_HEADER
    ct.font = FONT_TITULO
    ct.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 36

    taxa_em_dia = (kpis["em_dia"] / kpis["total_ativos"]) if kpis["total_ativos"] > 0 else 1.0

    cards_l1 = [
        ("TOTAL CHAMADOS PENDENTES", f"{kpis['total_ativos']:,}", f"{kpis['descartados']} Finalizados/Cancelados ignorados", C_AZUL_MARINHO, 1, 3),
        ("CHAMADOS EM DIA (REGULARES)", f"{kpis['em_dia']:,}", f"Taxa Dentro do SLA: {taxa_em_dia:.1%}", "166534", 4, 6),
        ("TOTAL STATUS 'VENCIDO' (C/ MARGEM)", f"{kpis['vencidos_total']:,}", "Expirados + Margem Preventiva 3d", C_CRIT_TXT, 7, 9),
        ("VENCIDOS REAIS (EXPIRADOS)", f"{kpis['vencidos_expirados']:,}", "Prazo Final Ultrapassado", C_CRIT_TXT, 10, 12)
    ]

    cards_l2 = [
        ("NA MARGEM DE SEGURANÇA (3 DIAS)", f"{kpis['vencidos_margem_3d']:,}", "Vencendo nos próximos 3 dias", C_ALERT_TXT, 1, 4),
        ("FILA CONTROLE DE FREQUÊNCIA", f"{stats_origem['Controle de Frequência']['total']:,}", f"SLA 15 dias | {stats_origem['Controle de Frequência']['vencido_tot']} Vencidos", C_AZUL_MARINHO, 5, 8),
        ("FILAS AFASTAMENTO & FERIADOS", f"{stats_origem['Afastamento']['total'] + stats_origem['Cadastro Feriados']['total']:,}", "SLA 8 dias | 100% Prioridade", C_CRIT_TXT, 9, 12)
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

    # TABELA 1: CONSOLIDAÇÃO POR ORIGEM / FILA DO CSC
    cur_r = 11
    ws.merge_cells(f"A{cur_r}:L{cur_r}")
    t1 = ws[f"A{cur_r}"]
    t1.value = "1. CONSOLIDAÇÃO DE SLA POR FILA DE ORIGEM NO PORTAL CSC"
    t1.fill = FILL_HEADER_SEC
    t1.font = FONT_SUBTITULO
    t1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[cur_r].height = 24
    cur_r += 1

    headers_t1 = [
        "Fila de Origem (Portal CSC)", "Prazo SLA (Dias)", "Total Pendentes",
        "Vencidos (Expirados)", "Vencidos (Margem 3d)", "Total Status 'Vencido'",
        "Em Dia", "% Em Dia", "Status da Fila"
    ]
    cols_t1 = [(1, 3), (4, 4), (5, 5), (6, 6), (7, 7), (8, 8), (9, 9), (10, 10), (11, 12)]
    ws.row_dimensions[cur_r].height = 25
    for (ci, cf), h in zip(cols_t1, headers_t1):
        if ci != cf:
            ws.merge_cells(start_row=cur_r, start_column=ci, end_row=cur_r, end_column=cf)
        c = ws.cell(row=cur_r, column=ci, value=h)
        c.alignment = ALIGN_CENTER
        for c_i in range(ci, cf + 1):
            ws.cell(row=cur_r, column=c_i).fill = FILL_HEADER
            ws.cell(row=cur_r, column=c_i).font = FONT_CABECALHO
            ws.cell(row=cur_r, column=c_i).border = BORDER_THIN
    cur_r += 1

    for orig_nome in ["Controle de Frequência", "Afastamento", "Cadastro Feriados"]:
        so = stats_origem[orig_nome]
        pct = (so["em_dia"] / so["total"]) if so["total"] > 0 else 1.0
        st_fila = "Crítico (Atuação Imediata)" if so["vencido_tot"] > 0 else "Regular"
        fill_r = FILL_ZEBRA if (cur_r % 2 == 0) else FILL_BRANCO

        vals_t1 = [
            orig_nome, f"{so['prazo']} dias", so["total"],
            so["expirado"], so["margem"], so["vencido_tot"],
            so["em_dia"], pct, st_fila
        ]
        ws.row_dimensions[cur_r].height = 21
        for idx_col, ((ci, cf), val) in enumerate(zip(cols_t1, vals_t1)):
            if ci != cf:
                ws.merge_cells(start_row=cur_r, start_column=ci, end_row=cur_r, end_column=cf)
            c = ws.cell(row=cur_r, column=ci, value=val)
            for c_i in range(ci, cf + 1):
                ws.cell(row=cur_r, column=c_i).fill = fill_r
                ws.cell(row=cur_r, column=c_i).border = BORDER_THIN
            c.font = FONT_DADOS_BOLD if idx_col in (0, 2, 5) else FONT_DADOS
            if idx_col == 0:
                c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            elif idx_col == 7:
                c.alignment = ALIGN_RIGHT
                c.number_format = "0.0%"
            elif idx_col == 8:
                c.alignment = ALIGN_CENTER
                for c_i in range(ci, cf + 1):
                    ws.cell(row=cur_r, column=c_i).fill = FILL_CRIT if so["vencido_tot"] > 0 else FILL_REG
                c.font = FONT_CRIT if so["vencido_tot"] > 0 else FONT_REG
            else:
                c.alignment = ALIGN_CENTER
        cur_r += 1

    # Total Tabela 1
    tot_t1 = [
        "TOTAL GERAL (3 FILAS CSC)", "-", kpis["total_ativos"],
        kpis["vencidos_expirados"], kpis["vencidos_margem_3d"], kpis["vencidos_total"],
        kpis["em_dia"], taxa_em_dia, f"{kpis['vencidos_total']} Vencidos"
    ]
    ws.row_dimensions[cur_r].height = 24
    for idx_col, ((ci, cf), val) in enumerate(zip(cols_t1, tot_t1)):
        if ci != cf:
            ws.merge_cells(start_row=cur_r, start_column=ci, end_row=cur_r, end_column=cf)
        c = ws.cell(row=cur_r, column=ci, value=val)
        for c_i in range(ci, cf + 1):
            ws.cell(row=cur_r, column=c_i).fill = FILL_TOTAL
            ws.cell(row=cur_r, column=c_i).border = BORDER_TOTAL
        c.font = Font(name="Calibri", size=10, bold=True, color=C_AZUL_MARINHO)
        if idx_col == 0:
            c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        elif idx_col == 7:
            c.alignment = ALIGN_RIGHT
            c.number_format = "0.0%"
        else:
            c.alignment = ALIGN_CENTER
    cur_r += 3

    # TABELA 2: CONSOLIDAÇÃO POR BANDEIRA / REGIONAL
    ws.merge_cells(f"A{cur_r}:L{cur_r}")
    t2 = ws[f"A{cur_r}"]
    t2.value = "2. DISTRIBUIÇÃO DE CHAMADOS POR BANDEIRA / REGIONAL"
    t2.fill = FILL_HEADER_SEC
    t2.font = FONT_SUBTITULO
    t2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[cur_r].height = 24
    cur_r += 1

    headers_t2 = [
        "Bandeira / Regional", "Total Chamados", "Vencidos (Expirados)",
        "Vencidos (Margem 3d)", "Total Status 'Vencido'", "Em Dia", "% Dentro do SLA", "Status Bandeira"
    ]
    cols_t2 = [(1, 3), (4, 4), (5, 5), (6, 6), (7, 7), (8, 8), (9, 10), (11, 12)]
    ws.row_dimensions[cur_r].height = 25
    for (ci, cf), h in zip(cols_t2, headers_t2):
        if ci != cf:
            ws.merge_cells(start_row=cur_r, start_column=ci, end_row=cur_r, end_column=cf)
        c = ws.cell(row=cur_r, column=ci, value=h)
        c.alignment = ALIGN_CENTER
        for c_i in range(ci, cf + 1):
            ws.cell(row=cur_r, column=c_i).fill = FILL_HEADER
            ws.cell(row=cur_r, column=c_i).font = FONT_CABECALHO
            ws.cell(row=cur_r, column=c_i).border = BORDER_THIN
    cur_r += 1

    for band_nome, sb in sorted(stats_bandeira.items(), key=lambda x: (-x[1]["vencido_tot"], -x[1]["total"])):
        pct_b = (sb["em_dia"] / sb["total"]) if sb["total"] > 0 else 1.0
        st_b = f"{sb['vencido_tot']} Vencidos" if sb["vencido_tot"] > 0 else "100% Em Dia"
        fill_r = FILL_ZEBRA if (cur_r % 2 == 0) else FILL_BRANCO

        vals_t2 = [
            band_nome, sb["total"], sb["expirado"], sb["margem"],
            sb["vencido_tot"], sb["em_dia"], pct_b, st_b
        ]
        ws.row_dimensions[cur_r].height = 20
        for idx_col, ((ci, cf), val) in enumerate(zip(cols_t2, vals_t2)):
            if ci != cf:
                ws.merge_cells(start_row=cur_r, start_column=ci, end_row=cur_r, end_column=cf)
            c = ws.cell(row=cur_r, column=ci, value=val)
            for c_i in range(ci, cf + 1):
                ws.cell(row=cur_r, column=c_i).fill = fill_r
                ws.cell(row=cur_r, column=c_i).border = BORDER_THIN
            c.font = FONT_DADOS_BOLD if idx_col in (0, 1, 4) else FONT_DADOS
            if idx_col == 0:
                c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            elif idx_col == 6:
                c.alignment = ALIGN_RIGHT
                c.number_format = "0.0%"
            elif idx_col == 7:
                c.alignment = ALIGN_CENTER
                for c_i in range(ci, cf + 1):
                    ws.cell(row=cur_r, column=c_i).fill = FILL_CRIT if sb["vencido_tot"] > 0 else FILL_REG
                c.font = FONT_CRIT if sb["vencido_tot"] > 0 else FONT_REG
            else:
                c.alignment = ALIGN_CENTER
        cur_r += 1

    cur_r += 2

    # TABELA 3: CONSOLIDAÇÃO POR TIPO DE SOLICITAÇÃO
    ws.merge_cells(f"A{cur_r}:L{cur_r}")
    t3 = ws[f"A{cur_r}"]
    t3.value = "3. DETALHAMENTO POR TIPO DE SOLICITAÇÃO NO CSC"
    t3.fill = FILL_HEADER_SEC
    t3.font = FONT_SUBTITULO
    t3.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[cur_r].height = 24
    cur_r += 1

    cols_t3 = [(1, 3, "Fila de Origem"), (4, 7, "Tipo de Solicitação"), (8, 8, "Total"), (9, 10, "Status 'Vencido'"), (11, 12, "Em Dia")]
    ws.row_dimensions[cur_r].height = 24
    for ci, cf, h in cols_t3:
        if ci != cf:
            ws.merge_cells(start_row=cur_r, start_column=ci, end_row=cur_r, end_column=cf)
        c = ws.cell(row=cur_r, column=ci, value=h)
        c.alignment = ALIGN_CENTER
        for c_i in range(ci, cf + 1):
            ws.cell(row=cur_r, column=c_i).fill = FILL_HEADER
            ws.cell(row=cur_r, column=c_i).font = FONT_CABECALHO
            ws.cell(row=cur_r, column=c_i).border = BORDER_THIN
    cur_r += 1

    for tipo_nome, st in sorted(stats_tipo.items(), key=lambda x: (-x[1]["vencido_tot"], -x[1]["total"])):
        fill_r = FILL_ZEBRA if (cur_r % 2 == 0) else FILL_BRANCO
        ws.row_dimensions[cur_r].height = 20
        vals_t3 = [st["origem"], tipo_nome, st["total"], st["vencido_tot"], st["em_dia"]]
        for idx_col, ((ci, cf, _), val) in enumerate(zip(cols_t3, vals_t3)):
            if ci != cf:
                ws.merge_cells(start_row=cur_r, start_column=ci, end_row=cur_r, end_column=cf)
            c = ws.cell(row=cur_r, column=ci, value=val)
            for c_i in range(ci, cf + 1):
                ws.cell(row=cur_r, column=c_i).fill = fill_r
                ws.cell(row=cur_r, column=c_i).border = BORDER_THIN
            c.font = FONT_DADOS_BOLD if idx_col in (1, 3) and st["vencido_tot"] > 0 else FONT_DADOS
            if idx_col in (0, 1):
                c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            else:
                c.alignment = ALIGN_CENTER
        cur_r += 1

    for col in range(1, 13):
        ws.column_dimensions[get_column_letter(col)].width = 16

def gravar_aba_chamados(ws, linhas, titulo_banner):
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A3"

    ws.merge_cells("A1:O1")
    ct = ws["A1"]
    ct.value = titulo_banner
    ct.fill = FILL_HEADER
    ct.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    ct.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 30

    # As 13 primeiras colunas reproduzem exatamente o modelo CONTROLE SLA - CSC.xlsx
    headers = [
        "Origem", "Criacao", "Prazo Atendimento", "Data de Vencimento", "Situação",
        "Identificador", "Status", "Data Ult Tramitacao", "Tipo de solicitacao",
        "Bandeira", "Loja do Colaborador", "Matricula", "Nome completo do colaborador",
        "Detalhamento SLA (Margem 3d)", "Ação Recomendada RH"
    ]
    ws.append(headers)
    ws.row_dimensions[2].height = 25
    for col in range(1, 16):
        c = ws.cell(row=2, column=col)
        c.fill = FILL_HEADER_SEC
        c.font = FONT_CABECALHO
        c.alignment = ALIGN_CENTER
        c.border = BORDER_THIN

    for r_idx, row_data in enumerate(linhas, start=3):
        ws.append(row_data)
        ws.row_dimensions[r_idx].height = 20
        fill_z = FILL_ZEBRA if (r_idx % 2 == 0) else FILL_BRANCO
        sit = row_data[4]
        sub_st = row_data[13]

        for col in range(1, 16):
            c = ws.cell(row=r_idx, column=col)
            c.border = BORDER_THIN
            c.fill = fill_z
            c.font = FONT_DADOS

            if col in (2, 3, 4, 6, 8, 10, 11, 12):
                c.alignment = ALIGN_CENTER
                if col in (4, 6):
                    c.font = FONT_DADOS_BOLD
            elif col == 5:  # Situação (Vencido / Em dia)
                c.alignment = ALIGN_CENTER
                if sit == "Vencido":
                    c.fill = FILL_CRIT
                    c.font = FONT_CRIT
                else:
                    c.fill = FILL_REG
                    c.font = FONT_REG
            elif col == 14:  # Detalhamento SLA
                c.alignment = ALIGN_CENTER
                if "Expirado" in sub_st:
                    c.fill = FILL_CRIT
                    c.font = FONT_CRIT
                elif "Margem" in sub_st:
                    c.fill = FILL_ALERT
                    c.font = FONT_ALERT
                else:
                    c.fill = FILL_REG
                    c.font = FONT_REG
            else:
                c.alignment = ALIGN_LEFT

    ws.auto_filter.ref = f"A2:O{max(2, len(linhas) + 2)}"

    larguras = {
        "A": 24, "B": 13, "C": 18, "D": 18, "E": 14, "F": 15, "G": 26,
        "H": 19, "I": 38, "J": 16, "K": 18, "L": 14, "M": 34, "N": 28, "O": 58
    }
    for col_l, w in larguras.items():
        ws.column_dimensions[col_l].width = w

# ======================================================================================
# 6. EXECUÇÃO PRINCIPAL
# ======================================================================================

def main():
    t0 = time.time()
    print("=" * 85)
    print("CENCOSUD BRASIL - MÓDULO 8: CONTROLE E AUDITORIA DE SLA — PORTAL CSC")
    print("=" * 85)

    linhas_geral, linhas_divergencias, kpis, stats_origem, stats_bandeira, stats_tipo = processar_relatorios_csc()

    print("\n-> Gerando Workbook Excel Oficial no Padrão Cencosud de 4 Abas...")
    wb = openpyxl.Workbook()
    default_ws = wb.active

    ws_ori = wb.create_sheet(title="Orientações")
    construir_aba_orientacoes(ws_ori)

    ws_res = wb.create_sheet(title="Resumo da Auditoria")
    construir_aba_resumo(ws_res, kpis, stats_origem, stats_bandeira, stats_tipo)

    ws_conf = wb.create_sheet(title="Conferência Geral")
    gravar_aba_chamados(
        ws_conf,
        linhas_geral,
        f"CENCOSUD BRASIL — CONSOLIDADO GERAL DE SLA DO PORTAL CSC ({len(linhas_geral)} CHAMADOS PENDENTES)"
    )

    ws_div = wb.create_sheet(title="Divergências")
    gravar_aba_chamados(
        ws_div,
        linhas_divergencias,
        f"CENCOSUD BRASIL — FILA PRIORITÁRIA DE CHAMADOS VENCIDOS E NA MARGEM DE 3 DIAS ({len(linhas_divergencias)} CHAMADOS)"
    )

    wb.remove(default_ws)

    print(f"-> Salvando relatório em:\n   {OUTPUT_FILE}")
    wb.save(OUTPUT_FILE)

    print("=" * 85)
    print(f"CONCLUÍDO COM SUCESSO EM {time.time() - t0:.2f} SEGUNDOS!")
    print(f" - Total Chamados Ativos:          {kpis['total_ativos']} ({kpis['descartados']} finalizados descartados)")
    print(f" - Status 'Vencido' (c/ margem):   {kpis['vencidos_total']} ({kpis['vencidos_expirados']} Expirados + {kpis['vencidos_margem_3d']} na Margem de 3 dias)")
    print(f" - Status 'Em dia':                {kpis['em_dia']} ({kpis['em_dia']/kpis['total_ativos']:.1%})")
    print("=" * 85)

if __name__ == "__main__":
    main()
