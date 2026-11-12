"""
Módulo de Consolidação em Excel - Robô Ticket Plus (GoodManager)

Responsável por:
1. Extrair os dados cadastrais, financeiros e de limites dos relatórios em PDF
   (Mês Vigente e Compras Próximo Período) gerados pelo robô.
2. Identificar ocorrências de colaboradores não localizados ou sem cartão ativo
   a partir da leitura do log de execução (execucao.log) e arquivos de matrículas.
3. Gerar a planilha 'Resumo_Extratos_Ticket.xlsx' com formatação profissional
   institucional (Cencosud / Good Card), contendo:
   - Aba 'Importacao_Valores': Visão direta para importação/conferência contábil e de folha.
   - Aba 'Detalhamento_Completo': Auditoria detalhada com limites, saldos, cartões e emissão.
"""

import os
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Any

import pymupdf  # PyMuPDF
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from utils import BASE_DIR, OUTPUT_DIR, LOG_FILE, get_logger, load_matriculas, load_matriculas_e_nomes

logger = get_logger("ConsolidadorExcel")
ARQUIVO_EXCEL_PADRAO = BASE_DIR / "Resumo_Extratos_Ticket.xlsx"


def parse_moeda_br(val: Any) -> float:
    """
    Converte uma string monetária brasileira ('331,45', 'R$ 1.250,00') em float decimal (331.45).
    """
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return 0.0
    cleaned = re.sub(r"[^\d,\.-]", "", s)
    # Formato brasileiro com milhar em ponto e decimal em vírgula
    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    elif "," in cleaned:
        cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def extrair_dados_pdf_mes_atual(caminho_pdf: Path) -> Dict[str, Any]:
    """
    Extrai informações do PDF do Mês Atual (Transações Faturadas).
    """
    dados = {
        "competencia": "Outubro/2026",
        "nome": "",
        "cartao": "",
        "situacao_cartao": "",
        "valor_mes_atual": 0.0,
        "limite_geral": 0.0,
        "valor_utilizado": 0.0,
        "saldo_disponivel": 0.0,
        "data_emissao": "",
    }
    try:
        doc = pymupdf.open(caminho_pdf)
        texto = "\n".join(pagina.get_text() for pagina in doc)

        # Competência do título (ex: EXTRATO DE COMPRAS DO USUÁRIO - OUTUBRO/2026)
        m_comp = re.search(r"EXTRATO DE COMPRAS DO USU[ÁA]RIO\s*-\s*([A-ZÇÃÕ]+/\d{4})", texto, re.I)
        if m_comp:
            dados["competencia"] = m_comp.group(1).title()

        # Nome
        m_nome = re.search(r"Nome:\s*(.+)", texto)
        if m_nome:
            dados["nome"] = m_nome.group(1).strip()

        # Cartão
        m_card = re.search(r"Cart[ãa]o:\s*([\d\*\s]+)", texto)
        if m_card:
            dados["cartao"] = m_card.group(1).strip()

        # Situação do Cartão
        m_sit = re.search(r"Situa[çc][ãa]o do cart[ãa]o:\s*(.+)", texto, re.I)
        if m_sit:
            dados["situacao_cartao"] = m_sit.group(1).strip()

        # Total Faturado
        m_fat = re.search(r"Total j[áa] Faturado:\s*([\d\.,]+)", texto, re.I) or re.search(r"Total no Per[íi]odo:\s*([\d\.,]+)", texto, re.I)
        if m_fat:
            dados["valor_mes_atual"] = parse_moeda_br(m_fat.group(1))

        # Limite Geral
        m_lim = re.search(r"Limite de cr[ée]dito GERAL:\s*([\d\.,]+)", texto, re.I)
        if m_lim:
            dados["limite_geral"] = parse_moeda_br(m_lim.group(1))

        # Valor Total Utilizado
        m_util = re.search(r"Valor total utilizado:\s*([\d\.,]+)", texto, re.I)
        if m_util:
            dados["valor_utilizado"] = parse_moeda_br(m_util.group(1))

        # Saldo Disponível
        m_saldo = re.search(r"Saldo:\s*([\d\.,]+)", texto, re.I)
        if m_saldo:
            dados["saldo_disponivel"] = parse_moeda_br(m_saldo.group(1))

        # Data de Emissão
        m_emissao = re.search(r"Emiss[ãa]o:\s*([\d/:\s]+)", texto, re.I)
        if m_emissao:
            dados["data_emissao"] = m_emissao.group(1).strip()

    except Exception as e:
        logger.warning(f"Erro ao ler PDF do Mês Atual {caminho_pdf.name}: {e}")

    return dados


def extrair_dados_pdf_proximo_periodo(caminho_pdf: Path) -> Dict[str, Any]:
    """
    Extrai informações do PDF de Compras Próximo Período (Transações Não Faturadas).
    """
    dados = {
        "valor_proximo_periodo": 0.0,
        "parcelas_proximos": 0.0,
        "nome": "",
        "cartao": "",
        "limite_geral": 0.0,
        "valor_utilizado": 0.0,
        "saldo_disponivel": 0.0,
        "data_emissao": "",
    }
    try:
        doc = pymupdf.open(caminho_pdf)
        texto = "\n".join(pagina.get_text() for pagina in doc)

        # Saldo a faturar
        m_fat = re.search(r"SALDO A FATURAR:\s*([\d\.,]+)", texto, re.I) or re.search(r"Total no Per[íi]odo:\s*([\d\.,]+)", texto, re.I)
        if m_fat:
            dados["valor_proximo_periodo"] = parse_moeda_br(m_fat.group(1))

        # Parcelas próximos períodos
        m_parc = re.search(r"Parcelas pr[óo]ximos per[íi]odos:\s*([\d\.,]+)", texto, re.I)
        if m_parc:
            dados["parcelas_proximos"] = parse_moeda_br(m_parc.group(1))

        # Nome
        m_nome = re.search(r"Nome:\s*(.+)", texto)
        if m_nome:
            dados["nome"] = m_nome.group(1).strip()

        # Cartão
        m_card = re.search(r"Cart[ãa]o:\s*([\d\*\s]+)", texto)
        if m_card:
            dados["cartao"] = m_card.group(1).strip()

        # Limite Geral
        m_lim = re.search(r"Limite de cr[ée]dito GERAL:\s*([\d\.,]+)", texto, re.I)
        if m_lim:
            dados["limite_geral"] = parse_moeda_br(m_lim.group(1))

        # Valor Total Utilizado
        m_util = re.search(r"Valor total utilizado:\s*([\d\.,]+)", texto, re.I)
        if m_util:
            dados["valor_utilizado"] = parse_moeda_br(m_util.group(1))

        # Saldo Disponível
        m_saldo = re.search(r"Saldo:\s*([\d\.,]+)", texto, re.I)
        if m_saldo:
            dados["saldo_disponivel"] = parse_moeda_br(m_saldo.group(1))

        # Emissão
        m_emissao = re.search(r"Emiss[ãa]o:\s*([\d/:\s]+)", texto, re.I)
        if m_emissao:
            dados["data_emissao"] = m_emissao.group(1).strip()

    except Exception as e:
        logger.warning(f"Erro ao ler PDF do Próximo Período {caminho_pdf.name}: {e}")

    return dados


def mapear_falhas_do_log() -> Dict[str, Dict[str, str]]:
    """
    Inspeciona execucao.log para resgatar o histórico de matrículas com exceções/falhas.
    """
    falhas = {}
    if not LOG_FILE.exists():
        return falhas

    try:
        with open(LOG_FILE, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                # 1. Matrícula não localizada
                m_nl = re.search(r"\[Matr[íi]cula\s+(\d+)\]\s+N[ãa]o localizada", line, re.I)
                if m_nl:
                    matr = m_nl.group(1).strip()
                    falhas[matr] = {
                        "matricula": matr,
                        "nome": "NÃO LOCALIZADO",
                        "cartao": "-",
                        "situacao_cartao": "-",
                        "status": "Matrícula não localizada",
                    }

                # 2. Usuário Cancelado / Sem cartão Ativo (A) e Titular (T)
                m_no_card = re.search(
                    r"\[Matr[íi]cula\s+(\d+)\]\s+Nenhum cart[ãa]o com '\(A\)' e '\(T\)'.+Op[çc][õo]es listadas:\s*\[(.*)\]",
                    line,
                    re.I,
                )
                if m_no_card:
                    matr = m_no_card.group(1).strip()
                    opts = m_no_card.group(2)
                    nome = ""
                    m_nome = re.search(r"\([A-Z]\)\s*\([A-Z]\)\s+([A-Z\s]+?)\s+\d+", opts)
                    if m_nome:
                        nome = m_nome.group(1).strip()

                    status = (
                        "Usuário Cancelado (U) - Sem cartão Ativo (A)"
                        if "(U)" in opts
                        else "Sem cartão Ativo/Titular"
                    )
                    falhas[matr] = {
                        "matricula": matr,
                        "nome": nome or "NÃO LOCALIZADO",
                        "cartao": "-",
                        "situacao_cartao": "USUÁRIO CANCELADO" if "(U)" in opts else "SEM CARTÃO ATIVO",
                        "status": status,
                    }
    except Exception as e:
        logger.warning(f"Erro ao mapear falhas do log: {e}")

    return falhas


def coletar_todos_os_dados() -> List[Dict[str, Any]]:
    """
    Percorre a pasta 'Relatórios extraídos' e consolida todas as matrículas com PDF,
    mais as matrículas identificadas com exceção.
    """
    logger.info("Coletando dados dos PDFs e logs de execução...")
    registros: Dict[str, Dict[str, Any]] = {}

    # 1. Extração a partir das pastas de colaboradores em 'Relatórios extraídos'
    if OUTPUT_DIR.exists():
        for pasta in sorted(OUTPUT_DIR.iterdir()):
            if not pasta.is_dir():
                continue

            # Nome da pasta geralmente é 'MATRICULA - NOME'
            partes = pasta.name.split(" - ", 1)
            matr_pasta = partes[0].strip()
            nome_pasta = partes[1].strip() if len(partes) > 1 else ""

            pdfs = list(pasta.glob("*.pdf"))
            pdf_mes_atual = next((p for p in pdfs if "Mes_Atual" in p.name), None)
            pdf_prox_periodo = next((p for p in pdfs if "Compras_Proximo_Periodo" in p.name), None)

            dados_mes = extrair_dados_pdf_mes_atual(pdf_mes_atual) if pdf_mes_atual else {}
            dados_prox = extrair_dados_pdf_proximo_periodo(pdf_prox_periodo) if pdf_prox_periodo else {}

            nome = dados_mes.get("nome") or dados_prox.get("nome") or nome_pasta
            cartao = dados_mes.get("cartao") or dados_prox.get("cartao") or "-"
            situacao = dados_mes.get("situacao_cartao") or "CARTAO ATIVO"
            competencia = dados_mes.get("competencia") or "Outubro/2026"
            val_mes = dados_mes.get("valor_mes_atual", 0.0)
            val_prox = dados_prox.get("valor_proximo_periodo", 0.0)
            total_2_comp = val_mes + val_prox
            parc_prox = dados_prox.get("parcelas_proximos", 0.0)
            lim_geral = dados_mes.get("limite_geral") or dados_prox.get("limite_geral", 0.0)
            val_util = dados_mes.get("valor_utilizado") or dados_prox.get("valor_utilizado", 0.0)
            saldo = dados_mes.get("saldo_disponivel") or dados_prox.get("saldo_disponivel", 0.0)
            emissao = dados_mes.get("data_emissao") or dados_prox.get("data_emissao", "-")

            registros[matr_pasta] = {
                "matricula": matr_pasta,
                "nome": nome,
                "cartao": cartao,
                "situacao_cartao": situacao,
                "competencia": competencia,
                "valor_mes_atual": val_mes,
                "valor_proximo_periodo": val_prox,
                "parcelas_proximos": parc_prox,
                "total_2_competencias": total_2_comp,
                "limite_geral": lim_geral,
                "valor_utilizado": val_util,
                "saldo_disponivel": saldo,
                "data_emissao": emissao,
                "status": "Extraído com Sucesso",
            }

    # 2. Adiciona registros de exceções/falhas conhecidas do log
    falhas_log = mapear_falhas_do_log()
    for matr, falha in falhas_log.items():
        if matr not in registros:
            registros[matr] = {
                "matricula": matr,
                "nome": falha.get("nome", "NÃO LOCALIZADO"),
                "cartao": falha.get("cartao", "-"),
                "situacao_cartao": falha.get("situacao_cartao", "-"),
                "competencia": "Outubro/2026",
                "valor_mes_atual": 0.0,
                "valor_proximo_periodo": 0.0,
                "parcelas_proximos": 0.0,
                "total_2_competencias": 0.0,
                "limite_geral": 0.0,
                "valor_utilizado": 0.0,
                "saldo_disponivel": 0.0,
                "data_emissao": "-",
                "status": falha.get("status", "Não localizado"),
            }

    # 3. Lê lista oficial e mapa de nomes da planilha de desligados
    matr_oficiais, mapa_nomes = load_matriculas_e_nomes()
    for matr in matr_oficiais:
        if matr not in registros:
            registros[matr] = {
                "matricula": matr,
                "nome": mapa_nomes.get(matr, "NÃO PROCESSADO"),
                "cartao": "-",
                "situacao_cartao": "-",
                "competencia": "Outubro/2026",
                "valor_mes_atual": 0.0,
                "valor_proximo_periodo": 0.0,
                "parcelas_proximos": 0.0,
                "total_2_competencias": 0.0,
                "limite_geral": 0.0,
                "valor_utilizado": 0.0,
                "saldo_disponivel": 0.0,
                "data_emissao": "-",
                "status": "Pendente de Execução",
            }

    # Atualiza nomes a partir do mapa oficial caso esteja vago ou não localizado
    for matr, reg in registros.items():
        if matr in mapa_nomes and (not reg["nome"] or reg["nome"] in ["NÃO LOCALIZADO", "NÃO PROCESSADO", "SEM_NOME"]):
            reg["nome"] = mapa_nomes[matr]

    # Ordenação: segue a ordem original da planilha oficial de desligados,
    # seguida por quaisquer matrículas adicionais (ex: testes anteriores)
    ordem_oficial = {matr: idx for idx, matr in enumerate(matr_oficiais)}

    def sort_key(item):
        m = item["matricula"]
        return ordem_oficial.get(m, 100000 + (int(m) if m.isdigit() else 99999))

    lista_ordenada = sorted(registros.values(), key=sort_key)
    logger.info(f"Total de registros consolidados: {len(lista_ordenada)} (Base oficial: {len(matr_oficiais)})")
    return lista_ordenada


def gerar_consolidado_excel(
    caminho_saida: Optional[Path] = None,
    dados: Optional[List[Dict[str, Any]]] = None,
) -> Path:
    """
    Gera a planilha Excel 'Resumo_Extratos_Ticket.xlsx' formatada profissionalmente
    com openpyxl contendo duas abas: 'Importacao_Valores' e 'Detalhamento_Completo'.
    """
    if caminho_saida is None:
        caminho_saida = ARQUIVO_EXCEL_PADRAO

    if dados is None:
        dados = coletar_todos_os_dados()
    wb = openpyxl.Workbook()

    # Estilos Padrão Institucional Cencosud / Good Card
    AZUL_HEADER = "1F4E78"       # Azul institucional corporativo
    AZUL_TOTAL = "D9E1F2"        # Azul claro de destaque para linha de totais
    VERDE_SUCESSO = "E2EFDA"     # Fundo suave para status de sucesso
    VERMELHO_FALHA = "FCE4D6"    # Fundo suave para exceção/falha

    fonte_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    fonte_corpo = Font(name="Segoe UI", size=10, color="000000")
    fonte_total = Font(name="Segoe UI", size=11, bold=True, color="000000")
    fonte_alerta = Font(name="Segoe UI", size=10, bold=True, color="C00000")
    fonte_ok = Font(name="Segoe UI", size=10, bold=True, color="375623")

    fill_header = PatternFill(start_color=AZUL_HEADER, end_color=AZUL_HEADER, fill_type="solid")
    fill_total = PatternFill(start_color=AZUL_TOTAL, end_color=AZUL_TOTAL, fill_type="solid")
    fill_sucesso = PatternFill(start_color=VERDE_SUCESSO, end_color=VERDE_SUCESSO, fill_type="solid")
    fill_falha = PatternFill(start_color=VERMELHO_FALHA, end_color=VERMELHO_FALHA, fill_type="solid")

    alinhamento_centro = Alignment(horizontal="center", vertical="center")
    alinhamento_esquerda = Alignment(horizontal="left", vertical="center")
    alinhamento_direita = Alignment(horizontal="right", vertical="center")

    borda_fina = Side(style="thin", color="D9D9D9")
    borda_celula = Border(left=borda_fina, right=borda_fina, top=borda_fina, bottom=borda_fina)
    borda_dupla_total = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))

    FORMATO_MOEDA = 'R$ #,##0.00'

    # ==========================================================
    # ABA 1: Importacao_Valores (Enxuta e Direta para Conferência)
    # ==========================================================
    ws1 = wb.active
    ws1.title = "Importacao_Valores"
    ws1.views.sheetView[0].showGridLines = True

    cabecalhos_aba1 = [
        "Matrícula",
        "Nome do Colaborador",
        "Valor Mês Atual (Outubro/2026)",
        "Valor Próximo Período",
        "Total das 2 Competências",
        "Status / Observação",
    ]

    ws1.append(cabecalhos_aba1)
    ws1.row_dimensions[1].height = 28

    for col_idx in range(1, len(cabecalhos_aba1) + 1):
        cell = ws1.cell(row=1, column=col_idx)
        cell.font = fonte_header
        cell.fill = fill_header
        cell.alignment = alinhamento_centro

    for r_idx, reg in enumerate(dados, start=2):
        ws1.row_dimensions[r_idx].height = 20
        row_values = [
            reg["matricula"],
            reg["nome"],
            reg["valor_mes_atual"],
            reg["valor_proximo_periodo"],
            reg["total_2_competencias"],
            reg["status"],
        ]
        ws1.append(row_values)

        # Formatação das células
        c_matr = ws1.cell(row=r_idx, column=1)
        c_matr.alignment = alinhamento_centro
        c_matr.font = fonte_corpo
        c_matr.border = borda_celula

        c_nome = ws1.cell(row=r_idx, column=2)
        c_nome.alignment = alinhamento_esquerda
        c_nome.font = fonte_corpo
        c_nome.border = borda_celula

        # Valores monetários (Colunas 3, 4, 5)
        for c_idx in [3, 4, 5]:
            c_val = ws1.cell(row=r_idx, column=c_idx)
            c_val.number_format = FORMATO_MOEDA
            c_val.alignment = alinhamento_direita
            c_val.font = fonte_corpo
            c_val.border = borda_celula

        # Status (Coluna 6)
        c_stat = ws1.cell(row=r_idx, column=6)
        c_stat.alignment = alinhamento_centro
        c_stat.border = borda_celula
        if reg["status"] == "Extraído com Sucesso":
            c_stat.font = fonte_ok
            c_stat.fill = fill_sucesso
        else:
            c_stat.font = fonte_alerta
            c_stat.fill = fill_falha

    # Linha de Totais da Aba 1
    total_row_aba1 = len(dados) + 2
    ws1.row_dimensions[total_row_aba1].height = 24
    ws1.cell(row=total_row_aba1, column=1, value="TOTAL GERAL").alignment = alinhamento_centro
    ws1.cell(row=total_row_aba1, column=1).font = fonte_total
    ws1.cell(row=total_row_aba1, column=1).fill = fill_total
    ws1.cell(row=total_row_aba1, column=1).border = borda_dupla_total

    ws1.cell(row=total_row_aba1, column=2, value=f"{len(dados)} Colaboradores").alignment = alinhamento_esquerda
    ws1.cell(row=total_row_aba1, column=2).font = fonte_total
    ws1.cell(row=total_row_aba1, column=2).fill = fill_total
    ws1.cell(row=total_row_aba1, column=2).border = borda_dupla_total

    for c_idx, col_letter in [(3, "C"), (4, "D"), (5, "E")]:
        c_tot = ws1.cell(row=total_row_aba1, column=c_idx)
        c_tot.value = f"=SUM({col_letter}2:{col_letter}{total_row_aba1 - 1})"
        c_tot.number_format = FORMATO_MOEDA
        c_tot.alignment = alinhamento_direita
        c_tot.font = fonte_total
        c_tot.fill = fill_total
        c_tot.border = borda_dupla_total

    c_tot_stat = ws1.cell(row=total_row_aba1, column=6, value="")
    c_tot_stat.fill = fill_total
    c_tot_stat.border = borda_dupla_total

    ws1.auto_filter.ref = f"A1:F{total_row_aba1 - 1}"

    # Auto ajuste de largura das colunas
    for col in ws1.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws1.column_dimensions[col_letter].width = max(max_len + 4, 14)
    ws1.column_dimensions["A"].width = 14
    ws1.column_dimensions["B"].width = 38
    ws1.column_dimensions["C"].width = 30
    ws1.column_dimensions["D"].width = 24
    ws1.column_dimensions["E"].width = 28
    ws1.column_dimensions["F"].width = 44

    # ==========================================================
    # ABA 2: Detalhamento_Completo (Limites, Saldos e Auditoria)
    # ==========================================================
    ws2 = wb.create_sheet(title="Detalhamento_Completo")
    ws2.views.sheetView[0].showGridLines = True

    cabecalhos_aba2 = [
        "Matrícula",
        "Nome do Colaborador",
        "Número do Cartão",
        "Situação Cartão",
        "Competência Vigente",
        "Valor Mês Atual (Faturado)",
        "Valor Próximo Período (A Faturar)",
        "Parcelas Próx. Períodos",
        "Total das 2 Competências",
        "Limite de Crédito Geral",
        "Valor Total Utilizado",
        "Saldo Disponível",
        "Data de Emissão",
        "Status / Observação",
    ]

    ws2.append(cabecalhos_aba2)
    ws2.row_dimensions[1].height = 28

    for col_idx in range(1, len(cabecalhos_aba2) + 1):
        cell = ws2.cell(row=1, column=col_idx)
        cell.font = fonte_header
        cell.fill = fill_header
        cell.alignment = alinhamento_centro

    for r_idx, reg in enumerate(dados, start=2):
        ws2.row_dimensions[r_idx].height = 20
        row_values = [
            reg["matricula"],
            reg["nome"],
            reg["cartao"],
            reg["situacao_cartao"],
            reg["competencia"],
            reg["valor_mes_atual"],
            reg["valor_proximo_periodo"],
            reg["parcelas_proximos"],
            reg["total_2_competencias"],
            reg["limite_geral"],
            reg["valor_utilizado"],
            reg["saldo_disponivel"],
            reg["data_emissao"],
            reg["status"],
        ]
        ws2.append(row_values)

        # Formatações
        ws2.cell(row=r_idx, column=1).alignment = alinhamento_centro  # Matrícula
        ws2.cell(row=r_idx, column=2).alignment = alinhamento_esquerda  # Nome
        ws2.cell(row=r_idx, column=3).alignment = alinhamento_centro  # Cartão
        ws2.cell(row=r_idx, column=4).alignment = alinhamento_centro  # Situação
        ws2.cell(row=r_idx, column=5).alignment = alinhamento_centro  # Competência

        # Colunas Monetárias: 6, 7, 8, 9, 10, 11, 12
        for c_idx in [6, 7, 8, 9, 10, 11, 12]:
            c_val = ws2.cell(row=r_idx, column=c_idx)
            c_val.number_format = FORMATO_MOEDA
            c_val.alignment = alinhamento_direita
            c_val.font = fonte_corpo
            c_val.border = borda_celula

        # Data Emissão e Status
        ws2.cell(row=r_idx, column=13).alignment = alinhamento_centro
        c_stat = ws2.cell(row=r_idx, column=14)
        c_stat.alignment = alinhamento_centro
        c_stat.border = borda_celula
        if reg["status"] == "Extraído com Sucesso":
            c_stat.font = fonte_ok
            c_stat.fill = fill_sucesso
        else:
            c_stat.font = fonte_alerta
            c_stat.fill = fill_falha

        for c_idx in range(1, len(cabecalhos_aba2) + 1):
            ws2.cell(row=r_idx, column=c_idx).border = borda_celula

    # Linha de Totais da Aba 2
    total_row_aba2 = len(dados) + 2
    ws2.row_dimensions[total_row_aba2].height = 24
    ws2.cell(row=total_row_aba2, column=1, value="TOTAL GERAL").alignment = alinhamento_centro
    ws2.cell(row=total_row_aba2, column=1).font = fonte_total
    ws2.cell(row=total_row_aba2, column=1).fill = fill_total
    ws2.cell(row=total_row_aba2, column=1).border = borda_dupla_total

    ws2.cell(row=total_row_aba2, column=2, value=f"{len(dados)} Registros").alignment = alinhamento_esquerda
    ws2.cell(row=total_row_aba2, column=2).font = fonte_total
    ws2.cell(row=total_row_aba2, column=2).fill = fill_total
    ws2.cell(row=total_row_aba2, column=2).border = borda_dupla_total

    for c_idx in range(3, 6):
        c_empty = ws2.cell(row=total_row_aba2, column=c_idx, value="")
        c_empty.fill = fill_total
        c_empty.border = borda_dupla_total

    # Fórmulas de soma para as colunas monetárias
    for c_idx, col_letter in [(6, "F"), (7, "G"), (8, "H"), (9, "I"), (10, "J"), (11, "K"), (12, "L")]:
        c_tot = ws2.cell(row=total_row_aba2, column=c_idx)
        c_tot.value = f"=SUM({col_letter}2:{col_letter}{total_row_aba2 - 1})"
        c_tot.number_format = FORMATO_MOEDA
        c_tot.alignment = alinhamento_direita
        c_tot.font = fonte_total
        c_tot.fill = fill_total
        c_tot.border = borda_dupla_total

    for c_idx in [13, 14]:
        c_empty = ws2.cell(row=total_row_aba2, column=c_idx, value="")
        c_empty.fill = fill_total
        c_empty.border = borda_dupla_total

    ws2.auto_filter.ref = f"A1:N{total_row_aba2 - 1}"

    # Auto ajuste de largura das colunas Aba 2
    ws2.column_dimensions["A"].width = 14
    ws2.column_dimensions["B"].width = 38
    ws2.column_dimensions["C"].width = 24
    ws2.column_dimensions["D"].width = 20
    ws2.column_dimensions["E"].width = 22
    ws2.column_dimensions["F"].width = 26
    ws2.column_dimensions["G"].width = 30
    ws2.column_dimensions["H"].width = 24
    ws2.column_dimensions["I"].width = 26
    ws2.column_dimensions["J"].width = 22
    ws2.column_dimensions["K"].width = 22
    ws2.column_dimensions["L"].width = 20
    ws2.column_dimensions["M"].width = 22
    ws2.column_dimensions["N"].width = 44

    # Salva o arquivo Excel
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(caminho_saida)
    logger.info(f"Planilha consolidada gerada com sucesso em: {caminho_saida}")
    return caminho_saida


def gerar_resumos_separados_por_planilha() -> List[Path]:
    """
    Gera um arquivo de resumo individual para cada planilha presente na pasta
    'dados e matriculas p consulta', contendo exclusivamente os colaboradores daquela base.
    """
    from utils import DADOS_DIR

    gerados: List[Path] = []
    if not DADOS_DIR.exists():
        return gerados

    planilhas = [
        p for p in DADOS_DIR.glob("*")
        if p.suffix.lower() in [".xlsx", ".xls", ".xlsb"] and not p.name.startswith(("~", "_temp_"))
    ]

    if not planilhas:
        return gerados

    todos_dados = coletar_todos_os_dados()
    mapa_dados = {reg["matricula"]: reg for reg in todos_dados}

    for planilha in sorted(planilhas, key=lambda x: x.name):
        nome_base = planilha.stem
        # Nome do arquivo de saída limpo
        saida_nome = f"Resumo_{nome_base}.xlsx"
        saida_path = BASE_DIR / saida_nome

        mats_arquivo, nomes_arquivo = load_matriculas_e_nomes(planilha)
        logger.info(f"Gerando resumo individual para '{planilha.name}': {len(mats_arquivo)} matrículas -> '{saida_nome}'")

        dados_filtrados = []
        for m in mats_arquivo:
            if m in mapa_dados:
                reg = dict(mapa_dados[m])
                if m in nomes_arquivo and nomes_arquivo[m]:
                    reg["nome"] = nomes_arquivo[m]
                dados_filtrados.append(reg)
            else:
                dados_filtrados.append({
                    "matricula": m,
                    "nome": nomes_arquivo.get(m, "NÃO LOCALIZADO"),
                    "cartao": "-",
                    "situacao_cartao": "-",
                    "competencia": "Outubro/2026",
                    "valor_mes_atual": 0.0,
                    "valor_proximo_periodo": 0.0,
                    "parcelas_proximos": 0.0,
                    "total_2_competencias": 0.0,
                    "limite_geral": 0.0,
                    "valor_utilizado": 0.0,
                    "saldo_disponivel": 0.0,
                    "data_emissao": "-",
                    "status": "Matrícula não localizada",
                })

        gerar_consolidado_excel(saida_path, dados_filtrados)
        gerados.append(saida_path)

    return gerados


if __name__ == "__main__":
    caminho = gerar_consolidado_excel()
    print(f"Planilha consolidada geral: {caminho}")
    separadas = gerar_resumos_separados_por_planilha()
    for s in separadas:
        print(f"Planilha individual gerada: {s.name}")
