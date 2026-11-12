# -*- coding: utf-8 -*-
"""
processar_auditoria_inss.py
---------------------------
Motor de Auditoria e Cruzamento de Afastados (INSS Empresa x Folha SAP Cencosud).
Cruza os extratos oficiais do portal empresa.inss.gov.br com a base cadastral COLA_Colaboradores.
Gera relatório executivo com dashboard e alertas de conformidade.
"""

import os
import re
import datetime
import zipfile
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pypdfium2 as pdfium

def parse_inss_pdf(pdf_path):
    """Lê e extrai todos os registros do extrato em PDF do INSS Empresa."""
    if not os.path.exists(pdf_path):
        print(f"Arquivo não encontrado: {pdf_path}")
        return []

    pdf = pdfium.PdfDocument(pdf_path)
    full_text = []
    for page in pdf:
        full_text.append(page.get_textpage().get_text_range())
    raw_content = "\n".join(full_text)

    # Identificar CNPJ do cabeçalho
    current_cnpj = "NÃO IDENTIFICADO"
    cnpj_match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", raw_content)
    if cnpj_match:
        current_cnpj = cnpj_match.group(1)

    lines = [l.strip() for l in raw_content.splitlines() if l.strip()]
    records = []
    current_especie = "NÃO INFORMADA"

    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        
        # Mudar de espécie
        if re.match(r"^\d{2}\s*-\s*[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ\s\-]+", line) and ("AUXÍLIO" in line or "APOSENTADORIA" in line or "PENSÃO" in line):
            current_especie = line
            i += 1
            continue

        # Linha de cabeçalho da pessoa: NOME CPF NIT DATA NASCIMENTO
        m_person = re.search(r"^([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ'\s]+?)\s+(\d{3}\.\d{3}\.\d{3}-\d{2})\s+(\d{3}\.\d{5}\.\d{2}-\d)\s+(\d{2}/\d{2}/\d{4})$", line)
        if m_person:
            nome = m_person.group(1).strip()
            nome = re.sub(r"^(NOME\s+CPF\s+NIT\s+DATA\s+NASCIMENTO\s+)+", "", nome).strip()
            cpf = m_person.group(2)
            nit = m_person.group(3)
            data_nasc = m_person.group(4)

            # Procurar bloco do benefício
            j = i + 1
            nb, situacao, dt_req, dt_ini, dt_desp, dt_pericia, conclusao, dt_limite, dt_cessacao = "", "", "", "", "", "", "", "", ""
            while j < min(i + 25, n):
                line_j = lines[j]
                m_benef = re.search(r"(\d{3}\.\d{3}\.\d{3}-\d)\s+(ATIVO|CESSADO)\b", line_j)
                if m_benef:
                    nb = m_benef.group(1)
                    situacao = m_benef.group(2)

                    # Coletar apenas até antes da próxima pessoa ou rodapé
                    block_lines = []
                    for k in range(j, min(j+10, n)):
                        lk = lines[k]
                        if k > j and ("NOME CPF NIT" in lk or "INSS - INSTITUTO" in lk or "BENEFICIÁRIOS POR" in lk):
                            break
                        # Se a própria linha tiver "NOME CPF NIT" no meio, truncar
                        if "NOME CPF NIT" in lk:
                            lk_cut = lk.split("NOME CPF NIT")[0]
                            block_lines.append(lk_cut)
                            break
                        block_lines.append(lk)

                    block_text = " ".join(block_lines)
                    all_dates = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", block_text)

                    if "Data de cessação do benefício" in block_text:
                        conclusao = "Data de cessação do benefício"
                    elif "Contrário" in block_text or "Contrria" in block_text or "Contrária" in block_text:
                        conclusao = "Contrário"
                    elif "Favorável" in block_text or "Favorvel" in block_text or "Favorável" in block_text:
                        conclusao = "Favorável"
                    else:
                        conclusao = "Avaliação Realizada"

                    # Datas válidas que não sejam a data de nascimento da pessoa
                    valid_dates = [d for d in all_dates if d != data_nasc]
                    if len(valid_dates) >= 1: dt_req = valid_dates[0]
                    if len(valid_dates) >= 2: dt_ini = valid_dates[1]
                    if len(valid_dates) >= 3: dt_desp = valid_dates[2]
                    if len(valid_dates) >= 4: dt_pericia = valid_dates[3]
                    
                    # Na estrutura do INSS, a Data de Cessação é a última data válida do bloco
                    if len(valid_dates) >= 6:
                        dt_limite = valid_dates[4]
                        dt_cessacao = valid_dates[5]
                    elif len(valid_dates) == 5:
                        dt_cessacao = valid_dates[4]
                    elif len(valid_dates) == 4 and situacao == "CESSADO":
                        dt_cessacao = valid_dates[3]

                    break
                j += 1

            records.append({
                "CNPJ_INSS": current_cnpj,
                "Especie_INSS": current_especie,
                "Nome_INSS": nome,
                "Nome_norm": re.sub(r"\s+", " ", nome).strip().upper(),
                "CPF_INSS": cpf,
                "CPF_clean": re.sub(r"\D", "", cpf).zfill(11),
                "NIT_INSS": nit,
                "Data_Nasc_INSS": data_nasc,
                "NB": nb,
                "Situacao_INSS": situacao,
                "Data_Req_INSS": dt_req,
                "Data_Inicio_DIB": dt_ini,
                "Data_Despacho_INSS": dt_desp,
                "Data_Pericia_INSS": dt_pericia,
                "Conclusao_Pericia": conclusao,
                "Data_Limite_INSS": dt_limite,
                "Data_Cessacao_DCB": dt_cessacao
            })
            i = max(i, j)
        i += 1

    return records


def carregar_base_colaboradores():
    """Carrega a base COLA_Colaboradores da Cencosud do zip mais recente ou pasta tratada."""
    # 1. Tentar pegar o zip mais recente em Downloads
    zip_path = r"c:\Users\matcaval\Downloads\Relatórios Nativos LG_10.2026 e 11.2026.zip"
    if os.path.exists(zip_path):
        with zipfile.ZipFile(zip_path, 'r') as z:
            for name in z.namelist():
                if "COLA_Colaboradores_CSV" in name:
                    with z.open(name) as f:
                        df = pd.read_csv(f, sep=';', skiprows=1, encoding='latin1', low_memory=False)
                        df['Nome_norm'] = df['Colaborador'].astype(str).str.strip().str.upper()
                        cpf_col = [c for c in df.columns if 'CPF' in c][0]
                        df['CPF_clean'] = df[cpf_col].astype(str).str.replace(r'\D', '', regex=True).str.zfill(11)
                        return df

    # 2. Fallback para Excel tratado
    excel_path = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\04 - GOOD CARD\Relatório Tratado\COLA_Colaboradores.xlsx"
    if os.path.exists(excel_path):
        df = pd.read_excel(excel_path)
        df['Nome_norm'] = df['Colaborador'].astype(str).str.strip().str.upper()
        cpf_col = [c for c in df.columns if 'CPF' in c][0]
        df['CPF_clean'] = df[cpf_col].astype(str).str.replace(r'\D', '', regex=True).str.zfill(11)
        return df

    raise FileNotFoundError("Não foi possível localizar a base COLA_Colaboradores.")


def diagnosticar_registro(row):
    """Aplica as regras de auditoria corporativa e SST."""
    sit_inss = str(row.get("Situacao_INSS", "")).upper()
    sit_folha = str(row.get("Situação", "")).strip()
    especie = str(row.get("Especie_INSS", ""))
    
    is_acidente = "91" in especie or "92" in especie or "ACIDENTE" in especie.upper()
    estabilidade = "SIM (Estabilidade 12 meses após retorno)" if is_acidente else "NÃO"

    if not row.get("Matrícula") or pd.isna(row.get("Matrícula")):
        return (
            "⚠️ NÃO LOCALIZADO NA FOLHA",
            "Colaborador com benefício no INSS não localizado na base ativa do SAP. Checar se foi desligado ou transferido de CNPJ.",
            estabilidade
        )

    # Colaborador está "Ativo" / "Atividade Normal" na Folha
    is_ativo_folha = "ATIVIDADE NORMAL" in sit_folha.upper() or sit_folha.upper() == "ATIVO"

    if is_ativo_folha:
        if sit_inss == "ATIVO":
            return (
                "🚨 CRÍTICO - ATIVO NA FOLHA COM BENEFÍCIO VIGENTE NO INSS",
                "Colaborador consta trabalhando na loja/unidade, mas possui benefício vigente concedido pelo INSS. Risco grave de duplicidade salarial e ausência de ASO de retorno.",
                estabilidade
            )
        else: # CESSADO
            return (
                "✅ CONFORME - RETORNO REALIZADO APÓS CESSAÇÃO",
                "Colaborador consta com benefício cessado no INSS e está ativo na folha. Validar se o ASO de retorno está arquivado no prontuário.",
                estabilidade
            )
    else:
        # Colaborador está como "Afastado" na Folha
        if sit_inss == "ATIVO":
            return (
                "✅ REGULAR - AFASTADO COM BENEFÍCIO VIGENTE",
                "Afastamento e benefício perfeitamente sincronizados entre INSS e Folha. Acompanhar data limite de perícia.",
                estabilidade
            )
        else: # CESSADO
            return (
                "⚠️ ATENÇÃO - BENEFÍCIO CESSADO NO INSS E AFASTADO NA FOLHA",
                "Data de cessação (DCB) atingida no INSS, mas o colaborador ainda permanece afastado no SAP. Convocar para Exame de Retorno ou solicitar laudo de pedido de prorrogação.",
                estabilidade
            )


def gerar_relatorio_excel(df_resultado, caminho_saida):
    """Gera o arquivo Excel corporativo no padrão visual Cencosud com OpenPyXL."""
    wb = openpyxl.Workbook()
    
    # ----------------------------------------------------
    # Cores e Estilos Corporativos
    # ----------------------------------------------------
    azul_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    azul_kpi = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    vermelho_critico = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
    amarelo_atencao = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    verde_conforme = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    cinza_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_normal = Font(name="Calibri", size=10)
    font_kpi_num = Font(name="Calibri", size=20, bold=True, color="FFFFFF")
    font_kpi_label = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # ----------------------------------------------------
    # ABA 1: RESUMO EXECUTIVO
    # ----------------------------------------------------
    ws_resumo = wb.active
    ws_resumo.title = "Resumo Executivo"
    ws_resumo.views.sheetView[0].showGridLines = True

    # Título Principal
    ws_resumo.merge_cells("B2:K2")
    ws_resumo["B2"] = "CENCOSUD BRASIL - AUDITORIA DE BENEFÍCIOS PREVIDENCIÁRIOS & AFASTADOS (INSS x FOLHA SAP)"
    ws_resumo["B2"].font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    ws_resumo["B2"].alignment = Alignment(vertical="center")

    ws_resumo["B3"] = f"Competência: Outubro/2026 | Gerado em: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws_resumo["B3"].font = Font(name="Calibri", size=10, italic=True, color="595959")

    # Contadores KPI
    total_inss = len(df_resultado)
    total_critico = len(df_resultado[df_resultado["Diagnóstico"].str.contains("CRÍTICO", na=False)])
    total_cessado_pendente = len(df_resultado[df_resultado["Diagnóstico"].str.contains("BENEFÍCIO CESSADO", na=False)])
    total_regular = len(df_resultado[df_resultado["Diagnóstico"].str.contains("CONFORME|REGULAR", na=False)])
    total_acidentario = len(df_resultado[df_resultado["Estabilidade_Acidentária"].str.contains("SIM", na=False)])

    kpis = [
        ("Total Benefícios INSS", total_inss, "2F5597"),
        ("🚨 Crítico (Ativo c/ Benefício)", total_critico, "C00000"),
        ("⚠️ Cessados sem Retorno", total_cessado_pendente, "ED7D31"),
        ("✅ Regulares / Conformes", total_regular, "385723"),
        ("🛡️ Acidentários (B91 Estabilidade)", total_acidentario, "7030A0")
    ]

    col_start = 2
    for label, val, hex_color in kpis:
        c_letter = get_column_letter(col_start)
        c_next = get_column_letter(col_start + 1)
        ws_resumo.merge_cells(f"{c_letter}5:{c_next}5")
        ws_resumo.merge_cells(f"{c_letter}6:{c_next}6")

        fill_k = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
        
        ws_resumo[f"{c_letter}5"] = str(val)
        ws_resumo[f"{c_letter}5"].font = font_kpi_num
        ws_resumo[f"{c_letter}5"].fill = fill_k
        ws_resumo[f"{c_letter}5"].alignment = Alignment(horizontal="center", vertical="center")

        ws_resumo[f"{c_letter}6"] = label
        ws_resumo[f"{c_letter}6"].font = font_kpi_label
        ws_resumo[f"{c_letter}6"].fill = fill_k
        ws_resumo[f"{c_letter}6"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        col_start += 2

    # Tabela 1: Distribuição por Diagnóstico
    ws_resumo["B9"] = "DISTRIBUIÇÃO POR DIAGNÓSTICO DE AUDITORIA"
    ws_resumo["B9"].font = font_bold
    ws_resumo["B10"] = "Diagnóstico"
    ws_resumo["C10"] = "Qtd"
    ws_resumo["D10"] = "% Total"
    for col_c in ["B10", "C10", "D10"]:
        ws_resumo[col_c].font = font_header
        ws_resumo[col_c].fill = azul_header
        ws_resumo[col_c].alignment = Alignment(horizontal="center", vertical="center")

    diag_counts = df_resultado["Diagnóstico"].value_counts()
    r = 11
    for diag, qtd in diag_counts.items():
        ws_resumo[f"B{r}"] = diag
        ws_resumo[f"B{r}"].font = font_normal
        ws_resumo[f"B{r}"].border = thin_border
        
        ws_resumo[f"C{r}"] = qtd
        ws_resumo[f"C{r}"].font = font_normal
        ws_resumo[f"C{r}"].border = thin_border
        ws_resumo[f"C{r}"].alignment = Alignment(horizontal="center")

        pct = (qtd / total_inss) if total_inss > 0 else 0
        ws_resumo[f"D{r}"] = f"{pct:.1%}"
        ws_resumo[f"D{r}"].font = font_normal
        ws_resumo[f"D{r}"].border = thin_border
        ws_resumo[f"D{r}"].alignment = Alignment(horizontal="center")
        r += 1

    # Tabela 2: Distribuição por Espécie de Benefício
    ws_resumo["F9"] = "DISTRIBUIÇÃO POR ESPÉCIE DE BENEFÍCIO"
    ws_resumo["F9"].font = font_bold
    ws_resumo["F10"] = "Espécie INSS"
    ws_resumo["G10"] = "Qtd"
    for col_c in ["F10", "G10"]:
        ws_resumo[col_c].font = font_header
        ws_resumo[col_c].fill = azul_header
        ws_resumo[col_c].alignment = Alignment(horizontal="center", vertical="center")

    esp_counts = df_resultado["Especie_INSS"].value_counts()
    r_esp = 11
    for esp, qtd in esp_counts.items():
        ws_resumo[f"F{r_esp}"] = esp
        ws_resumo[f"F{r_esp}"].font = font_normal
        ws_resumo[f"F{r_esp}"].border = thin_border
        
        ws_resumo[f"G{r_esp}"] = qtd
        ws_resumo[f"G{r_esp}"].font = font_normal
        ws_resumo[f"G{r_esp}"].border = thin_border
        ws_resumo[f"G{r_esp}"].alignment = Alignment(horizontal="center")
        r_esp += 1

    # Ajustar larguras do resumo
    ws_resumo.column_dimensions["B"].width = 55
    ws_resumo.column_dimensions["C"].width = 10
    ws_resumo.column_dimensions["D"].width = 12
    ws_resumo.column_dimensions["E"].width = 4
    ws_resumo.column_dimensions["F"].width = 55
    ws_resumo.column_dimensions["G"].width = 10
    ws_resumo.column_dimensions["H"].width = 15
    ws_resumo.column_dimensions["I"].width = 15
    ws_resumo.column_dimensions["J"].width = 15
    ws_resumo.column_dimensions["K"].width = 15

    # ----------------------------------------------------
    # ABA 2: DIVERGÊNCIAS PRIORITÁRIAS
    # ----------------------------------------------------
    ws_prio = wb.create_sheet(title="Divergências Prioritárias")
    ws_prio.views.sheetView[0].showGridLines = True

    df_prio = df_resultado[df_resultado["Diagnóstico"].str.contains("CRÍTICO|BENEFÍCIO CESSADO", na=False)]

    cols_prio = [
        "Diagnóstico", "Ação Recomendada", "Matrícula", "Colaborador", "CPF_INSS",
        "Razão Social", "Estabelecimento", "Cargo", "Situação na Folha",
        "Especie_INSS", "NB", "Situação INSS", "Data_Inicio_DIB", "Data_Cessacao_DCB"
    ]

    for col_idx, col_name in enumerate(cols_prio, start=1):
        cell = ws_prio.cell(row=1, column=col_idx, value=col_name)
        cell.font = font_header
        cell.fill = azul_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_idx, (_, row) in enumerate(df_prio.iterrows(), start=2):
        is_critico = "CRÍTICO" in str(row.get("Diagnóstico", ""))
        row_fill = vermelho_critico if is_critico else amarelo_atencao

        for col_idx, col_name in enumerate(cols_prio, start=1):
            val = row.get(col_name, "")
            cell = ws_prio.cell(row=row_idx, column=col_idx, value=str(val) if pd.notna(val) else "")
            cell.font = font_normal
            cell.border = thin_border
            cell.fill = row_fill
            if col_name in ["Matrícula", "CPF_INSS", "NB", "Situação INSS", "Data_Inicio_DIB", "Data_Cessacao_DCB"]:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # ----------------------------------------------------
    # ABA 3: BASE COMPLETA
    # ----------------------------------------------------
    ws_base = wb.create_sheet(title="Base Completa")
    ws_base.views.sheetView[0].showGridLines = True

    cols_base = [
        "Matrícula", "Colaborador", "CPF_INSS", "NIT_INSS", "Data_Nasc_INSS",
        "Razão Social", "Código Estabelecimento", "Estabelecimento",
        "Código do Centro de Custo", "Centro de Custo", "Código do Cargo", "Cargo",
        "Situação na Folha", "Tipo de Situação", "CNPJ_INSS", "Especie_INSS",
        "NB", "Situação INSS", "Data_Req_INSS", "Data_Inicio_DIB", "Data_Despacho_INSS",
        "Data_Pericia_INSS", "Conclusao_Pericia", "Data_Limite_INSS", "Data_Cessacao_DCB",
        "Estabilidade_Acidentária", "Diagnóstico", "Ação Recomendada"
    ]

    for col_idx, col_name in enumerate(cols_base, start=1):
        cell = ws_base.cell(row=1, column=col_idx, value=col_name)
        cell.font = font_header
        cell.fill = azul_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_idx, (_, row) in enumerate(df_resultado.iterrows(), start=2):
        diag = str(row.get("Diagnóstico", ""))
        if "CRÍTICO" in diag:
            row_fill = vermelho_critico
        elif "ATENÇÃO" in diag or "CESSADO" in diag:
            row_fill = amarelo_atencao
        elif "CONFORME" in diag or "REGULAR" in diag:
            row_fill = verde_conforme
        else:
            row_fill = cinza_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None)

        for col_idx, col_name in enumerate(cols_base, start=1):
            val = row.get(col_name, "")
            cell = ws_base.cell(row=row_idx, column=col_idx, value=str(val) if pd.notna(val) else "")
            cell.font = font_normal
            cell.border = thin_border
            cell.fill = row_fill
            if col_name in ["Matrícula", "CPF_INSS", "NIT_INSS", "NB", "Situação INSS", "Data_Inicio_DIB", "Data_Cessacao_DCB"]:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # Auto-fit columns
    for ws in [ws_prio, ws_base]:
        ws.freeze_panes = "A2"
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    wb.save(caminho_saida)
    print(f"Planilha gravada com sucesso em: {caminho_saida}")


def main():
    print("=" * 70)
    print("INICIANDO MOTOR DE AUDITORIA INSS EMPRESA x FOLHA SAP (CENCOSUD)")
    print("=" * 70)

    # 1. Ler Extratos do INSS
    pasta_downloads = r"c:\Users\matcaval\Downloads"
    pdf_atacado = os.path.join(pasta_downloads, "inss-empresa-22102026.pdf")
    pdf_perini = os.path.join(pasta_downloads, "inss-empresa-20102026.pdf")

    registros_inss = []
    if os.path.exists(pdf_atacado):
        print(f"Lendo extrato Cencosud Atacado: {pdf_atacado}")
        reg_atacado = parse_inss_pdf(pdf_atacado)
        registros_inss.extend(reg_atacado)
        print(f" -> {len(reg_atacado)} registros extraídos.")

    if os.path.exists(pdf_perini):
        print(f"Lendo extrato Perini: {pdf_perini}")
        reg_perini = parse_inss_pdf(pdf_perini)
        registros_inss.extend(reg_perini)
        print(f" -> {len(reg_perini)} registros extraídos.")

    df_inss = pd.DataFrame(registros_inss)
    print(f"Total consolidado de benefícios INSS: {len(df_inss)}")

    # 2. Carregar Base COLA_Colaboradores da Cencosud
    print("Carregando base COLA_Colaboradores...")
    df_cola = carregar_base_colaboradores()
    print(f"Base de colaboradores carregada com {len(df_cola)} linhas.")

    # 3. Cruzamento Híbrido (CPF primeiro, depois Nome caso anonimizado/sem match)
    print("Cruzando dados entre INSS e SAP...")
    
    # Preparar df_cola deduplicada por CPF_clean e por Nome_norm
    cols_cola_meta = [c for c in df_cola.columns if c not in ["CPF_clean", "Nome_norm"]]
    df_cola_cpf = df_cola.dropna(subset=["CPF_clean"]).drop_duplicates(subset=["CPF_clean"])
    df_cola_nome = df_cola.dropna(subset=["Nome_norm"]).drop_duplicates(subset=["Nome_norm"])
    
    # Adicionar sufixo no df_cola para evitar colisão
    df_merge_cpf = pd.merge(
        df_inss,
        df_cola_cpf[["CPF_clean"] + cols_cola_meta],
        on="CPF_clean",
        how="left"
    )
    
    sem_match = df_merge_cpf[df_merge_cpf["Matrícula"].isna()].copy()
    com_match = df_merge_cpf[df_merge_cpf["Matrícula"].notna()].copy()

    # Tentativa 2 para os sem match: por Nome Normalizado
    if len(sem_match) > 0:
        print(f"Aplicando fallback por Nome Normalizado para {len(sem_match)} registros...")
        sem_match_inss = sem_match[list(df_inss.columns)].copy()
        sem_match_repaired = pd.merge(
            sem_match_inss,
            df_cola_nome[["Nome_norm"] + cols_cola_meta],
            on="Nome_norm",
            how="left"
        )
        df_consolidado = pd.concat([com_match, sem_match_repaired], ignore_index=True)
    else:
        df_consolidado = df_merge_cpf

    # Renomear / Organizar colunas
    df_consolidado["Colaborador"] = df_consolidado["Colaborador"].fillna(df_consolidado["Nome_INSS"])
    df_consolidado["Situação na Folha"] = df_consolidado["Situação"].fillna("NÃO ENCONTRADO NA FOLHA")
    df_consolidado["Situação INSS"] = df_consolidado["Situacao_INSS"]
    df_consolidado["Matrícula"] = df_consolidado["Matrícula"].apply(
        lambda x: str(int(float(x))) if pd.notna(x) and str(x).strip() != "" and str(x) != "nan" else ""
    )

    # 4. Diagnosticar cada caso
    diagnosticos = []
    acoes = []
    estabilidades = []
    for _, row in df_consolidado.iterrows():
        diag, acao, estab = diagnosticar_registro(row)
        diagnosticos.append(diag)
        acoes.append(acao)
        estabilidades.append(estab)

    df_consolidado["Diagnóstico"] = diagnosticos
    df_consolidado["Ação Recomendada"] = acoes
    df_consolidado["Estabilidade_Acidentária"] = estabilidades

    # 5. Exportar Relatório Excel
    pasta_destino = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\afastados"
    os.makedirs(pasta_destino, exist_ok=True)
    caminho_final = os.path.join(pasta_destino, "Auditoria_Afastados_INSS_Cencosud_2026.xlsx")

    print(f"Gerando relatório corporativo em: {caminho_final}")
    gerar_relatorio_excel(df_consolidado, caminho_final)

    print("\n" + "=" * 70)
    print("RESUMO DO PROCESSAMENTO:")
    print(f" - Total de benefícios analisados no INSS: {len(df_consolidado)}")
    print(" - Casos Críticos (Ativos na folha com benefício no INSS):", len(df_consolidado[df_consolidado["Diagnóstico"].str.contains("CRÍTICO")]))
    print(" - Casos de Atenção (Benefício Cessado sem retorno na folha):", len(df_consolidado[df_consolidado["Diagnóstico"].str.contains("BENEFÍCIO CESSADO")]))
    print(" - Casos Regulares / Conformes:", len(df_consolidado[df_consolidado["Diagnóstico"].str.contains("CONFORME|REGULAR")]))
    print("=" * 70)

if __name__ == "__main__":
    main()
