import pandas as pd
import requests
import json
import time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
from docx.shared import Inches, Pt, RGBColor
import os
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

TOKEN = os.getenv("SINTEGRAWS_TOKEN", "")
URL = "https://www.sintegraws.com.br/api/v1/execute-api.php"

# Consulta salva (Óbito) - Modelo para homologação/teste
DENISE_DADOS = {
    "matricula": "AF-ÓBITO",
    "cpf_origem": "000.000.000-00",
    "nome_origem": "COLABORADOR FALECIDO EXEMPLO",
    "nasc_origem": "01/01/1975",
    "json": {
        "code": "0",
        "status": "OK",
        "message": "Pesquisa realizada com sucesso.",
        "cpf": "000.000.000-00",
        "nome": "Colaborador Falecido Exemplo",
        "nome_mae": "Mae Exemplo",
        "genero": {
            "nome_social": None,
            "nome_civil": "Colaborador Falecido Exemplo",
            "sexo": "F",
            "transicao_genero": False
        },
        "data_nascimento": "01/01/1975",
        "situacao_cadastral": "Titular falecido",
        "data_inscricao": "anterior a 10/11/1990",
        "uf": ["SP"],
        "digito_verificador": "00",
        "comprovante": "XXXX.XXXX.XXXX.XXXX",
        "idade": "50",
        "ano_obito": "2025",
        "Comprovantes": {
            "digito_verificador": "00",
            "codigo_controle": "XXXX.XXXX.XXXX.XXXX"
        },
        "version": "1"
    }
}

# Cache de consultas já realizadas para poupar saldo de API
CACHE_SUCESSO = {}

def tratar_data_nascimento_8digitos(data_raw):
    data_str = str(data_raw).strip()
    if "/" in data_str:
        partes = data_str.split("/")
        dia = partes[0].zfill(2)
        mes = partes[1].zfill(2)
        ano = partes[2]
        if len(ano) == 2:
            ano_int = int(ano)
            if ano_int > 30:
                ano = f"19{ano}"
            else:
                ano = f"20{ano}"
        return f"{dia}{mes}{ano}"
    else:
        num = "".join(filter(str.isdigit, data_str))
        if len(num) == 6:
            dia, mes, ano = num[:2], num[2:4], num[4:]
            ano_int = int(ano)
            if ano_int > 30:
                ano = f"19{ano}"
            else:
                ano = f"20{ano}"
            return f"{dia}{mes}{ano}"
        return num

df = pd.read_excel(r"robo_afastados\colaboradores_temp.xlsx")
resultados = []

print(f"Iniciando varredura de {len(df)} colaboradores...")

for idx, row in df.iterrows():
    matr = row.get('CDN_FUNCIONARIO')
    cpf_raw = str(row.get('COD_ID_FEDER'))
    cpf_limpo = "".join(filter(str.isdigit, cpf_raw)).zfill(11)
    nome_planilha = str(row.get('NOM_PESSOA_FISIC'))
    nasc_raw = row.get('DAT_NASCIMENTO')
    nasc_8dig = tratar_data_nascimento_8digitos(nasc_raw)
    
    print(f"[{idx+1}/{len(df)}] Processando: {nome_planilha} | CPF: {cpf_limpo} | Nasc: {nasc_8dig}")
    
    if cpf_limpo in CACHE_SUCESSO:
        res_json = CACHE_SUCESSO[cpf_limpo]
        print(f"  -> Recuperado de cache de sucesso: {res_json.get('nome')}")
    else:
        params = {
            "token": TOKEN,
            "cpf": cpf_limpo,
            "data-nascimento": nasc_8dig,
            "plugin": "CPF"
        }
        try:
            r = requests.get(URL, params=params, timeout=30)
            res_json = r.json()
            print(f"  -> Retorno: {res_json.get('status')} | {res_json.get('nome')} | {res_json.get('situacao_cadastral')}")
        except Exception as e:
            res_json = {"code": "-1", "status": "ERROR", "message": f"Erro de conexao: {str(e)}"}
            print(f"  -> Erro: {e}")
        time.sleep(1.2)
        
    resultados.append({
        "matricula": matr,
        "cpf_origem": cpf_limpo,
        "nome_origem": nome_planilha,
        "nasc_origem": f"{nasc_8dig[:2]}/{nasc_8dig[2:4]}/{nasc_8dig[4:]}",
        "json": res_json
    })

# Unir com Denise no topo
lista_geral = [DENISE_DADOS] + resultados

# Salvar JSON
with open(r"robo_afastados\consultas_completas_oficiais.json", "w", encoding="utf-8") as f:
    json.dump(lista_geral, f, ensure_ascii=False, indent=2)

# ==================== GERANDO EXCEL ====================
print("\nGerando planilha oficial estilizada...")
wb = Workbook()
ws = wb.active
ws.title = "Auditoria Afastados RFB"
ws.views.sheetView[0].showGridLines = True

headers = [
    "Status Geral", "Matrícula", "CPF", "Nome Colaborador", "Data Nascimento",
    "Situação Cadastral", "Ano Óbito", "Nome na Receita Federal", "Nome da Mãe",
    "Idade", "UF", "Data Inscrição", "Código Comprovante", "Link PDF Comprovante", "Mensagem Retorno"
]

header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

ws.append(headers)
for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for item in lista_geral:
    j = item["json"]
    status = j.get("status", "")
    sit_cad = j.get("situacao_cadastral", "")
    ano_ob = j.get("ano_obito", "")
    
    if sit_cad and "falecido" in str(sit_cad).lower():
        alerta = "🚨 ÓBITO CONFIRMADO"
    elif status == "OK":
        alerta = "✅ REGULAR"
    else:
        alerta = "⚠️ DIVERGÊNCIA"
        
    comprovantes_obj = j.get("Comprovantes", {})
    pdf_url = ""
    if isinstance(comprovantes_obj, dict):
        arq = comprovantes_obj.get("arquivo", {})
        if isinstance(arq, dict):
            pdf_url = arq.get("url", "")

    row_data = [
        alerta,
        item["matricula"],
        j.get("cpf") or item["cpf_origem"],
        item["nome_origem"],
        j.get("data_nascimento") or item["nasc_origem"],
        sit_cad or "—",
        ano_ob if ano_ob not in [None, ""] else "Não consta",
        j.get("nome", "—"),
        j.get("nome_mae", "—"),
        j.get("idade", "—"),
        ", ".join(j.get("uf", [])) if isinstance(j.get("uf"), list) else str(j.get("uf", "")),
        j.get("data_inscricao", "—"),
        j.get("comprovante", "—"),
        pdf_url,
        j.get("message", "")
    ]
    ws.append(row_data)

for row_idx in range(2, ws.max_row + 1):
    alerta_val = ws.cell(row=row_idx, column=1).value
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=row_idx, column=col_idx)
        c.font = Font(name="Calibri", size=10)
        c.border = thin_border
        
        if col_idx in [1, 2, 3, 5, 6, 7, 10, 11, 12, 13]:
            c.alignment = Alignment(horizontal="center", vertical="center")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")
            
        if "ÓBITO" in str(alerta_val):
            c.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            if col_idx in [1, 6, 7]:
                c.font = Font(name="Calibri", size=10, bold=True, color="9C0006")
        elif "REGULAR" in str(alerta_val):
            if col_idx == 1:
                c.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                c.font = Font(name="Calibri", size=10, bold=True, color="006100")
        elif "DIVERGÊNCIA" in str(alerta_val):
            if col_idx == 1:
                c.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
                c.font = Font(name="Calibri", size=10, bold=True, color="9C6500")

for col in ws.columns:
    max_len = 0
    col_letter = get_column_letter(col[0].column)
    for cell in col:
        val_str = str(cell.value or "")
        if len(val_str) > max_len:
            max_len = len(val_str)
    ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 40)

caminho_excel = r"robo_afastados\relatorio_afastados_sintegraws.xlsx"
wb.save(caminho_excel)

# ==================== GERANDO DOCX ====================
print("Gerando relatório em Word com todas as saídas...")
doc = docx.Document()
for s in doc.sections:
    s.top_margin = Inches(0.8)
    s.bottom_margin = Inches(0.8)
    s.left_margin = Inches(0.8)
    s.right_margin = Inches(0.8)

# Titulo
p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_title = p_title.add_run("RELATÓRIO DE AUDITORIA CADASTRAL E PROVA DE VIDA\nCONSULTA RECEITA FEDERAL VIA SINTEGRAWS")
run_title.font.name = "Calibri"
run_title.font.size = Pt(15)
run_title.font.bold = True
run_title.font.color.rgb = RGBColor(31, 78, 120)

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_sub = p_sub.add_run(f"Emissão: 28/09/2026 | Sistema de RH - Colaboradores Afastados")
run_sub.font.name = "Calibri"
run_sub.font.size = Pt(10)
run_sub.font.italic = True
run_sub.font.color.rgb = RGBColor(89, 89, 89)

# Box Destaque de Óbito
p_alert_box = doc.add_paragraph()
p_alert_box.paragraph_format.space_before = Pt(12)
p_alert_box.paragraph_format.space_after = Pt(14)
r_alert = p_alert_box.add_run("🚨 CASO CRÍTICO DE ÓBITO IDENTIFICADO NA RECEITA FEDERAL\n")
r_alert.font.name = "Calibri"
r_alert.font.size = Pt(12)
r_alert.font.bold = True
r_alert.font.color.rgb = RGBColor(192, 0, 0)

r_alert_det = p_alert_box.add_run(
    f"• Colaboradora: DENISE BAGANO ROCHA CORREIA\n"
    f"• CPF: 143.559.338-30 | Data de Nascimento: 24/04/1974\n"
    f"• Situação Cadastral: TITULAR FALECIDO\n"
    f"• Ano do Óbito: 2025\n"
    f"• Comprovante RFB: B92B.4F26.7AD7.8BAB\n"
    f"• Link do PDF Oficial: https://sintegraws.com.br/api/v1/cpf/tipo-retorno/comprovante-pdf/3d58a00c-ff76-46d4-aeb1-77fcdff6851b\n"
    f"-> Recomendação de RH: Bloqueio de folha/benefícios e abertura imediata de rescisão por falecimento."
)
r_alert_det.font.name = "Calibri"
r_alert_det.font.size = Pt(10)
r_alert_det.font.color.rgb = RGBColor(120, 0, 0)

# Resumo
tot_sucesso = sum(1 for r in lista_geral if r["json"].get("status") == "OK")
tot_obito = sum(1 for r in lista_geral if "falecido" in str(r["json"].get("situacao_cadastral", "")).lower())

p_sum = doc.add_paragraph()
p_sum.add_run("1. RESUMO EXECUTIVO DO PROCESSAMENTO\n").font.bold = True
p_sum.paragraph_format.space_after = Pt(8)
p_sum.add_run(f"• Total de Colaboradores Processados: {len(lista_geral)}\n")
p_sum.add_run(f"• Consultas com Sucesso na Receita Federal: {tot_sucesso}\n")
p_sum.add_run(f"• Casos com Óbito Confirmado: {tot_obito}\n")
p_sum.add_run(f"• Consultas com Divergência Cadastral: {len(lista_geral) - tot_sucesso}\n")

# Detalhamento
p_det = doc.add_paragraph()
p_det.paragraph_format.space_before = Pt(10)
p_det.add_run("2. DETALHAMENTO INDIVIDUAL DE CADA REGISTRO CONSULTADO\n").font.bold = True

for idx, item in enumerate(lista_geral, 1):
    j = item["json"]
    status = j.get("status", "")
    sit_cad = j.get("situacao_cadastral", "")
    ano_ob = j.get("ano_obito")
    is_obito = sit_cad and "falecido" in str(sit_cad).lower()
    
    p_item = doc.add_paragraph()
    titulo_texto = f"Registro #{idx:02d} — {item['nome_origem']} (Matrícula: {item['matricula']})"
    if is_obito:
        titulo_texto += " [🚨 ÓBITO CONFIRMADO]"
    elif status == "OK":
        titulo_texto += " [✅ REGULAR]"
        
    r_idx = p_item.add_run(titulo_texto)
    r_idx.font.name = "Calibri"
    r_idx.font.size = Pt(11)
    r_idx.font.bold = True
    if is_obito:
        r_idx.font.color.rgb = RGBColor(192, 0, 0)
    elif status == "OK":
        r_idx.font.color.rgb = RGBColor(0, 128, 0)
    else:
        r_idx.font.color.rgb = RGBColor(31, 78, 120)
        
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    largura_col1 = Inches(2.3)
    largura_col2 = Inches(4.5)
    
    def add_row_info(lbl, val, destaque_vermelho=False, destaque_verde=False):
        row = table.add_row()
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = largura_col1
        c1.width = largura_col2
        
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(lbl)
        r0.font.name = "Calibri"
        r0.font.size = Pt(9.5)
        r0.font.bold = True
        
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(str(val if val not in [None, ""] else "—"))
        r1.font.name = "Calibri"
        r1.font.size = Pt(9.5)
        if destaque_vermelho:
            r1.font.bold = True
            r1.font.color.rgb = RGBColor(192, 0, 0)
        elif destaque_verde:
            r1.font.bold = True
            r1.font.color.rgb = RGBColor(0, 128, 0)
            
    add_row_info("CPF:", j.get("cpf") or item["cpf_origem"])
    add_row_info("Data Nascimento:", j.get("data_nascimento") or item["nasc_origem"])
    add_row_info("Status na Receita Federal:", f"{status} (Código: {j.get('code')})", destaque_verde=(status == "OK"), destaque_vermelho=(status != "OK"))
    add_row_info("Mensagem / Retorno:", j.get("message"))
    
    if status == "OK":
        add_row_info("Nome na RFB:", j.get("nome"))
        add_row_info("Nome da Mãe:", j.get("nome_mae"))
        add_row_info("Situação Cadastral:", sit_cad, destaque_vermelho=is_obito, destaque_verde=(sit_cad == "Regular"))
        add_row_info("Ano de Óbito:", ano_ob if ano_ob not in [None, ""] else "Não consta óbito", destaque_vermelho=bool(ano_ob))
        add_row_info("Idade:", f"{j.get('idade')} anos" if j.get('idade') else "—")
        uf_val = ", ".join(j.get("uf", [])) if isinstance(j.get("uf"), list) else str(j.get("uf", ""))
        add_row_info("UF:", uf_val)
        add_row_info("Data de Inscrição:", j.get("data_inscricao"))
        add_row_info("Código Comprovante:", j.get("comprovante"))
        
        comp_dict = j.get("Comprovantes", {})
        if isinstance(comp_dict, dict):
            arq = comp_dict.get("arquivo", {})
            if isinstance(arq, dict) and arq.get("url"):
                add_row_info("Link Comprovante PDF:", arq.get("url"))
                
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

caminho_docx = r"robo_afastados\relatorio_detalhado_afastados.docx"
doc.save(caminho_docx)
print("Processamento concluído com sucesso total!")
