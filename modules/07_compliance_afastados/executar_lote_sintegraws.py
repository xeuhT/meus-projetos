import pandas as pd
import requests
import json
import time
import os
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

TOKEN = os.getenv("SINTEGRAWS_TOKEN", "")
URL = "https://www.sintegraws.com.br/api/v1/execute-api.php"

# Consulta salva da linha 0 (Modelo para teste/homologação)
RESULTADO_LINHA_0 = {
    "code": "0",
    "status": "OK",
    "message": "Pesquisa realizada com sucesso.",
    "cpf": "000.000.000-00",
    "nome": "Colaborador Exemplo",
    "nome_mae": "Mae Exemplo",
    "genero": {
        "nome_social": None,
        "nome_civil": "Colaborador Exemplo",
        "sexo": "F",
        "transicao_genero": False
    },
    "data_nascimento": "01/01/1980",
    "situacao_cadastral": "Regular",
    "data_inscricao": "01/01/2000",
    "uf": ["BA", "SE"],
    "digito_verificador": "00",
    "comprovante": "XXXX.XXXX.XXXX.XXXX",
    "idade": "45",
    "ano_obito": "",
    "qsa": [],
    "Comprovantes": {
        "digito_verificador": "00",
        "codigo_controle": "XXXX.XXXX.XXXX.XXXX"
    },
    "version": "1"
}

def formatar_data_nascimento(data_raw):
    data_str = str(data_raw).strip()
    if "/" in data_str:
        partes = data_str.split("/")
        dia = partes[0].zfill(2)
        mes = partes[1].zfill(2)
        ano = partes[2]
        if len(ano) == 2:
            ano_int = int(ano)
            # Regra para século
            if ano_int > 30:
                ano = f"19{ano}"
            else:
                ano = f"20{ano}"
        return f"{dia}/{mes}/{ano}"
    return data_str

def consultar_sintegraws(cpf, data_nasc_formatada):
    cpf_limpo = "".join(filter(str.isdigit, str(cpf))).zfill(11)
    params = {
        "token": TOKEN,
        "cpf": cpf_limpo,
        "data_nascimento": data_nasc_formatada,
        "plugin": "CPF"
    }
    try:
        resp = requests.get(URL, params=params, timeout=30)
        return resp.json()
    except Exception as e:
        return {"code": "-1", "status": "ERROR", "message": f"Erro de conexao: {str(e)}"}

print("Iniciando processamento...")
df = pd.read_excel(r"robo_afastados\colaboradores_temp.xlsx")

resultados_lista = []

for idx, row in df.iterrows():
    matr = row.get('CDN_FUNCIONARIO')
    cpf_raw = str(row.get('COD_ID_FEDER'))
    nome_planilha = str(row.get('NOM_PESSOA_FISIC'))
    nasc_raw = row.get('DAT_NASCIMENTO')
    nasc_format = formatar_data_nascimento(nasc_raw)
    
    print(f"[{idx+1}/{len(df)}] Processando: {nome_planilha} (CPF: {cpf_raw}) | Nasc: {nasc_format}")
    
    if idx == 0:
        # Reaproveita a consulta ja feita pelo usuario
        res_json = RESULTADO_LINHA_0
    else:
        res_json = consultar_sintegraws(cpf_raw, nasc_format)
        time.sleep(1.2) # intervalo seguro entre consultas
        
    resultados_lista.append({
        "matricula": matr,
        "cpf_origem": cpf_raw,
        "nome_origem": nome_planilha,
        "nasc_origem": nasc_format,
        "json": res_json
    })

# Salvar json de backup
with open(r"robo_afastados\consultas_completas.json", "w", encoding="utf-8") as f:
    json.dump(resultados_lista, f, ensure_ascii=False, indent=2)

print("\n--- GERANDO EXCEL PROFISSIONAL ---")
wb = Workbook()
ws = wb.active
ws.title = "Afastados - RFB SintegraWS"
ws.views.sheetView[0].showGridLines = True

headers = [
    "Matrícula", "CPF Informado", "Nome Planilha", "Data Nasc Informada",
    "Status Consulta", "Mensagem Retorno", "CPF RFB", "Nome RFB", 
    "Situação Cadastral", "Ano Óbito", "Data Nascimento RFB", "Idade", 
    "UF", "Data Inscrição", "Código Comprovante", "Link Comprovante PDF"
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
for col_idx, cell in enumerate(ws[1], 1):
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for item in resultados_lista:
    j = item["json"]
    status = j.get("status", "")
    code = str(j.get("code", ""))
    msg = j.get("message", "")
    
    cpf_rfb = j.get("cpf", "")
    nome_rfb = j.get("nome", "")
    sit_cadastral = j.get("situacao_cadastral", "")
    ano_obito = j.get("ano_obito", "")
    if ano_obito is None or ano_obito == "":
        ano_obito = "Não consta"
    
    nasc_rfb = j.get("data_nascimento", "")
    idade = j.get("idade", "")
    uf = ", ".join(j.get("uf", [])) if isinstance(j.get("uf"), list) else str(j.get("uf", ""))
    dt_inscricao = j.get("data_inscricao", "")
    comprovante_cod = j.get("comprovante", "")
    
    comprovantes_obj = j.get("Comprovantes", {})
    if isinstance(comprovantes_obj, dict):
        arquivo_obj = comprovantes_obj.get("arquivo", {})
        pdf_url = arquivo_obj.get("url", "") if isinstance(arquivo_obj, dict) else ""
    else:
        pdf_url = ""

    row_data = [
        item["matricula"],
        item["cpf_origem"],
        item["nome_origem"],
        item["nasc_origem"],
        status,
        msg,
        cpf_rfb,
        nome_rfb,
        sit_cadastral,
        ano_obito,
        nasc_rfb,
        idade,
        uf,
        dt_inscricao,
        comprovante_cod,
        pdf_url
    ]
    ws.append(row_data)

# Estilização das linhas de dados
for row_idx in range(2, ws.max_row + 1):
    status_val = ws.cell(row=row_idx, column=5).value
    sit_val = ws.cell(row=row_idx, column=9).value
    obito_val = ws.cell(row=row_idx, column=10).value
    
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=row_idx, column=col_idx)
        c.font = Font(name="Calibri", size=10)
        c.border = thin_border
        
        # Alinhamentos específicos
        if col_idx in [1, 2, 4, 5, 7, 10, 11, 12, 13, 14, 15]:
            c.alignment = Alignment(horizontal="center", vertical="center")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")
            
        # Destaque de óbito se houver
        if col_idx == 10 and obito_val not in ["Não consta", "", None]:
            c.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            c.font = Font(name="Calibri", size=10, bold=True, color="9C0006")
            
        # Destaque de status
        if col_idx == 5:
            if status_val == "OK":
                c.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                c.font = Font(name="Calibri", size=10, bold=True, color="006100")
            else:
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
print(f"Excel salvo com sucesso: {caminho_excel}")

print("\n--- GERANDO RELATORIO EM DOCX ---")
doc = docx.Document()

# Margens
sections = doc.sections
for s in sections:
    s.top_margin = Inches(0.8)
    s.bottom_margin = Inches(0.8)
    s.left_margin = Inches(0.8)
    s.right_margin = Inches(0.8)

# Titulo
p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_title = p_title.add_run("RELATÓRIO DE CONSULTA CADASTRAL NA RECEITA FEDERAL (SINTEGRAWS)")
run_title.font.name = "Calibri"
run_title.font.size = Pt(16)
run_title.font.bold = True
run_title.font.color.rgb = RGBColor(31, 78, 120)

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_sub = p_sub.add_run(f"Auditoria de Colaboradores Afastados - Gerado em {datetime.datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}")
run_sub.font.name = "Calibri"
run_sub.font.size = Pt(10)
run_sub.font.italic = True
run_sub.font.color.rgb = RGBColor(89, 89, 89)

doc.add_paragraph().paragraph_format.space_after = Pt(10)

# Resumo
p_sum = doc.add_paragraph()
p_sum.add_run("1. RESUMO EXECUTIVO DO PROCESSAMENTO\n").font.bold = True
total = len(resultados_lista)
sucessos = sum(1 for r in resultados_lista if r["json"].get("status") == "OK")
divergentes = total - sucessos
obitos_count = sum(1 for r in resultados_lista if r["json"].get("ano_obito") not in ["", None])

p_sum.add_run(f"• Total de registros na planilha: {total}\n")
p_sum.add_run(f"• Consultas com sucesso na Receita Federal: {sucessos}\n")
p_sum.add_run(f"• Consultas com divergência / erro: {divergentes}\n")
p_sum.add_run(f"• Óbitos confirmados na base: {obitos_count}\n")

doc.add_paragraph().paragraph_format.space_after = Pt(10)

p_det = doc.add_paragraph()
p_det.add_run("2. DETALHAMENTO INDIVIDUAL DE CADA CONSULTA\n").font.bold = True
p_det.paragraph_format.space_after = Pt(8)

for idx, item in enumerate(resultados_lista, 1):
    j = item["json"]
    status = j.get("status", "")
    
    # Bloco individual
    p_item = doc.add_paragraph()
    r_idx = p_item.add_run(f"Registro #{idx:02d} — {item['nome_origem']} (Matrícula: {item['matricula']})")
    r_idx.font.name = "Calibri"
    r_idx.font.size = Pt(11)
    r_idx.font.bold = True
    r_idx.font.color.rgb = RGBColor(31, 78, 120)
    
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    largura_col1 = Inches(2.2)
    largura_col2 = Inches(4.5)
    
    def add_row_info(lbl, val, destaque_vermelho=False, destaque_verde=False):
        row = table.add_row()
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        cell_lbl.width = largura_col1
        cell_val.width = largura_col2
        
        p0 = cell_lbl.paragraphs[0]
        r0 = p0.add_run(lbl)
        r0.font.name = "Calibri"
        r0.font.size = Pt(9.5)
        r0.font.bold = True
        
        p1 = cell_val.paragraphs[0]
        r1 = p1.add_run(str(val if val not in [None, ""] else "—"))
        r1.font.name = "Calibri"
        r1.font.size = Pt(9.5)
        if destaque_vermelho:
            r1.font.bold = True
            r1.font.color.rgb = RGBColor(192, 0, 0)
        elif destaque_verde:
            r1.font.bold = True
            r1.font.color.rgb = RGBColor(0, 128, 0)
            
    add_row_info("CPF Informado:", item["cpf_origem"])
    add_row_info("Data Nasc. Informada:", item["nasc_origem"])
    add_row_info("Status da API:", f"{status} (Código: {j.get('code')})", destaque_verde=(status == "OK"), destaque_vermelho=(status != "OK"))
    add_row_info("Mensagem:", j.get("message"))
    
    if status == "OK":
        add_row_info("Nome na RFB:", j.get("nome"))
        add_row_info("CPF na RFB:", j.get("cpf"))
        add_row_info("Situação Cadastral:", j.get("situacao_cadastral"))
        ano_ob = j.get("ano_obito")
        add_row_info("Ano de Óbito:", ano_ob if ano_ob not in [None, ""] else "Não consta óbito", destaque_vermelho=(ano_ob not in [None, ""]))
        add_row_info("Data Nasc. na RFB:", j.get("data_nascimento"))
        add_row_info("Idade:", f"{j.get('idade')} anos" if j.get('idade') else "—")
        uf_val = ", ".join(j.get("uf", [])) if isinstance(j.get("uf"), list) else str(j.get("uf", ""))
        add_row_info("UF:", uf_val)
        add_row_info("Data de Inscrição:", j.get("data_inscricao"))
        add_row_info("Código de Comprovante:", j.get("comprovante"))
        
        comp_dict = j.get("Comprovantes", {})
        if isinstance(comp_dict, dict):
            arq = comp_dict.get("arquivo", {})
            url_pdf = arq.get("url") if isinstance(arq, dict) else ""
            if url_pdf:
                add_row_info("Link Comprovante PDF:", url_pdf)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

caminho_docx = r"robo_afastados\relatorio_detalhado_afastados.docx"
doc.save(caminho_docx)
print(f"Docx salvo com sucesso: {caminho_docx}")
print("Processamento 100% concluído!")
