import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Carregar consultas realizadas
with open(r"robo_afastados\consultas_completas.json", "r", encoding="utf-8") as f:
    consultas = json.load(f)

# Consulta da Denise (Obito confirmado)
DENISE_DADOS = {
    "matricula": "AF-ÓBITO",
    "cpf_origem": "143.559.338-30",
    "nome_origem": "DENISE BAGANO ROCHA CORREIA",
    "nasc_origem": "24/04/1974",
    "json": {
        "code": "0",
        "status": "OK",
        "message": "Pesquisa realizada com sucesso.",
        "cpf": "143.559.338-30",
        "nome": "Denise Bagano Rocha Correia",
        "nome_mae": "Alice Bagano Rocha",
        "genero": {
            "nome_social": None,
            "nome_civil": "Denise Bagano Rocha Correia",
            "sexo": "F",
            "transicao_genero": False
        },
        "data_nascimento": "24/04/1974",
        "situacao_cadastral": "Titular falecido",
        "data_inscricao": "anterior a 10/11/1990",
        "uf": ["SP"],
        "digito_verificador": "03",
        "comprovante": "B92B.4F26.7AD7.8BAB",
        "idade": "52",
        "ano_obito": "2025",
        "Comprovantes": {
            "digito_verificador": "03",
            "codigo_controle": "B92B.4F26.7AD7.8BAB",
            "qrcode": "https://sintegraws.com.br/api/v1/cpf/tipo-retorno/comprovante-png/748793e2-65ff-47b4-a120-be4215b9442b",
            "data_consulta": "28/09/2026 09:15:24",
            "arquivo": {
                "extensao": "pdf",
                "url": "https://sintegraws.com.br/api/v1/cpf/tipo-retorno/comprovante-pdf/3d58a00c-ff76-46d4-aeb1-77fcdff6851b"
            }
        },
        "version": "1"
    }
}

# Colocar Denise na primeira posicao da lista
lista_completa = [DENISE_DADOS] + consultas

# ==================== GERANDO EXCEL ====================
wb = Workbook()
ws = wb.active
ws.title = "Auditoria Afastados - SintegraWS"
ws.views.sheetView[0].showGridLines = True

headers = [
    "Alerta / Status", "Matrícula", "CPF Informado", "Nome Informado", "Data Nasc Informada",
    "Status RFB", "Situação Cadastral", "Ano Óbito", "Nome na Receita Federal",
    "Data Nasc RFB", "Idade", "UF", "Data Inscrição RFB", "Código Comprovante", "Link Comprovante PDF", "Mensagem Retorno"
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

for item in lista_completa:
    j = item["json"]
    status = j.get("status", "")
    sit_cad = j.get("situacao_cadastral", "")
    ano_ob = j.get("ano_obito", "")
    
    # Classificacao de Alerta
    if sit_cad and "falecido" in str(sit_cad).lower():
        alerta = "ALERTA: ÓBITO CONFIRMADO"
    elif status == "OK" and sit_cad == "Regular":
        alerta = "REGULAR"
    else:
        alerta = "VERIFICAR DADOS"
        
    comprovantes_obj = j.get("Comprovantes", {})
    pdf_url = ""
    if isinstance(comprovantes_obj, dict):
        arq = comprovantes_obj.get("arquivo", {})
        if isinstance(arq, dict):
            pdf_url = arq.get("url", "")

    row_data = [
        alerta,
        item["matricula"],
        item["cpf_origem"],
        item["nome_origem"],
        item["nasc_origem"],
        status,
        sit_cad if sit_cad else "—",
        ano_ob if ano_ob not in [None, ""] else "Não consta",
        j.get("nome", "—"),
        j.get("data_nascimento", "—"),
        j.get("idade", "—"),
        ", ".join(j.get("uf", [])) if isinstance(j.get("uf"), list) else str(j.get("uf", "")),
        j.get("data_inscricao", "—"),
        j.get("comprovante", "—"),
        pdf_url,
        j.get("message", "")
    ]
    ws.append(row_data)

# Estilizacao
for row_idx in range(2, ws.max_row + 1):
    alerta_val = ws.cell(row=row_idx, column=1).value
    
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=row_idx, column=col_idx)
        c.font = Font(name="Calibri", size=10)
        c.border = thin_border
        
        # Alinhamentos
        if col_idx in [1, 2, 3, 5, 6, 7, 8, 10, 11, 12, 13, 14]:
            c.alignment = Alignment(horizontal="center", vertical="center")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")
            
        # Cores de destaque
        if "ÓBITO" in str(alerta_val):
            c.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # Vermelho claro
            if col_idx in [1, 7, 8]:
                c.font = Font(name="Calibri", size=10, bold=True, color="9C0006")
        elif alerta_val == "REGULAR":
            if col_idx == 1:
                c.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                c.font = Font(name="Calibri", size=10, bold=True, color="006100")
        elif alerta_val == "VERIFICAR DADOS":
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
r_alert = p_alert_box.add_run("🚨 ATENÇÃO: CASO CRÍTICO DE ÓBITO IDENTIFICADO NA RECEITA FEDERAL\n")
r_alert.font.name = "Calibri"
r_alert.font.size = Pt(12)
r_alert.font.bold = True
r_alert.font.color.rgb = RGBColor(192, 0, 0)

r_alert_det = p_alert_box.add_run(
    f"• Colaboradora: DENISE BAGANO ROCHA CORREIA\n"
    f"• CPF: 143.559.338-30 | Data de Nascimento: 24/04/1974\n"
    f"• Situação Cadastral na RFB: TITULAR FALECIDO\n"
    f"• Ano do Óbito: 2025\n"
    f"• Comprovante RFB: B92B.4F26.7AD7.8BAB\n"
    f"• Link do PDF Oficial: https://sintegraws.com.br/api/v1/cpf/tipo-retorno/comprovante-pdf/3d58a00c-ff76-46d4-aeb1-77fcdff6851b\n"
    f"-> Recomendação de RH: Bloqueio imediato de folha/benefícios e abertura de procedimento de rescisão por falecimento."
)
r_alert_det.font.name = "Calibri"
r_alert_det.font.size = Pt(10)
r_alert_det.font.color.rgb = RGBColor(120, 0, 0)

# Resumo
p_sum = doc.add_paragraph()
p_sum.add_run("1. RESUMO EXECUTIVO DO PROCESSAMENTO\n").font.bold = True
p_sum.paragraph_format.space_after = Pt(8)

p_sum.add_run(f"• Total de Colaboradores Auditados: {len(lista_completa)}\n")
p_sum.add_run(f"• Casos com Óbito Confirmado: 1 (Denise Bagano Rocha Correia - Óbito 2025)\n")
p_sum.add_run(f"• Casos com Cadastro Regular: 1 (Vera Lucia Sousa Feitoza)\n")
p_sum.add_run(f"• Casos com Necessidade de Validação da Data de Nascimento: 20 registros\n")
p_sum.add_run(f"  (Observação técnica: Na planilha de origem, algumas datas de nascimento estão divergentes da base cadastral da Receita Federal ou com formato de ano de 2 dígitos. Ajustando a data no sistema, a consulta retorna automaticamente).\n")

# Detalhamento
p_det = doc.add_paragraph()
p_det.paragraph_format.space_before = Pt(10)
p_det.add_run("2. DETALHAMENTO INDIVIDUAL DE CADA REGISTRO CONSULTADO\n").font.bold = True

for idx, item in enumerate(lista_completa, 1):
    j = item["json"]
    status = j.get("status", "")
    sit_cad = j.get("situacao_cadastral", "")
    ano_ob = j.get("ano_obito")
    
    p_item = doc.add_paragraph()
    is_obito = sit_cad and "falecido" in str(sit_cad).lower()
    
    titulo_texto = f"Registro #{idx:02d} — {item['nome_origem']} (Matrícula: {item['matricula']})"
    if is_obito:
        titulo_texto += " [🚨 ÓBITO CONFIRMADO]"
        
    r_idx = p_item.add_run(titulo_texto)
    r_idx.font.name = "Calibri"
    r_idx.font.size = Pt(11)
    r_idx.font.bold = True
    if is_obito:
        r_idx.font.color.rgb = RGBColor(192, 0, 0)
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
            
    add_row_info("CPF Informado:", item["cpf_origem"])
    add_row_info("Data Nasc. Informada:", item["nasc_origem"])
    add_row_info("Status na Receita Federal:", f"{status} (Código: {j.get('code')})", destaque_verde=(status == "OK"), destaque_vermelho=(status != "OK"))
    add_row_info("Mensagem / Retorno:", j.get("message"))
    
    if status == "OK":
        add_row_info("Nome na RFB:", j.get("nome"))
        add_row_info("CPF na RFB:", j.get("cpf"))
        add_row_info("Situação Cadastral:", sit_cad, destaque_vermelho=is_obito, destaque_verde=(sit_cad == "Regular"))
        add_row_info("Ano de Óbito:", ano_ob if ano_ob not in [None, ""] else "Não consta óbito", destaque_vermelho=bool(ano_ob))
        add_row_info("Data Nasc. na RFB:", j.get("data_nascimento"))
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
print("Arquivos atualizados com sucesso com o caso de óbito!")
