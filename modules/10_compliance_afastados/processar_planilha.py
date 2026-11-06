import pandas as pd
import requests
import os
import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

API_TOKEN = os.getenv("INFOSIMPLES_TOKEN", "")
URL_API = "https://api.infosimples.com/api/v2/consultas/receita-federal/cpf"

def consultar_infosimples(cpf, data_nascimento):
    try:
        cpf_limpo = "".join(filter(str.isdigit, str(cpf))).zfill(11) # Garante 11 digitos
        
        data_str = str(data_nascimento).strip()
        if "/" in data_str:
            partes = data_str.split("/")
            dia = partes[0].zfill(2)
            mes = partes[1].zfill(2)
            ano = partes[2]
            
            if len(ano) == 2:
                ano_int = int(ano)
                if ano_int > datetime.datetime.now().year % 100:
                    ano = f"19{ano}"
                else:
                    ano = f"20{ano}"
                    
            data_formatada = f"{ano}-{mes}-{dia}"
        else:
            data_formatada = data_str.split(" ")[0]
            
        payload = {
            "token": API_TOKEN,
            "cpf": cpf_limpo,
            "birthdate": data_formatada
        }
        
        response = requests.post(URL_API, data=payload, timeout=30)
        res = response.json()
        
        if res.get("code") == 200:
            dados = res["data"][0]
            situacao = dados.get("situacao_cadastral", "Desconhecido")
            obito = dados.get("ano_obito", "Nenhum")
            if not obito: obito = "Nenhum"
            return situacao, obito, "Sucesso"
        else:
            erro_msg = res.get("code_message", "Erro")
            erros_array = res.get("errors", [])
            return "Erro", "Erro", f"{erro_msg} - {erros_array}"
            
    except Exception as e:
        return "Erro", "Erro", str(e)

def processar_planilha(caminho_entrada, caminho_saida):
    print(f"Lendo planilha: {caminho_entrada}")
    
    try:
        df = pd.read_excel(caminho_entrada)
    except Exception as e:
        print(f"Erro ao ler planilha: {e}")
        return
    
    situacoes = []
    obitos = []
    status_robo = []
    
    for index, row in df.iterrows():
        cpf = str(row.get('COD_ID_FEDER', ''))
        data_nasc = str(row.get('DAT_NASCIMENTO', ''))
        nome = str(row.get('NOM_PESSOA_FISIC', ''))
        
        if cpf.strip() == '' or cpf.strip() == 'nan':
            situacoes.append("Sem CPF")
            obitos.append("Sem CPF")
            status_robo.append("Ignorado")
            continue
            
        situacao, obito, status = consultar_infosimples(cpf, data_nasc)
        
        situacoes.append(situacao)
        obitos.append(obito)
        status_robo.append(status)
        
    df['Situação Receita'] = situacoes
    df['Registro de Óbito'] = obitos
    df['Status da Consulta'] = status_robo
    
    print("\nSalvando arquivo de resultado...")
    df.to_excel(caminho_saida, index=False, engine='openpyxl')
    
    wb = load_workbook(caminho_saida)
    ws = wb.active
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    
    for col_num, column_title in enumerate(ws[1], 1):
        column_title.fill = header_fill
        column_title.font = header_font
        column_title.alignment = Alignment(horizontal="center", vertical="center")
        
        letter = get_column_letter(col_num)
        ws.column_dimensions[letter].width = max(len(str(column_title.value)) + 5, 20)
        
    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            if ws.cell(row=1, column=col).value == "Registro de Óbito":
                if cell.value not in ["Nenhum", "Erro", "Sem CPF"]:
                    cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                    cell.font = Font(color="FFFFFF", bold=True)
    
    wb.save(caminho_saida)

if __name__ == "__main__":
    entrada = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\afastados\colaboradores_afastados.xlsx"
    saida = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\Python Antigravity Agent\robo_afastados\resultado_afastados_atualizado.xlsx"
    processar_planilha(entrada, saida)
