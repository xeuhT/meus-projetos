import requests
import json
import sys
import os

API_TOKEN = os.getenv("INFOSIMPLES_TOKEN", "")

def consultar_cpf(cpf: str, data_nascimento: str):
    url = "https://api.infosimples.com/api/v2/consultas/receita-federal/cpf"
    
    cpf_limpo = "".join(filter(str.isdigit, cpf))
    
    # Converte DD/MM/AAAA para AAAA-MM-DD para a API
    if "/" in data_nascimento:
        dia, mes, ano = data_nascimento.split("/")
        data_formatada = f"{ano}-{mes}-{dia}"
    else:
        data_formatada = data_nascimento
    
    payload = {
        "token": API_TOKEN,
        "cpf": cpf_limpo,
        "birthdate": data_formatada
    }
    
    print(f"\n[INFO] Consultando CPF {cpf_limpo}...")
    
    try:
        response = requests.post(url, data=payload, timeout=30)
        resposta_json = response.json()
        
        codigo_retorno = resposta_json.get("code")
        mensagem = resposta_json.get("code_message")
        
        if codigo_retorno == 200:
            dados = resposta_json.get("data", [{}])[0]
            
            nome = dados.get("nome")
            situacao = dados.get("situacao_cadastral")
            ano_obito = dados.get("ano_obito")
            
            print("\n" + "="*45)
            print(" RESULTADO DA CONSULTA NA RECEITA FEDERAL")
            print("="*45)
            print(f"Nome:               {nome}")
            print(f"Situação Cadastral: {situacao}")
            print(f"Ano de Óbito:       {ano_obito if ano_obito else 'Nenhum'}")
            print("="*45)
            return dados
        else:
            erros = resposta_json.get("errors", [])
            print(f"[ERRO] {mensagem} | Detalhes: {erros}")
            return None
            
    except Exception as e:
        print(f"[ERRO DE CONEXÃO] {str(e)}")
        return None

if __name__ == "__main__":
    if len(sys.argv) == 3:
        cpf_input = sys.argv[1]
        nasc_input = sys.argv[2]
    else:
        print("--- TESTE DE CONSULTA CPF (INFOSIMPLES) ---")
        cpf_input = input("Digite o CPF: ").strip()
        nasc_input = input("Digite a Data de Nascimento (DD/MM/AAAA): ").strip()
        
    if cpf_input and nasc_input:
        consultar_cpf(cpf_input, nasc_input)
    else:
        print("CPF ou Data de Nascimento não informados.")
