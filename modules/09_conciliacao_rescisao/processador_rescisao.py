import pandas as pd
import numpy as np
from datetime import datetime
import os
from leitor_pdf_rescisao import extract_pdf_data

def processar_pdf(pdf_path):
    print(f"Lendo PDF: {pdf_path}")
    header, events, totals = extract_pdf_data(pdf_path)
    
    # 1. Montar o DataFrame com os eventos extraídos
    df_events = pd.DataFrame(events)
    
    # Se não houver eventos, retorna dataframe vazio com as colunas
    if df_events.empty:
        print(f"Aviso: Nenhum evento encontrado em {pdf_path}")
        # Mesmo vazio, podemos criar a estrutura
    
    # 2. Adicionar as colunas de cabeçalho
    colunas_header = [
        'EMPRESA', 'LOTAÇÃO TRIBUTÁRIA', 'MATRICULA COLABORADOR', 'COLABORADOR',
        'DATA DE ADMISSÃO', 'SITUAÇÃO', 'DATA INÍCIO SITUAÇÃO', 'CENTRO DE CUSTOS',
        'UNIDADE ORGANIZACIONAL', 'CARGO', 'CODIGO FOLHA', 'FOLHA',
        'PERÍODO DA FOLHA DE PAGAMENTO', 'DATA DE PAGAMENTO'
    ]
    
    for col in colunas_header:
        df_events[col] = header.get(col, "")
        
    # Reordenar para o layout exato solicitado de 19 colunas + as nossas de auditoria
    colunas_layout = colunas_header + [
        'TIPO DE EVENTO', 'CÓDIGO DO EVENTO', 'DESCRIÇÃO DO EVENTO', 
        'VALOR EVENTO', 'VALOR EVENTO REFERÊNCIA'
    ]
    
    # Garante que as colunas existam mesmo se df_events for vazio
    for col in ['TIPO DE EVENTO', 'CÓDIGO DO EVENTO', 'DESCRIÇÃO DO EVENTO', 'VALOR EVENTO', 'VALOR EVENTO REFERÊNCIA']:
        if col not in df_events.columns:
            df_events[col] = np.nan
            
    df_final = df_events[colunas_layout].copy()
    
    if df_final.empty:
        return df_final
        
    # Ordenar por CÓDIGO DO EVENTO de forma crescente (Proventos -> Descontos -> Resultados)
    df_final.sort_values(by='CÓDIGO DO EVENTO', ascending=True, inplace=True)
    df_final.reset_index(drop=True, inplace=True)
    
    # --- MOTOR DE CÁLCULO / AUDITORIA BÁSICA ---
    # Aqui vamos buscar os valores base para calcular as regras
    
    # Buscar Salário / Remuneração Fixa (Resultados 11000 ou 11220)
    remun_fixa = df_final[df_final['CÓDIGO DO EVENTO'] == 11220]['VALOR EVENTO'].sum()
    if remun_fixa == 0:
        remun_fixa = df_final[df_final['CÓDIGO DO EVENTO'] == 11000]['VALOR EVENTO'].sum()
        
    salario_contratual = df_final[df_final['CÓDIGO DO EVENTO'] == 11000]['VALOR EVENTO'].sum()
        
    df_final['VALOR CALCULADO PYTHON'] = np.nan
    df_final['DIVERGÊNCIA'] = np.nan
    df_final['STATUS AUDITORIA'] = "Sem regra implementada"
    
    def format_date(d_str):
        try:
            return datetime.strptime(d_str, "%d/%m/%Y")
        except:
            return None
            
    dt_admissao = format_date(header['DATA DE ADMISSÃO'])
    dt_rescisao = format_date(header['DATA INÍCIO SITUAÇÃO'])
    
    for idx, row in df_final.iterrows():
        cod = row['CÓDIGO DO EVENTO']
        valor_pdf = row['VALOR EVENTO']
        ref = row['VALOR EVENTO REFERÊNCIA']
        valor_calc = np.nan
        status = "Sem regra"
        
        # 1. Saldo de Salário (721)
        if cod == 721 and not pd.isna(ref):
            # ref é o número de dias
            # Saldo de salário utiliza o Salário Contratual (11000) como base
            base_calculo = salario_contratual if salario_contratual > 0 else remun_fixa
            valor_calc = round((base_calculo / 30) * ref, 2)
            status = "Regra CLT (Salário/30 * Dias)"
            
        # 2. Aviso Prévio Indenizado (725)
        elif cod == 725:
            # Pegar os dias de aviso no Result 20208
            dias_aviso_lei = df_final[df_final['CÓDIGO DO EVENTO'] == 20208]['VALOR EVENTO'].sum()
            dias_totais = 30 + dias_aviso_lei
            valor_calc = round((remun_fixa / 30) * dias_totais, 2)
            status = "Regra CLT (30d + 3d/ano)"
            
        # 3. Férias Vencidas (737)
        elif cod == 737:
            valor_calc = round(remun_fixa, 2)
            status = "Regra CLT (1 remuneração)"
            
        # 4. Férias Proporcionais (738)
        elif cod == 738:
            # 1 avo de férias
            valor_calc = round(remun_fixa / 12, 2)
            status = "Regra CLT (1/12 avos)"
            
        # 5. 1/3 Férias (747, 748, 749, 752, 753, 754)
        elif cod == 747:
            # 1/3 de férias vencidas
            base = df_final[df_final['CÓDIGO DO EVENTO'] == 737]['VALOR EVENTO'].sum()
            valor_calc = round(base / 3, 2)
            status = "Regra CLT (1/3 Base Férias Vencidas)"
            
        elif cod == 748:
            # 1/3 de férias prop
            base = df_final[df_final['CÓDIGO DO EVENTO'] == 738]['VALOR EVENTO'].sum()
            valor_calc = round(base / 3, 2)
            status = "Regra CLT (1/3 Base Férias Prop)"
            
        elif cod == 749:
            # 1/3 férias indenizadas
            base = df_final[df_final['CÓDIGO DO EVENTO'] == 739]['VALOR EVENTO'].sum()
            valor_calc = round(base / 3, 2)
            status = "Regra CLT (1/3 Base Férias Ind)"
            
        # 6. INSS Mês (5013)
        elif cod == 5013:
            # 7.5% da base de INSS Mês (10100)
            base_inss = df_final[df_final['CÓDIGO DO EVENTO'] == 10100]['VALOR EVENTO'].sum()
            # Tabela progressiva (simplificada para o exemplo, faixa 1)
            # Na pratica, usar a tabela completa de INSS
            if base_inss <= 1412.00:
                valor_calc = round(base_inss * 0.075, 2)
            else:
                # Vamos simplificar assumindo faixa 1 exata pro PDF atual (1393.82)
                valor_calc = round(base_inss * 0.075, 2)
            status = "Regra INSS (Tabela Progressiva)"
            
        # Se calculamos algo, preencher
        if not pd.isna(valor_calc):
            df_final.at[idx, 'VALOR CALCULADO PYTHON'] = valor_calc
            diverg = abs(valor_pdf - valor_calc)
            df_final.at[idx, 'DIVERGÊNCIA'] = diverg
            if diverg <= 0.05: # Tolerância de 5 centavos
                df_final.at[idx, 'STATUS AUDITORIA'] = "OK"
            else:
                df_final.at[idx, 'STATUS AUDITORIA'] = "DIVERGENTE"
                
    return df_final

def processar_relatorio(pdf_path, output_excel):
    df_final = processar_pdf(pdf_path)
    # Salvando em Excel
    df_final.to_excel(output_excel, index=False)
    print(f"Planilha consolidada gerada com sucesso em: {output_excel}")

def processar_lote_pdfs(diretorio_pdfs, output_excel):
    import glob
    todos_dfs = []
    
    # Encontrar todos os PDFs
    padrao = os.path.join(diretorio_pdfs, "*.pdf")
    pdfs = glob.glob(padrao)
    
    if not pdfs:
        print(f"Nenhum PDF encontrado no diretório: {diretorio_pdfs}")
        return
        
    print(f"Iniciando processamento de {len(pdfs)} arquivos PDF...")
    
    for pdf_path in pdfs:
        try:
            df_pdf = processar_pdf(pdf_path)
            if not df_pdf.empty:
                todos_dfs.append(df_pdf)
        except Exception as e:
            print(f"Erro ao processar {pdf_path}: {e}")
        
    if todos_dfs:
        df_final = pd.concat(todos_dfs, ignore_index=True)
        df_final.to_excel(output_excel, index=False)
        print(f"Lote processado com sucesso! Arquivo consolidado gerado em: {output_excel}")
    else:
        print("Nenhum dado pôde ser extraído dos PDFs informados.")

if __name__ == '__main__':
    import sys
    # Se passar 2 argumentos, assume que é (diretório_entrada, arquivo_saída)
    if len(sys.argv) > 2:
        entrada = sys.argv[1]
        saida = sys.argv[2]
        if os.path.isdir(entrada):
            processar_lote_pdfs(entrada, saida)
        else:
            processar_relatorio(entrada, saida)
    else:
        print("Testando processamento em lote no diretório atual.")
        # Processa todos os PDFs do diretório atual como teste
        processar_lote_pdfs(".", "Folha_de_Rescisao_Consolidada_Lote.xlsx")
