import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta

def preparar_dataframe(df):
    colunas_data = ['Data_Admissao', 'Data_Demissao', 'Inicio Período Aquisitivo']
    for col in colunas_data:
        df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

def classificar_direitos_rescisao(row):
    if pd.isna(row['Data_Admissao']) or pd.isna(row['Data_Demissao']):
        return pd.Series({'calcula_aviso': False, 'calcula_multa_fgts': False, 'paga_ferias_prop': False, 'paga_13_prop': False, 'dias_aviso_indenizado': 0})
        
    tipo = str(row['Tipo_Rescisao']).strip().upper()
    
    direitos = {
        'calcula_aviso': False,
        'calcula_multa_fgts': False,
        'paga_ferias_prop': True,
        'paga_13_prop': True,
        'dias_aviso_indenizado': 0
    }

    if tipo == 'DESLIGAMENTO SEM JUSTA CAUSA POR INICIATIVA DA EMPRESA':
        direitos['calcula_aviso'] = True
        direitos['calcula_multa_fgts'] = True
        
        anos_trabalhados = relativedelta(row['Data_Demissao'], row['Data_Admissao']).years
        direitos['dias_aviso_indenizado'] = min(30 + (anos_trabalhados * 3), 90)

    elif tipo == 'PEDIDO DE DEMISSÃO':
        pass 

    elif tipo == 'DESLIGAMENTO POR JUSTA CAUSA':
        direitos['paga_ferias_prop'] = False
        direitos['paga_13_prop'] = False

    elif tipo == 'TÉRMINO DE CONTRATO DE EXPERIÊNCIA':
        pass 

    return pd.Series(direitos)

def calcular_avos_ferias(inicio, fim):
    if pd.isna(inicio) or pd.isna(fim):
        return 0
    delta = relativedelta(fim, inicio)
    avos = delta.years * 12 + delta.months
    if delta.days >= 14: # Considera 15 dias trabalhados no mês incompleto como 1 avo
        avos += 1
    return min(avos, 12)

def calcular_avos_13(admissao, demissao):
    if pd.isna(admissao) or pd.isna(demissao):
        return 0
    inicio_ano = datetime(demissao.year, 1, 1)
    data_inicio = max(admissao, inicio_ano)
    
    avos = 0
    current_date = data_inicio
    while current_date.month <= demissao.month:
        if current_date.month == demissao.month:
            # último mês
            dias_no_mes = demissao.day
            if current_date.month == data_inicio.month and data_inicio.month == demissao.month:
                # começou e terminou no mesmo mês
                dias_no_mes = (demissao - data_inicio).days + 1
            if dias_no_mes >= 15:
                avos += 1
        else:
            if current_date.month == data_inicio.month:
                # primeiro mês (se não for o último)
                # quantos dias trabalhou no mês de início?
                # calcular o último dia do mês
                next_month = current_date.replace(day=28) + relativedelta(days=4)
                last_day = next_month - relativedelta(days=next_month.day)
                dias_no_mes = (last_day - current_date).days + 1
                if dias_no_mes >= 15:
                    avos += 1
            else:
                avos += 1
        current_date += relativedelta(months=1)
        current_date = current_date.replace(day=1)
    
    return min(avos, 12)

def calcular_motor_financeiro(row):
    salario = row['Salario_Contratual']
    
    # 1. Saldo de Salário
    # Considera os dias trabalhados no mês da demissão
    dias_saldo = row['Data_Demissao'].day
    valor_saldo_salario = (salario / 30.0) * dias_saldo
    
    # 2. Aviso Prévio Indenizado
    valor_aviso_previo = 0.0
    if row['calcula_aviso']:
        valor_aviso_previo = (salario / 30.0) * row['dias_aviso_indenizado']
        
    # 3. Férias Proporcionais
    valor_ferias_prop = 0.0
    valor_terco_ferias = 0.0
    if row['paga_ferias_prop']:
        avos_ferias = calcular_avos_ferias(row['Inicio Período Aquisitivo'], row['Data_Demissao'])
        # Adicionar projeção do aviso nas férias (se indenizado)
        if row['calcula_aviso'] and row['dias_aviso_indenizado'] > 0:
            data_projecao = row['Data_Demissao'] + relativedelta(days=row['dias_aviso_indenizado'])
            avos_ferias = calcular_avos_ferias(row['Inicio Período Aquisitivo'], data_projecao)
        
        valor_ferias_prop = (salario / 12.0) * avos_ferias
        valor_terco_ferias = valor_ferias_prop / 3.0

    # 4. 13º Salário Proporcional
    valor_13_prop = 0.0
    if row['paga_13_prop']:
        avos_13 = calcular_avos_13(row['Data_Admissao'], row['Data_Demissao'])
        # Adicionar projeção do aviso no 13º
        if row['calcula_aviso'] and row['dias_aviso_indenizado'] > 0:
            data_projecao = row['Data_Demissao'] + relativedelta(days=row['dias_aviso_indenizado'])
            avos_13 = calcular_avos_13(row['Data_Admissao'], data_projecao)
            
        valor_13_prop = (salario / 12.0) * avos_13

    return pd.Series({
        'Valor_Calculado_Saldo_Salario': round(valor_saldo_salario, 2),
        'Valor_Calculado_Ferias_Prop': round(valor_ferias_prop, 2),
        'Valor_Calculado_13_Terco': round(valor_terco_ferias, 2),
        'Valor_Calculado_13_Prop': round(valor_13_prop, 2),
        'Valor_Calculado_Aviso_Previo': round(valor_aviso_previo, 2)
    })

if __name__ == '__main__':
    # Ler os dados gerados
    df_base = pd.read_excel('base_funcionarios_mock.xlsx')
    df_base = preparar_dataframe(df_base)
    
    # Classificar direitos
    direitos = df_base.apply(classificar_direitos_rescisao, axis=1)
    df_base = pd.concat([df_base, direitos], axis=1)
    
    # Rodar Motor Financeiro
    financeiro = df_base.apply(calcular_motor_financeiro, axis=1)
    df_base = pd.concat([df_base, financeiro], axis=1)
    
    # Exibir resultados
    pd.set_option('display.max_columns', None)
    print("=== Resultado do Motor Financeiro ===")
    print(df_base[['Nome', 'Valor_Calculado_Saldo_Salario', 'Valor_Calculado_Ferias_Prop', 'Valor_Calculado_13_Terco', 'Valor_Calculado_13_Prop', 'Valor_Calculado_Aviso_Previo']])
    
    # Exportar para conferência
    df_base.to_excel('resultado_motor_financeiro.xlsx', index=False)
    print("\nArquivo 'resultado_motor_financeiro.xlsx' gerado com os cálculos.")
