import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta

def calcular_avos_ferias(inicio, fim):
    if pd.isna(inicio) or pd.isna(fim):
        return 0
    delta = relativedelta(fim, inicio)
    avos = delta.years * 12 + delta.months
    if delta.days >= 14:
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
            dias_no_mes = demissao.day
            if current_date.month == data_inicio.month and data_inicio.month == demissao.month:
                dias_no_mes = (demissao - data_inicio).days + 1
            if dias_no_mes >= 15:
                avos += 1
        else:
            if current_date.month == data_inicio.month:
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

def fator_faltas_ferias(faltas):
    if pd.isna(faltas) or faltas <= 5:
        return 1.0
    elif 6 <= faltas <= 14:
        return 24 / 30.0
    elif 15 <= faltas <= 23:
        return 18 / 30.0
    elif 24 <= faltas <= 32:
        return 12 / 30.0
    else:
        return 0.0

def inss_2024(base):
    if base <= 0: return 0.0, 0.0
    if base <= 1412.00:
        return base * 0.075, 7.5
    elif base <= 2666.68:
        return (base * 0.09) - 21.18, 9.0
    elif base <= 4000.03:
        return (base * 0.12) - 101.18, 12.0
    elif base <= 7786.02:
        return (base * 0.14) - 181.18, 14.0
    else:
        return (7786.02 * 0.14) - 181.18, 14.0

def irrf_2024(base):
    if base <= 2259.20:
        return 0.0, 0.0
    elif base <= 2826.65:
        return (base * 0.075) - 169.44, 7.5
    elif base <= 3751.05:
        return (base * 0.15) - 381.44, 15.0
    elif base <= 4664.68:
        return (base * 0.225) - 662.77, 22.5
    else:
        return (base * 0.275) - 896.00, 27.5

def calcular_rescisao(row):
    # Inputs
    salario = float(row.get('Salario_Contratual', 0))
    tipo = str(row.get('Tipo_Rescisao', '')).strip().upper()
    admissao = row.get('Data_Admissao')
    demissao = row.get('Data_Demissao')
    inicio_aquisitivo = row.get('Inicio Período Aquisitivo')
    faltas = float(row.get('Nº de Faltas', 0))
    if pd.isna(faltas): faltas = 0
    tem_ferias_vencidas = str(row.get('Ferias_Vencidas', 'Não')).strip().upper() == 'SIM'
    
    calcula_aviso = False
    paga_ferias_prop = True
    paga_13_prop = True
    dias_aviso = 0
    
    if tipo == 'DESLIGAMENTO SEM JUSTA CAUSA POR INICIATIVA DA EMPRESA':
        calcula_aviso = True
        anos = relativedelta(demissao, admissao).years
        dias_aviso = min(30 + (anos * 3), 90)
    elif tipo == 'DESLIGAMENTO POR JUSTA CAUSA':
        paga_ferias_prop = False
        paga_13_prop = False

    fator_ferias = fator_faltas_ferias(faltas)
    eventos = []
    
    def add_evento(codigo, descricao, tipo_evento, valor, ref):
        if valor > 0 or tipo_evento == 'Resultado':
            eventos.append({
                'CÓDIGO DO EVENTO': codigo,
                'DESCRIÇÃO DO EVENTO': descricao,
                'TIPO DE EVENTO': tipo_evento,
                'VALOR EVENTO': round(valor, 2),
                'VALOR EVENTO REFERÊNCIA': round(ref, 2) if ref else np.nan
            })

    # 721: Saldo de Salario
    dias_saldo = demissao.day if not pd.isna(demissao) else 0
    v_721 = (salario / 30.0) * dias_saldo
    add_evento(721, 'Saldo de Salario', 'PROVENTO', v_721, dias_saldo)
    
    # 20205 & 725: Aviso Previo Indenizado
    v_20205 = dias_aviso if calcula_aviso else 0
    v_725 = (salario / 30.0) * v_20205
    if v_725 > 0:
        add_evento(725, 'Aviso Previo Indenizado', 'PROVENTO', v_725, v_20205)
        add_evento(20205, 'Num Dias Aviso Previo Indenizado', 'RESULTADO', v_20205, v_20205)
    
    # Ferias
    v_20260 = 0
    if paga_ferias_prop:
        fim_projecao = demissao + relativedelta(days=v_20205) if demissao else None
        v_20260 = calcular_avos_ferias(inicio_aquisitivo, fim_projecao)
        v_738 = (salario / 12.0) * v_20260 * fator_ferias
        if v_738 > 0:
            add_evento(738, 'Ferias Proporcionais', 'PROVENTO', v_738, v_20260)
            add_evento(748, '1/3 Ferias Proporcionais', 'PROVENTO', v_738 / 3.0, np.nan)
            add_evento(20260, 'Num Meses Ferias Proporcionais (Avos de Férias)', 'RESULTADO', v_20260, v_20260)
        
    if tem_ferias_vencidas:
        v_737 = salario * fator_ferias
        add_evento(737, 'Ferias Vencidas', 'PROVENTO', v_737, 30.0)
        add_evento(747, '1/3 Ferias Vencidas', 'PROVENTO', v_737 / 3.0, np.nan)
        
    # 756: 13o Proporcional
    if paga_13_prop:
        fim_projecao = demissao + relativedelta(days=v_20205) if demissao else None
        avos_13 = calcular_avos_13(admissao, fim_projecao)
        v_756 = (salario / 12.0) * avos_13
        if v_756 > 0:
            add_evento(756, '13 Salario Proporcional', 'PROVENTO', v_756, avos_13)
    
    # 5013: INSS
    v_5013, aliquota_inss = inss_2024(v_721)
    if v_5013 > 0:
        add_evento(5013, 'Inss Mes', 'DESCONTO', v_5013, aliquota_inss)
    
    # 5020: IRRF
    base_irrf = v_721 - v_5013
    base_irrf_simplificada = v_721 - 564.80
    irrf_legal, aliq_leg = irrf_2024(base_irrf)
    irrf_simp, aliq_simp = irrf_2024(base_irrf_simplificada)
    if irrf_legal < irrf_simp and irrf_legal > 0:
        add_evento(5020, 'Irrf Mes', 'DESCONTO', irrf_legal, aliq_leg)
    elif irrf_simp > 0:
        add_evento(5020, 'Irrf Mes', 'DESCONTO', irrf_simp, aliq_simp)
    
    # Descontos Externos
    cencopay = float(row.get('Cencopay', 0)) if not pd.isna(row.get('Cencopay')) else 0.0
    if cencopay > 0: add_evento(5081, 'Desconto Gasto Cartão Virtual Benefício (Cencopay)', 'DESCONTO', cencopay, np.nan)
    
    goodcard = float(row.get('GoodCard', 0)) if not pd.isna(row.get('GoodCard')) else 0.0
    if goodcard > 0: add_evento(6280, 'Desc Cartão Benefícios Especiais (Goodcard)', 'DESCONTO', goodcard, np.nan)
    
    # FGTS
    saldo_fgts = float(row.get('Saldo_FGTS', 0)) if not pd.isna(row.get('Saldo_FGTS')) else 0.0
    fgts_mes = v_721 * 0.08
    fgts_aviso = v_725 * 0.08
    v_17190 = saldo_fgts + fgts_mes + fgts_aviso
    if v_17190 > 0:
        add_evento(17190, 'Base para fins rescisorios (FGTS Rescisorio)', 'RESULTADO', v_17190, np.nan)
        if tipo == 'DESLIGAMENTO SEM JUSTA CAUSA POR INICIATIVA DA EMPRESA':
            add_evento(11520, 'Fgts Multa Rescisao', 'RESULTADO', v_17190 * 0.40, 40.0)

    return eventos

def gerar_relatorio_final():
    print("Iniciando leitura das bases...")
    df_base = pd.read_excel('Conferência de Rescisão_Coletivo (1) (1).xlsb', sheet_name='BASE_FUNCIONARIOS', engine='pyxlsb')
    df_fopa = pd.read_excel('Conferência de Rescisão_Coletivo (1) (1).xlsb', sheet_name='FOPA_Eventos_por_Colaborador - ', engine='pyxlsb')
    
    df_base['Data_Admissao'] = pd.to_datetime(df_base['Data_Admissao'], errors='coerce')
    df_base['Data_Demissao'] = pd.to_datetime(df_base['Data_Demissao'], errors='coerce')
    df_base['Inicio Período Aquisitivo'] = pd.to_datetime(df_base['Inicio Período Aquisitivo'], errors='coerce')
    
    # Fazer merge com FOPA para trazer dados adicionais de cadastro (Centro de Custo, Cargo, etc)
    df_fopa['Matricula'] = df_fopa['Matrícula']
    df_merge = pd.merge(df_base, df_fopa, on='Matricula', how='left')
    
    print("Processando cálculos e gerando eventos...")
    linhas_relatorio = []
    
    for _, row in df_merge.iterrows():
        eventos = calcular_rescisao(row)
        
        empresa_nome = str(row.get('Código da Empresa', '')) + ' - ' + str(row.get('Nome da Empresa', ''))
        cc = str(row.get('Código Centro Custo', '')) + ' - ' + str(row.get('Centro Custo', ''))
        cargo = str(row.get('Código Cargo', '')) + ' - ' + str(row.get('Cargo', ''))
        sit = row.get('Tipo Situação', '')
        
        for ev in eventos:
            linhas_relatorio.append({
                'EMPRESA': empresa_nome if empresa_nome != 'nan - nan' else np.nan,
                'LOTAÇÃO TRIBUTÁRIA': np.nan,
                'MATRICULA COLABORADOR': row['Matricula'],
                'COLABORADOR': row['Nome'],
                'DATA DE ADMISSÃO': row['Data_Admissao'].strftime('%d/%m/%Y') if not pd.isna(row['Data_Admissao']) else np.nan,
                'SITUAÇÃO': sit,
                'DATA INÍCIO SITUAÇÃO': row['Data_Demissao'].strftime('%d/%m/%Y') if not pd.isna(row['Data_Demissao']) else np.nan,
                'CENTRO DE CUSTOS': cc if cc != 'nan - nan' else np.nan,
                'UNIDADE ORGANIZACIONAL': np.nan,
                'CARGO': cargo if cargo != 'nan - nan' else np.nan,
                'CODIGO FOLHA': 5,
                'FOLHA': 'Rescisao',
                'PERÍODO DA FOLHA DE PAGAMENTO': np.nan,
                'DATA DE PAGAMENTO': np.nan,
                'TIPO DE EVENTO': ev['TIPO DE EVENTO'],
                'CÓDIGO DO EVENTO': ev['CÓDIGO DO EVENTO'],
                'DESCRIÇÃO DO EVENTO': ev['DESCRIÇÃO DO EVENTO'],
                'VALOR EVENTO': ev['VALOR EVENTO'],
                'VALOR EVENTO REFERÊNCIA': ev['VALOR EVENTO REFERÊNCIA']
            })
            
    df_final = pd.DataFrame(linhas_relatorio)
    
    print("Gerando relatório final no formato Folha de Rescisão...")
    df_final.to_excel('Relatorio_Calculo_Rescisao.xlsx', index=False)
    print("Finalizado! Arquivo salvo como 'Relatorio_Calculo_Rescisao.xlsx'")

if __name__ == '__main__':
    gerar_relatorio_final()
