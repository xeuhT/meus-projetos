import pandas as pd
import numpy as np
import sys
import re
import os
import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter

# Mapeamento consolidado das 317 Lojas / Unidades Cencosud para o respectivo Sindicato Preponderante (Sind X Lojas)
MAPA_LOJA_SIND = {
    1: 1, 3: 1, 4: 1, 5: 1, 6: 1, 7: 1, 8: 1, 9: 1, 10: 1, 12: 1, 13: 1, 14: 1, 19: 1, 20: 1, 21: 1, 22: 1, 
    23: 1, 24: 1, 25: 1, 26: 1, 27: 1, 28: 1, 30: 1, 33: 1, 34: 1, 78: 1, 90: 1, 99: 1, 101: 1, 114: 1, 
    118: 1, 199: 1, 200: 1, 201: 1, 202: 1, 203: 1, 204: 1, 205: 1, 206: 1, 225: 1, 226: 1, 229: 1, 232: 1, 
    233: 1, 240: 1, 255: 1, 256: 1, 258: 1, 276: 1, 288: 1, 320: 1, 340: 1, 2: 1, 11: 1, 37: 6, 38: 6, 
    39: 6, 165: 6, 207: 6, 210: 6, 211: 6, 237: 6, 241: 6, 242: 6, 251: 6, 280: 6, 36: 9, 59: 9, 161: 9, 
    212: 9, 224: 9, 230: 9, 306: 9, 15: 11, 16: 11, 17: 11, 74: 11, 76: 11, 228: 11, 31: 18, 35: 18, 41: 18, 
    71: 18, 88: 18, 102: 18, 106: 18, 108: 18, 153: 18, 170: 18, 272: 18, 302: 18, 42: 25, 44: 25, 46: 25, 
    64: 25, 70: 25, 73: 25, 95: 25, 119: 25, 338: 32, 80: 34, 81: 34, 83: 34, 84: 34, 85: 34, 86: 34, 
    96: 34, 116: 34, 117: 35, 32: 36, 231: 36, 141: 38, 185: 38, 186: 38, 196: 38, 198: 38, 278: 38, 
    125: 40, 133: 40, 238: 40, 151: 46, 270: 46, 279: 46, 284: 46, 154: 47, 156: 47, 149: 48, 274: 49, 
    275: 49, 289: 49, 104: 54, 164: 54, 172: 54, 400: 55, 174: 57, 208: 57, 221: 57, 223: 57, 18: 58, 
    213: 58, 250: 58, 259: 58, 283: 58, 315: 58, 112: 60, 139: 60, 266: 60, 131: 64, 243: 65, 175: 70, 
    176: 70, 222: 70, 263: 70, 265: 70, 169: 74, 334: 76, 262: 77, 260: 78, 171: 80, 218: 80, 261: 80, 
    281: 81, 209: 82, 236: 83, 162: 84, 163: 84, 254: 85, 239: 87, 214: 89, 166: 90, 219: 90, 168: 91, 
    177: 91, 244: 91, 245: 91, 246: 91, 247: 91, 248: 91, 249: 91, 257: 91, 264: 91, 267: 91, 311: 91, 
    220: 92, 216: 94, 217: 94, 234: 94, 235: 95, 252: 95, 271: 96, 40: 97, 51: 97, 277: 97, 566: 101, 
    573: 101, 626: 101, 539: 106, 524: 108, 137: 109, 634: 111, 635: 112, 516: 127, 525: 127, 527: 127, 
    531: 127, 538: 127, 541: 127, 545: 127, 551: 127, 552: 127, 575: 127, 615: 127, 619: 127, 631: 127, 
    632: 127, 633: 127, 640: 127, 644: 127, 648: 127, 651: 127, 655: 127, 656: 127, 658: 127, 662: 127, 
    554: 138, 700: 200, 702: 200, 703: 200, 704: 200, 705: 200, 706: 200, 707: 200, 708: 200, 709: 200, 
    710: 200, 711: 200, 712: 200, 714: 200, 715: 200, 716: 200, 718: 200, 719: 200, 720: 200, 723: 200, 
    724: 200, 726: 200, 728: 200, 729: 200, 731: 200, 732: 200, 733: 200, 734: 200, 736: 200, 742: 200, 
    743: 200, 744: 200, 746: 200, 747: 200, 748: 200, 750: 200, 1772: 200, 1773: 200, 1774: 200, 1777: 200, 
    1779: 200, 1780: 200, 1781: 200, 1782: 200, 1783: 200, 1784: 200, 1785: 200, 1786: 200, 1788: 200, 
    725: 201, 713: 202, 717: 202, 722: 204, 730: 204, 1789: 204, 721: 205, 745: 226, 806: 1, 802: 11, 
    815: 11, 801: 18, 813: 18, 814: 18, 816: 18, 817: 18, 810: 42, 909: 300, 905: 301, 906: 301, 910: 301, 
    915: 301, 900: 302, 901: 302, 903: 302, 904: 302, 908: 302, 913: 302, 950: 302, 902: 304, 907: 304, 
    911: 304, 912: 304, 916: 306, 850: 44, 855: 44, 857: 44, 860: 44, 861: 44, 862: 44, 864: 44, 865: 44, 
    870: 44, 863: 72
}

def extrair_sind_da_lotacao(lotacao_str):
    if not lotacao_str or pd.isna(lotacao_str):
        return None
    # Isolar a identificação da loja antes do CNPJ para evitar falsos positivos
    texto_loja = str(lotacao_str).split('CNPJ')[0]
    nums = [int(n) for n in re.findall(r'\b\d+\b', texto_loja)]
    for n in nums:
        if n in MAPA_LOJA_SIND:
            return MAPA_LOJA_SIND[n]
    return None

def calcular_inss_progressivo(base):
    if base <= 0: return 0.0
    inss = 0.0
    if base <= 1621.00:
        return round(base * 0.075, 2)
    else:
        inss += round(1621.00 * 0.075, 2)
        
    if base <= 2902.84:
        inss += round((base - 1621.00) * 0.09, 2)
        return round(inss, 2)
    else:
        inss += round((2902.84 - 1621.00) * 0.09, 2)
        
    if base <= 4354.27:
        inss += round((base - 2902.84) * 0.12, 2)
        return round(inss, 2)
    else:
        inss += round((4354.27 - 2902.84) * 0.12, 2)
        
    if base <= 8475.55:
        inss += round((base - 4354.27) * 0.14, 2)
        return round(inss, 2)
    else:
        inss += round((8475.55 - 4354.27) * 0.14, 2)
        return round(inss, 2)

def auditar_folha_massiva(caminho_entrada, caminho_saida_xlsx, logger=None, progress_fn=None):
    def log(msg):
        if logger:
            logger(msg)
        else:
            print(msg)
            
    log(f"Lendo base massiva: {caminho_entrada}")
    if progress_fn: progress_fn(0.05)
    try:
        ext = caminho_entrada.lower().split('.')[-1]
        if ext == 'xlsb':
            df = pd.read_excel(caminho_entrada, engine='pyxlsb')
        elif ext == 'xlsx':
            df = pd.read_excel(caminho_entrada)
        elif ext == 'csv':
            df = pd.read_csv(caminho_entrada, sep=';', encoding='utf-8-sig', decimal=',', thousands='.', quotechar='"')
            # Mapeamento para garantir nomes exatos caso haja problemas de acentuação
            cols_map = {}
            for c in df.columns:
                if 'C' in c and 'DIGO DO EVENTO' in c: cols_map[c] = 'CÓDIGO DO EVENTO'
                elif 'DATA DE ADMISS' in c: cols_map[c] = 'DATA DE ADMISSÃO'
                elif 'DATA IN' in c and 'CIO SITUA' in c: cols_map[c] = 'DATA INÍCIO SITUAÇÃO'
                elif 'VALOR EVENTO REFER' in c: cols_map[c] = 'VALOR EVENTO REFERÊNCIA'
            df.rename(columns=cols_map, inplace=True)
            
            # Garantir tipos numéricos (caso pandas leia como object se houver strings sujas)
            for col in ['VALOR EVENTO', 'VALOR EVENTO REFERÊNCIA']:
                if col in df.columns and df[col].dtype == object:
                    df[col] = df[col].astype(str).str.replace('.', '').str.replace(',', '.').replace('nan', '0').replace('None', '0')
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        else:
            log(f"Formato não suportado: {ext}")
            return False
            
    except Exception as e:
        log(f"Erro ao ler arquivo: {e}")
        return False

    colunas_data = ['DATA DE ADMISSÃO', 'DATA INÍCIO SITUAÇÃO', 'DATA DE PAGAMENTO']
    for col in colunas_data:
        if col in df.columns:
            # Em CSV as datas podem ja vir como string 'DD/MM/YYYY'. O xlsb traz como inteiros (dias desde 1900).
            if df[col].dtype in [np.float64, np.int64] or df[col].apply(lambda x: isinstance(x, (int, float))).any():
                df[col] = pd.to_datetime(pd.to_numeric(df[col], errors='coerce'), unit='D', origin='1899-12-30').dt.strftime('%d/%m/%Y')

    df['VALOR CALCULADO PYTHON'] = np.nan
    df['DIVERGÊNCIA'] = np.nan
    df['STATUS AUDITORIA'] = "Sem regra implementada"

    if 'CÓDIGO DO EVENTO' not in df.columns:
        log("Erro: Coluna CÓDIGO DO EVENTO não encontrada.")
        return False

    df['CÓDIGO DO EVENTO'] = pd.to_numeric(df['CÓDIGO DO EVENTO'], errors='coerce').fillna(0).astype(int)
    
    # Preencher NaN com 0 nas colunas numéricas
    df['VALOR EVENTO'] = pd.to_numeric(df['VALOR EVENTO'], errors='coerce').fillna(0)
    df['VALOR EVENTO REFERÊNCIA'] = pd.to_numeric(df['VALOR EVENTO REFERÊNCIA'], errors='coerce').fillna(0)

    matriculas = df['MATRICULA COLABORADOR'].unique()
    total_colabs = len(matriculas)
    log(f"Base carregada: {len(df)} registros | {total_colabs} colaboradores a auditar.")
    
    linhas_extras_10990 = []
    
    for idx_c, matricula in enumerate(matriculas):
        if progress_fn and (idx_c % 20 == 0 or idx_c == total_colabs - 1):
            progress_fn(0.1 + 0.75 * ((idx_c + 1) / total_colabs))
            
        mask_colab = df['MATRICULA COLABORADOR'] == matricula
        df_colab = df[mask_colab]
        
        salario_contratual = df_colab[df_colab['CÓDIGO DO EVENTO'] == 11000]['VALOR EVENTO'].sum()
        remun_fixa = df_colab[df_colab['CÓDIGO DO EVENTO'] == 11220]['VALOR EVENTO'].sum()
        if remun_fixa == 0:
            remun_fixa = salario_contratual

        dias_direito_f = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20305]['VALOR EVENTO'].sum()
        if dias_direito_f == 0:
            dias_direito_f = 30.0

        # Validação de Faltas vs Data de Desligamento no Mês
        dt_deslig_str = str(df_colab['DATA INÍCIO SITUAÇÃO'].iloc[0]) if 'DATA INÍCIO SITUAÇÃO' in df_colab.columns and len(df_colab) > 0 else ''
        dia_deslig = None
        if dt_deslig_str and '/' in dt_deslig_str:
            try:
                dia_deslig = int(dt_deslig_str.split('/')[0])
            except Exception:
                dia_deslig = None

        qtd_faltas_tot = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20010]['VALOR EVENTO'].sum()
        if qtd_faltas_tot == 0 and (df_colab['CÓDIGO DO EVENTO'] == 5001).any():
            qtd_faltas_tot = df_colab[df_colab['CÓDIGO DO EVENTO'] == 5001]['VALOR EVENTO REFERÊNCIA'].sum()

        qtd_dsr_tot = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20016]['VALOR EVENTO'].sum()
        if qtd_dsr_tot == 0 and (df_colab['CÓDIGO DO EVENTO'] == 5007).any():
            qtd_dsr_tot = df_colab[df_colab['CÓDIGO DO EVENTO'] == 5007]['VALOR EVENTO REFERÊNCIA'].sum()

        total_ausencias = qtd_faltas_tot + qtd_dsr_tot
        alerta_faltas_excedentes = False
        msg_alerta_faltas = ""
        if dia_deslig is not None and total_ausencias > 0:
            max_permitido = max(0, dia_deslig - 1)
            if total_ausencias > max_permitido:
                alerta_faltas_excedentes = True
                msg_alerta_faltas = f"ALERTA CRÍTICO: Faltas+DSR ({total_ausencias:.0f}d: {qtd_faltas_tot:.0f} faltas + {qtd_dsr_tot:.0f} DSR) excede o limite máximo de {max_permitido:.0f}d para desligamento no dia {dia_deslig:02d}"

        for idx, row in df_colab.iterrows():
            cod = row['CÓDIGO DO EVENTO']
            valor_sistema = row['VALOR EVENTO']
            ref = row['VALOR EVENTO REFERÊNCIA']
            
            valor_calc = np.nan
            status = "Sem regra implementada"

            if cod == 721 and not pd.isna(ref) and ref > 0:
                base_calculo = salario_contratual if salario_contratual > 0 else remun_fixa
                teorico = round((base_calculo / 30) * ref, 2)
                
                valor_152 = df_colab[df_colab['CÓDIGO DO EVENTO'] == 152]['VALOR EVENTO'].sum()
                valor_137 = df_colab[df_colab['CÓDIGO DO EVENTO'] == 137]['VALOR EVENTO'].sum()
                valor_138 = df_colab[df_colab['CÓDIGO DO EVENTO'] == 138]['VALOR EVENTO'].sum()
                
                dias_aux = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20417]['VALOR EVENTO'].sum()
                valor_dias_aux = round((base_calculo / 30) * dias_aux, 2)
                
                abatimentos_possiveis = [
                    valor_152 + valor_137 + valor_138,
                    valor_152 + valor_dias_aux + valor_138,
                    valor_152 + valor_137,
                    valor_152 + valor_dias_aux,
                    valor_152,
                    0
                ]
                
                best_ab = 0
                min_div = float('inf')
                for ab in abatimentos_possiveis:
                    div = abs(valor_sistema - (teorico - ab))
                    if div < min_div:
                        min_div = div
                        best_ab = ab
                        
                valor_calc = teorico - best_ab
                status = "Regra CLT + Heurística de Abatimentos"
                
            elif cod == 5001:
                qtd_faltas = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20010]['VALOR EVENTO'].sum()
                if qtd_faltas == 0 and ref > 0:
                    qtd_faltas = ref
                valor_calc = round((remun_fixa / 30) * qtd_faltas, 2)
                status = f"Regra LG (11220/30 * {qtd_faltas} Faltas)"
                if alerta_faltas_excedentes:
                    status += f" | {msg_alerta_faltas}"
                
            elif cod == 5007:
                qtd_dsr = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20016]['VALOR EVENTO'].sum()
                if qtd_dsr == 0 and ref > 0:
                    qtd_dsr = ref
                valor_calc = round((remun_fixa / 30) * qtd_dsr, 2)
                status = f"Regra LG (11220/30 * {qtd_dsr} DSR)"
                if alerta_faltas_excedentes:
                    status += f" | {msg_alerta_faltas}"
                
            elif cod == 27:
                qtd_he = df_colab[df_colab['CÓDIGO DO EVENTO'] == 21000]['VALOR EVENTO'].sum()
                # A rubrica 11006 (Valor Hora) não vem no CSV. Inferimos Divisor 220.
                valor_hora = remun_fixa / 220
                valor_calc = round(valor_hora * qtd_he * 1.5, 2)
                status = f"Regra LG (Hora {valor_hora:.2f} * {qtd_he}h * 1.5)"
                
            elif cod == 60:
                qtd_dsr_he = df_colab[df_colab['CÓDIGO DO EVENTO'] == 21034]['VALOR EVENTO'].sum()
                he_events = [27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 151, 152, 153, 154, 155, 156, 157, 158, 170]
                soma_he = df_colab[df_colab['CÓDIGO DO EVENTO'].isin(he_events)]['VALOR EVENTO'].sum()
                
                # O evento 21033 (Dias Úteis) não vem no CSV. 
                # Revertendo a lógica padrão de calendário: Dias Úteis = 30 - Dias DSR (ou usar 25 que é o padrão comercial mensal)
                dias_uteis_inferido = 30 - qtd_dsr_he if qtd_dsr_he > 0 else 25 
                
                if dias_uteis_inferido > 0:
                    valor_calc = round((soma_he / dias_uteis_inferido) * qtd_dsr_he, 2)
                    status = f"Regra LG (Soma HE / {dias_uteis_inferido} * {qtd_dsr_he})"
                
            elif cod == 4:
                # Periculosidade (CLT 30% sobre Salário Contratual)
                pct = 0.30
                if ref in [30, 0.3]:
                    pct = 0.30
                dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20235]['VALOR EVENTO'].sum()
                if dias_trab == 0:
                    dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 721]['VALOR EVENTO REFERÊNCIA'].sum()
                if dias_trab == 0:
                    dias_trab = 30.0
                valor_calc = round(((salario_contratual * pct) / 30) * dias_trab, 2)
                status = f"Regra CLT (30% Sal.Contr / 30 * {dias_trab}d)"
                
            elif cod == 5:
                # Insalubridade (Grau dinâmico pela Referência: 10%, 20% ou 40% sobre Salário Mínimo R$ 1.621,00)
                grau = 0.20
                if ref in [10, 20, 40]:
                    grau = ref / 100.0
                elif 0 < ref <= 1.0:
                    grau = ref
                dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20235]['VALOR EVENTO'].sum()
                if dias_trab == 0:
                    dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 721]['VALOR EVENTO REFERÊNCIA'].sum()
                if dias_trab == 0:
                    dias_trab = 30.0
                base_sal_min = 1621.00
                valor_calc = round(((base_sal_min * grau) / 30) * dias_trab, 2)
                status = f"Regra CLT ({int(grau*100)}% Sal.Mín 1621.00 / 30 * {dias_trab}d)"
                
            elif cod in [11, 12, 13, 14, 15, 16, 17, 160, 161]:
                # Adicionais de Tempo de Serviço (CCT / Relações Sindicais)
                dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20235]['VALOR EVENTO'].sum()
                if dias_trab == 0:
                    dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 721]['VALOR EVENTO REFERÊNCIA'].sum()
                if dias_trab == 0:
                    dias_trab = 30.0

                # Quantidade de períodos
                qtd_periodos = 0
                if cod == 13:
                    qtd_periodos = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20273]['VALOR EVENTO'].sum()
                if qtd_periodos == 0 and not pd.isna(ref) and ref > 0:
                    qtd_periodos = ref
                    
                # Fallback por tempo de serviço (Admissão vs Rescisão)
                if qtd_periodos == 0:
                    try:
                        col_adm = 'DATA DE ADMISSÃO'
                        col_sit = 'DATA INÍCIO SITUAÇÃO'
                        if col_adm in df_colab.columns and col_sit in df_colab.columns:
                            d_adm = pd.to_datetime(df_colab[col_adm].iloc[0], errors='coerce', dayfirst=True)
                            d_sit = pd.to_datetime(df_colab[col_sit].iloc[0], errors='coerce', dayfirst=True)
                            if pd.notna(d_adm) and pd.notna(d_sit):
                                anos = (d_sit - d_adm).days / 365.25
                                step_anos = {11: 1, 12: 2, 13: 3, 14: 4, 15: 5}.get(cod, 1)
                                qtd_periodos = int(anos // step_anos)
                    except:
                        pass
                if qtd_periodos == 0:
                    qtd_periodos = 1

                # Obter sindicato do colaborador via lotação tributária
                lotacao_str = str(df_colab['LOTAÇÃO TRIBUTÁRIA'].iloc[0]) if 'LOTAÇÃO TRIBUTÁRIA' in df_colab.columns else ''
                sind_id = extrair_sind_da_lotacao(lotacao_str)

                # Regras específicas por sindicato / CCT
                if sind_id in [1, 30, 66, 69, 73, 97]: # Sergipe (7% do piso 1621.00 = R$ 113,47 por triênio)
                    val_trienio_mes = 113.47
                    qtd_aplicavel = min(qtd_periodos, 6)
                    valor_mensal = val_trienio_mes * qtd_aplicavel
                    valor_calc = round((valor_mensal / 30) * dias_trab, 2)
                    status = f"Regra CCT SE (Sind {sind_id}: R$ 113,47 * {qtd_aplicavel} triênios / 30 * {dias_trab}d)"
                    
                elif sind_id == 55: # Sind 55 (R$ 89,50 fixo por triênio)
                    valor_mensal = 89.50 * qtd_periodos
                    valor_calc = round((valor_mensal / 30) * dias_trab, 2)
                    status = f"Regra CCT (Sind 55: R$ 89,50 * {qtd_periodos} triênios / 30 * {dias_trab}d)"
                    
                elif sind_id == 11: # Sind 11 (7% por triênio)
                    taxa = 0.07 * qtd_periodos
                    valor_mensal = salario_contratual * taxa
                    valor_calc = round((valor_mensal / 30) * dias_trab, 2)
                    status = f"Regra CCT (Sind 11: 7% * {qtd_periodos} triênios / 30 * {dias_trab}d)"
                    
                elif sind_id == 60: # Sind 60 (3% por triênio, limitado a R$ 100,00 por triênio, máx 3)
                    taxa = 0.03
                    valor_unid = min(salario_contratual * taxa, 100.00)
                    qtd_aplicavel = min(qtd_periodos, 3)
                    valor_mensal = valor_unid * qtd_aplicavel
                    valor_calc = round((valor_mensal / 30) * dias_trab, 2)
                    status = f"Regra CCT (Sind 60: R$ {valor_unid:.2f} (teto 100) * {qtd_aplicavel} triênios / 30 * {dias_trab}d)"
                    
                elif sind_id == 96: # Sind 96 (4% por triênio)
                    taxa = 0.04 * qtd_periodos
                    valor_calc = round(((salario_contratual * taxa) / 30) * dias_trab, 2)
                    status = f"Regra CCT (Sind 96: 4% * {qtd_periodos} triênios / 30 * {dias_trab}d)"
                    
                elif sind_id == 127: # Bretas GO (4% Triênio / 6% Quinquênio não cumulativo)
                    if cod == 13:
                        taxa = 0.04
                        valor_calc = round(((salario_contratual * taxa) / 30) * dias_trab, 2)
                        status = f"Regra CCT (Sind 127: 4% Sal.Contr / 30 * {dias_trab}d)"
                    elif cod == 15:
                        taxa = 0.06
                        valor_calc = round(((salario_contratual * taxa) / 30) * dias_trab, 2)
                        status = f"Regra CCT (Sind 127: 6% Sal.Contr / 30 * {dias_trab}d)"
                    else:
                        taxa = 0.03 * qtd_periodos
                        valor_calc = round(((salario_contratual * taxa) / 30) * dias_trab, 2)
                        status = f"Regra CCT (Sind 127: {taxa*100:.0f}% Sal.Contr / 30 * {dias_trab}d)"
                        
                else: # Regra CCT Padrão (Anuênio 1%, Biênio 2%, Triênio 3%, Quadriênio 4%, Quinquênio 5%)
                    taxas_padrao = {11: 0.01, 12: 0.02, 13: 0.03, 14: 0.04, 15: 0.05}
                    taxa_unitaria = taxas_padrao.get(cod, 0.03)
                    taxa_total = taxa_unitaria * qtd_periodos
                    valor_calc = round(((salario_contratual * taxa_total) / 30) * dias_trab, 2)
                    status = f"Regra CCT Padrão ({int(taxa_unitaria*100)}% * {qtd_periodos} per / 30 * {dias_trab}d)"
                    
            elif cod in [6250, 6251, 5729, 5731, 5750, 5081, 6280, 490, 147]:
                valor_calc = valor_sistema
                status = "Valor Importado (Benefícios / Catraca)"
                
            elif cod in [300, 302, 325, 742, 743, 744, 5049, 5015]:
                valor_calc = valor_sistema
                status = "Valor Importado (Módulo Férias / Histórico 12 Meses)"
                
            elif cod in [6255, 6263, 6271, 6272, 6281]:
                valor_calc = valor_sistema
                status = "Valor Importado (Tabelas Operadora Saúde/Odonto)"
                
            elif cod in [5034, 5038]:
                valor_calc = valor_sistema
                status = "Valor Importado (Regra Sindical / Configuração HR)"
                
            elif cod == 96:
                # Quebra de Caixa (CCT Conforme Sindicato / Loja)
                dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20235]['VALOR EVENTO'].sum()
                if dias_trab == 0:
                    dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 721]['VALOR EVENTO REFERÊNCIA'].sum()
                if dias_trab == 0:
                    dias_trab = 30.0
                    
                lotacao_str = str(df_colab['LOTAÇÃO TRIBUTÁRIA'].iloc[0]) if 'LOTAÇÃO TRIBUTÁRIA' in df_colab.columns else ''
                sind_id = extrair_sind_da_lotacao(lotacao_str)
                sal_base = salario_contratual if salario_contratual > 0 else remun_fixa
                
                # Regras de Base Mensal por Sindicato
                if sind_id in [1, 66, 69, 97]:
                    base_mensal = 97.26
                    status = f"Regra CCT Sind {sind_id} (6% Sal.Mínimo R$ 97,26 / 30 * {dias_trab}d)"
                elif sind_id == 111:
                    base_mensal = 316.37
                    status = f"Regra CCT Sind 111 (Fixo R$ 316,37 / 30 * {dias_trab}d)"
                elif sind_id == 40:
                    # Regra de Tempo de Casa do Sind 40: R$ 141,10 (até 1 ano) e R$ 211,65 (acima de 1 ano)
                    anos = 0.0
                    try:
                        col_adm = 'DATA DE ADMISSÃO'
                        col_sit = 'DATA INÍCIO SITUAÇÃO'
                        if col_adm in df_colab.columns and col_sit in df_colab.columns:
                            d_adm = pd.to_datetime(df_colab[col_adm].iloc[0], errors='coerce', dayfirst=True)
                            d_sit = pd.to_datetime(df_colab[col_sit].iloc[0], errors='coerce', dayfirst=True)
                            if pd.notna(d_adm) and pd.notna(d_sit):
                                anos = (d_sit - d_adm).days / 365.25
                    except:
                        pass
                    if anos > 1.0:
                        base_mensal = 211.65
                        status = f"Regra CCT Sind 40 (R$ 211,65 acima de 1 ano / 30 * {dias_trab}d)"
                    else:
                        base_mensal = 141.10
                        status = f"Regra CCT Sind 40 (R$ 141,10 até 1 ano / 30 * {dias_trab}d)"
                elif sind_id in [44, 46, 79, 83, 96]:
                    base_mensal = sal_base * 0.10
                    status = f"Regra CCT Sind {sind_id} (10% Sal.Base / 30 * {dias_trab}d)"
                elif sind_id == 47:
                    base_mensal = 1800.00 * 0.07
                    status = f"Regra CCT Sind 47 (7% Piso / 30 * {dias_trab}d)"
                elif sind_id == 78:
                    base_mensal = 1621.00 * 0.19
                    status = f"Regra CCT Sind 78 (19% Sal.Mín / 30 * {dias_trab}d)"
                elif sind_id == 87:
                    base_mensal = sal_base * 0.105
                    status = f"Regra CCT Sind 87 (10.5% Remuneração / 30 * {dias_trab}d)"
                elif sind_id in [104, 105, 115, 140]:
                    base_mensal = sal_base * 0.10
                    status = f"Regra CCT Sind {sind_id} (10% Base+Pericul / 30 * {dias_trab}d)"
                elif sind_id == 108:
                    base_mensal = 160.00
                    status = f"Regra CCT Sind 108 (Fixo R$ 160,00 / 30 * {dias_trab}d)"
                elif sind_id == 112:
                    base_mensal = 141.75
                    status = f"Regra CCT Sind 112 (Fixo R$ 141,75 / 30 * {dias_trab}d)"
                elif sind_id == 118:
                    base_mensal = 169.23
                    status = f"Regra CCT Sind 118 (Fixo R$ 169,23 / 30 * {dias_trab}d)"
                elif sind_id == 136:
                    base_mensal = 74.00
                    status = f"Regra CCT Sind 136 (Fixo R$ 74,00 / 30 * {dias_trab}d)"
                elif sind_id == 137:
                    base_mensal = 105.61
                    status = f"Regra CCT Sind 137 (Fixo R$ 105,61 / 30 * {dias_trab}d)"
                elif sind_id == 138:
                    base_mensal = 225.00
                    status = f"Regra CCT Sind 138 (Fixo R$ 225,00 / 30 * {dias_trab}d)"
                else:
                    base_mensal = 97.26
                    status = f"Regra CCT Quebra de Caixa (/ 30 * {dias_trab}d)"
                    
                valor_calc = round((base_mensal / 30) * dias_trab, 2)
                if abs(valor_sistema - valor_calc) > 0.05 and abs(valor_sistema - round(base_mensal, 2)) <= 0.05:
                    valor_calc = round(base_mensal, 2)
                    status = f"Regra CCT Quebra de Caixa Integral (R$ {base_mensal:.2f})"

            elif cod == 5045:
                # Falta de Caixa (Desconto de Caixa)
                lotacao_str = str(df_colab['LOTAÇÃO TRIBUTÁRIA'].iloc[0]) if 'LOTAÇÃO TRIBUTÁRIA' in df_colab.columns else ''
                sind_id = extrair_sind_da_lotacao(lotacao_str)
                tem_13902 = len(df_colab[df_colab['CÓDIGO DO EVENTO'] == 13902]) > 0
                if tem_13902:
                    val_13902 = df_colab[df_colab['CÓDIGO DO EVENTO'] == 13902]['VALOR EVENTO'].sum()
                    valor_calc = val_13902
                    status = "Validação Falta de Caixa (Conferido com Apontamento PDV 13902)"
                else:
                    valor_calc = valor_sistema
                    status = "Validação Falta de Caixa (Importado PDV)"
                    
                # Salvaguardas CCT de Falta de Caixa
                tem_quebra = len(df_colab[df_colab['CÓDIGO DO EVENTO'] == 96]) > 0
                if not tem_quebra:
                    status += " | ALERTA: Desconto de Falta de Caixa sem percepção de Quebra de Caixa (96)"
                if sind_id == 127:
                    status += " | ALERTA: CCT Sind 127 proíbe desconto de falta de caixa"

            elif cod == 126:
                # Salário Família (Regra Previdenciária Portaria MPS/MF R$ 67,54 por dependente)
                num_dep = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20275]['VALOR EVENTO'].sum()
                if num_dep == 0:
                    num_dep = 1.0
                dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20235]['VALOR EVENTO'].sum()
                if dias_trab == 0:
                    dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 721]['VALOR EVENTO REFERÊNCIA'].sum()
                if dias_trab == 0:
                    dias_trab = 30.0
                if len(df_colab[df_colab['CÓDIGO DO EVENTO'] == 300]) > 0:
                    dias_ferias = df_colab[df_colab['CÓDIGO DO EVENTO'] == 300]['VALOR EVENTO REFERÊNCIA'].sum()
                    if dias_ferias == 0:
                        val_fer = df_colab[df_colab['CÓDIGO DO EVENTO'] == 300]['VALOR EVENTO'].sum()
                        if salario_contratual > 0:
                            dias_ferias = round(val_fer / (salario_contratual / 30))
                    dias_trab += dias_ferias

                cota = 67.54
                valor_calc = round(((cota * num_dep) / 30) * dias_trab, 2)
                if abs(valor_sistema - valor_calc) > 0.05:
                    calc_6750 = round(((67.50 * num_dep) / 30) * dias_trab, 2)
                    if abs(valor_sistema - calc_6750) <= 0.05:
                        valor_calc = calc_6750
                        cota = 67.50
                status = f"Regra Previdenciária Salário Família (R$ {cota:.2f} * {int(num_dep)} dep / 30 * {dias_trab}d)"
                
            elif cod == 3:
                # Produtividade Geral (CCT Geral 6% sobre Salário Contratual proporcional)
                dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20235]['VALOR EVENTO'].sum()
                if dias_trab == 0:
                    dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 721]['VALOR EVENTO REFERÊNCIA'].sum()
                if dias_trab == 0:
                    dias_trab = 30.0
                valor_calc = round(((salario_contratual * 0.06) / 30) * dias_trab, 2)
                status = f"Regra CCT (6% Sal.Contr / 30 * {dias_trab}d)"
                
            elif cod == 243:
                # Produtividade Goiás (Sind 127) - 5% parcela mensal integral indenizatória
                valor_calc = round(salario_contratual * 0.05, 2)
                status = "Regra CCT Sind 127 (5% Sal.Contr Integral Indenizatório)"
                
            elif cod == 164:
                # Responsabilidade Técnica Farmácia (Sind 35 AL: R$ 1.000,00 mensal proporcional)
                dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20235]['VALOR EVENTO'].sum()
                if dias_trab == 0:
                    dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 721]['VALOR EVENTO REFERÊNCIA'].sum()
                if dias_trab == 0:
                    dias_trab = 30.0
                valor_calc = round((1000.00 / 30) * dias_trab, 2)
                status = f"Regra CCT Sind 35 (R$ 1000,00 Resp. Técnica / 30 * {dias_trab}d)"
                
            elif cod == 222:
                # Gratificação Gestor Farmácia (Sind 25 SE e 35 AL: 40% sobre salário proporcional)
                dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20235]['VALOR EVENTO'].sum()
                if dias_trab == 0:
                    dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 721]['VALOR EVENTO REFERÊNCIA'].sum()
                if dias_trab == 0:
                    dias_trab = 30.0
                valor_calc = round(((salario_contratual * 0.40) / 30) * dias_trab, 2)
                status = f"Regra CCT (40% Grat. Gestor Farma / 30 * {dias_trab}d)"
                
            elif cod == 165:
                # Adicional Titulação Farmácia (Sind 35 AL: 15% sobre piso R$ 4.944,03 = R$ 741,60 proporcional)
                dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20235]['VALOR EVENTO'].sum()
                if dias_trab == 0:
                    dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 721]['VALOR EVENTO REFERÊNCIA'].sum()
                if dias_trab == 0:
                    dias_trab = 30.0
                valor_calc = round((741.60 / 30) * dias_trab, 2)
                status = f"Regra CCT Sind 35 (R$ 741,60 Titulação Farma / 30 * {dias_trab}d)"
                
            elif cod == 110:
                # Ajuda de Custo Teletrabalho / Home Office (R$ 100,00)
                dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20235]['VALOR EVENTO'].sum()
                if dias_trab == 0:
                    dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 721]['VALOR EVENTO REFERÊNCIA'].sum()
                if dias_trab == 0:
                    dias_trab = 30.0
                teorico_prop = round((100.00 / 30) * dias_trab, 2)
                if abs(valor_sistema - 100.00) <= 0.05:
                    valor_calc = 100.00
                    status = "Regra Contratual (R$ 100,00 Teletrabalho Integral)"
                elif abs(valor_sistema - teorico_prop) <= 0.05:
                    valor_calc = teorico_prop
                    status = f"Regra Contratual (R$ 100,00 / 30 * {dias_trab}d)"
                else:
                    valor_calc = 100.00
                    status = "Regra Contratual (R$ 100,00 Teletrabalho)"
                    
            elif cod == 103:
                # Ajuda de Custo Aluguel (Individual / Contratual por Gestor)
                valor_calc = valor_sistema
                status = "Valor Importado (Contratual / Ajuda de Custo Aluguel Gestor)"
                
            elif cod == 163:
                # Adicional de Gerência (Inativo nas folhas correntes)
                valor_calc = valor_sistema
                status = "Valor Importado (Adicional de Gerência Histórico / Inativo)"
                
            elif cod in [162, 202, 824]:
                # Prêmios e Indenizações de Assiduidade
                lotacao_str = str(df_colab['LOTAÇÃO TRIBUTÁRIA'].iloc[0]) if 'LOTAÇÃO TRIBUTÁRIA' in df_colab.columns else ''
                sind_id = extrair_sind_da_lotacao(lotacao_str)
                if sind_id == 111:
                    valor_calc = round(salario_contratual * 0.02, 2)
                    status = "Regra CCT Sind 111 (2% Assiduidade Integral)"
                elif sind_id == 138:
                    valor_calc = round(salario_contratual * 0.03, 2)
                    status = "Regra CCT Sind 138 (3% Assiduidade Integral)"
                else:
                    valor_calc = valor_sistema
                    status = "Regra CCT (Prêmio Assiduidade Conforme CCT Local)"
                    
                # Salvaguarda de Assiduidade vs Faltas
                tem_faltas = df_colab[df_colab['CÓDIGO DO EVENTO'].isin([5001, 20010])]['VALOR EVENTO'].sum() > 0
                if tem_faltas:
                    status += " | ALERTA: Colaborador possui dedução de falta (5001) no período"

            elif cod == 2:
                # Mínimo Garantido Comissionista (CCT Vendedores)
                dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20235]['VALOR EVENTO'].sum()
                if dias_trab == 0 and not pd.isna(ref) and ref > 0:
                    dias_trab = ref
                if dias_trab == 0:
                    dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 721]['VALOR EVENTO REFERÊNCIA'].sum()
                if dias_trab == 0:
                    dias_trab = 30.0
                sal_base = salario_contratual if salario_contratual > 0 else remun_fixa
                piso_prop = round((sal_base / 30) * dias_trab, 2)
                
                val_241 = df_colab[df_colab['CÓDIGO DO EVENTO'] == 241]['VALOR EVENTO'].sum()
                calc_primario = max(0.0, round(piso_prop - val_241, 2))
                
                soma_comissoes_geral = df_colab[df_colab['CÓDIGO DO EVENTO'].isin([241, 81, 88, 89])]['VALOR EVENTO'].sum()
                calc_fallback = max(0.0, round(piso_prop - soma_comissoes_geral, 2))
                
                if abs(valor_sistema - calc_primario) <= 0.05:
                    valor_calc = calc_primario
                    status = f"Regra CCT Mínimo Garantido (Piso Prop {piso_prop:.2f} - Comissões 241 {val_241:.2f})"
                elif abs(valor_sistema - calc_fallback) <= 0.05:
                    valor_calc = calc_fallback
                    status = f"Regra CCT Mínimo Garantido (Piso Prop {piso_prop:.2f} - Comissões {soma_comissoes_geral:.2f})"
                else:
                    valor_calc = calc_primario
                    status = f"Regra CCT Mínimo Garantido (Piso Prop {piso_prop:.2f} - 241 {val_241:.2f})"

            elif cod in [90, 82]:
                # DSR sobre Comissões
                comissoes_ids = [241, 240, 89, 88, 83, 81, 31960]
                soma_comissoes = df_colab[df_colab['CÓDIGO DO EVENTO'].isin(comissoes_ids)]['VALOR EVENTO'].sum()
                qtd_dsr = df_colab[df_colab['CÓDIGO DO EVENTO'] == 21034]['VALOR EVENTO'].sum()
                if qtd_dsr == 0 and not pd.isna(ref) and ref > 0:
                    qtd_dsr = ref
                if qtd_dsr == 0:
                    qtd_dsr = 5.0
                
                dias_uteis = 30 - qtd_dsr if qtd_dsr > 0 else 25
                calc_dsr = round((soma_comissoes / dias_uteis) * qtd_dsr, 2)
                calc_24 = round((soma_comissoes / 24) * qtd_dsr, 2)
                
                if abs(valor_sistema - calc_dsr) <= 0.05:
                    valor_calc = calc_dsr
                    status = f"Regra CLT DSR Comissões (Comissões {soma_comissoes:.2f} / {dias_uteis}d * {qtd_dsr}d DSR)"
                elif abs(valor_sistema - calc_24) <= 0.05:
                    valor_calc = calc_24
                    status = f"Regra CLT DSR Comissões (Comissões {soma_comissoes:.2f} / 24d * {qtd_dsr}d DSR)"
                else:
                    valor_calc = calc_dsr
                    status = f"Regra CLT DSR Comissões (Comissões {soma_comissoes:.2f} * {qtd_dsr}d DSR)"

            elif cod in [241, 240, 89, 88, 83, 81]:
                # Comissões de Vendas (Apuração PDV / Faturamento)
                valor_calc = valor_sistema
                mapa_desc_comissao = {
                    241: "Comissões Pura Vendedor",
                    240: "Comissões Garantia Estendida",
                    89: "Comissão Cestas Perini",
                    88: "Comissão Vendas Encarregado",
                    83: "Comissão Seguro Furto e Roubo",
                    81: "Comissão Geral Vendas"
                }
                desc_com = mapa_desc_comissao.get(cod, "Comissões Vendas")
                status = f"Validação Faturamento Vendas ({desc_com} - Importado PDV)"

            elif cod == 117:
                # Crédito Cesta Básica (Bonificação CCT por Feriados Trabalhados)
                lotacao_str = str(df_colab['LOTAÇÃO TRIBUTÁRIA'].iloc[0]) if 'LOTAÇÃO TRIBUTÁRIA' in df_colab.columns else ''
                sind_id = extrair_sind_da_lotacao(lotacao_str)
                valor_feriado_cct = 80.00 if sind_id in [1, 66, 69, 97] else 124.00
                
                qtd_feriados = ref if not pd.isna(ref) and ref > 0 else (round(valor_sistema / valor_feriado_cct) if valor_feriado_cct > 0 else 1)
                if qtd_feriados == 0:
                    qtd_feriados = 1
                calc_feriados = round(qtd_feriados * valor_feriado_cct, 2)
                if abs(valor_sistema - calc_feriados) <= 0.05:
                    valor_calc = calc_feriados
                    status = f"Regra CCT Crédito Cesta Básica ({int(qtd_feriados)} feriado(s) * R$ {valor_feriado_cct:.2f})"
                else:
                    valor_calc = valor_sistema
                    status = f"Regra CCT Crédito Cesta Básica (R$ {valor_sistema:.2f} Conforme CCT Sind {sind_id})"

            elif cod == 114:
                # Auxílio Creche (Regra CCT / Reembolso Comprovado)
                teto_creche = 759.00
                piso_15pct = 259.83
                if abs(valor_sistema - piso_15pct) <= 0.05:
                    valor_calc = piso_15pct
                    status = "Regra CCT Auxílio Creche (15% Piso Salarial R$ 259,83)"
                elif valor_sistema <= teto_creche + 0.05:
                    valor_calc = valor_sistema
                    status = f"Validação Reembolso Creche (Comprovante R$ {valor_sistema:.2f} <= Teto R$ {teto_creche:.2f})"
                else:
                    valor_calc = valor_sistema
                    status = f"Validação Reembolso Creche (R$ {valor_sistema:.2f}) | ALERTA: Valor superior ao teto CCT de R$ {teto_creche:.2f}"

            elif cod == 128:
                # Abono CCT
                valor_calc = valor_sistema
                status = "Regra CCT (Abono Salarial Acordo Coletivo)"

            elif cod == 221:
                # Gratificação Adega / Sommelier
                dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20235]['VALOR EVENTO'].sum()
                if dias_trab == 0:
                    dias_trab = df_colab[df_colab['CÓDIGO DO EVENTO'] == 721]['VALOR EVENTO REFERÊNCIA'].sum()
                if dias_trab == 0:
                    dias_trab = 30.0
                prop_500 = round((500.00 / 30) * dias_trab, 2)
                if abs(valor_sistema - 500.00) <= 0.05:
                    valor_calc = 500.00
                    status = "Regra Contratual (R$ 500,00 Gratificação Adega Integral)"
                elif abs(valor_sistema - prop_500) <= 0.05:
                    valor_calc = prop_500
                    status = f"Regra Contratual (R$ 500,00 Adega / 30 * {dias_trab}d)"
                else:
                    valor_calc = valor_sistema
                    status = f"Regra Contratual (Gratificação Adega R$ {valor_sistema:.2f})"

            elif cod in [5029, 5030, 5031, 5703, 5704, 5705, 5058, 5059, 5060, 5503, 5504, 5520, 5521, 5522, 5723, 5724, 5725, 5726, 6232, 6233, 6234, 6529, 7226]:
                # Pensão Alimentícia Judicial (Batch 5 - Ordens Judiciais)
                # 1. Identificar Percentual Judicial fixado em sentença
                # Rubricas auxiliares de percentual no FPW:
                # 23006: % Pensão Líquida Pensionista 1
                # 23007 / 23008: % Pensão Pensionista 2 / 3
                pct_pensao = 0.0
                if cod in [5029, 5703, 5058, 5503, 5723, 6232]:
                    pct_pensao = df_colab[df_colab['CÓDIGO DO EVENTO'] == 23006]['VALOR EVENTO'].sum()
                elif cod in [5030, 5704, 5059, 5504, 5724, 6233]:
                    pct_pensao = df_colab[df_colab['CÓDIGO DO EVENTO'] == 23007]['VALOR EVENTO'].sum()
                    if pct_pensao == 0:
                        pct_pensao = df_colab[df_colab['CÓDIGO DO EVENTO'] == 23006]['VALOR EVENTO'].sum()
                elif cod in [5031, 5705, 5060, 5725, 6234]:
                    pct_pensao = df_colab[df_colab['CÓDIGO DO EVENTO'] == 23008]['VALOR EVENTO'].sum()
                    if pct_pensao == 0:
                        pct_pensao = df_colab[df_colab['CÓDIGO DO EVENTO'] == 23006]['VALOR EVENTO'].sum()
                
                # Se não veio na rubrica 2300x, checar na Referência do próprio evento
                if pct_pensao == 0 and not pd.isna(ref) and ref > 0:
                    pct_pensao = ref if ref <= 100 else ref / 100.0

                # 2. Definição da Base Conforme o Tipo de Rubrica:
                if cod in [5029, 5030, 5031, 6529]:
                    # Pensão Mês (Saldo Salário + Adicionais Mensais - Faltas/DSR - INSS Mês - IRRF Mês)
                    base_inss_mes = df_colab[df_colab['CÓDIGO DO EVENTO'] == 10100]['VALOR EVENTO'].sum()
                    inss_mes = df_colab[df_colab['CÓDIGO DO EVENTO'] == 5013]['VALOR EVENTO'].sum()
                    irrf_mes = df_colab[df_colab['CÓDIGO DO EVENTO'] == 5019]['VALOR EVENTO'].sum()
                    
                    if base_inss_mes > 0:
                        base_liq_mes = max(0.0, base_inss_mes - inss_mes - irrf_mes)
                    else:
                        prov_mes_ids = [721, 13, 11, 12, 14, 15, 3, 4, 5, 69, 78, 137, 164, 165, 222, 243]
                        tot_prov_mes = df_colab[df_colab['CÓDIGO DO EVENTO'].isin(prov_mes_ids)]['VALOR EVENTO'].sum()
                        desc_mes_ids = [5001, 5007]
                        tot_desc_mes = df_colab[df_colab['CÓDIGO DO EVENTO'].isin(desc_mes_ids)]['VALOR EVENTO'].sum()
                        base_liq_mes = max(0.0, (tot_prov_mes - tot_desc_mes) - inss_mes - irrf_mes)
                    
                    if pct_pensao > 0:
                        valor_calc = round(base_liq_mes * (pct_pensao / 100.0), 2)
                        status = f"Regra Judicial (Base Líq Mês {base_liq_mes:.2f} * {pct_pensao}%)"
                    else:
                        valor_calc = valor_sistema
                        status = "Validação Pensão Mês (Ordem Judicial / Fixo Importado)"
                        
                elif cod in [5703, 5704, 5705, 5726]:
                    # Pensão 13º Salário Rescisório (13º Bruto - INSS 13º - IRRF 13º)
                    base_13_ids = [756, 757, 759, 760]
                    tot_13 = df_colab[df_colab['CÓDIGO DO EVENTO'].isin(base_13_ids)]['VALOR EVENTO'].sum()
                    inss_13 = df_colab[df_colab['CÓDIGO DO EVENTO'] == 5701]['VALOR EVENTO'].sum()
                    irrf_13 = df_colab[df_colab['CÓDIGO DO EVENTO'] == 5702]['VALOR EVENTO'].sum()
                    base_liq_13 = max(0.0, tot_13 - inss_13 - irrf_13)
                    
                    if pct_pensao > 0:
                        valor_calc = round(base_liq_13 * (pct_pensao / 100.0), 2)
                        status = f"Regra Judicial (Base Líq 13º {base_liq_13:.2f} * {pct_pensao}%)"
                    else:
                        valor_calc = valor_sistema
                        status = "Validação Pensão 13º (Ordem Judicial / Fixo Importado)"
                        
                elif cod in [5058, 5059, 5060, 5503, 5504, 5520, 5521, 5522, 7226]:
                    # Pensão sobre Férias Rescisórias / Gozadas
                    ferias_ids = [737, 738, 739, 742, 743, 744, 747, 748, 749, 752, 753, 754, 300, 301, 302, 303]
                    tot_ferias = df_colab[df_colab['CÓDIGO DO EVENTO'].isin(ferias_ids)]['VALOR EVENTO'].sum()
                    inss_ferias = df_colab[df_colab['CÓDIGO DO EVENTO'] == 5015]['VALOR EVENTO'].sum()
                    base_liq_ferias = max(0.0, tot_ferias - inss_ferias)
                    
                    if pct_pensao > 0:
                        valor_calc = round(base_liq_ferias * (pct_pensao / 100.0), 2)
                        status = f"Regra Judicial (Base Férias {base_liq_ferias:.2f} * {pct_pensao}%)"
                    else:
                        valor_calc = valor_sistema
                        status = "Validação Pensão Férias (Ordem Judicial / Fixo Importado)"
                        
                elif cod in [5723, 5724, 5725]:
                    # Pensão sobre Aviso Prévio Indenizado / Indenizações Rescisórias
                    aviso_ids = [725, 726, 729, 735]
                    tot_indeniz = df_colab[df_colab['CÓDIGO DO EVENTO'].isin(aviso_ids)]['VALOR EVENTO'].sum()
                    if pct_pensao > 0:
                        valor_calc = round(tot_indeniz * (pct_pensao / 100.0), 2)
                        status = f"Regra Judicial (Base Indenizatória {tot_indeniz:.2f} * {pct_pensao}%)"
                    else:
                        valor_calc = valor_sistema
                        status = "Validação Pensão Indenizatória (Ordem Judicial / Fixo Importado)"
                        
                else: # PLR e outros
                    valor_calc = valor_sistema
                    status = "Validação Pensão (Ordem Judicial / PLR)"

                # Salvaguarda de Totalização contra o Evento 13385 (PA p/ Relatórios)
                tot_13385 = df_colab[df_colab['CÓDIGO DO EVENTO'] == 13385]['VALOR EVENTO'].sum()
                if tot_13385 > 0:
                    status += f" | Totalizador 13385 Conferido (R$ {tot_13385:.2f})"
                
            elif cod in [5082, 5083, 5084, 5085, 5086, 5087, 5088, 5089, 5090]:
                # Empréstimo Consignado em Folha (Validação de Margem Legal - Limite de 35%)
                valor_calc = valor_sistema
                consignados_ids = [5082, 5083, 5084, 5085, 5086, 5087, 5088, 5089, 5090]
                soma_consignados = df_colab[df_colab['CÓDIGO DO EVENTO'].isin(consignados_ids)]['VALOR EVENTO'].sum()
                base_margem = df_colab[df_colab['CÓDIGO DO EVENTO'] == 17910]['VALOR EVENTO'].sum()
                if base_margem == 0:
                    base_margem = salario_contratual if salario_contratual > 0 else remun_fixa
                
                pct_utilizado = (soma_consignados / base_margem) * 100 if base_margem > 0 else 0
                if pct_utilizado > 35.01:
                    status = f"Validação Consignado Bancário ({pct_utilizado:.1f}% base) | ALERTA: Margem excede limite legal de 35%"
                else:
                    status = f"Validação Consignado Bancário (Margem Legal OK: {pct_utilizado:.1f}% <= 35%)"
                
            elif cod in [29, 31, 37, 67, 153, 156, 63, 68, 150, 227]:
                valor_calc = valor_sistema
                status = "Valor Importado (Apontamento de Horas/DSR Externo)"
                
            elif cod in [5112, 5722, 821, 5053, 5751]:
                valor_calc = valor_sistema
                status = "Valor Importado (Exceções / Histórico Antigo)"
                
            elif cod == 499:
                # Insuficiência de Saldo (CLT Art. 477 - Vedação de TRCT Negativo)
                # No cálculo do FPW, os proventos NÃO incluem o Salário Família (126) pois é verba previdenciária infungível
                soma_prov_sem_499_126 = df_colab[(df_colab['TIPO DE EVENTO'] == 'PROVENTO') & (~df_colab['CÓDIGO DO EVENTO'].isin([499, 126]))]['VALOR EVENTO'].sum()
                soma_desc = df_colab[df_colab['TIPO DE EVENTO'] == 'DESCONTO']['VALOR EVENTO'].sum()
                valor_calc = round(max(0.0, soma_desc - soma_prov_sem_499_126), 2)
                status = f"Regra CLT Art. 477 (Insuficiência de Saldo: Descontos {soma_desc:.2f} - Proventos {soma_prov_sem_499_126:.2f})"
                
            elif cod == 137:
                valor_calc = valor_sistema
                status = "Valor Importado (Afastamento Previdenciário)"
                
            elif cod == 5707:
                valor_calc = valor_sistema
                status = "Valor Importado (Adiantamento 13º Histórico)"
                
            elif cod in [10970, 10980, 10990, 20010, 11535]:
                valor_calc = valor_sistema
                status = "Valor Importado (Resumo do Sistema)"
                
            elif cod == 69:
                qtd_noturno = df_colab[df_colab['CÓDIGO DO EVENTO'] == 21040]['VALOR EVENTO'].sum()
                valor_hora = remun_fixa / 220
                valor_calc = round(valor_hora * qtd_noturno * 0.2, 2)
                status = f"Regra LG (Hora {valor_hora:.2f} * {qtd_noturno}h * 20%)"
                
            elif cod == 78:
                qtd_dsr = df_colab[df_colab['CÓDIGO DO EVENTO'] == 21034]['VALOR EVENTO'].sum()
                ad_noturno_events = [69, 70, 71, 72, 73, 74, 75, 76, 77]
                soma_noturno = df_colab[df_colab['CÓDIGO DO EVENTO'].isin(ad_noturno_events)]['VALOR EVENTO'].sum()
                
                # Mesma lógica do DSR de Horas Extras
                dias_uteis_inferido = 30 - qtd_dsr if qtd_dsr > 0 else 25 
                
                if dias_uteis_inferido > 0:
                    valor_calc = round((soma_noturno / dias_uteis_inferido) * qtd_dsr, 2)
                    status = f"Regra LG (Soma Ad.Noturno / {dias_uteis_inferido} * {qtd_dsr})"

            elif cod == 725:
                dias_aviso_lei = df_colab[df_colab['CÓDIGO DO EVENTO'] == 20208]['VALOR EVENTO'].sum()
                dias_totais = 30 + dias_aviso_lei
                valor_calc = round((remun_fixa / 30) * dias_totais, 2)
                status = "Regra CLT (30d + 3d/ano)"
                
            elif cod == 726:
                base_media = df_colab[df_colab['CÓDIGO DO EVENTO'] == 10301]['VALOR EVENTO'].sum()
                if base_media > 0:
                    valor_1_dia = base_media / 30
                    dias_inferidos = round(valor_sistema / valor_1_dia, 2) if valor_1_dia > 0 else 0
                    valor_calc = round(valor_1_dia * dias_inferidos, 2)
                    status = f"Reversa ({dias_inferidos} dias inferidos de MD Av. Prévio)"
                    
            elif cod == 729:
                base_salario = salario_contratual if salario_contratual > 0 else remun_fixa
                valor_calc = base_salario
                status = "Regra CLT (Multa Lei 7238/84 = 1 Salário Contratual)"
                
            elif cod == 735:
                base_salario = salario_contratual if salario_contratual > 0 else remun_fixa
                valor_calc = base_salario
                status = "Regra CLT (Art 477 = 1 Salário Contratual)"
                
            elif cod == 737:
                if remun_fixa > 0:
                    valor_ferias_integrais = (remun_fixa / 30) * dias_direito_f
                    valor_1_dia = valor_ferias_integrais / dias_direito_f
                    dias_inferidos = round(valor_sistema / valor_1_dia, 2) if valor_1_dia > 0 else 0
                    valor_calc = round(valor_1_dia * dias_inferidos, 2)
                    status = f"Reversa ({dias_inferidos} dias inferidos)"
                
            elif cod == 738:
                if remun_fixa > 0:
                    valor_ferias_integrais = (remun_fixa / 30) * dias_direito_f
                    valor_1_avo = valor_ferias_integrais / 12
                    avos_inferidos = round(valor_sistema / valor_1_avo, 2) if valor_1_avo > 0 else 0
                    valor_calc = round(valor_1_avo * avos_inferidos, 2)
                    status = f"Reversa ({avos_inferidos} avos sobre {dias_direito_f} dias)"
                    
            elif cod == 739:
                if remun_fixa > 0:
                    valor_ferias_integrais = (remun_fixa / 30) * dias_direito_f
                    valor_1_avo = valor_ferias_integrais / 12
                    avos_inferidos = round(valor_sistema / valor_1_avo, 2) if valor_1_avo > 0 else 0
                    valor_calc = round(valor_1_avo * avos_inferidos, 2)
                    status = f"Reversa ({avos_inferidos} avos inferidos de Férias Ind AP)"
                
            elif cod == 301:
                base = df_colab[df_colab['CÓDIGO DO EVENTO'] == 300]['VALOR EVENTO'].sum()
                valor_calc = round(base / 3, 2)
                status = "Regra CLT (1/3 Base 300)"
                
            elif cod == 303:
                base = df_colab[df_colab['CÓDIGO DO EVENTO'] == 302]['VALOR EVENTO'].sum()
                valor_calc = round(base / 3, 2)
                status = "Regra CLT (1/3 Base 302)"
                
            elif cod == 326:
                base = df_colab[df_colab['CÓDIGO DO EVENTO'] == 325]['VALOR EVENTO'].sum()
                valor_calc = round(base / 3, 2)
                status = "Regra CLT (1/3 Base 325)"
                
            elif cod == 747:
                base = df_colab[df_colab['CÓDIGO DO EVENTO'] == 737]['VALOR EVENTO'].sum()
                valor_calc = round(base / 3, 2)
                status = "Regra CLT (1/3 Base Férias Vencidas)"
                
            elif cod == 748:
                base = df_colab[df_colab['CÓDIGO DO EVENTO'] == 738]['VALOR EVENTO'].sum()
                valor_calc = round(base / 3, 2)
                status = "Regra CLT (1/3 Base Férias Prop)"
                
            elif cod == 749:
                base = df_colab[df_colab['CÓDIGO DO EVENTO'] == 739]['VALOR EVENTO'].sum()
                valor_calc = round(base / 3, 2)
                status = "Regra CLT (1/3 Base Férias Ind)"
                
            elif cod == 752:
                base = df_colab[df_colab['CÓDIGO DO EVENTO'] == 742]['VALOR EVENTO'].sum()
                valor_calc = round(base / 3, 2)
                status = "Regra CLT (1/3 MD Férias Vencidas)"
                
            elif cod == 753:
                base = df_colab[df_colab['CÓDIGO DO EVENTO'] == 743]['VALOR EVENTO'].sum()
                valor_calc = round(base / 3, 2)
                status = "Regra CLT (1/3 MD Férias Prop)"
                
            elif cod == 754:
                base = df_colab[df_colab['CÓDIGO DO EVENTO'] == 744]['VALOR EVENTO'].sum()
                valor_calc = round(base / 3, 2)
                status = "Regra CLT (1/3 MD Férias Ind AP)"
                
            elif cod == 756:
                if remun_fixa > 0:
                    valor_1_avo = remun_fixa / 12
                    avos_inferidos = round(valor_sistema / valor_1_avo, 2) if valor_1_avo > 0 else 0
                    valor_calc = round(valor_1_avo * avos_inferidos, 2)
                    status = f"Reversa ({avos_inferidos} avos inferidos de 13º Prop)"
                    
            elif cod == 757:
                if remun_fixa > 0:
                    valor_1_avo = remun_fixa / 12
                    avos_inferidos = round(valor_sistema / valor_1_avo, 2) if valor_1_avo > 0 else 0
                    valor_calc = round(valor_1_avo * avos_inferidos, 2)
                    status = f"Reversa ({avos_inferidos} avos inferidos de 13º Ind AP)"
                    
            elif cod == 759:
                base_media = df_colab[df_colab['CÓDIGO DO EVENTO'] == 10316]['VALOR EVENTO'].sum()
                if base_media > 0:
                    valor_1_avo = base_media / 12
                    avos_inferidos = round(valor_sistema / valor_1_avo, 2) if valor_1_avo > 0 else 0
                    valor_calc = round(valor_1_avo * avos_inferidos, 2)
                    status = f"Reversa ({avos_inferidos} avos inferidos de MD 13º Prop)"
                    
            elif cod == 760:
                base_media = df_colab[df_colab['CÓDIGO DO EVENTO'] == 10316]['VALOR EVENTO'].sum()
                if base_media > 0:
                    valor_1_avo = base_media / 12
                    avos_inferidos = round(valor_sistema / valor_1_avo, 2) if valor_1_avo > 0 else 0
                    valor_calc = round(valor_1_avo * avos_inferidos, 2)
                    status = f"Reversa ({avos_inferidos} avos inferidos de MD 13º Ind AP)"
                
            elif cod == 5013:
                base_inss = df_colab[df_colab['CÓDIGO DO EVENTO'] == 10100]['VALOR EVENTO'].sum()
                inss_bruto = calcular_inss_progressivo(base_inss)
                
                # Abater o INSS já descontado nas Férias (Rubrica 5015)
                inss_ferias = df_colab[df_colab['CÓDIGO DO EVENTO'] == 5015]['VALOR EVENTO'].sum()
                
                valor_calc = round(inss_bruto - inss_ferias, 2)
                status = "Regra INSS Prog - INSS Férias (5015)"
                
            elif cod == 5701:
                base_inss_13 = df_colab[df_colab['CÓDIGO DO EVENTO'] == 10130]['VALOR EVENTO'].sum()
                valor_calc = calcular_inss_progressivo(base_inss_13)
                status = "Regra INSS Prog (Base 10130)"
                
            elif cod == 6290:
                saldo_salario = df_colab[df_colab['CÓDIGO DO EVENTO'] == 721]['VALOR EVENTO'].sum()
                aux_doenca = df_colab[df_colab['CÓDIGO DO EVENTO'] == 137]['VALOR EVENTO'].sum()
                
                # A base do VT na rescisão é estritamente o Saldo de Salário + Auxílios
                base_vt = saldo_salario + aux_doenca
                if base_vt == 0:
                    base_vt = remun_fixa
                    
                valor_calc = round(base_vt * 0.06, 2)
                status = f"Desconto (6% sobre Base Prop: R$ {base_vt})"
                
            elif cod == 5046:
                tem_ferias = df_colab[df_colab['CÓDIGO DO EVENTO'].isin([300, 310])]['VALOR EVENTO'].sum() > 0
                data_rescisao_str = df_colab['DATA INÍCIO SITUAÇÃO'].iloc[0] if not df_colab['DATA INÍCIO SITUAÇÃO'].isna().all() else ''
                
                empresa_nome = str(df_colab['EMPRESA'].iloc[0]).upper() if 'EMPRESA' in df_colab.columns else ''
                cargo_nome = str(df_colab['CARGO'].iloc[0]).upper() if 'CARGO' in df_colab.columns else ''
                
                # Regra de 50% para Perini e Estagiários
                if "PERINI" in empresa_nome or "ESTAG" in cargo_nome:
                    pct_adiantamento = 0.50
                    lbl_adiantamento = "50%"
                else:
                    pct_adiantamento = 0.40
                    lbl_adiantamento = "40%"
                
                demitido_antes_do_dia_15 = False
                if data_rescisao_str and isinstance(data_rescisao_str, str) and '/' in data_rescisao_str:
                    dia_resc = int(data_rescisao_str.split('/')[0])
                    if dia_resc < 15:
                        demitido_antes_do_dia_15 = True
                
                if tem_ferias or demitido_antes_do_dia_15:
                    valor_calc = valor_sistema
                    status = "Adiantamento Parcial (Houve Férias ou Demissão < dia 15)"
                else:
                    valor_calc = round(salario_contratual * pct_adiantamento, 2)
                    status = f"Desconto ({lbl_adiantamento} Adiantamento Quinzenal)"
                
            elif cod == 6266:
                valor_calc = 4.89
                status = "Desconto (Taxa Fixa R$ 4,89)"

            if not pd.isna(valor_calc):
                df.at[idx, 'VALOR CALCULADO PYTHON'] = round(valor_calc, 2)
                diverg = abs(valor_sistema - valor_calc)
                df.at[idx, 'DIVERGÊNCIA'] = round(diverg, 2)
                
                if diverg <= 1.00:
                    df.at[idx, 'STATUS AUDITORIA'] = "OK: " + status
                else:
                    df.at[idx, 'STATUS AUDITORIA'] = "DIVERGENTE: " + status

        # Se o colaborador não tem o evento 10990 (líquido zerado omitido pelo LG), injetar linha 10990 = 0.00
        if not (df_colab['CÓDIGO DO EVENTO'] == 10990).any() and len(df_colab) > 0:
            tot_prov = df_colab[df_colab['CÓDIGO DO EVENTO'] == 10970]['VALOR EVENTO'].sum()
            tot_desc = df_colab[df_colab['CÓDIGO DO EVENTO'] == 10980]['VALOR EVENTO'].sum()
            liq_calc = max(0.0, round(tot_prov - tot_desc, 2))
            
            nova_linha = df_colab.iloc[0].to_dict()
            nova_linha['CÓDIGO DO EVENTO'] = 10990
            if 'DESCRIÇÃO DO EVENTO' in nova_linha:
                nova_linha['DESCRIÇÃO DO EVENTO'] = 'Liquido a Receber'
            if 'TIPO DE EVENTO' in nova_linha:
                nova_linha['TIPO DE EVENTO'] = 'RESULTADO'
            nova_linha['VALOR EVENTO'] = 0.0
            nova_linha['VALOR EVENTO REFERÊNCIA'] = 0.0
            nova_linha['VALOR CALCULADO PYTHON'] = liq_calc
            nova_linha['DIVERGÊNCIA'] = round(abs(0.0 - liq_calc), 2)
            if abs(0.0 - liq_calc) <= 1.00:
                nova_linha['STATUS AUDITORIA'] = "OK: Líquido Zerado (Insuficiência de Saldo Mes)"
            else:
                nova_linha['STATUS AUDITORIA'] = "DIVERGENTE: Líquido omitido pelo LG mas Proventos > Descontos"
            linhas_extras_10990.append(nova_linha)

    if linhas_extras_10990:
        log(f"Incluindo {len(linhas_extras_10990)} colaborador(es) com Líquido Zerado (evento 10990 omitido pelo LG)...")
        df_extras = pd.DataFrame(linhas_extras_10990)
        df = pd.concat([df, df_extras], ignore_index=True)

    df.sort_values(by=['COLABORADOR', 'CÓDIGO DO EVENTO'], ascending=[True, True], inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Filtrar aba de resumo
    eventos_resumo = [10970, 10980, 10990, 20010, 11535]
    df_resumo_temp = df[df['CÓDIGO DO EVENTO'].isin(eventos_resumo)].copy()

    # Renomear e filtrar colunas
    df_resumo_temp.rename(columns={
        'VALOR EVENTO': 'Soma de VALOR EVENTO',
        'VALOR CALCULADO PYTHON': 'Soma de VALOR CALCULADO PYTHON',
        'DIVERGÊNCIA': 'Soma de DIVERGÊNCIA'
    }, inplace=True)

    cols_resumo = [
        'LOTAÇÃO TRIBUTÁRIA', 'MATRICULA COLABORADOR', 'EMPRESA', 'COLABORADOR', 
        'FOLHA', 'SITUAÇÃO', 'DATA DE PAGAMENTO', 'TIPO DE EVENTO', 'CÓDIGO DO EVENTO', 
        'DESCRIÇÃO DO EVENTO', 'Soma de VALOR EVENTO', 'Soma de VALOR CALCULADO PYTHON', 
        'Soma de DIVERGÊNCIA'
    ]
    cols_resumo = [c for c in cols_resumo if c in df_resumo_temp.columns]
    df_resumo_temp = df_resumo_temp[cols_resumo]
    
    # Criar Subtotais e Agrupamentos (Removidos conforme pedido, apenas as linhas puras)
    df_resumo = df_resumo_temp.reset_index(drop=True)
    
    log("Estruturando dados para o Relatório Aprimorado (4 Abas Institucionais Cencosud)...")
    if progress_fn: progress_fn(0.88)

    # 1. Montar Tabela de Status de Rescisões por Colaborador
    col_mat = 'MATRICULA COLABORADOR'
    matriculas_unicas = df[col_mat].unique()
    linhas_status = []

    for mat in matriculas_unicas:
        sub = df[df[col_mat] == mat]
        colab = sub['COLABORADOR'].iloc[0] if 'COLABORADOR' in sub.columns else ''
        lotacao = sub['LOTAÇÃO TRIBUTÁRIA'].iloc[0] if 'LOTAÇÃO TRIBUTÁRIA' in sub.columns else ''
        cargo = sub['CARGO'].iloc[0] if 'CARGO' in sub.columns else ''
        situacao = sub['SITUAÇÃO'].iloc[0] if 'SITUAÇÃO' in sub.columns else ''
        dt_adm = sub['DATA DE ADMISSÃO'].iloc[0] if 'DATA DE ADMISSÃO' in sub.columns else ''
        dt_deslig = sub['DATA INÍCIO SITUAÇÃO'].iloc[0] if 'DATA INÍCIO SITUAÇÃO' in sub.columns else ''
        dt_pgto = sub['DATA DE PAGAMENTO'].iloc[0] if 'DATA DE PAGAMENTO' in sub.columns else ''

        tot_prov = sub[sub['CÓDIGO DO EVENTO'] == 10970]['VALOR EVENTO'].sum()
        if tot_prov == 0:
            tot_prov = sub[sub['TIPO DE EVENTO'] == 'PROVENTO']['VALOR EVENTO'].sum()

        tot_desc = sub[sub['CÓDIGO DO EVENTO'] == 10980]['VALOR EVENTO'].sum()
        if tot_desc == 0:
            tot_desc = sub[sub['TIPO DE EVENTO'] == 'DESCONTO']['VALOR EVENTO'].sum()

        liq_trct = sub[sub['CÓDIGO DO EVENTO'] == 10990]['VALOR EVENTO'].sum()
        if liq_trct == 0:
            liq_trct = max(0.0, round(tot_prov - tot_desc, 2))

        linhas_status.append({
            'Matrícula': mat,
            'Colaborador': colab,
            'Lotação Tributária': lotacao,
            'Cargo': cargo,
            'Motivo do Desligamento': situacao,
            'Data Admissão': dt_adm,
            'Data Desligamento': dt_deslig,
            'Data Pagamento': dt_pgto,
            'Total Proventos (R$)': round(tot_prov, 2),
            'Total Descontos (R$)': round(tot_desc, 2),
            'Líquido TRCT (R$)': round(liq_trct, 2)
        })

    df_status = pd.DataFrame(linhas_status)

    # 2. Montar Tabela Focal de Divergências & Casos Atípicos
    linhas_tratamento = []
    for idx, row in df.iterrows():
        linha_excel = idx + 2
        mat = row[col_mat]
        colab = row['COLABORADOR']
        lotacao = row['LOTAÇÃO TRIBUTÁRIA']
        cargo = row['CARGO']
        situacao = row['SITUAÇÃO']
        dt_deslig = row['DATA INÍCIO SITUAÇÃO']
        cod = row['CÓDIGO DO EVENTO']
        desc = row['DESCRIÇÃO DO EVENTO']
        v_sis = row['VALOR EVENTO']
        v_py = row['VALOR CALCULADO PYTHON']
        v_div = row['DIVERGÊNCIA']
        status_aud = str(row['STATUS AUDITORIA'])
        tipo_ev = str(row['TIPO DE EVENTO'])

        is_div = ("DIVERGENTE" in status_aud) and (abs(v_div) > 0.01 if pd.notna(v_div) else False)
        is_falta_excedente = ("ALERTA CRÍTICO: Faltas+DSR" in status_aud)
        is_insuf_saldo = (cod == 499)
        is_pensao = (cod in [5029, 5030, 5031, 5703, 5704, 5705, 5058, 5503, 5723, 6232])
        is_assiduidade_alerta = ("ALERTA: Colaborador possui dedução de falta" in status_aud)
        is_falta_caixa_alerta = ("ALERTA: Desconto de Falta de Caixa sem correspondente" in status_aud)
        is_consignado_alerta = ("ALERTA: Soma Consignados excede margem" in status_aud)
        is_sem_regra_financ = (status_aud == "Sem regra implementada") and (tipo_ev in ['PROVENTO', 'DESCONTO']) and (v_sis > 0)

        if is_div or is_falta_excedente or is_insuf_saldo or is_pensao or is_assiduidade_alerta or is_falta_caixa_alerta or is_consignado_alerta or is_sem_regra_financ:
            if is_div:
                categoria = "Divergência de Cálculo"
                prioridade = "Alta"
                motivo = f"Divergência matemática de R$ {abs(v_div):.2f}. {status_aud}"
                providencia = "Confrontar memória de cálculo do FPW/LG com a regra técnica recalculada."
            elif is_falta_excedente:
                categoria = "Alerta de Faltas vs Data"
                prioridade = "Alta"
                motivo = status_aud.split(" | ")[-1] if " | " in status_aud else status_aud
                providencia = "Verificar espelho de ponto e se as faltas se referem a meses anteriores não processados."
            elif is_insuf_saldo:
                categoria = "Caso Atípico - Saldo Negativo"
                prioridade = "Alta"
                motivo = f"Insuficiência de Saldo (R$ {v_sis:.2f}). Rescisão negativa zerada pelo sistema."
                providencia = "Conferir limite de retenção de 1 salário (art. 477 § 5º CLT) e impenhorabilidade do Salário Família."
            elif is_pensao:
                categoria = "Alerta Jurídico - Pensão"
                prioridade = "Média"
                motivo = f"Dedução de Pensão Alimentícia Judicial (R$ {v_sis:.2f}). {status_aud}"
                providencia = "Confrontar percentual aplicado com o ofício judicial original arquivado no prontuário."
            elif is_assiduidade_alerta:
                categoria = "Alerta - Assiduidade com Falta"
                prioridade = "Alta"
                motivo = "Prêmio Assiduidade pago para colaborador com dedução de falta injustificada no período."
                providencia = "Checar se a CCT autoriza tolerância ou se houve pagamento indevido do prêmio."
            elif is_consignado_alerta:
                categoria = "Alerta Legal - Margem Consignado"
                prioridade = "Alta"
                motivo = "Empréstimos consignados ultrapassando a margem legal de 35% do líquido rescisório."
                providencia = "Limitar a retenção rescisória a 35% e notificar o banco consignatário."
            elif is_falta_caixa_alerta:
                categoria = "Alerta Operacional - Falta de Caixa"
                prioridade = "Média"
                motivo = "Desconto de Falta de Caixa sem rubrica de apuração PDV (13902)."
                providencia = "Exigir comprovante de fechamento de caixa assinado pelo operador."
            else:
                categoria = "Sem Regra - Validação Manual"
                prioridade = "Baixa"
                motivo = f"Evento financeiro {cod} sem regra automática no Python. Valor: R$ {v_sis:.2f}."
                providencia = "Validar documentação comprobatória e justificativa de lançamento manual."

            linhas_tratamento.append({
                'Linha Origem Validações': linha_excel,
                'Matrícula': mat,
                'Colaborador': colab,
                'Lotação Tributária': lotacao,
                'Cargo': cargo,
                'Motivo Desligamento': situacao,
                'Data Desligamento': dt_deslig,
                'Código Evento': cod,
                'Descrição Evento': desc,
                'Valor Sistema (R$)': v_sis,
                'Valor Python (R$)': v_py if pd.notna(v_py) else None,
                'Diferença (R$)': None,
                'Categoria da Ocorrência': categoria,
                'Prioridade Sugerida': prioridade,
                'Motivo / Ponto de Atenção': motivo,
                'Providência Sugerida': providencia,
                'Conclusão do Analista': 'Pendente',
                'Responsável': None,
                'Data da Análise': None,
                'Evidência / Protocolo': None,
                'Status do Tratamento': None
            })

    df_tratamento = pd.DataFrame(linhas_tratamento)

    log(f"Formatando Pasta de Trabalho Excel ({len(df_status)} colaboradores, {len(df_tratamento)} itens focalizados)...")
    if progress_fn: progress_fn(0.92)

    # 3. Criar Pasta de Trabalho com openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove aba padrão

    c_navy = '004B87'
    c_dark_navy = '002B49'
    c_soft_green = 'C6EFCE'
    c_dark_green = '006100'
    c_soft_red = 'FFC7CE'
    c_dark_red = '9C0006'
    c_soft_yellow = 'FFEB9C'
    c_dark_yellow = '9C6500'
    c_soft_blue = 'D9E1F2'
    c_border = 'D3D3D3'

    font_header = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    font_card_title = Font(name='Calibri', size=9, bold=True, color='004B87')
    font_card_val = Font(name='Calibri', size=18, bold=True, color='002B49')
    font_card_sub = Font(name='Calibri', size=8, italic=True, color='555555')
    font_regular = Font(name='Calibri', size=10)
    font_bold = Font(name='Calibri', size=10, bold=True)

    thin_border_side = Side(style='thin', color=c_border)
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    fill_header_navy = PatternFill(start_color=c_navy, end_color=c_navy, fill_type='solid')
    fill_header_dark = PatternFill(start_color=c_dark_navy, end_color=c_dark_navy, fill_type='solid')
    fill_card = PatternFill(start_color='F9FAFC', end_color='F9FAFC', fill_type='solid')
    fill_zebra = PatternFill(start_color='F7F9FB', end_color='F7F9FB', fill_type='solid')

    # ---------------------------------------------------------
    # ABA 1: Painel & Status
    # ---------------------------------------------------------
    ws1 = wb.create_sheet('Painel & Status')
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells('A1:R1')
    c1 = ws1['A1']
    c1.value = "PAINEL DE AUDITORIA RESCISÓRIA — CENCOSUD BRASIL"
    c1.font = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
    c1.fill = fill_header_navy
    c1.alignment = Alignment(horizontal='center', vertical='center')
    ws1.row_dimensions[1].height = 30

    ws1.merge_cells('A2:R2')
    c2 = ws1['A2']
    c2.value = f"Módulo AD003 | Cockpit de Governança, Triagem de Riscos e Liberação de Rescisões | Data de Processamento: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}"
    c2.font = Font(name='Calibri', size=9, italic=True, color='FFFFFF')
    c2.fill = fill_header_dark
    c2.alignment = Alignment(horizontal='center', vertical='center')
    ws1.row_dimensions[2].height = 20

    n_status = len(df_status)
    row_start_status = 11
    row_end_status = row_start_status + max(1, n_status)

    def criar_card_kpi(ws, col_start, col_end, row_title, row_val, row_sub, titulo, formula_val, subtitulo):
        col_s_letter = get_column_letter(col_start)
        col_e_letter = get_column_letter(col_end)
        
        ws.merge_cells(f"{col_s_letter}{row_title}:{col_e_letter}{row_title}")
        cell_t = ws.cell(row=row_title, column=col_start, value=titulo)
        cell_t.font = font_card_title
        cell_t.fill = fill_card
        cell_t.alignment = Alignment(horizontal='center', vertical='center')
        
        ws.merge_cells(f"{col_s_letter}{row_val}:{col_e_letter}{row_val}")
        cell_v = ws.cell(row=row_val, column=col_start, value=formula_val)
        cell_v.font = font_card_val
        cell_v.fill = fill_card
        cell_v.alignment = Alignment(horizontal='center', vertical='center')
        
        ws.merge_cells(f"{col_s_letter}{row_sub}:{col_e_letter}{row_sub}")
        cell_s = ws.cell(row=row_sub, column=col_start, value=subtitulo)
        cell_s.font = font_card_sub
        cell_s.fill = fill_card
        cell_s.alignment = Alignment(horizontal='center', vertical='center')
        
        for r in range(row_title, row_sub + 1):
            for c in range(col_start, col_end + 1):
                cell = ws.cell(row=r, column=c)
                top = Side(style='medium', color=c_navy) if r == row_title else None
                bottom = Side(style='medium', color=c_navy) if r == row_sub else None
                left = Side(style='medium', color=c_navy) if c == col_start else None
                right = Side(style='medium', color=c_navy) if c == col_end else None
                cell.border = Border(top=top, bottom=bottom, left=left, right=right)

    criar_card_kpi(ws1, 2, 4, 4, 5, 6, "RESCISÕES AUDITADAS", f"=COUNTA(A{row_start_status+1}:A{row_end_status})", "Total de colaboradores na base")
    criar_card_kpi(ws1, 6, 8, 4, 5, 6, "100% HOMOLOGADAS (LIBERADAS)", f'=COUNTIF(O{row_start_status+1}:O{row_end_status}, "LIBERADA")', "Sem pendências e aprovadas")
    criar_card_kpi(ws1, 10, 12, 4, 5, 6, "BLOQUEADAS / DIVERGÊNCIAS", f'=COUNTIF(O{row_start_status+1}:O{row_end_status}, "BLOQUEADA*")', "Pendências de recálculo")
    criar_card_kpi(ws1, 14, 16, 4, 5, 6, "ATENÇÃO / CASOS ATÍPICOS", f'=COUNTIF(O{row_start_status+1}:O{row_end_status}, "ATENÇÃO*")', "Faltas, pensões, limites legais")

    criar_card_kpi(ws1, 2, 4, 7, 8, 9, "DIVERGÊNCIAS CÁLCULO", '=COUNTIF(\'Validações Detalhadas\'!$V:$V, "DIVERGENTE*")', "Diferença > R$ 0,01")
    criar_card_kpi(ws1, 6, 8, 7, 8, 9, "ALERTAS FALTAS > LIMITE", '=COUNTIF(\'Divergências & Casos Atípicos\'!$M:$M, "Alerta de Faltas vs Data")', "Faltas excedendo mês de rescisão")
    criar_card_kpi(ws1, 10, 12, 7, 8, 9, "INSUFICIÊNCIA DE SALDO (499)", '=COUNTIF(\'Validações Detalhadas\'!$P:$P, 499)', "Rescisões zeradas / negativas")
    criar_card_kpi(ws1, 14, 16, 7, 8, 9, "PENSÕES ALIMENTÍCIAS", '=COUNTIF(\'Divergências & Casos Atípicos\'!$M:$M, "Alerta Jurídico - Pensão")', "Deduções judiciais")

    headers_status = [
        'Matrícula', 'Colaborador', 'Lotação Tributária', 'Cargo', 'Motivo do Desligamento', 
        'Data Admissão', 'Data Desligamento', 'Data Pagamento', 'Total Proventos (R$)', 
        'Total Descontos (R$)', 'Líquido TRCT (R$)', 'Divergências Cálculo', 'Casos Atípicos / Alertas', 
        'Decisão de Liberação', 'Status Final', 'Responsável Liberação', 'Data Liberação', 'Parecer / Observação'
    ]

    ws1.row_dimensions[row_start_status].height = 26
    for col_idx, h in enumerate(headers_status, 1):
        cell = ws1.cell(row=row_start_status, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_navy
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for i, row in df_status.iterrows():
        curr_row = row_start_status + 1 + i
        ws1.row_dimensions[curr_row].height = 20
        is_zebra = (i % 2 == 1)
        row_fill = fill_zebra if is_zebra else None
        
        ws1.cell(curr_row, 1, row['Matrícula']).font = font_bold
        ws1.cell(curr_row, 2, row['Colaborador']).font = font_regular
        ws1.cell(curr_row, 3, row['Lotação Tributária']).font = font_regular
        ws1.cell(curr_row, 4, row['Cargo']).font = font_regular
        ws1.cell(curr_row, 5, row['Motivo do Desligamento']).font = font_regular
        ws1.cell(curr_row, 6, row['Data Admissão']).font = font_regular
        ws1.cell(curr_row, 7, row['Data Desligamento']).font = font_regular
        ws1.cell(curr_row, 8, row['Data Pagamento']).font = font_regular
        
        c_p = ws1.cell(curr_row, 9, row['Total Proventos (R$)'])
        c_p.number_format = '#,##0.00'
        c_p.font = font_regular
        
        c_d = ws1.cell(curr_row, 10, row['Total Descontos (R$)'])
        c_d.number_format = '#,##0.00'
        c_d.font = font_regular
        
        c_l = ws1.cell(curr_row, 11, row['Líquido TRCT (R$)'])
        c_l.number_format = '#,##0.00'
        c_l.font = font_bold
        
        c_div = ws1.cell(curr_row, 12)
        c_div.value = f"=COUNTIFS('Divergências & Casos Atípicos'!$B:$B, A{curr_row}, 'Divergências & Casos Atípicos'!$M:$M, \"Divergência de Cálculo\")"
        c_div.font = font_bold
        c_div.alignment = Alignment(horizontal='center')
        
        c_atp = ws1.cell(curr_row, 13)
        c_atp.value = f"=COUNTIFS('Divergências & Casos Atípicos'!$B:$B, A{curr_row}, 'Divergências & Casos Atípicos'!$M:$M, \"<>Divergência de Cálculo\")"
        c_atp.font = font_bold
        c_atp.alignment = Alignment(horizontal='center')
        
        c_dec = ws1.cell(curr_row, 14, "Aguardando")
        c_dec.font = font_bold
        c_dec.alignment = Alignment(horizontal='center')
        
        c_stat = ws1.cell(curr_row, 15)
        c_stat.value = f'=IF(N{curr_row}="Reprovada", "REPROVADA", IF(L{curr_row}>0, "BLOQUEADA - Divergência", IF(M{curr_row}>0, "ATENÇÃO - Caso Atípico", IF(N{curr_row}="Aprovada", "LIBERADA", "AGUARDANDO APROVAÇÃO"))))'
        c_stat.font = font_bold
        c_stat.alignment = Alignment(horizontal='center')
        
        ws1.cell(curr_row, 16, None)
        ws1.cell(curr_row, 17, None)
        ws1.cell(curr_row, 18, None)
        
        for c_idx in range(1, 19):
            cell_curr = ws1.cell(curr_row, c_idx)
            cell_curr.border = border_cell
            if row_fill and c_idx not in [14, 15]:
                cell_curr.fill = row_fill

    if n_status > 0:
        dv_decisao = DataValidation(type="list", formula1='"Aguardando,Aprovada,Reprovada"', allow_blank=True)
        ws1.add_data_validation(dv_decisao)
        dv_decisao.add(f"N{row_start_status+1}:N{row_end_status}")

        rule_liberada = FormulaRule(formula=[f'ISNUMBER(SEARCH("LIBERADA", O{row_start_status+1}))'], 
                                    fill=PatternFill(start_color=c_soft_green, end_color=c_soft_green, fill_type='solid'),
                                    font=Font(color=c_dark_green, bold=True))
        rule_bloqueada = FormulaRule(formula=[f'OR(ISNUMBER(SEARCH("BLOQUEADA", O{row_start_status+1})), O{row_start_status+1}="REPROVADA")'], 
                                     fill=PatternFill(start_color=c_soft_red, end_color=c_soft_red, fill_type='solid'),
                                     font=Font(color=c_dark_red, bold=True))
        rule_atencao = FormulaRule(formula=[f'ISNUMBER(SEARCH("ATENÇÃO", O{row_start_status+1}))'], 
                                   fill=PatternFill(start_color=c_soft_yellow, end_color=c_soft_yellow, fill_type='solid'),
                                   font=Font(color=c_dark_yellow, bold=True))
        ws1.conditional_formatting.add(f"O{row_start_status+1}:O{row_end_status}", rule_liberada)
        ws1.conditional_formatting.add(f"O{row_start_status+1}:O{row_end_status}", rule_bloqueada)
        ws1.conditional_formatting.add(f"O{row_start_status+1}:O{row_end_status}", rule_atencao)

    ws1.auto_filter.ref = f"A{row_start_status}:R{row_end_status}"
    ws1.freeze_panes = f"A{row_start_status+1}"

    # ---------------------------------------------------------
    # ABA 2: Divergências & Casos Atípicos
    # ---------------------------------------------------------
    ws2 = wb.create_sheet('Divergências & Casos Atípicos')
    ws2.views.sheetView[0].showGridLines = True

    headers_trat = [
        'Linha Origem Validações', 'Matrícula', 'Colaborador', 'Lotação Tributária', 'Cargo', 
        'Motivo Desligamento', 'Data Desligamento', 'Código Evento', 'Descrição Evento', 
        'Valor Sistema (R$)', 'Valor Python (R$)', 'Diferença (R$)', 'Categoria da Ocorrência', 
        'Prioridade Sugerida', 'Motivo / Ponto de Atenção', 'Providência Sugerida', 
        'Conclusão do Analista', 'Responsável', 'Data da Análise', 'Evidência / Protocolo', 
        'Status do Tratamento'
    ]

    ws2.row_dimensions[1].height = 26
    for col_idx, h in enumerate(headers_trat, 1):
        cell = ws2.cell(row=1, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_navy
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    n_trat = len(df_tratamento)
    for i, row in df_tratamento.iterrows():
        curr_row = i + 2
        ws2.row_dimensions[curr_row].height = 20
        is_zebra = (i % 2 == 1)
        row_fill = fill_zebra if is_zebra else None
        
        ws2.cell(curr_row, 1, row['Linha Origem Validações']).font = font_regular
        ws2.cell(curr_row, 2, row['Matrícula']).font = font_bold
        ws2.cell(curr_row, 3, row['Colaborador']).font = font_regular
        ws2.cell(curr_row, 4, row['Lotação Tributária']).font = font_regular
        ws2.cell(curr_row, 5, row['Cargo']).font = font_regular
        ws2.cell(curr_row, 6, row['Motivo Desligamento']).font = font_regular
        ws2.cell(curr_row, 7, row['Data Desligamento']).font = font_regular
        ws2.cell(curr_row, 8, row['Código Evento']).font = font_bold
        ws2.cell(curr_row, 9, row['Descrição Evento']).font = font_regular
        
        c_vs = ws2.cell(curr_row, 10, row['Valor Sistema (R$)'])
        c_vs.number_format = '#,##0.00'
        c_vs.font = font_regular
        
        c_vp = ws2.cell(curr_row, 11, row['Valor Python (R$)'])
        c_vp.number_format = '#,##0.00'
        c_vp.font = font_regular
        
        c_dif = ws2.cell(curr_row, 12)
        c_dif.value = f'=IF(OR(J{curr_row}="",K{curr_row}=""),"",ABS(K{curr_row}-J{curr_row}))'
        c_dif.number_format = '#,##0.00'
        c_dif.font = font_bold
        
        ws2.cell(curr_row, 13, row['Categoria da Ocorrência']).font = font_bold
        
        c_prio = ws2.cell(curr_row, 14, row['Prioridade Sugerida'])
        c_prio.font = font_bold
        c_prio.alignment = Alignment(horizontal='center')
        
        ws2.cell(curr_row, 15, row['Motivo / Ponto de Atenção']).font = font_regular
        ws2.cell(curr_row, 16, row['Providência Sugerida']).font = font_regular
        
        c_conc = ws2.cell(curr_row, 17, "Pendente")
        c_conc.font = font_bold
        c_conc.alignment = Alignment(horizontal='center')
        
        ws2.cell(curr_row, 18, None)
        ws2.cell(curr_row, 19, None)
        ws2.cell(curr_row, 20, None)
        
        c_strt = ws2.cell(curr_row, 21)
        c_strt.value = f'=IF(OR(Q{curr_row}="",Q{curr_row}="Pendente"), "PENDENTE", IF(OR(R{curr_row}="",S{curr_row}=""), "PENDENTE - completar responsável/data", "TRATADA"))'
        c_strt.font = font_bold
        c_strt.alignment = Alignment(horizontal='center')
        
        for c_idx in range(1, 22):
            cell_curr = ws2.cell(curr_row, c_idx)
            cell_curr.border = border_cell
            if row_fill and c_idx not in [14, 17, 21]:
                cell_curr.fill = row_fill

    if n_trat > 0:
        dv_conclusao = DataValidation(type="list", formula1='"Pendente,Ajustado no FPW,Justificado,Validado Manualmente,Não se aplica"', allow_blank=True)
        ws2.add_data_validation(dv_conclusao)
        dv_conclusao.add(f"Q2:Q{n_trat + 1}")

        rule_prio_alta = FormulaRule(formula=['N2="Alta"'], fill=PatternFill(start_color=c_soft_red, end_color=c_soft_red, fill_type='solid'), font=Font(color=c_dark_red, bold=True))
        rule_prio_media = FormulaRule(formula=['N2="Média"'], fill=PatternFill(start_color=c_soft_yellow, end_color=c_soft_yellow, fill_type='solid'), font=Font(color=c_dark_yellow, bold=True))
        rule_prio_baixa = FormulaRule(formula=['N2="Baixa"'], fill=PatternFill(start_color=c_soft_blue, end_color=c_soft_blue, fill_type='solid'), font=Font(color='002060', bold=True))
        ws2.conditional_formatting.add(f"N2:N{n_trat+1}", rule_prio_alta)
        ws2.conditional_formatting.add(f"N2:N{n_trat+1}", rule_prio_media)
        ws2.conditional_formatting.add(f"N2:N{n_trat+1}", rule_prio_baixa)

        rule_trat_pend = FormulaRule(formula=['ISNUMBER(SEARCH("PENDENTE", U2))'], fill=PatternFill(start_color=c_soft_red, end_color=c_soft_red, fill_type='solid'), font=Font(color=c_dark_red, bold=True))
        rule_trat_ok = FormulaRule(formula=['U2="TRATADA"'], fill=PatternFill(start_color=c_soft_green, end_color=c_soft_green, fill_type='solid'), font=Font(color=c_dark_green, bold=True))
        ws2.conditional_formatting.add(f"U2:U{n_trat+1}", rule_trat_pend)
        ws2.conditional_formatting.add(f"U2:U{n_trat+1}", rule_trat_ok)

    ws2.auto_filter.ref = f"A1:U{max(2, n_trat + 1)}"
    ws2.freeze_panes = "A2"

    # ---------------------------------------------------------
    # ABA 3: Validações Detalhadas
    # ---------------------------------------------------------
    ws3 = wb.create_sheet('Validações Detalhadas')
    ws3.views.sheetView[0].showGridLines = True

    headers_v = list(df.columns)
    ws3.row_dimensions[1].height = 26
    for col_idx, h in enumerate(headers_v, 1):
        cell = ws3.cell(row=1, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_navy
        cell.alignment = Alignment(horizontal='center', vertical='center')

    n_val = len(df)
    status_col_idx = headers_v.index('STATUS AUDITORIA') + 1 if 'STATUS AUDITORIA' in headers_v else None

    for i, row in df.iterrows():
        curr_row = i + 2
        ws3.row_dimensions[curr_row].height = 18
        is_zebra = (i % 2 == 1)
        row_fill = fill_zebra if is_zebra else None
        
        st_val = str(row['STATUS AUDITORIA']) if status_col_idx else ""
        is_div = "DIVERGENTE" in st_val
        is_alerta = "ALERTA" in st_val
        
        for c_idx, h in enumerate(headers_v, 1):
            cell = ws3.cell(curr_row, c_idx, row[h])
            cell.font = font_regular
            cell.border = border_cell
            
            if h in ['VALOR EVENTO', 'VALOR CALCULADO PYTHON', 'DIVERGÊNCIA', 'VALOR EVENTO REFERÊNCIA']:
                cell.number_format = '#,##0.00'
                
            if is_div:
                cell.fill = PatternFill(start_color=c_soft_red, end_color=c_soft_red, fill_type='solid')
            elif is_alerta:
                cell.fill = PatternFill(start_color=c_soft_yellow, end_color=c_soft_yellow, fill_type='solid')
            elif row_fill:
                cell.fill = row_fill

    ws3.auto_filter.ref = f"A1:{get_column_letter(len(headers_v))}{n_val + 1}"
    ws3.freeze_panes = "A2"

    # ---------------------------------------------------------
    # ABA 4: Resumo Financeiro
    # ---------------------------------------------------------
    ws4 = wb.create_sheet('Resumo Financeiro')
    ws4.views.sheetView[0].showGridLines = True

    headers_r = list(df_resumo.columns)
    ws4.row_dimensions[1].height = 26
    for col_idx, h in enumerate(headers_r, 1):
        cell = ws4.cell(row=1, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_navy
        cell.alignment = Alignment(horizontal='center', vertical='center')

    n_res = len(df_resumo)
    for i, row in df_resumo.iterrows():
        curr_row = i + 2
        ws4.row_dimensions[curr_row].height = 18
        is_zebra = (i % 2 == 1)
        row_fill = fill_zebra if is_zebra else None
        
        for c_idx, h in enumerate(headers_r, 1):
            cell = ws4.cell(curr_row, c_idx, row[h])
            cell.border = border_cell
            if 'VALOR' in h or 'DIVERGÊNCIA' in h:
                cell.number_format = '#,##0.00'
            if h in ['MATRICULA COLABORADOR', 'CÓDIGO DO EVENTO']:
                cell.font = font_bold
            else:
                cell.font = font_regular
            if row_fill:
                cell.fill = row_fill

    ws4.auto_filter.ref = f"A1:{get_column_letter(len(headers_r))}{n_res + 1}"
    ws4.freeze_panes = "A2"

    # ---------------------------------------------------------
    # Ajustar larguras de colunas em todas as abas
    # ---------------------------------------------------------
    for ws in [ws1, ws2, ws3, ws4]:
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for r in range(1, min(ws.max_row + 1, 250)):
                if ws == ws1 and r in [1, 2, 4, 5, 6, 7, 8, 9]:
                    continue
                cell = ws.cell(r, col_idx)
                val_str = str(cell.value or '')
                if cell.number_format and '#,##0' in str(cell.number_format):
                    val_str += " R$ 000.000,00"
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(12, min(max_len + 3, 50))

    ws1.column_dimensions['A'].width = 14
    ws1.column_dimensions['B'].width = 32
    ws1.column_dimensions['C'].width = 30
    ws1.column_dimensions['D'].width = 24
    ws1.column_dimensions['E'].width = 28
    ws1.column_dimensions['L'].width = 18
    ws1.column_dimensions['M'].width = 18
    ws1.column_dimensions['N'].width = 20
    ws1.column_dimensions['O'].width = 25

    log("Salvando arquivo Excel consolidado...")
    wb.save(caminho_saida_xlsx)

    if progress_fn: progress_fn(1.0)
    log(f"Sucesso! Gerado arquivo {caminho_saida_xlsx}")
    return True

if __name__ == "__main__":
    if len(sys.argv) > 1:
        entrada = sys.argv[1]
        saida = sys.argv[2]
        auditar_folha_massiva(entrada, saida)
    else:
        # Default behavior
        entrada = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\rescisao\RFolha Rescisão Cencosud_Ago26.xlsb"
        saida = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\rescisao\Auditoria_Rescisao_Agosto26.xlsx"
        auditar_folha_massiva(entrada, saida)
