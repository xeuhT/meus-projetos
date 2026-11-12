import pandas as pd
import numpy as np
import re
import os
import time
import csv
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import warnings

warnings.filterwarnings('ignore')

# Arquivos padrão no ambiente corporativo
DEFAULT_PLANILHAS_DIR = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\férias\planilhas"
DEFAULT_OUTPUT_DIR = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\férias\conferencias"

def safe_read_csv(filepath, sep=';', encoding='latin1', skiprows=0, dtype=str, **kwargs):
    """
    Lê arquivos CSV nativos do sistema LG de forma ultra-resiliente, suportando:
    1. Variação no número de campos entre linhas (ex: linhas sem marcação de férias com delimitadores extras).
    2. Linhas de título isoladas antes do cabeçalho real.
    3. Normalização e truncamento de colunas para o tamanho exato do cabeçalho, evitando descartar linhas.
    """
    try:
        return pd.read_csv(filepath, sep=sep, encoding=encoding, skiprows=skiprows, dtype=dtype, **kwargs)
    except Exception:
        pass

    with open(filepath, 'r', encoding=encoding, errors='replace') as f:
        for _ in range(skiprows):
            f.readline()
        reader = csv.reader(f, delimiter=sep)
        header = None
        rows = []
        nrows = kwargs.get('nrows', None)
        for row in reader:
            if not row or not any(row):
                continue
            if header is None:
                # Se for linha de título isolada (1 coluna com texto descritivo e sem palavras-chave de colunas)
                if len(row) == 1 and not any(kw in str(row[0]).lower() for kw in ('matr', 'empresa', 'codigo', 'estabelecimento', 'colaborador')):
                    continue
                header = [c.strip() for c in row]
                max_cols = len(header)
                continue
            
            # Normaliza a quantidade de colunas da linha ao cabeçalho
            if len(row) > max_cols:
                row = row[:max_cols]
            elif len(row) < max_cols:
                row = row + [''] * (max_cols - len(row))
            rows.append(row)
            if nrows is not None and len(rows) >= nrows:
                break
            
        if header is None:
            return pd.DataFrame()
        return pd.DataFrame(rows, columns=header)

def safe_save_writer(writer_save_fn, target_path, prompt_lock_cb=None, log_cb=None):
    """Salva a planilha final com salvaguarda contra arquivos bloqueados no Excel."""
    os.makedirs(os.path.dirname(os.path.abspath(target_path)), exist_ok=True)
    tmp_path = target_path + f".tmp_{os.getpid()}_{int(time.time())}.xlsx"
    writer_save_fn(tmp_path)

    while True:
        try:
            if os.path.exists(target_path):
                with open(target_path, 'r+b'):
                    pass
                os.remove(target_path)
            os.rename(tmp_path, target_path)
            if log_cb:
                log_cb(f"Arquivo gerado e salvo: {os.path.basename(target_path)}")
            return target_path
        except (PermissionError, OSError):
            if log_cb:
                log_cb(f"Arquivo bloqueado pelo Excel: {os.path.basename(target_path)}")
            if prompt_lock_cb:
                action = prompt_lock_cb(target_path)
                if action == 'retry':
                    time.sleep(0.5)
                    continue
                elif action == 'copy':
                    base, ext = os.path.splitext(target_path)
                    new_path = base + f"_novo_{int(time.time())}" + ext
                    try:
                        os.rename(tmp_path, new_path)
                        if log_cb:
                            log_cb(f"Salvo como cópia alternativa: {os.path.basename(new_path)}")
                        return new_path
                    except Exception:
                        pass
            base, ext = os.path.splitext(target_path)
            new_path = base + f"_novo_{int(time.time())}" + ext
            try:
                os.rename(tmp_path, new_path)
                if log_cb:
                    log_cb(f"Salvo como cópia de emergência: {os.path.basename(new_path)}")
                return new_path
            except Exception:
                raise PermissionError(f"Arquivo bloqueado no Excel: {target_path}")

def run_conferencia_formatada(planilhas_dir=None, output_dir=None, file_map=None, progress_cb=None, log_cb=None, prompt_lock_cb=None):
    """
    Executa a geração homologada da planilha 'Automacao_Conferencia_Ferias_Formatada.xlsx'.
    Mantém 100% de paridade com o modelo oficial Cencosud validado.
    """
    def log(msg):
        print(msg)
        if log_cb:
            log_cb(msg)

    def progress(pct, msg):
        log(f"[{int(pct*100)}%] {msg}")
        if progress_cb:
            progress_cb(pct, msg)

    progress(0.05, "Inicializando motor de Conferência Férias Formatada...")

    if not planilhas_dir:
        planilhas_dir = DEFAULT_PLANILHAS_DIR
    if not output_dir:
        output_dir = DEFAULT_OUTPUT_DIR

    # Mapeamento de arquivos de entrada
    f_colab = None
    f_faltas = None
    f_ferias_prog = None
    f_ferias_colab = None
    f_orga = None
    f_base_pa = None
    f_afas = None
    f_feriados = None

    if file_map:
        f_colab = file_map.get('colab_master')
        f_faltas = file_map.get('faltas_pa')
        if file_map.get('prog_files'):
            f_ferias_prog = file_map.get('prog_files')[-1]
        f_ferias_colab = file_map.get('ferias_hist')
        f_orga = file_map.get('orga_file')
        f_base_pa = file_map.get('pa_aberto_model')
        f_afas = file_map.get('afastados')
        f_feriados = file_map.get('feriados')

    # Fallbacks relativos caso algum arquivo não esteja explicitamente no file_map
    def resolve_path(p):
        if not p:
            return ""
        if os.path.isabs(p):
            return p
        return os.path.join(planilhas_dir, p)

    if not f_colab:
        f_colab = resolve_path('bases compartilhadas/COLA_Colaboradores_CSV - 10092026_183455_824810.CSV')
    if not f_faltas:
        f_faltas = resolve_path('relatorios conferencia ferias formatada/COLA_Falta_Injustifica_PA - 10092026_184706_237965.CSV')
    if not f_ferias_prog:
        f_ferias_prog = resolve_path('relatorios conferencia ferias formatada/Férias Programadas ou Gozadas em CSV - 10092026_181249_423274.CSV')
    if not f_ferias_colab:
        f_ferias_colab = resolve_path('bases compartilhadas/Cencosud_Ferias_Colaboradores_CSV 01 - 10092026_181158_492867.CSV')
    if not f_orga:
        f_orga = resolve_path('relatorios conferencia ferias formatada/ORGA_Estabelecimento_CSV - 10092026_Confere Feriados Out2026.xlsx')
    if not f_base_pa:
        f_base_pa = resolve_path('planilhas de conferencia MODELO/COLA_Conferencia_PA_Aberto - 08092026_Férias Out 2026 MODELO.xlsx')
    if not f_afas:
        f_afas = resolve_path('relatorios folha mensal/COLA_Colaboradores_Afastados - 15092026_163904_199734 (1).CSV')
    if not f_feriados:
        f_feriados = resolve_path('relatorios conferencia ferias formatada/Feriados_Cadastrados10092026 (1).csv')

    target_file = os.path.join(output_dir, "Automacao_Conferencia_Ferias_Formatada.xlsx")

    progress(0.15, "Lendo arquivos de entrada e bases mestres...")
    import unicodedata
    def norm_str(s):
        if pd.isna(s): return ''
        return ''.join(c for c in unicodedata.normalize('NFD', str(s)) if unicodedata.category(c) != 'Mn').upper().strip()

    def clean_ctrl(s):
        if pd.isna(s): return ''
        return str(s).replace('\x1d', '').strip()

    # 1. Colaboradores
    df_colab = safe_read_csv(f_colab, sep=';', encoding='latin1', dtype=str)
    if not any('matr' in norm_str(c).lower() for c in df_colab.columns):
        df_colab = safe_read_csv(f_colab, sep=';', encoding='latin1', skiprows=1, dtype=str)
    mat_colab_name = next((c for c in df_colab.columns if 'matr' in norm_str(c).lower() and 'esocial' not in norm_str(c).lower()), df_colab.columns[3])
    sit_colab_name = next((c for c in df_colab.columns if 'situac' in norm_str(c).lower() and 'tipo' not in norm_str(c).lower()), df_colab.columns[-3])
    dict_situacao = dict(zip(df_colab[mat_colab_name].apply(clean_ctrl), df_colab[sit_colab_name].apply(clean_ctrl)))

    # 2. Faltas
    df_faltas = safe_read_csv(f_faltas, sep=';', encoding='latin1', dtype=str)
    if len(df_faltas.columns) < 3:
        df_faltas = safe_read_csv(f_faltas, sep=',', encoding='latin1', dtype=str)
    if not any('matr' in norm_str(c).lower() or 'colaborador' in norm_str(c).lower() for c in df_faltas.columns):
        df_faltas = safe_read_csv(f_faltas, sep=';', encoding='latin1', skiprows=1, dtype=str)
        if len(df_faltas.columns) < 3:
            df_faltas = safe_read_csv(f_faltas, sep=',', encoding='latin1', skiprows=1, dtype=str)
    dict_faltas = dict(zip(df_faltas.iloc[:, 2].apply(clean_ctrl), df_faltas.iloc[:, 8].apply(clean_ctrl)))

    # 3. Férias Programadas
    df_ferias_prog = safe_read_csv(f_ferias_prog, sep=';', encoding='latin1', dtype=str)
    if len(df_ferias_prog.columns) > 4 and 'matr' not in norm_str(df_ferias_prog.columns[4]).lower():
        df_ferias_prog = safe_read_csv(f_ferias_prog, sep=';', encoding='latin1', skiprows=1, dtype=str)
    idx_mat_prog = next((i for i, c in enumerate(df_ferias_prog.columns) if 'matr' in norm_str(c).lower()), 4)
    df_ferias_prog.rename(columns={df_ferias_prog.columns[idx_mat_prog]: 'Matricula'}, inplace=True)
    df_ferias_prog['Matricula'] = df_ferias_prog['Matricula'].apply(clean_ctrl)

    # 4. Férias Colaboradores
    df_ferias_colab = safe_read_csv(f_ferias_colab, sep=';', encoding='latin1', dtype=str)
    if not any('matr' in norm_str(c).lower() for c in df_ferias_colab.columns):
        df_ferias_colab = safe_read_csv(f_ferias_colab, sep=';', encoding='latin1', skiprows=1, dtype=str)
    if len(df_ferias_colab.columns) < 3:
        df_ferias_colab = safe_read_csv(f_ferias_colab, sep=',', encoding='latin1', dtype=str)
        if not any('matr' in norm_str(c).lower() for c in df_ferias_colab.columns):
            df_ferias_colab = safe_read_csv(f_ferias_colab, sep=',', encoding='latin1', skiprows=1, dtype=str)
    idx_mat_fcol = next((i for i, c in enumerate(df_ferias_colab.columns) if 'matr' in norm_str(c).lower()), 2)
    df_ferias_colab.rename(columns={df_ferias_colab.columns[idx_mat_fcol]: 'Matricula'}, inplace=True)
    df_ferias_colab['Matricula'] = df_ferias_colab['Matricula'].apply(clean_ctrl)

    # 5. Afastamentos
    df_afas = safe_read_csv(f_afas, sep=';', encoding='latin1', dtype=str)
    if not any('matr' in norm_str(c).lower() for c in df_afas.columns):
        df_afas = safe_read_csv(f_afas, sep=';', encoding='latin1', skiprows=1, dtype=str)
    col_mat_afas = next((c for c in df_afas.columns if 'matr' in norm_str(c).lower()), df_afas.columns[0])
    col_term_afas = next((c for c in df_afas.columns if 'termino' in norm_str(c).lower() or 'fim' in norm_str(c).lower()), df_afas.columns[-1])
    df_afas['dt_term'] = pd.to_datetime(df_afas[col_term_afas], format='%d/%m/%Y', dayfirst=True, errors='coerce')
    dict_afas_dt = {}
    for _, r in df_afas.iterrows():
        m = clean_ctrl(r[col_mat_afas])
        if pd.notna(r['dt_term']):
            dict_afas_dt[m] = r['dt_term']

    # 6. Feriados
    try:
        df_fer = safe_read_csv(f_feriados, sep=';', encoding='utf-8-sig')
    except Exception:
        df_fer = safe_read_csv(f_feriados, sep=';', encoding='latin1')
    df_fer['dt'] = pd.to_datetime(df_fer['Data'], format='%d/%m/%Y', dayfirst=True, errors='coerce')
    holidays_by_mun = df_fer.dropna(subset=['Municipio']).groupby(df_fer['Municipio'].apply(norm_str))['dt'].apply(lambda x: set(x.dropna().dt.date)).to_dict()
    vc_fer = df_fer['dt'].dt.date.value_counts()
    national_holidays = set(df_fer[df_fer['Municipio'].isna()]['dt'].dropna().dt.date) | set(vc_fer[vc_fer > 20].index)

    # 7. ORGA (suporta Excel ou CSV nativo LG)
    if str(f_orga).lower().endswith(('.xlsx', '.xls')):
        df_orga = pd.read_excel(f_orga)
        has_feriado_col_orga = len(df_orga.columns) > 12 and 'feriado' in norm_str(df_orga.columns[12]).lower()
    else:
        df_orga = safe_read_csv(f_orga, sep=';', encoding='latin1', dtype=str)
        if len(df_orga.columns) > 2 and 'estabelecimento' not in norm_str(df_orga.columns[2]).lower():
            df_orga = safe_read_csv(f_orga, sep=';', encoding='latin1', skiprows=1, dtype=str)
        has_feriado_col_orga = False

    def extrair_bandeira(est):
        e = str(est).upper()
        for b in ['ELETROSHOW', 'FARMACIA', 'GBARBOSA', 'BRETAS', 'PREZUNIC', 'GIGA', 'MERCANTIL', 'PERINI', 'SPID', 'ADM']:
            if b in e:
                return b
        if 'CD 703' in e or '700' in e:
            return 'PREZUNIC'
        return 'GBARBOSA'

    # 8. Base PA Aberto: suporta Excel Modelo (28 cols), CSV nativo LG (COLA_Conferencia_PA_Aberto) ou construção dinâmica
    if f_base_pa and os.path.exists(f_base_pa):
        if str(f_base_pa).lower().endswith(('.xlsx', '.xls')):
            df_raw_pa = pd.read_excel(f_base_pa, sheet_name=0)
            if not any('matr' in norm_str(c).lower() or 'empresa' in norm_str(c).lower() for c in df_raw_pa.columns):
                df_raw_pa = pd.read_excel(f_base_pa, sheet_name=0, skiprows=1)
        else:
            df_raw_pa = safe_read_csv(f_base_pa, sep=';', encoding='latin1', dtype=str)
            if not any('matr' in norm_str(c).lower() or 'empresa' in norm_str(c).lower() for c in df_raw_pa.columns):
                df_raw_pa = safe_read_csv(f_base_pa, sep=';', encoding='latin1', skiprows=1, dtype=str)
            if len(df_raw_pa.columns) < 3:
                df_raw_pa = safe_read_csv(f_base_pa, sep=',', encoding='latin1', dtype=str)
                if not any('matr' in norm_str(c).lower() or 'empresa' in norm_str(c).lower() for c in df_raw_pa.columns):
                    df_raw_pa = safe_read_csv(f_base_pa, sep=',', encoding='latin1', skiprows=1, dtype=str)

        # Deduplica colunas de df_raw_pa mantendo a primeira ocorrencia
        df_raw_pa = df_raw_pa.loc[:, ~df_raw_pa.columns.duplicated()]

        # Verifica se já é exatamente a planilha MODELO de 28 colunas (com Bandeira na col 2 e Situação na col 9)
        is_exact_modelo_28 = (
            len(df_raw_pa.columns) == 28
            and 'bandeira' in norm_str(df_raw_pa.columns[2]).lower()
            and 'situac' in norm_str(df_raw_pa.columns[9]).lower()
        )
        if is_exact_modelo_28:
            df_base_pa = df_raw_pa.copy()
            cols_list = list(df_base_pa.columns)
            cols_list[7] = 'Matricula'
            cols_list[10] = 'Inicio PA'
            cols_list[11] = 'Final PA'
            df_base_pa.columns = cols_list
            df_base_pa = df_base_pa.loc[:, ~df_base_pa.columns.duplicated()]
            df_base_pa['Matricula'] = df_base_pa['Matricula'].apply(clean_ctrl)
        else:
            log("Mapeando colunas do relatório nativo COLA_Conferencia_PA_Aberto para o layout de 28 colunas...")
            norm_cols = {c: norm_str(c) for c in df_raw_pa.columns}
            def find_col_pa(pred_fn):
                for c, cn in norm_cols.items():
                    if pred_fn(cn):
                        return c
                return None

            def get_series_pa(pred_fn, default_val='', clean=True, is_num=False, is_int=False):
                c = find_col_pa(pred_fn)
                if c is None:
                    return pd.Series([default_val] * len(df_raw_pa), index=df_raw_pa.index)
                s = df_raw_pa[c]
                if isinstance(s, pd.DataFrame):
                    s = s.iloc[:, 0]
                if is_num or is_int:
                    num_s = pd.to_numeric(s.astype(str).str.replace('\x1d', '', regex=False).str.replace(',', '.', regex=False).str.strip(), errors='coerce').fillna(default_val)
                    return num_s.astype(int) if is_int else num_s
                if clean:
                    return s.apply(clean_ctrl)
                return s

            s_cod_emp = get_series_pa(lambda cn: 'CODIGO' in cn and 'EMPRESA' in cn, default_val=2, is_int=True)
            s_emp = get_series_pa(lambda cn: cn in ('EMPRESA', 'NOME DA EMPRESA', 'RAZAO SOCIAL'))
            s_cod_est = get_series_pa(lambda cn: 'CODIGO' in cn and 'ESTABELECIMENTO' in cn, default_val=1, is_int=True)
            s_est = get_series_pa(lambda cn: cn in ('ESTABELECIMENTO', 'NOME DO ESTABELECIMENTO', 'UNIDADE ORGANIZACIONAL'))
            c_band = find_col_pa(lambda cn: 'BANDEIRA' in cn)
            s_band = get_series_pa(lambda cn: 'BANDEIRA' in cn) if c_band else s_est.apply(extrair_bandeira)
            s_cod_cc = get_series_pa(lambda cn: 'CODIGO' in cn and 'CENTRO' in cn)
            s_cc = get_series_pa(lambda cn: 'CENTRO' in cn and 'CUSTO' in cn and 'CODIGO' not in cn)
            s_mat = get_series_pa(lambda cn: 'MATR' in cn and 'ESOCIAL' not in cn)
            s_colab = get_series_pa(lambda cn: cn in ('COLABORADOR', 'NOME DO COLABORADOR', 'NOME COLABORADOR', 'NOME'))
            s_sit = get_series_pa(lambda cn: 'SITUAC' in cn and 'TIPO' not in cn)
            s_ini_pa = get_series_pa(lambda cn: ('INICIO' in cn or 'INI' in cn) and ('PA' in cn or 'AQUISITIVO' in cn), clean=False)
            s_fim_pa = get_series_pa(lambda cn: ('FINAL' in cn or 'FIM' in cn or 'TERMINO' in cn) and ('PA' in cn or 'AQUISITIVO' in cn), clean=False)
            s_avos = get_series_pa(lambda cn: 'AVOS' in cn)
            s_ini_fru = get_series_pa(lambda cn: ('INICIO' in cn or 'INI' in cn) and ('FRUICAO' in cn or 'CONCESSIVO' in cn), clean=False)
            s_fim_fru = get_series_pa(lambda cn: ('FINAL' in cn or 'FIM' in cn or 'LIMITE' in cn) and ('FRUICAO' in cn or 'CONCESSIVO' in cn), clean=False)
            s_regra = get_series_pa(lambda cn: 'REGRA' in cn)
            s_tipo_fer = get_series_pa(lambda cn: 'TIPO' in cn and 'FERIAS' in cn)
            s_ret_afas = get_series_pa(lambda cn: 'RETORNO' in cn and 'AFAS' in cn)
            s_ini_desc = get_series_pa(lambda cn: ('INICIO' in cn or 'INI' in cn) and ('DESCANSO' in cn or 'GOZO' in cn), clean=False)
            s_fim_desc = get_series_pa(lambda cn: ('FIM' in cn or 'FINAL' in cn or 'TERMINO' in cn) and ('DESCANSO' in cn or 'GOZO' in cn), clean=False)
            s_dias_desc = get_series_pa(lambda cn: 'DIAS' in cn and ('DESCANSO' in cn or 'GOZO' in cn), default_val=0, is_num=True)
            s_dias_fer = get_series_pa(lambda cn: 'DIAS' in cn and 'FERIAS' in cn, default_val=30, is_num=True)
            s_dias_lic = get_series_pa(lambda cn: 'DIAS' in cn and 'LICENCA' in cn, default_val=0, is_num=True)
            s_dias_abono = get_series_pa(lambda cn: 'ABONO' in cn, default_val=0, is_num=True)
            s_faltas = get_series_pa(lambda cn: 'FALTA' in cn, default_val=0, is_num=True)
            s_dias_descons = get_series_pa(lambda cn: 'DESCONSIDERAR' in cn, default_val=0, is_num=True)
            s_parcela = get_series_pa(lambda cn: 'PARCELA' in cn and '13' not in cn)
            s_13 = get_series_pa(lambda cn: '13' in cn or 'DECIMO' in cn, default_val='N')

            df_base_pa = pd.DataFrame({
                'Código Empresa': s_cod_emp,
                'Empresa': s_emp,
                'Bandeira': s_band,
                'Código Estabelecimento': s_cod_est,
                'Estabelecimento': s_est,
                'Código Centro Custo': s_cod_cc,
                'Centro de Custo': s_cc,
                'Matricula': s_mat,
                'Colaborador': s_colab,
                'Situação': s_sit,
                'Inicio PA': s_ini_pa,
                'Final PA': s_fim_pa,
                'Avos': s_avos,
                'Início Fruição': s_ini_fru,
                'Final Fruição': s_fim_fru,
                'Regra': s_regra,
                'Tipo Férias': s_tipo_fer,
                'Retorno Afas': s_ret_afas,
                'Início Descanso': s_ini_desc,
                'Fim Descanso': s_fim_desc,
                'Dias Descanso': s_dias_desc,
                'Dias Férias': s_dias_fer,
                'Dias Licença': s_dias_lic,
                'Dias Abono': s_dias_abono,
                'Faltas': s_faltas,
                'Dias a Desconsiderar': s_dias_descons,
                'Parcela': s_parcela,
                '13º': s_13,
            })
    else:
        log("Construindo estrutura PA Aberto no layout do modelo a partir das bases CSV...")

        # Mapear programações de férias (de Cencosud_Ferias_Colaboradores_CSV e Férias Programadas)
        prog_by_mat = {}
        for _, r in df_ferias_colab.iterrows():
            m = str(r['Matricula']).strip()
            if not m or m == 'nan':
                continue
            ini_pa = str(r.iloc[6]).strip() if pd.notna(r.iloc[6]) else ''
            fim_pa = str(r.iloc[7]).strip() if pd.notna(r.iloc[7]) else ''
            ini_gozo = str(r.iloc[9]).strip() if pd.notna(r.iloc[9]) else ''
            fim_gozo = str(r.iloc[10]).strip() if pd.notna(r.iloc[10]) else ''
            dias_gozo = pd.to_numeric(r.iloc[11], errors='coerce')
            dias_desc = pd.to_numeric(r.iloc[12], errors='coerce')
            abono = pd.to_numeric(r.iloc[13], errors='coerce')
            parcela = str(r.iloc[14]).strip() if pd.notna(r.iloc[14]) else ''
            dec_terc = 'S' if str(r.iloc[15]).strip().lower().startswith('s') else 'N'
            prog_by_mat.setdefault(m, []).append({
                'Inicio PA': ini_pa,
                'Final PA': fim_pa,
                'Tipo Férias': 'Normais' if ini_gozo else '',
                'Início Descanso': ini_gozo,
                'Fim Descanso': fim_gozo,
                'Dias Descanso': 0 if pd.isna(dias_gozo) else int(dias_gozo),
                'Dias Abono': 0 if pd.isna(abono) else int(abono),
                'Dias a Desconsiderar': 0 if pd.isna(dias_desc) else int(dias_desc),
                'Parcela': parcela if parcela != 'nan' else '',
                '13º': dec_terc
            })

        rows_pa = []
        ref_year = pd.Timestamp.today().year
        for _, r in df_colab.iterrows():
            m = str(r[mat_colab_name]).strip()
            if not m or m == 'nan':
                continue
            sit = str(r[sit_colab_name]).strip()
            if 'demitido' in sit.lower() or 'rescis' in sit.lower():
                continue
            cod_emp = pd.to_numeric(r.iloc[0], errors='coerce')
            emp = str(r.iloc[1]).strip()
            cod_est = pd.to_numeric(r.iloc[44], errors='coerce')
            est = str(r.iloc[45]).strip()
            band = extrair_bandeira(est)
            cod_cc = str(r.iloc[52]).strip().lstrip('\x1d')
            cc = str(r.iloc[53]).strip()
            nome = str(r.iloc[4]).strip()
            dt_adm = pd.to_datetime(r.iloc[7], format='%d/%m/%Y', dayfirst=True, errors='coerce')

            progs = prog_by_mat.get(m, [])
            if progs:
                for p in progs:
                    rows_pa.append({
                        'Código Empresa': int(cod_emp) if pd.notna(cod_emp) else r.iloc[0],
                        'Empresa': emp,
                        'Bandeira': band,
                        'Código Estabelecimento': int(cod_est) if pd.notna(cod_est) else r.iloc[44],
                        'Estabelecimento': est,
                        'Código Centro Custo': cod_cc,
                        'Centro de Custo': cc,
                        'Matricula': m,
                        'Colaborador': nome,
                        'Situação': sit,
                        'Inicio PA': p['Inicio PA'],
                        'Final PA': p['Final PA'],
                        'Avos': '',
                        'Início Fruição': '',
                        'Final Fruição': '',
                        'Regra': '',
                        'Tipo Férias': p['Tipo Férias'],
                        'Retorno Afas': '',
                        'Início Descanso': p['Início Descanso'],
                        'Fim Descanso': p['Fim Descanso'],
                        'Dias Descanso': p['Dias Descanso'],
                        'Dias Férias': 30,
                        'Dias Licença': 0,
                        'Dias Abono': p['Dias Abono'],
                        'Faltas': 0,
                        'Dias a Desconsiderar': p['Dias a Desconsiderar'],
                        'Parcela': p['Parcela'],
                        '13º': p['13º']
                    })
            else:
                ini_pa_str, fim_pa_str = '', ''
                if pd.notna(dt_adm):
                    try:
                        ini_pa_dt = dt_adm.replace(year=ref_year - 1)
                    except ValueError:
                        ini_pa_dt = dt_adm + pd.DateOffset(years=ref_year - 1 - dt_adm.year)
                    fim_pa_dt = ini_pa_dt + pd.DateOffset(years=1) - pd.Timedelta(days=1)
                    ini_pa_str = ini_pa_dt.strftime('%d/%m/%Y')
                    fim_pa_str = fim_pa_dt.strftime('%d/%m/%Y')
                rows_pa.append({
                    'Código Empresa': int(cod_emp) if pd.notna(cod_emp) else r.iloc[0],
                    'Empresa': emp,
                    'Bandeira': band,
                    'Código Estabelecimento': int(cod_est) if pd.notna(cod_est) else r.iloc[44],
                    'Estabelecimento': est,
                    'Código Centro Custo': cod_cc,
                    'Centro de Custo': cc,
                    'Matricula': m,
                    'Colaborador': nome,
                    'Situação': sit,
                    'Inicio PA': ini_pa_str,
                    'Final PA': fim_pa_str,
                    'Avos': '',
                    'Início Fruição': '',
                    'Final Fruição': '',
                    'Regra': '',
                    'Tipo Férias': '',
                    'Retorno Afas': '',
                    'Início Descanso': '',
                    'Fim Descanso': '',
                    'Dias Descanso': 0,
                    'Dias Férias': 30,
                    'Dias Licença': 0,
                    'Dias Abono': 0,
                    'Faltas': 0,
                    'Dias a Desconsiderar': 0,
                    'Parcela': '',
                    '13º': 'N'
                })
        df_base_pa = pd.DataFrame(rows_pa)

    # ==========================================
    # ABA 1: Conferencia_PA_Aberto
    # ==========================================
    progress(0.35, "Processando Aba 1: Conferencia_PA_Aberto...")
    df_aba1 = df_base_pa.copy()
    df_aba1 = df_aba1.loc[:, ~df_aba1.columns.duplicated()]

    # Situação da coluna DC
    df_aba1['Situação'] = df_aba1['Matricula'].map(dict_situacao).fillna('Atividade Normal')

    # Datas do PA
    df_aba1['dt_ini_pa'] = pd.to_datetime(df_aba1['Inicio PA'], dayfirst=True, errors='coerce')
    df_aba1['dt_fim_pa'] = pd.to_datetime(df_aba1['Final PA'], dayfirst=True, errors='coerce')

    # Fruição e Regra
    df_aba1['Início Fruição'] = df_aba1['dt_fim_pa'] + pd.Timedelta(days=1)
    df_aba1['Final Fruição'] = df_aba1['Início Fruição'] + pd.DateOffset(months=11) - pd.Timedelta(days=1)
    df_aba1['Regra'] = df_aba1['Final Fruição'] - pd.Timedelta(days=60)

    # Tipo Férias: quem está vazio fica 'Pendente de Programação'
    if 'Tipo Férias' not in df_aba1.columns:
        df_aba1['Tipo Férias'] = ''
    df_aba1['Tipo Férias'] = df_aba1['Tipo Férias'].replace('', np.nan).fillna('Pendente de Programação')

    # Retorno Afas: data se estiver dentro do PA
    def get_retorno_afas(row):
        m = row['Matricula']
        dt_t = dict_afas_dt.get(m, None)
        if dt_t is not None and pd.notna(row['dt_ini_pa']) and pd.notna(row['dt_fim_pa']):
            if row['dt_ini_pa'] <= dt_t <= row['dt_fim_pa']:
                return dt_t.strftime('%d/%m/%Y')
        return ''

    df_aba1['Retorno Afas'] = df_aba1.apply(get_retorno_afas, axis=1)

    # Risco de Dobra
    today_ts = pd.Timestamp.now()
    def get_risco_dobra(row):
        if pd.notna(row['Final Fruição']):
            if (row['Final Fruição'] - today_ts).days <= 60 and row['Tipo Férias'] == 'Pendente de Programação':
                return 'Risco de Dobra'
        return ''
    df_aba1.insert(18, 'Risco de Dobra', df_aba1.apply(get_risco_dobra, axis=1))
    df_aba1.insert(19, 'Conferencia PA', '')

    # Faltas
    df_aba1['Faltas'] = df_aba1['Matricula'].map(dict_faltas).fillna(0)
    def obter_dias_ferias(faltas):
        try: f = int(float(faltas))
        except: return 30
        if f <= 5: return 30
        elif f <= 14: return 24
        elif f <= 23: return 18
        elif f <= 32: return 12
        else: return 0
    df_aba1['Dias Férias'] = df_aba1['Faltas'].apply(obter_dias_ferias)
    if 'Dias Licença' in df_aba1.columns:
        df_aba1['Dias Licença'] = 0

    # Formatar datas como texto dd/mm/yyyy
    date_cols_aba1 = ['Inicio PA', 'Final PA', 'Início Fruição', 'Final Fruição', 'Regra', 'Início Descanso', 'Fim Descanso']
    for c in date_cols_aba1:
        if c in df_aba1.columns:
            df_aba1[c] = pd.to_datetime(df_aba1[c], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y').fillna('')

    def calc_avos(row):
        try:
            dt_fim = pd.to_datetime(row['Final PA'], dayfirst=True, errors='coerce')
            dt_ini = pd.to_datetime(row['Inicio PA'], dayfirst=True, errors='coerce')
            hoje = pd.Timestamp.today() + pd.offsets.MonthEnd(0)
            if pd.isna(dt_fim) or pd.isna(dt_ini): return ''
            if hoje >= dt_fim: return 12
            if hoje < dt_ini: return 0
            diff_months = (hoje.year - dt_ini.year) * 12 + hoje.month - dt_ini.month
            if hoje.day < dt_ini.day: diff_months -= 1
            dt_temp = dt_ini + pd.DateOffset(months=diff_months)
            diff_days = (hoje - dt_temp).days
            avos = diff_months + (1 if diff_days >= 15 else 0)
            return min(12, max(0, avos))
        except:
            return ''

    df_aba1['Avos'] = df_aba1.apply(calc_avos, axis=1)

    def calc_conf_pa(row):
        try:
            dt_fim = pd.to_datetime(row['Final PA'], dayfirst=True, errors='coerce')
            dt_ini = pd.to_datetime(row['Inicio PA'], dayfirst=True, errors='coerce')
            if pd.isna(dt_fim) or pd.isna(dt_ini): return 'Verificar'
            days = (dt_fim - dt_ini).days
            return 'OK' if days in [364, 365] else 'Verificar'
        except:
            return 'Verificar'

    df_aba1['Conferencia PA'] = df_aba1.apply(calc_conf_pa, axis=1)

    df_aba1.drop(columns=['dt_ini_pa', 'dt_fim_pa'], inplace=True, errors='ignore')
    df_aba1 = df_aba1.fillna('')

    # ==========================================
    # ABA 2: Ferias_Colaboradores
    # ==========================================
    progress(0.60, "Processando Aba 2: Ferias_Colaboradores...")
    df_aba2 = df_ferias_colab.copy()

    def extrair_loja(unidade):
        match = re.search(r'\b\d{1,4}\b', str(unidade))
        return match.group(0) if match else ''

    idx_unid = next((i for i, c in enumerate(df_aba2.columns) if 'UNIDADE' in norm_str(c)), 5)
    df_aba2.insert(2, 'Loja', df_aba2.iloc[:, idx_unid].apply(extrair_loja))
    df_aba2.insert(13, 'Faltas', df_aba2['Matricula'].map(dict_faltas).fillna(0))

    col_fim_aquisitivo = next((c for c in df_aba2.columns if 'FIM' in norm_str(c) and 'AQUISITIVO' in norm_str(c)), df_aba2.columns[8])
    col_inicio_gozo = next((c for c in df_aba2.columns if ('INICIO' in norm_str(c) or 'INI' in norm_str(c)) and 'GOZO' in norm_str(c)), df_aba2.columns[10])
    col_fim_gozo = next((c for c in df_aba2.columns if 'FIM' in norm_str(c) and 'GOZO' in norm_str(c)), df_aba2.columns[11])
    col_limite = next((c for c in df_aba2.columns if 'LIMITE' in norm_str(c)), df_aba2.columns[9])
    col_retorno = next((c for c in df_aba2.columns if 'RETORNO' in norm_str(c)), df_aba2.columns[-3])
    col_dias_gozo = next((c for c in df_aba2.columns if 'DIAS' in norm_str(c) and 'GOZO' in norm_str(c)), df_aba2.columns[12])

    dt_fim_gozo_s = pd.to_datetime(df_aba2[col_fim_gozo], format='%d/%m/%Y', dayfirst=True, errors='coerce')
    dt_limite_s = pd.to_datetime(df_aba2[col_limite], format='%d/%m/%Y', dayfirst=True, errors='coerce')
    dt_retorno_s = pd.to_datetime(df_aba2[col_retorno], format='%d/%m/%Y', dayfirst=True, errors='coerce')
    dt_inicio_gozo_s = pd.to_datetime(df_aba2[col_inicio_gozo], format='%d/%m/%Y', dayfirst=True, errors='coerce')
    dt_fim_aquisitivo_s = pd.to_datetime(df_aba2[col_fim_aquisitivo], format='%d/%m/%Y', dayfirst=True, errors='coerce')

    dict_mun_orga = dict(zip(df_orga.iloc[:, 2].astype(str).str.strip(), df_orga.iloc[:, 11].apply(norm_str)))
    if has_feriado_col_orga:
        dict_feriados_orga = dict(zip(df_orga.iloc[:, 2].astype(str).str.strip(), df_orga.iloc[:, 12]))
        df_aba2.insert(14, 'Feriados', df_aba2['Loja'].map(dict_feriados_orga))
        df_aba2['Feriados'] = pd.to_datetime(df_aba2['Feriados'], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y').fillna('#N/A')
    else:
        meses_gozo = dt_inicio_gozo_s.dropna().dt.month.tolist()
        mes_ref = max(set(meses_gozo), key=meses_gozo.count) if meses_gozo else 10
        df_mun_fer = df_fer[(df_fer['dt'].dt.month == mes_ref) & (df_fer['Municipio'].notna()) & (~df_fer['dt'].dt.date.isin(national_holidays))].drop_duplicates(subset=['Municipio'], keep='first')
        map_mun_fer = {norm_str(r['Municipio']): r['dt'].strftime('%d/%m/%Y') for _, r in df_mun_fer.iterrows()}
        dict_feriados_orga = {str(r.iloc[2]).strip(): map_mun_fer.get(norm_str(r.iloc[11]), '#N/A') for _, r in df_orga.iterrows()}
        df_aba2.insert(14, 'Feriados', df_aba2['Loja'].map(dict_feriados_orga).fillna('#N/A'))

    df_aba2['Municipio_Check'] = df_aba2['Loja'].map(dict_mun_orga)

    # 1. Validação Dias de Gozo x Faltas (Art. 130 CLT)
    def validar_gozo_faltas(row, idx):
        try: f = int(float(row['Faltas']))
        except: f = 0
        try: d = int(float(row[col_dias_gozo]))
        except: d = 30
        lim = obter_dias_ferias(f)
        if f > 6 and d > lim:
            return 'Ajustar Dias de Gozo'
        return 'OK'

    val_gozo_faltas_s = [validar_gozo_faltas(row, i) for i, row in df_aba2.iterrows()]
    df_aba2.insert(14, 'Validação Dias de Gozo x Faltas', val_gozo_faltas_s)

    # 2. Validação da Dt Início Gozo x Feriado (CLT Art. 134 § 3º)
    def validar_inicio_feriado(dt, mun):
        if pd.isna(dt): return None
        wd = dt.weekday()
        if wd in [4, 5, 6]:
            return 'Ajustar Dt. Início de Gozo'
        
        local_holidays = holidays_by_mun.get(norm_str(mun), set()) | national_holidays
        
        d0 = dt.date()
        d1 = (dt + pd.Timedelta(days=1)).date()
        d2 = (dt + pd.Timedelta(days=2)).date()
        if d0 in local_holidays or d1 in local_holidays or d2 in local_holidays:
            return 'Ajustar Dt. Início de Gozo'
        return 'OK'

    val_inicio_feriado_s = [validar_inicio_feriado(dt, mun) for dt, mun in zip(dt_inicio_gozo_s, df_aba2['Municipio_Check'])]

    del df_aba2['Municipio_Check']
    df_aba2.insert(16, 'Validação da Dt Início Gozo x Feriado', val_inicio_feriado_s)

    # 3. Conf PA x Gozo e Conf Dobra
    df_aba2.insert(17, 'Conf PA x Gozo', dt_inicio_gozo_s >= dt_fim_aquisitivo_s)
    df_aba2.insert(18, 'Conf Dobra', dt_fim_gozo_s <= dt_limite_s)

    # 4. Conf.Dt Retorno
    conf_retorno_days = (dt_retorno_s - dt_fim_gozo_s).dt.days
    df_aba2.insert(24, 'Conf.Dt Retorno', conf_retorno_days)

    # 5. Validação Dt de retorno para Centro de Custo de Loja logistica
    def validar_retorno_loja(row, idx):
        unid = str(row.get('Unidade Organizacional', '')).upper()
        is_loja_log = any(k in unid for k in ['LOJA', 'SUPER', 'HIPER', 'GIGA', 'CD', 'CENTRO', 'LOGISTICA', 'POSTO', 'FARMACIA']) or str(row.get('Loja', '')).strip() != ''
        diff = conf_retorno_days.iloc[idx]
        if is_loja_log and pd.notna(diff) and diff > 1:
            return 'Ajustar Dt de retorno'
        return 'OK'

    val_retorno_loja_s = [validar_retorno_loja(row, i) for i, row in df_aba2.iterrows()]
    df_aba2.insert(25, 'Validação Dt de retorno para Centro de Custo de Loja logistica ', val_retorno_loja_s)

    # 6. Ação Consolidada
    col_13 = next((c for c in df_aba2.columns if '13' in norm_str(c) or 'decimo' in norm_str(c)), None)

    def consolidar_acoes(row, idx):
        acoes = []
        if val_gozo_faltas_s[idx] == 'Ajustar Dias de Gozo':
            acoes.append('Dias de gozo pendente de ajuste')
        if val_inicio_feriado_s[idx] == 'Ajustar Dt. Início de Gozo':
            acoes.append('Dt início de gozo pendente de ajuste')
        if val_retorno_loja_s[idx] == 'Ajustar Dt de retorno':
            dt_ret = dt_retorno_s.iloc[idx]
            diff = conf_retorno_days.iloc[idx]
            if pd.notna(dt_ret) and dt_ret.weekday() == 0 and diff in [2, 3]:
                acoes.append('Dt Retorno Ajustada')
            else:
                acoes.append('Dt de retorno pendente de ajuste')
        if col_13:
            v13 = str(row.get(col_13, '')).strip().lower()
            if v13 in ['sim', 's', 'true', '1']:
                dt_ini = dt_inicio_gozo_s.iloc[idx]
                if pd.notna(dt_ini) and dt_ini.month in [11, 12, 1]:
                    acoes.append('Não permitido adiantamento de 13º em Nov/Dez/Jan')
        return ', '.join(acoes) if acoes else 'OK'

    df_aba2.insert(28, 'Ação', [consolidar_acoes(row, i) for i, row in df_aba2.iterrows()])

    df_aba2.rename(columns={col_limite: 'Data Limite para Final do Gozo'}, inplace=True)

    for col in [col_fim_gozo, col_retorno, col_inicio_gozo, col_fim_aquisitivo, 'Data Limite para Final do Gozo']:
        if col in df_aba2.columns:
            df_aba2[col] = pd.to_datetime(df_aba2[col], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y').fillna('')

    df_aba2 = df_aba2.fillna('')

    # ==========================================
    # GRAVANDO E APLICANDO ESTILOS NO EXCEL
    # ==========================================
    progress(0.85, "Gravando arquivo Excel e aplicando estilização visual corporativa...")

    def write_excel(file_to_save):
        with pd.ExcelWriter(file_to_save, engine='openpyxl') as writer:
            df_aba1.to_excel(writer, sheet_name='Conferencia_PA_Aberto', index=False)
            df_aba2.to_excel(writer, sheet_name='Ferias_Colaboradores', index=False)
            
            ws1 = writer.sheets['Conferencia_PA_Aberto']
            ws2 = writer.sheets['Ferias_Colaboradores']
            
            header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            red_font = Font(color="9C0006", bold=True)
            
            left_cols = ['Empresa', 'Estabelecimento', 'Centro de Custo', 'Colaborador', 'Bandeira', 'Razão Social', 'Unidade Organizacional']
            
            # Formatação Aba 1
            for r in range(2, ws1.max_row + 1):        
                val_s = ws1.cell(row=r, column=19).value
                if val_s == 'Risco de Dobra':
                    ws1.cell(row=r, column=19).fill = red_fill
                    ws1.cell(row=r, column=19).font = red_font
                    
                val_r = ws1.cell(row=r, column=18).value
                if val_r and str(val_r).strip() != '':
                    ws1.cell(row=r, column=18).fill = yellow_fill

            for row in ws1.iter_rows(min_row=2):
                for cell in row:
                    col_name = str(ws1.cell(row=1, column=cell.column).value)
                    if any(x.lower() in col_name.lower() for x in left_cols):
                        cell.alignment = Alignment(horizontal='left', vertical='center')
                    else:
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        
            for cell in ws1[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
            for col in ws1.columns:
                col_letter = col[0].column_letter
                max_len = max(len(str(c.value or '')) for c in col[:100])
                ws1.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 40)
            ws1.column_dimensions['J'].width = 25
            ws1.column_dimensions['M'].width = 10
            ws1.column_dimensions['Q'].width = 25
            ws1.column_dimensions['R'].width = 15
            ws1.column_dimensions['S'].width = 16
            ws1.column_dimensions['T'].width = 16

            # Formatação Aba 2
            for r in range(2, ws2.max_row + 1):        
                val_ac = ws2.cell(row=r, column=29).value
                if val_ac and val_ac != 'OK':
                    ws2.cell(row=r, column=29).fill = yellow_fill
                    ws2.cell(row=r, column=29).font = Font(color="000000", bold=True)

            for row in ws2.iter_rows(min_row=2):
                for cell in row:
                    col_name = str(ws2.cell(row=1, column=cell.column).value)
                    if any(x.lower() in col_name.lower() for x in left_cols):
                        cell.alignment = Alignment(horizontal='left', vertical='center')
                    else:
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        
            for cell in ws2[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
            for col in ws2.columns:
                col_letter = col[0].column_letter
                max_len = max(len(str(c.value or '')) for c in col[:100])
                ws2.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 42)
            ws2.column_dimensions['J'].width = 28
            ws2.column_dimensions['O'].width = 28
            ws2.column_dimensions['Q'].width = 30
            ws2.column_dimensions['Z'].width = 32
            ws2.column_dimensions['AC'].width = 35

    saved_path = safe_save_writer(write_excel, target_file, prompt_lock_cb=prompt_lock_cb, log_cb=log)
    progress(1.0, f"Sucesso! {os.path.basename(saved_path)} atualizada!")
    return saved_path

if __name__ == '__main__':
    run_conferencia_formatada()
