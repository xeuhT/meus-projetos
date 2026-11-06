import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import os
import shutil
import datetime
import time
import re
import csv
import unicodedata

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
                if len(row) == 1 and not any(kw in str(row[0]).lower() for kw in ('matr', 'empresa', 'codigo', 'estabelecimento', 'colaborador')):
                    continue
                header = [c.strip() for c in row]
                max_cols = len(header)
                continue
            
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

DEFAULT_PLANILHAS_DIR = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\férias\planilhas"
DEFAULT_OUTPUT_DIR = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\férias\conferencias"

MESES_PT = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}
MESES_ABREV = {
    1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr',
    5: 'Mai', 6: 'Jun', 7: 'Jul', 8: 'Ago',
    9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'
}

def extract_fopa_competencia(fopa_path):
    """Extrai mês e ano da competência do cabeçalho do arquivo FOPA (ex: (10, 2026))."""
    if not fopa_path or not os.path.exists(fopa_path):
        return None, None
    try:
        with open(fopa_path, 'r', encoding='latin1') as f:
            for _ in range(5):
                line = next(f, '')
                if 'Compet' in line:
                    parts = [p.strip().strip('"') for p in line.split(';')]
                    for p in parts:
                        m = re.search(r'(\d{1,2})/(\d{4})', p)
                        if m:
                            return int(m.group(1)), int(m.group(2))
    except Exception:
        pass
    return None, None

# Cabeçalhos Homologados e Validados pela Equipe de Auditoria
OFFICIAL_HEADERS_CALC = [
    'Código da Empresa', 'Nome da Empresa', 'Código Estabelecimento', 'Estabelecimento',
    'Código Centro Custo', 'Centro Custo', 'Matrícula', 'Nome do Colaborador', 'ID Colaborador',
    'Data Admissão', 'Confere', 'Tempo de Casa', 'Código Cargo', 'Cargo', 'Situação', 'Sindicato',
    'PA', 'Férias Outubro', 'Insalubridade/Periculosidade', 'Validação Sal Insalub/Pericul',
    '11220 - Remuneracao Fixa Mensal', '11000 - Salario Contratual', 'Conf 501', 'Validação 501',
    'Validação Fracionadas', 'Validação x Férias Fracionadas/ Faltas', 'Validar Vendedor', 'Validar Vendedor II',
    '501 - Ferias Mes', 'Conf 502', 'Validação 502', '502 - 1/3 Ferias Mes', '503 - Media Ferias Mes',
    'Conf 504', 'Validação 504', '504 - Media 1/3 Ferias Mes', '505 - Abono Pecuniario Mes', 'Conf 506',
    'Validação 506', '506 - 1/3 Abono Pecuniario Mes', '507 - Media Abono Pecuniario Mes', '508 - Media 1/3 Abono Pecuniario Mes',
    '509 - Ferias Mes Seguinte', 'Conf 510', 'Validação 510', '510 - 1/3 Ferias Mes Seguinte', '511 - Media Ferias Mes Seguinte',
    '512 - Media 1/3 Ferias Mes Seguinte', '513 - Abono Pecuniario Mes Seguinte', 'Conf 514', 'Validação 514',
    '514 - 1/3 Abono Pecuniario Mes Seguinte', '515 - Media Abono Pecuniario Mes Seguinte', '516 - Media 1/3 Abono Pecuniario Mes Seguinte',
    '517 - Adiantamento 13 Salario Ferias', '518 - Ferias Mes Seguinte 2', '519 - 1/3 Ferias Mes Seguinte 2',
    '520 - Media Ferias Mes Seguinte 2', '521 - Media 1/3 Ferias Mes Seguinte 2', '522 - Ferias em Dobro',
    '526 - Abono Pecuniario em Dobro', 'Confere 545', 'Validação 545', '545 - Abono Ferias Farmacia',
    '546 - Gratificacao Ferias', 'Conf 5501', 'Validação 5501', '5501 - Inss Ferias', '5502 - Irrf Ferias Normais',
    'Pensionista', '5503 - Pensao Alimenticia Ferias I', '5504 - Pensao Alimenticia Ferias II', '5505 - Pensao Alimenticia Ferias III',
    '10120 - Base Inss Ferias', '10205 - Base Bruta Irrf Ferias Normais', '10225 - Base Liquida Irrf Ferias Normais',
    '20215 - Num Dias Ferias Mes I', '20220 - Num Dias Ferias Mes Seguinte I', '20280 - Num Dias Ab Pec Total I',
    '20285 - Num Dias Ab Pec Total II', '20315 - Num Dias Ferias Mes II', '20374 - Num Dias Dobro Ferias',
    '20415 - Num Dias Ferias Mes III', '20420 - Num Dias Ferias Mes Seguinte III', 'TT dias Férias', 'Justificativa'
]

OFFICIAL_HEADERS_FOLHA = [
    'Código da Empresa', 'Nome da Empresa', 'Código Estabelecimento', 'Estabelecimento',
    'Código Centro Custo', 'Centro Custo', 'Matrícula', 'Nome do Colaborador', 'ID Colaborador',
    'Data Admissão', 'Código Cargo', 'Cargo', 'Situação', 'Data Rescisão',
    'Férias Agosto', 'Férias Setembro', 'Validação Entre Fração e Programação de Gozo',
    '1 - Salario', '2 - Minimo Garantido Comissionista', '4 - Adicional de Periculosidade',
    '5 - Adicional de Insalubridade', '7 - Horas Trabalhadas', '8 - DSR Horas Trabalhadas',
    '11 - Anuenio', '12 - Bienio', '13 - Trienio', '14 - Quadrienio', '15 - Quinquenio',
    '16 - Sexenio', '17 - Septenio', '18 - Salario Maternidade - Lei N 10.710',
    '137 - Auxilio Doenca Previdenciario', '138 - Auxilio Acid de Trabalho Previdenciario',
    '160 - Octenio', '161 - Decenio', '163 - Adicional de Gerencia', '164 - Responsabilidade Tecnica',
    '165 - Adicional de Titulacao', '81 Comissões', '90 DSR Comissões', '240 Garantia Estendida',
    '241 - Comissoes Pura', '300 - Ferias 1 Ocorrencia', '325 - Compl Ferias Recalculo 1 Ocorrencia',
    '333 - Compl Ferias Recalculo 2 Ocorrencia', '302 - Media Ferias 1 Ocorrencia',
    '350 - Ferias em Dobro 1 Ocorrencia', '351 - 1/3 Ferias em Dobro 1 Ocorrencia',
    '352 - Media Ferias em Dobro 1 Ocorrencia', '353 - 1/3 Media Ferias em Dobro 1 Ocorrencia',
    '354 - Ferias em Dobro 2 Ocorrencia', '355 - 1/3 Ferias em Dobro 2 Ocorrencia',
    '356 - Media Ferias em Dobro 2 Ocorrencia', '522 - Ferias em Dobro', '802 - Ferias em Dobro Adicional 1/3',
    '5026 - Irrf Ferias 1 Oc - Recalculo', '11220 - Remuneracao Fixa Mensal', '20215 - Num Dias Ferias Mes I',
    '20235 - Num Dias Trabalhados Mes', '20292 - Num Dias Afastamento Mes', 'TT dias', 'TT de Proventos',
    'Confere vendedor', 'Validação', 'Obs.:'
]

# Estilos Visuais
header_font_white = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
header_font_black = Font(name="Calibri", size=11, bold=True, color="000000")
header_fill_blue = PatternFill("solid", fgColor="002060") # Azul Institucional Cencosud
header_fill_yellow = PatternFill("solid", fgColor="FFC000") # Amarelo Ouro (Validações)
alert_fill_red = PatternFill("solid", fgColor="FFC7CE") # Fundo Alerta Vermelho
alert_font_red = Font(name="Calibri", size=11, color="9C0006", bold=True)
yellow_alert_fill = PatternFill("solid", fgColor="FFFF99")

center_align = Alignment(horizontal="center", vertical="center")
left_align = Alignment(horizontal="left", vertical="center")

def safe_round(val):
    if val is None: return None
    try:
        fval = float(val)
        if abs(fval) < 0.001: return 0.0
        return round(fval, 2)
    except:
        return val

def parse_num(v):
    if v is None: return None
    s = str(v).strip()
    if s in ['', '-', ' -', 'nan', 'None']: return None
    if isinstance(v, (int, float)): return float(v)
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        f = float(s)
        return f if f != 0.0 else None
    except:
        return None

def is_code_col(h):
    hl = str(h).lower().strip()
    codes = [
        'código da empresa', 'codigo da empresa', 'código empresa', 'codigo empresa',
        'código estabelecimento', 'codigo estabelecimento',
        'código centro custo', 'codigo centro custo',
        'matrícula', 'matricula',
        'id colaborador',
        'código cargo', 'codigo cargo'
    ]
    return any(c in hl for c in codes)

def is_text_col(h):
    hl = str(h).lower().strip()
    texts = [
        'nome da empresa', 'empresa', 'estabelecimento', 'centro de custo', 'centro custo',
        'nome do colaborador', 'nome colaborador', 'colaborador', 'cargo',
        'situação', 'situacao', 'sindicato', 'justificativa', 'obs', 'tempo de casa', 'confere', 'pa', 'validar vendedor', 'pensionista'
    ]
    return any(t in hl for t in texts) and not is_code_col(h)

def is_date_col(h):
    hl = str(h).lower().strip()
    dates = [
        'data admissão', 'data admissao', 'admissão', 'admissao',
        'data rescisão', 'data rescisao', 'rescisão', 'rescisao'
    ]
    return any(d in hl for d in dates)

def is_day_col(h):
    hl = str(h).lower().strip()
    return 'num dias' in hl or hl == 'tt dias férias' or hl == 'tt dias ferias' or hl == 'tt dias'

def is_conf_val_col(h):
    hl = str(h).lower().strip()
    return ('conf' in hl or 'valida' in hl) and not is_day_col(h)

# Algoritmo Oficial de INSS Progressivo 2025/2026
def calc_inss_2026_official(b):
    if b <= 0: return 0.0
    if b >= 8475.55: return 988.0914
    if b <= 1621.00:
        return b * 0.075
    elif b <= 2902.84:
        return b * 0.09 - 24.315
    elif b <= 4354.27:
        return b * 0.12 - 111.4002
    else:
        return b * 0.14 - 198.4856

def find_files_by_prefix(directory, prefix):
    if not os.path.exists(directory):
        return []
    prefix_norm = ''.join(c for c in unicodedata.normalize('NFD', prefix) if unicodedata.category(c) != 'Mn').lower()
    matches = []
    for fname in os.listdir(directory):
        if not fname.lower().endswith('.csv'):
            continue
        fname_norm = ''.join(c for c in unicodedata.normalize('NFD', fname) if unicodedata.category(c) != 'Mn').lower()
        if fname_norm.startswith(prefix_norm):
            matches.append(os.path.join(directory, fname))
    return matches

def find_file_by_prefix(directory, prefix, fallback_name=None):
    matches = find_files_by_prefix(directory, prefix)
    if matches:
        return max(matches, key=os.path.getmtime)
    if fallback_name:
        p = os.path.join(directory, fallback_name)
        if os.path.exists(p):
            return p
    return None

def safe_save_workbook(wb, target_path, prompt_lock_cb=None, log_cb=None):
    """Salva o workbook tratando bloqueios do Excel (PermissionError)."""
    os.makedirs(os.path.dirname(os.path.abspath(target_path)), exist_ok=True)
    tmp_path = target_path + f".tmp_{os.getpid()}_{int(time.time())}.xlsx"
    wb.save(tmp_path)

    while True:
        try:
            if os.path.exists(target_path):
                with open(target_path, 'r+b'):
                    pass
                os.remove(target_path)
            os.rename(tmp_path, target_path)
            if log_cb:
                log_cb(f"Arquivo gravado com sucesso: {os.path.basename(target_path)}")
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
                raise PermissionError(f"Arquivo {target_path} bloqueado no Excel.")

def carregar_contexto_compartilhado(planilhas_dir=None, output_dir=None, file_map=None, log_cb=None, progress_cb=None):
    """Carrega as bases mestres compartilhadas entre os motores de Folha e Cálculo."""
    if not planilhas_dir:
        planilhas_dir = DEFAULT_PLANILHAS_DIR
    if not output_dir:
        output_dir = DEFAULT_OUTPUT_DIR

    shared_dir = os.path.join(planilhas_dir, "bases compartilhadas")
    calc_dir = os.path.join(planilhas_dir, "relatorios calculo ferias")
    folha_dir = os.path.join(planilhas_dir, "relatorios folha mensal")
    conf_dir = os.path.join(planilhas_dir, "relatorios conferencia ferias formatada")

    def log(msg):
        print(msg)
        if log_cb:
            log_cb(msg)

    # Identificação de arquivos
    file_colab_master = (file_map.get('colab_master') if file_map else None) or find_file_by_prefix(shared_dir, "COLA_Colaboradores_CSV", "COLA_Colaboradores_CSV - 10092026_183455_824810.CSV")
    file_fer_hist = (file_map.get('ferias_hist') if file_map else None) or find_file_by_prefix(shared_dir, "Cencosud_Ferias_Colaboradores_CSV", "Cencosud_Ferias_Colaboradores_CSV 01 - 10092026_181158_492867.CSV")
    file_pens = (file_map.get('pensionistas') if file_map else None) or find_file_by_prefix(calc_dir, "COLA_Beneficiario_Pensionista", "COLA_Beneficiario_Pensionista - 23092026_092746_655088.CSV")
    file_afas = (file_map.get('afastados') if file_map else None) or find_file_by_prefix(folha_dir, "COLA_Colaboradores_Afastados", "COLA_Colaboradores_Afastados - 15092026_163904_199734 (1).CSV")
    file_fopa_folha = (file_map.get('fopa_folha') if file_map else None) or find_file_by_prefix(folha_dir, "FOPA_Eventos_por_Colaborador", "FOPA_Eventos_por_Colaborador - 22092026_084440_677704 (1).CSV")
    file_fopa_calc = (file_map.get('fopa_calc') if file_map else None) or find_file_by_prefix(calc_dir, "FOPA_Eventos_por_Colaborador", "FOPA_Eventos_por_Colaborador - 22092026_083606_981325 (1).CSV")

    log("Carregando base mestre de colaboradores (COLA_Colaboradores)...")
    df_colab_m = safe_read_csv(file_colab_master, sep=';', encoding='latin1', skiprows=1, dtype=str)
    if not any('matr' in str(c).lower() for c in df_colab_m.columns):
        df_colab_m = safe_read_csv(file_colab_master, sep=';', encoding='latin1', dtype=str)
    mat_col_m = [c for c in df_colab_m.columns if 'Matr' in c and 'eSocial' not in c][0]
    df_colab_m[mat_col_m] = df_colab_m[mat_col_m].astype(str).str.replace('\x1d', '', regex=False).str.strip()

    dict_sindicato = {}
    dict_admissao = {}
    dict_sal_contratual = {}
    dict_cargo_mestre = {}

    col_sind = [c for c in df_colab_m.columns if 'Sindicato' in c]
    col_sind_name = col_sind[0] if col_sind else None
    col_adm = [c for c in df_colab_m.columns if 'Admiss' in c][0]
    col_sal = [c for c in df_colab_m.columns if 'Salário Contratual' in c or 'Salario Contratual' in c or 'Salário' in c]
    col_sal_name = col_sal[0] if col_sal else None
    col_cargo_m = [c for c in df_colab_m.columns if str(c).strip() == 'Cargo']
    col_cargo_name = col_cargo_m[0] if col_cargo_m else [c for c in df_colab_m.columns if 'Nome Cargo' in c or 'Cargo' in c][0]

    for _, r in df_colab_m.iterrows():
        m = r[mat_col_m]
        if col_sind_name: dict_sindicato[m] = str(r[col_sind_name]).strip()
        dict_admissao[m] = str(r[col_adm]).strip()
        if col_sal_name: dict_sal_contratual[m] = parse_num(r[col_sal_name])
        dict_cargo_mestre[m] = str(r[col_cargo_name]).strip().upper()

    log(f"Base mestre carregada: {len(df_colab_m)} colaboradores mapeados.")

    # Arquivos de Férias Programadas
    prog_files = []
    if file_map and file_map.get('prog_files'):
        prog_files = [p for p in file_map.get('prog_files') if os.path.exists(p)]
    if not prog_files:
        prog_files_shared = find_files_by_prefix(shared_dir, "Ferias Programadas ou Gozadas em CSV")
        prog_files_conf = find_files_by_prefix(conf_dir, "Ferias Programadas ou Gozadas em CSV")
        prog_files = sorted(set(prog_files_shared + prog_files_conf), key=os.path.getmtime)

    def load_ferias(f_path):
        m_fer, d_dt = {}, {}
        if os.path.exists(f_path):
            df_p = safe_read_csv(f_path, sep=';', encoding='latin1', dtype=str)
            if 'Matr' not in df_p.columns[4]:
                df_p = safe_read_csv(f_path, sep=';', encoding='latin1', skiprows=1, dtype=str)
            for _, r in df_p.iterrows():
                m = str(r.iloc[4]).strip()
                val_g = str(r.iloc[14]).strip().replace('  a  ', ' a ')
                val_d = str(r.iloc[15]).strip()
                m_fer[m] = f"{val_g} ({val_d}d)"
                if ' a ' in val_g:
                    p = val_g.split(' a ')
                    di = pd.to_datetime(p[0].strip(), format='%d/%m/%Y', errors='coerce')
                    dfim = pd.to_datetime(p[1].strip(), format='%d/%m/%Y', errors='coerce')
                    d_dt[m] = (di, dfim)
        return m_fer, d_dt

    # Detecção dinâmica de competência dos FOPAs
    mes_folha, ano_folha = extract_fopa_competencia(file_fopa_folha)
    if not mes_folha:
        mes_folha, ano_folha = 9, 2026
    mes_calc, ano_calc = extract_fopa_competencia(file_fopa_calc)
    if not mes_calc:
        mes_calc, ano_calc = 10, 2026

    map_ferias_by_month = {}
    dt_by_month = {}

    for f_path in prog_files:
        m_fer, d_dt = load_ferias(f_path)
        if not d_dt:
            continue
        meses = [dt[0].month for dt in d_dt.values() if pd.notna(dt[0])]
        if not meses:
            continue
        mes_predominante = max(set(meses), key=meses.count)
        map_ferias_by_month[mes_predominante] = m_fer
        dt_by_month[mes_predominante] = d_dt

    map_ferias_ago = map_ferias_by_month.get(8, {})
    map_ferias_set = map_ferias_by_month.get(9, {})
    map_ferias_out = map_ferias_by_month.get(10, {})

    mes_ant = 12 if mes_folha == 1 else mes_folha - 1
    ano_ant = ano_folha - 1 if mes_folha == 1 else ano_folha

    dt_ant = dt_by_month.get(mes_ant, {})
    dt_atual = dt_by_month.get(mes_folha, {})

    map_validar_frac = {}
    for mat_ in set(dt_ant.keys()).intersection(set(dt_atual.keys())):
        fim_a = dt_ant[mat_][1]
        ini_s = dt_atual[mat_][0]
        if pd.notna(fim_a) and pd.notna(ini_s):
            diff = (ini_s - fim_a).days
            if ini_s.month == fim_a.month:
                map_validar_frac[mat_] = "Retorno e novo gozo no mesmo mês"
            elif diff < 30:
                map_validar_frac[mat_] = f"Intervalo < 30 dias ({diff}d)"
            else:
                map_validar_frac[mat_] = "OK"

    # Afastados (com filtro estrito de Status == 'Ativo')
    set_afastados = set()
    if file_afas and os.path.exists(file_afas):
        df_afas = safe_read_csv(file_afas, sep=';', encoding='latin1', dtype=str)
        if not any('matr' in str(c).lower() for c in df_afas.columns):
            df_afas = safe_read_csv(file_afas, sep=';', encoding='latin1', skiprows=1, dtype=str)
        col_mat_afas = [c for c in df_afas.columns if 'Matr' in c or 'matr' in str(c).lower()][0]
        col_status_afas = [c for c in df_afas.columns if any(k in str(c).lower() for k in ['status', 'situac'])]
        if col_status_afas:
            col_st = col_status_afas[0]
            df_afas = df_afas[df_afas[col_st].astype(str).str.strip().str.lower() == 'ativo']
        set_afastados = set(df_afas[col_mat_afas].astype(str).str.replace('\x1d', '', regex=False).str.strip())
        log(f"Afastados ativos mapeados: {len(set_afastados)}")

    # Pensionistas
    dict_pens = {}
    if file_pens and os.path.exists(file_pens):
        df_pens = safe_read_csv(file_pens, sep=';', encoding='latin1', dtype=str)
        for _, row in df_pens.iterrows():
            mat = str(row.iloc[2]).strip()
            sit = str(row.iloc[5]).strip()
            if mat not in dict_pens or sit.lower() == 'ativo':
                dict_pens[mat] = sit

    # PA
    dict_pa = {}
    def carregar_pa(f_path, destino):
        df_prog = safe_read_csv(f_path, sep=';', encoding='latin1', dtype=str)
        if 'Matr' not in str(df_prog.columns[4]):
            df_prog = safe_read_csv(f_path, sep=';', encoding='latin1', skiprows=1, dtype=str)
        for _, r in df_prog.iterrows():
            m = str(r.iloc[4]).strip()
            pa_val = str(r.iloc[12]).strip()
            if m and m != 'nan' and pa_val and pa_val != 'nan':
                destino[m] = pa_val

    for f_path in prog_files:
        if os.path.exists(f_path):
            carregar_pa(f_path, dict_pa)
    log(f"PA carregado para {len(dict_pa)} matrículas ({len(prog_files)} arquivos de férias programadas detectados).")

    return {
        'dict_sindicato': dict_sindicato,
        'dict_admissao': dict_admissao,
        'dict_sal_contratual': dict_sal_contratual,
        'dict_cargo_mestre': dict_cargo_mestre,
        'mes_folha': mes_folha,
        'ano_folha': ano_folha,
        'mes_calc': mes_calc,
        'ano_calc': ano_calc,
        'map_ferias_by_month': map_ferias_by_month,
        'dt_by_month': dt_by_month,
        'map_validar_frac': map_validar_frac,
        'map_ferias_ago': map_ferias_ago,
        'map_ferias_set': map_ferias_set,
        'map_ferias_out': map_ferias_out,
        'set_afastados': set_afastados,
        'dict_pens': dict_pens,
        'dict_pa': dict_pa,
        'file_fopa_folha': file_fopa_folha,
        'file_fopa_calc': file_fopa_calc,
        'output_dir': output_dir,
        'planilhas_dir': planilhas_dir
    }

def run_auditoria_folha(planilhas_dir=None, output_dir=None, file_map=None, progress_cb=None, log_cb=None, prompt_lock_cb=None):
    """Executa a Auditoria Oficial da Folha Mensal (Set/2026) - 65 Colunas homologadas."""
    def log(msg):
        print(msg)
        if log_cb:
            log_cb(msg)

    def progress(pct, msg):
        log(f"[{int(pct*100)}%] {msg}")
        if progress_cb:
            progress_cb(pct, msg)

    progress(0.05, "Inicializando motor da Folha Mensal...")
    if not planilhas_dir: planilhas_dir = DEFAULT_PLANILHAS_DIR
    if not output_dir: output_dir = DEFAULT_OUTPUT_DIR

    ctx = carregar_contexto_compartilhado(planilhas_dir, output_dir, file_map, log_cb=log, progress_cb=progress_cb)
    file_fopa_folha = ctx['file_fopa_folha']

    progress(0.20, "Identificando cabeçalhos oficiais da Folha...")
    headers_folha = None
    validadas_dir = os.path.join(output_dir, "ja validadas")
    candidates = [
        file_map.get('ref_folha') if file_map else None,
        os.path.join(validadas_dir, "Auditoria_Folha_Mensal_Set2026_v6.xlsx"),
        os.path.join(output_dir, "Auditoria_Folha_Mensal_Set2026_v6.xlsx")
    ]
    if os.path.exists(validadas_dir):
        for f in os.listdir(validadas_dir):
            if 'folha_mensal' in f.lower() and f.endswith('.xlsx') and not f.startswith('~$'):
                candidates.append(os.path.join(validadas_dir, f))
    for cand in candidates:
        if cand and os.path.exists(cand):
            try:
                tmp_ref = cand + f".tmp_{os.getpid()}.xlsx"
                shutil.copy2(cand, tmp_ref)
                wb_rf = openpyxl.load_workbook(tmp_ref, data_only=False)
                ws_rf = wb_rf.worksheets[0]
                loaded = [str(ws_rf.cell(row=1, column=c).value).strip() for c in range(1, ws_rf.max_column + 1)]
                wb_rf.close()
                if os.path.exists(tmp_ref): os.remove(tmp_ref)
                if len(loaded) >= 65:
                    headers_folha = loaded
                    break
            except Exception:
                pass

    if not headers_folha:
        headers_folha = list(OFFICIAL_HEADERS_FOLHA)

    # Competência dinâmica da Folha
    mes_folha = ctx.get('mes_folha') or 9
    ano_folha = ctx.get('ano_folha') or 2026
    mes_ant = 12 if mes_folha == 1 else mes_folha - 1
    ano_ant = ano_folha - 1 if mes_folha == 1 else ano_folha
    nome_mes_ant = MESES_PT.get(mes_ant, 'Agosto')
    nome_mes_atual = MESES_PT.get(mes_folha, 'Setembro')
    abrev_mes_atual = MESES_ABREV.get(mes_folha, 'Set')

    col_fer_ant = f"Férias {nome_mes_ant}"
    col_fer_atual = f"Férias {nome_mes_atual}"

    # Ajusta nomes das colunas de férias conforme a competência
    idx_val_frac = next((i for i, h in enumerate(headers_folha) if 'valida' in str(h).lower() and ('gozo' in str(h).lower() or 'fra' in str(h).lower())), 16)
    if idx_val_frac >= 2:
        headers_folha[idx_val_frac - 2] = col_fer_ant
        headers_folha[idx_val_frac - 1] = col_fer_atual

    log(f"Competência Folha Mensal: {nome_mes_atual}/{ano_folha} (Colunas: {col_fer_ant} e {col_fer_atual})")
    log(f"Total de colunas Folha Mensal: {len(headers_folha)}")

    progress(0.30, "Lendo CSV da Folha Mensal...")
    df_fopa_folha = safe_read_csv(file_fopa_folha, sep=';', encoding='latin1', skiprows=3, dtype=str)
    if len(df_fopa_folha.columns) < 30:
        log(f"AVISO: O arquivo FOPA selecionado ({os.path.basename(file_fopa_folha)}) possui apenas {len(df_fopa_folha.columns)} colunas. Buscando base completa de proventos...")
        cand_dir = os.path.dirname(file_fopa_folha)
        candidates_f = find_files_by_prefix(cand_dir, "FOPA_Eventos_por_Colaborador")
        for cand in candidates_f:
            if cand != file_fopa_folha:
                try:
                    df_test = safe_read_csv(cand, sep=';', encoding='latin1', skiprows=3, nrows=5, dtype=str)
                    if len(df_test.columns) >= 30 and any('salario' in c.lower() or '11220' in c for c in df_test.columns):
                        log(f"FOPA completo identificado automaticamente: {os.path.basename(cand)}")
                        file_fopa_folha = cand
                        df_fopa_folha = safe_read_csv(file_fopa_folha, sep=';', encoding='latin1', skiprows=3, dtype=str)
                        break
                except Exception:
                    pass
    mat_col_ff = [c for c in df_fopa_folha.columns if 'Matr' in c][0]
    df_fopa_folha[mat_col_ff] = df_fopa_folha[mat_col_ff].str.strip()
    total_colab = len(df_fopa_folha)
    log(f"Total colaboradores Folha Mensal: {total_colab}")

    wb_out_f = openpyxl.Workbook()
    ws_out_f = wb_out_f.active
    ws_out_f.title = f"Auditoria Folha {abrev_mes_atual} {ano_folha}"

    # Cabeçalhos
    for col_idx, h in enumerate(headers_folha, 1):
        cell = ws_out_f.cell(row=1, column=col_idx, value=h)
        if is_conf_val_col(h):
            cell.fill = header_fill_yellow
            cell.font = header_font_black
        else:
            cell.fill = header_fill_blue
            cell.font = header_font_white
        cell.alignment = center_align

    progress(0.40, "Calculando e preenchendo colaboradores da Folha...")
    step_notify = max(1, total_colab // 5)

    for idx, row in df_fopa_folha.iterrows():
        out_row = idx + 2
        mat = str(row[mat_col_ff]).strip()

        if idx > 0 and idx % step_notify == 0:
            pct = 0.40 + (0.45 * (idx / total_colab))
            progress(pct, f"Processando colaborador {idx:,} de {total_colab:,}...")

        col_vals = {}
        for h in headers_folha:
            hl = str(h).lower().strip()
            raw_val = None
            for fc in df_fopa_folha.columns:
                fcl = fc.lower().strip()
                if fcl == hl:
                    raw_val = row[fc]
                    break
            if raw_val is None:
                for fc in df_fopa_folha.columns:
                    fcl = fc.lower().strip()
                    if hl in fcl:
                        if hl in ['cargo', 'estabelecimento', 'centro de custo', 'empresa', 'nome da empresa', 'nome do colaborador']:
                            if 'código' not in fcl and 'codigo' not in fcl and 'id ' not in fcl and 'matr' not in fcl:
                                raw_val = row[fc]
                                break
                        else:
                            if fcl.startswith(hl + " ") or fcl == hl:
                                raw_val = row[fc]
                                break

            val = None
            if is_code_col(h):
                s = str(raw_val).strip() if raw_val is not None else ""
                if s and s not in ['nan', 'None', '-']:
                    if '.' in s: s = s.split('.')[0]
                    val = s
            elif is_text_col(h):
                s = str(raw_val).strip() if raw_val is not None else ""
                val = s if s and s not in ['nan', 'None', '-'] else None
            elif is_date_col(h):
                s = str(raw_val).strip() if raw_val is not None else ""
                if s and s not in ['nan', 'None', '-']:
                    dt = pd.to_datetime(s, format='%d/%m/%Y', dayfirst=True, errors='coerce')
                    val = dt.strftime('%d/%m/%Y') if pd.notna(dt) else s
            else:
                num = parse_num(raw_val)
                if num is not None and num != 0.0:
                    val = float(num)

            col_vals[h] = val

        # Regras de Negócio e Cálculos
        col_vals[col_fer_ant] = ctx['map_ferias_by_month'].get(mes_ant, {}).get(mat, None)
        col_vals[col_fer_atual] = ctx['map_ferias_by_month'].get(mes_folha, {}).get(mat, None)

        val_frac_col = None
        for h in headers_folha:
            if headers_folha.index(h) == 16 or ('valida' in h.lower() and ('gozo' in h.lower() or 'fra' in h.lower())):
                val_frac_col = h
                break
        if val_frac_col:
            col_vals[val_frac_col] = ctx.get('map_validar_frac', {}).get(mat, 'OK')

        # TT dias
        v1 = parse_num(next((v for k,v in row.items() if '20215' in str(k)), None)) or 0.0
        v2 = parse_num(next((v for k,v in row.items() if '20235' in str(k)), None)) or 0.0
        v3 = parse_num(next((v for k,v in row.items() if '20292' in str(k)), None)) or 0.0
        col_vals['TT dias'] = v1 + v2 + v3

        # TT de Proventos (Cols R a AS, indices 17 a 44)
        try:
            tt_prov = sum(parse_num(col_vals[headers_folha[i]]) or 0.0 for i in range(17, 45) if i < len(headers_folha))
        except: tt_prov = 0.0
        col_vals['TT de Proventos'] = tt_prov

        # TT Descontos (Cols AU a BD, indices 46 a 55)
        try:
            tt_desc = sum(parse_num(col_vals[headers_folha[i]]) or 0.0 for i in range(46, 56) if i < len(headers_folha))
        except: tt_desc = 0.0
        col_vals['TT Descontos'] = tt_desc

        col_vals['Líquido'] = tt_prov - tt_desc

        val_liq = parse_num(row.get('11200 - Liquido Folha', None))
        col_vals['Liquido Folha'] = float(val_liq) if val_liq is not None else None

        val_col_name = next((h for h in headers_folha if 'valida' in h.lower() and len(h) < 15), 'Validação')
        col_vals[val_col_name] = safe_round(tt_prov - (parse_num(next((v for k,v in row.items() if '11220' in str(k)), None)) or 0.0))

        # Confere vendedor
        cargo_val = str(col_vals.get('Cargo') or '').strip().upper()
        if cargo_val == 'VENDEDOR':
            soma_prov = 0.0
            for i in range(17, min(44, len(headers_folha))):
                c_prov = headers_folha[i]
                v_prov = parse_num(col_vals.get(c_prov, None))
                if v_prov: soma_prov += v_prov
            col_vals['Confere vendedor'] = float(soma_prov)
        else:
            col_vals['Confere vendedor'] = None

        # Obs.:
        v_afas = parse_num(row.get('20292 - Num Dias Afastamento Mes', None)) or 0.0
        obs_msgs = []
        if 'VENDEDOR' in cargo_val:
            obs_msgs.append("Vendedor")
        val_folha = col_vals.get(val_col_name, 0.0)
        if val_folha is not None and float(val_folha) > 0.02:
            obs_msgs.append("Diferença de valor Positiva - Conferir")
        if mat in ctx['set_afastados'] or v_afas > 0:
            obs_msgs.append("Afastamento INSS")
        col_vals['Obs.:'] = '; '.join(obs_msgs) if obs_msgs else None

        # Escrita nas células
        for col_idx, h in enumerate(headers_folha, 1):
            cell = ws_out_f.cell(row=out_row, column=col_idx)
            val = col_vals.get(h, None)
            cell.value = val

            if val is not None:
                is_red_alert = (col_vals.get(val_frac_col) == 'Retorno e novo gozo no mesmo mês' and (h.startswith('Férias') or h == val_frac_col))
                if is_red_alert:
                    cell.fill = alert_fill_red
                    cell.font = Font(name="Calibri", size=11, bold=True, color='9C0006')
                elif h == val_frac_col and val != 'OK':
                    cell.fill = yellow_alert_fill
                    cell.font = Font(name="Calibri", size=11, bold=True)

                if is_code_col(h): cell.number_format = '@'
                elif is_text_col(h): pass
                elif is_date_col(h): pass
                elif isinstance(val, (int, float)):
                    if h == 'TT dias': cell.number_format = '0'
                    else: cell.number_format = '#,##0.00;-#,##0.00;0.00'

            hl = h.lower().strip()
            cell.alignment = left_align if any(k in hl for k in ['nome', 'empresa', 'estabelecimento', 'cargo', 'centro']) else center_align

    progress(0.90, "Formatando larguras e congelando painéis...")
    fixed_widths_f = [5.5, 15, 6, 10, 12, 10, 8, 40, 7, 11, 6.5, 45, 23, 14, 25, 25, 32]
    for col_idx in range(1, len(headers_folha) + 1):
        col_letter = get_column_letter(col_idx)
        if col_idx <= len(fixed_widths_f):
            ws_out_f.column_dimensions[col_letter].width = fixed_widths_f[col_idx - 1]
        else:
            h_text = headers_folha[col_idx - 1]
            ws_out_f.column_dimensions[col_letter].width = max(len(str(h_text)) + 3, 10)

    ws_out_f.freeze_panes = "A2"

    target_folha = os.path.join(output_dir, f"Auditoria_Folha_Mensal_{abrev_mes_atual}{ano_folha}_v6.xlsx")
    saved_path = safe_save_workbook(wb_out_f, target_folha, prompt_lock_cb=prompt_lock_cb, log_cb=log)
    progress(1.0, f"Sucesso! {os.path.basename(saved_path)} gerada com 100% de integridade!")
    return saved_path

def run_auditoria_calculo(planilhas_dir=None, output_dir=None, file_map=None, progress_cb=None, log_cb=None, prompt_lock_cb=None):
    """Executa a Auditoria Oficial do Cálculo de Férias (Out/2026) - 86 Colunas homologadas."""
    def log(msg):
        print(msg)
        if log_cb:
            log_cb(msg)

    def progress(pct, msg):
        log(f"[{int(pct*100)}%] {msg}")
        if progress_cb:
            progress_cb(pct, msg)

    progress(0.05, "Inicializando motor do Cálculo de Férias...")
    if not planilhas_dir: planilhas_dir = DEFAULT_PLANILHAS_DIR
    if not output_dir: output_dir = DEFAULT_OUTPUT_DIR

    ctx = carregar_contexto_compartilhado(planilhas_dir, output_dir, file_map, log_cb=log, progress_cb=progress_cb)
    file_fopa_folha = ctx['file_fopa_folha']
    file_fopa_calc = ctx['file_fopa_calc']

    progress(0.20, "Identificando cabeçalhos oficiais do Cálculo...")
    headers_calc = None
    validadas_dir = os.path.join(output_dir, "ja validadas")
    candidates = [
        file_map.get('ref_calc') if file_map else None,
        os.path.join(validadas_dir, "Auditoria_Calculo_Ferias_Out2026_v6.xlsx"),
        os.path.join(output_dir, "Auditoria_Calculo_Ferias_Out2026_v6.xlsx")
    ]
    if os.path.exists(validadas_dir):
        for f in os.listdir(validadas_dir):
            if 'calculo_ferias' in f.lower() and f.endswith('.xlsx') and not f.startswith('~$'):
                candidates.append(os.path.join(validadas_dir, f))
    for cand in candidates:
        if cand and os.path.exists(cand):
            try:
                tmp_ref = cand + f".tmp_{os.getpid()}.xlsx"
                shutil.copy2(cand, tmp_ref)
                wb_rc = openpyxl.load_workbook(tmp_ref, data_only=False)
                ws_rc = wb_rc.worksheets[0]
                loaded = [str(ws_rc.cell(row=1, column=c).value).strip() for c in range(1, ws_rc.max_column + 1)]
                wb_rc.close()
                if os.path.exists(tmp_ref): os.remove(tmp_ref)
                if len(loaded) >= 86:
                    headers_calc = loaded
                    break
            except Exception:
                pass

    if not headers_calc:
        headers_calc = list(OFFICIAL_HEADERS_CALC)

    # Competência dinâmica do Cálculo
    mes_calc = ctx.get('mes_calc') or 10
    ano_calc = ctx.get('ano_calc') or 2026
    nome_mes_calc = MESES_PT.get(mes_calc, 'Outubro')
    abrev_mes_calc = MESES_ABREV.get(mes_calc, 'Out')
    col_fer_calc = f"Férias {nome_mes_calc}"

    for idx_c, h_c in enumerate(headers_calc):
        if str(h_c).strip().startswith('Férias') or str(h_c).strip().startswith('Ferias'):
            headers_calc[idx_c] = col_fer_calc
            break

    log(f"Competência Cálculo de Férias: {nome_mes_calc}/{ano_calc} (Coluna: {col_fer_calc})")
    log(f"Total de colunas Cálculo de Férias: {len(headers_calc)}")

    progress(0.30, "Mapeando adicionais de Insalubridade e Periculosidade...")
    df_fopa_folha = safe_read_csv(file_fopa_folha, sep=';', encoding='latin1', skiprows=3, dtype=str)
    mat_col_ff = [c for c in df_fopa_folha.columns if 'Matr' in c][0]
    cols_insalub = [c for c in df_fopa_folha.columns if c.startswith('4 -') or c.startswith('5 -') or 'adicional de insalubridade' in c.lower() or 'adicional de periculosidade' in c.lower()]
    dict_insalub = {}
    for _, row in df_fopa_folha.iterrows():
        mat = str(row[mat_col_ff]).strip()
        total = sum(parse_num(row[col]) or 0.0 for col in cols_insalub)
        if total > 0: dict_insalub[mat] = total

    progress(0.40, "Lendo CSV do Cálculo de Férias...")
    df_fopa_calc = safe_read_csv(file_fopa_calc, sep=';', encoding='latin1', skiprows=3, dtype=str)
    mat_col_calc = [c for c in df_fopa_calc.columns if 'Matr' in c][0]
    df_fopa_calc[mat_col_calc] = df_fopa_calc[mat_col_calc].str.strip()
    total_colab = len(df_fopa_calc)
    log(f"Total colaboradores Cálculo: {total_colab}")

    wb_out_c = openpyxl.Workbook()
    ws_out_c = wb_out_c.active
    ws_out_c.title = "Conferencia Calculo Ferias"

    for col_idx, h in enumerate(headers_calc, 1):
        cell = ws_out_c.cell(row=1, column=col_idx, value=h)
        if is_conf_val_col(h):
            cell.fill = header_fill_yellow
            cell.font = header_font_black
        else:
            cell.fill = header_fill_blue
            cell.font = header_font_white
        cell.alignment = center_align

    progress(0.50, "Processando linhas do Cálculo de Férias...")
    step_notify = max(1, total_colab // 5)

    for idx, row in df_fopa_calc.iterrows():
        out_row = idx + 2
        mat = str(row[mat_col_calc]).strip()
        cargo_str = str(row.get('Cargo', '')).strip().upper()
        is_vendedor = (cargo_str == 'VENDEDOR')

        sit_pens = ctx['dict_pens'].get(mat, None)
        is_pens_ativo = (sit_pens is not None and 'ativo' in sit_pens.lower())

        if idx > 0 and idx % step_notify == 0:
            pct = 0.50 + (0.35 * (idx / total_colab))
            progress(pct, f"Processando colaborador {idx:,} de {total_colab:,}...")

        col_vals = {}
        for h in headers_calc:
            hl = str(h).lower().strip()
            raw_val = None
            for fc in df_fopa_calc.columns:
                fcl = fc.lower().strip()
                if fcl == hl:
                    raw_val = row[fc]
                    break
            if raw_val is None:
                for fc in df_fopa_calc.columns:
                    fcl = fc.lower().strip()
                    if hl in fcl:
                        if hl in ['cargo', 'estabelecimento', 'centro de custo', 'empresa', 'nome da empresa', 'nome do colaborador']:
                            if 'código' not in fcl and 'codigo' not in fcl and 'id ' not in fcl and 'matr' not in fcl:
                                raw_val = row[fc]
                                break
                        else:
                            if fcl.startswith(hl + " ") or fcl == hl:
                                raw_val = row[fc]
                                break

            val = None
            if is_code_col(h):
                s = str(raw_val).strip() if raw_val is not None else ""
                if s and s not in ['nan', 'None', '-']:
                    if '.' in s: s = s.split('.')[0]
                    val = s
            elif is_text_col(h):
                s = str(raw_val).strip() if raw_val is not None else ""
                val = s if s and s not in ['nan', 'None', '-'] else None
            elif is_date_col(h):
                s = str(raw_val).strip() if raw_val is not None else ""
                if s and s not in ['nan', 'None', '-']:
                    dt = pd.to_datetime(s, format='%d/%m/%Y', dayfirst=True, errors='coerce')
                    val = dt.strftime('%d/%m/%Y') if pd.notna(dt) else s
            else:
                num = parse_num(raw_val)
                if num is not None and num != 0.0:
                    val = float(num)

            col_vals[h] = val

        # Override com lógica de dados consolidados
        col_vals['Salário Contratual'] = float(ctx['dict_sal_contratual'].get(mat)) if ctx['dict_sal_contratual'].get(mat) else col_vals.get('Salário Contratual')
        col_vals['Cargo'] = ctx['dict_cargo_mestre'].get(mat, str(row.get('Cargo', '')).strip())
        col_vals['Situação'] = str(row.get('Situação', '')).strip()
        col_vals['PA'] = ctx['dict_pa'].get(mat, None)
        col_vals[col_fer_calc] = ctx['map_ferias_by_month'].get(mes_calc, {}).get(mat, None)

        insalub = dict_insalub.get(mat, 0.0)
        col_vals['Insalubridade/Periculosidade'] = float(insalub) if insalub > 0 else None

        v_u = parse_num(next((v for k,v in row.items() if '11220' in str(k)), None)) or 0.0
        v_v = parse_num(col_vals.get('Salário Contratual')) or 0.0
        col_vals['Validação Sal Insalub/Pericul'] = v_u - v_v - insalub

        for h in headers_calc:
            if '11220 -' in h:
                v = parse_num(row.get(h, None))
                col_vals[h] = float(v) if v else None
            elif '11000 -' in h:
                v = parse_num(row.get(h, None)) or ctx['dict_sal_contratual'].get(mat, None)
                col_vals[h] = float(v) if v else None

        # Fórmulas de Validação
        v_ac = parse_num(col_vals.get('501 - Ferias Mes')) or 0.0
        v_ak = parse_num(col_vals.get('503 - Media Ferias Mes')) or 0.0
        v_aq = parse_num(col_vals.get('505 - Abono Pecuniario Mes')) or 0.0
        v_aw = parse_num(col_vals.get('509 - Ferias Mes Seguinte')) or 0.0
        v_513 = parse_num(next((v for k,v in row.items() if str(k).strip().startswith('513')), None)) or 0.0

        col_vals['Conf 501'] = v_ac + v_aq + v_aw + v_513
        col_vals['Validação 501'] = safe_round(col_vals['Conf 501'] - v_u)

        v_z = parse_num(next((v for k,v in row.items() if '11000' in str(k)), None)) or ctx['dict_sal_contratual'].get(mat, 0.0)
        col_vals['Insalubridade/Periculosidade'] = dict_insalub.get(mat, 0.0)
        col_vals['Validação Sal Insalub/Pericul'] = safe_round(v_u - v_z)

        v_by = parse_num(next((v for k,v in row.items() if '20215' in str(k)), None)) or 0.0
        v_bz = parse_num(next((v for k,v in row.items() if '20220' in str(k)), None)) or 0.0
        tt_dias_ferias = v_by + v_bz
        col_vals['TT dias Férias'] = tt_dias_ferias
        col_vals['Sindicato'] = ctx['dict_sindicato'].get(mat, None)

        col_vals['Validação x Férias Fracionadas/ Faltas'] = safe_round((v_u / 30) * tt_dias_ferias) if tt_dias_ferias else 0.0
        col_vals['Validação Fracionadas'] = safe_round((v_ac + v_aw) - col_vals['Validação x Férias Fracionadas/ Faltas'])

        if is_vendedor:
            soma_v = 0.0
            for c_name in ['501 - Ferias Mes', '503 - Media Ferias Mes', '505 - Abono Pecuniario Mes', '507 - Media Abono Pecuniario Mes', '509 - Ferias Mes Seguinte', '511 - Media Ferias Mes Seguinte']:
                v_c = parse_num(row.get(c_name, None))
                if v_c: soma_v += v_c
            col_vals['Validar Vendedor'] = float(soma_v)
            col_vals['Validar Vendedor II'] = (col_vals['Validar Vendedor'] == v_u) if col_vals.get('Validar Vendedor') is not None else False
        else:
            col_vals['Validar Vendedor'] = None
            col_vals['Validar Vendedor II'] = None

        col_vals['Conf 502'] = v_ac / 3
        col_vals['Validação 502'] = safe_round(col_vals['Conf 502'] - (parse_num(col_vals.get('502 - 1/3 Ferias Mes')) or 0.0))

        col_vals['Conf 504'] = v_ak / 3
        v_504 = parse_num(next((v for k,v in row.items() if str(k).strip().startswith('504')), None)) or 0.0
        col_vals['Validação 504'] = safe_round(col_vals['Conf 504'] - v_504)

        col_vals['Conf 506'] = v_aq / 3
        col_vals['Validação 506'] = safe_round(col_vals['Conf 506'] - (parse_num(col_vals.get('506 - 1/3 Abono Pecuniario Mes')) or 0.0))

        col_vals['Conf 510'] = v_aw / 3
        col_vals['Validação 510'] = safe_round(col_vals['Conf 510'] - (parse_num(col_vals.get('510 - 1/3 Ferias Mes Seguinte')) or 0.0))

        col_vals['Conf 514'] = v_513 / 3
        v_514 = parse_num(next((v for k,v in row.items() if str(k).strip().startswith('514')), None)) or 0.0
        col_vals['Validação 514'] = safe_round(col_vals['Conf 514'] - v_514)

        v_545 = parse_num(row.get('545 - Abono Ferias Farmacia', None))
        col_vals['Confere 545'] = float(v_545) if v_545 else None
        col_vals['Validação 545'] = safe_round((parse_num(col_vals.get('Confere 545')) or 0.0) - (parse_num(col_vals.get('545 - Abono Ferias Farmacia')) or 0.0))

        # INSS Progressivo 2026 oficial (rubrica 10120)
        base_inss = 0.0
        for fc in df_fopa_calc.columns:
            if '10120' in fc:
                num = parse_num(row[fc])
                if num: base_inss = num
        inss_val = calc_inss_2026_official(base_inss)
        col_vals['Conf 5501'] = float(inss_val)
        col_vals['Validação 5501'] = safe_round(float(inss_val) - (parse_num(col_vals.get('5501 - Inss Ferias')) or parse_num(col_vals.get('5501 - INSS Ferias')) or 0.0))

        v5503 = parse_num(row.get('5503 - Pensao Alimenticia Ferias I', None))
        v5504 = parse_num(row.get('5504 - Pensao Alimenticia Ferias II', None))
        v5505 = parse_num(row.get('5505 - Pensao Alimenticia Ferias III', None))
        tem_desc = ((v5503 and v5503 > 0) or (v5504 and v5504 > 0) or (v5505 and v5505 > 0))

        if is_pens_ativo:
            if tem_desc: col_vals['Pensionista'] = "OK"
            else: col_vals['Pensionista'] = "Pensionista Ativo sem desconto de pensão (5503/5504/5505)"
        else: col_vals['Pensionista'] = sit_pens if sit_pens else "Não"

        # Justificativa Oficial Homologada
        divergencias = []
        if is_pens_ativo and not tem_desc: divergencias.append("Pensao alimenticia nao descontada nas ferias")

        v_insalub_check = parse_num(col_vals.get('Insalubridade/Periculosidade')) or 0.0
        if v_insalub_check > 0 and v_u <= v_z:
            divergencias.append("Divergência em Validação Sal Insalub/Pericul")

        tt_dias = parse_num(col_vals.get('TT dias Férias')) or 0.0
        for v_col in ['Validação 501', 'Validação 502', 'Validação 504', 'Validação 506', 'Validação 510', 'Validação 514', 'Validação 5501', 'Validação Fracionadas']:
            v_val = col_vals.get(v_col, 0.0)

            # Ignorar Validação 501 se foram férias fracionadas
            if v_col == 'Validação 501' and tt_dias < 30:
                continue

            # Ignorar Validação Fracionadas se foram férias de 30 dias (pois a 501 já cobre isso)
            if v_col == 'Validação Fracionadas' and tt_dias == 30:
                continue

            if v_val is not None and abs(float(v_val)) > 0.02:
                divergencias.append(f"Divergência em {v_col}")

        col_vals['Justificativa'] = '; '.join(divergencias) if divergencias else None

        dt_adm = pd.to_datetime(col_vals.get('Data Admissão'), format='%d/%m/%Y', dayfirst=True, errors='coerce')
        if pd.notna(dt_adm):
            years = pd.Timestamp.today().year - dt_adm.year - ((pd.Timestamp.today().month, pd.Timestamp.today().day) < (dt_adm.month, dt_adm.day))
            col_vals['Tempo de Casa'] = f"{years} anos"

        col_vals['Confere'] = None

        # Escrita nas células
        for col_idx, h in enumerate(headers_calc, 1):
            cell = ws_out_c.cell(row=out_row, column=col_idx)
            val = col_vals.get(h, None)
            cell.value = val

            if val is not None:
                if h == 'Pensionista' and val == "Pensionista Ativo sem desconto de pensão (5503/5504/5505)":
                    cell.fill = yellow_alert_fill
                    cell.font = Font(name="Calibri", size=11, bold=True)

                if is_code_col(h): cell.number_format = '@'
                elif is_text_col(h): pass
                elif is_date_col(h): pass
                elif isinstance(val, (int, float)):
                    if h == 'TT dias Férias': cell.number_format = '0'
                    else: cell.number_format = '#,##0.00;-#,##0.00;0.00'

            hl = h.lower().strip()
            cell.alignment = left_align if any(k in hl for k in ['nome', 'empresa', 'estabelecimento', 'cargo', 'centro', 'justificativa', 'pensionista']) else center_align

    progress(0.90, "Formatando larguras e congelando painéis...")
    semantic_widths_calc = {
        'Confere': 9, 'Tempo de Casa': 13, 'Salário Contratual': 12, 'Cargo': 50,
        'Situação': 28, 'Sindicato': 8, 'PA': 23, 'Férias Outubro': 25, col_fer_calc: 25,
        'Insalubridade/Periculosidade': 22, 'Validação Sal Insalub/Pericul': 22,
        'Pensionista': 45, 'Justificativa': 40
    }

    for col_idx in range(1, len(headers_calc) + 1):
        col_letter = get_column_letter(col_idx)
        h_text = headers_calc[col_idx - 1]
        if h_text in semantic_widths_calc:
            ws_out_c.column_dimensions[col_letter].width = semantic_widths_calc[h_text]
        else:
            ws_out_c.column_dimensions[col_letter].width = max(len(str(h_text)) + 3, 11)

    ws_out_c.freeze_panes = "A2"

    target_calc = os.path.join(output_dir, f"Auditoria_Calculo_Ferias_{abrev_mes_calc}{ano_calc}_v6.xlsx")
    saved_path = safe_save_workbook(wb_out_c, target_calc, prompt_lock_cb=prompt_lock_cb, log_cb=log)
    progress(1.0, f"Sucesso! {os.path.basename(saved_path)} gerada com 100% de integridade!")
    return saved_path

def run_todas_auditorias(planilhas_dir=None, output_dir=None, file_map=None, progress_cb=None, log_cb=None, prompt_lock_cb=None):
    """Executa sequencialmente as auditorias da Folha e do Cálculo de Férias."""
    f1 = run_auditoria_folha(planilhas_dir, output_dir, file_map, progress_cb, log_cb, prompt_lock_cb)
    f2 = run_auditoria_calculo(planilhas_dir, output_dir, file_map, progress_cb, log_cb, prompt_lock_cb)
    return f1, f2

if __name__ == '__main__':
    run_todas_auditorias()
