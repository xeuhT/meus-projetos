import os
import re
import csv
import unicodedata
import pandas as pd

def safe_read_csv(filepath, sep=';', encoding='latin1', skiprows=0, dtype=str, **kwargs):
    """Lê arquivos CSV nativos do sistema LG de forma ultra-resiliente."""
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

def normalize_text(s):
    """Remove acentos, pontuações e converte para minúsculas."""
    if not s:
        return ""
    n = ''.join(c for c in unicodedata.normalize('NFD', str(s)) if unicodedata.category(c) != 'Mn')
    return n.lower().strip()

def is_csv_fopa_calc(f_path):
    """Verifica se o arquivo FOPA é do Cálculo de Férias (Folha 2)."""
    try:
        with open(f_path, 'r', encoding='latin1') as f:
            for _ in range(4):
                line = next(f, "")
                if '"Folha ";' in line:
                    return '"2"' in line
        return False
    except:
        return False

def is_csv_fopa_folha(f_path):
    """Verifica se o arquivo FOPA é da Folha Mensal (Folha 1)."""
    try:
        with open(f_path, 'r', encoding='latin1') as f:
            for _ in range(4):
                line = next(f, "")
                if '"Folha ";' in line:
                    return '"1"' in line
        return False
    except:
        return False

def is_csv_fopa_folha_completa(f_path):
    """Verifica se o arquivo FOPA é da Folha Mensal Completa (Folha 1 com > 30 colunas e eventos de salário/férias)."""
    try:
        with open(f_path, 'r', encoding='latin1') as f:
            lines = [next(f, '') for _ in range(5)]
            is_folha1 = any('"Folha ";' in l and '"1"' in l for l in lines)
            if not is_folha1:
                return False
            if len(lines) >= 4:
                header_line = lines[3]
                cols = [c.strip().strip('"').lower() for c in header_line.split(';')]
                has_salario = any('salario' in c or 'salário' in c or '11220' in c for c in cols)
                return len(cols) >= 30 and has_salario
        return False
    except Exception:
        return False

def is_active_afastados(f_path):
    """Verifica se o arquivo de afastados contém colaboradores com status Ativo (evita arquivos históricos de Cessados)."""
    try:
        with open(f_path, 'r', encoding='latin1') as f:
            for _ in range(50):
                line = next(f, "")
                if ';ativo;' in line.lower() or ';"ativo";' in line.lower():
                    return True
        return False
    except Exception:
        return False

def scan_ferias_prog_months(f_path):
    """Identifica o mês predominante e total de registros em um arquivo de Férias Programadas."""
    try:
        df = safe_read_csv(f_path, sep=';', encoding='latin1', dtype=str)
        if 'Matr' not in str(df.columns[4]):
            df = safe_read_csv(f_path, sep=';', encoding='latin1', skiprows=1, dtype=str)
        
        idx_gozo = next((i for i, c in enumerate(df.columns) if 'per' in normalize_text(c) and 'gozo' in normalize_text(c)), 14)
        meses = []
        for v in df.iloc[:, idx_gozo].dropna():
            m = re.search(r'(\d{2})/(\d{2})/(\d{4})', str(v))
            if m:
                meses.append(int(m.group(2)))
        if meses:
            mes_num = max(set(meses), key=meses.count)
            nomes_meses = {1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril', 5: 'Maio', 6: 'Junho',
                           7: 'Julho', 8: 'Agosto', 9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'}
            return mes_num, nomes_meses.get(mes_num, f'Mês {mes_num}'), len(df)
        return None, "Indefinido", len(df)
    except Exception as e:
        return None, "Erro na leitura", 0

class FileDetector:
    def __init__(self, root_dir, output_dir=None):
        self.root_dir = os.path.abspath(root_dir)
        self.output_dir = os.path.abspath(output_dir) if output_dir else os.path.join(self.root_dir, '..', 'conferencias')
        self.all_files = []
        self.catalog = {}

    def scan(self):
        """Varre recursivamente a pasta raiz (os.walk) e cataloga todos os arquivos."""
        self.all_files = []
        if not os.path.exists(self.root_dir):
            return self.diagnose()

        for root, dirs, files in os.walk(self.root_dir):
            dirs[:] = [d for d in dirs if not any(d.lower().startswith(p) for p in ['planilhas organizadas', 'backup', 'old', 'arquivo', 'temp', '.'])]
            for f in files:
                if f.startswith('~$') or f.startswith('.'):
                    continue
                full_path = os.path.join(root, f)
                try:
                    # Suporte para caminhos longos no Windows se necessário
                    stat_path = full_path
                    if os.name == 'nt' and not stat_path.startswith('\\\\?\\') and len(stat_path) > 240:
                        stat_path = '\\\\?\\' + os.path.abspath(stat_path)
                    st = os.stat(stat_path)
                    mtime = st.st_mtime
                    size = st.st_size
                except Exception:
                    mtime = 0.0
                    size = 0

                self.all_files.append({
                    'name': f,
                    'full_path': full_path,
                    'norm': normalize_text(f),
                    'mtime': mtime,
                    'size': size,
                    'ext': os.path.splitext(f)[1].lower()
                })

        # Ordenar por data de modificação decrescente (mais recente primeiro)
        self.all_files.sort(key=lambda x: x['mtime'], reverse=True)
        return self.diagnose()

    def find_best(self, prefix, exts=('.csv',), subfolder_hint=None, validator=None):
        """Localiza o melhor arquivo que atende ao prefixo e critérios opcionais."""
        p_norm = normalize_text(prefix)
        candidates = []
        for item in self.all_files:
            if item['ext'] in exts and item['norm'].startswith(p_norm):
                if subfolder_hint and subfolder_hint not in normalize_text(item['full_path']):
                    continue
                if validator and not validator(item['full_path']):
                    continue
                candidates.append(item)

        if not candidates and subfolder_hint:
            # Tentar sem o subfolder hint se não achou com ele
            for item in self.all_files:
                if item['ext'] in exts and item['norm'].startswith(p_norm):
                    if validator and not validator(item['full_path']):
                        continue
                    candidates.append(item)

        if candidates:
            # Retorna o mais recente
            return candidates[0]
        return None

    def find_all_matching(self, prefix, exts=('.csv',)):
        """Localiza todos os arquivos com o prefixo."""
        p_norm = normalize_text(prefix)
        matches = [item for item in self.all_files if item['ext'] in exts and item['norm'].startswith(p_norm)]
        # Ordenados do mais antigo para o mais recente (para que novos sobrescrevam)
        matches.sort(key=lambda x: x['mtime'])
        return matches

    def diagnose(self):
        """Diagnostica as entradas para cada uma das 3 conferências."""
        if not self.all_files and os.path.exists(self.root_dir):
            return self.scan()

        # 1. Arquivos base gerais
        colab_master = self.find_best('cola_colaboradores_csv')
        ferias_hist = self.find_best('cencosud_ferias_colaboradores_csv')
        faltas_pa = self.find_best('cola_falta_injustifica_pa')
        # Prioriza arquivo de afastados que contém colaboradores com status Ativo
        afastados = self.find_best('cola_colaboradores_afastados', validator=is_active_afastados) or self.find_best('cola_colaboradores_afastados')
        feriados = self.find_best('feriados_cadastrados')
        pensionistas = self.find_best('cola_beneficiario_pensionista')

        # ORGA Estabelecimento (prioriza Excel com Feriados)
        orga_xlsx = self.find_best('orga_estabelecimento_csv', exts=('.xlsx', '.xls'))
        orga_csv = self.find_best('orga_estabelecimento_csv', exts=('.csv',))
        orga_file = orga_xlsx or orga_csv

        # Base PA Aberto (COLA_Conferencia_PA_Aberto em CSV ou Excel)
        pa_aberto_model = self.find_best('cola_conferencia_pa_aberto', exts=('.csv', '.xlsx', '.xls'))
        if not pa_aberto_model:
            for item in self.all_files:
                if item['ext'] in ('.csv', '.xlsx', '.xls') and 'pa_aberto' in item['norm']:
                    pa_aberto_model = item
                    break

        # Férias Programadas (múltiplos meses)
        prog_items = self.find_all_matching('ferias programadas ou gozadas em csv')
        prog_details = []
        months_found = {}
        for item in prog_items:
            m_num, m_nome, count = scan_ferias_prog_months(item['full_path'])
            prog_details.append({
                'item': item,
                'mes_num': m_num,
                'mes_nome': m_nome,
                'count': count
            })
            if m_num:
                months_found[m_nome] = item['full_path']

        # FOPA Cálculo vs FOPA Folha (prioriza sempre o arquivo mais recente que atende aos validadores completos)
        fopa_calc = self.find_best('fopa_eventos_por_colaborador', validator=is_csv_fopa_calc) or self.find_best('fopa_eventos_por_colaborador', subfolder_hint='calculo', validator=is_csv_fopa_calc)
        fopa_folha = self.find_best('fopa_eventos_por_colaborador', validator=is_csv_fopa_folha_completa) or self.find_best('fopa_eventos_por_colaborador', subfolder_hint='folha', validator=is_csv_fopa_folha) or self.find_best('fopa_eventos_por_colaborador', validator=is_csv_fopa_folha)

        # Arquivos de referência/modelo de Auditoria (Layout de colunas)
        # Prioridade máxima para a pasta 'ja validadas'
        validadas_dir = os.path.join(self.output_dir, "ja validadas")
        ref_calc = None
        ref_folha = None
        if os.path.exists(validadas_dir):
            for f in os.listdir(validadas_dir):
                if f.lower().endswith('.xlsx') and not f.startswith('~$'):
                    fnorm = normalize_text(f)
                    if 'calculo_ferias' in fnorm and not ref_calc:
                        ref_calc = {'name': f, 'full_path': os.path.join(validadas_dir, f)}
                    elif 'folha_mensal' in fnorm and not ref_folha:
                        ref_folha = {'name': f, 'full_path': os.path.join(validadas_dir, f)}

        if not ref_calc:
            ref_calc = self.find_best('auditoria_calculo_ferias', exts=('.xlsx',))
        if not ref_calc and os.path.exists(self.output_dir):
            for f in os.listdir(self.output_dir):
                if f.lower().endswith('.xlsx') and 'calculo_ferias' in normalize_text(f) and not f.startswith('~$'):
                    ref_calc = {'name': f, 'full_path': os.path.join(self.output_dir, f)}
                    break

        if not ref_folha:
            ref_folha = self.find_best('auditoria_folha_mensal', exts=('.xlsx',))
        if not ref_folha and os.path.exists(self.output_dir):
            for f in os.listdir(self.output_dir):
                if f.lower().endswith('.xlsx') and 'folha_mensal' in normalize_text(f) and not f.startswith('~$'):
                    ref_folha = {'name': f, 'full_path': os.path.join(self.output_dir, f)}
                    break

        # -------------------------------------------------------------
        # Montagem do Diagnóstico da ABA 1: Conferência Férias Formatada
        # -------------------------------------------------------------
        aba1_items = [
            {'label': 'Base Colaboradores (COLA)', 'file': colab_master, 'req': True, 'status': 'OK' if colab_master else 'MISSING'},
            {'label': 'Faltas Injustificadas PA', 'file': faltas_pa, 'req': True, 'status': 'OK' if faltas_pa else 'MISSING'},
            {'label': 'Férias Programadas / Gozadas', 'file': prog_details[-1]['item'] if prog_details else None, 'req': True,
             'status': 'OK' if prog_details else 'MISSING',
             'obs': f"{len(prog_details)} arquivo(s) detectado(s)" if prog_details else ""},
            {'label': 'Histórico Férias Colaboradores', 'file': ferias_hist, 'req': True, 'status': 'OK' if ferias_hist else 'MISSING'},
            {'label': 'Estabelecimentos / Feriados (ORGA)', 'file': orga_file, 'req': True, 'status': 'OK' if orga_file else 'MISSING'},
            {'label': 'Conferência PA Aberto (COLA)', 'file': pa_aberto_model, 'req': False,
             'status': 'OK' if pa_aberto_model else 'OPTIONAL',
             'obs': '' if pa_aberto_model else 'Opcional (gerado via bases CSV se ausente)'},
            {'label': 'Colaboradores Afastados', 'file': afastados, 'req': True, 'status': 'OK' if afastados else 'MISSING'},
            {'label': 'Tabela Feriados Cadastrados', 'file': feriados, 'req': True, 'status': 'OK' if feriados else 'MISSING'},
        ]
        aba1_ready = all(i['file'] is not None for i in aba1_items if i['req'])

        # -------------------------------------------------------------
        # Montagem do Diagnóstico da ABA 2: Auditoria Cálculo de Férias
        # -------------------------------------------------------------
        aba2_items = [
            {'label': 'Eventos FOPA - Cálculo (Folha 2)', 'file': fopa_calc, 'req': True, 'status': 'OK' if fopa_calc else 'MISSING'},
            {'label': 'Base Colaboradores Master', 'file': colab_master, 'req': True, 'status': 'OK' if colab_master else 'MISSING'},
            {'label': 'Histórico Férias Colaboradores', 'file': ferias_hist, 'req': True, 'status': 'OK' if ferias_hist else 'MISSING'},
            {'label': 'Beneficiários / Pensionistas', 'file': pensionistas, 'req': True, 'status': 'OK' if pensionistas else 'MISSING'},
            {'label': 'Férias Programadas (PA Aquisitivo)', 'file': prog_details[-1]['item'] if prog_details else None, 'req': True,
             'status': 'OK' if prog_details else 'MISSING',
             'obs': f"{len(prog_details)} base(s) detectada(s) para PA" if prog_details else ""},
            {'label': 'Eventos FOPA - Folha (Insalubridade)', 'file': fopa_folha, 'req': True, 'status': 'OK' if fopa_folha else 'MISSING'},
        ]
        aba2_ready = all(i['file'] is not None for i in aba2_items if i['req'])

        # -------------------------------------------------------------
        # Montagem do Diagnóstico da ABA 3: Auditoria Folha Mensal
        # -------------------------------------------------------------
        prog_months_text = ", ".join(months_found.keys()) if months_found else "Nenhum"
        aba3_items = [
            {'label': 'Eventos FOPA - Folha (Folha 1)', 'file': fopa_folha, 'req': True, 'status': 'OK' if fopa_folha else 'MISSING'},
            {'label': 'Base Colaboradores Master', 'file': colab_master, 'req': True, 'status': 'OK' if colab_master else 'MISSING'},
            {'label': 'Histórico Férias Colaboradores', 'file': ferias_hist, 'req': True, 'status': 'OK' if ferias_hist else 'MISSING'},
            {'label': 'Colaboradores Afastados', 'file': afastados, 'req': True, 'status': 'OK' if afastados else 'MISSING'},
            {'label': 'Férias Programadas Múltiplos Meses', 'file': prog_details[-1]['item'] if prog_details else None, 'req': True,
             'status': 'OK' if prog_details else 'MISSING',
             'obs': f"{len(prog_details)} arquivo(s): {prog_months_text}" if prog_details else ""},
        ]
        aba3_ready = all(i['file'] is not None for i in aba3_items if i['req'])

        return {
            'aba1': {'items': aba1_items, 'ready': aba1_ready},
            'aba2': {'items': aba2_items, 'ready': aba2_ready},
            'aba3': {'items': aba3_items, 'ready': aba3_ready},
            'raw_map': {
                'colab_master': colab_master['full_path'] if colab_master else None,
                'ferias_hist': ferias_hist['full_path'] if ferias_hist else None,
                'faltas_pa': faltas_pa['full_path'] if faltas_pa else None,
                'afastados': afastados['full_path'] if afastados else None,
                'feriados': feriados['full_path'] if feriados else None,
                'pensionistas': pensionistas['full_path'] if pensionistas else None,
                'orga_file': orga_file['full_path'] if orga_file else None,
                'pa_aberto_model': pa_aberto_model['full_path'] if pa_aberto_model else None,
                'prog_files': [p['item']['full_path'] for p in prog_details],
                'prog_details': prog_details,
                'fopa_calc': fopa_calc['full_path'] if fopa_calc else None,
                'fopa_folha': fopa_folha['full_path'] if fopa_folha else None,
                'ref_calc': ref_calc['full_path'] if ref_calc else None,
                'ref_folha': ref_folha['full_path'] if ref_folha else None,
            }
        }
