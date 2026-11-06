"""
leitores.py - Leitura e padronização de TODAS as planilhas.
V19 - Máquina do Tempo Automática (Lê a Competência na linha 3 da Fopag).
"""
from pathlib import Path
import logging
import pandas as pd
import glob

import config
from utils import limpar_cpf, limpar_matricula, normalizar_nome

log = logging.getLogger(__name__)

def _remover_fantasmas(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty: return df
    return df.dropna(how='all').copy()

def _ler_excel_dinamico(caminho, aba_esperada, header_idx):
    xls = pd.ExcelFile(caminho)
    abas = xls.sheet_names
    if aba_esperada in abas: return _remover_fantasmas(pd.read_excel(xls, sheet_name=aba_esperada, header=header_idx, dtype=str))
    abas_parciais = [a for a in abas if str(aba_esperada).lower() in a.lower()]
    if abas_parciais: return _remover_fantasmas(pd.read_excel(xls, sheet_name=abas_parciais[0], header=header_idx, dtype=str))
    abas_lixo = ["legenda", "orgânico unificado", "organico unificado", "solicitação rejeitada", "solicitacao rejeitada", "chamados rejeitados"]
    abas_uteis = [a for a in abas if str(a).lower().strip() not in abas_lixo]
    if len(abas_uteis) == 1: return _remover_fantasmas(pd.read_excel(xls, sheet_name=abas_uteis[0], header=header_idx, dtype=str))
    raise ValueError(f"Aba correta para {aba_esperada} não localizada.")

def _padronizar(df: pd.DataFrame, mapeamento: dict) -> pd.DataFrame:
    import unicodedata
    saida = df.copy()
    saida.columns = [str(c).strip() for c in saida.columns]
    
    def norm(s):
        return ''.join(c for c in unicodedata.normalize('NFKD', str(s)) if not unicodedata.combining(c)).lower().strip()
    
    col_map = {norm(c): c for c in saida.columns}
    
    for nome_padrao, nome_origem in mapeamento.items():
        origem_limpo = norm(nome_origem)
        if origem_limpo in col_map:
            saida[nome_padrao] = saida[col_map[origem_limpo]]
        else:
            saida[nome_padrao] = ""
    return saida

def _aplicar_limpeza_comum(df: pd.DataFrame) -> pd.DataFrame:
    if "cpf" in df.columns: df["cpf"] = df["cpf"].apply(limpar_cpf)
    if "matricula" in df.columns: df["matricula"] = df["matricula"].apply(limpar_matricula)
    for col_nome in ("titular", "beneficiario", "nome"):
        if col_nome in df.columns: df[col_nome] = df[col_nome].apply(normalizar_nome)
        
    # Remover linhas fantasmas que s tm CPF (sem nome, sem matricula, etc)
    # Se todas as colunas de nome estiverem vazias, consideramos lixo.
    colunas_nome_presentes = [c for c in ("titular", "beneficiario", "nome") if c in df.columns]
    if colunas_nome_presentes:
        # Pelo menos um nome tem que ser no vazio
        mask = df[colunas_nome_presentes].replace("", float("NaN")).notna().any(axis=1)
        df = df[mask].copy()
        
    return df

def ler_ativos(caminho: Path, operadora: str) -> pd.DataFrame:
    cfg = config.ATIVOS.get(operadora)
    if not cfg: return pd.DataFrame()
    if not caminho or not Path(caminho).exists(): return pd.DataFrame()
    try: df = _ler_excel_dinamico(caminho, cfg["aba"], cfg["header"])
    except Exception: return pd.DataFrame()
    if "sulam" in operadora.lower():
        if len(df.columns) > 3: df["bandeira_garantida"] = df.iloc[:, 3]
        if len(df.columns) > 8:
            df["nascimento_garantido_7"] = df.iloc[:, 7]
            df["nascimento_garantido_8"] = df.iloc[:, 8]
    df = _padronizar(df, cfg["colunas"])
    df = _aplicar_limpeza_comum(df)
    df["operadora"] = operadora
    return df

def ler_fopag(caminho: Path, operadora: str) -> pd.DataFrame:
    cfg = config.FOPAG.get(operadora)
    if not cfg: return pd.DataFrame()
    if not caminho or not Path(caminho).exists(): return pd.DataFrame()
    try: df = _ler_excel_dinamico(caminho, cfg["aba"], cfg["header"])
    except Exception: return pd.DataFrame()
    df = _padronizar(df, cfg["colunas"])
    df = _aplicar_limpeza_comum(df)
    df["operadora"] = operadora
    return df

def ler_colaboradores(caminho: Path) -> pd.DataFrame:
    import unicodedata
    cfg = config.COLABORADORES
    if not caminho or not Path(caminho).exists(): return pd.DataFrame()
    extensao = str(caminho).lower()
    kws_header = ["matricula", "colaborador", "cpf", "situacao", "cargo"]

    def _norm_txt(s: str) -> str:
        return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c)).lower()

    try:
        header_idx = cfg.get("header", 0)
        if extensao.endswith(".csv"):
            linhas_topo = []
            for enc in ("utf-8", "latin-1"):
                try:
                    with open(caminho, "r", encoding=enc) as f:
                        linhas_topo = [f.readline() for _ in range(10)]
                    break
                except UnicodeDecodeError:
                    continue
            for idx, ln in enumerate(linhas_topo):
                txt = _norm_txt(ln)
                if sum(1 for k in kws_header if k in txt) >= 3:
                    header_idx = idx
                    break
            try: df = pd.read_csv(caminho, sep=cfg.get("separador", ";"), encoding="utf-8", header=header_idx, dtype=str, low_memory=False)
            except UnicodeDecodeError: df = pd.read_csv(caminho, sep=cfg.get("separador", ";"), encoding="latin-1", header=header_idx, dtype=str, low_memory=False)
        else:
            aba = cfg.get("aba", 0)
            df_raw_top = pd.read_excel(caminho, sheet_name=aba, header=None, nrows=10, dtype=str)
            for idx, row in df_raw_top.iterrows():
                txt = _norm_txt(" ".join(row.dropna().astype(str).values))
                if sum(1 for k in kws_header if k in txt) >= 3:
                    header_idx = int(idx)
                    break
            df = pd.read_excel(caminho, sheet_name=aba, header=header_idx, dtype=str)

        df = _remover_fantasmas(df)
        df = df.loc[:, ~df.columns.duplicated()].copy()
        df = _padronizar(df, cfg["colunas"])

        if "situacao" in df.columns and "tipo_situacao" in df.columns:
            mask_vazio = df["situacao"].astype(str).str.strip().isin(["", "nan", "NaN", "None"])
            df.loc[mask_vazio, "situacao"] = df.loc[mask_vazio, "tipo_situacao"]

        if ("situacao" not in df.columns or (df["situacao"].astype(str).str.strip() == "").all()) and len(df.columns) > 105:
            df["situacao"] = df.iloc[:, 105]

        df = _aplicar_limpeza_comum(df)
        if not df.empty:
            df["cpf_colaborador"] = df["cpf"] if "cpf" in df.columns else ""
            df["matricula_colaborador_rh"] = df["matricula"] if "matricula" in df.columns else ""
        log.info("Base de Colaboradores carregada: %d registros (header=%d) de '%s'.", len(df), header_idx, Path(caminho).name)
        return df
    except Exception as e:
        log.error("Erro ao ler Base de Colaboradores '%s': %s", caminho, e)
        return pd.DataFrame()

def ler_demitidos(caminho: Path) -> pd.DataFrame:
    cfg = config.DEMITIDOS
    if not caminho or not Path(caminho).exists(): return pd.DataFrame()
    extensao = str(caminho).lower()
    try:
        if extensao.endswith(".csv"):
            try: df = pd.read_csv(caminho, sep=cfg.get("separador", ";"), encoding="utf-8", header=cfg["header"], dtype=str, low_memory=False)
            except UnicodeDecodeError: df = pd.read_csv(caminho, sep=cfg.get("separador", ";"), encoding="latin-1", header=cfg["header"], dtype=str, low_memory=False)
        else: df = pd.read_excel(caminho, sheet_name=cfg.get("aba", 0), header=cfg["header"], dtype=str)
        df = _remover_fantasmas(df)
        df = df.loc[:, ~df.columns.duplicated()].copy()
        df = _padronizar(df, cfg["colunas"])
        df = _aplicar_limpeza_comum(df)
        return df
    except Exception: return pd.DataFrame()

def ler_dependentes(caminho: Path) -> pd.DataFrame:
    cfg = config.DEPENDENTES
    if not caminho or not Path(caminho).exists(): return pd.DataFrame()
    try: df = pd.read_csv(caminho, sep=cfg.get("separador", ";"), encoding="utf-8", header=cfg.get("header", 0), dtype=str, low_memory=False)
    except UnicodeDecodeError: df = pd.read_csv(caminho, sep=cfg.get("separador", ";"), encoding="latin-1", header=cfg.get("header", 0), dtype=str, low_memory=False)
    except Exception: df = pd.read_excel(caminho, header=cfg.get("header", 0), dtype=str)
    df = _remover_fantasmas(df)
    df = df.loc[:, ~df.columns.duplicated()].copy()
    if len(df.columns) >= 27:
        df["codigo_dependente_garantido"] = df.iloc[:, 11]
        df["cpf_marsh_garantido"] = df.iloc[:, 13]
        df["nascimento_marsh_garantido"] = df.iloc[:, 14]
        df["codigo_instrucao_garantido"] = df.iloc[:, 24]
        df["desc_instrucao_garantido"] = df.iloc[:, 25]
        df["invalido_garantido"] = df.iloc[:, 26]
    saida = df.copy()
    for idx, nome_padrao in cfg["colunas_indices"].items():
        if idx < len(df.columns): saida[nome_padrao] = df.iloc[:, idx]
        else: saida[nome_padrao] = ""
    saida = _aplicar_limpeza_comum(saida)
    return saida

# =========================================================================
# LEITORES FINANCEIROS
# =========================================================================

def ler_dicionario_eventos(caminho: Path) -> pd.DataFrame:
    if not caminho or not Path(caminho).exists(): return pd.DataFrame()
    extensao = str(caminho).lower()
    try:
        if extensao.endswith(".xlsx") or extensao.endswith(".xls"):
            df = pd.read_excel(caminho, dtype=str)
        else:
            try: df = pd.read_csv(caminho, sep=";", encoding="latin-1", dtype=str, on_bad_lines='skip')
            except: df = pd.read_csv(caminho, sep=",", encoding="utf-8", dtype=str, on_bad_lines='skip')

        if len(df.columns) >= 3:
            df = df.iloc[:, :3].copy()
            df.columns = ["CODIGO", "DESCRICAO", "PRODUTO"]
            log.info("Dicionário de Eventos carregado com Sucesso. (%d rubricas lidas)", len(df))
            return df
        else:
            log.error("Dicionário de Eventos inválido: Possui menos de 3 colunas.")
            return pd.DataFrame()
    except Exception as e:
        log.error("Erro Crítico ao ler Dicionário de Eventos: %s", e)
        return pd.DataFrame()

def ler_financeiro_fopag(diretorio: Path, prefixo: str = "FOPA_Eventos_por_Colaborador") -> pd.DataFrame:
    arquivos = list(Path(diretorio).glob(f"{prefixo}*.*"))
    arquivos = [a for a in arquivos if str(a).lower().endswith(('.csv', '.xlsx', '.xls'))]
    if not arquivos:
        log.warning("Nenhum arquivo financeiro encontrado com o prefixo: %s", prefixo)
        return pd.DataFrame()
    lista_dfs = []
    competencia_dt = None
    import re
    meses_map = {'jan':1, 'fev':2, 'mar':3, 'abr':4, 'mai':5, 'jun':6, 'jul':7, 'ago':8, 'set':9, 'out':10, 'nov':11, 'dez':12}

    def _parse_comp(texto_raw: str):
        if not texto_raw: return None
        s = str(texto_raw).strip()
        m_iso = re.search(r"(\d{4})-(\d{2})-(\d{2})", s)
        if m_iso:
            return pd.to_datetime(f"{m_iso.group(1)}-{m_iso.group(2)}-01")
        m_br_full = re.search(r"(\d{2})/(\d{2})/(\d{4})", s)
        if m_br_full:
            return pd.to_datetime(f"{m_br_full.group(3)}-{m_br_full.group(2)}-01")
        m_my = re.search(r"([A-Za-zçÇ]{3}|\d{1,2})/(\d{2,4})", s)
        if m_my:
            p1, p2 = m_my.group(1).lower(), m_my.group(2)
            m_num = int(p1) if p1.isdigit() else meses_map.get(p1[:3])
            if m_num and 1 <= m_num <= 12:
                y_num = 2000 + int(p2) if len(p2) == 2 else int(p2)
                return pd.to_datetime(f"{y_num}-{m_num:02d}-01")
        return None

    for arq in arquivos:
        log.info("Processando base financeira: %s", arq.name)
        try:
            ext = str(arq).lower()
            if ext.endswith('.csv'):
                with open(arq, 'r', encoding='latin-1') as f:
                    for i, line in enumerate(f):
                        if i <= 3 and "compet" in line.lower():
                            dt_c = _parse_comp(line)
                            if dt_c is not None: competencia_dt = dt_c
                        if i > 3: break
                df = pd.read_csv(arq, sep=";", encoding="latin-1", skiprows=3, dtype=str, low_memory=False)
            else:
                df_header = pd.read_excel(arq, header=None, nrows=4, dtype=str)
                for i, row in df_header.iterrows():
                    row_str = " ".join(row.dropna().astype(str).values)
                    if "compet" in row_str.lower():
                        dt_c = _parse_comp(row_str)
                        if dt_c is not None: competencia_dt = dt_c
                df = pd.read_excel(arq, skiprows=3, dtype=str)
            if competencia_dt is None:
                m_fn = re.search(r"(0[1-9]|1[0-2])(20\d{2})", arq.name)
                if m_fn:
                    competencia_dt = pd.to_datetime(f"{m_fn.group(2)}-{m_fn.group(1)}-01")
            lista_dfs.append(df)
        except Exception as e:
            log.error("Erro na leitura financeira: %s", e)
    if not lista_dfs: return pd.DataFrame()
    df_consolidado = pd.concat(lista_dfs, ignore_index=True)
    df_consolidado.columns = [str(c).strip() for c in df_consolidado.columns]
    df_consolidado = df_consolidado.dropna(how='all')

    colunas_cadastro = ["Código da Empresa", "Nome da Empresa", "Código Estabelecimento", "Estabelecimento", "Código Centro Custo", "Centro Custo", "Matrícula", "Nome do Colaborador", "ID Colaborador", "Data Admissão", "Código Cargo", "Cargo", "Situação", "Tipo Situação", "Data Rescisão"]
    colunas_id = [c for c in colunas_cadastro if c in df_consolidado.columns]
    colunas_eventos = [c for c in df_consolidado.columns if c not in colunas_id and "Unnamed" not in c]

    df_melt = pd.melt(df_consolidado, id_vars=colunas_id, value_vars=colunas_eventos, var_name="EVENTO_NOME", value_name="VALOR")
    df_melt = df_melt.dropna(subset=["VALOR"])
    df_melt = df_melt[~df_melt["VALOR"].astype(str).str.contains(r'^\s*-\s*$')]

    def _clean_valor(v):
        if pd.isna(v): return 0.0
        s = str(v).strip()
        if not s or s == "-": return 0.0
        if "." in s and "," in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except:
            return 0.0
            
    df_melt["VALOR_FLOAT"] = df_melt["VALOR"].apply(_clean_valor)
    df_melt = df_melt[df_melt["VALOR_FLOAT"] > 0]

    df_melt["CODIGO_EVENTO"] = df_melt["EVENTO_NOME"].astype(str).str.extract(r'(\d+)')[0]
    if "Matrícula" in df_melt.columns: df_melt["matricula"] = df_melt["Matrícula"].apply(limpar_matricula).astype(str)

    # Anexa o mês da máquina do tempo aos dados
    df_melt["COMPETENCIA_DATA"] = competencia_dt if competencia_dt else pd.NaT

    log.info("Base Financeira consolidada. %d rubricas de desconto localizadas.", len(df_melt))
    return df_melt


def ler_faturas_operadoras(diretorio: Path) -> pd.DataFrame:
    import os
    import glob
    import pandas as pd
    import logging
    from utils import limpar_cpf, limpar_matricula

    log = logging.getLogger(__name__)
    registros = []
    arquivos = glob.glob(os.path.join(str(diretorio), "*.*"))

    log.info(f"🔎 Iniciando leitura de faturas na pasta...")

    def get_col(df, keywords):
        import re
        def kw_in_str(kw, text):
            if len(kw) <= 3:
                return bool(re.search(r'\b' + re.escape(kw) + r'\b', text))
            return kw in text
        for col in df.columns:
            col_str = str(col).upper()
            for kw in keywords:
                if kw_in_str(kw, col_str): return col
        return None

    def clean_val(v):
        if pd.isna(v): return 0.0
        s = str(v).upper().replace("R$", "").strip().replace(" ", "")
        if not s: return 0.0
        if ',' in s and '.' in s: s = s.replace('.', '')
        s = s.replace(',', '.')
        try:
            return float(s)
        except:
            return 0.0

    def ajustar_cabecalho(df, palavras_chave):
        import re
        def kw_in_str(kw, text):
            if len(kw) <= 3:
                return bool(re.search(r'\b' + re.escape(kw) + r'\b', text))
            return kw in text
        col_str = " ".join([str(c).upper() for c in df.columns])
        if sum(1 for w in palavras_chave if kw_in_str(w, col_str)) >= 2:
            return df

        for idx, row in df.iterrows():
            row_str = " ".join([str(c).upper() for c in row.values])
            if sum(1 for w in palavras_chave if kw_in_str(w, row_str)) >= 2:
                df.columns = df.iloc[idx]
                return df.iloc[idx + 1:].reset_index(drop=True)
        return df

    def extrair_tabela(arq, chaves_busca, eh_saude=False):
        # 1. TENTA LER COMO EXCEL NATIVO (Rápido)
        try:
            import warnings
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                if eh_saude:
                    df = pd.read_excel(arq, sheet_name="SeguradosAtivos", header=None)
                else:
                    df = pd.read_excel(arq, header=None)

            df = ajustar_cabecalho(df, chaves_busca)
            if sum(1 for w in chaves_busca if w in " ".join([str(c).upper() for c in df.columns])) >= 2:
                return df
        except:
            pass

        # 2. MODO DEUS: Invoca o MS Excel para consertar o arquivo corrompido
        if str(arq).upper().endswith(('.XLS', '.XLSX')):
            try:
                import win32com.client as win32
                import pythoncom

                # CoInitialize é crucial para não travar a Thread do CustomTkinter
                pythoncom.CoInitialize()

                excel = win32.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False

                wb = excel.Workbooks.Open(os.path.abspath(arq))

                if eh_saude:
                    ws = wb.Sheets("SeguradosAtivos")
                else:
                    ws = wb.Sheets(1)

                dados = ws.UsedRange.Value
                wb.Close(False)
                excel.Quit()
                pythoncom.CoUninitialize()

                if dados:
                    df = pd.DataFrame(dados)
                    df = ajustar_cabecalho(df, chaves_busca)
                    if sum(1 for w in chaves_busca if w in " ".join([str(c).upper() for c in df.columns])) >= 2:
                        return df
            except Exception as e:
                log.debug(f"Modo Deus falhou em {os.path.basename(arq)}: {e}")
                try:
                    excel.Quit()
                    pythoncom.CoUninitialize()
                except:
                    pass

        # 3. TENTA LER COMO TSV/CSV
        for separador in ['\t', ';', ',']:
            try:
                df = pd.read_csv(arq, sep=separador, encoding='latin-1', header=None, on_bad_lines='skip',
                                 engine='python')
                df = ajustar_cabecalho(df, chaves_busca)
                if sum(1 for w in chaves_busca if w in " ".join([str(c).upper() for c in df.columns])) >= 2:
                    return df
            except:
                pass

        # 4. TENTA LER COMO HTML (Sem vomitar no terminal)
        try:
            import warnings
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                dfs = pd.read_html(arq, decimal=',', thousands='.', encoding='latin-1')
                for df in dfs:
                    df = ajustar_cabecalho(df, chaves_busca)
                    if sum(1 for w in chaves_busca if w in " ".join([str(c).upper() for c in df.columns])) >= 2:
                        return df
        except:
            pass

        return pd.DataFrame()

    for arq in arquivos:
        nome_arq = os.path.basename(arq).upper()
        if not nome_arq.endswith(('.XLS', '.XLSX', '.CSV')): continue

        # Pula bases do RH/Fopag/Ativos
        if any(x in nome_arq for x in
               ["COLA_", "DEPENDENTES", "DESL", "FOPA_", "EVENTOS", "MOVIMENT", "ATIVOS", "ATIV_"]):
            continue

        try:
            op = ""
            df_target = pd.DataFrame()

            if "SUL AMERICA" in nome_arq and "SAUDE" in nome_arq and "PREMIO" in nome_arq:
                op = "SULAMERICA_SAUDE"
                chaves_busca = ["ID FUNCIONAL", "CPF", "NOME SEGURADO"]
                df_target = extrair_tabela(arq, chaves_busca, eh_saude=True)

            elif "KLINI" in nome_arq and "MENSALIDADE" in nome_arq:
                op = "KLINI"
                chaves_busca = ["MATRÍCULA DO ESTIPULANTE", "BENEFICIÁRIO"]
                df_target = extrair_tabela(arq, chaves_busca)

            elif "SUL AMERICA" in nome_arq and "ODONTO" in nome_arq and "PREMIO" in nome_arq:
                op = "SULAMERICA_ODONTO"
                chaves_busca = ["RE", "NOME", "CPF", "TOTAL"]
                df_target = extrair_tabela(arq, chaves_busca)

            elif "HAPVIDA" in nome_arq and "PREMIO" in nome_arq:
                op = "HAPVIDA"
                chaves_busca = ["MATRÍCULA", "NOME DO BENEFICIÁRIO"]
                df_target = extrair_tabela(arq, chaves_busca)

            elif "GNDI" in nome_arq and "MENSALIDADE" in nome_arq:
                op = "GNDI"
                chaves_busca = ["MATRICULA", "BENEFICIÁRIO", "COBRADO"]
                df_target = extrair_tabela(arq, chaves_busca)

            elif "UNIMED" in nome_arq and "MENSALIDADE" in nome_arq:
                op = "UNIMED"
                chaves_busca = ["MATRICULA", "BENEFICIARIO_NOME"]
                df_target = extrair_tabela(arq, chaves_busca)

            # --- VALIDAÇÃO COM LOGS DE ALERTA ---
            if df_target is not None and not df_target.empty and op:
                if op == "SULAMERICA_SAUDE":
                    col_mat = get_col(df_target, ["ID FUNCIONAL", "MATR"])
                    col_cpf = get_col(df_target, ["CPF"])
                    col_nome = get_col(df_target, ["NOME SEGURADO", "NOME"])
                    col_val = get_col(df_target, ["PRÊMIO", "PREMIO"])
                elif op == "KLINI":
                    col_mat = get_col(df_target, ["MATRÍCULA DO ESTIPULANTE", "MATR"])
                    col_cpf = get_col(df_target, ["CPF"])
                    col_nome = get_col(df_target, ["BENEFICIÁRIO", "NOME"])
                    col_val = get_col(df_target, ["VALOR DO ITEM"])
                elif op == "SULAMERICA_ODONTO":
                    col_mat = get_col(df_target, ["RE", "MATR"])
                    col_cpf = get_col(df_target, ["CPF"])
                    col_nome = get_col(df_target, ["NOME", "TITULAR"])
                    col_val = get_col(df_target, ["TOTAL"])
                elif op == "HAPVIDA":
                    col_mat = get_col(df_target, ["MATRÍCULA", "MATR"])
                    col_cpf = get_col(df_target, ["CPF"])
                    col_nome = get_col(df_target, ["NOME DO BENEFICIÁRIO", "NOME"])
                    col_val = get_col(df_target, ["TOTAL", "MENSALIDADE"])
                elif op == "GNDI":
                    col_mat = get_col(df_target, ["MATRICULA", "MATR"])
                    col_cpf = get_col(df_target, ["CPF BENEF", "CPF"])
                    col_nome = get_col(df_target, ["BENEFICIÁRIO", "NOME"])
                    col_val = get_col(df_target, ["COBRADO", "TOTAL"])
                elif op == "UNIMED":
                    col_mat = get_col(df_target, ["MATRICULA", "MATR"])
                    col_cpf = get_col(df_target, ["BENEFICIARIO_CPF", "CPF"])
                    col_nome = get_col(df_target, ["BENEFICIARIO_NOME", "NOME"])
                    col_val = get_col(df_target, ["VALOR_MENSALIDADE", "VALOR"])

                if col_mat and col_val:
                    linhas_antes = len(registros)
                    for _, r in df_target.iterrows():
                        mat = limpar_matricula(r.get(col_mat, ""))
                        cpf = limpar_cpf(r.get(col_cpf, ""))
                        nome = str(r.get(col_nome, "")).strip().upper()
                        val = clean_val(r.get(col_val, 0))

                        if val > 0:
                            registros.append({"OPERADORA": op, "MATRICULA_FAT": mat, "CPF_FAT": cpf, "NOME_FAT": nome,
                                              "VALOR_FATURADO": val})

                    lidas = len(registros) - linhas_antes
                    if lidas > 0:
                        log.info(f"✅ {op} lida: {lidas} faturamentos encontrados em '{nome_arq}'.")
                    else:
                        log.warning(
                            f"⚠️ {op}: As colunas foram mapeadas, mas não havia faturamentos maiores que zero em '{nome_arq}'.")
                else:
                    log.warning(
                        f"⚠️ {op}: A tabela foi aberta, mas não achei as colunas essenciais no arquivo '{nome_arq}'.")
            elif op:
                log.warning(
                    f"⚠️ {op}: O arquivo '{nome_arq}' falhou em todas as tentativas ou está completamente vazio.")

        except Exception as e:
            log.error(f"❌ Erro processando fatura {nome_arq}: {e}")

    df_final = pd.DataFrame(registros)
    log.info(f"💰 Leitura de faturas concluída. Total mapeado: {len(df_final)} cobranças.")
    return df_final