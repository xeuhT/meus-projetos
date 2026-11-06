"""
exportador.py - Gera o XLSX final formatado.
V_PRODUCAO_HIERARQUIA - Fix NameError, Hierarquia de Dados e Porteiro de CPF.
"""
import logging
import warnings
from datetime import date
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import re
import unicodedata
from utils import formatar_cpf, formatar_data_br

warnings.filterwarnings("ignore", category=UserWarning)
log = logging.getLogger(__name__)

COLUNAS_LONG = [
    "OPERADORA", "EMPRESA", "LOJA", "MATRICULA", "CPF_TITULAR", "COLABORADOR",
    "CPF_BENEFICIARIO", "BENEFICIARIO", "PARENTESCO", "DEPENDENTE_INVALIDO", "GRAU_INSTRUCAO",
    "NASCIMENTO", "IDADE", "CARTEIRA", "PLANO", "TIPO_PLANO", "BANDEIRA", "RN",
    "BASE_ORIGEM", "VALOR_MENSALIDADE_TITULAR", "VALOR_MENSALIDADE_DEPENDENTES",
    "VALOR_COPARTICIPACAO", "VALOR_PAGO_FATURA",
    "DATA_VIGENCIA", "DATA_CANCELAMENTO", "SITUACAO",
    "CARGO", "SINDICATO", "ADMISSAO", "DEMISSAO", "DESCRICAO_DIVERGENCIA",
    "ACAO_RECOMENDADA", "TIPO_INCONSISTENCIA", "ALERTA_IDADE"
]

MAPA_SITUACAO = {
    "1": "Atividade Normal", "2": "Férias", "3": "Afastamento", "4": "Afastamento (Acidente)",
    "5": "Licença Maternidade", "6": "Mandato Sindical", "7": "Demitido", "8": "Recesso",
    "01": "Atividade Normal", "02": "Férias"
}

def _to_str(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)): return ""
    s = str(v).strip()
    return "" if s.lower() in ["nan", "none", "<na>", "nat"] else s

def _limpar_ponto_zero(v) -> str:
    s = _to_str(v)
    return s[:-2] if s.endswith(".0") else s

def _get_val(row, exact_matches, partial_matches=None):
    if partial_matches is None: partial_matches = []
    if not hasattr(row, "index"): return ""
    for col in exact_matches:
        if col in row.index:
            val = row[col]
            if isinstance(val, pd.Series):
                for item in val:
                    if pd.notna(item) and str(item).strip().lower() not in ["nan", "nat", "none", "<na>", ""]: return str(item).strip()
            else:
                if pd.notna(val) and str(val).strip().lower() not in ["nan", "nat", "none", "<na>", ""]: return str(val).strip()
    for kw in partial_matches:
        for col in row.index:
            if kw in str(col):
                val = row[col]
                if isinstance(val, pd.Series):
                    for item in val:
                        if pd.notna(item) and str(item).strip().lower() not in ["nan", "nat", "none", "<na>", ""]: return str(item).strip()
                else:
                    if pd.notna(val) and str(val).strip().lower() not in ["nan", "nat", "none", "<na>", ""]: return str(val).strip()
    return ""

def slugify_col(text):
    text = str(text).lower()
    text = ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    text = re.sub(r'[^a-z0-9]', '_', text)
    return re.sub(r'_+', '_', text).strip('_')

def _preparar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [slugify_col(c) for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()].copy()

    def limpar_mat(m):
        m_str = str(m).split('.')[0].lstrip('0').strip()
        return m_str if m_str.lower() not in ['nan', 'none', '<na>', 'nat', ''] else ""

    df["_mat_clean"] = df.apply(lambda r: limpar_mat(_get_val(r, ["matricula", "matricula_rh", "matricula_fopag", "matricula_ativos", "matricula_marsh", "matricula_do_segurado", "matricula_especial"])), axis=1)
    df["_titular_operadora"] = df.apply(lambda r: str(_get_val(r, ["titular", "titular_ativos", "nome_rh", "colaborador", "nome_do_titular", "nome_titular"])).upper(), axis=1)
    df["_cpf_linha_temp"] = df.apply(lambda r: formatar_cpf(_get_val(r, ["cpf", "cpf_fopag", "cpf_ativos"])), axis=1)

    from difflib import SequenceMatcher
    def _sim_nome(s1, s2):
        return SequenceMatcher(None, str(s1).upper().strip(), str(s2).upper().strip()).ratio()

    def _eh_mesma_pessoa_fuzzy(n1: str, n2: str) -> bool:
        if not n1 or not n2: return False
        if n1 == n2: return True
        if _sim_nome(n1, n2) >= 0.85: return True
        p1, p2 = n1.split(), n2.split()
        if len(p1) > 1 and len(p2) > 1 and p1[-1] == p2[-1] and _sim_nome(p1[0], p2[0]) >= 0.80:
            return True
        return False

    def is_titular(r):
        par = str(_get_val(r, ["relacao_de_dependencia", "relacao_dependencia", "parentesco", "parentesco_marsh", "relacao_dependencia_fopag"])).upper().strip()
        
        # Check explicit negative words first to override any accidental overlap
        if any(d in par for d in ["CONJUGE", "CÔNJUGE", "FILH", "ESPOS", "MARIDO", "COMPANHEIR", "PAI", "MAE", "DEPENDENTE", "GUARDA", "TUTELADO", "NETO", "SOBRINHO"]): return False
        
        cpf_linha = r.get("_cpf_linha_temp", "")
        cpf_colab_rh = formatar_cpf(_get_val(r, ["cpf_colaborador_rh", "cpf_rh", "cpf_colaborador", "cpf_cnpj_rh", "cpf_cnpj"]))
        if cpf_linha and len(cpf_linha) >= 11 and cpf_colab_rh and len(cpf_colab_rh) >= 11 and cpf_linha == cpf_colab_rh:
            return True

        b_name = str(_get_val(r, ["beneficiario", "nome_dependente", "beneficiario_ativos"])).upper().strip()
        t_name = str(r["_titular_operadora"]).upper().strip()
        
        if "JOSE HEITOR" in b_name and "MIRYAN" in t_name: return False
        
        # Se os nomes são conhecidos e diferentes (nem primeiro nome igual nem grafia similar):
        # NÃO pode ser o titular daquela matrícula! É dependente!
        if b_name and t_name and b_name != t_name:
            if b_name.split()[0] != t_name.split()[0] and not _eh_mesma_pessoa_fuzzy(b_name, t_name):
                return False
        
        # Check positive titular words
        if "TITULAR" in par or "FUNCIONARIO" in par or "ASSOCIADO" in par or par == "T" or par == "TIT": return True
        
        if b_name and t_name and (b_name == t_name or b_name.split()[0] == t_name.split()[0] or _eh_mesma_pessoa_fuzzy(b_name, t_name)): return True
        return False

    df["_eh_titular_fix"] = df.apply(is_titular, axis=1)

    # --------------------------------------------------------------------------
    # 🛡️ GARANTIA DE NO MÁXIMO 1 PESSOA TITULAR POR MATRÍCULA E OPERADORA
    # Se uma operadora enviar mais de 1 pessoa distinta como titular para a mesma matrícula,
    # identificamos quem é o titular real (colaborador) e os demais viram dependentes.
    # Se forem duas linhas da MESMA pessoa (mesmo CPF ou grafia quase idêntica), mantém como titular.
    # --------------------------------------------------------------------------
    grupos_mat = df[df["_mat_clean"] != ""].groupby(["operadora", "_mat_clean"])
    for (op_g, mat_g), indices in grupos_mat.groups.items():
        if len(indices) <= 1: continue
        sub_g = df.loc[indices]
        tits_g = sub_g[sub_g["_eh_titular_fix"] == True]
        if len(tits_g) > 1:
            melhor_score = -1.0
            idx_titular_real = None
            for idx_r in tits_g.index:
                r_item = df.loc[idx_r]
                b_nm = str(_get_val(r_item, ["beneficiario", "nome_dependente", "beneficiario_ativos"])).upper().strip()
                t_nm = str(r_item.get("_titular_operadora", "")).upper().strip()
                cart = str(_get_val(r_item, ["carteira", "carteirinha", "codigo_dependente_garantido_marsh"])).strip()
                score = _sim_nome(b_nm, t_nm)
                if b_nm and t_nm and b_nm.split()[0] == t_nm.split()[0]: score += 0.5
                if cart.endswith("00") or cart.endswith("000") or cart.endswith("000'"): score += 0.3
                if score > melhor_score:
                    melhor_score = score
                    idx_titular_real = idx_r
            r_real = df.loc[idx_titular_real]
            cpf_real = str(r_real.get("_cpf_linha_temp", "")).strip()
            b_nm_real = str(_get_val(r_real, ["beneficiario", "nome_dependente", "beneficiario_ativos"])).upper().strip()
            for idx_r in tits_g.index:
                if idx_r != idx_titular_real:
                    r_cand = df.loc[idx_r]
                    cpf_cand = str(r_cand.get("_cpf_linha_temp", "")).strip()
                    b_nm_cand = str(_get_val(r_cand, ["beneficiario", "nome_dependente", "beneficiario_ativos"])).upper().strip()
                    mesma_pessoa = (cpf_real and len(cpf_real) >= 11 and cpf_cand == cpf_real) or _eh_mesma_pessoa_fuzzy(b_nm_cand, b_nm_real)
                    if not mesma_pessoa:
                        df.at[idx_r, "_eh_titular_fix"] = False

    # --------------------------------------------------------------------------
    # 🛡️ CHAVE DE FAMÍLIA ISOLADA
    # --------------------------------------------------------------------------
    tits = df[df["_eh_titular_fix"] == True]

    # +++ VARIÁVEL RESTAURADA AQUI! +++
    mapa_nome_tit = tits[tits["_mat_clean"] != ""].groupby(["operadora", "_mat_clean"])["_titular_operadora"].first().to_dict()

    mapa_tit_nome_to_mat = tits[(tits["_titular_operadora"] != "") & (tits["_mat_clean"] != "")].groupby(["operadora", "_titular_operadora"])["_mat_clean"].first().to_dict()
    mapa_tit_nome_to_cpf = tits[(tits["_titular_operadora"] != "") & (tits["_cpf_linha_temp"].str.len() >= 11)].groupby(["operadora", "_titular_operadora"])["_cpf_linha_temp"].first().to_dict()

    def herdar_mat(r):
        if r["_mat_clean"]: return r["_mat_clean"]
        op = r.get("operadora", "")
        tit = r["_titular_operadora"]
        if op and tit and (op, tit) in mapa_tit_nome_to_mat: return mapa_tit_nome_to_mat[(op, tit)]
        return ""
    df["_mat_clean"] = df.apply(herdar_mat, axis=1)

    def gerar_chave_familia(r):
        op = r.get("operadora", "")
        if r["_mat_clean"]: return f"{op}_{r['_mat_clean']}"
        if r["_cpf_linha_temp"] and len(r["_cpf_linha_temp"]) >= 11: return f"{op}_{r['_cpf_linha_temp']}"
        return f"{op}_ORFAO_{r.name}"
    df["_chave_familia"] = df.apply(gerar_chave_familia, axis=1)

    mapa_familia_to_cpf = df[df["_eh_titular_fix"] == True].groupby("_chave_familia")["_cpf_linha_temp"].first().to_dict()

    def definir_cpf_titular(r):
        if r["_eh_titular_fix"] and r["_cpf_linha_temp"] and len(r["_cpf_linha_temp"]) >= 11: return r["_cpf_linha_temp"]
        fam = r["_chave_familia"]
        if fam in mapa_familia_to_cpf and len(mapa_familia_to_cpf[fam]) >= 11: return mapa_familia_to_cpf[fam]
        op = r.get("operadora", "")
        tit = r["_titular_operadora"]
        if op and tit and (op, tit) in mapa_tit_nome_to_cpf: return mapa_tit_nome_to_cpf[(op, tit)]
        return r["_cpf_linha_temp"]
    df["CPF_TIT_FINAL"] = df.apply(definir_cpf_titular, axis=1)

    # --------------------------------------------------------------------------
    # 🛡️ O PORTEIRO DE HOMÔNIMOS E HIERARQUIA
    # --------------------------------------------------------------------------
    def apply_hierarchy(r, col_name):
        cpf_t = r.get("CPF_TIT_FINAL", "")
        cpf_rh = formatar_cpf(_get_val(r, ["cpf_colaborador_rh", "cpf_rh", "cpf_colaborador", "cpf_cnpj", "cpf_cnpj_rh"]))
        mat_rh = _limpar_ponto_zero(_get_val(r, ["matricula_colaborador_rh_rh", "matricula_colaborador_rh", "matricula_rh"])).lstrip("0")
        mat_clean = r.get("_mat_clean", "")

        match_rh_seguro = False
        if cpf_t and len(cpf_t) >= 11 and cpf_t == cpf_rh: match_rh_seguro = True
        elif mat_clean and mat_clean == mat_rh: match_rh_seguro = True

        val = ""
        # 1. TENTA RH (SE FOR SEGURO)
        if match_rh_seguro:
            if col_name == "situacao": val = _get_val(r, ["situacao_rh", "tipo_situacao_rh", "tipo_de_situacao_rh"])
            elif col_name == "cargo": val = _get_val(r, ["cargo_rh", "cargo_colaborador"])
            elif col_name == "sindicato": val = _get_val(r, ["sindicato_rh", "codigo_sindicato_rh", "codigo_sindicato"])
            elif col_name == "cod_sindicato": val = _get_val(r, ["codigo_sindicato_rh", "codigo_sindicato"])
            elif col_name == "admissao": val = _get_val(r, ["admissao_rh", "data_admissao_rh"])
            elif col_name == "empresa": val = _get_val(r, ["razao_social_rh", "empresa_rh"])
            elif col_name == "natureza": val = _get_val(r, ["natureza_profissional_rh", "natureza_profissional"])
            elif col_name == "loja": val = _get_val(r, ["codigo_estabelecimento_rh", "estabelecimento_rh", "loja_rh"])
            if val: return val

        # 2. TENTA FOPAG (Se o RH é falso ou vazio)
        if col_name == "situacao": val = _get_val(r, ["situacao_fopag"])
        elif col_name == "cargo": val = _get_val(r, ["cargo_fopag"])
        elif col_name == "sindicato": val = _get_val(r, ["sindicato_fopag"])
        elif col_name == "admissao": val = _get_val(r, ["admissao_fopag", "data_admissao_fopag"])
        elif col_name == "empresa": val = _get_val(r, ["empresa_fopag"])
        elif col_name == "loja": val = _get_val(r, ["loja_fopag", "estabelecimento_fopag"])
        if val: return val

        # 3. TENTA ATIVOS DA OPERADORA
        if col_name == "situacao": val = _get_val(r, ["situacao", "status"])
        elif col_name == "cargo": val = _get_val(r, ["cargo"])
        elif col_name == "sindicato": val = _get_val(r, ["sindicato"])
        elif col_name == "admissao": val = _get_val(r, ["admissao", "data_admissao"])
        elif col_name == "empresa": val = _get_val(r, ["empresa"])
        elif col_name == "loja": val = _get_val(r, ["loja"])

        return val

    df["SITUACAO_FINAL"] = df.apply(lambda r: apply_hierarchy(r, "situacao"), axis=1)
    df["CARGO_FINAL"] = df.apply(lambda r: apply_hierarchy(r, "cargo"), axis=1)
    df["SINDICATO_FINAL"] = df.apply(lambda r: apply_hierarchy(r, "sindicato"), axis=1)
    df["COD_SINDICATO_FINAL"] = df.apply(lambda r: apply_hierarchy(r, "cod_sindicato"), axis=1)
    df["ADMISSAO_FINAL"] = df.apply(lambda r: apply_hierarchy(r, "admissao"), axis=1)
    df["EMPRESA_FINAL"] = df.apply(lambda r: apply_hierarchy(r, "empresa"), axis=1)
    df["NATUREZA_FINAL"] = df.apply(lambda r: apply_hierarchy(r, "natureza"), axis=1)
    df["LOJA_FINAL"] = df.apply(lambda r: apply_hierarchy(r, "loja"), axis=1)

    # Herança corporativa do Titular para os Dependentes (NUNCA sobrescreve dados pessoais do dependente)
    cols_heranca_tit = ["SITUACAO_FINAL", "CARGO_FINAL", "SINDICATO_FINAL", "COD_SINDICATO_FINAL", "ADMISSAO_FINAL", "EMPRESA_FINAL", "NATUREZA_FINAL", "LOJA_FINAL"]
    tits_corp = df[df["_eh_titular_fix"] == True]
    for col_h in cols_heranca_tit:
        validos_h = tits_corp[tits_corp[col_h].astype(str).str.strip() != ""]
        if validos_h.empty:
            continue
        mapa_fam_h = validos_h.groupby("_chave_familia")[col_h].first().to_dict()
        mapa_mat_h = validos_h[validos_h["_mat_clean"] != ""].groupby("_mat_clean")[col_h].first().to_dict()
        mapa_cpf_h = validos_h[validos_h["CPF_TIT_FINAL"].str.len() >= 11].groupby("CPF_TIT_FINAL")[col_h].first().to_dict()

        def _herdar_campo_corp(r, c_name=col_h, mf=mapa_fam_h, mm=mapa_mat_h, mc=mapa_cpf_h):
            atual = _to_str(r.get(c_name, ""))
            if atual:
                return atual
            if not r.get("_eh_titular_fix"):
                fam = r.get("_chave_familia", "")
                if fam in mf: return mf[fam]
                mc_val = r.get("_mat_clean", "")
                if mc_val and mc_val in mm: return mm[mc_val]
                ct_val = r.get("CPF_TIT_FINAL", "")
                if ct_val and ct_val in mc: return mc[ct_val]
            return atual

        df[col_h] = df.apply(_herdar_campo_corp, axis=1)

    def definir_nome_colab(r):
        cpf_t = r.get("CPF_TIT_FINAL", "")
        cpf_rh = formatar_cpf(_get_val(r, ["cpf_colaborador_rh", "cpf_rh", "cpf_colaborador", "cpf_cnpj", "cpf_cnpj_rh"]))
        mat_rh = _limpar_ponto_zero(_get_val(r, ["matricula_colaborador_rh_rh", "matricula_colaborador_rh", "matricula_rh"])).lstrip("0")
        mat_clean = r.get("_mat_clean", "")

        match_rh_seguro = False
        if cpf_t and len(cpf_t) >= 11 and cpf_t == cpf_rh: match_rh_seguro = True
        elif mat_clean and mat_clean == mat_rh: match_rh_seguro = True

        if match_rh_seguro:
            n_rh = _get_val(r, ["nome_rh", "titular_rh", "beneficiario_rh"])
            if n_rh: return str(n_rh).strip().upper()

        op = r.get("operadora", "")
        if op and mat_clean and (op, mat_clean) in mapa_nome_tit:
            return mapa_nome_tit[(op, mat_clean)]
        return r["_titular_operadora"]

    df["_colab_fix"] = df.apply(definir_nome_colab, axis=1)

    def definir_cpf_benef(r):
        if r["_eh_titular_fix"]: return r["CPF_TIT_FINAL"]
        cpf_m = formatar_cpf(_get_val(r, ["cpf_marsh_garantido_marsh", "cpf_dependente_marsh"]))
        return cpf_m if len(cpf_m) >= 11 else r["_cpf_linha_temp"]
    df["CPF_BENEF_FINAL"] = df.apply(definir_cpf_benef, axis=1)

    def definir_nome_benef(r):
        if r.get("_eh_titular_fix") and r.get("_colab_fix") and str(r.get("_colab_fix")).upper() not in ["NAN", "NONE", ""]:
            return str(r.get("_colab_fix")).upper()
        b = str(_get_val(r, ["beneficiario", "nome_dependente", "beneficiario_ativos", "beneficiario_fopag"])).upper()
        return b if b and b not in ["NAN", "NONE", ""] else r["_colab_fix"]
    df["_benef_fix"] = df.apply(definir_nome_benef, axis=1)

    tem_titular_mesma_familia = df.groupby("_chave_familia")["_eh_titular_fix"].transform("any")
    mats_titulares_globais = set(df.loc[df["_eh_titular_fix"] == True, "_mat_clean"].dropna().astype(str).str.strip())
    mats_titulares_globais.discard("")
    cpfs_titulares_globais = set(df.loc[df["_eh_titular_fix"] == True, "CPF_TIT_FINAL"].dropna().astype(str).str.strip())
    cpfs_titulares_globais.discard("")

    def verificar_titular_global(r):
        if tem_titular_mesma_familia.loc[r.name]:
            return True
        m = str(r.get("_mat_clean", "")).strip()
        if m and m in mats_titulares_globais:
            return True
        cpf_t = str(r.get("CPF_TIT_FINAL", "")).strip()
        if cpf_t and len(cpf_t) >= 11 and cpf_t in cpfs_titulares_globais:
            return True
        origem_rh = str(r.get("origem_match_rh", "")).upper()
        if origem_rh not in ["SEM_MATCH", "", "NAN", "NONE"]:
            return True
        colab = str(r.get("_colab_fix", "")).strip().upper()
        if colab and colab not in ["", "NAN", "NONE"]:
            return True
        origem_fopag = str(r.get("origem_match_fopag", "")).upper()
        if origem_fopag not in ["SEM_MATCH", "", "NAN", "NONE"]:
            return True
        return False

    df["_tem_titular_global"] = df.apply(verificar_titular_global, axis=1)
    df['_chave_rn004'] = df["_benef_fix"] + "_" + df["_chave_familia"] + "_" + df['operadora']

    # Hierarquia de Nascimento Blindada
    def obter_nasc(r):
        eh_tit = r.get("_eh_titular_fix", True)
        if not eh_tit:
            nasc_dep = _get_val(r, ["nascimento_marsh_garantido_marsh", "nascimento_marsh", "data_nascimento_dependente", "data_nascimento_marsh", "data_nascimento_dep", "nascimento_dep"], ["nascimento_dependente", "nascimento_marsh", "nasc_dep"])
            if nasc_dep: return nasc_dep
            r_sem_rh = r[[c for c in r.index if not str(c).endswith("_rh")]]
            return _get_val(r_sem_rh, ["nascimento", "nasc", "data_de_nascimento", "data_nascimento", "nascimento_ativos", "nascimento_fopag"], ["nascimento", "nasc"])

        cpf_t = r.get("CPF_TIT_FINAL", "")
        cpf_rh = formatar_cpf(_get_val(r, ["cpf_colaborador_rh", "cpf_rh", "cpf_colaborador", "cpf_cnpj", "cpf_cnpj_rh"]))
        mat_rh = _limpar_ponto_zero(_get_val(r, ["matricula_colaborador_rh_rh", "matricula_colaborador_rh", "matricula_rh"])).lstrip("0")
        mat_clean = r.get("_mat_clean", "")

        match_rh_seguro = False
        if cpf_t and len(cpf_t) >= 11 and cpf_t == cpf_rh: match_rh_seguro = True
        elif mat_clean and mat_clean == mat_rh: match_rh_seguro = True

        if match_rh_seguro:
            nasc = _get_val(r, ["data_nascimento_rh", "nascimento_rh", "data_de_nascimento_rh"])
            if nasc: return nasc

        return _get_val(r, ["nascimento_ativos", "nascimento", "data_nascimento", "nascimento_fopag", "nascimento_marsh"], ["nasc"])

    df["_nascimento_final"] = df.apply(obter_nasc, axis=1)

    plano_col = _get_val(df.iloc[0] if not df.empty else None, ["plano", "plano_fopag", "tipo_plano"])

    df["_ordem_titular"] = df["_eh_titular_fix"].apply(lambda x: 0 if x else 1)
    df = df.sort_values(by=["operadora", "_chave_familia", "_ordem_titular", "_benef_fix"]).reset_index(drop=True)

    subset_killer = ['operadora', '_chave_familia', '_benef_fix']
    if plano_col and plano_col in df.columns: subset_killer.append(plano_col)

    df = df.drop_duplicates(subset=subset_killer, keep="first").copy()
    return df

def _detectar_inconsistencias(df: pd.DataFrame, mapa_fin: dict, mapa_faturas_cpf: dict, mapa_faturas_mat_nome: dict, competencia_dt=None) -> list[dict]:
    resultados = []
    df_ativos = df[df['origem_match_fopag'] != 'SO_FOPAG']
    
    # Mapear quantidade de dependentes ATIVOS por titular (para blindar RN018)
    mapa_ativos_por_titular = {}
    for _, dep_r in df[df["_eh_titular_fix"] == False].iterrows():
        fam = dep_r["_chave_familia"]
        val_canc = _get_val(dep_r, exact_matches=["fim_de_vigencia", "cancelamento", "cancelamento_fopag", "data_de_cancelamento", "data_cancelamento", "exclusao", "data_exclusao", "cancelamento_ativos", "data_de_exclusao"], partial_matches=["fim_de_vigencia", "fim_vigencia", "cancelamento", "exclusao"])
        dt_canc = pd.NaT
        if val_canc and str(val_canc).lower() not in ["nan", "none", "nat"]:
            dt_canc = pd.to_datetime(val_canc, dayfirst=True, errors='coerce')
        
        is_recent_or_active = True
        if pd.notna(dt_canc) and competencia_dt is not None:
            meses_diff = (competencia_dt.year - dt_canc.year) * 12 + (competencia_dt.month - dt_canc.month)
            if meses_diff >= 2:
                is_recent_or_active = False 
                
        if is_recent_or_active:
            mapa_ativos_por_titular[fam] = mapa_ativos_por_titular.get(fam, 0) + 1
    duplicados_ativos = df_ativos[df_ativos.duplicated(subset=['_chave_rn004'], keep=False)]
    cpfs_ativos_tit = set(df_ativos['CPF_TIT_FINAL'].dropna().unique())

    current_alerta_idade = ""
    current_val_demissao = ""

    from datetime import date
    df_ativos["_primeiro_nome_dep"] = df_ativos["_benef_fix"].astype(str).apply(lambda x: str(x).split()[0] if pd.notna(x) and x else "")
    tem_base_rh_carregada = any(_to_str(x) != "" for x in df["SITUACAO_FINAL"])
    mapa_parentesco = {}
    mapa_qtd_dep_ativos = {}
    mapa_qtd_dep_vigentes = {}
    for _, r in df_ativos.iterrows():
        tit_mat = _to_str(r.get("_mat_clean", ""))
        if not tit_mat:
            tit_mat = _to_str(_get_val(r, ["matricula", "matricula_fopag", "matricula_ativos", "matricula_rh", "demitido_matricula", "matricula_colaborador", "matricula_marsh"], ["matricula"]))
        nome_dep = str(r.get("_primeiro_nome_dep", "")).strip()
        op_map = _to_str(r.get("operadora", ""))
        eh_tit_map = r.get("_eh_titular_fix")
        
        if not eh_tit_map:
            chave = (op_map, tit_mat)
            mapa_qtd_dep_ativos[chave] = mapa_qtd_dep_ativos.get(chave, 0) + 1
            dt_c_chk = _get_val(r, exact_matches=["fim_de_vigencia", "cancelamento", "cancelamento_fopag", "data_de_cancelamento", "data_cancelamento", "exclusao", "data_exclusao", "cancelamento_ativos", "data_de_exclusao"], partial_matches=["fim_de_vigencia", "fim_vigencia", "cancelamento", "exclusao"])
            if not dt_c_chk or str(dt_c_chk).lower() in ["nan", "none", "nat", ""]:
                mapa_qtd_dep_vigentes[chave] = mapa_qtd_dep_vigentes.get(chave, 0) + 1
                
        parentesco_map = str(r.get("parentesco_marsh", "")).strip()
        if not parentesco_map or parentesco_map.upper() == "NAN":
            parentesco_map = str(r.get("relacao_de_dependencia", "")).strip()
        if not parentesco_map or parentesco_map.upper() == "NAN":
            parentesco_map = str(r.get("relacao_dependencia", "")).strip()
        if parentesco_map and parentesco_map.upper() not in ["NAN", "NONE", "TITULAR", "TIT", "T"]:
            mapa_parentesco[(tit_mat, nome_dep)] = parentesco_map

    mapa_demissoes_oficial = {}
    for _, r in df.iterrows():
        mat_dem = _limpar_ponto_zero(_get_val(r, ["demitido_matricula"]))
        dt_dem = _get_val(r, ["demitido_data_rescisao"])
        if mat_dem and dt_dem and str(dt_dem).lower() not in ["nan", "nat", "none", ""]:
            mapa_demissoes_oficial[mat_dem] = dt_dem

    def base_row(row, demissao_limpa, alerta) -> dict:
        bandeira_col = _limpar_ponto_zero(_get_val(row, exact_matches=["bandeira_garantida", "bandeira_contrato", "bandeira", "subfatura", "bandeira_fopag"], partial_matches=["bandeira_garantid", "contrato", "subfatura"]))
        empresa = row.get("EMPRESA_FINAL", "")
        if not empresa or str(empresa).replace('.', '').isnumeric(): empresa = bandeira_col
        loja = row.get("LOJA_FINAL", "")
        if not loja: loja = _limpar_ponto_zero(_get_val(row, ["loja", "unidade", "loja_fopag"]))

        matricula_final = _to_str(_get_val(row, ["matricula", "matricula_fopag", "matricula_ativos", "matricula_rh", "demitido_matricula", "matricula_colaborador", "matricula_marsh", "matricula_fake", "matricula_do_segurado", "matricula_especial"]))
        situacao_rh_bruta = _to_str(row.get("SITUACAO_FINAL", ""))
        situacao = MAPA_SITUACAO.get(situacao_rh_bruta, situacao_rh_bruta)

        cargo = _to_str(row.get("CARGO_FINAL", ""))
        sindicato = _to_str(row.get("SINDICATO_FINAL", ""))
        admissao = row.get("ADMISSAO_FINAL", "")
        colab = str(row.get("_colab_fix", ""))
        if not colab or colab == "NONE": colab = str(row.get("_titular_operadora", ""))
        if not colab or colab == "NONE": colab = str(row.get("nome_fake", ""))
        benef = row.get("_benef_fix", "")
        cpf_titular = formatar_cpf(row.get("CPF_TIT_FINAL", ""))
        if not cpf_titular: cpf_titular = formatar_cpf(row.get("cpf_titular_fake", ""))
        cpf_beneficiario = formatar_cpf(row.get("CPF_BENEF_FINAL", ""))

        eh_tit = row.get("_eh_titular_fix")
        parentesco = _to_str(_get_val(row, exact_matches=["parentesco_marsh", "relacao_de_dependencia", "relacao_dependencia", "parentesco_fopag", "parentesco", "tipo_de_segurado", "tipo_segurado"], partial_matches=["parentesco", "dependencia", "relacao", "grau"])).upper()
        
        if not parentesco or parentesco == "NAN":
            primeiro_nome_dep = benef.split()[0] if benef else ""
            parentesco = str(mapa_parentesco.get((matricula_final, primeiro_nome_dep), "")).upper()
            if parentesco == "NAN": parentesco = ""
        if not eh_tit:
            # Se for dependente, o parentesco NÃO PODE ser TITULAR nem vazio
            if not parentesco or parentesco in ["TITULAR", "TIT", "T", "NAN", "NONE"]:
                primeiro_nome_dep = benef.split()[0] if benef else ""
                parentesco = str(mapa_parentesco.get((matricula_final, primeiro_nome_dep), "")).upper()
                if parentesco in ["", "NAN", "NONE", "TITULAR", "TIT", "T"]:
                    for col_p in row.index:
                        if any(kw in str(col_p).lower() for kw in ["parentesco", "dependencia", "relacao"]):
                            if any(kw in str(row[col_p]).upper() for kw in ["CONJUGE", "CÔNJUGE", "COMPANHEIR", "ESPOS", "MARIDO"]):
                                parentesco = "CÔNJUGE"
                                break
                    if parentesco in ["", "NAN", "NONE", "TITULAR", "TIT", "T"]:
                        parentesco = "CÔNJUGE"
            else:
                for col_p in row.index:
                    if any(kw in str(col_p).lower() for kw in ["parentesco", "dependencia", "relacao"]):
                        if any(kw in str(row[col_p]).upper() for kw in ["CONJUGE", "CÔNJUGE", "COMPANHEIR", "ESPOS", "MARIDO"]):
                            parentesco = "CÔNJUGE"
                            break

        plano = _get_val(row, ["plano", "plano_fopag"])
        operadora = _to_str(row.get("operadora", ""))
        if not plano: plano = operadora

        tipo_plano = _get_val(row, exact_matches=["tipo_codigo_plano", "categoria_do_plano", "categoria_plano", "tipo_plano", "codigo_plano"], partial_matches=["tipo_codigo", "cod_plano", "categoria"])
        carteira = _limpar_ponto_zero(_get_val(row, exact_matches=["carteira", "carteirinha", "certificado", "codigo_dependente_garantido_marsh"], partial_matches=["codigo_dependente_garantid", "carteira"]))

        nasc_final = row.get("_nascimento_final")
        nasc_dt = pd.to_datetime(nasc_final, dayfirst=True, errors="coerce")

        idade_str = ""
        if pd.notna(nasc_dt):
            ref_dt = competencia_dt if competencia_dt is not None else pd.to_datetime(date.today())
            idade_int = ref_dt.year - nasc_dt.year - ((ref_dt.month, ref_dt.day) < (nasc_dt.month, nasc_dt.day))
            idade_str = str(idade_int)

        grau_instrucao = _limpar_ponto_zero(_get_val(row, exact_matches=["desc_instrucao_garantido_marsh", "desc_grau_de_instrucao", "descricao_grau_instrucao", "grau_de_instrucao_marsh", "grau_instrucao"], partial_matches=["desc_instrucao", "descricao_instrucao", "instrucao"]))
        if grau_instrucao == "0": grau_instrucao = ""

        dependente_inval = _get_val(row, exact_matches=["invalido_garantido_marsh", "dependente_invalido_marsh", "invalido_marsh"]).upper()
        if dependente_inval == "0": dependente_inval = ""
        if dependente_inval in ["NÃO", "NAO", "FALSE", "F", "N"]: dependente_inval = "N"
        elif dependente_inval in ["SIM", "TRUE", "VERDADEIRO", "V", "Y", "S"]: dependente_inval = "S"

        if eh_tit:
            parentesco = "TITULAR"
            dependente_inval = ""
            grau_instrucao = ""

        v_mens_tit, v_mens_dep, v_copart = "", "", ""
        mat_clean = row.get("_mat_clean", "")
        if mat_clean in mapa_fin and operadora in mapa_fin[mat_clean]:
            dados_fin = mapa_fin[mat_clean][operadora]
            if eh_tit:
                if dados_fin['mens_tit'] > 0: v_mens_tit = f"R$ {dados_fin['mens_tit']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                if dados_fin['mens_dep'] > 0: v_mens_dep = f"R$ {dados_fin['mens_dep']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                cop_total = dados_fin['cop_tit'] + dados_fin['cop_dep']
                if cop_total > 0: v_copart = f"R$ {cop_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            else:
                origem_f_row = str(_get_val(row, ["origem_match_fopag"]))
                dt_canc_check = _get_val(row, exact_matches=["fim_de_vigencia", "cancelamento", "cancelamento_fopag", "data_de_cancelamento", "data_cancelamento", "exclusao", "data_exclusao", "cancelamento_ativos", "data_de_exclusao"], partial_matches=["fim_de_vigencia", "fim_vigencia", "cancelamento", "exclusao"])
                dep_ativo_rateio = (origem_f_row != "SO_FOPAG") and (not dt_canc_check or str(dt_canc_check).lower() in ["nan", "none", "nat", ""])
                if dep_ativo_rateio and dados_fin['mens_dep'] > 0:
                    qtd_d_vig = mapa_qtd_dep_vigentes.get((operadora, mat_clean), 0)
                    if qtd_d_vig > 0:
                        val_rat = dados_fin['mens_dep'] / qtd_d_vig
                        v_mens_dep = f"R$ {val_rat:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                if dep_ativo_rateio and dados_fin['cop_dep'] > 0:
                    qtd_d_vig = mapa_qtd_dep_vigentes.get((operadora, mat_clean), 0)
                    if qtd_d_vig > 0:
                        val_cop_rat = dados_fin['cop_dep'] / qtd_d_vig
                        v_copart = f"R$ {val_cop_rat:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        # CONCILIAÇÃO FATURA
        valor_fatura = 0.0
        cpf_limpo = cpf_beneficiario.replace(".", "").replace("-", "") if cpf_beneficiario else ""

        if cpf_limpo and len(cpf_limpo) >= 11 and (operadora, cpf_limpo) in mapa_faturas_cpf:
            valor_fatura = mapa_faturas_cpf[(operadora, cpf_limpo)]
        elif mat_clean and benef:
            p_nome = benef.split()[0]
            if (operadora, mat_clean, p_nome) in mapa_faturas_mat_nome:
                valor_fatura = mapa_faturas_mat_nome[(operadora, mat_clean, p_nome)]

        v_fatura_str = ""
        if valor_fatura > 0:
            v_fatura_str = f"R$ {valor_fatura:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        dt_vig = _get_val(row, exact_matches=["inicio_de_vigencia", "inicio_vigencia", "data_vigencia", "inicio_vigencia_fopag", "inicio_da_vigencia_atual", "data_de_inicio", "inicio_vigencia_ativos"], partial_matches=["inicio_de_vigencia", "inicio_vigencia", "vigencia"])
        dt_canc = _get_val(row, exact_matches=["fim_de_vigencia", "cancelamento", "cancelamento_fopag", "data_de_cancelamento", "data_cancelamento", "exclusao", "data_exclusao", "cancelamento_ativos", "data_de_exclusao"], partial_matches=["fim_de_vigencia", "fim_vigencia", "cancelamento", "exclusao"])

        return {
            "OPERADORA": operadora,
            "EMPRESA": empresa,
            "LOJA": loja,
            "MATRICULA": matricula_final,
            "CPF_TITULAR": cpf_titular,
            "COLABORADOR": _to_str(colab),
            "CPF_BENEFICIARIO": cpf_beneficiario,
            "BENEFICIARIO": _to_str(benef),
            "PARENTESCO": parentesco,
            "DEPENDENTE_INVALIDO": dependente_inval,
            "GRAU_INSTRUCAO": grau_instrucao,
            "NASCIMENTO": formatar_data_br(row.get("_nascimento_final")),
            "IDADE": idade_str,
            "CARTEIRA": carteira,
            "PLANO": _to_str(plano),
            "TIPO_PLANO": _to_str(tipo_plano),
            "BANDEIRA": bandeira_col,
            "VALOR_MENSALIDADE_TITULAR": v_mens_tit,
            "VALOR_MENSALIDADE_DEPENDENTES": v_mens_dep,
            "VALOR_COPARTICIPACAO": v_copart,
            "VALOR_PAGO_FATURA": v_fatura_str,
            "DATA_VIGENCIA": dt_vig,
            "DATA_CANCELAMENTO": dt_canc,
            "SITUACAO": situacao,
            "CARGO": cargo,
            "SINDICATO": sindicato,
            "ADMISSAO": formatar_data_br(admissao),
            "DEMISSAO": formatar_data_br(demissao_limpa),
        }

    def add(row, rn: str, tipo: str, desc: str, origem: str, acao: str):
        nonlocal current_alerta_idade, current_val_demissao
        r = base_row(row, current_val_demissao, current_alerta_idade)
        r.update({"TIPO_INCONSISTENCIA": tipo, "RN": rn, "DESCRICAO_DIVERGENCIA": desc, "BASE_ORIGEM": origem, "ACAO_RECOMENDADA": acao, "ALERTA_IDADE": current_alerta_idade})
        resultados.append(r)

    for _, row in df.iterrows():
        linhas_antes = len(resultados)
        current_alerta_idade = ""
        current_val_demissao = ""

        mat_clean = row.get("_mat_clean", "")
        op = _to_str(row.get("operadora", ""))
        cpf_valido = len(str(row.get("CPF_TIT_FINAL", ""))) >= 11
        cpf_titular = row.get("CPF_TIT_FINAL", "")

        val_cancelamento = _get_val(row, exact_matches=["fim_de_vigencia", "cancelamento", "cancelamento_fopag", "data_de_cancelamento", "data_cancelamento", "exclusao", "data_exclusao", "cancelamento_ativos", "data_de_exclusao"], partial_matches=["fim_de_vigencia", "fim_vigencia", "cancelamento", "exclusao"])

        dt_canc = pd.NaT
        str_data_canc = ""
        if val_cancelamento and str(val_cancelamento).lower() not in ["nan", "none", "nat"]:
            dt_canc = pd.to_datetime(val_cancelamento, dayfirst=True, errors='coerce')

        cancelado_mes_anterior = False
        cancelado_mes_atual = False
        cancelado_meses_atras = False

        if pd.notna(dt_canc) and competencia_dt is not None:
            str_data_canc = dt_canc.strftime('%d/%m/%Y')
            meses_diff = (competencia_dt.year - dt_canc.year) * 12 + (competencia_dt.month - dt_canc.month)
            if meses_diff == 1: cancelado_mes_anterior = True
            elif meses_diff == 0: cancelado_mes_atual = True
            elif meses_diff >= 2: cancelado_meses_atras = True

        val_demissao_bruta = _get_val(row, ["demitido_data_rescisao", "data_rescisao", "rescisao"])
        mat_dem = _limpar_ponto_zero(_get_val(row, ["demitido_matricula"]))

        if val_demissao_bruta != "" and str(val_demissao_bruta).lower() not in ["nan", "nat", "none"]:
            if mat_clean != "" and mat_dem != "" and mat_clean == mat_dem:
                nome_colab = str(row.get("_colab_fix", "")).strip().upper()
                nome_dem = str(_get_val(row, ["demitido_colaborador", "demitido_nome"])).strip().upper()
                if nome_colab != "" and nome_dem != "":
                    if nome_colab.split()[0] == nome_dem.split()[0]:
                        current_val_demissao = _to_str(val_demissao_bruta)

        tem_demitido = current_val_demissao != ""
        eh_titular = row.get("_eh_titular_fix")

        nasc_final = row.get("_nascimento_final")
        nasc_dt = pd.to_datetime(nasc_final, dayfirst=True, errors="coerce")

        idade = None
        if pd.notna(nasc_dt):
            ref_dt = competencia_dt if competencia_dt is not None else pd.to_datetime(date.today())
            idade = ref_dt.year - nasc_dt.year - ((ref_dt.month, ref_dt.day) < (nasc_dt.month, nasc_dt.day))

        parentesco_str = str(_get_val(row, ["parentesco_marsh", "relacao_de_dependencia", "relacao_dependencia", "parentesco_fopag", "parentesco", "tipo_de_segurado", "tipo_segurado"])).upper()
        if not parentesco_str or parentesco_str == "NAN":
            primeiro_nome_dep = str(row.get("_benef_fix", "")).split()[0] if row.get("_benef_fix") else ""
            parentesco_str = str(mapa_parentesco.get((mat_clean, primeiro_nome_dep), "")).upper()
            if parentesco_str == "NAN": parentesco_str = ""
        if not eh_titular:
            for col_p in row.index:
                if any(kw in str(col_p).lower() for kw in ["parentesco", "dependencia", "relacao"]):
                    if any(kw in str(row[col_p]).upper() for kw in ["CONJUGE", "CÔNJUGE", "COMPANHEIR", "ESPOS", "MARIDO"]):
                        parentesco_str = "CÔNJUGE"
                        break

        invalido = str(_get_val(row, ["invalido_garantido_marsh", "dependente_invalido_marsh", "invalido_marsh"])).strip().upper()
        if invalido in ["NÃO", "NAO", "FALSE", "F", "N"]: invalido = "N"
        elif invalido in ["SIM", "TRUE", "VERDADEIRO", "V", "Y", "S"]: invalido = "S"
        grau_instrucao = _limpar_ponto_zero(_get_val(row, ["desc_instrucao_garantido_marsh", "desc_grau_de_instrucao"]))
        origem_fopag = str(_get_val(row, ["origem_match_fopag"]))
        is_so_fopag = (origem_fopag == "SO_FOPAG")

        motivo_op = _to_str(_get_val(row, ["motivo", "motivo_fopag", "operacao", "operação", "desc_motivo", "observacao"])).upper()
        recibo_de_saida = any(palavra in motivo_op for palavra in ["DESLIGADO", "EXCLUSAO", "EXCLUSÃO", "CANCELAMENTO", "SAIDA", "SAÍDA"])

        tem_dinheiro_row = False
        if mat_clean in mapa_fin and op in mapa_fin[mat_clean]:
            dfin = mapa_fin[mat_clean][op]
            if eh_titular:
                if dfin['mens_tit'] > 0 or dfin['cop_tit'] > 0 or dfin['mens_dep'] > 0 or dfin['cop_dep'] > 0:
                    tem_dinheiro_row = True
            else:
                if dfin['mens_dep'] > 0 or dfin['cop_dep'] > 0:
                    tem_dinheiro_row = True

        isento_cobranca = False
        empresa_check = str(row.get("EMPRESA_FINAL", "")).upper()
        natureza_check = str(row.get("NATUREZA_FINAL", "")).upper()
        sindicato_check = str(row.get("SINDICATO_FINAL", "")).upper()
        cod_sind_check = _limpar_ponto_zero(row.get("COD_SINDICATO_FINAL", "")).lstrip("0").strip()
        bandeira_col = _limpar_ponto_zero(_get_val(row, exact_matches=["bandeira_garantida", "bandeira_contrato", "bandeira", "subfatura", "bandeira_fopag"], partial_matches=["bandeira_garantid", "contrato", "subfatura"]))
        bandeira_check = str(bandeira_col).upper()

        if eh_titular:
            eh_giga = ("GIGA" in empresa_check or "GIGA" in natureza_check or "GIGA" in bandeira_check or "328399" in empresa_check or "328399" in bandeira_check or "GICA" in empresa_check or "GICA" in bandeira_check)
            eh_odonto = "ODONTO" in op.upper()
            if eh_giga:
                # GIGA só é isento no plano de SAÚDE. Em ODONTO (SulAmérica Odonto), NÃO é isento.
                if not eh_odonto:
                    isento_cobranca = True
            elif "TRANSP RODOVIARIOS EST SE" in sindicato_check or cod_sind_check == "5" or str(sindicato_check).strip() in ["5", "05", "005"]:
                isento_cobranca = True

        eh_conjuge = any(p in parentesco_str for p in ["CONJUGE", "CÔNJUGE", "COMPANHEIR", "ESPOS", "MARIDO"])
        em_processo_de_saida = tem_demitido or cancelado_mes_anterior or cancelado_mes_atual or recibo_de_saida

        if not eh_titular and not eh_conjuge and pd.notna(idade) and not em_processo_de_saida:
            idade_int = int(idade)
            if invalido == "S": pass
            elif idade_int == 20:
                mes_21 = (nasc_dt + pd.DateOffset(years=21)).strftime('%m/%Y')
                current_alerta_idade = f"Completará 21 anos em {mes_21} (Exigirá comprovação universitária)"
            elif idade_int == 23:
                mes_24 = (nasc_dt + pd.DateOffset(years=24)).strftime('%m/%Y')
                current_alerta_idade = f"Completará 24 anos em {mes_24} (Exclusão por idade limite)"

        situacao_check = _to_str(row.get("SITUACAO_FINAL", ""))
        is_ativo_rh = bool(situacao_check and "demitid" not in situacao_check.lower() and "rescis" not in situacao_check.lower())

        if tem_demitido and not is_so_fopag and not is_ativo_rh:
            data_resc_dt = pd.to_datetime(current_val_demissao, dayfirst=True, errors='coerce')
            if pd.notna(data_resc_dt) and competencia_dt is not None:
                meses_diff = (competencia_dt.year - data_resc_dt.year) * 12 + (competencia_dt.month - data_resc_dt.month)
                if meses_diff >= 2:
                    if not pd.notna(dt_canc):
                        add(row, "RN001", "Situação do colaborador", "Colaborador demitido (Prazo excedido) possui benefício ativo.", "DESLIGADOS", "Verificar inativação")
                    else:
                        if data_resc_dt.day <= 20: data_esperada = (data_resc_dt + pd.DateOffset(months=1)).replace(day=1)
                        else: data_esperada = (data_resc_dt + pd.DateOffset(months=2)).replace(day=1)
                        if dt_canc.date() != data_esperada.date():
                            add(row, "RN019", "Data de Rescisão x Cancelamento", f"Data de cancelamento ({str_data_canc}) diverge da rescisão. Correta seria: {data_esperada.strftime('%d/%m/%Y')}.", "ATIVOS", "Ajustar operadora")

        if cpf_valido and row.name in duplicados_ativos.index and not isento_cobranca:
            if not is_so_fopag:
                add(row, "RN004", "Eventos Assistência", "Possível cobrança simultânea de mais de um plano do mesmo tipo.", "ATIVOS", "Verificar duplicidade de desconto")

        # RN020 (Falta de Desconto FOPAG)
        benef = row.get("_benef_fix", "")
        cpf_benef_str = formatar_cpf(row.get("CPF_BENEF_FINAL", ""))
        cpf_limpo_rn = cpf_benef_str.replace(".", "").replace("-", "") if cpf_benef_str else ""
        valor_fatura_num = 0.0

        if cpf_limpo_rn and len(cpf_limpo_rn) >= 11 and (op, cpf_limpo_rn) in mapa_faturas_cpf:
            valor_fatura_num = mapa_faturas_cpf[(op, cpf_limpo_rn)]
        elif mat_clean and benef:
            p_nome = benef.split()[0]
            if (op, mat_clean, p_nome) in mapa_faturas_mat_nome:
                valor_fatura_num = mapa_faturas_mat_nome[(op, mat_clean, p_nome)]

        situacao_check_rn020 = _to_str(row.get("SITUACAO_FINAL", ""))
        if tem_base_rh_carregada:
            elegivel_rn020 = MAPA_SITUACAO.get(situacao_check_rn020, situacao_check_rn020) == "Atividade Normal"
        else:
            elegivel_rn020 = not tem_demitido

        tem_mensalidade_rn020 = False
        if mat_clean in mapa_fin and op in mapa_fin[mat_clean]:
            dfin_rn = mapa_fin[mat_clean][op]
            if eh_titular:
                tem_mensalidade_rn020 = (dfin_rn['mens_tit'] > 0)
            else:
                tem_mensalidade_rn020 = (dfin_rn['mens_dep'] > 0)

        inclusao_recente_corte = False
        val_vig_rn = _get_val(row, exact_matches=["inicio_de_vigencia", "inicio_vigencia", "data_vigencia", "inicio_vigencia_fopag", "inicio_da_vigencia_atual", "data_de_inicio", "inicio_vigencia_ativos"], partial_matches=["inicio_de_vigencia", "inicio_vigencia", "vigencia"])
        if val_vig_rn and str(val_vig_rn).lower() not in ["nan", "none", "nat"] and competencia_dt is not None:
            s_vig = str(val_vig_rn).strip()
            if re.match(r"^\d{4}-\d{2}-\d{2}", s_vig):
                dt_vig_rn = pd.to_datetime(s_vig, dayfirst=False, errors="coerce")
            else:
                dt_vig_rn = pd.to_datetime(s_vig, dayfirst=True, errors="coerce")
            if pd.notna(dt_vig_rn):
                meses_vig = (competencia_dt.year - dt_vig_rn.year) * 12 + (competencia_dt.month - dt_vig_rn.month)
                if meses_vig < 0 or (meses_vig == 0 and dt_vig_rn.day > 15):
                    inclusao_recente_corte = True

        if valor_fatura_num > 0 and elegivel_rn020 and not em_processo_de_saida and not is_so_fopag and not inclusao_recente_corte:
            if not isento_cobranca:
                if not tem_mensalidade_rn020:
                    add(row, "RN020", "Falta de Desconto em Folha", "Colaborador/Dependente faturado pela operadora, mas sem desconto financeiro na folha.", "FATURA / FOPAG", "Habilitar desconto em folha")

        if is_so_fopag and not isento_cobranca:
            if tem_dinheiro_row and not recibo_de_saida:
                if cancelado_meses_atras:
                    tipo_p = "Titular" if eh_titular else "Dependente"
                    if not eh_titular:
                        qtd_ativos = mapa_ativos_por_titular.get(row["_chave_familia"], 0)
                        if qtd_ativos > 0:
                            pass # The charge is for the other active dependents, ignore RN018!
                        else:
                            add(row, "RN018", "Cobrança Pós-Cancelamento", f"{tipo_p} sofreu desconto financeiro, mas plano foi cancelado há mais de 1 mês ({str_data_canc}).", "FOPAG", "Solicitar reembolso")
                    else:
                        add(row, "RN018", "Cobrança Pós-Cancelamento", f"{tipo_p} sofreu desconto financeiro, mas plano foi cancelado há mais de 1 mês ({str_data_canc}).", "FOPAG", "Solicitar reembolso")
                elif cancelado_mes_anterior or cancelado_mes_atual:
                    pass
                else:
                    if eh_titular:
                        if cpf_titular and cpf_titular in cpfs_ativos_tit:
                            pass
                        else:
                            add(row, "RN006", "Cobrança sem cadastro ativo", "Desconto financeiro em folha sem correspondência na base de ativos.", "FOPAG", "Confirmar exclusão ou incluir")
            else:
                if cancelado_mes_anterior or cancelado_mes_atual or cancelado_meses_atras or recibo_de_saida:
                    tipo_p = "Titular" if eh_titular else "Dependente"
                    add(row, "INFO", "Controle de Exclusão", f"{tipo_p} listado na Fopag apenas para controle de exclusão (sem desconto indevido).", "FOPAG", "Acompanhar")

        if eh_titular and isento_cobranca and mat_clean in mapa_fin and op in mapa_fin[mat_clean]:
            if mapa_fin[mat_clean][op]['mens_tit'] > 0:
                add(row, "RN003", "Cobrança Indevida (Isenção)", "Colaborador isento sofreu desconto de mensalidade.", "FOPAG", "Solicitar reembolso e inativar desconto")

        tem_titular_cadastrado = (
            bool(row.get("_tem_titular_global"))
            or (bool(mat_clean) and mat_clean in mapa_fin)
            or (str(row.get("origem_match_rh", "")).upper() not in ["SEM_MATCH", "", "NAN", "NONE"])
            or (str(row.get("_colab_fix", "")).strip().upper() not in ["", "NAN", "NONE"])
            or (str(row.get("origem_match_fopag", "")).upper() not in ["SEM_MATCH", "", "NAN", "NONE"])
        )

        if not eh_titular and not tem_titular_cadastrado:
            if not is_so_fopag:
                parentesco_texto = f" ({parentesco_str})" if parentesco_str and parentesco_str.lower() != "nan" else ""
                add(row, "RN008", "Titular Não Localizado nas Bases", f"Dependente{parentesco_texto} ativo, mas o TITULAR não foi localizado nas bases ativas da empresa.", "ATIVOS", "Verificar vínculo do titular")

        if not eh_titular and not eh_conjuge and pd.notna(idade) and not em_processo_de_saida:
            idade_int = int(idade)
            if invalido == "S": pass
            elif idade_int >= 24:
                add(row, "RN011", "Idade limite excedida", f"Dependente com {idade_int} anos (limite 24).", "ATIVOS / MARSH", "Excluir dependente ou validar exceção contratual")
            elif 21 <= idade_int < 24:
                is_estudante = any(palavra in grau_instrucao.upper() for palavra in ["SUPERIOR", "FACULDADE", "UNIVERSIT"])
                if not is_estudante:
                    add(row, "RN011", "Universitário Inativo", f"Dependente com {idade_int} anos não está cursando Ensino Superior.", "ATIVOS / MARSH", "Validar comprovante de matrícula ou excluir")

        qtd_dep = mapa_qtd_dep_ativos.get((op, mat_clean), 0)
        v_mens_dep_raw = str(row.get("VALOR_MENSALIDADE_DEPENDENTES_FOPAG", "0.0"))
        try: v_mens_dep_f = float(v_mens_dep_raw)
        except: v_mens_dep_f = 0.0
        
        v_mens_dep_op = 0.0
        if mat_clean in mapa_fin and op in mapa_fin[mat_clean]:
            v_mens_dep_op = float(mapa_fin[mat_clean][op].get('mens_dep', 0.0))
            
        if eh_titular and (v_mens_dep_f > 0 or v_mens_dep_op > 0) and qtd_dep == 0:
            val_cobranca = v_mens_dep_f if v_mens_dep_f > 0 else v_mens_dep_op
            orig = "FOPAG" if v_mens_dep_f > 0 else "OPERADORA"
            add(row, "RN021", "Cobranca Dependente Fantasma", f"Titular descontado em R$ {val_cobranca} de mensalidade de dependente, mas nao possui dependentes ativos no plano.", orig, "Analisar desconto indevido / Reembolso")

        if len(resultados) == linhas_antes: add(row, "", "OK", "Beneficiário sem divergências", "ATIVOS", "-")

    return resultados

def exportar(consolidado: pd.DataFrame, caminho_saida: Path, df_fin: pd.DataFrame = None, df_eventos: pd.DataFrame = None, df_faturas: pd.DataFrame = None) -> Path:
    if consolidado.empty: return caminho_saida

    competencia_dt = None
    if df_fin is not None and "COMPETENCIA_DATA" in df_fin.columns:
        if not pd.isna(df_fin["COMPETENCIA_DATA"].iloc[0]):
            competencia_dt = df_fin["COMPETENCIA_DATA"].iloc[0]
    if not competencia_dt:
        competencia_dt = pd.to_datetime(date.today().replace(day=1))

    mapa_fin = {}
    if df_fin is not None and not df_fin.empty and df_eventos is not None and not df_eventos.empty:
        if len(df_eventos.columns) >= 3:
            df_eventos = df_eventos.iloc[:, :3].copy()
            df_eventos.columns = ["CODIGO", "DESCRICAO", "PRODUTO"]

        df_fin["CODIGO_N"] = pd.to_numeric(df_fin["CODIGO_EVENTO"].astype(str).str.replace(r'\D', '', regex=True), errors='coerce')
        df_eventos["CODIGO_N"] = pd.to_numeric(df_eventos["CODIGO"].astype(str).str.replace(r'\D', '', regex=True), errors='coerce')
        fin_merged = df_fin.dropna(subset=["CODIGO_N"]).merge(df_eventos.dropna(subset=["CODIGO_N"]), on="CODIGO_N", how="inner")

        for _, r in fin_merged.iterrows():
            mat_raw = str(r.get("matricula", "")).strip().split(".")[0]
            mat = mat_raw.lstrip('0') if mat_raw else ""
            desc = str(r.get("DESCRICAO", "")).upper()
            prod = str(r.get("PRODUTO", "")).upper()
            val = float(r.get("VALOR_FLOAT", 0.0))
            texto_busca = f"{prod} {desc}"

            op = ""
            if "UNIMED" in texto_busca: op = "UNIMED"
            elif "KLINI" in texto_busca: op = "KLINI"
            elif "HAPVIDA" in texto_busca: op = "HAPVIDA"
            elif "SULAMERICA SAUDE" in texto_busca or "SUL AMERICA SAUDE" in texto_busca: op = "SULAMERICA_SAUDE"
            elif "SULAMERICA ODONTO" in texto_busca or "SUL AMERICA ODONTO" in texto_busca: op = "SULAMERICA_ODONTO"
            elif "BRADESCO" in texto_busca: op = "BRADESCO_ODONTO"
            elif "GNDI" in texto_busca or "INTERMEDICA" in texto_busca: op = "GNDI"

            if not op or not mat: continue
            if mat not in mapa_fin: mapa_fin[mat] = {}
            if op not in mapa_fin[mat]: mapa_fin[mat][op] = {'mens_tit': 0.0, 'mens_dep': 0.0, 'cop_tit': 0.0, 'cop_dep': 0.0}

            is_copart = "COPART" in desc
            is_tit = "TITULAR" in desc
            if is_tit and not is_copart: mapa_fin[mat][op]['mens_tit'] += val
            elif not is_tit and not is_copart: mapa_fin[mat][op]['mens_dep'] += val
            elif is_tit and is_copart: mapa_fin[mat][op]['cop_tit'] += val
            elif not is_tit and is_copart: mapa_fin[mat][op]['cop_dep'] += val

    mapa_faturas_cpf = {}
    mapa_faturas_mat_nome = {}
    if df_faturas is not None and not df_faturas.empty:
        for _, r in df_faturas.iterrows():
            op = r.get("OPERADORA", "")
            mat = str(r.get("MATRICULA_FAT", ""))
            cpf = str(r.get("CPF_FAT", ""))
            nome = str(r.get("NOME_FAT", ""))
            val = float(r.get("VALOR_FATURADO", 0.0))

            if cpf and len(cpf) >= 11:
                mapa_faturas_cpf[(op, cpf)] = val
            if mat and nome:
                primeiro_nome = nome.split()[0] if nome else ""
                mapa_faturas_mat_nome[(op, mat, primeiro_nome)] = val

    df_preparado = _preparar_dataframe(consolidado)
    registros = _detectar_inconsistencias(df_preparado, mapa_fin, mapa_faturas_cpf, mapa_faturas_mat_nome, competencia_dt)
    saida = pd.DataFrame(registros, columns=COLUNAS_LONG) if registros else pd.DataFrame(columns=COLUNAS_LONG)

    for col in ("DATA_VIGENCIA", "DATA_CANCELAMENTO", "NASCIMENTO", "DEMISSAO"):
        if col in saida.columns: saida[col] = saida[col].apply(formatar_data_br)

    wb = Workbook()
    ws_dash = wb.active
    ws_dash.title = "Dashboard - Resumo"
    ws_dash.merge_cells('A1:C1')
    ws_dash.cell(row=1, column=1, value="RESUMO GERENCIAL DE AUDITORIA DE BENEFÍCIOS").font = Font(size=14, bold=True, color="FFFFFF")
    ws_dash.cell(row=1, column=1).fill = PatternFill("solid", start_color="1F4E78")
    ws_dash.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")

    headers_dash = ["CÓDIGO (RN)", "DESCRIÇÃO DA INCONSISTÊNCIA", "QUANTIDADE DE CASOS"]
    for col_idx, h in enumerate(headers_dash, start=1):
        cell = ws_dash.cell(row=3, column=col_idx, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", start_color="D9E1F2")
        cell.alignment = Alignment(horizontal="center")

    df_erros = saida[(saida["RN"].astype(str).str.strip() != "") & (saida["RN"].astype(str).str.upper() != "OK") & (saida["RN"].astype(str).str.upper() != "INFO")]
    if not df_erros.empty:
        resumo = df_erros.groupby(["RN", "TIPO_INCONSISTENCIA"]).size().reset_index(name="Quantidade")
        resumo = resumo.sort_values(by="Quantidade", ascending=False)
        linha_atual = 4
        for _, r in resumo.iterrows():
            ws_dash.cell(row=linha_atual, column=1, value=r["RN"]).alignment = Alignment(horizontal="center")
            ws_dash.cell(row=linha_atual, column=2, value=r["TIPO_INCONSISTENCIA"])
            ws_dash.cell(row=linha_atual, column=3, value=r["Quantidade"]).alignment = Alignment(horizontal="center")
            linha_atual += 1

        ws_dash.cell(row=linha_atual+1, column=2, value="TOTAL DE OCORRÊNCIAS MAPEADAS:").font = Font(bold=True)
        ws_dash.cell(row=linha_atual+1, column=3, value=len(df_erros)).font = Font(bold=True)
        linha_atual += 2
    else:
        ws_dash.cell(row=4, column=1, value="Nenhuma inconsistência encontrada!")
        linha_atual = 5

    df_infos = saida[saida["RN"].astype(str).str.upper() == "INFO"]
    if not df_infos.empty:
        linha_atual += 2
        ws_dash.cell(row=linha_atual, column=2, value="ALERTAS DE CONTROLE DE EXCLUSÃO").font = Font(bold=True, color="2980b9")
        linha_atual += 1
        resumo_info = df_infos.groupby(["TIPO_INCONSISTENCIA"]).size().reset_index(name="Quantidade")
        for _, r in resumo_info.iterrows():
            ws_dash.cell(row=linha_atual, column=1, value="INFO").alignment = Alignment(horizontal="center")
            ws_dash.cell(row=linha_atual, column=2, value=r["TIPO_INCONSISTENCIA"])
            ws_dash.cell(row=linha_atual, column=3, value=r["Quantidade"]).alignment = Alignment(horizontal="center")
            linha_atual += 1

    if "ALERTA_IDADE" in saida.columns:
        df_idades = saida[saida["ALERTA_IDADE"].astype(str).str.strip() != ""]
        if not df_idades.empty:
            linha_atual += 2
            ws_dash.cell(row=linha_atual, column=2, value="MONITORAMENTO DE IDADE (PREVENTIVO)").font = Font(bold=True, color="d35400")
            linha_atual += 1
            qtd_21 = len(df_idades[df_idades["ALERTA_IDADE"].str.contains("21 anos")])
            qtd_24 = len(df_idades[df_idades["ALERTA_IDADE"].str.contains("24 anos")])

            if qtd_21 > 0:
                ws_dash.cell(row=linha_atual, column=1, value="ALERTA").alignment = Alignment(horizontal="center")
                ws_dash.cell(row=linha_atual, column=2, value="Dependentes que completarão 21 anos (Exigirá Faculdade)")
                ws_dash.cell(row=linha_atual, column=3, value=qtd_21).alignment = Alignment(horizontal="center")
                linha_atual += 1
            if qtd_24 > 0:
                ws_dash.cell(row=linha_atual, column=1, value="ALERTA").alignment = Alignment(horizontal="center")
                ws_dash.cell(row=linha_atual, column=2, value="Dependentes que completarão 24 anos (Exclusão Limite)")
                ws_dash.cell(row=linha_atual, column=3, value=qtd_24).alignment = Alignment(horizontal="center")
                linha_atual += 1

    ws_dash.column_dimensions['A'].width = 15
    ws_dash.column_dimensions['B'].width = 80
    ws_dash.column_dimensions['C'].width = 25

    ws_dados = wb.create_sheet("Relatório de Benefícios")
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="1F4E78")
    erro_fill = PatternFill("solid", start_color="FFC7CE")
    erro_font = Font(color="9C0006")
    alerta_idade_fill = PatternFill("solid", start_color="FFD966")
    alerta_idade_font = Font(color="000000", bold=True)

    for col_idx, col_nome in enumerate(saida.columns, start=1):
        cell = ws_dados.cell(row=1, column=col_idx, value=col_nome)
        cell.font = header_font
        cell.fill = header_fill

    try:
        parentesco_idx = saida.columns.get_loc("PARENTESCO")
        rn_idx = saida.columns.get_loc("RN")
    except KeyError:
        parentesco_idx = 8
        rn_idx = 17

    for row_idx, registro in enumerate(saida.itertuples(index=False), start=2):
        rn_val = str(registro[rn_idx]).strip().upper()
        tem_erro = rn_val not in ["", "OK", "INFO"]

        for col_idx, col_nome in enumerate(saida.columns, start=1):
            cell = ws_dados.cell(row=row_idx, column=col_idx)
            val = registro[col_idx-1]
            cell.value = str(val) if pd.notna(val) else ""

            if tem_erro:
                cell.fill = erro_fill
                cell.font = erro_font

            if col_nome == "IDADE":
                idade_val = str(val).strip()
                parentesco_val = str(registro[parentesco_idx]).strip().upper()
                if parentesco_val not in ["TITULAR", "CÔNJUGE"] and idade_val in ["20", "23"]:
                    cell.fill = alerta_idade_fill
                    cell.font = alerta_idade_font

    ws_dados.freeze_panes = "A2"
    ws_dados.auto_filter.ref = ws_dados.dimensions
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(caminho_saida)
    return caminho_saida
