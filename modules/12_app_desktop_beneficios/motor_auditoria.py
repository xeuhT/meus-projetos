"""
motor_auditoria.py - Cruza Ativos × Fopag × RH × Demitidos × Dependentes Marsh.
V_ESTAVEL - Cruzamento cascata seguro e mapeamento de colunas do RH.
"""
import logging
import pandas as pd
from utils import limpar_cpf, limpar_matricula, normalizar_nome

log = logging.getLogger(__name__)

def _sanear_chaves(df: pd.DataFrame, contexto: str) -> pd.DataFrame:
    if df is None or df.empty: return df
    df = df.copy()
    if "cpf" not in df.columns: df["cpf"] = ""
    df["cpf"] = df["cpf"].apply(limpar_cpf).astype(str)
    if "matricula" not in df.columns: df["matricula"] = ""
    df["matricula"] = df["matricula"].apply(limpar_matricula).astype(str)

    base_nome = df["beneficiario"] if "beneficiario" in df.columns else df.get("nome", "")
    df["beneficiario_norm"] = base_nome.apply(normalizar_nome) if hasattr(base_nome, "apply") else ""
    return df

def _merge_cascata(esquerda, direita, sufixo_direita, coluna_origem, nome_esq="esq", nome_dir="dir"):
    if esquerda is None or esquerda.empty: return pd.DataFrame()
    esquerda = _sanear_chaves(esquerda, contexto=nome_esq)
    direita = _sanear_chaves(direita, contexto=nome_dir)

    if direita is None or direita.empty:
        return esquerda.assign(**{coluna_origem: "SEM_MATCH"})

    chaves = {"cpf", "matricula", "beneficiario_norm"}
    direita_ren = direita.rename(columns={c: f"{c}_{sufixo_direita}" for c in direita.columns if c not in chaves})

    def _tentar(left, right, chave, drop_keys):
        if left.empty: return left, left
        right_sel = right.drop(columns=drop_keys, errors="ignore")
        merged = left.merge(right_sel, on=chave, how="left", indicator="_flag")
        sucesso = merged[merged["_flag"] == "both"].drop(columns="_flag").copy()
        cols_drop = ["_flag"] + [c for c in right_sel.columns if c not in (chave if isinstance(chave, list) else [chave])]
        falha = merged[merged["_flag"] == "left_only"].drop(columns=cols_drop, errors="ignore").copy()
        return sucesso, falha

    com_mat_nome = esquerda[(esquerda["matricula"].str.len() > 0) & (esquerda["beneficiario_norm"].str.len() > 0)].copy()
    sem_mat_nome = esquerda[~((esquerda["matricula"].str.len() > 0) & (esquerda["beneficiario_norm"].str.len() > 0))].copy()

    dir_mat_nome = direita_ren[(direita_ren["matricula"].str.len() > 0) & (direita_ren["beneficiario_norm"].str.len() > 0)].drop_duplicates(subset=["matricula", "beneficiario_norm"], keep="first")
    casados_mat_nome, resto_mat_nome = _tentar(com_mat_nome, dir_mat_nome, ["matricula", "beneficiario_norm"], ["cpf"])
    casados_mat_nome[coluna_origem] = "MATRICULA_NOME"

    resto1 = pd.concat([resto_mat_nome, sem_mat_nome], ignore_index=True)
    com_cpf = resto1[resto1["cpf"].str.len() == 11].copy()
    sem_cpf = resto1[resto1["cpf"].str.len() != 11].copy()

    dir_cpf = direita_ren[direita_ren["cpf"].str.len() == 11].drop_duplicates(subset=["cpf"], keep="first")
    casados_cpf, resto_cpf = _tentar(com_cpf, dir_cpf, ["cpf"], ["matricula", "beneficiario_norm"])
    casados_cpf[coluna_origem] = "CPF"

    resto2 = pd.concat([resto_cpf, sem_cpf], ignore_index=True)
    com_nome = resto2[resto2["beneficiario_norm"].str.len() > 0].copy()
    sem_nome = resto2[resto2["beneficiario_norm"].str.len() == 0].copy()

    dir_nome = direita_ren[direita_ren["beneficiario_norm"].str.len() > 0].drop_duplicates(subset=["beneficiario_norm"], keep="first")
    casados_nome, nao_casados = _tentar(com_nome, dir_nome, ["beneficiario_norm"], ["cpf", "matricula"])
    casados_nome[coluna_origem] = "NOME"

    nao_casados[coluna_origem] = "SEM_MATCH"
    sem_nome[coluna_origem] = "SEM_MATCH"

    return pd.concat([casados_mat_nome, casados_cpf, casados_nome, nao_casados, sem_nome], ignore_index=True)

def auditar_operadora(ativos, fopag, colaboradores, demitidos, dependentes, operadora):
    ativos = _sanear_chaves(ativos, contexto=f"{operadora}-ATIVOS")
    fopag = _sanear_chaves(fopag, contexto=f"{operadora}-FOPAG")
    colaboradores = _sanear_chaves(colaboradores, contexto="RH-COLABORADORES")
    demitidos = _sanear_chaves(demitidos, contexto="RH-DEMITIDOS")

    cruzado = _merge_cascata(ativos, fopag, "fopag", "origem_match_fopag", f"{operadora}-ATIVOS", f"{operadora}-FOPAG")

    if fopag is not None and not fopag.empty:
        num_ativos = len(ativos) if ativos is not None else 0
        cpfs_at = set(ativos.loc[ativos["cpf"].str.len() == 11, "cpf"]) if num_ativos > 0 else set()
        mats_at = set(ativos.loc[ativos["matricula"].str.len() > 0, "matricula"]) if num_ativos > 0 else set()
        nomes_at = set(ativos.loc[ativos["beneficiario_norm"].str.len() > 0, "beneficiario_norm"]) if num_ativos > 0 else set()
        so_fopag = fopag[~fopag["cpf"].isin(cpfs_at) & ~fopag["matricula"].isin(mats_at) & ~fopag["beneficiario_norm"].isin(nomes_at)].copy()
        if not so_fopag.empty:
            so_fopag = so_fopag.rename(columns={c: f"{c}_fopag" for c in so_fopag.columns if c not in {"cpf", "matricula", "beneficiario_norm"}})
            so_fopag["origem_match_fopag"] = "SO_FOPAG"
            cruzado = pd.concat([cruzado, so_fopag], ignore_index=True)

    palavras_chave = ["matricula", "cpf", "beneficiario", "nome", "cargo", "sindicato", "admissao", "situacao", "empresa", "estabelecimento", "loja", "nascimento", "nasc", "razao_social", "natureza"]
    cols_para_manter = [c for c in colaboradores.columns if any(p in c.lower() for p in palavras_chave)]
    colaboradores_clean = colaboradores[cols_para_manter].copy()
    if not colaboradores_clean.empty:
        if "cpf" in colaboradores_clean.columns:
            colaboradores_clean["cpf_colaborador"] = colaboradores_clean["cpf"]
        if "matricula" in colaboradores_clean.columns:
            colaboradores_clean["matricula_colaborador_rh"] = colaboradores_clean["matricula"]

    cruzado = _merge_cascata(cruzado, colaboradores_clean, "rh", "origem_match_rh", f"{operadora}-CONSOL", "RH-COLABORADORES")

    if "cpf" in cruzado.columns and "matricula" in cruzado.columns and not colaboradores_clean.empty:
        df_colab_validos = colaboradores_clean[colaboradores_clean["cpf"].str.len() == 11].drop_duplicates(subset=["cpf"], keep="first")
        mapa_cpf_mat = df_colab_validos.set_index("cpf")["matricula"].to_dict()
        def preencher_matricula(row):
            mat = str(row.get("matricula", "")).strip()
            if mat == "" or mat.lower() == "nan":
                cpf_val = str(row.get("cpf", "")).strip()
                if len(cpf_val) == 11: return mapa_cpf_mat.get(cpf_val, "")
            return mat
        cruzado["matricula"] = cruzado.apply(preencher_matricula, axis=1)

        # Resgate de dados corporativos do RH para Dependentes pela matrícula do Titular
        mask_sem_rh = (cruzado["origem_match_rh"] == "SEM_MATCH") & (cruzado["matricula"].astype(str).str.strip() != "")
        if mask_sem_rh.any():
            chaves_base = {"cpf", "matricula", "beneficiario_norm"}
            colab_mat_ren = colaboradores_clean.rename(columns={c: f"{c}_rh" for c in colaboradores_clean.columns if c not in chaves_base}).copy()
            colab_mat_ren["_mat_join"] = colab_mat_ren["matricula"].astype(str).str.split(".").str[0].str.lstrip("0").str.strip()
            colab_mat_ren = colab_mat_ren[colab_mat_ren["_mat_join"] != ""].drop_duplicates(subset=["_mat_join"], keep="first")
            cols_rh_resgate = [c for c in colab_mat_ren.columns if c.endswith("_rh") and not any(p in c.lower() for p in ["nasc", "beneficiario"])]
            if cols_rh_resgate and not colab_mat_ren.empty:
                temp_esq = cruzado.loc[mask_sem_rh, ["matricula"]].copy()
                temp_esq["_mat_join"] = temp_esq["matricula"].astype(str).str.split(".").str[0].str.lstrip("0").str.strip()
                resgate_rh = temp_esq.reset_index().merge(colab_mat_ren[["_mat_join"] + cols_rh_resgate], on="_mat_join", how="inner").set_index("index")
                if not resgate_rh.empty:
                    for col_r in cols_rh_resgate:
                        if col_r not in cruzado.columns:
                            cruzado[col_r] = ""
                        cruzado.loc[resgate_rh.index, col_r] = resgate_rh[col_r]
                    cruzado.loc[resgate_rh.index, "origem_match_rh"] = "MATRICULA_TITULAR"

    if demitidos is not None and not demitidos.empty and "cpf" in demitidos.columns:
        dem = demitidos[demitidos["cpf"].str.len() == 11].copy()
        dem = dem.drop_duplicates(subset=["cpf"], keep="last").add_prefix("demitido_").rename(columns={"demitido_cpf": "cpf"})
        if "cpf" not in cruzado.columns: cruzado["cpf"] = ""
        cruzado["cpf"] = cruzado["cpf"].apply(limpar_cpf).astype(str)
        cruzado = cruzado.merge(dem, on="cpf", how="left")

    if dependentes is not None and not dependentes.empty and not cruzado.empty:
        dependentes = _sanear_chaves(dependentes, contexto="RH-DEPENDENTES")
        dependentes["_primeiro_nome"] = dependentes["beneficiario_norm"].str.split().str[0].fillna("")
        if "_primeiro_nome" not in cruzado.columns:
            cruzado["_primeiro_nome"] = cruzado["beneficiario_norm"].str.split().str[0].fillna("")

        chaves_dep = {"matricula", "beneficiario_norm", "cpf", "_primeiro_nome"}
        cols_originais = [c for c in dependentes.columns if c not in chaves_dep]
        dep_ren = dependentes.rename(columns={c: f"{c}_marsh" for c in cols_originais})
        cols_marsh = [f"{c}_marsh" for c in cols_originais]

        dep_mat_nome = dep_ren.drop(columns=["cpf"]).drop_duplicates(subset=["matricula", "beneficiario_norm"])
        cruzado = cruzado.merge(dep_mat_nome, on=["matricula", "beneficiario_norm"], how="left")

        if cols_marsh:
            col_check = cols_marsh[0]
            if "cpf" in cruzado.columns and "cpf" in dep_ren.columns:
                dep_cpf = dep_ren.drop(columns=["matricula", "beneficiario_norm"]).dropna(subset=["cpf"])
                dep_cpf = dep_cpf[dep_cpf["cpf"].str.len() == 11].drop_duplicates(subset=["cpf"])
                mask_cpf = cruzado[col_check].isna() & (cruzado["cpf"].str.len() == 11)
                if mask_cpf.any():
                    resgate_cpf = cruzado.loc[mask_cpf, ["cpf"]].reset_index().merge(dep_cpf, on="cpf", how="inner").set_index("index")
                    cruzado.update(resgate_cpf)

    cruzado["operadora"] = operadora
    return cruzado

def auditar_todas(ativos_por_operadora, fopag_por_operadora, colaboradores, demitidos, dependentes):
    resultados = []
    for op in ativos_por_operadora:
        res = auditar_operadora(ativos_por_operadora.get(op, pd.DataFrame()), fopag_por_operadora.get(op, pd.DataFrame()), colaboradores, demitidos, dependentes, op)
        if not res.empty: resultados.append(res)
    return pd.concat(resultados, ignore_index=True) if resultados else pd.DataFrame()