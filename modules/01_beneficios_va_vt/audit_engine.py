"""
Módulo do Motor de Cálculo e Regras de Auditoria de 6% (VT).
Execução 100% vetorizada e em memória via Pandas.
"""

import pandas as pd
import numpy as np
from datetime import datetime
from config import TOLERANCIA_CENTAVOS, COLUNAS_OFICIAIS_CONFERENCIA

class AuditEngine:
    def __init__(self, competencia, logger=None):
        self.competencia = competencia
        self.logger = logger
        self.mes_comp = int(competencia.split("/")[0])
        self.ano_comp = int(competencia.split("/")[1])

    def _log(self, msg):
        if self.logger:
            self.logger.log(msg)
        else:
            print(msg)

    def process(self, df_folha, df_dem, df_afast, df_ferias, df_nom, df_compra=None, df_retorno=None):
        self._log("Iniciando motor de cruzamento e cálculo de auditoria...")
        
        # 1. Base mestre é df_folha
        df = df_folha.copy()
        
        # Garantir tipo da Matrícula
        df["Matrícula"] = df["Matrícula"].astype(str).str.strip().str.lstrip("0")
        
        # 2. Joins Relacionais
        # Join com Demitidos
        if df_dem is not None and not df_dem.empty:
            df_dem_sub = df_dem[["Matrícula", "Data Rescisão"]].copy() if "Data Rescisão" in df_dem.columns else pd.DataFrame(columns=["Matrícula", "Data Rescisão"])
            df = df.merge(df_dem_sub.rename(columns={"Data Rescisão": "Desligados"}), on="Matrícula", how="left")
        else:
            df["Desligados"] = np.nan

        # Join com Afastados
        if df_afast is not None and not df_afast.empty:
            cols_afast = ["Matrícula", "Início Afastamento", "Final Afastamento"]
            df_afast_sub = df_afast[[c for c in cols_afast if c in df_afast.columns]].copy()
            df = df.merge(df_afast_sub, on="Matrícula", how="left")
        else:
            df["Início Afastamento"] = np.nan
            df["Final Afastamento"] = np.nan

        # Join com Férias
        if df_ferias is not None and not df_ferias.empty:
            cols_fer = ["Matrícula", "Início Férias", "Final Férias", "Qtd de Dias Férias"]
            df_fer_sub = df_ferias[[c for c in cols_fer if c in df_ferias.columns]].copy()
            df = df.merge(df_fer_sub, on="Matrícula", how="left")
        else:
            df["Início Férias"] = np.nan
            df["Final Férias"] = np.nan
            df["Qtd de Dias Férias"] = np.nan

        # Join com Nomina
        if df_nom is not None and not df_nom.empty:
            sit_col = "Situação" if "Situação" in df_nom.columns else ("Situacao" if "Situacao" in df_nom.columns else None)
            if sit_col:
                df_nom_sub = df_nom[["Matrícula", sit_col]].copy().rename(columns={sit_col: "Situação_Nomina"})
                df = df.merge(df_nom_sub, on="Matrícula", how="left")
            else:
                df["Situação_Nomina"] = np.nan
        else:
            df["Situação_Nomina"] = np.nan

        # Preencher Situação (Coluna 38 - AL)
        if "Situação" in df.columns and "Situação_Nomina" in df.columns:
            df["Situação"] = df["Situação_Nomina"].fillna(df["Situação"]).fillna("Atividade Normal")
        elif "Situação_Nomina" in df.columns:
            df["Situação"] = df["Situação_Nomina"].fillna("Atividade Normal")
        elif "Situação" not in df.columns:
            df["Situação"] = "Atividade Normal"

        # 3. Cálculo da Coluna 21 (U): Soma - Base de cálculo 6%
        # Regra validada pela equipe (Tamires): somar estritamente Coluna M (1 - Salario) + Coluna S (137) + Coluna T (138)
        df["Soma - Base de cálculo 6%"] = (
            df["1 - Salario"] +
            df["137 - Auxilio Doenca Previdenciario"] +
            df["138 - Auxilio Acid de Trabalho Previdenciario"]
        ).round(2)

        # 4. Cálculo da Coluna 24 (X): Custo Total Vale Transporte
        df["Custo Total Vale Transporte"] = (
            df["31903 - Custo Total Vale Transporte"] + df["31905 - Aux Credito Vale Transporte"]
        ).round(2)

        # 5. Coluna 27 (AA): Cálculo dos 6%
        # Regra validada pela equipe (Tamires): Coluna N (11010 - Salario Mensal) multiplicado por 6%
        df["Cálculo dos 6% "] = (df["11010 - Salario Mensal"] * 0.06).round(2)

        # 6. Coluna 28 (AB): 6% ou Valor da compra
        # Fórmula: =SE([@[Cálculo dos 6% ]]<=[@[Custo Total Vale Transporte]];"6%";"Valor da compra")
        df["6% ou Valor da compra"] = np.where(
            df["Cálculo dos 6% "] <= df["Custo Total Vale Transporte"],
            "6%",
            "Valor da compra"
        )

        # 7. Coluna 29 (AC): Validação Desconto - 6%
        # Regra validada pela equipe (Tamires): comparar Coluna Z (6290 - Vale Transporte) com Coluna AA (Cálculo dos 6%)
        df["Validação Desconto - 6%"] = (
            df["6290 - Vale Transporte "] - df["Cálculo dos 6% "]
        ).round(2)

        # 8. Coluna 30 (AD): Validação Desconto - Valor da Compra
        # Regra validada pela equipe (Tamires): comparar Coluna Z (6290 - Vale Transporte) com Coluna X (Custo Total Vale Transporte)
        df["Validação Desconto - Valor da Compra"] = (
            df["6290 - Vale Transporte "] - df["Custo Total Vale Transporte"]
        ).round(2)

        # 9. Coluna 31 (AE): Validação Custo empresa + Desc 6% x Custo Total de Vale Transporte
        # Fórmula: =[@[31906 - Custo Empresa Vale Transporte]]+[@[6290 - Vale Transporte ]]=[@[Custo Total Vale Transporte]]
        soma_emp_func = (df["31906 - Custo Empresa Vale Transporte"] + df["6290 - Vale Transporte "]).round(2)
        diff_soma = (soma_emp_func - df["Custo Total Vale Transporte"]).abs()
        df["Validação Custo empresa + Desc 6% x\nCusto Total de Vale Transporte"] = np.where(
            diff_soma <= TOLERANCIA_CENTAVOS,
            "VERDADEIRO",
            "FALSO"
        )

        # 10. Identificadores auxiliares para classificação da Coluna 39 (AM: Obs. Conferência desconto 6%)
        df["_dt_adm"] = pd.to_datetime(df["Data Admissão"], format="%d/%m/%Y", errors="coerce")
        df["_eh_novo_admitido"] = (
            (df["_dt_adm"].dt.month == self.mes_comp) &
            (df["_dt_adm"].dt.year == self.ano_comp)
        )

        c_tot = df["Custo Total Vale Transporte"]
        d_6290 = df["6290 - Vale Transporte "]
        c_emp = df["31906 - Custo Empresa Vale Transporte"]
        fechou_custo = df["Validação Custo empresa + Desc 6% x\nCusto Total de Vale Transporte"] == "VERDADEIRO"

        # Diferenças contra os possíveis tetos válidos de 6%:
        # a) Teto sobre 11010 - Salario Mensal (Coluna N)
        esperado_n = np.minimum(df["Cálculo dos 6% "], c_tot)
        diff_n = (d_6290 - esperado_n).abs()

        # b) Teto sobre Soma - Base de cálculo 6% (Coluna U = M + S + T)
        calc_6_u = (df["Soma - Base de cálculo 6%"] * 0.06).round(2)
        esperado_u = np.minimum(calc_6_u, c_tot)
        diff_u = (d_6290 - esperado_u).abs()

        # c) Teto sobre Salário Proporcional de Admissão no mês
        dias_trab = np.where(df["_eh_novo_admitido"], (30 - df["_dt_adm"].dt.day + 1).clip(lower=1, upper=30), 30)
        calc_6_adm = (((df["11010 - Salario Mensal"] / 30.0) * dias_trab) * 0.06).round(2)
        esperado_adm = np.minimum(calc_6_adm, c_tot)
        diff_adm = (d_6290 - esperado_adm).abs()

        # d) Teto sobre Mínimo Garantido / Horas Comissionista
        base_comiss = (df["2 - Minimo Garantido Comissionista"] + df["7 - Horas Trabalhadas"] + df["8 - DSR Horas Trabalhadas"]).round(2)
        calc_6_comiss = (base_comiss * 0.06).round(2)
        esperado_comiss = np.minimum(calc_6_comiss, c_tot)
        diff_comiss = (d_6290 - esperado_comiss).abs()

        # Desconto válido segundo qualquer uma das bases aplicáveis ao perfil
        desc_valido_n_ou_u = (diff_n <= TOLERANCIA_CENTAVOS) | (diff_u <= TOLERANCIA_CENTAVOS)
        desc_valido_adm = desc_valido_n_ou_u | (diff_adm <= TOLERANCIA_CENTAVOS)

        # Coluna 39 (AM): Obs. Conferência desconto 6%
        # Regra validada pela equipe (Tamires): As observações detalhadas devem ficar na Coluna AM
        df["Obs. Conferência desconto 6%"] = ""

        # Condição 1: Sem custo de VT e sem desconto dos 6% - Ok
        mask_sem_vt = (c_tot == 0) & (d_6290 == 0) & (c_emp == 0)
        df.loc[mask_sem_vt, "Obs. Conferência desconto 6%"] = "Sem custo de VT e sem desconto dos 6% - Ok"

        # Condição 2: Estagiário (com compra de VT e sem desconto 6290 -> Regular; com desconto 6290 > 0 -> Indevido)
        eh_estagiario = (
            df["Cargo"].str.contains("ESTAG", case=False, na=False) |
            df["Situação"].str.contains("Estagiario|Estagiário", case=False, na=False)
        )
        mask_estagio_ok = eh_estagiario & (c_tot > 0) & (d_6290 == 0)
        df.loc[mask_estagio_ok, "Obs. Conferência desconto 6%"] = "Estagiário"

        mask_estagio_indevido = eh_estagiario & (d_6290 > 0)
        df.loc[mask_estagio_indevido, "Obs. Conferência desconto 6%"] = "Estagiário - Desconto Indevido"

        # Regras de Negócio Validadas com a Equipe (Tamires - Versão 3):
        # 1. Afastamento (Colunas AG e AH): Só considera se houver período de afastamento preenchido nas colunas AG ou AH
        tem_afast = (
            df["Início Afastamento"].notna() &
            ~df["Início Afastamento"].astype(str).str.strip().isin(["", "nan", "None", "NaT"])
        ) | (
            df["Final Afastamento"].notna() &
            ~df["Final Afastamento"].astype(str).str.strip().isin(["", "nan", "None", "NaT"])
        )

        # 2. Férias (Colunas AI e AJ): Só considera se apresentar período de férias nas colunas AI e AJ
        tem_ferias = (
            df["Início Férias"].notna() &
            ~df["Início Férias"].astype(str).str.strip().isin(["", "nan", "None", "NaT"]) &
            df["Final Férias"].notna() &
            ~df["Final Férias"].astype(str).str.strip().isin(["", "nan", "None", "NaT"])
        )

        # Condição 3: Vendedor Comissionista em Férias / Comissionista - Desconto Ok
        eh_comiss = df["Cargo"].str.contains("VENDEDOR|COMISSION", case=False, na=False) & (df["2 - Minimo Garantido Comissionista"] > 0)
        mask_comiss_ferias = (
            (df["Obs. Conferência desconto 6%"] == "") &
            eh_comiss &
            tem_ferias &
            (d_6290 > 0)
        )
        df.loc[mask_comiss_ferias, "Obs. Conferência desconto 6%"] = "Vendedor Comissionista - Desconto Ok"

        # Condição 4: Novo Admitido - Desconto Ok
        mask_ok_adm = (
            (df["Obs. Conferência desconto 6%"] == "") &
            df["_eh_novo_admitido"] &
            (d_6290 > 0) &
            desc_valido_adm &
            fechou_custo
        )
        df.loc[mask_ok_adm, "Obs. Conferência desconto 6%"] = "Novo Admitido - Desconto Ok"

        # Condição 5: Afastamento - Desconto 6% Ok (teve desconto proporcional validado)
        mask_ok_afast = (
            (df["Obs. Conferência desconto 6%"] == "") &
            tem_afast &
            (d_6290 > 0) &
            desc_valido_n_ou_u &
            fechou_custo
        )
        df.loc[mask_ok_afast, "Obs. Conferência desconto 6%"] = "Afastamento - Desconto 6% Ok"

        # Condição 6: Férias - Desconto 6% Ok (teve desconto proporcional ou integral validado)
        mask_ok_ferias = (
            (df["Obs. Conferência desconto 6%"] == "") &
            tem_ferias &
            (d_6290 > 0) &
            (desc_valido_n_ou_u | (diff_comiss <= TOLERANCIA_CENTAVOS)) &
            fechou_custo
        )
        df.loc[mask_ok_ferias, "Obs. Conferência desconto 6%"] = "Férias - Desconto 6% Ok"

        # Condição 7: Desconto 6% Ok padrão (colaborador ativo sem férias/afastamento com 6290 batendo com teto de N)
        mask_ok_padrao = (
            (df["Obs. Conferência desconto 6%"] == "") &
            (c_tot > 0) &
            (d_6290 > 0) &
            (diff_n <= TOLERANCIA_CENTAVOS) &
            fechou_custo
        )
        df.loc[mask_ok_padrao, "Obs. Conferência desconto 6%"] = "Desconto 6% Ok"

        # Condição 8: Salário proporcional - Desconto 6% Ok (ativo sem férias/afastamento cujo 6290 bate com 6% da Soma U)
        mask_prop = (
            (df["Obs. Conferência desconto 6%"] == "") &
            (c_tot > 0) &
            (d_6290 > 0) &
            (diff_u <= TOLERANCIA_CENTAVOS) &
            fechou_custo
        )
        df.loc[mask_prop, "Obs. Conferência desconto 6%"] = "Salário proporcional - Desconto 6% Ok"

        # Condição 9: Vendedor Comissionista geral (sem férias, batendo com mínimo garantido/horas)
        mask_comiss_geral = (
            (df["Obs. Conferência desconto 6%"] == "") &
            eh_comiss &
            (d_6290 > 0) &
            (diff_comiss <= TOLERANCIA_CENTAVOS) &
            fechou_custo
        )
        df.loc[mask_comiss_geral, "Obs. Conferência desconto 6%"] = "Vendedor Comissionista - Desconto Ok"

        # --- DIVERGÊNCIAS / CASOS DE ATENÇÃO NA COLUNA AM ---
        # 10. Afastamento (teve custo/compra de VT ou lançamento sem desconto 6290 por estar afastado)
        mask_afast_sem_desc = (
            (df["Obs. Conferência desconto 6%"] == "") &
            tem_afast &
            ((c_tot > 0) | (c_emp > 0) | (d_6290 > 0)) &
            (d_6290 == 0)
        )
        df.loc[mask_afast_sem_desc, "Obs. Conferência desconto 6%"] = "Afastamento"

        # 11. Férias (teve custo/compra de VT sem desconto 6290 por estar 30 dias em férias)
        mask_ferias_sem_desc = (
            (df["Obs. Conferência desconto 6%"] == "") &
            tem_ferias &
            ((c_tot > 0) | (c_emp > 0) | (d_6290 > 0)) &
            (d_6290 == 0)
        )
        df.loc[mask_ferias_sem_desc, "Obs. Conferência desconto 6%"] = "Férias"

        # 12. Compra sem Desconto dos 6% (Ativo normal com compra e sem desconto)
        mask_compra_sem_desc = (
            (df["Obs. Conferência desconto 6%"] == "") &
            (c_tot > 0) &
            (d_6290 == 0)
        )
        df.loc[mask_compra_sem_desc, "Obs. Conferência desconto 6%"] = "Compra sem Desconto dos 6%"

        # 13. Desconto em Folha sem Compra de VT
        mask_desc_sem_compra = (
            (df["Obs. Conferência desconto 6%"] == "") &
            (c_tot == 0) &
            (d_6290 > 0)
        )
        df.loc[mask_desc_sem_compra, "Obs. Conferência desconto 6%"] = "Desconto em Folha sem Compra de VT"

        # 14. Divergência no Valor do Desconto
        mask_div_valor = (
            (df["Obs. Conferência desconto 6%"] == "") &
            (~desc_valido_adm)
        )
        df.loc[mask_div_valor, "Obs. Conferência desconto 6%"] = "Divergência no Valor do Desconto"

        # Fallback
        df.loc[df["Obs. Conferência desconto 6%"] == "", "Obs. Conferência desconto 6%"] = "Verificar - Não categorizado"

        # 11. Coluna 40 (AN): Obs. Validação Custo empresa + Desc 6%
        # Regra validada pela equipe (Tamires): Baseada na validação da Coluna AE:
        # Se VERDADEIRO -> "Conciliação Ok" | Se FALSO -> "Inconsistente / Revalidar"
        df["Obs. Validação Custo empresa + Desc 6%"] = np.where(
            fechou_custo,
            "Conciliação Ok",
            "Inconsistente / Revalidar"
        )

        # 12. Coluna 41 (AO): Tratativas
        condicoes = [
            df["Obs. Validação Custo empresa + Desc 6%"] == "Inconsistente / Revalidar",
            df["Obs. Conferência desconto 6%"] == "Vendedor Comissionista - Desconto Ok",
            df["Obs. Conferência desconto 6%"] == "Divergência no Valor do Desconto",
            df["Obs. Conferência desconto 6%"] == "Afastamento",
            df["Obs. Conferência desconto 6%"] == "Férias",
            df["Obs. Conferência desconto 6%"] == "Compra sem Desconto dos 6%",
            df["Obs. Conferência desconto 6%"] == "Desconto em Folha sem Compra de VT",
            df["Obs. Conferência desconto 6%"] == "Estagiário - Desconto Indevido",
        ]
        escolhas = [
            "Fechamento contábil incorreto: 31906 + 6290 ≠ Custo Total VT. Revalidar eventos com o DP.",
            "Verificar com quality como é o cálculo de vendedor comissionista quando está de férias",
            "Valor descontado (6290) difere do limite legal de 6% ou do custo da compra. Validar cálculo com o DP e solicitar ajuste.",
            "Colaborador em afastamento com custo de VT (31903/31906) e sem desconto 6290. Verificar período de afastamento.",
            "Colaborador em férias integrais com custo de VT (31903/31906) e sem desconto 6290. Verificar compra no período de férias.",
            "Colaborador ativo com custo de VT na folha porém sem desconto no evento 6290. Verificar com o DP.",
            "Desconto no evento 6290 sem Custo Total de VT correspondente na folha. Verificar com o DP.",
            "Desconto de VT indevido para estagiário (regime não CLT). Solicitar estorno ao DP.",
        ]
        df["Tratativas"] = np.select(condicoes, escolhas, default="")

        # Garantir presença e ordenação estrita das 41 colunas oficiais
        for col in COLUNAS_OFICIAIS_CONFERENCIA:
            if col not in df.columns:
                df[col] = ""

        df_final = df[COLUNAS_OFICIAIS_CONFERENCIA].copy()

        if df_compra is not None and not df_compra.empty:
            df_compra_enriquecida = self._process_compra(df_compra, df_final, df_retorno=df_retorno, df_dem=df_dem)
        else:
            df_compra_enriquecida = None
        if df_retorno is not None and not df_retorno.empty:
            df_retorno_enriquecida = self._process_retorno(df_retorno, df_final, df_dem=df_dem)
        else:
            df_retorno_enriquecida = None

        self._log("Cálculo e categorização da auditoria concluídos com sucesso.")
        return df_final, df_compra_enriquecida, df_retorno_enriquecida

    def _process_compra(self, df_compra, df_final, df_retorno=None, df_dem=None):
        self._log("Enriquecendo Relatório de Compra com dados da folha e retorno (35 colunas)...")
        df_c = df_compra.copy()
        
        # Identificar coluna de matrícula
        mat_col = [c for c in df_c.columns if "matr" in c.lower()][0]
        df_c["_matricula_join"] = df_c[mat_col].astype(str).str.strip().str.lstrip("0")
        
        # Limpar coluna de valor para somar
        val_col = [c for c in df_c.columns if c.strip().lower() == "valor"][0]
        if pd.api.types.is_numeric_dtype(df_c[val_col]):
            df_c["_valor_float"] = df_c[val_col].fillna(0.0).round(2)
        else:
            df_c["_valor_float"] = (
                df_c[val_col].astype(str)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
                .str.replace("-", "0", regex=False)
                .str.strip()
            )
            df_c["_valor_float"] = pd.to_numeric(df_c["_valor_float"], errors="coerce").fillna(0.0).round(2)
        
        # Coluna 25 (Y): Valor da compra Consolidado por Matrícula
        df_c["Valor da compra Consolidado"] = df_c.groupby("_matricula_join")["_valor_float"].transform("sum").round(2)
        
        # Mapear dados da Folha (df_final)
        df_f_sub = df_final[[
            "Matrícula", 
            "Custo Total Vale Transporte", 
            "Situação", 
            "Cargo", 
            "Data Admissão", 
            "Obs. Conferência desconto 6%", 
            "Tratativas"
        ]].drop_duplicates(subset=["Matrícula"])
        
        df_c = df_c.merge(df_f_sub, left_on="_matricula_join", right_on="Matrícula", how="left")
        
        # Mapear Desligados diretamente de df_dem (pois desligados fora da folha ativa não constam em df_final)
        if df_dem is not None and not df_dem.empty and "Data Rescisão" in df_dem.columns:
            df_dem_sub = (
                df_dem[["Matrícula", "Data Rescisão"]]
                .drop_duplicates(subset=["Matrícula"])
                .rename(columns={"Matrícula": "_mat_dem", "Data Rescisão": "Desligado"})
            )
            df_c = df_c.merge(df_dem_sub, left_on="_matricula_join", right_on="_mat_dem", how="left")
        else:
            df_c["Desligado"] = ""
        df_c["Desligado"] = df_c["Desligado"].fillna("")

        # Coluna 28 (AB): Arquivo de retorno
        # Regra validada pela equipe (Tamires - AB1):
        # Sinalizar os colaboradores que tiveram importação relacionada ao arquivo de retorno,
        # buscando essa informação na aba "Arquivo de Retorno", coluna "G" (Ocorrência).
        if df_retorno is not None and not df_retorno.empty and "Ocorrência" in df_retorno.columns:
            df_ret_sub = df_retorno[["Matricula", "Ocorrência"]].copy()
            df_ret_sub["_mat_ret"] = df_ret_sub["Matricula"].astype(str).str.strip().str.lstrip("0")
            df_ret_sub = df_ret_sub.drop_duplicates(subset=["_mat_ret"])[["_mat_ret", "Ocorrência"]]
            df_c = df_c.merge(df_ret_sub, left_on="_matricula_join", right_on="_mat_ret", how="left")
            df_c["Arquivo de retorno"] = df_c["Ocorrência"].fillna("")
        else:
            df_c["Arquivo de retorno"] = ""

        # Coluna 27 (AA): Compra x Folha
        na_folha = df_c["Custo Total Vale Transporte"].notna()
        custo_folha = df_c["Custo Total Vale Transporte"].fillna(0.0).round(2)
        diff = (df_c["Valor da compra Consolidado"] - custo_folha).abs()
        df_c["Compra x Folha"] = np.where(na_folha & (diff <= TOLERANCIA_CENTAVOS), "VERDADEIRO", "FALSO")
        
        # Renomear/Padronizar colunas da folha
        df_c["Observação de Conferência"] = df_c["Obs. Conferência desconto 6%"].fillna("")
        df_c["Admissão"] = df_c["Data Admissão"].fillna("")
        df_c["Tratativa"] = df_c["Tratativas"].fillna("")

        # Coluna 34 (AH): Obs. Conferência 2
        # Regra validada pela equipe (Tamires - AH1):
        # 1. Se verdadeiro na coluna "AA" (Compra x Folha) -> mesma observação da coluna "AC" (Observação de Conferência)
        # 2. Se falso na coluna "AA":
        #    - Se colaborador desligado (consta data em Desligado e não bateu folha) -> "Desligado"
        #    - Se valor maior na coluna "Y" (Valor da compra Consolidado) do que na coluna "Z" (Custo Total Vale Transporte)
        #      e na coluna "AB" (Arquivo de retorno) está sinalizado como "Arquivo de Retorno - Gestão de Saldo Parcial"
        #      -> "Arquivo de Retorno - Gestão de Saldo Parcial"
        #    - Se apresentar valor na coluna "Y" (Valor da compra Consolidado > 0) e valor zerado na coluna "Z" (Custo Total == 0)
        #      e na coluna "AB" (Arquivo de retorno) está sinalizado como "Arquivo de Retorno - Gestão de Saldo Total"
        #      -> "Arquivo de Retorno - Gestão de Saldo Total"
        #    - Demais casos -> "Divergência: VT vs. Folha"
        y_val = df_c["Valor da compra Consolidado"]
        z_val = custo_folha
        ab_val = df_c["Arquivo de retorno"]

        cond_obs2 = [
            df_c["Compra x Folha"] == "VERDADEIRO",
            (df_c["Compra x Folha"] == "FALSO") & (df_c["Desligado"].astype(str).str.strip() != ""),
            (df_c["Compra x Folha"] == "FALSO") & (y_val > z_val) & (ab_val == "Arquivo de Retorno - Gestão de Saldo Parcial"),
            (df_c["Compra x Folha"] == "FALSO") & (y_val > 0) & (z_val == 0) & (ab_val == "Arquivo de Retorno - Gestão de Saldo Total"),
        ]
        escolhas_obs2 = [
            df_c["Observação de Conferência"],
            "Desligado",
            "Arquivo de Retorno - Gestão de Saldo Parcial",
            "Arquivo de Retorno - Gestão de Saldo Total",
        ]
        df_c["Obs. Conferência 2"] = np.select(cond_obs2, escolhas_obs2, default="Divergência: VT vs. Folha")
        
        # Montar as 35 colunas na ordem estrita
        cols_base = list(df_compra.columns)
        cols_adicionais = [
            "Valor da compra Consolidado",
            "Custo Total Vale Transporte",
            "Compra x Folha",
            "Arquivo de retorno",
            "Observação de Conferência",
            "Situação",
            "Desligado",
            "Cargo",
            "Admissão",
            "Obs. Conferência 2",
            "Tratativa"
        ]
        
        colunas_finais = cols_base + cols_adicionais
        for c in colunas_finais:
            if c not in df_c.columns:
                df_c[c] = ""
                
        df_c_final = df_c[colunas_finais].copy()
        return df_c_final

    def _process_retorno(self, df_retorno, df_final, df_dem=None):
        self._log("Enriquecendo Arquivo de Retorno com dados da folha (12 colunas)...")
        df_r = df_retorno.copy()
        
        # O DataLoader já padronizou para "Matricula" (sem acento)
        df_r["_matricula_join"] = df_r["Matricula"].astype(str).str.strip().str.lstrip("0")
        
        # Extrair valor a ser importado para float preservando decimais se já for numérico
        if pd.api.types.is_numeric_dtype(df_r["Valor a ser importado"]):
            df_r["_valor_importar"] = df_r["Valor a ser importado"].fillna(0.0).round(2)
        else:
            df_r["_valor_importar"] = (
                df_r["Valor a ser importado"].astype(str)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
                .str.replace("-", "0", regex=False)
                .str.strip()
            )
            df_r["_valor_importar"] = pd.to_numeric(df_r["_valor_importar"], errors="coerce").fillna(0.0).round(2)
        
        # Mapear dados da Folha (df_final)
        # Regra validada pela equipe (Tamires - H1):
        # Na coluna "Folha" (H) apresentar os valores relacionados ao custo total de VT buscados na aba
        # "Conferência" coluna "X" ("Custo Total Vale Transporte")
        df_f_sub = df_final[[
            "Matrícula",
            "Custo Total Vale Transporte"
        ]].drop_duplicates(subset=["Matrícula"])
        
        df_r = df_r.merge(df_f_sub, left_on="_matricula_join", right_on="Matrícula", how="left")

        # Mapear Desligados diretamente de df_dem
        if df_dem is not None and not df_dem.empty and "Data Rescisão" in df_dem.columns:
            df_dem_sub = (
                df_dem[["Matrícula", "Data Rescisão"]]
                .drop_duplicates(subset=["Matrícula"])
                .rename(columns={"Matrícula": "_mat_dem", "Data Rescisão": "Desligado"})
            )
            df_r = df_r.merge(df_dem_sub, left_on="_matricula_join", right_on="_mat_dem", how="left")
        else:
            df_r["Desligado"] = ""
        df_r["Desligado"] = df_r["Desligado"].fillna("")
        
        # Coluna 8 (H): Folha (Custo Total Vale Transporte da Coluna X da aba Conferência)
        na_folha = df_r["Custo Total Vale Transporte"].notna()
        df_r["Folha"] = df_r["Custo Total Vale Transporte"].fillna(0.0).round(2)
        
        # Coluna 9 (I): Importado x Folha
        diff = (df_r["_valor_importar"] - df_r["Folha"]).abs()
        df_r["Importado x Folha"] = np.where(na_folha & (diff <= TOLERANCIA_CENTAVOS), "VERDADEIRO", "FALSO")
        
        # Coluna 11 (K): Obs
        condicoes_obs = [
            df_r["Importado x Folha"] == "VERDADEIRO",
            df_r["Desligado"].astype(str).str.strip() != "",
        ]
        escolhas_obs = [
            "Ok",
            "Desligado",
        ]
        df_r["Obs"] = np.select(condicoes_obs, escolhas_obs, default="Inconsistência: Importado != Folha")
        
        # Coluna 12 (L): Tratativa
        df_r["Tratativa"] = np.where(
            df_r["Obs"] == "Inconsistência: Importado != Folha",
            "Verificar diferença entre valor do Arquivo de Retorno e Custo Total VT na Folha.",
            ""
        )
        
        colunas_oficiais_retorno = [
            "Emp", "Loja", "Matricula", "Funcionario", "Evento",
            "Valor a ser importado", "Ocorrência", "Folha",
            "Importado x Folha", "Desligado", "Obs", "Tratativa"
        ]
        
        for c in colunas_oficiais_retorno:
            if c not in df_r.columns:
                df_r[c] = ""
                
        return df_r[colunas_oficiais_retorno].copy()
