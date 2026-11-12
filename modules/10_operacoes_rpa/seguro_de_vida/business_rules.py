"""
Motor de Regras de Negócio, Classificação de Planos e Conciliação Trilateral de Seguro de Vida (Cencosud Brasil).
Aplica rigorosamente:
1. Regras de elegibilidade e corte etário/admissão (Bradesco x Zurich).
2. Segregação auditada de Seguro Complementar 12x (Plano 05) e 24x (Plano 06).
3. Tratamento explícito de situações especiais: Afastados (INSS/Acidente), Férias, Menores Aprendizes e Desligados/Rescisões.
4. Cálculo e validação de capitais segurados, desconto em folha e custeio empresa.
5. Conciliação Trilateral (Nomina x FOPA x Fatura da Seguradora).
"""

from typing import Dict, Any, List, Tuple
import pandas as pd
import numpy as np

from config import (
    DATA_CORTE_ADESAO_IDADE,
    IDADE_LIMITE_ADESAO,
    PLANO_GBARBOSA,
    PLANO_PREZUNIC,
    PLANO_MERCANTIL,
    PLANO_GIGA,
    PLANO_PERINI,
    PLANO_BRETAS_LOJAS,
    PLANO_BRETAS_POSTOS,
    PLANO_COMP_12X,
    PLANO_COMP_24X,
    PLANO_JAGUAQUARA,
    PLANO_DIRETOR_12X,
    PLANO_DIRETOR_24X,
    SINDICATO_JAGUAQUARA
)


class BusinessRulesEngine:
    """Motor que executa todas as validações, cruzamentos e classificações de seguro de vida."""

    def __init__(self, raw_data: Dict[str, Any]):
        self.df_nomina = raw_data["nomina"].copy()
        self.df_oposicao = raw_data["oposicao"].copy()
        self.df_fopa = raw_data["fopa"].copy()
        self.optantes_comp = raw_data["optantes_complementar"]
        self.diretores = raw_data["diretores"]
        self.df_fatura = raw_data["fatura_seguradora"].copy()

    def run_all(self) -> Dict[str, Any]:
        """Executa o processamento completo de elegibilidade, conciliação e carga."""
        print("[*] Iniciando motor de regras de negócio e validações técnicas...")

        # 1. Cruzamento de Nomina com Folha (FOPA) e Oposições
        df_base = self._merge_sources()

        # 2. Avaliação de Regras de Exclusão, Elegibilidade e Situações Especiais
        df_base = self._apply_eligibility_and_special_situations(df_base)

        # 3. Classificação de Bandeira e Tipo de Plano
        df_base = self._assign_bandeira_and_plan(df_base)

        # 4. Cálculo de Capitais Segurados, Descontos Esperados e Custeio Empresa
        df_base = self._calculate_capital_and_discounts(df_base)

        # Contingência: inclui diretores manuais apenas se não vieram pela aba 'Com Gestor'
        df_base = self._append_diretores_to_base(df_base)

        # Garante Diretores ordenados no topo da Base Consolidada
        df_dir_top = df_base[df_base["is_diretor"]].sort_values(by="colaborador")
        df_rest = df_base[~df_base["is_diretor"]]
        df_base = pd.concat([df_dir_top, df_rest], ignore_index=True)

        # 5. Segregação e Montagem das Cargas (Normal, Comp 12x, Comp 24x e Diretores)
        df_seguro_normal, df_comp_12x, df_comp_24x = self._build_cargas(df_base)

        # 6. Conciliação Trilateral com a Fatura da Seguradora (Critério 3)
        conciliacao_fatura = self._reconcile_with_fatura(df_base, df_seguro_normal)

        # 7. Análise de Divergências de Folha
        divergencias_folha = self._analyze_folha_divergences(df_base)

        # 8. Consolidação dos KPIs Executivos
        kpis = self._calculate_kpis(df_base, df_seguro_normal, df_comp_12x, df_comp_24x, divergencias_folha, conciliacao_fatura)

        print("[+] Processamento de regras e conciliações concluído com sucesso.")
        return {
            "base_consolidada": df_base,
            "seguro_normal": df_seguro_normal,
            "comp_12x": df_comp_12x,
            "comp_24x": df_comp_24x,
            "exclusoes": df_base[df_base["status_conferencia"] == "EXCLUIDO"].copy(),
            "divergencias": divergencias_folha,
            "conciliacao_fatura": conciliacao_fatura,
            "kpis": kpis
        }

    def _merge_sources(self) -> pd.DataFrame:
        """Faz a junção segura entre Nomina (incluindo aba Com Gestor), Oposições e FOPA com base na matrícula."""
        df = self.df_nomina.copy()
        if "is_diretor" not in df.columns:
            df["is_diretor"] = df["natureza_profissional"].astype(str).str.strip().str.lower() == "diretores"
        else:
            df["is_diretor"] = df["is_diretor"].fillna(False).astype(bool)

        # Merge com Oposição
        opo_mats = set(self.df_oposicao["matricula"].dropna().astype(int))
        df["tem_oposicao"] = df["matricula"].isin(opo_mats)

        # Merge com FOPA (Descontos)
        if not self.df_fopa.empty and "matricula" in self.df_fopa.columns:
            df = df.merge(
                self.df_fopa[["matricula", "desconto_6266", "desconto_6267"]],
                on="matricula",
                how="left"
            )
            df["desconto_6266"] = df["desconto_6266"].fillna(0.0)
            df["desconto_6267"] = df["desconto_6267"].fillna(0.0)
        else:
            df["desconto_6266"] = 0.0
            df["desconto_6267"] = 0.0

        return df

    def _apply_eligibility_and_special_situations(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Aplica as regras de elegibilidade, cortes etários e tratamento de situações especiais:
        1. Rescisões/Desligados: fora da carga de vidas ativas.
        2. Idade > 65 anos admitido após 31/12/2015: fora da carga (acordo Bradesco x Zurich).
        3. Intermitentes: fora da carga.
        4. Horistas: fora da carga.
        5. Autônomos: fora da carga.
        6. Oposição formal ao SGV: fora da carga.
        7. Estagiários: MANTIDOS mesmo sem desconto em folha (isenção corporativa).
        8. Menores Aprendizes: CLT regulares elegíveis (se sem oposição).
        9. Afastados (INSS/Acidente) e Férias: cobertura mantida pela política de benefícios.
        """
        # --- Identificação de Situações Especiais ---
        # 1. Estagiários
        is_estag = (
            df["categoria"].str.contains("estag", case=False, na=False) |
            df["cargo"].str.contains("estag", case=False, na=False) |
            df["vinculo"].str.contains("estag", case=False, na=False)
        )
        df["is_estagiario"] = is_estag

        # 2. Menores Aprendizes
        is_aprendiz = (
            df["categoria"].str.contains("aprendiz", case=False, na=False) |
            df["cargo"].str.contains("aprendiz", case=False, na=False)
        )
        df["is_aprendiz"] = is_aprendiz

        # 3. Afastados (INSS, Doença, Acidente de Trabalho, Licença Maternidade)
        is_afastado = (
            df["tipo_situacao"].str.contains("afast", case=False, na=False) |
            df["situacao"].str.contains("afast", case=False, na=False) |
            df["situacao"].str.contains("licen", case=False, na=False)
        )
        df["is_afastado"] = is_afastado

        # 4. Colaboradores em Férias
        is_ferias = (
            df["tipo_situacao"].str.contains("férias", case=False, na=False) |
            df["tipo_situacao"].str.contains("ferias", case=False, na=False) |
            df["situacao"].str.contains("férias", case=False, na=False) |
            df["situacao"].str.contains("ferias", case=False, na=False)
        )
        df["is_ferias"] = is_ferias

        # 5. Desligados / Rescisões Efetivadas
        is_desligado = (
            df["data_rescisao"].notna() |
            df["tipo_situacao"].str.contains("deslig", case=False, na=False) |
            df["tipo_situacao"].str.contains("rescis", case=False, na=False) |
            df["situacao"].str.contains("rescis", case=False, na=False)
        )
        df["is_desligado"] = is_desligado
        df["motivo_exclusao_desligado"] = np.where(
            is_desligado,
            "Colaborador Desligado / Rescisão registrada",
            ""
        )

        # --- Regras de Exclusão da Carga ---
        # 6. Idade > 65 anos admitido após 31/12/2015
        dt_corte = pd.to_datetime(DATA_CORTE_ADESAO_IDADE)
        is_idoso_pos2015 = (df["idade"] > IDADE_LIMITE_ADESAO) & (df["data_admissao"] > dt_corte)
        df["motivo_exclusao_idoso"] = np.where(
            is_idoso_pos2015,
            f"Admitido após 2015 com idade > {IDADE_LIMITE_ADESAO} anos (Regra Bradesco x Zurich)",
            ""
        )

        # 7. Intermitentes
        is_intermitente = (
            df["categoria"].str.contains("intermitente", case=False, na=False) |
            df["vinculo"].str.contains("intermitente", case=False, na=False)
        )
        df["motivo_exclusao_intermitente"] = np.where(
            is_intermitente,
            "Contrato Intermitente (Fica fora da carga)",
            ""
        )

        # 8. Horistas
        is_horista = df["tipo_salario"].str.contains("hor", case=False, na=False)
        df["motivo_exclusao_horista"] = np.where(
            is_horista,
            "Colaborador Horista (Fica fora da carga)",
            ""
        )

        # 9. Autônomos
        is_autonomo = (
            df["categoria"].str.contains("autonomo", case=False, na=False) |
            df["vinculo"].str.contains("autonomo", case=False, na=False)
        )
        df["motivo_exclusao_autonomo"] = np.where(
            is_autonomo,
            "Contrato Autônomo (Fica fora da carga)",
            ""
        )

        # 10. Oposição Formal SGV
        is_oposicao = df["tem_oposicao"]
        df["motivo_exclusao_oposicao"] = np.where(
            is_oposicao,
            "Oposição formal ao desconto do seguro (SGV)",
            ""
        )

        # Máscara de Exclusão Consolidada:
        # Estagiários são protegidos de cortes por falta de desconto, mas se tiverem rescisão são excluídos
        excluir_mask = (
            is_desligado |
            ((is_idoso_pos2015 | is_intermitente | is_horista | is_autonomo | is_oposicao) & (~is_estag))
        )

        df["status_conferencia"] = np.where(excluir_mask, "EXCLUIDO", "ELEGIVEL")

        # Motivo detalhado e rótulo de situação especial
        motivos = []
        situacoes_especiais = []

        for idx, row in df.iterrows():
            # Rótulo de situação especial
            sit_esp = []
            if row["is_diretor"]:
                sit_esp.append("Diretoria Executiva (Aba Com Gestor)")
            if row["is_estagiario"]:
                sit_esp.append("Estagiário (Isenção Corporativa)")
            if row["is_aprendiz"]:
                sit_esp.append("Menor Aprendiz (CLT)")
            if row["is_afastado"]:
                sit_esp.append("Afastado INSS/Acidente (Cobertura Mantida)")
            if row["is_ferias"]:
                sit_esp.append("Férias (Cobertura Mantida)")
            if row["is_desligado"]:
                sit_esp.append("Rescisão/Desligado")
            situacoes_especiais.append("; ".join(sit_esp) if sit_esp else "Atividade Normal")

            # Motivo de conferência
            if row["status_conferencia"] == "ELEGIVEL":
                if row["is_diretor"]:
                    motivos.append("Elegível (Diretor localizado na aba 'Com Gestor' - Seguro Básico GB PLANO 01)")
                elif row["is_estagiario"]:
                    motivos.append("Elegível (Estagiário mantido por benefício corporativo)")
                elif row["is_afastado"]:
                    motivos.append("Elegível (Afastado mantido por apólice corporativa)")
                elif row["is_ferias"]:
                    motivos.append("Elegível (Colaborador em gozo de férias)")
                elif row["is_aprendiz"]:
                    motivos.append("Elegível (Menor Aprendiz CLT ativo)")
                else:
                    motivos.append("Elegível")
            else:
                m_list = [
                    row["motivo_exclusao_desligado"],
                    row["motivo_exclusao_idoso"],
                    row["motivo_exclusao_intermitente"],
                    row["motivo_exclusao_horista"],
                    row["motivo_exclusao_autonomo"],
                    row["motivo_exclusao_oposicao"]
                ]
                m_list = [m for m in m_list if m]
                motivos.append("; ".join(m_list) if m_list else "Excluído por regra interna")

        df["situacao_especial"] = situacoes_especiais
        df["motivo_detalhado"] = motivos
        return df

    def _assign_bandeira_and_plan(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Atribui a Bandeira, o Tipo de Plano e o arquivo de carga destino:
        - Diretores (aba Com Gestor): PLANO 01 (arquivo Gbarbosa e Jaguaquara)
        - GBarbosa: PLANO 01 (exceto Jaguaquara)
        - Jaguaquara (Sindicato Jaguaquara): PLANO 07 (arquivo Gbarbosa e Jaguaquara)
        - Bretas Lojas: PLANO 03
        - Bretas Postos: PLANO 04 (Estabelecimento contendo POSTO)
        - Prezunic (com SPID): PLANO 01
        - Giga e Mercantil: PLANO 01
        - Perini: PLANO 01
        """
        bandeira_list = []
        plano_list = []
        arquivo_carga_list = []

        for idx, row in df.iterrows():
            natureza = str(row["natureza_profissional"]).strip().upper()
            estab = str(row["estabelecimento"]).strip().upper()
            sindicato = str(row["sindicato"]).strip().upper()

            if row["is_diretor"] or "DIRETOR" in natureza:
                bandeira_list.append("DIRETORES (GBARBOSA)")
                plano_list.append(PLANO_GBARBOSA)
                arquivo_carga_list.append("Gbarbosa e Jaguaquara")

            elif SINDICATO_JAGUAQUARA in sindicato:
                bandeira_list.append("JAGUAQUARA")
                plano_list.append(PLANO_JAGUAQUARA)
                arquivo_carga_list.append("Gbarbosa e Jaguaquara")

            elif "BRETAS" in natureza:
                if "POSTO" in estab:
                    bandeira_list.append("BRETAS POSTOS")
                    plano_list.append(PLANO_BRETAS_POSTOS)
                    arquivo_carga_list.append("Bretas Postos")
                else:
                    bandeira_list.append("BRETAS LOJAS")
                    plano_list.append(PLANO_BRETAS_LOJAS)
                    arquivo_carga_list.append("Bretas Lojas")

            elif "GBARBOSA" in natureza:
                bandeira_list.append("GBARBOSA")
                plano_list.append(PLANO_GBARBOSA)
                arquivo_carga_list.append("Gbarbosa e Jaguaquara")

            elif "PREZUNIC" in natureza or "SPID" in natureza:
                bandeira_list.append("PREZUNIC")
                plano_list.append(PLANO_PREZUNIC)
                arquivo_carga_list.append("Prezunic")

            elif "GIGA" in natureza or "MERCANTIL" in natureza:
                bandeira_list.append("GIGA E MERCANTIL")
                plano_list.append(PLANO_GIGA)
                arquivo_carga_list.append("Giga e Mercantil")

            elif "PERINI" in natureza:
                bandeira_list.append("PERINI")
                plano_list.append(PLANO_PERINI)
                arquivo_carga_list.append("Perini")

            else:
                bandeira_list.append(natureza if natureza else "CORPORATIVO")
                plano_list.append(PLANO_GBARBOSA)
                arquivo_carga_list.append("Gbarbosa e Jaguaquara")

        df["bandeira"] = bandeira_list
        df["plano_atribuido"] = plano_list
        df["arquivo_destino"] = arquivo_carga_list
        return df

    def _calculate_capital_and_discounts(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Calcula os capitais segurados estimados, descontos em folha e custeio empresa
        conforme apólice corporativa e CCT (Critério 1 da Auditoria).
        """
        sal_num = df["salario_atual"].fillna(0.0)
        capital_normal = np.where(
            df["is_diretor"],
            np.maximum(sal_num * 12, 100000.0),
            np.maximum(sal_num * 12, 20000.0)
        )
        df["capital_segurado_normal"] = np.where(df["status_conferencia"] == "ELEGIVEL", capital_normal, 0.0)

        # Custeio Empresa Estimado (Prêmio total médio R$ 5,50 por vida)
        custo_premio_vida = 5.50
        df["premio_total_estimado"] = np.where(df["status_conferencia"] == "ELEGIVEL", custo_premio_vida, 0.0)
        df["custeio_empresa_estimado"] = np.maximum(df["premio_total_estimado"] - df["desconto_6266"], 0.0)

        # Checagem de conformidade do desconto
        df["conformidade_desconto"] = np.where(
            df["status_conferencia"] == "EXCLUIDO",
            np.where(df["desconto_6266"] > 0, "NÃO CONFORME (Desconto em colaborador excluído)", "CONFORME"),
            np.where(
                df["is_diretor"],
                "CONFORME (Diretoria - Aba Com Gestor)",
                np.where(
                    df["is_estagiario"],
                    "CONFORME (Isenção Estagiário)",
                    np.where(df["desconto_6266"] > 0, "CONFORME (Descontado)", "PENDENTE (Sem desconto no FOPA)")
                )
            )
        )
        return df

    def _append_diretores_to_base(self, df_base: pd.DataFrame) -> pd.DataFrame:
        """
        Contingência: caso algum diretor venha de cadastro manual (quando a aba 'Com Gestor'
        não estiver presente no arquivo), insere apenas matrículas ausentes da base.
        """
        if not self.diretores:
            return df_base

        mats_existentes = set(df_base["matricula"].dropna().astype(int))
        diretores_faltantes = [d for d in self.diretores if int(d["matricula"]) not in mats_existentes]
        if not diretores_faltantes:
            return df_base

        diretores_base_rows = []
        for d in diretores_faltantes:
            diretores_base_rows.append({
                "matricula": int(d["matricula"]),
                "colaborador": str(d["nome"]).strip().upper(),
                "cpf": d["cpf"],
                "data_nascimento": pd.to_datetime(d["nascimento"], dayfirst=True),
                "idade": 0,
                "data_admissao": pd.NaT,
                "cargo": d.get("cargo", "DIRETOR"),
                "estabelecimento": "CORPORATIVO",
                "natureza_profissional": "Diretores",
                "sindicato": "DIRETORIA",
                "salario_atual": 0.0,
                "categoria": "DIRETORIA",
                "vinculo": "DIRETOR ESTATUTARIO",
                "tipo_salario": "Mensal",
                "tipo_situacao": "Atividade Normal",
                "situacao": "Atividade Normal",
                "data_rescisao": pd.NaT,
                "motivo_desligamento": "",
                "nivel_hierarquico": "Executivo",
                "is_diretor": True,
                "tem_oposicao": False,
                "desconto_6266": 0.0,
                "desconto_6267": 0.0,
                "fopa_valida": True,
                "is_estagiario": False,
                "is_aprendiz": False,
                "is_afastado": False,
                "is_ferias": False,
                "is_desligado": False,
                "motivo_exclusao_idoso": "",
                "motivo_exclusao_intermitente": "",
                "motivo_exclusao_horista": "",
                "motivo_exclusao_autonomo": "",
                "motivo_exclusao_oposicao": "",
                "motivo_exclusao_desligado": "",
                "status_conferencia": "ELEGIVEL",
                "situacao_especial": "Diretoria Executiva (Contingência)",
                "motivo_detalhado": "Elegível (Diretor - Contingência Seguro Básico GB PLANO 01)",
                "bandeira": "DIRETORES (GBARBOSA)",
                "plano_atribuido": d.get("plano_normal", PLANO_GBARBOSA),
                "arquivo_destino": "Gbarbosa e Jaguaquara",
                "capital_segurado_normal": 100000.0,
                "premio_total_estimado": 5.50,
                "custeio_empresa_estimado": 5.50,
                "conformidade_desconto": "CONFORME (Diretoria)"
            })

        df_dir = pd.DataFrame(diretores_base_rows)
        df_base = pd.concat([df_dir, df_base], ignore_index=True).drop_duplicates(subset=["matricula"])
        return df_base

    def _build_cargas(
        self,
        df_base: pd.DataFrame
    ) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """
        Segrega as bases de carga para Seguro Normal e Seguro Complementar:
        - Seguro Normal: Vidas Elegíveis + Diretores (da aba Com Gestor) fixados no topo de GB (PLANO 01).
        - Seguro Complementar: Segregação auditada entre 12x (Plano 05 / Dir Plano 08) e 24x (Plano 06 / Dir Plano 09),
          apenas para vidas ativas/elegíveis no mês corrente.
        """
        df_normal = df_base[df_base["status_conferencia"] == "ELEGIVEL"].copy()

        # Garante que todos os Diretores fiquem ordenados alfabeticamente no topo absoluto
        df_dir = df_normal[df_normal["is_diretor"]].sort_values(by="colaborador")
        df_demais = df_normal[~df_normal["is_diretor"]]
        df_normal = pd.concat([df_dir, df_demais], ignore_index=True).drop_duplicates(subset=["matricula"])

        # --- SEGURO COMPLEMENTAR: Segregação 12x e 24x unificada sobre vidas elegíveis ---
        comp_12x_rows = []
        comp_24x_rows = []

        # Ordena com Diretores primeiro (em ordem alfabética) e colaboradores regulares em seguida
        colabs_elegiveis = pd.concat([df_dir, df_demais], ignore_index=True)

        for idx, row in colabs_elegiveis.iterrows():
            mat = int(row["matricula"])
            sal = row["salario_atual"]
            tem_desc_comp = row["desconto_6267"] > 0
            esta_no_catalogo = mat in self.optantes_comp

            if esta_no_catalogo or tem_desc_comp:
                info_cat = self.optantes_comp.get(mat, {})
                tipo_optante = str(info_cat.get("tipo", "12x")).lower()

                if row["is_diretor"]:
                    sal_out = float(sal) if (pd.notna(sal) and float(sal) > 0) else np.nan
                    origem = "Diretoria (Aba Com Gestor)"
                    if "12" in tipo_optante:
                        plano = info_cat.get("plano", PLANO_DIRETOR_12X)
                        cap_est = sal_out * 12 if pd.notna(sal_out) else np.nan
                        comp_12x_rows.append({
                            "matricula": mat,
                            "nome": row["colaborador"],
                            "nascimento": row["data_nascimento"],
                            "cpf": row["cpf"],
                            "salario": sal_out,
                            "plano": plano,
                            "capital_segurado_estimado": cap_est,
                            "origem": origem
                        })
                    else:
                        plano = info_cat.get("plano", PLANO_DIRETOR_24X)
                        cap_est = sal_out * 24 if pd.notna(sal_out) else np.nan
                        comp_24x_rows.append({
                            "matricula": mat,
                            "nome": row["colaborador"],
                            "nascimento": row["data_nascimento"],
                            "cpf": row["cpf"],
                            "salario": sal_out,
                            "plano": plano,
                            "capital_segurado_estimado": cap_est,
                            "origem": origem
                        })
                else:
                    sal_out = float(sal) if pd.notna(sal) else 0.0
                    origem = "Colaborador Regular"
                    if "12" in tipo_optante:
                        comp_12x_rows.append({
                            "matricula": mat,
                            "nome": row["colaborador"],
                            "nascimento": row["data_nascimento"],
                            "cpf": row["cpf"],
                            "salario": sal_out,
                            "plano": PLANO_COMP_12X,
                            "capital_segurado_estimado": sal_out * 12,
                            "origem": origem
                        })
                    else:
                        comp_24x_rows.append({
                            "matricula": mat,
                            "nome": row["colaborador"],
                            "nascimento": row["data_nascimento"],
                            "cpf": row["cpf"],
                            "salario": sal_out,
                            "plano": PLANO_COMP_24X,
                            "capital_segurado_estimado": sal_out * 24,
                            "origem": origem
                        })

        df_12x = pd.DataFrame(comp_12x_rows).drop_duplicates(subset=["matricula"]) if comp_12x_rows else pd.DataFrame(columns=["matricula", "nome", "nascimento", "cpf", "salario", "plano"])
        df_24x = pd.DataFrame(comp_24x_rows).drop_duplicates(subset=["matricula"]) if comp_24x_rows else pd.DataFrame(columns=["matricula", "nome", "nascimento", "cpf", "salario", "plano"])

        return df_normal, df_12x, df_24x

    def _reconcile_with_fatura(self, df_base: pd.DataFrame, df_normal: pd.DataFrame) -> Dict[str, pd.DataFrame]:
        """
        Conciliação Trilateral: Vidas Ativas Cencosud x Fatura da Seguradora/Corretora (Critério 3).
        Identifica:
        1. Vidas Faturadas que estão Desligadas/Inelegíveis na Empresa -> Cobrança Indevida / Exclusão Pendente na Seguradora.
        2. Vidas Ativas Elegíveis na Empresa que NÃO constam na Fatura -> Inclusão Pendente na Seguradora.
        3. Vidas em Conformidade (Ativa e Faturada).
        """
        mats_elegiveis = set(df_normal["matricula"].dropna().astype(int))
        mats_desligados = set(df_base[df_base["is_desligado"]]["matricula"].dropna().astype(int))
        mats_excluidos = set(df_base[df_base["status_conferencia"] == "EXCLUIDO"]["matricula"].dropna().astype(int))

        df_fat = self.df_fatura.copy()

        # 1. Cobrança Indevida: Seguradora faturou vidas que são desligadas ou excluídas por regra
        cobranca_indevida = df_fat[df_fat["matricula"].isin(mats_excluidos)].copy()
        if not cobranca_indevida.empty:
            cobranca_indevida = cobranca_indevida.merge(
                df_base[["matricula", "colaborador", "cpf", "bandeira", "cargo", "motivo_detalhado"]],
                on="matricula",
                how="left"
            )
            cobranca_indevida["tipo_conciliacao"] = "Cobrança Indevida na Fatura (Vidas Inelegíveis/Desligadas)"
            cobranca_indevida["acao_sugerida"] = "Solicitar estorno/glosa do prêmio à Seguradora"
        else:
            cobranca_indevida = pd.DataFrame(columns=["matricula", "colaborador", "cpf", "bandeira", "cargo", "premio_faturado", "tipo_conciliacao", "acao_sugerida"])

        # 2. Inclusões Pendentes: Vidas ativas elegíveis na empresa que não foram faturadas pela seguradora
        mats_faturadas = set(df_fat["matricula"].dropna().astype(int))
        inclusoes_pendentes = df_normal[~df_normal["matricula"].isin(mats_faturadas)].copy()
        if not inclusoes_pendentes.empty:
            inclusoes_pendentes["tipo_conciliacao"] = "Inclusão Pendente na Fatura (Vidas Elegíveis não Faturadas)"
            inclusoes_pendentes["acao_sugerida"] = "Enviar movimento de inclusão para a Seguradora"
            inclusoes_pendentes["premio_faturado"] = 0.0
        else:
            inclusoes_pendentes = pd.DataFrame(columns=["matricula", "colaborador", "cpf", "bandeira", "cargo", "premio_faturado", "tipo_conciliacao", "acao_sugerida"])

        # 3. Vidas em Conformidade
        conformes = df_normal[df_normal["matricula"].isin(mats_faturadas)].copy()
        conformes["tipo_conciliacao"] = "Vidas em Plena Conformidade (Ativa na Carga e Faturada)"

        return {
            "cobranca_indevida": cobranca_indevida,
            "inclusoes_pendentes": inclusoes_pendentes,
            "conformes": conformes
        }

    def _analyze_folha_divergences(self, df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
        """Audita batimento entre Folha de Pagamento (FOPA) e as regras da apólice."""
        # 1. Desconto indevido em colaborador excluído
        teve_desc_excluido = df[(df["desconto_6266"] > 0) & (df["status_conferencia"] == "EXCLUIDO")].copy()
        teve_desc_excluido["tipo_divergencia"] = "Desconto em folha realizado para colaborador inelegível"

        # 2. Oposição ativa mas com desconto
        oposicao_com_desconto = df[(df["tem_oposicao"]) & (df["desconto_6266"] > 0)].copy()
        oposicao_com_desconto["tipo_divergencia"] = "Colaborador com oposição ao SGV descontado em folha"

        # 3. Elegível ativo na Nomina sem desconto em folha (exceto estagiários e diretores)
        ativo_sem_desconto = df[
            (df["status_conferencia"] == "ELEGIVEL") &
            (df["desconto_6266"] == 0) &
            (~df["is_estagiario"]) &
            (~df["is_diretor"])
        ].copy()
        ativo_sem_desconto["tipo_divergencia"] = "Colaborador ativo sem desconto do evento 6266"

        return {
            "desconto_indevido": teve_desc_excluido,
            "oposicao_descontada": oposicao_com_desconto,
            "ativo_sem_desconto": ativo_sem_desconto
        }

    def _calculate_kpis(
        self,
        df_base: pd.DataFrame,
        df_normal: pd.DataFrame,
        df_12x: pd.DataFrame,
        df_24x: pd.DataFrame,
        divergencias_folha: Dict[str, pd.DataFrame],
        conciliacao_fatura: Dict[str, pd.DataFrame]
    ) -> Dict[str, Any]:
        """Calcula os indicadores consolidados para o Resumo Executivo e Auditoria."""
        total_nomina = len(df_base)
        total_elegiveis = len(df_normal)
        total_diretores_elegiveis = int(df_normal["is_diretor"].sum()) if "is_diretor" in df_normal.columns else 0
        total_excluidos = len(df_base[df_base["status_conferencia"] == "EXCLUIDO"])
        total_estagiarios = len(df_base[df_base["is_estagiario"]])
        total_aprendizes = len(df_base[df_base["is_aprendiz"]])
        total_afastados = len(df_base[df_base["is_afastado"]])
        total_ferias = len(df_base[df_base["is_ferias"]])
        total_desligados = len(df_base[df_base["is_desligado"]])

        excl_idosos = len(df_base[df_base["motivo_exclusao_idoso"] != ""])
        excl_inter = len(df_base[df_base["motivo_exclusao_intermitente"] != ""])
        excl_hor = len(df_base[df_base["motivo_exclusao_horista"] != ""])
        excl_aut = len(df_base[df_base["motivo_exclusao_autonomo"] != ""])
        excl_opo = len(df_base[df_base["motivo_exclusao_oposicao"] != ""])
        excl_deslig = len(df_base[df_base["motivo_exclusao_desligado"] != ""])

        bandeira_counts = df_normal["bandeira"].value_counts().to_dict()

        return {
            "total_nomina": total_nomina,
            "total_elegiveis": total_elegiveis,
            "total_diretores_elegiveis": total_diretores_elegiveis,
            "total_excluidos": total_excluidos,
            "total_estagiarios_mantidos": total_estagiarios,
            "total_aprendizes": total_aprendizes,
            "total_afastados_mantidos": total_afastados,
            "total_ferias_mantidos": total_ferias,
            "total_desligados_excluidos": total_desligados,
            "total_comp_12x": len(df_12x),
            "total_comp_24x": len(df_24x),
            "exclusoes_por_motivo": {
                "Rescisões / Desligados": excl_deslig,
                "Idade > 65 anos pós-2015": excl_idosos,
                "Contratos Intermitentes": excl_inter,
                "Horistas": excl_hor,
                "Autônomos": excl_aut,
                "Oposição formal SGV": excl_opo
            },
            "contagem_por_bandeira": bandeira_counts,
            "divergencias_folha": {
                "descontos_indevidos": len(divergencias_folha["desconto_indevido"]),
                "oposicoes_descontadas": len(divergencias_folha["oposicao_descontada"]),
                "ativos_sem_desconto": len(divergencias_folha["ativo_sem_desconto"])
            },
            "conciliacao_fatura": {
                "cobrancas_indevidas": len(conciliacao_fatura["cobranca_indevida"]),
                "inclusoes_pendentes": len(conciliacao_fatura["inclusoes_pendentes"]),
                "vidas_em_conformidade": len(conciliacao_fatura["conformes"])
            }
        }
