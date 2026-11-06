"""
Motor de Regras de Negócio para Auditoria de Valores Pré-Folha VT (Competência 09/2026).
Valida:
1. Equilíbrio contábil entre os 10 Eventos de Crédito/Pagamento Pré-Folha (147, 148, 159,
   490, 9421, 9422, 9423, 9428, 9429 e 31935) e a rubrica de saída 5750 (Adiantamento VT Através Alçadas).
2. Conciliação da importação do Evento 31935 (VT Alternativo Mês Seguinte - 387 registros).
3. Conciliação da importação do Evento 6292 (Desconto Cartão VTE - 30 registros).
"""

import sys
from pathlib import Path
from typing import Any, Dict, List

import numpy as np
import pandas as pd

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
if str(SCRIPT_DIR.parent) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR.parent))

try:
    from config import CONFERENCIA_COLUMNS, EVENTOS_SOMA_5750
except ImportError:
    from scripts.config import CONFERENCIA_COLUMNS, EVENTOS_SOMA_5750


class BusinessRulesPrefolha:
    """Processa e audita os eventos de Valores Pré-Folha VT."""

    def process_conferencia(self, data: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
        """Executa o cruzamento integral entre FOPA Pré-Folha e importações 31935 / 6292."""
        df = data["df_fopa"].copy()
        df_31935 = data["df_31935"].copy()
        df_6292 = data["df_6292"].copy()

        # 1. Cálculo da Soma dos 10 Eventos de Pagamento que devem espelhar o Evento 5750
        df["Soma Eventos Pagamento"] = df[EVENTOS_SOMA_5750].sum(axis=1).round(2)
        df["Diferença (Soma x 5750)"] = (
            df["Soma Eventos Pagamento"] - df["5750 - Adiantamento VT Atraves Alcadas"]
        ).round(2)

        match_5750 = df["Diferença (Soma x 5750)"].abs() < 0.01
        df["Evts Pgtos x Evt 5750"] = np.where(match_5750, "VERDADEIRO", "FALSO")

        # 2. Mapeamento das bases de importação 31935 e 6292 na FOPA
        map_31935 = (
            df_31935.groupby("Matrícula")["Valor Importado 31935"].sum().to_dict()
            if not df_31935.empty
            else {}
        )
        map_6292 = (
            df_6292.groupby("Matrícula")["Valor Importado 6292"].sum().to_dict()
            if not df_6292.empty
            else {}
        )

        df["Importado 31935"] = df["Matrícula"].map(map_31935).fillna(0.0).round(2)
        df["Importado 6292"] = df["Matrícula"].map(map_6292).fillna(0.0).round(2)

        # 3. Classificação de Status Auditoria e Observação na FOPA
        statuses: List[str] = []
        obs_list: List[str] = []

        for _, row in df.iterrows():
            diff = float(row["Diferença (Soma x 5750)"])
            orig_obs = str(row.get("Obs_Original", "Ok") or "Ok").strip()
            imp_31935 = float(row["Importado 31935"])
            fopa_31935 = float(row["31935 - Valor Vale Transporte Alternativo Mes Seguinte"])
            imp_6292 = float(row["Importado 6292"])
            fopa_6292 = float(row["6292 - Desconto Cartao VTE"])

            if abs(diff) >= 0.01:
                statuses.append("Divergente - Evento 5750")
                if orig_obs and orig_obs != "Ok":
                    obs_list.append(orig_obs)
                else:
                    obs_list.append(
                        f"Sinalizar à Quality - Diferença de R$ {diff:,.2f} entre Soma Pagamentos e Evento 5750."
                    )
            elif imp_31935 > 0 and abs(imp_31935 - fopa_31935) >= 0.01:
                statuses.append("Divergente - Importação 31935")
                obs_list.append(
                    f"Valor importado 31935 (R$ {imp_31935:,.2f}) difere da FOPA (R$ {fopa_31935:,.2f})."
                )
            elif imp_6292 > 0 and abs(imp_6292 - fopa_6292) >= 0.01:
                statuses.append("Divergente - Importação 6292")
                obs_list.append(
                    f"Valor importado 6292 (R$ {imp_6292:,.2f}) difere da FOPA (R$ {fopa_6292:,.2f})."
                )
            else:
                statuses.append("Conciliado (Ok)")
                obs_list.append("Ok")

        df["Status Auditoria"] = statuses
        df["Obs."] = obs_list

        # 4. Identificação de colaboradores das abas 31935 e 6292 ausentes na FOPA (ex.: Desligados)
        fopa_mats = set(df["Matrícula"].unique())
        extra_divergencias: List[Dict[str, Any]] = []

        for _, r31 in df_31935.iterrows():
            mat = int(r31["Matrícula"])
            val_imp = float(r31["Valor Importado 31935"])
            fopa_val = r31["Fopa Mensal 31935"]
            obs_31 = str(r31["Obs 31935"] or "Ok")
            if mat not in fopa_mats or fopa_val is None or abs(val_imp - float(fopa_val)) >= 0.01:
                extra_divergencias.append(
                    {
                        "Código da Empresa": r31["Emp"],
                        "Nome da Empresa": "CENCOSUD BRASIL COMERCIAL LTDA.",
                        "Código Estabelecimento": None,
                        "Estabelecimento": "Base Importação 31935 (Ausente FOPA Ativa)",
                        "Código Centro Custo": "-",
                        "Centro Custo": "-",
                        "Matrícula": mat,
                        "Nome do Colaborador": "COLABORADOR DESLIGADO (VER RESCISÃO)",
                        "ID Colaborador": None,
                        "Data Admissão": None,
                        "Código Cargo": None,
                        "Cargo": "-",
                        "Situação": "Demitido / Rescisão",
                        "Tipo Situação": "Demitido",
                        "Data Rescisão": None,
                        "118 - Credito Vale Transporte": 0.0,
                        "147 - Credito Vale Transporte Dev Operadora FL 40": 0.0,
                        "148 - Credito Vale Transporte Dev Operadora FL 41": 0.0,
                        "159 - Deslocamento Atividades Externa": 0.0,
                        "490 - Credito Vale Transporte": 0.0,
                        "5731 - Saldo Vale Transporte Rescisao": 0.0,
                        "5750 - Adiantamento VT Atraves Alcadas": 0.0,
                        "6292 - Desconto Cartao VTE": 0.0,
                        "9290 - Dev Vale Transporte": 0.0,
                        "9421 - Reembolso VT Sem Dedução - Folha 40": 0.0,
                        "9422 - Reembolso VT Sem Dedução - Folha 41": 0.0,
                        "9423 - Reembolso VT Sem Dedução - Folha 42": 0.0,
                        "9424 - Credito Vale Transporte - Folha 41": 0.0,
                        "9425 - Credito Vale Transporte - Folha 42": 0.0,
                        "9426 - Dev Vale Transporte - Folha 41": 0.0,
                        "9427 - Dev Vale Transporte - Folha 42": 0.0,
                        "9428 - Deslocamento Atividades Externa - Folha 41": 0.0,
                        "9429 - Deslocamento Atividades Externa - Folha 42": 0.0,
                        "31935 - Valor Vale Transporte Alternativo Mes Seguinte": 0.0,
                        "Soma Eventos Pagamento": 0.0,
                        "Diferença (Soma x 5750)": 0.0,
                        "Evts Pgtos x Evt 5750": "N/A (Rescisão)",
                        "Importado 31935": val_imp,
                        "Importado 6292": 0.0,
                        "Status Auditoria": "Não Importado - Rescisão (31935)",
                        "Obs.": obs_31 if obs_31 != "Ok" else "Desligado - Valor Transitando em rescisão",
                    }
                )

        for _, r62 in df_6292.iterrows():
            mat = int(r62["Matrícula"])
            val_imp = float(r62["Valor Importado 6292"])
            fopa_val = r62["Fopa Mensal 6292"]
            obs_62 = str(r62["Obs 6292"] or "Ok")
            if mat not in fopa_mats or fopa_val is None or abs(val_imp - float(fopa_val)) >= 0.01:
                extra_divergencias.append(
                    {
                        "Código da Empresa": r62["Emp"],
                        "Nome da Empresa": "CENCOSUD BRASIL COMERCIAL LTDA.",
                        "Código Estabelecimento": r62["Loja"],
                        "Estabelecimento": f"LOJA {r62['Loja']} (Base Importação 6292)",
                        "Código Centro Custo": "-",
                        "Centro Custo": "-",
                        "Matrícula": mat,
                        "Nome do Colaborador": r62["Colaborador"],
                        "ID Colaborador": None,
                        "Data Admissão": None,
                        "Código Cargo": None,
                        "Cargo": "-",
                        "Situação": "Demitido / Rescisão",
                        "Tipo Situação": "Demitido",
                        "Data Rescisão": None,
                        "118 - Credito Vale Transporte": 0.0,
                        "147 - Credito Vale Transporte Dev Operadora FL 40": 0.0,
                        "148 - Credito Vale Transporte Dev Operadora FL 41": 0.0,
                        "159 - Deslocamento Atividades Externa": 0.0,
                        "490 - Credito Vale Transporte": 0.0,
                        "5731 - Saldo Vale Transporte Rescisao": 0.0,
                        "5750 - Adiantamento VT Atraves Alcadas": 0.0,
                        "6292 - Desconto Cartao VTE": 0.0,
                        "9290 - Dev Vale Transporte": 0.0,
                        "9421 - Reembolso VT Sem Dedução - Folha 40": 0.0,
                        "9422 - Reembolso VT Sem Dedução - Folha 41": 0.0,
                        "9423 - Reembolso VT Sem Dedução - Folha 42": 0.0,
                        "9424 - Credito Vale Transporte - Folha 41": 0.0,
                        "9425 - Credito Vale Transporte - Folha 42": 0.0,
                        "9426 - Dev Vale Transporte - Folha 41": 0.0,
                        "9427 - Dev Vale Transporte - Folha 42": 0.0,
                        "9428 - Deslocamento Atividades Externa - Folha 41": 0.0,
                        "9429 - Deslocamento Atividades Externa - Folha 42": 0.0,
                        "31935 - Valor Vale Transporte Alternativo Mes Seguinte": 0.0,
                        "Soma Eventos Pagamento": 0.0,
                        "Diferença (Soma x 5750)": 0.0,
                        "Evts Pgtos x Evt 5750": "N/A (Rescisão)",
                        "Importado 31935": 0.0,
                        "Importado 6292": val_imp,
                        "Status Auditoria": "Não Importado - Rescisão (6292)",
                        "Obs.": obs_62 if obs_62 != "Ok" else "Desligado - Não descontado em folha",
                    }
                )

        df_geral = df[CONFERENCIA_COLUMNS].copy()
        df_fopa_div = df_geral[df_geral["Status Auditoria"] != "Conciliado (Ok)"].copy()
        df_extra_div = (
            pd.DataFrame(extra_divergencias)[CONFERENCIA_COLUMNS]
            if extra_divergencias
            else pd.DataFrame(columns=CONFERENCIA_COLUMNS)
        )
        df_divergencias = pd.concat([df_fopa_div, df_extra_div], ignore_index=True)

        # 5. Consolidação de KPIs Executivos
        kpis = self._build_kpis(df_geral, df_31935, df_6292, df_divergencias)

        return {
            "df_geral": df_geral,
            "df_divergencias": df_divergencias,
            "kpis": kpis,
        }

    def _build_kpis(
        self,
        df_geral: pd.DataFrame,
        df_31935: pd.DataFrame,
        df_6292: pd.DataFrame,
        df_divergencias: pd.DataFrame,
    ) -> Dict[str, Any]:
        """Calcula os indicadores financeiros e operacionais para a aba Resumo da Auditoria."""
        total_fopa = len(df_geral)
        conciliados_fopa = int((df_geral["Status Auditoria"] == "Conciliado (Ok)").sum())
        divergentes_fopa = total_fopa - conciliados_fopa

        soma_pgtos = float(df_geral["Soma Eventos Pagamento"].sum())
        soma_5750 = float(df_geral["5750 - Adiantamento VT Atraves Alcadas"].sum())
        diff_5750 = round(soma_pgtos - soma_5750, 2)

        # Totais por evento principal
        tot_31935_fopa = float(
            df_geral["31935 - Valor Vale Transporte Alternativo Mes Seguinte"].sum()
        )
        tot_31935_imp = float(df_31935["Valor Importado 31935"].sum()) if not df_31935.empty else 0.0
        ok_31935 = int((df_31935["Obs 31935"] == "Ok").sum()) if not df_31935.empty else 0
        div_31935 = len(df_31935) - ok_31935

        tot_6292_fopa = float(df_geral["6292 - Desconto Cartao VTE"].sum())
        tot_6292_imp = float(df_6292["Valor Importado 6292"].sum()) if not df_6292.empty else 0.0
        ok_6292 = int((df_6292["Obs 6292"] == "Ok").sum()) if not df_6292.empty else 0
        div_6292 = len(df_6292) - ok_6292

        # Resumo por evento componente do 5750
        resumo_eventos_rows: List[Dict[str, Any]] = []
        for col_ev in EVENTOS_SOMA_5750:
            qtd_colab = int((df_geral[col_ev] > 0).sum())
            val_tot = round(float(df_geral[col_ev].sum()), 2)
            resumo_eventos_rows.append(
                {
                    "Evento FOPA": col_ev,
                    "Tipo": "Base de Entrada (Crédito/Reembolso VT)",
                    "Qtd Colaboradores (> R$ 0)": qtd_colab,
                    "Valor Total FOPA (R$)": val_tot,
                    "% sobre Soma Pagamentos": round(
                        (val_tot / soma_pgtos * 100.0) if soma_pgtos else 0.0, 2
                    ),
                }
            )
        resumo_eventos_rows.append(
            {
                "Evento FOPA": "5750 - Adiantamento VT Atraves Alcadas",
                "Tipo": "Contrapartida Saída Alçadas (Deve Igualar Soma)",
                "Qtd Colaboradores (> R$ 0)": int(
                    (df_geral["5750 - Adiantamento VT Atraves Alcadas"] > 0).sum()
                ),
                "Valor Total FOPA (R$)": round(soma_5750, 2),
                "% sobre Soma Pagamentos": round(
                    (soma_5750 / soma_pgtos * 100.0) if soma_pgtos else 0.0, 2
                ),
            }
        )

        # Resumo por Tipo Situação
        resumo_situacao = (
            df_geral.groupby(["Tipo Situação", "Situação"], dropna=False)
            .agg(
                Colaboradores=("Matrícula", "count"),
                Soma_Pagamentos=("Soma Eventos Pagamento", "sum"),
                Total_5750=("5750 - Adiantamento VT Atraves Alcadas", "sum"),
                Total_31935=("31935 - Valor Vale Transporte Alternativo Mes Seguinte", "sum"),
                Total_6292=("6292 - Desconto Cartao VTE", "sum"),
            )
            .reset_index()
            .sort_values("Colaboradores", ascending=False)
        )

        return {
            "total_fopa": total_fopa,
            "conciliados_fopa": conciliados_fopa,
            "divergentes_fopa": divergentes_fopa,
            "taxa_conformidade_fopa": round(
                (conciliados_fopa / total_fopa * 100.0) if total_fopa else 0.0, 2
            ),
            "soma_pgtos": round(soma_pgtos, 2),
            "soma_5750": round(soma_5750, 2),
            "diff_5750": diff_5750,
            "qtd_31935_imp": len(df_31935),
            "ok_31935": ok_31935,
            "div_31935": div_31935,
            "tot_31935_imp": round(tot_31935_imp, 2),
            "tot_31935_fopa": round(tot_31935_fopa, 2),
            "qtd_6292_imp": len(df_6292),
            "ok_6292": ok_6292,
            "div_6292": div_6292,
            "tot_6292_imp": round(tot_6292_imp, 2),
            "tot_6292_fopa": round(tot_6292_fopa, 2),
            "total_divergencias_consolidadas": len(df_divergencias),
            "df_resumo_eventos": pd.DataFrame(resumo_eventos_rows),
            "df_resumo_situacao": resumo_situacao,
        }
