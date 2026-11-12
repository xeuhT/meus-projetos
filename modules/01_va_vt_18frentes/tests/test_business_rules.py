"""
Testes Unitários do Motor de Regras de Negócio (business_rules.py)
Auditoria "04 - GOOD CARD" - Cencosud Brasil

Valida de forma exaustiva:
1. merge_data: Cruzamentos relacionais (left joins por Matrícula) sem perda de linhas.
2. apply_financial_calculations: Fórmulas de saldo, margem 20%, status e equações de quitação.
3. apply_business_rules:
   - Regra 1: Desligados (quitação TRCT, saldo pendente, transferência sem ônus).
   - Regra 2: Afastados > 15 dias (bloqueio, provisionamento 29621, desconto indevido, suspensão).
   - Regra 3: Inelegibilidade (Aprendiz e Diretor com benefício indevido, exclusão de falsos positivos).
   - Regra 4: Margem 20%, quitação, estouro de margem, alerta de quality, subsídio 5% Estab 400/Giga.
4. build_conferencia_df: Formatação e conformidade com as 38 colunas oficiais.
5. reconcile_valores_importados: Conciliação de fornecedor, evento 29613 e 11 colunas oficiais.
6. compute_kpis: Apuração de totalizadores, conformidade e divisão por zero.
7. execute_audit: Execução ponta a ponta com dados sintéticos e dados reais.
"""

import os
import unittest
from pathlib import Path
import numpy as np
import pandas as pd

import config
import business_rules
from data_loader import load_all_data


class TestMergeData(unittest.TestCase):
    """Testes para a função merge_data."""

    def setUp(self):
        self.df_fopa = pd.DataFrame({
            "Matrícula": ["100", "200", "300"],
            "Nome do Colaborador": ["Ana", "Bruno", "Carlos"],
            "Cargo": ["Operador", "Caixa", "Repositor"],
            "11010 - Salario Mensal": [2000.0, 2500.0, 1800.0],
        })
        self.df_nomina = pd.DataFrame({
            "Matrícula": ["100", "200"],
            "Situação": ["Atividade Normal", "Afastado INSS"],
        })
        self.df_afast = pd.DataFrame({
            "Matrícula": ["200"],
            "Início Afastamento": ["01/09/2026"],
            "Término Afastamento": ["30/09/2026"],
            "Tipo Afastamento": ["5- Motivo de Doença, per. superior a 15 dias"],
            "Dias Afastamento": [30],
        })
        self.df_demit = pd.DataFrame({
            "Matrícula": ["300"],
            "Data Rescisão": ["15/09/2026"],
            "Tipo de Rescisão": ["2 - Sem justa causa do empregador"],
        })
        self.df_ferias = pd.DataFrame({
            "Matrícula": ["100"],
            "Início do Período de Gozo": ["01/09/2026"],
            "Fim do Período de Gozo": ["30/09/2026"],
            "Dias Gozo": [30],
        })

    def test_merge_preserves_row_count(self):
        """O número de linhas resultante deve ser estritamente igual ao do df_fopa."""
        merged = business_rules.merge_data(
            self.df_fopa, self.df_nomina, self.df_afast, self.df_demit, self.df_ferias
        )
        self.assertEqual(len(merged), 3)

    def test_merge_attaches_nomina_situacao(self):
        """Anexa situação cadastral do df_nomina."""
        merged = business_rules.merge_data(self.df_fopa, self.df_nomina)
        self.assertEqual(merged.loc[merged["Matrícula"] == "100", "Situação"].iloc[0], "Atividade Normal")
        self.assertEqual(merged.loc[merged["Matrícula"] == "200", "Situação"].iloc[0], "Afastado INSS")
        # Matrícula 300 não estava no nomina -> fallback para default
        self.assertEqual(merged.loc[merged["Matrícula"] == "300", "Situação"].iloc[0], "Atividade Normal")

    def test_merge_attaches_afastados_dates(self):
        """Anexa dados de afastamento renomeando colunas para o padrão oficial."""
        merged = business_rules.merge_data(self.df_fopa, df_afast=self.df_afast)
        row_200 = merged[merged["Matrícula"] == "200"].iloc[0]
        self.assertEqual(row_200["Início Afastamento"], "01/09/2026")
        self.assertEqual(row_200["Final Afastamento"], "30/09/2026")
        self.assertIn("superior a 15 dias", row_200["Tipo do Afastamento"])

    def test_merge_attaches_demitidos(self):
        """Anexa dados de rescisão mapeando para Desligamento."""
        merged = business_rules.merge_data(self.df_fopa, df_demit=self.df_demit)
        row_300 = merged[merged["Matrícula"] == "300"].iloc[0]
        self.assertEqual(row_300["Desligamento"], "15/09/2026")
        self.assertEqual(row_300["Tipo de Rescisão"], "2 - Sem justa causa do empregador")

    def test_merge_attaches_ferias(self):
        """Anexa dados de férias mapeando para Início Férias, Final Férias, Qtd de dias."""
        merged = business_rules.merge_data(self.df_fopa, df_ferias=self.df_ferias)
        row_100 = merged[merged["Matrícula"] == "100"].iloc[0]
        self.assertEqual(row_100["Início Férias"], "01/09/2026")
        self.assertEqual(row_100["Final Férias"], "30/09/2026")
        self.assertEqual(row_100["Qtd de dias"], 30)

    def test_merge_empty_optional_dataframes(self):
        """Lida graciosamente com dataframes opcionais None ou vazios."""
        merged = business_rules.merge_data(self.df_fopa)
        self.assertEqual(len(merged), 3)
        self.assertIn("Início Afastamento", merged.columns)
        self.assertIn("Desligamento", merged.columns)
        self.assertIn("Início Férias", merged.columns)


class TestFinancialCalculations(unittest.TestCase):
    """Testes para apply_financial_calculations."""

    def setUp(self):
        self.df = pd.DataFrame({
            "Matrícula": ["1", "2", "3", "4"],
            "11010 - Salario Mensal": [2000.0, 5000.0, 1500.55, 0.0],
            "29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ": [100.0, 0.0, 50.0, 0.0],
            "29613 - Ben Esp - Valor Informado para Desconto": [200.0, 317.06, 250.0, 0.0],
            "6280 - Desc Cartão Benefícios Especiais": [300.0, 317.06, 300.0, 0.0],
            "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": [0.0, 0.0, 0.0, 0.0],
            "29624 - Ben Esp - Custo Empresa Beneficios Especiais": [0.0, 0.0, 0.0, 0.0],
        })

    def test_total_saldo_devedor_sum(self):
        """Total Saldo Devedor deve ser a soma exata de 29622 e 29613."""
        res = business_rules.apply_financial_calculations(self.df)
        self.assertEqual(res.loc[0, "Total Saldo Devedor"], 300.0)
        self.assertEqual(res.loc[1, "Total Saldo Devedor"], 317.06)
        self.assertEqual(res.loc[2, "Total Saldo Devedor"], 300.0)
        self.assertEqual(res.loc[3, "Total Saldo Devedor"], 0.0)

    def test_margem_20_percent(self):
        """Margem 20% do Salário deve ser 11010 * 0.20 arredondada."""
        res = business_rules.apply_financial_calculations(self.df)
        self.assertEqual(res.loc[0, "Margem 20% do Salário"], 400.0)
        self.assertEqual(res.loc[1, "Margem 20% do Salário"], 1000.0)
        self.assertEqual(res.loc[2, "Margem 20% do Salário"], 300.11)
        self.assertEqual(res.loc[3, "Margem 20% do Salário"], 0.0)

    def test_status_dentro_fora_da_margem(self):
        """Avalia se Total Saldo Devedor está Dentro ou Fora da Margem."""
        res = business_rules.apply_financial_calculations(self.df)
        col = 'Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"'
        self.assertEqual(res.loc[0, col], "Dentro da Margem")
        self.assertEqual(res.loc[1, col], "Dentro da Margem")
        self.assertEqual(res.loc[2, col], "Dentro da Margem")
        self.assertEqual(res.loc[3, col], "Dentro da Margem")

    def test_status_desconto_x_margem(self):
        """Avalia se Evento 6280 está Dentro ou Fora da Margem."""
        res = business_rules.apply_financial_calculations(self.df)
        self.assertEqual(res.loc[0, "Desconto x Margem de 20%"], "Dentro da Margem")
        self.assertEqual(res.loc[1, "Desconto x Margem de 20%"], "Dentro da Margem")

    def test_saldo_devedor_quitacao_boolean(self):
        """Verifica equação de quitação 6280 + 29624 == Total Saldo Devedor."""
        res = business_rules.apply_financial_calculations(self.df)
        self.assertTrue(res.loc[0, "Saldo Devedor x Evt 6280"])
        self.assertTrue(res.loc[1, "Saldo Devedor x Evt 6280"])

    def test_custo_empresa_5_percent_calculation(self):
        """5% Saldo Devedor = Total Saldo Devedor * 0.05 e Cálculo 5% x 29624 = 29624 - 5% Saldo."""
        res = business_rules.apply_financial_calculations(self.df)
        # Row 0: Total = 300.0 -> 5% = 15.0 -> Calculo = 0.0 - 15.0 = -15.0
        self.assertEqual(res.loc[0, "5% Saldo devedor"], 15.0)
        self.assertEqual(res.loc[0, "Cálculo 5% x 29624"], -15.0)


class TestBusinessRulesClassification(unittest.TestCase):
    """Testes detalhados para apply_business_rules e classify_collaborator_row."""

    def test_regra_1_desligado_liquidado(self):
        """Demitido que quitou todo o saldo no acerto rescisório."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=300.0,
            salario_11010=2000.0,
            evt_6280=300.0,
            evt_29621=0.0,
            evt_29624=0.0,
            is_demitido=True,
            tipo_rescisao="2 - Sem justa causa",
        )
        self.assertEqual(obs, "Ok - Liquidou todo o saldo devedor ")
        self.assertIsNone(trat)

    def test_regra_1_desligado_saldo_pendente(self):
        """Demitido com saldo pendente não liquidado em TRCT -> Divergência."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=500.0,
            salario_11010=0.0,
            evt_6280=0.0,
            evt_29621=500.0,
            evt_29624=0.0,
            is_demitido=True,
            tipo_rescisao="2 - Sem justa causa",
        )
        self.assertEqual(obs, "Demitido - Saldo devedor pendente")
        self.assertIsNotNone(trat)
        self.assertIn("TRCT", trat)

    def test_regra_1_desligado_transferencia_sem_onus(self):
        """Demitido com tipo Transferência sem ônus é tratado à parte."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=100.0,
            salario_11010=2000.0,
            evt_6280=100.0,
            evt_29621=0.0,
            evt_29624=0.0,
            is_demitido=True,
            tipo_rescisao="5 - Transferência sem ônus",
        )
        self.assertEqual(obs, "Transferência sem ônus")
        self.assertIn("Transferência de empresa", trat)

    def test_regra_2_afastado_superior_15_dias_ok(self):
        """Afastado > 15 dias sem desconto e saldo levado para o evento 29621."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=250.0,
            salario_11010=0.0,
            evt_6280=0.0,
            evt_29621=250.0,
            evt_29624=0.0,
            has_afastamento=True,
            dias_afastamento=30,
        )
        self.assertEqual(
            obs,
            "Afastamento - Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ",
        )
        self.assertIsNone(trat)

    def test_regra_2_afastado_desconto_indevido(self):
        """Afastado com desconto em folha e salário zerado."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=200.0,
            salario_11010=0.0,
            evt_6280=200.0,
            evt_29621=0.0,
            evt_29624=0.0,
            has_afastamento=True,
            dias_afastamento=20,
        )
        self.assertEqual(obs, "Afastado - Desconto indevido em folha")
        self.assertIn("estorno", trat.lower())

    def test_regra_2_afastado_saldo_nao_provisionado(self):
        """Afastado com saldo devedor não levado para o 29621 na folha."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=150.0,
            salario_11010=0.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            has_afastamento=True,
            dias_afastamento=25,
        )
        self.assertEqual(obs, "Afastado - Saldo devedor não provisionado")
        self.assertIn("provisionamento", trat.lower())

    def test_regra_2_suspensao_de_contrato(self):
        """Colaborador em suspensão de contrato sem movimentação."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=0.0,
            salario_11010=0.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            situacao="Suspensão de Contrato",
        )
        self.assertEqual(obs, "Ok - Suspensão de Contrato")
        self.assertIsNone(trat)

    def test_regra_2_ferias_no_mes(self):
        """Colaborador em férias no mês com saldo levado para o 29621."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=300.0,
            salario_11010=2000.0,
            evt_6280=0.0,
            evt_29621=300.0,
            evt_29624=0.0,
            is_ferias=True,
        )
        self.assertEqual(
            obs,
            "Férias- Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ",
        )
        self.assertIsNone(trat)

    def test_regra_3_aprendiz_com_saldo_devedor(self):
        """Aprendiz com benefício Good Card indevido."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=100.0,
            salario_11010=1000.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="APRENDIZ REPOSITOR",
        )
        self.assertEqual(obs, "Aprendiz - Verificar")
        self.assertIn("cancelamento", trat.lower())

    def test_regra_3_aprendiz_sem_movimentacao(self):
        """Aprendiz sem saldo e sem desconto não deve ser apontado."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=0.0,
            salario_11010=1000.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="APRENDIZ REPOSITOR",
        )
        self.assertEqual(obs, "Ok - Sem saldo devedor e sem desconto")
        self.assertIsNone(trat)

    def test_regra_3_diretor_com_movimentacao(self):
        """Diretor com movimentação indevida de Good Card."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=500.0,
            salario_11010=25000.0,
            evt_6280=500.0,
            evt_29621=0.0,
            evt_29624=0.0,
            cargo="DIRETOR REGIONAL DE OPERAÇÕES",
        )
        self.assertEqual(obs, "Diretor - Verificar")
        self.assertIn("inelegível", trat.lower())

    def test_regra_4_liquidou_dentro_da_margem(self):
        """Colaborador normal liquidando todo o saldo dentro da margem de 20%."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=200.0,
            salario_11010=2000.0,
            evt_6280=200.0,
            evt_29621=0.0,
            evt_29624=0.0,
        )
        self.assertEqual(obs, "Ok - Liquidou todo o saldo devedor ")
        self.assertIsNone(trat)

    def test_regra_4_desconto_fora_da_margem_20_percent(self):
        """Desconto superior à margem permitida de 20%."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=500.0,
            salario_11010=2000.0,  # margem = 400.0
            evt_6280=450.0,        # > 400.0
            evt_29621=50.0,
            evt_29624=0.0,
        )
        self.assertEqual(obs, "Desconto fora da margem dos 20%")
        self.assertIn("Ajustar valor descontado", trat)

    def test_regra_4_descontou_limite_levou_restante_29621(self):
        """Saldo estoura margem, mas desconto respeitou margem e residual foi para 29621."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=600.0,
            salario_11010=2000.0,  # margem = 400.0
            evt_6280=400.0,        # dentro da margem
            evt_29621=200.0,       # residual exato
            evt_29624=0.0,
        )
        self.assertEqual(
            obs,
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621",
        )
        self.assertIsNone(trat)

    def test_regra_4_estab400_giga_5_percent_conforme(self):
        """Estabelecimento 400 ou Giga com 5% custo empresa e 95% desconto colaborador."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=200.0,
            salario_11010=3000.0,
            evt_6280=190.0,
            evt_29621=0.0,
            evt_29624=10.0,
            is_estab400_or_giga=True,
        )
        self.assertEqual(
            obs,
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)",
        )
        self.assertIsNone(trat)

    def test_regra_4_estab400_giga_divergencia_5_percent(self):
        """Estabelecimento 400 ou Giga com erro no subsídio de 5% da empresa."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=200.0,
            salario_11010=3000.0,
            evt_6280=180.0,
            evt_29621=0.0,
            evt_29624=20.0,  # Esperado 10.0
            is_estab400_or_giga=True,
        )
        self.assertEqual(obs, "Divergência Custo Empresa 5%")
        self.assertIn("diverge dos 5% regulamentares", trat)

    def test_regra_4_estab400_giga_com_saldo_residual_29621(self):
        """Estabelecimento 400 ou Giga com saldo residual no evento 29621, desconto e subsídio 5%."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=200.0,
            salario_11010=1000.0,
            evt_6280=95.0,
            evt_29621=95.0,
            evt_29624=10.0,
            is_estab400_or_giga=True,
        )
        self.assertEqual(
            obs,
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)",
        )
        self.assertIsNone(trat)

    def test_regra_4_estab400_giga_saldo_residual_estouro_margem(self):
        """Estabelecimento 400/Giga onde saldo devedor estoura margem e remanescente vai para 29621."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=1000.0,
            salario_11010=2000.0,
            evt_6280=400.0,
            evt_29621=550.0,
            evt_29624=50.0,
            is_estab400_or_giga=True,
        )
        self.assertEqual(
            obs,
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)",
        )
        self.assertIsNone(trat)

    def test_regra_4_estab400_giga_quitacao_integral_colaborador(self):
        """Estabelecimento 400 ou Giga onde colaborador quitou 100% do saldo (sem custo empresa)."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=300.0,
            salario_11010=2500.0,
            evt_6280=300.0,
            evt_29621=0.0,
            evt_29624=0.0,
            is_estab400_or_giga=True,
        )
        self.assertEqual(obs, "Ok - Liquidou todo o saldo devedor ")
        self.assertIsNone(trat)

    def test_regra_4_estab400_giga_casos_reais_liquidacao_100_porcento(self):
        """Valida casos reais da base (Matrículas 324777 e 327053) no Estabelecimento 917 Giga."""
        # Matrícula 324777: José Flor de Maio
        obs_1, trat_1 = business_rules.classify_collaborator_row(
            total_saldo_devedor=419.29,
            salario_11010=2312.45,
            evt_6280=419.29,
            evt_29621=0.0,
            evt_29624=0.0,
            is_estab400_or_giga=True,
        )
        self.assertEqual(obs_1, "Ok - Liquidou todo o saldo devedor ")
        self.assertIsNone(trat_1)

        # Matrícula 327053: Leandro Araujo Leite da Silva
        obs_2, trat_2 = business_rules.classify_collaborator_row(
            total_saldo_devedor=1795.61,
            salario_11010=12870.00,
            evt_6280=1795.61,
            evt_29621=0.0,
            evt_29624=0.0,
            is_estab400_or_giga=True,
        )
        self.assertEqual(obs_2, "Ok - Liquidou todo o saldo devedor ")
        self.assertIsNone(trat_2)

    def test_regra_4_estab400_giga_casos_reais_saldo_residual_fora_margem(self):
        """Valida casos reais de residual fora da margem no Estabelecimento 950 (Matrículas 320923 e 321226)."""
        # Matrícula 320923: Jailson Francisco de Oliveira (GIGA 950)
        obs_1, trat_1 = business_rules.classify_collaborator_row(
            total_saldo_devedor=3569.08,
            salario_11010=3981.38,
            evt_6280=756.47,
            evt_29621=2772.80,
            evt_29624=39.81,
            is_estab400_or_giga=True,
        )
        self.assertEqual(
            obs_1,
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621",
        )
        self.assertIsNone(trat_1)

        # Matrícula 321226: Danilo da Silva Dias (GIGA 950)
        obs_2, trat_2 = business_rules.classify_collaborator_row(
            total_saldo_devedor=716.14,
            salario_11010=2532.40,
            evt_6280=481.16,
            evt_29621=209.66,
            evt_29624=25.32,
            is_estab400_or_giga=True,
        )
        self.assertEqual(
            obs_2,
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621",
        )
        self.assertIsNone(trat_2)

    def test_regra_4_estab400_giga_divergencia_custo_empresa_superior(self):
        """Estabelecimento 400 ou Giga com custo empresa superior aos 5% regulamentares."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=200.0,
            salario_11010=3000.0,
            evt_6280=180.0,
            evt_29621=0.0,
            evt_29624=20.0,  # Esperado 10.0 (5%), diferença de 10.0 > 0.05
            is_estab400_or_giga=True,
        )
        self.assertEqual(obs, "Divergência Custo Empresa 5%")
        self.assertIn("diverge dos 5% regulamentares", trat)

    def test_regra_4_estab400_giga_divergencia_custo_empresa_inferior(self):
        """Estabelecimento 400 ou Giga com custo empresa inferior aos 5% regulamentares (> 5 centavos)."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=200.0,
            salario_11010=3000.0,
            evt_6280=190.10,
            evt_29621=0.0,
            evt_29624=9.90,  # Esperado 10.0, diferença de 0.10 > 0.05 de tolerância
            is_estab400_or_giga=True,
        )
        self.assertEqual(obs, "Divergência Custo Empresa 5%")
        self.assertIn("diverge dos 5% regulamentares", trat)

    def test_regra_4_estab400_giga_5_percent_com_saldo_residual_equacao_quebrada(self):
        """Estabelecimento 400 ou Giga com 5% custo empresa e equação contábil violada -> Divergência no saldo devedor."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=1000.0,
            salario_11010=1000.0,
            evt_6280=200.0,
            evt_29621=600.0,  # Faltam R$ 150,00 para fechar 1000.0
            evt_29624=50.0,
            is_estab400_or_giga=True,
        )
        self.assertEqual(obs, "Divergência no saldo devedor")
        self.assertIn("Soma do desconto (6280)", trat)

    def test_regra_4_quality_ativo_sem_afastamento_sem_desconto(self):
        """Colaborador ativo com saldo e SEM desconto em folha -> Quality alert."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=350.0,
            salario_11010=2500.0,
            evt_6280=0.0,
            evt_29621=0.0,
            evt_29624=0.0,
            has_afastamento=False,
            is_ferias=False,
        )
        self.assertEqual(
            obs,
            "Verificar com quality - Não estava em situação de afastamento e não houve o desconto",
        )
        self.assertIn("Quality/Folha", trat)


class TestBuildConferenciaDf(unittest.TestCase):
    """Testes para build_conferencia_df."""

    def test_build_conferencia_exact_38_columns(self):
        """Verifica se o DataFrame final possui rigorosamente as 38 colunas oficiais."""
        df_dummy = pd.DataFrame({
            "Código da Empresa": [2],
            "Matrícula": ["935"],
            "Total Saldo Devedor": [317.06],
            "Obs. Conferência": ["Ok - Liquidou todo o saldo devedor "],
            "Tratativa": [None],
        })
        df_conf = business_rules.build_conferencia_df(df_dummy)
        self.assertEqual(len(df_conf.columns), 38)
        self.assertEqual(list(df_conf.columns), config.COLUNAS_OFICIAIS_CONFERENCIA)
        self.assertEqual(df_conf.loc[0, "Total Saldo Devedor"], 317.06)
        self.assertIsNone(df_conf.loc[0, "Tratativa"])


class TestReconcileValoresImportados(unittest.TestCase):
    """Testes para reconcile_valores_importados."""

    def setUp(self):
        self.df_vi = pd.DataFrame({
            "Emp": [7, 7, 7],
            "Matrícula": ["1001", "1002", "1003"],
            "Evento": [29613, 29613, 29613],
            "Mês": [92026, 92026, 92026],
            "Valor": [300.0, 150.0, 200.0],
        })
        self.df_fopa = pd.DataFrame({
            "Matrícula": ["1001", "1002"],
            "29613 - Ben Esp - Valor Informado para Desconto": [300.0, 100.0],
        })
        self.df_demit = pd.DataFrame({
            "Matrícula": ["1003"],
            "Data Rescisão": ["10/09/2026"],
            "Tipo de Rescisão": ["2 - Sem justa causa"],
        })

    def test_reconcile_exact_11_columns(self):
        """Gera exatamente as 11 colunas oficiais de Valores Importados."""
        res = business_rules.reconcile_valores_importados(self.df_vi, self.df_fopa, self.df_demit)
        self.assertEqual(len(res.columns), 11)
        self.assertEqual(list(res.columns), config.COLUNAS_OFICIAIS_VALORES_IMPORTADOS)

    def test_reconcile_importado_x_folha_logic(self):
        """Avalia True quando Valor == Folha e False caso contrário."""
        res = business_rules.reconcile_valores_importados(self.df_vi, self.df_fopa, self.df_demit)
        # 1001: 300 == 300 -> True
        self.assertTrue(res.loc[res["Matrícula"] == "1001", "Importado x Folha"].iloc[0])
        # 1002: 150 != 100 -> False
        self.assertFalse(res.loc[res["Matrícula"] == "1002", "Importado x Folha"].iloc[0])
        # 1003: 200 != 0.0 (não encontrado na folha) -> False
        self.assertFalse(res.loc[res["Matrícula"] == "1003", "Importado x Folha"].iloc[0])

    def test_reconcile_desligado_attached(self):
        """Anexa data de rescisão para colaboradores demitidos."""
        res = business_rules.reconcile_valores_importados(self.df_vi, self.df_fopa, self.df_demit)
        row_1003 = res[res["Matrícula"] == "1003"].iloc[0]
        self.assertEqual(row_1003["Desligado"], "10/09/2026")
        self.assertEqual(row_1003["Tipo de Rescisão"], "2 - Sem justa causa")
        self.assertIn("desligado", row_1003["Tratativa"].lower())


class TestComputeKpis(unittest.TestCase):
    """Testes para compute_kpis."""

    def test_compute_kpis_metrics(self):
        df_conf = pd.DataFrame({
            "Total Saldo Devedor": [100.0, 200.0, 0.0],
            "6280 - Desc Cartão Benefícios Especiais": [100.0, 190.0, 0.0],
            "29621 - Ben Esp - Saldo Devedor Beneficios Especiais": [0.0, 0.0, 0.0],
            "29624 - Ben Esp - Custo Empresa Beneficios Especiais": [0.0, 10.0, 0.0],
            "29613 - Ben Esp - Valor Informado para Desconto": [100.0, 200.0, 0.0],
            'Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"': [
                "Dentro da Margem",
                "Dentro da Margem",
                "Dentro da Margem",
            ],
            "Obs. Conferência": [
                "Ok - Liquidou todo o saldo devedor ",
                "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)",
                "Ok - Sem saldo devedor e sem desconto",
            ],
            "Tratativa": [None, None, None],
            "Desligamento": ["", "", ""],
            "Início Afastamento": ["", "", ""],
        })
        df_vi = pd.DataFrame({"Valor": [300.0]})

        kpis = business_rules.compute_kpis(df_conf, df_vi)
        self.assertEqual(kpis["total_auditados"], 3)
        self.assertEqual(kpis["com_saldo_devedor"], 2)
        self.assertEqual(kpis["conformes"], 3)
        self.assertEqual(kpis["divergencias"], 0)
        self.assertEqual(kpis["taxa_conformidade"], 100.0)
        self.assertEqual(kpis["tot_saldo_devedor"], 300.0)
        self.assertEqual(kpis["tot_desc_6280"], 290.0)
        self.assertEqual(kpis["tot_custo_29624"], 10.0)

    def test_compute_kpis_empty_dataframe(self):
        """Dataframe vazio não deve causar divisão por zero."""
        df_empty = pd.DataFrame()
        df_vi_empty = pd.DataFrame()
        kpis = business_rules.compute_kpis(df_empty, df_vi_empty)
        self.assertEqual(kpis["total_auditados"], 0)
        self.assertEqual(kpis["taxa_conformidade"], 0.0)


class TestExecuteAuditIntegration(unittest.TestCase):
    """Testes de integração para execute_audit."""

    def test_execute_audit_with_real_files(self):
        """Executa a auditoria completa contra o dataset real de RH."""
        data_dir = Path("Relatorio Original")
        template_file = Path("Conferência Good Card_Base.xlsx")
        if not (data_dir.exists() and template_file.exists()):
            self.skipTest("Arquivos reais não encontrados.")

        data_dict = load_all_data(data_dir=data_dir, template_path=template_file)
        df_conf, df_vi, kpis = business_rules.execute_audit(data_dict)

        # 1. Validações da Conferência
        self.assertEqual(len(df_conf), 15913)
        self.assertEqual(len(df_conf.columns), 38)
        self.assertEqual(list(df_conf.columns), config.COLUNAS_OFICIAIS_CONFERENCIA)

        # 2. Validações de Valores Importados
        self.assertEqual(len(df_vi), 1334)
        self.assertEqual(len(df_vi.columns), 11)
        self.assertEqual(list(df_vi.columns), config.COLUNAS_OFICIAIS_VALORES_IMPORTADOS)

        # 3. Validações de KPIs
        self.assertEqual(kpis["total_auditados"], 15913)
        self.assertEqual(kpis["com_saldo_devedor"], 1334)
        self.assertAlmostEqual(kpis["tot_saldo_devedor"], 522186.19, places=1)
        self.assertGreater(kpis["taxa_conformidade"], 99.0)


if __name__ == "__main__":
    unittest.main()
