# 💳 Módulo: Pagadoria & Fluxo de Caixa RH (SAP ZFI051)

![Status Carga SAP](https://img.shields.io/badge/Status_Carga_SAP-Em_Validação_Rateio-yellow)
![Status Fluxo de Caixa](https://img.shields.io/badge/Status_Fluxo_Caixa-Modelagem_de_Dados-blue)
![SAP](https://img.shields.io/badge/Integração-SAP_ZFI051-navy)
![ROI](https://img.shields.io/badge/ROI-30.7h%2Fmês_salvas-brightgreen)

## 📌 Visão Geral & Escopo

Este módulo gerencia a automação da rotina de pagamentos de Recursos Humanos e a integração financeira com a Tesouraria e a Controladoria da Cencosud Brasil:

1. **Carga SAP ZFI051 & Consolidador de Alçadas (17,5h salvas/mês):**
   * Consolidação das aprovações de pagamentos de RH (folha, benefícios, encargos, pensões alimentícias e rescisões);
   * Validação de alçadas de autorização financeira conforme governança interna corporativa;
   * Geração automática do arquivo estruturado (formato posicional/delimitado `.txt`) para a transação **SAP ZFI051**;
   * *Status:* Motor funcional entregue; pendente ajuste de regra na aba de rateio por centro de custo para empacotamento no executável `.exe`.

2. **Projeção de Fluxo de Caixa de RH (13,2h salvas/mês):**
   * Motor preditivo e analítico de projeção diária e semanal de desembolsos de RH para a Tesouraria Corporativa;
   * Mapeamento de janelas críticas: adiantamento salarial, fechamento mensal, GPS/DARF DCTFWeb, FGTS Digital e vencimentos de faturas de fornecedores;
   * *Status:* Em fase final de modelagem de dados e conciliação de calendários financeiros.

---

## 🛠️ Arquitetura dos Scripts

* `carga_sap_zfi051.py`: Motor de processamento das planilhas de aprovação de pagamentos, validação de alçadas e exportação do arquivo de carga SAP ZFI051.
* `projecao_fluxo_caixa.py`: Motor analítico que consolida os eventos futuros da folha e gera a matriz de liquidação financeira diária para a Tesouraria.
* `config_pagadoria.py`: Dicionário corporativo de empresas (CNPJs/Empresas SAP), centros de custo e regras de alçadas de aprovação.

---

## 🚀 Como Executar

```bash
# Executar a consolidação e geração da Carga SAP ZFI051
python carga_sap_zfi051.py --planilha "caminho/para/aprovacoes_pagamentos.xlsx" --saida "carga_zfi051.txt"

# Gerar a projeção consolidada de Fluxo de Caixa para a Tesouraria
python projecao_fluxo_caixa.py --competencia "2026-05" --saida "fluxo_caixa_rh_projetado.xlsx"
```
