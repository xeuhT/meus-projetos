# 💰 Módulo 07: Auditoria de Folha de Pagamento & Encargos

![Status](https://img.shields.io/badge/Status-Concluído-brightgreen)
![Fiscal](https://img.shields.io/badge/Auditoria-Malha%20Fina%20de%20Folha-darkgreen)

## 📌 Objetivo do Módulo
Pipeline analítico de malha fina para o fechamento mensal da folha de pagamento corporativa, auditando eventos de proventos e descontos, validação de encargos trabalhistas/previdenciários (INSS patronal, RAT, FAP, Terceiros, FGTS) e controle de gratificações manuais e SLAs de CSC.

## ⚙️ Regras de Negócio e Validações
* **Malha Fina de Eventos:** Confronto entre movimentações contratuais (admissões, promoções, méritos salariais) e os contracheques emitidos.
* **Auditoria de Gratificações Manuais:** Rastreamento e justificativa de valores atípicos inseridos manualmente fora de esteiras automáticas de comissão ou política salarial.
* **Governança de SLAs de CSC:** Acompanhamento de tempos de resposta e conclusão de chamados de atendimento de RH e DP.
* **Consistência de Encargos Previdenciários:** Verificação de bases de cálculo de INSS e FGTS com identificação de alíquotas divergentes.

## 📂 Arquivos e Scripts do Módulo
* `gerar_relatorio_malha_fina_folha.py`: Motor de auditoria analítica de rubricas e eventos.
* `gerar_relatorio_gratificacao_manual.py`: Identificação de pagamentos manuais atípicos.
* `gerar_relatorio_sla_csc.py`: Indicadores operacionais e cumprimento de prazos do CSC.

## 🚀 Como Executar
```bash
python gerar_relatorio_malha_fina_folha.py
```
