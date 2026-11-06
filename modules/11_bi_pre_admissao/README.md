# 📊 Módulo 11: Dashboard Power BI de Pré-Admissão & Esteira de Admissão

![Status](https://img.shields.io/badge/Status-Concluído-brightgreen)
![BI](https://img.shields.io/badge/Power%20BI-Template%20Executivo-yellow)

## 📌 Objetivo do Módulo
Central analítica para acompanhamento da esteira de contratação e pré-admissão corporativa de candidatos em grande escala para as lojas e centros de distribuição da Cencosud, monitorando tempos de entrega de documentação, validação cadastral e aprovação no eSocial.

## ⚙️ Indicadores e Funcionalidades
* **Gargalos na Esteira:** Rastreio dos estágios de documentação pendente, envio de exame admissional (ASO) e qualificação cadastral.
* **SLA de Contratação (Time to Hire):** Métricas de tempo decorrido entre a abertura da vaga, aprovação do candidato e início efetivo em loja.
* **Taxa de Conversão e Desistência:** Análise de desistências na fase pré-admissional por regional e cargo.
* **Automação de Tratamento de Dados:** Script Python dedicado para higienizar e consolidar a base de admissões antes da ingestão pelo modelo Power BI.

## 📂 Arquivos do Módulo
* `deashboard_rh.pbix`: Arquivo template do Power BI com modelo de dados, medidas DAX e visualizações.
* `Dashboard_Pre_Admissao_RH_Apresentacao.pptx`: Apresentação executiva do projeto e storytelling dos indicadores para a diretoria.
* `app_relatorio_admissao.py`: Script Python para extração, normalização e carga dos dados da esteira.

## 🚀 Como Utilizar
1. Execute `python app_relatorio_admissao.py` para gerar a base higienizada.
2. Abra o `deashboard_rh.pbix` no Power BI Desktop e atualize as fontes de dados para visualizar o dashboard interativo.
