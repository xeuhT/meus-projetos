# 📊 Módulo 12: DHO & Recrutamento — Dashboard Power BI de Pré-Admissão

![Status](https://img.shields.io/badge/Status-Validação_Final_/_Homologação-yellow)
![Power BI](https://img.shields.io/badge/Power_BI-Analytics_em_Nuvem-yellow?logo=powerbi)
![Data Source](https://img.shields.io/badge/Fonte-SharePoint_Corporativo-blue)
![ROI](https://img.shields.io/badge/ROI-38.0h%2Fmês_salvas-brightgreen)

---

## 📌 Visão Geral & Escopo Estratégico

Painel analítico gerencial desenvolvido em **Microsoft Power BI** e publicado no serviço em nuvem (*Power BI Service*), com conexão direta e contínua à planilha de controle operacional hospedada no **SharePoint corporativo**.

O objetivo estratégico do painel é municiar **gerentes regionais, diretores de operações e a liderança de DHO/Recrutamento** com visibilidade executiva em tempo real sobre a esteira de atração e pré-admissão de novos talentos nas lojas, postos de combustíveis e centros de distribuição da Cencosud Brasil.

---

## 📈 Indicadores Chave & Métricas de Performance (KPIs)

* **Time-to-Hire (Tempo Médio de Contratação):** Medição precisa dos dias decorridos desde a abertura da vaga até a integração final do candidato.
* **Gargalos de Entrega de Documentos:** Identificação de candidatos travados em etapas de envio de documentação pessoal, certidões ou comprovantes de residência.
* **Status de Exames Admissionais (ASO):** Monitoramento de prazos de agendamento de exames clínicos/complementares e emissão do Atestado de Saúde Ocupacional junto às clínicas credenciadas.
* **Taxa de Conversão da Esteira:** Percentual de perda ou desistência de candidatos por fase da seleção e por unidade de negócio.

---

## ⚙️ Arquitetura Técnica

* `deashboard_rh.pbix`: Arquivo analítico do Power BI com modelagem dimensional (Star Schema).
* `queries_dax/`: Medidas DAX avançadas para cálculo de SLAs dinâmicos, comparativos de metas regionais e acumulados móveis.
* **Status Atual:** A integração técnica direta com o SharePoint corporativo está concluída e operacional; aguarda-se a padronização definitiva das colunas da planilha pela equipe de negócios para homologação e publicação final.
