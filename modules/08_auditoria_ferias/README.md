# 🏖️ Módulo 08: Auditoria Integral de Férias (Cálculo e Folha Mensal)

![Status](https://img.shields.io/badge/Status-Concluído-brightgreen)
![Versão](https://img.shields.io/badge/Motor-v008%20Consolidado-blue)

## 📌 Objetivo do Módulo
Motor financeiro avançado de conciliação dupla e conferência vetorial entre o cálculo original do recibo de férias (abono pecuniário, 1/3 constitucional, médias variáveis de horas extras/comissões, pensão alimentícia, retenção de INSS e IRRF) e o seu lançamento na folha de pagamento mensal oficial.

## ⚙️ Regras de Negócio e Validações
* **Confronto Vetorial Duplo:** Validação paralela centavo a centavo entre recibo pago e lançamento contábil em folha de fechamento.
* **Médias de Variáveis:** Verificação da correta incorporação de horas extras, adicionais noturnos, DSR e comissões do período aquisitivo.
* **Retenção Fiscal e Previdenciária:** Apuração da tabela progressiva de IRRF e alíquotas faixas de INSS nas férias gozadas dentro do mês e partidas entre competências distintas.
* **Geração de Evidências Executivas:** Exportação automatizada de relatórios em Excel com formatação condicional inteligente destacando divergências.

## 📂 Arquivos e Scripts do Módulo
* `atualizar_todas_auditorias.py`: Orquestrador consolidado v008 que executa a malha completa de auditorias de férias.
* `gerar_conferencia_ferias_formatado_v4.py`: Gerador de relatórios executivos formatados.
* `gerar_auditoria_calculo_v6.py`: Motor de recálculo independente de proventos de férias.
* `gerar_auditoria_folha_v6.py`: Confronto com o espelho oficial de folha mensal.
* `generate_doc_pdf.py`: Gerador de evidências e pareceres em PDF para auditoria externa.

## 🚀 Como Executar
```bash
python atualizar_todas_auditorias.py
```
