# 📄 Módulo 09: Conciliação e Auditoria de Rescisão (TRCTs em PDF)

![Status](https://img.shields.io/badge/Status-Em%20Homologação%20%2F%20Desenvolvimento-yellow)
![Tecnologia](https://img.shields.io/badge/OCR%20%7C%20Parser-PDFPlumber%20%7C%20PyPDF-purple)

## 📌 Objetivo do Módulo
Pipeline de inteligência documental e OCR/Parser para extração automatizada de dados tabulares e metadados de Termos de Rescisão de Contrato de Trabalho (TRCTs em formato `.pdf`), com confronto automático rubrica a rubrica contra o mapa coletivo de rescisões e bases oficiais do DP.

## ⚙️ Regras de Negócio e Validações
* **Parser Inteligente de TRCTs:** Extração de campos estruturados (Dados do Empregador, Trabalhador, Causa do Afastamento, Datas de Admissão/Aviso/Afastamento).
* **Validação Rubrica a Rubrica:**
  - Saldo de salário (cód. 721);
  - Aviso prévio proporcional indenizado e trabalhado (Lei 12.506/2011, rubricas 725/20208);
  - 13º salário proporcional e indenizado;
  - Férias proporcionais e vencidas com 1/3 constitucional;
  - Descontos de faltas, DSR e adiantamentos.
* **Motor Financeiro Rescisório:** Recálculo das bases de apuração da multa rescisória de 40% do FGTS e deduções fiscais (IRRF/INSS).
* **Documentação Técnica:** Acompanha o [Manual de Auditoria de Rescisão](Manual_Auditoria_Rescisao.pdf).

## 📂 Arquivos e Scripts do Módulo
* `conciliador_rescisao.py`: Script de orquestração do confronto coletivo.
* `leitor_pdf_rescisao.py`: Motor de extração de tabelas e texto de PDFs de TRCT.
* `motor_financeiro.py`: Regras de cálculo de verbas rescisórias e descontos legais.
* `processador_rescisao.py`: Pipeline de transformação e enriquecimento de dados.
* `app_rescisao.py`: Interface para execução da auditoria rescisória.
* `Manual_Auditoria_Rescisao.pdf`: Guia e documentação operacional de conformidade.

## 🗺️ Roadmap / Próximos Passos
- [ ] Refinamento do parser para variações de layout gráfico emitidas por filiais de outros estados.
- [ ] Ajuste fino do recálculo de alíquotas da tabela progressiva de INSS em rescisões com múltiplos vínculos.
- [ ] Implementação de exportação direta de divergências para pareceres jurídicos automatizados.

## 🚀 Como Executar
```bash
python conciliador_rescisao.py
```
