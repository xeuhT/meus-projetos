# 🏛️ Enterprise Payroll Audit & HR Data Engineering Suite

Um ecossistema robusto e modular de engenharia de dados, auditoria contínua e garantia de conformidade trabalhista desenvolvido para processar, conciliar e auditar bases de dados em escala corporativa (milhares de colaboradores e múltiplos CNPJs).

O projeto centraliza **18 frentes operacionais de conferência e conciliação financeira**, eliminando riscos de pagamento indevido, reduzindo glosas contratuais e garantindo compliance fiscal e trabalhista com governança de ponta a ponta.

---

## 📊 Arquitetura e Escopo de Atuação

### 1. 🛡️ Auditoria de Benefícios, Cartões e Faturamento (18 Módulos de Fechamento)
Pipelines automatizados de conferência entre folha de pagamento (LG Lugar de Gente), faturamento analítico de operadoras e bases operacionais:
* **Vale Transporte & Desconto Legal:** Motor de apuração dos 6% constitucionais vs. custo real de compra e cálculo automático de rateio de centros de custo (CVE).
* **Gestão de Cartões de Alimentação/Refeição:** Auditoria de créditos Good Card (Evt. 6280), bloqueios preventivos de cartões para demitidos/afastados e recargas Cencopay.
* **Conciliação de Saúde e Odonto:** Detecção ativa de colaboradores desligados mantidos indevidamente em operadoras (SulAmérica Saúde, GNDI, Hapvida, Unimed, etc.), coparticipações e conferência de "vidas fantasmas".
* **Conferências Específicas:** Relatórios para contratos intermitentes, apuração de refeitório de lojas, unidades de entrega física de benefícios e campanhas de Cesta Prêmio / Ticket Bretas.

### 2. 📑 Auditoria Integral de Férias e Fechamento Mensal
* **Confronto Vetorial Duplo:** Validação paralela entre o cálculo do recibo de férias (abono, 1/3 constitucional, médias de variáveis, pensão alimentícia e retenção de INSS/IRRF) contra o fechamento oficial da folha mensal.
* **Geração de Evidências:** Exportação automática de relatórios em Excel com marcação visual condicional de divergências de centavos e inconsistências de competência.

### 3. 🔍 Conciliação Automatizada de Rescisões (TRCTs em PDF)
* **Pipeline OCR/Parser de Documentos:** Extração automatizada de dados tabulares e metadados de Termos de Rescisão de Contrato de Trabalho (`.pdf`).
* **Validação Rubrica a Rubrica:** Motor financeiro que audita saldos de salário (721), aviso prévio proporcional à lei (725/20208), férias vencidas/indenizadas e base de cálculo rescisória em lote contra o mapa coletivo.

### 4. 🤖 Compliance Cadastral & Detecção de Fraudes (Robô de Afastados)
* **Integração de APIs Públicas:** Integração via REST API (SintegraWS e InfoSimples) para consulta cadastral em lote de colaboradores afastados junto à Receita Federal.
* **Inteligência de Riscos:** Identificação de participação ativa em Quadro de Sócios e Administradores (QSA), incompatibilidade cadastral e auditoria de óbitos não comunicados.
* **Dossiês Executivos:** Geração automatizada de laudos investigativos em Word (`.docx`) e relatórios consolidados em Excel.

### 5. 🛠️ Ferramental de Produtividade e Apoio
* **UI Gráfica:** Interfaces em CustomTkinter para operação simplificada por equipes de RH não técnicas.
* **Governança & LGPD:** Scripts dedicados de anonimização e mascaramento de PII (CPFs, nomes, salários) para compartilhamento seguro de bases analíticas.
* **Rotinas Operacionais:** Automação de enquadramento de Salário Família e normalização de relatórios gerais de admissão e cargos.

---

## 💻 Tecnologias Empregadas
* **Linguagem:** Python 3.10+
* **Engenharia & Tratamento de Dados:** Pandas, NumPy, OpenPyXL, PyXLSB, XLRD
* **Extração de Documentos & OCR:** PDFPlumber, PyPDF, Python-Docx
* **Integrações & Web Services:** Requests, APIs REST, JSON Cache Engine
* **Interfaces Gráficas:** CustomTkinter, Tkinter, Pillow
