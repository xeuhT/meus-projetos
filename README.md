# 🏛️ Enterprise Payroll Audit & HR Data Engineering Suite
### Ecossistema Integrado de Engenharia de Dados, Auditoria Contínua & Automação de Recursos Humanos

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)
![Power BI](https://img.shields.io/badge/Power_BI-Analytics-yellow?logo=powerbi)
![Selenium](https://img.shields.io/badge/Selenium-RPA-green?logo=selenium)
![CustomTkinter](https://img.shields.io/badge/CustomTkinter-Desktop_GUI-orange)
![Security & LGPD](https://img.shields.io/badge/Security-LGPD_Compliant-red)
![License](https://img.shields.io/badge/License-Proprietary%20%2F%20Enterprise-lightgrey)

---

## 📖 Visão Executiva

Este repositório consolida o **ecossistema corporativo completo de engenharia de dados, automações de processos (RPA), compliance trabalhista e auditoria contínua de folha de pagamento** desenvolvido para operar em escala corporativa (milhares de colaboradores, múltiplos CNPJs e unidades espalhadas em nível nacional).

O objetivo principal da suíte é **eliminar pagamentos indevidos, erradicar glosas contratuais de operadoras de benefícios, blindar passivos trabalhistas (CLT/eSocial/SST) e transformar processos manuais de Departamento Pessoal em pipelines analíticos de alta performance**.

---

## 🗺️ Mapa dos 13 Módulos Integrados

| # | Módulo | Foco de Negócio & Escopo | Status |
|:---:|:---|:---|:---:|
| **01** | [🍱 Benefícios Corporativos & VA/VT](modules/01_beneficios_va_vt/) | 18 conferências oficiais (VT 6%, Good Card, CVE, Cencopay, Refeitório) | `Concluído` |
| **02** | [🛡️ Auditoria de Seguro de Vida](modules/02_seguro_de_vida/) | Conciliação de apólice básica e capitais ampliados (12x e 24x salariais) | `Concluído` |
| **03** | [💳 Extrato de Ticket Plus / Compras](modules/03_extrato_ticket_plus/) | Robô de extração e consolidação analítica de extratos de convênio | `Concluído` |
| **04** | [🌐 Automação do Conectividade Social](modules/04_conectividade_social/) | RPA Selenium + Certificado Digital A1 no Windows para saldo FGTS | `Concluído` |
| **05** | [⏱️ Auditoria de Ponto e Frequência](modules/05_ponto_e_frequencia/) | Jornadas críticas (Art. 59 CLT), sobreaviso, atestados e controle de DSR | `Concluído` |
| **06** | [🤝 Relações Sindicais & CCTs](modules/06_relacoes_sindicais/) | Enquadramento por unidade, adicional por feriados e regras sindicais | `Concluído` |
| **07** | [💰 Folha de Pagamento & Encargos](modules/07_folha_e_encargos/) | Malha fina de eventos, gratificações manuais atípicas e SLAs do CSC | `Concluído` |
| **08** | [🏖️ Auditoria Integral de Férias](modules/08_auditoria_ferias/) | Confronto vetorial duplo (recibo de férias vs folha mensal oficial v008) | `Concluído` |
| **09** | [📄 Conciliação de Rescisões](modules/09_conciliacao_rescisao/) | Pipeline OCR/Parser de TRCTs em PDF e batimento financeiro de verbas | `Homologação` |
| **10** | [🤖 Compliance Cadastral & Afastados](modules/10_compliance_afastados/) | Robô SintegraWS para detecção de fraudes, sócios em QSA, INSS e óbitos | `Concluído` |
| **11** | [📊 BI de Pré-Admissão & Admissão](modules/11_bi_pre_admissao/) | Dashboard executivo Power BI (.pbix) e scripts de esteira de candidatos | `Concluído` |
| **12** | [🖥️ App Desktop de Benefícios](modules/12_app_desktop_beneficios/) | Interface gráfica CustomTkinter para auditoria de saúde e odontologia | `Concluído` |
| **--** | [🛠️ Utilitários & Governança LGPD](utils/) | Anonimizador de dados, Salário Família e normalização de cargos | `Concluído` |

---

## 📈 Métricas de Impacto e Ganho de Tempo (ROI do Ecossistema)

A implementação do ecossistema substituiu processos manuais altamente suscetíveis a falhas e gargalos de fechamento por pipelines automatizados em Python com validação matemática estrita, gerando uma redução drástica no tempo operacional:

| # | Projeto / Frente Operacional | Tempo Anterior (Manual) | Tempo Atual (Automação Python) | Redução de Tempo (%) | Horas Economizadas / Mês |
| :-: | :--- | :---: | :---: | :---: | :---: |
| **01** | Conferência Geral de Folha & Encargos (Malha Fina) | 7 dias úteis (~56h) | 1,5 a 2 dias (~14h) | **75%** | **42h / ciclo** |
| **02** | Auditoria de VT 6% (Teto Salarial vs Custo Compra) | 3 dias úteis (~24h) | 20 minutos | **98,3%** | **23,6h / mês** |
| **03** | Fechamento de Benefícios & Cartões (18 Módulos de VA/VT) | 10 dias úteis (~80h) | 1 dia útil (~8h) | **90%** | **72h / mês** |
| **04** | Auditoria de Férias (Cálculo vs Folha Mensal vs Médias) | 4 dias úteis (~32h) | 40 minutos | **97,9%** | **31,3h / mês** |
| **05** | Conciliação de Rescisões (Parsing de TRCT em PDF) | 40 min por rescisão (~40h/mês) | 30 seg por rescisão (~1h/mês) | **97,5%** | **39h / mês** |
| **06** | Auditoria de Faturas de Saúde/Odonto ("Vidas Fantasmas") | 5 dias úteis (~40h) | 1,5 hora | **96,2%** | **38,5h / mês** |
| **07** | Robô de Afastados, QSA e Óbitos (SintegraWS / RFB) | 6 dias de consulta (~48h) | 2 horas em lote | **95,8%** | **46h / lote** |
| **08** | Automação Web Conectividade Social (Extrato FGTS) | 15 min por guia/colab (~25h) | 1 min por consulta (~2h) | **92%** | **23h / mês** |
| **09** | Auditoria de Seguro de Vida (Básico + 12x/24x) | 2 dias úteis (~16h) | 10 minutos | **98,9%** | **15,8h / mês** |
| **10** | Auditoria de Ponto, Jornadas Críticas e Justificativas | 3 dias úteis (~24h) | 35 minutos | **97,5%** | **23,4h / mês** |
| **11** | Robô de Extratos Ticket Plus / Compras | 12 horas manuais | 20 minutos | **97,2%** | **11,6h / mês** |
| **12** | CCTs, Relações Sindicais & Perfil de Frequência | 16 horas manuais | 15 minutos | **98,4%** | **15,7h / mês** |
| **13** | Esteira Power BI & Normalização de Admissões | 10 horas semanais (~40h/mês) | Atualização contínua (~2h/mês) | **95%** | **38h / mês** |
| **TOTAL** | **IMPACTO CONSOLIDADO DO ECOSSISTEMA** | **~445 horas / mês** | **~48 horas / mês** | **89,2%** | **397 horas poupadas/mês (~2,5 FTEs)** |

---

## 🏗️ Arquitetura da Solução

```mermaid
flowchart TD
    subgraph INGESTION["📥 Camada de Ingestão de Dados"]
        LG["Folha de Pagamento LG"]
        OPERADORAS["Faturas Operadoras (Planilhas / XLSB)"]
        PDFS["TRCTs e Documentos em PDF"]
        WEBSERVICE["APIs Governamentais (Receita / Sintegra)"]
        PORTAL_CAIXA["Conectividade Social V2 (Caixa)"]
    end

    subgraph ENGINE["⚙️ Motores de Processamento & Auditoria"]
        M01["Auditor de Benefícios (18 Frentes)"]
        M04["RPA Selenium FGTS (Certificado A1)"]
        M08["Motor Vetorial de Férias"]
        M09["Parser OCR de Rescisões"]
        M10["Robô de Afastados (QSA / Óbitos)"]
        SHARED["Módulos Compartilhados & Anonimizador LGPD"]
    end

    subgraph OUTPUT["📊 Entregáveis & Tomada de Decisão"]
        REPORTS["Relatórios Excel Executivos com Marcação Condicional"]
        PBI["Dashboards Analíticos no Power BI"]
        DOCS["Dossiês Investigativos em Word (.docx)"]
        DESKTOP_APP["Interface Gráfica CustomTkinter para Analistas"]
    end

    LG --> M01 & M08 & M09 & M10
    OPERADORAS --> M01 & M08
    PDFS --> M09
    WEBSERVICE --> M10
    PORTAL_CAIXA --> M04

    M01 & M04 & M08 & M09 & M10 & SHARED --> REPORTS & PBI & DOCS & DESKTOP_APP
```

---

## 🛡️ Governança, Segurança e Conformidade com a LGPD

O repositório foi construído seguindo princípios rigorosos de **Security by Design** e **Privacidade de Dados (LGPD)**:
1. **Nenhum Dado Sensível no Versionamento:** Nenhuma planilha real (`.xlsx`, `.csv`, `.xlsb`), contracheque ou TRCT com dados de colaboradores (CPFs, nomes, salários, benefícios) é versionada no Git.
2. **Gestão de Segredos via `.env`:** Tokens de APIs (SintegraWS, InfoSimples) e credenciais de banco de dados são lidos exclusivamente a partir de variáveis de ambiente gerenciadas via `python-dotenv`.
3. **Anonimização Ativa:** Scripts de mascaramento irreversível em `utils/anonimizar_planilhas.py` garantem a criação segura de bases simuladas (*mock data*) para homologação e demonstrações públicas.

---

## 💻 Tecnologias Empregadas

* **Linguagem Principal:** Python 3.10+
* **Engenharia e Processamento de Dados:** Pandas, NumPy, OpenPyXL, PyXLSB, XLRD, XlsxWriter
* **Visão Computacional e Parsing Documental:** PDFPlumber, PyPDF, Python-Docx
* **RPA e Automação Web:** Selenium WebDriver, Windows CryptoAPI
* **Business Intelligence & Storytelling:** Microsoft Power BI, DAX, PowerPoint
* **Desenvolvimento de Interfaces Desktop:** CustomTkinter, Tkinter, Pillow
* **Integrações RESTful:** Requests, JSON Cache Engine, APIs SintegraWS e InfoSimples

---

## 🚀 Como Executar o Projeto Localmente

### 1. Clonar o Repositório
```bash
git clone https://github.com/xeuhT/meus-projetos.git
cd meus-projetos
```

### 2. Criar e Ativar o Ambiente Virtual
```bash
python -m venv .venv

# No Windows:
.venv\Scripts\activate

# No Linux / macOS:
source .venv/bin/activate
```

### 3. Instalar as Dependências
```bash
pip install -r requirements.txt
```

### 4. Configurar as Variáveis de Ambiente
Copie o template de variáveis e informe seus tokens locais:
```bash
cp .env.example .env
```

---

## 👨‍💻 Autor & Engenharia
Projeto concebido e desenvolvido por **Matheus Cavalcanti** — Especialista em Automação, Engenharia de Dados e Compliance em RH/DP.
