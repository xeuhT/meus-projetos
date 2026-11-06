# 🖥️ Módulo 12: App Desktop de Gestão e Auditoria de Benefícios (CustomTkinter)

![Status](https://img.shields.io/badge/Status-Concluído-brightgreen)
![UI](https://img.shields.io/badge/GUI-CustomTkinter-blue)

## 📌 Objetivo do Módulo
Aplicação desktop com interface moderna e amigável desenvolvida em CustomTkinter para permitir que analistas de RH e Departamento Pessoal operem auditorias de benefícios de saúde e odontologia (SulAmérica, GNDI, Hapvida, Unimed, OdontoPrev) sem necessidade de comandos em terminal ou conhecimento técnico de programação.

## ⚙️ Funcionalidades da Interface
* **Seleção Visual de Arquivos:** Seleção intuitiva de faturas de operadoras e arquivos de folha via diálogos do sistema.
* **Execução em Thread Separada:** Processamento não-bloqueante com barra de progresso em tempo real e console integrado de telemetria.
* **Exportação com 1 Clique:** Geração automatizada de relatórios em Excel com formatação executiva condicional.
* **Documentação Técnica Completa:** Manual técnico operacional incluso ([Manual Técnico AD002 V003](Manual_Tecnico_AD002_V003.pdf)).

## 📂 Arquivos do Módulo
* `app_beneficios.py`: Aplicação desktop principal em CustomTkinter.
* `leitores.py`: Ingestão de múltiplos formatos de faturas de operadoras de saúde e odonto.
* `motor_auditoria.py`: Motor de confronto cadastral e financeiro.
* `exportador.py`: Formatação e estilização de relatórios em Excel.
* `config.py`: Definições visuais, paletas de cores e configurações operacionais.
* `utils.py`: Formatadores de moeda, strings e validação de CPFs.
* `logo.png`: Identidade visual da aplicação.
* `Manual_Tecnico_AD002_V003.pdf`: Documentação técnica detalhada de operação e regras.

## 🚀 Como Executar
```bash
python app_beneficios.py
```
