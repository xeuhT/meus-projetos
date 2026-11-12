# 🌐 Módulo 04: Automação do Conectividade Social / FGTS (Selenium + Certificado A1)

![Status](https://img.shields.io/badge/Status-Concluído-brightgreen)
![Tecnologia](https://img.shields.io/badge/Stack-Selenium%20%7C%20Windows%20CryptoAPI-purple)

## 📌 Objetivo do Módulo
Automação robótica para navegação no portal Conectividade Social ICP V2 da Caixa Econômica Federal utilizando Selenium WebDriver integrado à camada nativa de certificados digitais do Windows (CryptoAPI / Certificado A1 / e-CNPJ), permitindo a extração automatizada em lote de Extratos Rescisórios e Saldo de FGTS para subsidiar rescisões contratuais.

## ⚙️ Arquitetura e Desafios Superados
* **Seleção Nativa de Certificado no Windows:** Bypass do modal nativo do navegador para seleção do Certificado Digital A1 via automação de eventos do sistema operacional.
* **Resiliência a Instabilidades da Caixa:** Mecanismo de retry com backoff exponencial e detecção de páginas de indisponibilidade/sessão expirada.
* **Extração Tabular em Lote:** Download silencioso, conversão dos relatórios HTML/TXT e consolidação direta em Excel com metadados rescisórios.

## 📂 Arquivos e Scripts do Módulo
* `main.py`: Ponto de entrada de execução da esteira.
* `browser_automation.py`: Driver Selenium configurado com profiles resilientes.
* `win_cert_helper.py`: Helper de automação de diálogo de Certificado A1 do Windows.
* `extractor.py`: Parser de páginas, extratos e tabelas do portal Caixa.
* `business_rules.py`: Regras de negócio de apuração de depósitos em aberto e saques.
* `excel_handler.py`: Exportação consolidada de dados de FGTS.
* `app_gui.py`: Interface gráfica opcional para operação simples pela equipe de RH.
* `config.py`: Parâmetros de timeouts, URLs do portal e caminhos de download.

## 🚀 Como Executar
```bash
python main.py
```
*(Requer ambiente Windows com navegador Google Chrome e Certificado A1 instalado ou configurado).*
