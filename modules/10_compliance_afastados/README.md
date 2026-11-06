# 🤖 Módulo 10: Compliance Cadastral, SST, INSS e Robô de Afastados

![Status](https://img.shields.io/badge/Status-Concluído-brightgreen)
![Integração](https://img.shields.io/badge/API-SintegraWS%20%7C%20Receita%20Federal-blue)

## 📌 Objetivo do Módulo
Sistema automatizado de compliance cadastral e inteligência preventiva de riscos para monitoramento de colaboradores afastados pelo INSS (auxílio-doença, B91 acidentário, aposentadoria por invalidez), verificação de participação societária ativa (QSA) e detecção de óbitos não comunicados ao RH.

## ⚙️ Regras de Negócio e Funcionalidades
* **Consulta Automatizada em Lote via REST API:** Integração com o SintegraWS (e InfoSimples) para validação cadastral junto à Receita Federal e Juntas Comerciais.
* **Detecção de Fraudes e Conflitos de Interesse:**
  - Identificação de colaboradores afastados pelo INSS que figuram como sócios administradores ou titulares de empresas ativas (QSA), caracterizando incompatibilidade com o benefício previdenciário.
  - Alerta de colaboradores com inscrição cancelada por óbito na Receita Federal sem baixa na folha de pagamento.
* **Segurança e Sanitização:**
  - O script foi sanitizado para utilizar a variável de ambiente `SINTEGRAWS_TOKEN` via `os.getenv("SINTEGRAWS_TOKEN")`, eliminando tokens expostos no código.
* **Dossiês Executivos Automatizados:** Geração de laudos individuais em Word (`.docx`) com histórico da investigação e consolidação gerencial em Excel.

## 📂 Arquivos e Scripts do Módulo
* `executar_lote_completo.py`: Orquestrador principal da esteira de consultas e processamento em lote.
* `processar_auditoria_inss.py`: Análise das regras de benefício e prazos de retorno/perícia.
* `processar_planilha.py`: Carga, limpeza e enriquecimento da base de afastados.
* `atualizar_qsa.py`: Módulo especialista em Quadro de Sócios e Administradores.
* `atualizar_documentos.py`: Gerador de dossiês em Word e relatórios consolidados.

## 🚀 Como Executar
1. Configure o token no arquivo `.env`:
   ```bash
   SINTEGRAWS_TOKEN=seu_token_aqui
   ```
2. Execute a esteira:
   ```bash
   python executar_lote_completo.py
   ```
