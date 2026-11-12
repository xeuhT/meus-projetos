# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH - Manual Técnico e Executivo Completo

**Motor Analítico de Benefícios, Ativos e FOPAG (Projeto AD002 - Versão V003)**
**Status:** Produção Homologada | **Atualizado em:** 22/09/2026

---

## 1. Visão Geral da Solução e Arquitetura Modular

O **Projeto AD002 V003** é uma solução corporativa de Engenharia de Dados desenvolvida para realizar a conciliação automatizada 360° entre todas as fontes de dados do ecossistema de benefícios do Grupo Cencosud.

**Cobertura Integral de Dados:** O sistema consolida automaticamente mais de 35 planilhas de entrada em cada rodada de execução, cruzando faturamentos das operadoras contra rubricas descontadas em folha e cruzando com a base mestre de RH em poucos minutos.

### Arquitetura Técnica do Sistema
* **app_beneficios.py**: Interface gráfica moderna (CustomTkinter) com execução assíncrona. [Novo V003] Validação flexível que permite bypass de arquivos ausentes com alertas visuais.
* **core.py / orquestrador.py**: Maestro do fluxo ETL (Extract, Transform, Load).
* **leitores.py**: Extrator avançado resiliente. [Novo V003] Inclui *Fuzzy Matching* para cabeçalhos corrompidos, parser universal de moedas (BR/US) e expurgo automático de linhas fantasmas.
* **motor_auditoria.py**: Mecanismo de cruzamento relacional cascateado em outer/left join.
* **exportador.py**: Motor de inteligência de negócios. [Novo V003] Maior transparência financeira nas linhas de dependentes.
* **utils.py & config.py**: Módulos utilitários para sanitização estrita.

---

## 2. Inventário Completo de Bases de Entrada (Todas as Fontes)

Para que o relatório Master seja gerado com 100% de acurácia, a automação consome três grandes grupos de arquivos na pasta de entrada:

### 2.1. Bases Estruturais Corporativas e Mestre (RH / FOPAG)

| Categoria da Base | Padrão de Nome do Arquivo / Identificador | Função no Processo de Auditoria |
| --- | --- | --- |
| **Colaboradores Mestre (RH)** | COLA_Colaboradores* | Base oficial de RH contendo CPF, Matrícula, Razão Social, Cargo, Sindicato e Situação legal. |
| **Dependentes (Marsh)** | Dependentes Marsh.xlsx ou *DEPENDENTES* | Cadastro complementar de dependentes com grau de instrução, condição PCD e vínculos. |
| **Desligados (RH)** | Desligados_2026_consolidado.xlsx ou *DESL* | Histórico oficial de rescisões e datas de demissão legal para aplicação das RN001 e RN019. |
| **Dicionário de Eventos** | Eventos 1.xlsx | De-Para unificado dos códigos de rubricas financeiras da FOPAG x Operadora e Categoria. |
| **FOPAG Financeira (Aspirador)** | FOPA_Eventos_por_Colaborador*.CSV | Extrato consolidado dos descontos reais efetuados na folha de pagamento de cada funcionário. |

### 2.2. Bases de Ativos e Movimentações das Operadoras

| Operadora / Serviço | Padrão do Arquivo de Ativos / Movimentação | Estrutura Extraída |
| --- | --- | --- |
| **Bradesco Odonto** | Listagem de Ativos - Bradesco Odonto*.xlsx / Movimentações Fopag - Bradesco* | Plano, Matrícula, CPF, Titular, Beneficiário, Parentesco, Carteira, Vigência. |
| **GNDI Saúde** | Listagem de Ativos - GNDI*.xlsx / Movimentações Fopag - GNDI* | Plano, Matrícula, CPF, Titular, Beneficiário, Carteira, Posição Cadastral. |
| **Hapvida Saúde** | Listagem de Ativos - Hapvida*.xlsx / Movimentações Fopag - HAPVIDA* | Plano, Matrícula, CPF, Beneficiário, Parentesco, Vigência, Organico Unificado. |
| **Klini Saúde** | Listagem de Ativos - Klini*.xlsx / Movimentações Fopag - KLINI* | Plano, Matrícula Estipulante, CPF, Titular, Beneficiário, Parentesco, Vigência. |
| **SulAmérica Odonto** | Listagem de Ativos - Sulamerica Odonto*.xlsx / Movimentações Fopag - SulAmérica Odonto* | Plano, RE (Matrícula), CPF, Titular, Beneficiário, Bandeira/Contrato, Vigência. |
| **SulAmérica Saúde** | Listagem de Ativos - Sulamerica Saúde*.xlsx / Movimentações Fopag - SulAmerica Saúde* | Plano, ID Funcional, CPF, Titular, Beneficiário, Bandeira/Contrato, Vigência. |
| **Unimed Vale do Aço** | Listagem de Ativos - Unimed*.xlsx / Movimentações Fopag - UNIMED* | Plano, Matrícula, Beneficiário CPF, Beneficiário Nome, Início Vigência. |

### 2.3. Faturas Financeiras das Operadoras (Prêmios / Mensalidades)
Conjunto de planilhas de faturamento real enviado pelas operadoras (arquivos iniciados com VENC... DEMONSTRATIVO) para conciliação financeira com a folha (RN020).

---

## 3. Matriz de Regras de Negócio e Inconsistências (RNs)

| Código | Tipo de Inconsistência | Descrição / Regra Aplicada | Base Origem | Ação Recomendada |
| --- | --- | --- | --- | --- |
| **RN001** | Situação do Colaborador | Colaborador demitido há mais de 2 meses com benefício mantido ativo na operadora. | DESLIGADOS | Solicitar inativação imediata |
| **RN003** | Cobrança Indevida | Colaborador categorizado como Isento (Ex: Empresa GIGA no Saúde) que sofreu desconto na folha. | FOPAG | Solicitar reembolso e inativar |
| **RN004** | Eventos Assistência | Possível cobrança simultânea em mais de um plano da mesma categoria para o mesmo CPF. | ATIVOS | Verificar duplicidade |
| **RN006** | Cobrança sem Cadastro | Desconto financeiro realizado em folha sem correspondência de cadastro ativo. | FOPAG | Confirmar exclusão ou incluir |
| **RN008** | Titular Não Localizado | Dependente constando como ativo na operadora, mas o TITULAR não foi localizado nas bases. | ATIVOS | Validar vínculo do titular |
| **RN011** | Idade Limite Excedida | Dependente com idade igual ou superior a 24 anos, ou entre 21 e 23 sem vínculo universitário. | ATIVOS/MARSH | Excluir dependente ou validar |
| **RN018** | Cobrança Pós-Cancelamento | Desconto financeiro em folha realizado para plano cancelado há mais de 1 mês. | FOPAG | Solicitar reembolso do valor |
| **RN019** | Data Rescisão x Cancel. | Divergência entre a data de cancelamento na operadora e a data legal de rescisão do contrato. | ATIVOS | Ajustar vigência na operadora |
| **RN020** | Falta de Desconto | Colaborador/Dependente com faturamento ativo cobrado, mas com desconto R$ 0,00 na folha. | FATURA/FOPAG | Habilitar desconto na folha |
| **INFO** | Controle de Exclusão | Beneficiário listado na FOPAG exclusivamente para controle de cancelamento. | FOPAG | Acompanhar inativação |

---

## 4. Mecanismos Avançados de Blindagem de Dados

* **O Porteiro de Homônimos e Hierarquia Estrita:** A busca por nomes gerava colisões. O sistema agora exige validação estrita por CPF ou Matrícula exata.
* **Isolamento Familiar:** O sistema implementou a Chave de Família Isolada (OPERADORA + MATRÍCULA_TITULAR) para impedir que titulares de operadoras diferentes misturem dependentes.
* **Filtro de Isenção GIGA:** Validação de isenção para colaboradores titulares da bandeira GIGA (aplicável exclusivamente aos planos de Saúde, não isento em Odonto) e Sindicato dos Rodoviários.
* **Dicionário Estrito de Demissões:** Datas de demissão só são atribuídas à linha se a Matrícula e o Nome coincidirem com o arquivo de Desligados.

---

## 5. [NOVIDADES V003] Changelog e Melhorias de Resiliência

A **Versão 003** introduz grandes novidades para sanar falhas nas bases das operadoras:

1. **Leitura Fuzzy de Cabeçalhos (Correção de Encoding):** Arquivos da GNDI e Klini apresentavam cabeçalhos com caracteres ocultos corrompidos (ex: BENEFICIRIO sem acento). O novo algoritmo rastreia palavras-chave por substring (dependencia, benefici), garantindo 100% de ingestão independente de falhas do Excel da operadora.
2. **Parser Financeiro Universal:** Nova função de limpeza resolve conflitos de localidade entre Excel e Pandas, processando nativamente formatos US (157.84) e BR (1.500,00), erradicando bugs que multiplicavam os valores por 100.
3. **Expurgo de CPFs Fantasmas:** A FOPAG da GNDI continha resíduos (CPFs fantasmas na linha 1.000.000). A nova etapa cruza campos; se o CPF existir mas as colunas de nome estiverem vazias, a linha é descartada automaticamente.
4. **Transparência da RN018:** Para sanar falsas impressões de erros na RN018 (Dependente cobrado após cancelamento), o relatório final agora injeta os valores que o Titular pagou na linha do dependente.
5. **Validação Flexível de UI (Soft-Block):** A ausência de faturamentos de operadoras sem movimento no mês (ex: Bradesco) não trava mais a aplicação. A interface gera um alerta visual (Aviso), mas habilita o processamento das demais bases.

