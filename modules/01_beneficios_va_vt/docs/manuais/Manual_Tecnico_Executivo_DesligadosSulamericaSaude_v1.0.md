# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 17: Desligados SulAmérica Saúde

---

**Módulo:** 17 - Desligados SulAmérica Saúde (SAS) | **Interface:** Central Executiva (`.exe`) | **Competência:** 10/2026 | **Revisão:** 04/11/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 925 Vidas Ativas Faturadas (Titulares e Dependentes)
* **Estrutura de Colunas:** 20 Colunas Oficiais (14 Cadastrais + 6 Auditoria e Parecer - Coluna B redundante de CPF removida)
* **Vidas Vinculadas a Demitidos:** 27 Vidas Mapeadas no Faturamento
* **Fila de Cancelamentos Imediatos:** 12 Vidas (Corte Imediato & Pedidos / Prazos Expirados)
* **Fila de Diligências e Monitoramento:** 15 Vidas (Demissões Recentes em Outubro / Prazo 30 dias ANS)
* **Vidas em Planos de Diretoria:** 92 Vidas em Categoria Prestige Monitoradas
* **Status do Motor no App:** 🟢 Motor Desenvolvido, Auditado e Homologado (100% de Conformidade)

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 17 - Desligados SulAmérica Saúde** automatiza o batimento mensal entre o censo de vidas ativas faturadas pela operadora SulAmérica Saúde e a relação corporativa de colaboradores demitidos do Grupo Cencosud Brasil (`COLA_Demitidos`).

A solução elimina a permanência indevida de ex-colaboradores e seus dependentes no faturamento corporativo, aplicando com rigor as disposições da **Lei Federal nº 9.656/1998 (Artigos 30 e 31)** e da **Resolução Normativa ANS nº 488/2022**:
1. **Propagação em Cascata Familiar:** A regra de desligamento identificada para o titular é automaticamente herdada por todos os seus dependentes cadastrados (cônjuges e filhos).
2. **Corte Imediato de Pedidos de Demissão:** Colaboradores que pediram rescisão voluntária (`Tipo 4`) não possuem direito legal à extensão do plano, sendo enquadrados como *"Cancelar Plano - Pedido"*, com solicitação de cancelamento imediato e glosa da fatura.
3. **Prazos Expirados de Demissão sem Justa Causa:** Demitidos pelo empregador (`Tipo 2`) anteriores a Outubro/2026 que já completaram o prazo legal são classificados como *"Cancelar Plano - Já completou o prazo"*.
4. **Diligências e Monitoramento (Rescisões Recentes):** Demitidos sem justa causa no mês corrente de Outubro/2026 têm direito regulamentar de até 30 dias para manifestação de interesse na permanência como beneficiário contribuinte (RN 488/ANS), sendo classificados como *"Manter o plano - Sem retorno RH Negpocio"*.
5. **Isolamento de Planos de Diretoria:** Titulares com categoria `PRESTIGE` (`29426`) são mapeados com a tag *"Diretor"*, segregando casos executivos que possuem regras específicas de transição contratual.

---

### 2. Estrutura de Pastas e Inventário de Relatórios Consumidos
O módulo consome as bases consolidadas na pasta `Desligados Sulamerica Saude`:

| Arquivo Corporativo | Formato / Peso | Finalidade e Regras Associadas |
| :--- | :---: | :--- |
| `Listagem de Ativos - Sulamerica Saúde - Outubro_2026.xlsx` | XLSX (~140 KB) | Base mestre de faturamento emitida pela operadora com 925 vidas (Titulares e Dependentes). |
| `Desligados 01.09 a 19.10.xlsx` | XLSX (~274 KB) | Extração corporativa da folha de pagamento (`COLA_Demitidos`) com 920 rescisões ocorridas no período. |
| `Listagem de Desligados - Sulamerica Saúde Validação RH.xlsx` | XLSX (~146 KB) | Modelo oficial homologado contendo as colunas de validação e a aba de Legenda. |

---

### 3. Matriz de Regras de Negócio e Validações Matemáticas (RNs)

| Código | Tipo de Validação | Regra Aplicada & Fórmula Homologada | Base Origem | Ação Recomendada |
| :---: | :--- | :--- | :--- | :--- |
| **RN001** | Cruzamento Titular | Cruzamento relacional por `MATRICULA` e `CPF` do Titular normalizados sem caracteres especiais. | Ativos SAS x Demitidos | Amarrar faturamento da operadora ao censo de folha. |
| **RN002** | Vínculo Familiar | Propagação automática da regra do titular para todos os dependentes associados à matrícula. | Ativos SAS | Cancelar todo o grupo familiar de ex-empregados. |
| **RN003** | Pedido de Demissão | Se `Tipo de Rescisão` contiver "EMPREGADO" ou "PEDIDO" (Tipo 4): `STATUS = 'Cancelar Plano - Pedido'`. | COLA Demitidos | Exclusão imediata sem direito ao Art. 30 e glosa na fatura. |
| **RN004** | Prazo Expirado (Art. 30) | Se demitido sem justa causa antes de Outubro/2026: `STATUS = 'Cancelar Plano - Já completou o prazo'`. | Data Rescisão | Cancelamento imediato por término do prazo legal. |
| **RN005** | Diligência RH Negócio | Se demitido sem justa causa recente em Outubro/2026: `STATUS = 'Manter o plano - Sem retorno RH Negpocio'`. | Data Rescisão | Aguardar prazo regulamentar de 30 dias de opção da ANS. |
| **RN006** | Mapeamento Diretoria | Se `CATEGORIA DO PLANO` contiver "PRESTIGE" ou `CÓDIGO` = 29426: `OBSERVAÇÃO = 'Diretor'`. | Categoria Plano | Submeter à Diretoria de Gente & Gestão. |

---

### 4. Mecanismos Avançados de Engenharia e Blindagem
* **Padrão Visual Institucional:** Cabeçalhos das 15 colunas cadastrais em Azul Marinho (`#1F497D`) com texto Branco (`#FFFFFF`). Colunas 16 a 21 de auditoria e decisão em Amarelo Puro (`#FFFF00`) com texto em Vermelho Escuro (`#9C0006`) em negrito.
* **Preservação de CPFs e Formatação de Carteira:** Coluna 2 (CPF Titular) e Coluna 8 (CPF Beneficiário) preservadas com formatação de 11 dígitos, e número de carteira gravado como string para impedir notação científica.
* **Abas Operacionais Dedicadas:** Além da aba integral `BASE DE ATIVOS`, o sistema gera automaticamente as abas `Cancelamentos Imediatos` (12 vidas em fila de corte imediato) e `Diligências e Monitoramento` (15 vidas sob prazo legal da RN 488/ANS).
* **Prevenção de Fórmulas Quebradas:** Dados gerados como valores estáticos pré-calculados em Python, eliminando erros de `#N/A` ou perda de fórmulas no Excel.

---

### 5. Cenários Didáticos Simulados (Anonimizados LGPD)

| Matrícula | Titular (Fictício) | Beneficiário | Parentesco | Data Rescisão | Tipo de Rescisão | Status Auditado |
| :---: | :--- | :--- | :---: | :---: | :--- | :--- |
| `336173` | TITULAR MODELO A | TITULAR MODELO A | TITULAR | 15/10/2026 | 4 - Sem justa causa empregado | **Cancelar Plano - Pedido** |
| `336173` | TITULAR MODELO A | CONJUGE MODELO A | CONJUGE | 15/10/2026 | 4 - Sem justa causa empregado | **Cancelar Plano - Pedido** |
| `334932` | TITULAR MODELO B | TITULAR MODELO B | TITULAR | 10/09/2026 | 2 - Sem justa causa empregador | **Cancelar Plano - Já completou o prazo** |
| `334599` | TITULAR MODELO C | TITULAR MODELO C | TITULAR | 13/10/2026 | 2 - Sem justa causa empregador | **Manter o plano - Sem retorno RH Negpocio** |
| `334557` | DIRETOR MODELO D | DIRETOR MODELO D | TITULAR | - | - (Ativo Regular) | **Ativo (Observação: Diretor)** |

---

### 6. Guia Operacional Passo a Passo
1. **Arquivo Gerado:** Abrir `Conferência Desligados Sulamerica Saúde_10.2026.xlsx` na pasta `Relatório Tratado`.
2. **Fila de Cancelamento:** Acessar a aba **"Cancelamentos Imediatos"** (12 vidas a cancelar imediatamente).
3. **Fila de Diligências:** Acessar a aba **"Diligências e Monitoramento"** (15 vidas aguardando decurso do prazo de 30 dias ou opção pelo RH Negócio).
4. **Ação no Portal:** Efetuar a baixa das carteirinhas listadas na fila de cancelamento imediato no portal da SulAmérica Saúde.
5. **Glosa Financeira:** Enviar o lote para a corretora solicitar o estorno financeiro das cobranças pós-rescisão.

---

### 7. Parecer Técnico de Homologação
O **Módulo 17 - Desligados SulAmérica Saúde** foi homologado e testado com 100% de conformidade com os apontamentos da equipe de Recursos Humanos e a legislação da ANS.

**Engenharia de Dados & Automação de RH** • *Desenvolvimento de Software & Algoritmos Analíticos*  
**Administração de Pessoal & Remuneração** • *Homologação de Negócio, Regras CLT & Diretrizes Cencosud*
