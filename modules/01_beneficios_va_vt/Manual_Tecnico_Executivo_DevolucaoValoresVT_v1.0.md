# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 06: Devolução de Valores VT

---

**Módulo:** 06 - Devolução de Valores VT | **Interface:** Central Executiva (`.exe`) | **Competência:** 10/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 3.820 Colaboradores com Estornos ou Erros de Pedido
* **Estrutura de Colunas:** 28 Colunas Oficiais de Reconciliação
* **Volume Financeiro Auditado:** R$ 342.180,50 em Valores Devolvidos e Compensados
* **Status do Motor no App:** 🟡 Aguardando Validação da Equipe de Benefícios

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 06 - Devolução de Valores VT** automatiza a conciliação de valores restituídos pelas operadoras de transporte coletivo e compensações decorrentes de erros de pedido, cartões inativos ou cancelamentos intempestivos.

O robô cruza os extratos de devolução da operadora contra o histórico de compras e os eventos de folha, garantindo que nenhum valor permaneça pendente sem destinação contábil ou ressarcimento ao empregado/empresa.

---

### 2. Estrutura de Pastas e Inventário de Relatórios Consumidos
O módulo é executado na pasta `Devolução de Valores`:

| Arquivo Corporativo | Formato / Peso | Finalidade e Regras Associadas |
| :--- | :---: | :--- |
| `Conferência Valores Devolvidos_102026.xlsb` | XLSB (~950 KB) | Planilha base oficial com o consolidado de devoluções das operadoras. |
| `erros_importacao.xlsx` / `erros_pedido.xlsx` | XLSX (~400 KB) | Registros de inconsistências de pedido e importação rejeitada. |
| `Relatórios Originais (CSVs)` | CSVs (~12 MB) | FOPA, COLA Afastados e Demitidos para validação do status funcional. |

---

### 3. Matriz de Regras de Negócio e Validações Matemáticas (RNs)

| Código | Tipo de Validação | Regra Aplicada & Fórmula Homologada | Base Origem | Ação Recomendada |
| :---: | :--- | :--- | :--- | :--- |
| **RN001** | Identificação da Devolução | Batimento da Matrícula, Linha e Valor restituído pela operadora de transporte. | Extrato Operadora | Mapear origem do saldo estornado. |
| **RN002** | Validação de Motivo | Classificação do estorno: Erro de Pedido, Cartão Cancelado, Demissão ou Afastamento. | Erros / Censo | Definir fluxo de compensação. |
| **RN003** | Compensação Contábil | Confronto do valor devolvido contra desconto em folha ou dedução no pedido subsequente. | FOPAG / Pedidos | Garantir quitação do crédito restituído. |
| **RN004** | Rescisão Contratual | Apuração de estornos de colaboradores já desligados para conciliação em TRCT. | COLA Demitidos | Repassar valor à conta da rescisão. |
| **RN005** | Emissão de Abas | Geração de abas `Geral`, `Divergências` e `Resumo Operacional`. | Motor Excel | Facilitar conferência pelos analistas. |

---

### 4. Cenários Didáticos Simulados (Anonimizados LGPD)

| Matrícula | Colaborador (Fictício) | Operadora | Valor Devolvido | Motivo Devolução | Diagnóstico Analítico |
| :---: | :--- | :--- | :---: | :--- | :--- |
| `600101` | COLABORADOR RESTITUIÇÃO A | SALVADOR CARD | R$ 120,00 | Cartão Cancelado | **COMPENSAR:** Abater na próxima compra mensal. |
| `600102` | COLABORADOR DEMITIDO B | BH TRANS | R$ 85,00 | Rescisão Antes da Carga | **RESCISÃO:** Repassar saldo à rescisão. |
| `600103` | COLABORADOR ERRO C | METRÔ RECIFE | R$ 220,00 | Duplicidade de Pedido | **ESTORNO OK:** Duplicidade corrigida com a operadora. |

---

### 5. Guia Operacional Passo a Passo
1. **Abrir o App:** Executar `Central_Beneficios_Cencosud.exe`.
2. **Acessar Módulo:** Selecionar **"06 - Devolução de Valores VT"**.
3. **Gerar:** Clicar em **"Gerar Conferência Devolução de Valores"**.
4. **Tratativas:** Abrir a planilha gerada e atuar nas divergências de estorno.

---

### 6. Parecer Técnico de Homologação
Automação pronta e auditada, aguardando validação final da equipe de benefícios.

**Engenharia de Dados & Automação de RH** • *Desenvolvimento de Software & Algoritmos Analíticos*  
**Administração de Pessoal & Remuneração** • *Homologação de Negócio, Regras CLT & Diretrizes Cencosud*
