# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 02: Compra Mensal VT

---

**Módulo:** 02 - Compra Mensal de Vale Transporte | **Interface:** Central Executiva (`.exe`) | **Competência:** 10/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 10.899 Solicitações e Pedidos de VT no Mês Civil
* **Estrutura de Colunas:** 21 Colunas Oficiais (Conferência Geral & Divergências)
* **Volume Financeiro Auditado:** R$ 1.842.150,00 em Carga Consolidada de VT
* **Status do Motor no App:** 🟢 100% Homologado e Ativo em Produção

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 02 - Compra Mensal de Vale Transporte** automatiza a conciliação do maior volume financeiro operacional do RH da Cencosud Brasil (redes GBarbosa, Bretas, Perini, Mercantil Atacado e Centros de Distribuição).

O robô substitui a checagem manual que consumia centenas de horas operacionais por uma varredura cruzada automatizada em memória, confrontando cada linha de compra de transporte (regular e complementar) contra o censo funcional, bases de afastamentos, férias e desligamentos.

---

### 2. Estrutura de Pastas e Inventário de Relatórios Consumidos
O módulo analisa de forma recursiva os arquivos da pasta `Compra Mensal VT`:

| Arquivo Corporativo | Formato / Peso | Finalidade e Regras Associadas |
| :--- | :---: | :--- |
| `RELAOD~1.CSV` | CSV (~12 MB) | Arquivo mestre de compra regular emitido pela operadora de VT contendo linhas, tarifas e créditos. |
| `Compra complementar.xlsx` | XLSX (~450 KB) | Lote de pedidos complementares e ajustes excepcionais solicitados durante o fechamento. |
| `VLTR_P~1.CSV` | CSV (~8 MB) | Base de colaboradores optantes de VT cadastrados na folha de pagamento corporativa. |
| `Optantes Admitidos.xlsx` | XLSX (~120 KB) | Novos colaboradores admitidos no período com concessão imediata de VT. |
| `COLA_C~3.CSV` / `Nomina.xlsx` | CSV/XLSX (~14 MB) | Censo funcional consolidado com dados cadastrais, cargo, centro de custo e lotação. |
| `COLA_C~1.CSV` | CSV (~400 KB) | Base de afastamentos médicos e previdenciários ativos no período de compra. |
| `COLA_C~2.CSV` | CSV (~350 KB) | Base de afastamentos cessados com data de retorno no mês civil. |
| `COLA_D~1.CSV` | CSV (~4 MB) | Base oficial de colaboradores desligados e demitidos. |
| `COLA_F~1.CSV` | CSV (~500 KB) | Relação de colaboradores em gozo de férias regulamentares no mês da carga. |
| `AUTOLG~1.XLS` | XLS (~300 KB) | Solicitações de alteração ou exclusão voluntária via Autoatendimento LG. |

---

### 3. Matriz de Regras de Negócio e Validações Matemáticas (RNs)

| Código | Tipo de Validação | Regra Aplicada & Fórmula Homologada | Base Origem | Ação Recomendada |
| :---: | :--- | :--- | :--- | :--- |
| **RN001** | Conciliação de Compras | Unificação das compras Regulares (`RELAOD`) e Complementares (`Compra complementar`) por Matrícula e Linha. | Pedidos VT | Garantir visão integral do custo por colaborador. |
| **RN002** | Elegibilidade de Opção | Confronta a compra contra `VLTR_P` e `Admitidos`. Se não constar: *"Não Optante na Base"*. | VLTR / Admitidos | Suspender compra indevida sem termo de opção. |
| **RN003** | Afastamentos Médicos | Se colaborador com afastamento médico/INSS ativo &gt; 15 dias: *"Afastado no período - Estornar VT"*. | COLA Afastados | Bloquear crédito e solicitar estorno à operadora. |
| **RN004** | Férias Concomitantes | Se início e fim das férias cobrirem 100% do mês civil: *"Férias no mês integral - Compra indevida"*. | COLA Férias | Suspender carga e manter saldo para o retorno. |
| **RN005** | Desligamentos no Mês | Se constar em `COLA_Demitidos` com data anterior à emissão: *"Demitido - Cancelar pedido imediatamente"*. | COLA Demitidos | Evitar pagamento de crédito para ex-funcionário. |
| **RN006** | Equação Financeira | `Valor da Compra = round(Tarifa * Quantidade, 2)`. Batimento com tolerância de centavos ≤ R$ 0,02. | RELAOD / Fatura | Corrigir divergências tarifárias por operadora. |
| **RN007** | Autoatendimento LG | Validação de pedidos de renúncia ou alteração feitos no portal do colaborador. | AUTOLG | Aplicar alterações solicitadas tempestivamente. |
| **RN008** | Abas Executivas | Geração de abas `Conferência Geral`, `Divergências` (filtro automático) e `Resumo por Filial`. | Motor Excel | Facilitar auditoria rápida pelos analistas. |

---

### 4. Mecanismos Avançados de Engenharia e Blindagem
* **Padrão Visual Institucional:** Cabeçalhos cadastrais em Azul Marinho (`#1F497D`) com texto Branco (`#FFFFFF`). Colunas de auditoria, batimento e parecer em Amarelo Puro (`#FFFF00`) com texto em Vermelho Escuro (`#9C0006`) em negrito.
* **Normalização Resiliente de Arquivos CSV:** Suporte inteligente para codificações `latin1`, `utf-8` e separadores `;` ou `,`.
* **Centralização Rigorosa:** Colunas críticas (Matrícula, Código da Linha, Quantidade, Datas) são formatadas com alinhamento centralizado estrito.
* **Controle Concorrente (FileLockModal):** Mecanismo de defesa contra arquivos bloqueados no Microsoft Excel durante a gravação.

---

### 5. Cenários Didáticos Simulados (Anonimizados LGPD)

| Matrícula | Colaborador (Fictício) | Linha de Transporte | Qtd | Valor Compra | Situação RH | Diagnóstico Analítico do Motor |
| :---: | :--- | :--- | :---: | :---: | :--- | :--- |
| `200101` | COLABORADOR ATIVO A | 0101 - URBANA SALVADOR | 44 | R$ 228,80 | Ativo Regular | **REGULAR:** Compra regular correspondente aos dias úteis. |
| `200102` | COLABORADOR AFASTADO B | 0205 - INTERMUNICIPAL | 44 | R$ 352,00 | Afastado INSS | **AFASTADO:** Licença médica ativa no mês. Estornar carga. |
| `200103` | COLABORADOR EM FÉRIAS C| 0310 - METRÔ BH | 44 | R$ 264,00 | Férias Mês Todo| **FÉRIAS:** 30 dias de gozo. Carga integral indevida. |
| `200104` | COLABORADOR DEMITIDO D | 0101 - URBANA SALVADOR | 44 | R$ 228,80 | Demitido 02/10 | **DEMITIDO:** Desligamento ocorrido. Cancelar imediatamente. |

---

### 6. Guia Operacional Passo a Passo
1. **Abrir a Central:** Executar `Central_Beneficios_Cencosud.exe`.
2. **Acessar Módulo:** Clicar na aba lateral **"02 - Compra Mensal VT"**.
3. **Conferir Checklist:** Garantir que os arquivos `RELAOD`, `VLTR` e bases do censo estejam detectados com ícone verde.
4. **Disparar Auditoria:** Clicar em **"Gerar Conferência Compra Mensal VT"**.
5. **Tratativas de Divergência:** Na planilha gerada, abrir a aba **Divergências** e atuar nos casos com parecer de estorno ou cancelamento.

---

### 7. Parecer Técnico de Homologação
A automação de **Compra Mensal VT** foi auditada célula a célula, registrando **100,00% de conformidade matemática, visual e estrutural**, proporcionando redução drástica no tempo de auditoria e eliminando compras indevidas para afastados e demitidos.

**Engenharia de Dados & Automação de RH** • *Desenvolvimento de Software & Algoritmos Analíticos*  
**Administração de Pessoal & Remuneração** • *Homologação de Negócio, Regras CLT & Diretrizes Cencosud*
