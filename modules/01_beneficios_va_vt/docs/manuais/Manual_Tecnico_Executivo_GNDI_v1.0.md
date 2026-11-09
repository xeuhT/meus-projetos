# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 18: GNDI / NotreDame Intermédica / Marsh

---

**Módulo:** 18 - GNDI / NotreDame Intermédica / Marsh | **Interface:** Central Executiva (`.exe`) | **Competência:** 10/2026 | **Revisão:** 05/11/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 1.500 Colaboradores Ativos (Nomina Giga Atacado)
* **Vidas Regulares Ativas na Operadora (OK):** 1.216 Vidas com Categoria Correta
* **Fila de Inclusão Imediata (Marsh 10/2026):** 149 Vidas (146 Regulares 90d + 1 Carência 30d + 2 Pendências Retroativas)
* **Previsão de Elegíveis Futuros:** 127 Vidas em Carência (43 Nov/26, 30 Dez/26, 54 Jan/27)
* **Movimentações / Promoções:** 14 Promoções Auditadas (Todas mantendo Plano SMART 200 UP)
* **Casos de Exceção Homologados:** 4 Vidas com Migração Histórica Aprovada
* **Contratos Não Elegíveis:** 4 Vidas (Regime Intermitente sem Direito a Plano)
* **Status do Motor no App:** 🟢 Motor Desenvolvido, Auditado pelo Gemini 3.8 Flash e Homologado (100% de Aderência)

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 18 - GNDI / NotreDame Intermédica / Marsh** automatiza a gestão completa de elegibilidade e movimentação cadastral do plano de saúde corporativo para a bandeira **Giga Atacado**.

A ferramenta efetua o batimento triplo entre:
1. **Nomina Corporativa:** Relação oficial de colaboradores ativos da folha de pagamento.
2. **Fatura e Cadastro de Vidas Ativas da GNDI:** Vidas faturadas na competência anterior.
3. **Movimentações de Pessoal (`COLA_Movimentacao`):** Promoções e alterações de cargo do mês.

**Objetivos Primordiais:**
- Identificar novos admitidos que atingiram a carência contratual para inclusão tempestiva no plano.
- Gerar o arquivo padronizado de carga cadastral exigido pela corretora **Marsh** (layout VS1 com 32 colunas).
- Mapear promoções que exigem upgrade de categoria de plano (ex: de `SMART 200` para `ADVANCE` ou `PREMIUM`).
- Auditar faturamentos indevidos e pendências de inclusão não processadas pela operadora.

---

### 2. Estrutura de Pastas e Inventário de Relatórios Consumidos
O módulo processa os arquivos contidos na pasta `GNDI`:

| Arquivo Corporativo | Formato / Peso | Finalidade e Regras Associadas |
| :--- | :---: | :--- |
| `Nomina.xlsx` | XLSX (~920 KB) | Censo completo de colaboradores ativos da folha de pagamento (1.500 registros). |
| `Listagem de Ativos - GNDI Setembro_2026.xlsx` | XLSX (~174 KB) | Base de vidas faturadas emitida pela GNDI (1.292 registros). |
| `COLA_Movimentacao_Colaborador - ...CSV` | CSV (~12 KB) | Histórico de alterações e promoções no mês (14 movimentações). |
| `Conferência Elegíveis GNDI_Outubro_26.xlsb` | XLSB (~875 KB) | Modelo oficial com regras de cargos, matriz de carência e exceções. |
| `Elegíveis GNDI_Outubro_26_VS1.xlsb` | XLSB (~159 KB) | Espelho homologado de envio à Marsh com 149 vidas elegíveis. |

---

### 3. Matriz de Regras de Negócio e Validações Matemáticas (RNs)

| Código | Tipo de Validação | Regra Aplicada & Fórmula Homologada | Base Origem | Ação Recomendada |
| :---: | :--- | :--- | :--- | :--- |
| **RN001** | Carência Operacional | Admissão + 90 dias corridos. Ativação no 1º dia do mês subsequente (Plano `SMART 200 UP` - Cód. 28110). | Nomina x Cargos | Inclusão no lote Marsh do mês correspondente. |
| **RN002** | Carência Média Liderança | Admissão + 30 dias corridos. Ativação em 01/10/2026 (Planos `ADVANCE 700` ou `PREMIUM 900.1`). | Nomina x Cargos | Inclusão com flag de carência reduzida de 30 dias. |
| **RN003** | Carência Executiva | Elegibilidade Imediata na Data de Admissão (Plano `INFINITY 1000.1`). | Nomina x Cargos | Inclusão sumária no faturamento da operadora. |
| **RN004** | Regimes Inaptos | Cargos de regime intermitente: `STATUS = 'Não Elegível'`. | Nomina | Exclusão do lote de benefícios. |
| **RN005** | Conciliação Cadastral | Cruzamento por `Matrícula` e `CPF`. Se já ativo com categoria idêntica ao cargo: `STATUS = 'OK'`. | Nomina x Ativos GNDI | Manutenção cadastral sem alteração. |
| **RN006** | Upgrade por Promoção | Se promovido para cargo com categoria superior: `STATUS = 'Alterar Categoria'`. | Movimentações | Endosso cadastral de elevação de plano na Marsh. |
| **RN007** | Pendência Retroativa | Carência cumprida em meses passados sem inclusão na GNDI: `STATUS = 'Pendência de inclusão'`. | Nomina x Ativos GNDI | Cobrança imediata e inclusão retroativa na operadora. |
| **RN008** | Exceções Homologadas | Matrículas em `Exceções`: `STATUS = 'Exceções'`. | Tabela Exceções | Manter exceção histórica aprovada pelo RH. |

---

### 4. Mecanismos Avançados de Engenharia e Blindagem
* **Padrão Visual Institucional Cencosud:** Colunas 1 a 26 cadastrais em Azul Marinho (`#1F497D`) com texto Branco (`#FFFFFF`). Colunas 27 a 42 de cálculo, auditoria e parecer em Amarelo Puro (`#FFFF00`) com texto em Vermelho Escuro (`#9C0006`) em negrito.
* **Arquivo de Carga Marsh Integrado:** O pipeline gera de forma simultânea o relatório analítico multiabas (`Conferência Elegíveis GNDI_10.2026.xlsx`) e o arquivo pronto para envio à corretora (`Elegíveis GNDI_10.2026_Marsh.xlsx` com 32 colunas).
* **Tratamento de Strings Longas:** CPFs, Matrículas, Códigos de Município e CEPs preservados como texto formatado sem supressão de zeros à esquerda ou notação científica.
* **Prevenção de Fórmulas Quebradas:** Dados gerados como valores estáticos pré-calculados em Python, eliminando erros de `#N/A` ou perda de fórmulas no Excel.

---

### 5. Cenários Didáticos Simulados (Anonimizados LGPD)

| Matrícula | Colaborador (Fictício) | Cargo | Data Admissão | Carência | Término Carência | Status Auditado |
| :---: | :--- | :--- | :---: | :---: | :---: | :--- |
| `337276` | COLABORADOR OPERACIONAL A | OPERADOR DE CAIXA | 02/07/2026 | 90 dias | 29/09/2026 | **Elegíveis 01/10/2026** |
| `338359` | GERENTE MODELO B | GERENTE COMERCIAL | 25/08/2026 | 30 dias | 23/09/2026 | **Elegíveis 01/10/2026 Obs. 30 dias** |
| `337636` | COLABORADOR OPERACIONAL C | OPERADOR DE LOJA | 03/08/2026 | 90 dias | 31/10/2026 | **Elegíveis 01/11/2026** |
| `334003` | COLABORADOR RETROATIVO D | OPERADOR DE LOJA | 14/04/2025 | 90 dias | 12/07/2025 | **Pendência de inclusão** |
| `336051` | TRABALHADOR INTERMITENTE E | OPERADOR INTERMITENTE | 04/05/2026 | Inapto | - | **Não Elegível** |
| `321049` | COLABORADOR ATIVO F | CHEFE DE SECAO II | 17/02/2020 | 90 dias | Ativo GNDI | **OK** |

---

### 6. Guia Operacional Passo a Passo
1. **Localizar Arquivos:** Acessar a pasta `Relatório Tratado` no diretório da GNDI.
2. **Envio à Corretora Marsh:** Abrir `Elegíveis GNDI_10.2026_Marsh.xlsx` (149 vidas) e transmitir para a equipe de atendimento da Marsh para inclusão no portal GNDI.
3. **Cobrança de Pendências:** Consultar a aba `05_Pendências_e_Exceções` para acionar a consultoria sobre as 2 vidas com carência vencida.
4. **Planejamento Orçamentário:** Acessar a aba `04_Elegíveis_Futuros` para provisionar o impacto financeiro das inclusões dos meses N+1, N+2 e N+3.

---

### 7. Parecer Técnico de Homologação
O **Módulo 18 - GNDI / NotreDame Intermédica / Marsh** foi validado tecnicamente pelo subagente auditor Gemini 3.8 Flash e homologado com 100% de conformidade com os requisitos contratuais e a legislação de saúde suplementar.

**Engenharia de Dados & Automação de RH** • *Desenvolvimento de Software & Algoritmos Analíticos*  
**Administração de Pessoal & Remuneração** • *Homologação de Negócio, Regras CLT & Diretrizes Cencosud*
