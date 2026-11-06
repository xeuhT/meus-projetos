# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 08: Rateio CVE

---

**Módulo:** 08 - Rateio CVE (Centro de Valor / Estrutura) | **Interface:** Central Executiva (`.exe`) | **Competência:** 07/2026 e 08/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 6.450 Colaboradores com Rateio Multifilial
* **Estrutura de Colunas:** 25 Colunas Oficiais de Distribuição Contábil
* **Volume Financeiro Auditado:** R$ 1.450.000,00 em Custos Rateados por Filial
* **Status do Motor no App:** 🟡 Aguardando Validação da Equipe de Benefícios

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 08 - Rateio CVE** automatiza a partilha dos custos de benefícios corporativos (Vale Transporte e Alimentação) entre os diversos Centros de Custo, Unidades de Negócio e Estabelecimentos do Grupo Cencosud Brasil.

A frente resolve a complexidade contábil de colaboradores que prestam serviços para múltiplas filiais ou mudam de lotação no decorrer do mês civil, aplicando as frações exatas de custo a cada Centro de Custo para garantir um fechamento contábil sem distorções de resultado gerencial.

---

### 2. Estrutura de Pastas e Inventário de Relatórios Consumidos
O módulo é executado na pasta `Rateio CVE`:

| Arquivo Corporativo | Formato / Peso | Finalidade e Regras Associadas |
| :--- | :---: | :--- |
| `FOPA_Eventos_por_Colaborador.CSV` | CSV (~2.8 MB) | Base de eventos da folha contendo o rateio gerencial lançado. |
| `Optantes 20.07.xlsx` | XLSX (~1.2 MB) | Censo de optantes e definição de centros de custo alocados. |
| `Relatório Tratado (XLSX)` | XLSX (~1.5 MB) | Histórico de saídas de conferência geradas. |

---

### 3. Matriz de Regras de Negócio e Validações Matemáticas (RNs)

| Código | Tipo de Validação | Regra Aplicada & Fórmula Homologada | Base Origem | Ação Recomendada |
| :---: | :--- | :--- | :--- | :--- |
| **RN001** | Proporcionalidade de Dias | Divisão do custo do benefício pela fração de dias trabalhados em cada Centro de Custo. | Escala/Lotação | Garantir rateio exato por filial. |
| **RN002** | Soma dos Rateios = 100% | `abs(SUM(Fração_Rateios) - Custo_Total_Benefício) <= 0.02`. | Contábil | Ajustar centavos de resíduo contábil. |
| **RN003** | Validação de Código CVE | Validação da existência do Centro de Custo no plano de contas Cencosud. | Plano Contas | Evitar lançamento em contas fantasmas. |

---

### 4. Cenários Didáticos Simulados (Anonimizados LGPD)

| Matrícula | Colaborador (Fictício) | Custo Total | Centro Custo 1 (60%) | Centro Custo 2 (40%) | Diagnóstico Analítico |
| :---: | :--- | :---: | :---: | :---: | :--- |
| `800101` | COLABORADOR MULTIFILIAL A | R$ 300,00 | R$ 180,00 (Loja 101) | R$ 120,00 (Loja 105) | **CONFORME:** Rateio 60/40 perfeito centavo a centavo. |
| `800102` | COLABORADOR FIXO B | R$ 250,00 | R$ 250,00 (Loja 201) | R$ 0,00 | **CONFORME:** 100% alocado na lotação fixa. |

---

### 5. Guia Operacional Passo a Passo
1. **Abrir o App:** Executar `Central_Beneficios_Cencosud.exe`.
2. **Acessar Módulo:** Selecionar **"08 - Rateio CVE"**.
3. **Gerar:** Clicar em **"Gerar Conferência Rateio CVE"**.
4. **Validar:** Conferir se a soma de rateio fecha com a despesa corporativa total.

---

### 6. Parecer Técnico de Homologação
Automação pronta e auditada, aguardando validação final da equipe de benefícios.

**Engenharia de Dados & Automação de RH** • *Desenvolvimento de Software & Algoritmos Analíticos*  
**Administração de Pessoal & Remuneração** • *Homologação de Negócio, Regras CLT & Diretrizes Cencosud*
