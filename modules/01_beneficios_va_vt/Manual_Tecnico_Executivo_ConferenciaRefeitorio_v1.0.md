# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 07: Conferência Refeitório (Evt 6251)

---

**Módulo:** 07 - Conferência Refeitório (Evt 6251) | **Interface:** Central Executiva (`.exe`) | **Competência:** 09/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 8.940 Colaboradores com Desconto de Refeição
* **Estrutura de Colunas:** 22 Colunas Oficiais de Validação
* **Volume Financeiro Auditado:** R$ 268.200,00 em Descontos de Alimentação em Folha
* **Status do Motor no App:** 🟡 Aguardando Validação da Equipe de Benefícios

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 07 - Conferência Refeitório (Evt 6251)** audita os descontos em folha referentes à coparticipação em refeições corporativas fornecidas nas lojas, centros de distribuição e cozinhas industriais do Grupo Cencosud Brasil.

A ferramenta confronta os lançamentos do **Evento 6251 (Desconto Refeitório)** contra as regras do Programa de Alimentação do Trabalhador (PAT), Acordos Coletivos de Trabalho (ACT) e dias efetivamente trabalhados, expurgando descontos em dias de férias, folgas ou licenças médicas.

---

### 2. Estrutura de Pastas e Inventário de Relatórios Consumidos
O módulo é executado a partir de `Conferencia Refeitorio`:

| Arquivo Corporativo | Formato / Peso | Finalidade e Regras Associadas |
| :--- | :---: | :--- |
| `FOPA_Eventos_por_Colaborador.xlsx` | XLSX (~2.8 MB) | Extrato da folha contendo o Evento 6251 lançado por colaborador. |
| `Conferência Refeitório_Base.xlsx` | XLSX (~1.2 MB) | Planilha base com registros de catraca e controle de refeições servidas. |
| `COLA_Ferias_Colaboradores.xlsx` / `COLA_Afastados.xlsx` | XLSX (~600 KB) | Mapeamento de afastamentos e férias para dedução de consumo indevido. |

---

### 3. Matriz de Regras de Negócio e Validações Matemáticas (RNs)

| Código | Tipo de Validação | Regra Aplicada & Fórmula Homologada | Base Origem | Ação Recomendada |
| :---: | :--- | :--- | :--- | :--- |
| **RN001** | Elegibilidade e Teto PAT | Validação da alíquota legal de desconto sobre a faixa salarial do empregado. | PAT / ACT | Garantir conformidade com as convenções sindicais. |
| **RN002** | Confronto Evt 6251 | `Diferença = Evt_6251_Folha - Valor_Calculado_Refeitorio`. Tolerância ≤ R$ 0,02. | FOPAG Evt 6251 | Apontar desconto a maior ou a menor em folha. |
| **RN003** | Férias e Afastamentos | Se colaborador esteve 100% afastado ou em férias no mês: `Evt 6251 deve ser R$ 0,00`. | Férias/Afastados | Estornar desconto indevido de refeição. |
| **RN004** | Emissão Executiva | Geração de abas com lista de inconsistências e resumo financeiro por filial. | Motor Excel | Facilitar envio de lote de correção para a folha. |

---

### 4. Cenários Didáticos Simulados (Anonimizados LGPD)

| Matrícula | Colaborador (Fictício) | Dias Trabalhados | Desconto Calculado | Evt 6251 Folha | Diagnóstico Analítico |
| :---: | :--- | :---: | :---: | :---: | :--- |
| `700101` | COLABORADOR REFEIÇÃO A | 22 dias | R$ 30,00 | R$ 30,00 | **CONFORME:** Desconto integral regular. |
| `700102` | COLABORADOR EM FÉRIAS B| 0 dias (Férias) | R$ 0,00 | R$ 30,00 | **DIVERGÊNCIA:** Férias no mês todo. Desconto indevido em folha. |
| `700103` | COLABORADOR AFASTADO C | 0 dias (INSS) | R$ 0,00 | R$ 0,00 | **CONFORME:** Afastado, sem desconto lançado. |

---

### 5. Guia Operacional Passo a Passo
1. **Abrir o App:** Executar `Central_Beneficios_Cencosud.exe`.
2. **Acessar Módulo:** Clicar em **"07 - Conferência Refeitório"**.
3. **Gerar:** Clicar em **"Gerar Conferência Refeitório"**.
4. **Analisar:** Filtrar as divergências para estorno em folha.

---

### 6. Parecer Técnico de Homologação
Automação pronta e auditada, aguardando validação final da equipe de benefícios.

**Engenharia de Dados & Automação de RH** • *Desenvolvimento de Software & Algoritmos Analíticos*  
**Administração de Pessoal & Remuneração** • *Homologação de Negócio, Regras CLT & Diretrizes Cencosud*
