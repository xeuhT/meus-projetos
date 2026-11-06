# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 09: Conferência Escritório

---

**Módulo:** 09 - Conferência Escritório (Evt 6251) | **Interface:** Central Executiva (`.exe`) | **Competência:** 09/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 1.280 Colaboradores dos Escritórios Corporativos
* **Estrutura de Colunas:** 24 Colunas Oficiais de Auditoria
* **Volume Financeiro Auditado:** R$ 185.400,00 em Benefícios Administrativos
* **Status do Motor no App:** 🟡 Aguardando Validação da Equipe de Benefícios

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 09 - Conferência Escritório** audita com exclusividade os benefícios concedidos aos colaboradores alocados nos escritórios administrativos e sedes corporativas do Grupo Cencosud Brasil (Salvador, Aracaju, Belo Horizonte, Fortaleza e São Paulo).

A ferramenta foca no desconto do **Evento 6251** e benefícios de refeição específicos do regime administrativo, cruzando a elegibilidade do pacote de benefícios contra o censo funcional e o regime de trabalho (presencial, híbrido ou teletrabalho).

---

### 2. Estrutura de Pastas e Inventário de Relatórios Consumidos
O módulo é executado na pasta `Conferencia escritorio`:

| Arquivo Corporativo | Formato / Peso | Finalidade e Regras Associadas |
| :--- | :---: | :--- |
| `Base de conferencia escritorio 092026.xlsx` | XLSX (~1.2 MB) | Planilha base homologada contendo a listagem de colaboradores administrativos. |
| `FOPA_ 23092026.xlsx` | XLSX (~2.5 MB) | Extrato da folha contendo os eventos lançados para o universo corporativo. |
| `Relatório Tratado (XLSX)` | XLSX (~800 KB) | Saída tratada com os pareceres da auditoria administrativa. |

---

### 3. Matriz de Regras de Negócio e Validações Matemáticas (RNs)

| Código | Tipo de Validação | Regra Aplicada & Fórmula Homologada | Base Origem | Ação Recomendada |
| :---: | :--- | :--- | :--- | :--- |
| **RN001** | Elegibilidade Escritório | Validação de centros de custo e cargos elegíveis ao pacote corporativo. | Censo RH | Mapear lançamentos em colaboradores não administrativos. |
| **RN002** | Batimento Evt 6251 | Confronto do desconto lançado em folha contra a tabela de coparticipação do escritório. | FOPAG Evt 6251 | Apontar divergências de cobrança. |
| **RN003** | Férias e Licenças | Expurgar descontos de refeição em períodos de ausência integral no mês. | Férias/Afastados | Estornar desconto indevido de refeição. |

---

### 4. Cenários Didáticos Simulados (Anonimizados LGPD)

| Matrícula | Colaborador (Fictício) | Lotação Corporativa | Evt 6251 Esperado | Evt 6251 Folha | Diagnóstico Analítico |
| :---: | :--- | :--- | :---: | :---: | :--- |
| `900101` | ANALISTA ADMINISTRATIVO A | SEDE SALVADOR | R$ 45,00 | R$ 45,00 | **CONFORME:** Lançamento regular escritório. |
| `900102` | COORDENADOR EM FÉRIAS B | SEDE ARACAJU | R$ 0,00 | R$ 45,00 | **DIVERGÊNCIA:** Em férias no mês civil. Estornar. |

---

### 5. Guia Operacional Passo a Passo
1. **Abrir a Central:** Executar `Central_Beneficios_Cencosud.exe`.
2. **Acessar Módulo:** Selecionar **"09 - Conferência Escritório"**.
3. **Gerar:** Clicar em **"Gerar Conferência Escritório"**.
4. **Tratar Divergências:** Filtrar divergências e enviar acertos para a folha.

---

### 6. Parecer Técnico de Homologação
Automação pronta e auditada, aguardando validação final da equipe de benefícios.

**Engenharia de Dados & Automação de RH** • *Desenvolvimento de Software & Algoritmos Analíticos*  
**Administração de Pessoal & Remuneração** • *Homologação de Negócio, Regras CLT & Diretrizes Cencosud*
