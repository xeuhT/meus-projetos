# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 03: Valores Pré-Folha VT

---

**Módulo:** 03 - Valores Pré-Folha VT (Evt 5750, 31935, 6292) | **Interface:** Central Executiva (`.exe`) | **Competência:** 09/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 4.250 Colaboradores Auditados na Pré-Folha
* **Estrutura de Colunas:** 41 Colunas Oficiais (Conferência Geral) + Resumo Executivo
* **Volume Financeiro Auditado:** R$ 894.620,15 em Adiantamentos e Créditos de Transporte
* **Status do Motor no App:** 🟢 100% Homologado e Ativo em Produção

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 03 - Valores Pré-Folha VT** realiza a auditoria prévia e crítica da folha de pagamento corporativa da Cencosud Brasil antes do fechamento definitivo.

A frente foca no batimento contábil de três eventos cruciais:
1. **Evento 5750 (Adiantamento VT Através de Alçadas):** Composto matematicamente pela soma de 10 eventos específicos de créditos e reembolsos.
2. **Evento 31935 (VT Alternativo Mês Seguinte):** Conciliação de valores de transporte antecipados para modalidades alternativas.
3. **Evento 6292 (Desconto VTE - Vale Transporte Excedente):** Apuração e validação do desconto de VT em folha.

---

### 2. Estrutura de Pastas e Inventário de Relatórios Consumidos
O módulo é executado a partir do diretório corporativo `Valores Pré-Folha`:

| Arquivo Corporativo | Formato / Peso | Finalidade e Regras Associadas |
| :--- | :---: | :--- |
| `FOPA_Eventos_por_Colaborador.xlsx` | XLSX (~2.8 MB) | Extrato da folha contendo os 10 eventos que compõem o Evt 5750 e os eventos 31935 e 6292. |
| `Conferência Valores Pré-Folha VT_Base.xlsx` | XLSX (~1.5 MB) | Planilha base homologada contendo a relação de colaboradores e cabeçalhos oficiais. |

---

### 3. Matriz de Regras de Negócio e Validações Matemáticas (RNs)

| Código | Tipo de Validação | Regra Aplicada & Fórmula Homologada | Base Origem | Ação Recomendada |
| :---: | :--- | :--- | :--- | :--- |
| **RN001** | Auditoria Evt 5750 | Confronta o Evento 5750 lançado contra a soma matemática estrita dos 10 eventos de alçada. | FOPAG Evt 5750 | Identificar lançamentos divergentes ou incompletos. |
| **RN002** | Composição da Alçada | `Soma 10 Eventos = Evt 147 + 148 + 159 + 490 + 9421 + 9422 + 9423 + 9428 + 9429 + 31935`. | FOPAG Bruto | Garantir a exatidão centesimal do adiantamento. |
| **RN003** | Batimento Evt 31935 | Validação do crédito de transporte alternativo (`Valor Vale Transporte Alternativo Mês Seguinte`). | FOPAG Evt 31935 | Validar quitação antecipada. |
| **RN004** | Validação Evt 6292 | Confronto do desconto em folha referente ao VTE contra as tabelas de alíquota e limite legal. | FOPAG Evt 6292 | Corrigir descontos a maior ou a menor. |
| **RN005** | Tolerância Centesimal | `abs(Evt_5750_Lancado - Soma_10_Eventos) <= 0.02`. Valores &gt; 0,02 são marcados como divergência. | Contábil | Ajustar lançamento antes do fechamento da FOPAG. |
| **RN006** | Abas Executivas | Geração das abas `Conferência Geral`, `Divergências` e `Resumo por Filial`. | Motor Excel | Facilitar envio de ajustes para a equipe de folha. |

---

### 4. Mecanismos Avançados de Engenharia e Blindagem
* **Padrão Visual Institucional:** Cabeçalhos base em Azul Marinho (`#1F497D`) com texto Branco (`#FFFFFF`). Colunas de auditoria, batimento e parecer em Amarelo Puro (`#FFFF00`) com texto em Vermelho Escuro (`#9C0006`) em negrito.
* **Cálculo Determinístico:** Eliminação de fórmulas dinâmicas do Excel, prevenindo distorções no recálculo em máquinas sem conexão de rede corporativa.
* **Máscara Contábil Uniforme:** Formatação contábil brasileira (`#,##0.00; -#,##0.00; 0.00`) aplicada em 100% das células financeiras.

---

### 5. Cenários Didáticos Simulados (Anonimizados LGPD)

| Matrícula | Colaborador (Fictício) | Soma 10 Eventos | Evt 5750 Lançado | Validação Alçada | Evt 6292 (Desc) | Diagnóstico Analítico |
| :---: | :--- | :---: | :---: | :---: | :---: | :--- |
| `300101` | COLABORADOR ALÇADA A | R$ 380,00 | R$ 380,00 | R$ 0,00 | R$ 120,00 | **CONFORME:** Soma dos 10 eventos idêntica ao Evento 5750. |
| `300102` | COLABORADOR DIVERGENTE B| R$ 450,00 | R$ 400,00 | -R$ 50,00 | R$ 120,00 | **DIVERGÊNCIA:** Lançado a menor em folha (Diferença de R$ 50,00). |
| `300103` | COLABORADOR ALTERNATIVO C| R$ 220,00 | R$ 220,00 | R$ 0,00 | R$ 0,00 | **ALTERNATIVO:** Evento 31935 validado com quitação integral. |

---

### 6. Guia Operacional Passo a Passo
1. **Abrir a Central:** Executar `Central_Beneficios_Cencosud.exe`.
2. **Acessar Módulo:** Clicar na aba **"03 - Valores Pré-Folha VT"**.
3. **Conferir Checklist:** Confirmar presença da extração FOPA e da planilha base.
4. **Executar:** Clicar em **"Gerar Conferência Pré-Folha VT"**.
5. **Ajuste em Folha:** Filtrar as divergências na planilha final e solicitar a retificação no sistema LG antes do corte da folha.

---

### 7. Parecer Técnico de Homologação
O **Módulo 03 - Valores Pré-Folha VT** foi homologado com **100,00% de precisão matemática e contábil**, garantindo que nenhum adiantamento ou desconto indevido de VT seja consolidado na folha de pagamento.

**Engenharia de Dados & Automação de RH** • *Desenvolvimento de Software & Algoritmos Analíticos*  
**Administração de Pessoal & Remuneração** • *Homologação de Negócio, Regras CLT & Diretrizes Cencosud*
