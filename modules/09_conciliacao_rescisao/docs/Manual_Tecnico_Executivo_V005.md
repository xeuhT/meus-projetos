# Manual Técnico e Executivo — Módulo AD003 (Versão V005)
## Sistema Analítico de Auditoria Rescisória & FOPAG — Cencosud Brasil
### Arquitetura de Governança, Validação de Faltas vs. Data de Desligamento e Relatório Institucional em 4 Abas

---

## 1. Visão Geral e Contexto Executivo

O **Módulo AD003 (Versão V005)** consolida o motor analítico de auditoria rescisória da Cencosud Brasil, unificando em um único fluxo automatizado:
1. **Recálculo Matemático CLT & CCT (100% Homologado):** Cobertura integral das regras legais de saldo de salário, aviso-prévio proporcional (Lei 12.506/11), férias vencidas/proporcionais com 1/3, 13º salário proporcional/indenizado, adicionais de tempo de serviço (Triênios/Anuênios - Lote 1), adicionais operacionais (Produtividade, Farmácia, Teletrabalho - Lote 2), quebra de caixa sindical, faltas de caixa e margem legal de consignados (Lote 3), comissionistas e créditos de cesta básica/creche (Lote 4), pensões alimentícias judiciais com desconto em cascata (Lote 5), e suficiência de saldo com impenhorabilidade do Salário Família (Evento 499).
2. **Nova Regra de Validação de Ausências vs. Competência do Desligamento:** Bloqueio preventivo e alerta para casos em que o total de faltas injustificadas (`5001` / `20010`) somado aos DSRs descontados (`5007` / `20016`) excede os dias corridos até a data de desligamento no mês da rescisão.
3. **Novo Padrão Visual e Operacional Consolidado em 4 Abas:** Redução estratégica de abas fragmentadas para um fluxo objetivo e direto, focado em **Divergências e Casos Atípicos**, com cockpit de liberação para a liderança e conformidade total com o padrão institucional Cencosud (Azul Petróleo `#004B87`, Laranja Cencosud `#F37021`, fórmulas dinâmicas e dropdowns interativos).

---

## 2. A Nova Regra de Validação: Faltas vs. Data de Desligamento

### 2.1. Fundamento Legal e Operacional
Em rescisões contratuais processadas no sistema de folha (FPW / LG), o colaborador desligado no dia $D$ do mês somente possui, na competência rescisória, $D$ dias corridos até o término do vínculo. Portanto, o número máximo de ausências descontáveis que podem ter ocorrido **exclusivamente dentro do mês da rescisão** é de no máximo $D - 1$ dias (garantindo ao menos a jornada do próprio dia de desligamento).

$$\text{Limite de Ausências Permitidas} = \max(0, D - 1)$$

Onde:
- $D$: Dia do desligamento extraído da `DATA INÍCIO SITUAÇÃO` (ex.: `05/10/2026` $\rightarrow D = 5$).
- $\text{Total de Ausências Descontadas} = \text{Qtd Faltas (5001 / 20010)} + \text{Qtd DSR Faltas (5007 / 20016)}$.

### 2.2. Tratamento de Exceções & Diagnóstico de Negócio
Quando $\text{Total de Ausências} > D - 1$, o sistema dispara automaticamente um **Alerta Crítico** e direciona o colaborador para a aba **`Divergências & Casos Atípicos`**:
- **Causa Operacional Frequente:** Faltas e DSRs ocorridos no mês anterior (ex.: setembro) que não foram fechados a tempo pelo ponto eletrônico e foram importados em atraso na rescisão do mês seguinte (ex.: outubro).
- **Risco Auditado:** Desconto duplicado de faltas ou inconsistência de saldo de salário negativo não homologável perante o MTE/Sindicato.
- **Mensagem Automática:**
  > `ALERTA CRÍTICO: Total de 12 dias de ausência (10 faltas + 2 DSR) excede o limite máximo de 4 dias para desligamento no dia 05.`

---

## 3. Estrutura do Novo Relatório Final em 4 Abas

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                      RELATÓRIO DE AUDITORIA RESCISÓRIA AD003 (V005)                    │
├──────────────────────────┬──────────────────────────┬───────────────────┬──────────────┤
│ 1. Painel & Status       │ 2. Divergências/Atípicos │ 3. Validações     │ 4. Resumo    │
│    Cockpit de Liberação  │    Tratamento Focal      │    Espelho Total  │    TRCT/FPW  │
└──────────────────────────┴──────────────────────────┴───────────────────┴──────────────┘
```

### Aba 1: `Painel & Status` (Cockpit Executivo & Liberação de Pagamentos)
- **Cards de Indicadores (KPIs no Topo):**
  - `RESCISÕES AUDITADAS`: Total de colaboradores processados no lote.
  - `100% HOMOLOGADAS (LIBERADAS)`: Rescisões com parecer "Aprovada" e zero pendências.
  - `BLOQUEADAS / DIVERGÊNCIAS`: Rescisões com inconsistência matemática de cálculo.
  - `ATENÇÃO / CASOS ATÍPICOS`: Rescisões com faltas excedentes, pensão, limite de consignado ou rescisão zerada.
  - `DIVERGÊNCIAS CÁLCULO`: Totalizador de eventos divergentes no lote.
  - `ALERTAS FALTAS > LIMITE`: Quantidade de rescisões com faltas desproporcionais aos dias do mês.
  - `INSUFICIÊNCIA DE SALDO (499)`: Ocorrências de rescisão com saldo negativo zerado.
  - `PENSÕES ALIMENTÍCIAS`: Ocorrências de deduções judiciais.
- **Tabela de Governança por Colaborador:**
  - `Matrícula`, `Colaborador`, `Lotação Tributária`, `Cargo`, `Motivo do Desligamento`, `Data Admissão`, `Data Desligamento`, `Data Pagamento`.
  - `Total Proventos (R$)`, `Total Descontos (R$)`, `Líquido TRCT (R$)`.
  - `Divergências Cálculo`: Fórmula dinâmica vinculada à Aba 2.
  - `Casos Atípicos / Alertas`: Fórmula dinâmica vinculada à Aba 2.
  - `Decisão de Liberação`: **Lista Suspensa interativa (Dropdown)** com as opções `Aguardando`, `Aprovada`, `Reprovada`.
  - `Status Final`: Fórmula automática de governança:
    - Se Decisão = `Reprovada` $\rightarrow$ `REPROVADA` (Vermelho)
    - Se Divergências $> 0$ $\rightarrow$ `BLOQUEADA - Divergência` (Vermelho)
    - Se Casos Atípicos $> 0$ $\rightarrow$ `ATENÇÃO - Caso Atípico` (Amarelo)
    - Se Decisão = `Aprovada` $\rightarrow$ `LIBERADA` (Verde)
    - Se não aprovada $\rightarrow$ `AGUARDANDO APROVAÇÃO` (Cinza)
  - `Responsável Liberação`, `Data Liberação`, `Parecer / Observação`.

---

### Aba 2: `Divergências & Casos Atípicos` (Aba Focal de Auditoria Operacional)
Isola 100% dos eventos que exigem intervenção humana, eliminando a necessidade de o analista auditar milhares de linhas conformes:
1. **Divergências de Cálculo:** Qualquer evento com $| \text{Sistema} - \text{Python} | > \text{R\$}~0,01$.
2. **Alerta de Faltas Excedentes:** Colaboradores com faltas + DSR superiores ao dia de desligamento.
3. **Insuficiência de Saldo (Evento 499):** Rescisões com débitos superiores aos proventos.
4. **Pensão Alimentícia Judicial:** Eventos `5029`, `5703`, etc., exigindo batimento com ofício judicial.
5. **Assiduidade com Falta Injustificada:** Pagamento indevido de prêmio com falta registrada no período.
6. **Consignados acima de 35%:** Excesso de retenção sobre a margem líquida rescisória (Lei 10.820/03).
7. **Falta de Caixa (5045) sem evento PDV correspondente.**
8. **Eventos manuais sem regra de recálculo cadastrada.**

**Colunas e Campos de Workflow:**
- `Linha Origem Validações`, `Matrícula`, `Colaborador`, `Lotação Tributária`, `Cargo`, `Motivo Desligamento`, `Data Desligamento`.
- `Código Evento`, `Descrição Evento`, `Valor Sistema (R$)`, `Valor Python (R$)`, `Diferença (R$)`.
- `Categoria da Ocorrência`, `Prioridade Sugerida` (`Alta`, `Média`, `Baixa`).
- `Motivo / Ponto de Atenção` e `Providência Sugerida`.
- `Conclusão do Analista`: **Dropdown interativo** (`Pendente`, `Ajustado no FPW`, `Justificado`, `Validado Manualmente`, `Não se aplica`).
- `Responsável`, `Data da Análise`, `Evidência / Protocolo`.
- `Status do Tratamento`: Fórmula automática (`PENDENTE` ou `TRATADA`).

---

### Aba 3: `Validações Detalhadas` (Base Analítica Completa)
- Espelho analítico completo de todas as linhas auditadas na folha.
- Cabeçalho institucional Azul Petróleo Cencosud (`#004B87`), congelamento de painéis na linha 1 e formatação condicional:
  - Fundo vermelho suave (`#FFC7CE`): Linhas divergentes.
  - Fundo amarelo suave (`#FFEB9C`): Linhas com alertas e casos atípicos.
  - Fundo verde suave (`#C6EFCE`): Linhas homologadas com 100% de precisão matemática.

---

### Aba 4: `Resumo Financeiro` (Conciliação Contábil / TRCT)
- Resumo estruturado por colaborador contendo os totalizadores oficiais:
  - `10970 - Total de Proventos`
  - `10980 - Total de Descontos`
  - `10990 - Líquido a Receber` (incluindo tratamento de líquidos zerados)
  - `11535 - Saldo para Fins Rescisórios FGTS`
  - `20010 - Qtd Dias Faltas Injustificadas`
  - `20016 - Qtd Dias DSR Faltas`

---

## 4. Simulação de Casos com Dados Sintéticos (LGPD Compliance)

Para fins de documentação, treinamento e auditoria técnica, todas as simulações utilizam dados descaracterizados (`999xxx`):

### Simulação 1: Colaborador com Faltas Excedentes (Alerta de Competência)
- **Matrícula Simulada:** `999010`
- **Colaborador Fictício:** *GHOST EMPLOYEE A*
- **Lotação Tributária:** `CENCOSUD BRASIL COMERCIAL SA - GBARBOSA 015`
- **Data de Desligamento:** `05/10/2026` ($D = 5$)
- **Ausências Lançadas no FPW:**
  - Evento `5001 (Desc Faltas Injust)`: 10 dias (R$ 600,00)
  - Evento `5007 (Desc DSR Faltas)`: 2 dias (R$ 120,00)
  - Total Ausências: 12 dias.
- **Limite Máximo Permitido para o Dia 05:** $5 - 1 = 4\text{ dias}$.
- **Resultado na Aba 1 (`Painel & Status`):**
  - Casos Atípicos / Alertas: `1`
  - Status Final: `ATENÇÃO - Caso Atípico` (Fundo Amarelo)
- **Resultado na Aba 2 (`Divergências & Casos Atípicos`):**
  - Categoria: `Alerta de Faltas vs Data` | Prioridade: `Alta`
  - Motivo: *"ALERTA CRÍTICO: Total de 12 dias de ausência (10 faltas + 2 DSR) excede o limite máximo de 4 dias para desligamento no dia 05."*
  - Providência: *"Verificar espelho de ponto e se as faltas se referem a meses anteriores não processados."*

---

### Simulação 2: Colaborador com Rescisão Negativa (Insuficiência de Saldo 499)
- **Matrícula Simulada:** `999020`
- **Colaborador Fictício:** *GHOST EMPLOYEE B*
- **Lotação Tributária:** `CENCOSUD BRASIL COMERCIAL SA - BRETAS 551`
- **Data de Desligamento:** `02/10/2026`
- **Total Proventos:** R$ 1.500,00 | **Total Descontos Originais:** R$ 2.657,54
- **Insuficiência de Saldo (499):** R$ 1.157,54
- **Líquido TRCT:** R$ 0,00
- **Resultado na Aba 1 (`Painel & Status`):**
  - Casos Atípicos / Alertas: `1`
  - Status Final: `ATENÇÃO - Caso Atípico` (Fundo Amarelo)
- **Resultado na Aba 2 (`Divergências & Casos Atípicos`):**
  - Categoria: `Caso Atípico - Saldo Negativo` | Prioridade: `Alta`
  - Motivo: *"Insuficiência de Saldo (R$ 1.157,54). Rescisão negativa zerada pelo sistema."*
  - Providência: *"Conferir limite de retenção de 1 salário (art. 477 § 5º CLT) e certificar que Salário Família não foi compensado."*

---

### Simulação 3: Colaborador 100% Homologado sem Inconsistências
- **Matrícula Simulada:** `999030`
- **Colaborador Fictício:** *GHOST EMPLOYEE C*
- **Lotação Tributária:** `CENCOSUD BRASIL COMERCIAL SA - PREZUNIC 734`
- **Data de Desligamento:** `14/10/2026`
- **Total Proventos:** R$ 8.540,20 | **Total Descontos:** R$ 940,10 | **Líquido:** R$ 7.600,10
- **Divergências:** `0` | **Casos Atípicos:** `0`
- **Resultado na Aba 1 (`Painel & Status`):**
  - Decisão: Analista marca `Aprovada`.
  - Status Final: `LIBERADA` (Fundo Verde)
  - Aba 2: Sem nenhum apontamento para esta matrícula.

---

## 5. Instruções de Operação e Compilação

### 5.1. Execução via Interface Gráfica (`Auditor_Rescisao_Cencosud.exe`)
1. Abrir o executável na pasta `dist/`.
2. O sistema exibe o badge atualizado: `Versão V005 (Painel & Governança)`.
3. Selecionar o arquivo de folha de rescisão (`.csv`, `.xlsb` ou `.xlsx`).
4. Clicar em **"Processar Auditoria Massiva"**.
5. O relatório gerado conterá automaticamente as 4 abas institucionais estruturadas.

### 5.2. Compilação Standalone via PyInstaller
Para gerar novamente o executável portátil:
```powershell
cd "c:\Users\matcaval\OneDrive - Cencosud\Documentos\matheus\codigos\codigo rescisao"
python -m PyInstaller --noconfirm Auditor_Rescisao_Cencosud.spec
```
O arquivo final gerado em `dist/Auditor_Rescisao_Cencosud.exe` possui ~46.8 MB e é 100% autônomo (não necessita de Python instalado na máquina de destino).
