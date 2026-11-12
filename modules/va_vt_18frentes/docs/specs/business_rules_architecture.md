# Arquitetura e Especificação Detalhada das Regras de Negócio: Auditoria "04 - GOOD CARD" (Cencosud)

**Data do Levantamento**: 06/10/2026  
**Autor**: explorer_survey_rules_3  
**Status**: Concluído / Pronto para Implementação  
**Ambiente de Execução**: Python 3.14.2 | Pandas 3.0.5 | OpenPyXL 3.1.5 | Windows (PowerShell)  
**Diretório do Projeto**: `C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\04 - GOOD CARD`

---

## 1. Visão Geral e Contexto de Negócio

A auditoria mensal do **Good Card** realizada pela equipe de Operações de Benefícios da Cencosud Brasil tem como objetivo auditar a concessão, bloqueio, cancelamento e os descontos em folha de pagamento de todos os colaboradores elegíveis ou portadores do cartão de benefícios especiais (Good Card).

A auditoria concilia as bases de dados extraídas dos sistemas corporativos de RH (Senior/Ronda/Folha) em formato CSV com os arquivos do fornecedor (Ticket/Good Card) e o modelo de homologação `Conferência Good Card_Base.xlsx`. O processamento deve ser 100% em memória (Pandas), sem geração de arquivos temporários em disco, e exportar uma planilha final em Excel formatada, com freeze pane em `'A2'`, destaque visual de divergências em vermelho e **sem nenhuma fórmula de cálculo** (valores estáticos gerados pelo Python).

---

## 2. Especificação Detalhada das 4 Regras de Negócio do Good Card

As regras de negócio foram mapeadas a partir de três fontes oficiais:
1. `Conferência Good Card_Base.xlsx` (abas `Orientações`, `Lista` e fórmulas da aba `Conferência`);
2. `Levantamento conferências da área Opreações de Benefícios.xlsb` (documentação oficial das rotinas de auditoria de benefícios);
3. Arquivos originais em `Relatorio Original/` (`COLA_Colaboradores.CSV`, `COLA_Colaboradores_Afastados.CSV`, `COLA_Demitidos.CSV`, `COLA_Ferias_Colaboradores.CSV`, `FOPA_Eventos_por_Colaborador - *.csv`).

### 2.1 Regra 1: Cartões de Desligados (Cancelamento para Demitidos)

- **Objetivo**: Identificar colaboradores rescindidos/desligados que possuam cartões ativos, saldo devedor pendente ou descontos efetuados indevidamente após a data de rescisão, garantindo o cancelamento imediato junto à operadora Ticket/Good Card e acerto rescisório.
- **Fontes de Dados**:
  - `COLA_Demitidos.CSV` (Campos: `Matrícula`, `Data Rescisão`, `Tipo de Rescisão`, `Motivo do Desligamento`).
  - Base de Folha (`FOPA_Eventos_por_Colaborador`): `Matrícula`, `Total Saldo Devedor`, `6280 - Desc Cartão Benefícios Especiais`, `29621 - Saldo Devedor`.
- **Condições e Filtros Críticos**:
  - **Exceção de Transferência**: Conforme instrução expressa da aba *Orientações*: *"Não deve considerar os colaboradores com tipo de desligamento 'Transferência sem ônus'"*. Colaboradores com `Tipo de Rescisão` contendo `"Transferência sem ônus"` (código 5) não são tratados como rescisão definitiva, devendo ser mantidos ou transferidos de empresa.
  - **Deduplicação**: Se a matrícula possuir múltiplos registros em `COLA_Demitidos` (ex.: recontratações históricas), considerar apenas o registro com a `Data Rescisão` mais recente e pertinente ao ciclo da folha auditada.
- **Classificação e Tratativa**:
  | Condição | Status / Obs. Conferência | Tratativa (Coluna 38) | Severidade / Destaque |
  | :--- | :--- | :--- | :--- |
  | Desligado com saldo devedor liquidado em rescisão (`Total Saldo Devedor > 0`, `6280 == Total Saldo Devedor`, `29621 == 0`) | `Ok - Liquidou todo o saldo devedor ` | `None` (Vazio) | Conforme (Sem destaque) |
  | Desligado sem saldo e sem desconto | `Ok - Sem saldo devedor e sem desconto` | `None` (Vazio) | Conforme (Sem destaque) |
  | Desligado com saldo devedor pendente não quitado (`29621 > 0` ou `Total Saldo Devedor > 6280`) | `Demitido - Saldo devedor pendente` | `Colaborador desligado com saldo devedor não liquidado. Verificar inclusão no TRCT e solicitar cancelamento do cartão.` | **Divergência (Vermelho)** |
  | Desligado com cartão ativo/movimentação após rescisão ou presente no arquivo da operadora | `Demitido - Cartão Ativo` | `Solicitar cancelamento imediato do cartão Good Card junto à operadora Ticket para colaborador desligado.` | **Divergência (Vermelho)** |
  | Desligado por transferência sem ônus com movimentação | `Transferência sem ônus` | `Transferência de empresa: verificar migração do cartão e acerto entre filiais.` | Informativo |

---

### 2.2 Regra 2: Afastados > 15 dias (Solicitação de Bloqueio de Cartão)

- **Objetivo**: Identificar colaboradores com afastamento previdenciário, médico ou acidentário superior a 15 dias (`Dias Afastamento > 15`) para solicitar o bloqueio temporário do cartão Good Card junto à operadora e monitorar o saldo devedor não descontado.
- **Fontes de Dados**:
  - `COLA_Colaboradores_Afastados.CSV` (Campos: `Matrícula`, `Início Afastamento`, `Término Afastamento`, `Dias Afastamento`, `Tipo Afastamento`, `Status Afastamento`).
  - Base de Folha (`FOPA`): `11010 - Salario Mensal`, `6280 - Desc Cartão`, `29621 - Saldo Devedor`, `Total Saldo Devedor`.
- **Condições e Filtros Críticos**:
  - **Afastamentos Finalizados no Passado**: Conforme *Orientações*: *"Não deve considerar os afastamentos que foram finalizados antes da primeira data do mês a que se refere a compra"* (Exemplo: se a folha/compra for Setembro/2026, ignorar afastamentos com `Término Afastamento < 01/09/2026`).
  - **Períodos Múltiplos**: Conforme *Orientações*: *"Verificar se existe colaboradores com mais de um período de afastamento e deve deixar somente o que possui data mais extensa"*. Ordenar por `(Término Afastamento - Início Afastamento)` decrescente e reter o mais abrangente.
  - **Tipos de Afastamento com Duração > 15 dias**:
    - `5 - Motivo de Doença, per. superior a 15 dias`
    - `2 - Por Acidente de Trabalho, per. superior a 15 dias`
    - `18 - Suspensão de Contrato de Trabalho`
    - Ou qualquer registro com `Dias Afastamento > 15` e sem data fim ou fim dentro/após o mês auditado.
- **Classificação e Tratativa**:
  | Condição | Status / Obs. Conferência | Tratativa (Coluna 38) | Severidade / Destaque |
  | :--- | :--- | :--- | :--- |
  | Afastado > 15 dias, sem desconto em folha (`6280 == 0`), com saldo devedor integralmente levado ao evento 29621 (`29621 == Total Saldo Devedor > 0`) | `Afastamento - Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ` | `None` (Vazio) | Conforme esperado em folha |
  | Afastado > 15 dias sem saldo devedor e sem desconto | `Ok - Sem saldo devedor e sem desconto` | `None` (Vazio) | Conforme |
  | Afastado com suspensão de contrato | `Ok - Suspensão de Contrato` | `None` (Vazio) | Conforme |
  | Afastado > 15 dias que sofreu desconto indevido em folha sem margem (`6280 > 0` com salário zerado/afastado) | `Afastado - Desconto indevido em folha` | `Colaborador afastado > 15 dias com desconto em folha. Validar margem líquida e solicitar estorno/regularização.` | **Divergência (Vermelho)** |
  | Afastado > 15 dias ativo com cartão desbloqueado e compras no mês | `Afastado > 15 dias - Cartão não bloqueado` | `Solicitar bloqueio temporário do cartão Good Card junto à operadora Ticket devido a afastamento previdenciário superior a 15 dias.` | **Divergência (Vermelho)** |
  | Afastado com saldo devedor não levado para o evento 29621 (`Total Saldo Devedor > 0`, `6280 == 0`, `29621 == 0`) | `Afastado - Saldo devedor não provisionado` | `Saldo devedor não transferido para o evento 29621 na folha. Acionar equipe de folha para provisionamento.` | **Divergência (Vermelho)** |

---

### 2.3 Regra 3: Novos Admitidos > 90 dias (Verificação de Elegibilidade)

- **Objetivo**: Auditar a concessão do cartão Good Card para colaboradores recém-admitidos que completaram o período de experiência (90 dias de empresa), garantindo a concessão do benefício aos elegíveis e impedindo concessões indevidas a cargos inelegíveis (Aprendizes e Diretores).
- **Fontes de Dados**:
  - `COLA_Colaboradores.CSV` / `FOPA_Eventos_por_Colaborador` (Campos: `Matrícula`, `Data Admissão`, `Código Cargo`, `Cargo`, `Situação`).
  - Base de Folha (`FOPA`): eventos `11010`, `6280`, `29613`, `29621`.
- **Condições e Filtros Críticos**:
  - **Cálculo dos 90 Dias**: `(Data Referência da Folha - Data Admissão).days > 90`. (Por exemplo, para competência 01/09/2026, admitidos até 02/06/2026 completaram mais de 90 dias).
  - **Exclusão Estrita de Aprendizes e Diretores**:
    - Conforme *Orientações*: *"Os colaboradores nos cargos de aprendiz e diretor não devem ter Good Card então deve conferi se existe algum colaborador em um desses cargos"*.
    - **Critério Aprendiz**: `Cargo.str.upper().str.contains("APRENDIZ")` (Ex.: `APRENDIZ REPOSITOR`, `JOVEM APRENDIZ`, `APRENDIZ EM FARMACIA`, `OPERADOR DE CAIXA APRENDIZ`, etc.).
    - **Critério Diretor**: `Cargo.str.upper().str.contains(r"\bDIRETOR\b|\bDIRETORA\b|\bDIRETORIA\b", regex=True)` (cuidado para não pegar `ANALISTA DE IMPOSTOS DIRETOS`).
- **Classificação e Tratativa**:
  | Condição | Status / Obs. Conferência | Tratativa (Coluna 38) | Severidade / Destaque |
  | :--- | :--- | :--- | :--- |
  | Cargo de Aprendiz com qualquer evento/movimentação de Good Card (`Total Saldo Devedor > 0` ou `6280 > 0`) | `Aprendiz - Verificar` | `Concessão indevida de Good Card para cargo de Aprendiz. Solicitar cancelamento do cartão junto à operadora.` | **Divergência (Vermelho)** |
  | Cargo de Diretor com movimentação de Good Card | `Diretor - Verificar` | `Concessão indevida de Good Card para cargo de Diretor. Cargo inelegível conforme política de benefícios.` | **Divergência (Vermelho)** |
  | Colaborador admitido há mais de 90 dias, elegível (não Aprendiz/Diretor), ativo e sem Good Card | `Elegível > 90 dias - Sem cartão` | `Colaborador completou período de experiência (> 90 dias). Solicitar emissão do cartão Good Card junto à operadora.` | Informativo / Oportunidade |
  | Colaborador admitido há menos de 90 dias com desconto em folha | `Admitido < 90 dias com movimentação` | `Movimentação de Good Card antes do término da experiência (90 dias). Verificar política de concessão.` | **Divergência (Vermelho)** |

---

### 2.4 Regra 4: Desconto em Folha e Margem de 20% (Eventos e Saldo Devedor)

- **Objetivo**: Conciliar matematicamente todos os eventos de provento e desconto do Good Card na folha de pagamento, validar se os descontos respeitam rigorosamente a margem consignável interna de 20% do salário base (`11010`), auditar a regra específica de 5% de custo empresa para a bandeira Giga e Estabelecimento 400, e validar a transferência do saldo devedor remanescente para o evento 29621.
- **Eventos Analisados**:
  - `11010`: Salário Mensal (Base de cálculo da margem de 20%).
  - `29622`: Saldo Devedor Benefícios Especiais Mês Anterior (`Bsc Saldo Dev Beneficios Especiais Mes Ant`).
  - `29613`: Valor Informado para Desconto (Compras do mês atual).
  - `6280`: Desconto Cartão Benefícios Especiais (Desconto efetivo realizado na folha).
  - `29621`: Saldo Devedor Benefícios Especiais (Saldo residual transportado para a próxima folha).
  - `29624`: Custo Empresa Benefícios Especiais (Subsídio pago pela empresa, restrito ao Giga e Estab 400).
- **Fórmulas e Relações Matemáticas Oficiais (Colunas R até AB da Conferência)**:
  1. `Total Saldo Devedor` (Col 18):
     $$\text{Total Saldo Devedor} = \text{Evento 29622} + \text{Evento 29613}$$
  2. `Margem 20% do Salário` (Col 19):
     $$\text{Margem 20\%} = \text{Evento 11010} \times 0.20$$
  3. `Total Saldo Devedor Dentro ou Fora da Margem` (Col 20):
     $$\text{Status Margem} = \begin{cases} \text{"Dentro da Margem"}, & \text{se } \text{round}(\text{Total Saldo Devedor}, 2) \le \text{round}(\text{Margem 20\%}, 2) \\ \text{"Fora da Margem"}, & \text{caso contrário} \end{cases}$$
  4. `Desconto x Margem de 20%` (Col 22):
     $$\text{Status Desconto} = \begin{cases} \text{"Dentro da Margem"}, & \text{se } \text{round}(\text{Evento 6280}, 2) \le \text{round}(\text{Margem 20\%}, 2) \\ \text{"Fora da Margem"}, & \text{caso contrário} \end{cases}$$
  5. `Saldo Devedor x Evt 6280` (Col 23) e `Total Saldo Devedor - Evt 6280 x Evt 29621` (Col 24):
     $$\text{Equação de Quitação} = (\text{round}(\text{Evento 6280} + \text{Evento 29624}, 2) == \text{round}(\text{Total Saldo Devedor}, 2))$$
     Ou na presença de saldo residual (29621):
     $$\text{Equação Contábil Completa} = (\text{round}(\text{Evento 6280} + \text{Evento 29621} + \text{Evento 29624}, 2) == \text{round}(\text{Total Saldo Devedor}, 2))$$
  6. `5% Saldo Devedor` (Col 27):
     $$\text{5\% Saldo Devedor} = \text{Total Saldo Devedor} \times 0.05$$
  7. `Cálculo 5% x 29624` (Col 28):
     $$\text{Cálculo 5\% x 29624} = \text{Evento 29624} - \text{5\% Saldo Devedor}$$
     *(Para Estab 400 ou Giga, essa diferença deve ser igual a 0.00 com tolerância de R$ 0.02).*

- **Regra Especial de 5% de Custo Empresa (Giga / Estabelecimento 400)**:
  - *Definição*: No Giga e no Estabelecimento 400 (Cencosud ADM 400), o desconto é fixado com subsídio de 5% da diretoria:
    - Empresa assume 5%: $\text{Evento 29624} = \text{round}(\text{Total Saldo Devedor} \times 0.05, 2)$
    - Colaborador paga 95%: $\text{Evento 6280} = \text{round}(\text{Total Saldo Devedor} \times 0.95, 2)$
    - Soma $\text{Evento 6280} + \text{Evento 29624} = \text{Total Saldo Devedor}$ (liquidação total dentro da margem).

- **Mapeamento de Casos e Rótulos Oficiais da Aba `Lista`**:
  | Cenário Matemático | Status / Obs. Conferência (Coluna 37) | Tratativa (Coluna 38) | Destaque |
  | :--- | :--- | :--- | :--- |
  | `Total Saldo Devedor == 0` e `6280 == 0` | `Ok - Sem saldo devedor e sem desconto` | `None` (Vazio) | Não |
  | `Total Saldo Devedor <= Margem 20%`, `6280 == Total Saldo Devedor`, `29621 == 0`, `29624 == 0` | `Ok - Liquidou todo o saldo devedor ` | `None` (Vazio) | Não |
  | Giga/Estab 400: `Total Saldo Devedor > 0`, `6280 == 95%`, `29624 == 5%`, `29621 == 0` | `Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)` | `None` (Vazio) | Não |
  | `Total Saldo Devedor > Margem 20%`, `6280 <= Margem 20%`, `29621 == Total Saldo Devedor - 6280` | `Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621` | `None` (Vazio) | Não |
  | Colaborador em Férias no mês, sem desconto, saldo levado para 29621 (`29621 == Total Saldo Devedor`) | `Férias- Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ` | `None` (Vazio) | Não |
  | Colaborador Afastado > 15 dias, sem desconto, saldo levado para 29621 | `Afastamento - Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ` | `None` (Vazio) | Não |
  | Colaborador ativo, não afastado, não de férias, com saldo devedor (`Total Saldo Devedor > 0`) e SEM desconto (`6280 == 0`) | `Verificar com quality - Não estava em situação de afastamento e não houve o desconto` | `Verificar com Quality/Folha: Colaborador com saldo devedor ativo e sem desconto em folha, não estando afastado nem em férias.` | **Divergência (Vermelho)** |
  | Desconto em folha superior à margem permitida (`6280 > Margem 20%`) | `Desconto fora da margem dos 20%` | `Desconto no evento 6280 superior à margem consignável de 20% do salário mensal (11010). Ajustar valor descontado.` | **Divergência (Vermelho)** |
  | Quebra da equação contábil (`6280 + 29621 + 29624 != Total Saldo Devedor`) com erro > R$ 0.02 | `Divergência no saldo devedor` | `Soma do desconto (6280) + saldo residual (29621) + custo empresa (29624) diverge do Total Saldo Devedor. Auditar folha.` | **Divergência (Vermelho)** |
  | Giga/Estab 400 com Custo Empresa incorreto (`abs(29624 - 5% * Total Saldo Devedor) > 0.05`) | `Divergência Custo Empresa 5%` | `Valor levado para o custo empresa (29624) diverge dos 5% regulamentares para o estabelecimento 400/Giga.` | **Divergência (Vermelho)** |

---

## 3. Fluxo de Dados e Cruzamento de Informações

### 3.1 Mapeamento de Arquivos Brutos para DataFrames Intermediários

```
                                  [Relatorios CSV Brutos]
                                             │
      ┌─────────────────┬────────────────────┼───────────────────┬─────────────────┐
      │                 │                    │                   │                 │
      ▼                 ▼                    ▼                   ▼                 ▼
[COLA_Colab.CSV] [COLA_Afast.CSV]     [COLA_Demit.CSV]    [COLA_Ferias.CSV] [FOPA_Eventos_*.csv]
  skiprows=1        skiprows=0           skiprows=0          skiprows=1        skiprows=3
  sep=';'           sep=';'              sep=';'             sep=';'           sep=';'
  encoding=latin1   encoding=latin1      encoding=latin1     encoding=latin1   encoding=latin1
      │                 │                    │                   │                 │
      ▼                 ▼                    ▼                   ▼                 ▼
 [df_nomina]       [df_afast]           [df_demit]          [df_ferias]        [df_fopa]
 (Matrícula,       (Matrícula,          (Matrícula,         (Matrícula,        (Matrícula,
  Situação,         Início Afast,        Data Rescisão,      Início Gozo,       Eventos 11010,
  Cargo, etc.)      Término Afast,       Tipo Rescisão)      Fim Gozo,          6280, 29613,
                    Dias, Tipo)                              Dias Gozo)         29621, 29624)
      │                 │                    │                   │                 │
      └─────────────────┴─────────────┬──────┴───────────────────┴─────────────────┘
                                      ▼
                        [Universal Join por Matrícula]
                                      │
                                      ▼
                           [df_conferencia_merged]
                               (38 Colunas)
                                      │
                                      ├───► [Reconciliação Valores Importados] (11 Colunas)
                                      │
                                      └───► [Cálculo de KPIs e Resumo da Auditoria]
```

### 3.2 Chave de Join Universal: Estritamente `Matrícula` como String

Para assegurar 100% de integridade nos merges e evitar problemas comuns de Pandas:
- A `Matrícula` deve ser convertida para string limpa, removendo espaços, caracteres nulos, e eliminando a terminação decimal `.0` comum quando colunas numéricas contêm NaNs:
  ```python
  def normalize_matricula(series: pd.Series) -> pd.Series:
      return (
          series.fillna("")
          .astype(str)
          .str.strip()
          .str.replace(r"\.0$", "", regex=True)
          .str.replace(r"\D+", "", regex=True) # ou preserva texto alfanumérico se houver
      )
  ```
- O join principal é um `LEFT JOIN` partindo do `df_fopa` (que contém todos os colaboradores do ciclo de folha auditado), associando ordenadamente `df_nomina`, `df_afast`, `df_ferias` e `df_demit`.

---

### 3.3 Especificação Exata das 38 Colunas da Aba "Conferência"

A tabela a seguir define a ordem rigorosa (Colunas A até AL), nome oficial do cabeçalho, origem dos dados, tipo, alinhamento, formato numérico e exemplos:

| Nº | Letra | Nome do Cabeçalho Oficial | Origem dos Dados / Regra de Cálculo | Tipo | Alinhamento | Formato Excel | Exemplo de Valor |
| :---: | :---: | :--- | :--- | :---: | :---: | :---: | :--- |
| 1 | A | `Código da Empresa` | `df_fopa['Código da Empresa']` | Int/Str | Centro | `@` / `0` | `2` |
| 2 | B | `Nome da Empresa` | `df_fopa['Nome da Empresa']` | Str | Esquerda | `@` | `CENCOSUD BRASIL COMERCIAL LTDA.` |
| 3 | C | `Código Estabelecimento` | `df_fopa['Código Estabelecimento']` | Int/Str | Centro | `@` / `0` | `176` |
| 4 | D | `Estabelecimento` | `df_fopa['Estabelecimento']` | Str | Esquerda | `@` | `CENCOSUD - ELETROSHOW 176` |
| 5 | E | `Código Centro Custo` | `df_fopa['Código Centro Custo']` | Str | Centro | `@` | `BWG1760439` |
| 6 | F | `Centro Custo` | `df_fopa['Centro Custo']` | Str | Esquerda | `@` | `LOJA-ELETRO-176` |
| 7 | G | `Matrícula` | `df_fopa['Matrícula']` | Str | Centro | `@` | `935` |
| 8 | H | `Nome do Colaborador` | `df_fopa['Nome do Colaborador']` | Str | Esquerda | `@` | `MARCIO AUGUSTO SANTOS GOES` |
| 9 | I | `ID Colaborador` | `df_fopa['ID Colaborador']` | Int/Str | Centro | `@` / `0` | `36261` |
| 10 | J | `Data Admissão` | `df_fopa['Data Admissão']` | Data | Centro | `DD/MM/YYYY` | `22/09/2008` |
| 11 | K | `Código Cargo` | `df_fopa['Código Cargo']` | Int/Str | Centro | `@` / `0` | `2019` |
| 12 | L | `Cargo` | `df_fopa['Cargo']` | Str | Esquerda | `@` | `GERENTE DE ELETRO SHOW` |
| 13 | M | `1 - Salario` | `df_fopa['1 - Salario']` (numérico) | Float | Centro | `R$ #,##0.00` | `5582.44` |
| 14 | N | `11010 - Salario Mensal` | `df_fopa['11010 - Salario Mensal']` | Float | Centro | `R$ #,##0.00` | `5582.44` |
| 15 | O | `10990 - Liquido a Receber` | `df_fopa['10990 - Liquido a Receber']` | Float | Centro | `R$ #,##0.00` | `2944.19` |
| 16 | P | `29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ` | `df_fopa['29622 - Ben Esp - ...']` | Float | Centro | `R$ #,##0.00` | `0.00` |
| 17 | Q | `29613 - Ben Esp - Valor Informado para Desconto` | `df_fopa['29613 - Ben Esp - ...']` | Float | Centro | `R$ #,##0.00` | `317.06` |
| 18 | R | `Total Saldo Devedor` | `Col 16 + Col 17` (`29622 + 29613`) | Float | Centro | `R$ #,##0.00` | `317.06` |
| 19 | S | `Margem 20% do Salário` | `Col 14 * 0.20` (`11010 * 20%`) | Float | Centro | `R$ #,##0.00` | `1116.49` |
| 20 | T | `Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"` | `Total Saldo Devedor <= Margem 20%` | Str | Centro | `@` | `Dentro da Margem` |
| 21 | U | `6280 - Desc Cartão Benefícios Especiais` | `df_fopa['6280 - Desc Cartão ...']` | Float | Centro | `R$ #,##0.00` | `317.06` |
| 22 | V | `Desconto x Margem de 20%` | `Col 21 <= Margem 20%` | Str | Centro | `@` | `Dentro da Margem` |
| 23 | W | `Saldo Devedor x Evt 6280` | `round(Col 26 + Col 21, 2) == round(Col 18, 2)` | Bool | Centro | `@` | `True` |
| 24 | X | `Total Saldo Devedor - Evt 6280 x Evt 29621` | `round(Col 26 + Col 21, 2) == round(Col 18, 2)` | Bool | Centro | `@` | `True` |
| 25 | Y | `29621 - Ben Esp - Saldo Devedor Beneficios Especiais` | `df_fopa['29621 - Ben Esp - ...']` | Float | Centro | `R$ #,##0.00` | `0.00` |
| 26 | Z | `29624 - Ben Esp - Custo Empresa Beneficios Especiais` | `df_fopa['29624 - Ben Esp - ...']` | Float | Centro | `R$ #,##0.00` | `0.00` |
| 27 | AA | `5% Saldo devedor` | `Col 18 * 0.05` | Float | Centro | `R$ #,##0.00` | `15.85` |
| 28 | AB | `Cálculo 5% x 29624` | `Col 26 - Col 27` (`29624 - 5%`) | Float | Centro | `R$ #,##0.00` | `-15.85` |
| 29 | AC | `Início Afastamento` | `df_afast['Início Afastamento']` | Data/Str | Centro | `DD/MM/YYYY` | `04/09/2026` |
| 30 | AD | `Final Afastamento` | `df_afast['Término Afastamento']` | Data/Str | Centro | `DD/MM/YYYY` | `""` / `30/09/2026` |
| 31 | AE | `Tipo do Afastamento` | `df_afast['Tipo Afastamento']` | Str | Esquerda | `@` | `5- Motivo de Doença...` |
| 32 | AF | `Início Férias` | `df_ferias['Início do Período de Gozo']` | Data/Str | Centro | `DD/MM/YYYY` | `01/09/2026` |
| 33 | AG | `Final Férias` | `df_ferias['Fim do Período de Gozo']` | Data/Str | Centro | `DD/MM/YYYY` | `30/09/2026` |
| 34 | AH | `Qtd de dias` | `df_ferias['Dias Gozo']` | Int/Str | Centro | `0` | `30` |
| 35 | AI | `Desligamento` | `df_demit['Data Rescisão']` | Data/Str | Centro | `DD/MM/YYYY` | `18/09/2026` |
| 36 | AJ | `Situação` | `df_nomina['Situação']` | Str | Esquerda | `@` | `Atividade Normal` |
| 37 | AK | `Obs. Conferência` | Classificação da Regra de Negócio | Str | Esquerda | `@` | `Ok - Liquidou todo o saldo devedor ` |
| 38 | AL | `Tratativa` | Texto de Ação Corretiva ou Vazio | Str | Esquerda | `@` | `None` / `Texto da Divergência` |

---

### 3.4 Especificação Exata das 11 Colunas da Aba "Valores Importados"

A aba **Valores Importados** tem a função precípua de auditar o lote enviado pela operadora Ticket/Good Card (ou importado para desconto em folha) contra a folha de pagamento real (evento 29613) e contra os desligados:

| Nº | Letra | Nome do Cabeçalho Oficial | Origem dos Dados / Lógica | Tipo | Alinhamento | Formato Excel | Exemplo |
| :---: | :---: | :--- | :--- | :---: | :---: | :---: | :--- |
| 1 | A | `Emp` | Código da Empresa no arquivo importado | Int/Str | Centro | `0` | `7` |
| 2 | B | `Matrícula` | Matrícula no arquivo importado | Str | Centro | `@` | `305204` |
| 3 | C | `Evento` | Código do evento (geralmente 29613) | Int/Str | Centro | `0` | `29613` |
| 4 | D | `Mês` | Ano/Mês de competência | Int/Str | Centro | `@` / `0` | `92026` |
| 5 | E | `Valor` | Valor informado pelo fornecedor para desconto | Float | Centro | `R$ #,##0.00` | `301.85` |
| 6 | F | `Folha` | Evento 29613 presente na Folha (`df_fopa`) | Float | Centro | `R$ #,##0.00` | `301.85` |
| 7 | G | `Importado x Folha` | `round(Valor, 2) == round(Folha, 2)` | Bool | Centro | `@` | `True` |
| 8 | H | `Desligado` | `Data Rescisão` do `df_demit` | Data/Str | Centro | `DD/MM/YYYY` | `""` / `15/09/2026` |
| 9 | I | `Tipo de Rescisão` | `Tipo de Rescisão` do `df_demit` | Str | Esquerda | `@` | `2 - Sem justa causa` |
| 10 | J | `Obs. Conferência` | Notas (ex.: `Transferencia de empresa...`) | Str | Esquerda | `@` | `None` |
| 11 | K | `Tratativa` | Ação necessária quando `Importado x Folha == False` ou desligado com valor importado | Str | Esquerda | `@` | `None` / `Divergência...` |

---

### 3.5 Especificação da Aba "Resumo da Auditoria" (Executive Dashboard)

Inspirado no modelo corporativo de sucesso da Cencosud (`auditoria_va_vt`), o painel executivo é gerado como a segunda aba do workbook, fornecendo visão estratégica imediata para os gestores de RH:

1. **Cabeçalho Visual (Banner Superior)**:
   - Células mescladas `B2:F2`, Altura 35pt.
   - Preenchimento: Azul Marinho Corporativo Cencosud (`#1F497D`).
   - Fonte: Calibri 14pt, Negrito, Branca (`#FFFFFF`), Alinhamento centralizado.
   - Texto: `PAINEL DE AUDITORIA - CONFERÊNCIA GOOD CARD (BENEFÍCIOS ESPECIAIS)`

2. **Cartões de Indicadores Chave (KPIs)**:
   - `Total de Colaboradores Auditados na Folha`: Contagem de linhas em `Conferência`.
   - `Colaboradores com Movimentação / Saldo Devedor`: Contagem com `Total Saldo Devedor > 0`.
   - `Colaboradores com Desconto Efetivo em Folha (Evento 6280)`: Contagem com `6280 > 0`.
   - `Colaboradores com Saldo Devedor Liquidado`: Contagem com status `Ok - Liquidou todo o saldo devedor`.
   - `Colaboradores com Saldo Devedor Carregado para Mês Seguinte (Evento 29621)`: Contagem com `29621 > 0`.
   - `Colaboradores Estabelecimento 400 / Giga (Custo Empresa 5%)`: Contagem de casos com regra 5%.
   - `Colaboradores Desligados Identificados na Base`: Contagem com `Desligamento` preenchido.
   - `Colaboradores Afastados > 15 dias com Saldo Bloqueado`: Contagem de afastados com saldo em 29621.
   - `Total de Registros Conformes (OK)`: Total auditado menos divergências apontadas.
   - `Total de Inconsistências / Divergências Apontadas`: Contagem de linhas com preenchimento em `Tratativa`.
   - `Índice de Conformidade da Folha`: Percentual calculado $\frac{\text{Conformes}}{\text{Total Auditados}} \times 100\%$.

3. **Reconciliação Financeira Consolidada (Totais em R$)**:
   - `Total Saldo Devedor Apurado`: $\sum \text{Total Saldo Devedor}$
   - `Total Descontado em Folha (Evento 6280)`: $\sum \text{Evento 6280}$
   - `Total Saldo Devedor Residual (Evento 29621)`: $\sum \text{Evento 29621}$
   - `Total Custo Empresa (Evento 29624)`: $\sum \text{Evento 29624}$
   - `Total de Valores Importados (Ticket)`: $\sum \text{Valor (Valores Importados)}$
   - `Divergência Financeira Folha x Importado`: Diferença absoluta entre importado e evento 29613.

4. **Tabela de Distribuição dos Status da Auditoria**:
   - Lista todas as ocorrências da coluna `Obs. Conferência`, com contagem de colaboradores e percentual, destacando em vermelho as linhas anômalas (ex.: `Verificar com quality...`, `Aprendiz - Verificar`, etc.).

---

## 4. Blueprint da Arquitetura do Software

A estrutura de código é dividida em 4 módulos coesos e desacoplados, implementando uma arquitetura limpa em camadas:

```
04 - GOOD CARD/
├── config.py             # Constantes, mapeamento de colunas, paleta de cores, tolerâncias
├── data_loader.py        # Ingestão resiliente, detecção de delimitador/encoding, parsing numérico
├── business_rules.py     # Motor analítico, implementação das regras 1 a 4, conciliação e KPIs
├── excel_exporter.py     # Geração OpenPyXL das 4 abas, estático sem fórmulas, formatação e estilos
├── main.py               # CLI runner, orquestrador do fluxo, telemetria e tratamento de exceções
└── tests/                # Testes unitários e de integração
    ├── test_data_loader.py
    ├── test_business_rules.py
    └── test_excel_exporter.py
```

### 4.1 Módulo `config.py`
- Define:
  - `COLUNAS_OFICIAIS_CONFERENCIA` (as 38 colunas exatas da aba Conferência).
  - `COLUNAS_OFICIAIS_VALORES_IMPORTADOS` (as 11 colunas exatas da aba Valores Importados).
  - `PALETA_CORES`:
    - `HEADER_FILL = "1F497D"` (Azul escuro corporativo)
    - `HEADER_FONT = "FFFFFF"` (Branco bold)
    - `DIVERGENCE_FILL = "FFC7CE"` (Rosa/vermelho suave)
    - `DIVERGENCE_FONT = "9C0006"` (Vermelho escuro bold)
    - `OK_FILL = "C6EFCE"` (Verde suave)
    - `OK_FONT = "006100"` (Verde escuro)
    - `SUB_HEADER_FILL = "DCE6F1"` (Azul gelo suave para seções)
  - `TOLERANCIA_CENTAVOS = 0.02`

### 4.2 Módulo `data_loader.py`
- **Responsabilidades**:
  - `detect_encoding(filepath)`: Detecta encoding com fallback prioritário para `latin1` (padrão absoluto do Senior/Ronda da Cencosud) e `utf-8`.
  - `detect_delimiter(filepath)`: Avalia os primeiros 4KB para identificar `;` vs `,`.
  - `parse_brazilian_currency(series)`: Trata rigorosamente a conversão de strings monetárias brasileiras:
    - Converte `"-"`, `" -"`, `""`, `None`, `np.nan` $\rightarrow 0.00$.
    - Remove pontos separadores de milhar (`.`), substitui vírgula decimal por ponto (`,`), e converte para `float64`.
  - `load_fopa(filepath)`: Ingestão de `FOPA_Eventos_por_Colaborador - *.csv` com `skiprows=3`, delimitador `;`, encoding `latin1`. Converte `Matrícula` para string pura e aplica `parse_brazilian_currency` em todas as colunas de eventos (`11010`, `6280`, `29613`, `29621`, `29622`, `29624`, `10990`, `1 - Salario`).
  - `load_colaboradores(filepath)`: Ingestão de `COLA_Colaboradores.CSV` com `skiprows=1`, `usecols` restrito a `Matrícula`, `Colaborador`, `Data Admissão`, `Cargo`, `Código do Cargo`, `Situação`.
  - `load_afastados(filepath)`: Ingestão de `COLA_Colaboradores_Afastados.CSV` com `skiprows=0`. Filtra afastamentos encerrados antes do mês da competência e deduplica por matrícula mantendo o período de maior duração (`Orientações`).
  - `load_demitidos(filepath)`: Ingestão de `COLA_Demitidos.CSV` com `skiprows=0`. Exclui `Tipo de Rescisão` == `"Transferência sem ônus"` e deduplica mantendo a rescisão mais recente.
  - `load_ferias(filepath)`: Ingestão de `COLA_Ferias_Colaboradores.CSV` com `skiprows=1`. Filtra períodos que abrangem o mês auditado e deduplica.
  - `load_valores_importados(filepath_or_template)`: Carrega o lote da operadora diretamente da aba `Valores Importados` do modelo base (ou arquivo fornecido), extraindo as 5 primeiras colunas (`Emp`, `Matrícula`, `Evento`, `Mês`, `Valor`).
  - **Eficiência de Memória**: O carregamento descarta colunas desnecessárias no momento da leitura (`usecols`), mantendo o consumo de RAM inferior a 400MB durante todo o processo.

### 4.3 Módulo `business_rules.py`
- **Responsabilidades**:
  - `execute_audit(df_fopa, df_colab, df_afast, df_demit, df_ferias, df_vi)`: Função mestre que orquestra os cruzamentos e a aplicação das regras.
  - `merge_data(df_fopa, df_colab, df_afast, df_demit, df_ferias)`: Realiza os `left joins` por `Matrícula`.
  - `apply_financial_calculations(df)`:
    - Calcula `Total Saldo Devedor = Col 16 + Col 17`.
    - Calcula `Margem 20% = Col 14 * 0.20`.
    - Avalia `Total Saldo Devedor Dentro/Fora da Margem`.
    - Avalia `Desconto x Margem 20%`.
    - Calcula equações de quitação e cálculo do subsídio 5% para Estabelecimento 400 / Giga.
  - `apply_rules_and_labeling(df)`:
    - Aplica Regra 1 (Desligados).
    - Aplica Regra 2 (Afastados > 15 dias).
    - Aplica Regra 3 (Admitidos > 90 dias / Aprendizes / Diretores).
    - Aplica Regra 4 (Desconto em Folha, Saldo Devedor, Quitação e 5% Custo Empresa).
    - Preenche as colunas 37 (`Obs. Conferência`) e 38 (`Tratativa`).
  - `reconcile_valores_importados(df_vi, df_conferencia)`:
    - Cruza os registros importados com o evento `29613` da conferência.
    - Avalia `Importado x Folha = (round(Valor, 2) == round(Folha, 2))`.
    - Anexa `Desligado` e `Tipo de Rescisão` do colaborador.
    - Preenche `Obs. Conferência` e `Tratativa` em caso de falta de desconto ou divergência de valor.
  - `compute_kpis(df_conferencia, df_vi)`: Retorna dicionário estruturado com todos os agregados e estatísticas para alimentar a aba `Resumo da Auditoria`.

### 4.4 Módulo `excel_exporter.py`
- **Responsabilidades**:
  - Instancia `openpyxl.Workbook()` e constrói as 4 abas na ordem exata:
    1. `Orientações`: Copia o layout e instruções oficiais da aba base ou grava a matriz padronizada de regras operacionais com títulos, larguras e quebras de linha preservadas.
    2. `Resumo da Auditoria`: Monta o painel executivo com banner navy blue, cartões de KPIs formatados, tabela financeira e matriz de distribuição de status.
    3. `Conferência`: Escreve os cabeçalhos das 38 colunas com estilo corporativo (`#1F497D`), formata dados linha a linha (centralização de códigos/números/datas, alinhamento à esquerda de nomes/tratativas, formato `R$ #,##0.00` em moedas). Aplica freeze pane restrito a `'A2'`. Aplica preenchimento rosa suave (`#FFC7CE`) e fonte vermelha escura (`#9C0006`) nas linhas com divergência (onde `Tratativa` não é vazia).
    4. `Valores Importados`: Escreve as 11 colunas de conciliação, freeze pane em `'A2'`, formatações monetárias e destaque em vermelho onde `Importado x Folha == False` ou `Tratativa` preenchida.
  - **Zero Fórmulas**: Nenhuma célula gerada contém `=VLOOKUP`, `=IF` ou `=SUM`. Todos os valores são estáticos, eliminando tempo de recálculo no Excel e impedindo erros de `#REF!`, `#NAME?` ou `#N/A`.
  - **Ajuste Automático de Largura**: Calcula a largura de cada coluna dinamicamente com base no comprimento do cabeçalho e dos dados (com limite mínimo e margem de segurança).

### 4.5 Módulo `main.py`
- **Responsabilidades**:
  - Ponto de entrada CLI (`argparse`) permitindo parametrização de pastas de entrada, arquivos específicos e pasta de saída.
  - Telemetria de execução: registra timestamp de início, tempo total decorrido, total de registros processados por fonte, contagem de divergências encontradas e caminho absoluto do arquivo `.xlsx` gerado.
  - Tratamento global de exceções: captura falhas de I/O, encoding ou tipagem, emitindo logs estruturados com rastreabilidade completa.
  - Garante processamento 100% em memória (zero criação de CSVs intermediários).

---

## 5. Casos de Borda, Vulnerabilidades Identificadas e Estratégia de Mitigação

| Vulnerabilidade / Caso de Borda | Causa Raiz Observada nos Dados | Estratégia de Mitigação Técnica no Código |
| :--- | :--- | :--- |
| **Hífen representando zero (`" -"`)** | O gerador de relatórios do Senior emite `" -"` em colunas de eventos sem lançamento. | A função `parse_brazilian_currency` trata expressamente strings contendo `"-"` ou `" -"` como `0.0`. |
| **Separadores numéricos brasileiros (`5.582,44`)** | Padrão PT-BR com ponto para milhar e vírgula para decimal. | Remoção de `.` e substituição de `,` por `.` antes da conversão para `float64`. |
| **Metadados no topo dos CSVs** | Arquivos como FOPA possuem 3 linhas de metadados; Férias e Colaboradores possuem 1 linha. | `skiprows` parametrizado por tipo de arquivo no `DataLoader`. |
| **Perda de zeros à esquerda ou `.0` na Matrícula** | Leitura padrão do Pandas como inteiro ou float quando há células nulas. | Forçar `dtype=str` na leitura e aplicar `normalize_matricula` para remover pontuações e sufixos `.0`. |
| **Múltiplos afastamentos para mesma matrícula** | Colaboradores com histórico de vários afastamentos na base. | Deduplicação retendo o período de maior duração e filtrando términos anteriores ao início do mês (`Orientações`). |
| **Férias bi-partidas entre dois meses** | Colaborador com férias em agosto e setembro no mesmo relatório. | Filtragem para manter apenas os períodos de gozo cujo intervalo compreenda o mês da competência auditada. |
| **Recontratações em Demitidos** | Colaborador demitido no passado, recontratado e ativo na folha atual. | Validar se a data de rescisão é recente e coincide com o desligamento da admissão atual; caso contrário, ignorar histórico antigo. |
| **Regra 5% Giga / Estab 400** | Estabelecimento 400 e lojas Giga possuem critério de desconto subsidiado de 5%. | Condicional por código de empresa/estabelecimento (`Estabelecimento == 400` ou rede Giga) aplicando divisão 95% colaborador / 5% custo empresa. |
| **Diferença de arredondamento de centavos** | Operações de ponto flutuante em percentuais (`20%`, `5%`, `95%`). | Uso sistemático de `round(valor, 2)` e tolerância de R$ 0.02 na comparação contábil. |

---

## 6. Estratégia Abrangente de Testes

1. **Testes Unitários (`tests/test_business_rules.py`)**:
   - Teste da Regra 1: Validação de cancelamento para demitido com saldo devedor; validação de desconsideração de "Transferência sem ônus".
   - Teste da Regra 2: Validação de bloqueio para afastamento > 15 dias; validação de afastamento com data fim no passado.
   - Teste da Regra 3: Validação de colaboradores com > 90 dias; detecção de cargo de Aprendiz com benefício indevido; detecção de Diretor.
   - Teste da Regra 4: Validação de margem de 20%; cálculo de saldo devedor e residual 29621; validação da regra de 5% para o estabelecimento 400.
2. **Testes de Ingestão (`tests/test_data_loader.py`)**:
   - Teste com dados ruidosos (hífens em valores, espaços extras, acentuação latin1).
   - Teste de preservação da matrícula como string íntegra.
3. **Testes de Exportação Excel (`tests/test_excel_exporter.py`)**:
   - Validação da existência das 4 abas com os nomes exatos.
   - Validação de 38 colunas na aba `Conferência` e 11 na aba `Valores Importados`.
   - Verificação de freeze pane em `'A2'` em todas as abas.
   - Verificação de ausência de fórmulas (`cell.value` nunca se inicia com `'='`).
   - Validação de preenchimento vermelho nas linhas de divergência.
4. **Teste de Integração End-to-End (`tests/test_e2e_integration.py`)**:
   - Execução do fluxo completo a partir dos arquivos reais em `Relatorio Original/`.
   - Comparação linha a linha dos totais financeiros gerados contra os totais da base de homologação `Conferência Good Card_Base.xlsx`.
