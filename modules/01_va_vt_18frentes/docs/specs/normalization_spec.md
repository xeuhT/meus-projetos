# Especificação Técnica de Normalização, Deduplicação e Pré-Filtros (Milestone 1)
**Auditoria de Benefícios "04 - GOOD CARD" — Cencosud Brasil**  
**Data da Especificação:** 07/10/2026  
**Autor:** teamwork_preview_explorer (`explorer_m1_3`)  
**Módulo Alvo:** `data_loader.py`  
**Ambiente de Execução:** Python 3.14.2 | Pandas 3.0.5 | OpenPyXL 3.1.5 | Windows (PowerShell)  
**Diretório do Projeto:** `C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\04 - GOOD CARD`

---

## 1. Visão Geral e Contexto Arquitetural

No fluxo do projeto **04 - GOOD CARD**, o módulo `data_loader.py` é responsável pela ingestão resiliente, saneamento sintático, tipagem uniforme e deduplicação de negócio dos 5 relatórios CSV corporativos de Recursos Humanos e da base de conciliação do fornecedor Ticket (`Valores Importados`).

```
[Relatório Original / CSVs de RH]        [Conferência Good Card_Base.xlsx]
          │                                              │
          ▼                                              ▼
┌────────────────────────────────────────────────────────────────────────┐
│                             data_loader.py                             │
│                                                                        │
│  1. Ingestão Resiliente (sep=';', encoding='latin1', skiprows)         │
│  2. Universal Matrícula Normalization (clean string, sem perda de 0s)   │
│  3. Parsing Monetário Brasileiro (5.582,44 -> 5582.44; '-' -> 0.00)    │
│  4. Deduplicação e Pré-Filtro Demitidos (exclui sem ônus, rescisão max)│
│  5. Deduplicação e Pré-Filtro Afastados (fim >= 01/09, maior duração)  │
│  6. Deduplicação e Pré-Filtro Férias (overlap set/26, férias ativas)   │
│  7. Ingestão Dual de Valores Importados (aba Excel ou CSV externo)     │
└────────────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
                   dict[str, pd.DataFrame] em Memória
             (Pronto para joins e business_rules.py)
```

### Princípios Arquiteturais Inegociáveis
1. **Processamento 100% em Memória**: Nenhuma etapa intermediária grava arquivos em disco. Todos os tratamentos ocorrem em objetos Pandas limpos.
2. **Preservação de Dados Brutos**: Os arquivos originais em `Relatorio Original/` são estritamente de leitura (`read-only`).
3. **Integridade Referencial Chave**: A coluna `Matrícula` é a chave única e universal de relacionamento entre todos os DataFrames. Nenhuma inconsistência de tipagem (ex.: `int` cruzando com `str`, espaços marginais ou sufixos `.0`) pode comprometer os joins.

---

## 2. Normalização Universal da `Matrícula`

### 2.1 Diagnóstico do Problema
Nos relatórios corporativos da Cencosud (Sistemas Senior/Ronda/Folha) e nas planilhas Excel manipuladas pela operação, a chave de colaborador apresenta múltiplos riscos de corrupção:
- **Conversão Automática para Float**: Quando uma coluna de IDs contém células vazias (`NaN` / `None`), o Pandas converte a coluna para `float64`, adicionando sufixos `.0` (ex: `305204` vira `305204.0`).
- **Perda de Zeros à Esquerda**: Castings diretos para inteiro (`int()`) eliminam zeros à esquerda caso existam matrículas formatadas com zeros prévios (ex: `'001234'` viraria `1234`).
- **Espaços Marginais**: Strings brutas frequentemente contêm espaços no início ou fim (`' 305204 '`).
- **Colunas Espúrias Homônimas**:
  - Em `COLA_Colaboradores.CSV`: existe `Matrícula eSocial` (Índice 2, com caracteres compostos como `3-289-00277064`) e `Matrícula` (Índice 3). O join **deve usar estritamente o Índice 3 (`Matrícula`)**.
  - Em `COLA_Demitidos.CSV`: existe `Matrícula` (Índice 5) e `Matrícula do Gestor` (Índice 38, contendo pipes como `36172 | 330516`). O join **deve usar estritamente o Índice 5 (`Matrícula`)**.

### 2.2 Mapeamento das Colunas de Matrícula por Arquivo

| Arquivo Fonte | Índice da Coluna | Nome Exato do Cabeçalho | skiprows | Exemplo Bruto |
| :--- | :---: | :--- | :---: | :--- |
| `FOPA_Eventos_por_Colaborador - *.csv` | 6 (7ª col) | `Matrícula` | 3 | `'935'`, `'331372'` |
| `COLA_Colaboradores.CSV` | 3 (4ª col) | `Matrícula` | 1 | `'935'`, `'16038'` |
| `COLA_Demitidos.CSV` | 5 (6ª col) | `Matrícula` | 0 | `'78'`, `'324474'` |
| `COLA_Colaboradores_Afastados.CSV` | 4 (5ª col) | `Matrícula` | 0 | `'11388'`, `'332560'` |
| `COLA_Ferias_Colaboradores.CSV` | 2 (3ª col) | `Matrícula` | 1 | `'20357'`, `'207813'` |
| `Conferência Good Card_Base.xlsx` (Valores Importados) | 1 (Col B) | `Matrícula` | 0 | `305204` (numérico) |

### 2.3 Especificação do Algoritmo Canônico

A função de normalização é implementada em duas variantes complementares: elemento a elemento (`normalize_matricula`) e vetorizada para Series Pandas (`normalize_matricula_series`).

```python
import re
from typing import Any
import pandas as pd
import numpy as np

def normalize_matricula(val: Any) -> str:
    """
    Normaliza uma matrícula individual garantindo:
    1. Retorno como string limpa.
    2. Eliminação de sufixos decimais (.0, .00).
    3. Remoção de espaços em branco marginais.
    4. Preservação de zeros à esquerda se existirem.
    5. Conversão de nulos/vazios/traços em string vazia ("").
    """
    if val is None or pd.isna(val):
        return ""
    
    # Se for tipo numérico nativo (float ou int)
    if isinstance(val, float):
        if np.isnan(val):
            return ""
        if val.is_integer():
            return str(int(val))
        return str(val)
    elif isinstance(val, int):
        return str(val)
    
    # Se for string ou objeto
    s = str(val).strip()
    
    # Sentinelas nulos ou ausentes
    if s in ("", "-", " -", "nan", "NaN", "None", "<NA>", "null", "NULL"):
        return ""
    
    # Remoção de terminação .0 ou .00 causada por conversões prévias
    if s.endswith(".0"):
        s = s[:-2]
    elif re.search(r"\.0+$", s):
        s = re.sub(r"\.0+$", "", s)
    
    return s.strip()

def normalize_matricula_series(series: pd.Series) -> pd.Series:
    """
    Aplica a normalização canônica em toda a Série do Pandas de forma robusta e otimizada.
    """
    return series.apply(normalize_matricula)
```

### 2.4 Matriz de Casos de Borda de Normalização

| Entrada Bruta (`val`) | Tipo de Dado Python | Saída Esperada (`normalize_matricula`) | Justificativa |
| :--- | :---: | :---: | :--- |
| `305204` | `int` | `'305204'` | Inteiro nativo vira string direta |
| `305204.0` | `float` | `'305204'` | Float com parte decimal zero elimina `.0` |
| `'305204'` | `str` | `'305204'` | String numérica mantida |
| `'  305204  '` | `str` | `'305204'` | Espaços laterais removidos via `.strip()` |
| `'305204.0'` | `str` | `'305204'` | String terminada em `.0` tem sufixo podado |
| `'305204.00'` | `str` | `'305204'` | String terminada em múltiplos zeros podada |
| `'001234'` | `str` | `'001234'` | **Zeros à esquerda preservados integralmente** |
| `'001234.0'` | `str` | `'001234'` | Sufixo `.0` removido mantendo zeros à esquerda |
| `None` | `NoneType` | `''` | Nulo vira string vazia |
| `np.nan` | `float` | `''` | NaN vira string vazia |
| `pd.NA` | `NAType` | `''` | NA vira string vazia |
| `'-'` | `str` | `''` | Traço indicativo de ausência vira vazio |
| `'78'` | `str` | `'78'` | Matrícula curta de 2 dígitos preservada |

---

## 3. Deduplicação e Pré-Filtro: `COLA_Demitidos`

### 3.1 Diagnóstico do Arquivo
- **Arquivo**: `Relatorio Original/COLA_Demitidos.CSV`
- **Total de Linhas Brutas**: 16.044 registros de dados
- **Duplicidades Encontradas**: 4 matrículas duplicadas (8 linhas no total: `324474`, `325212`, `327987`, `333635`).
- **Registros de Transferência**: 3.140 linhas possuem `Tipo de Rescisão` contendo `"Transferência sem ônus"` (código de desligamento interno entre filiais sem rescisão de vínculo com o grupo).

### 3.2 Regra de Negócio Oficial (Aba `Orientações`)
> *"Não deve considerar os colaboradores com tipo de desligamento 'Transferência sem ônus'."*  
> *"Caso haja múltiplos desligamentos para a mesma matrícula, manter o mais recente."*

### 3.3 Especificação do Algoritmo

1. **Leitura com Tipagem**: Ler com `skiprows=0`, `sep=';'`, `encoding='latin1'`, `dtype=str`.
2. **Normalização de Cabeçalhos**: `.str.strip()` em todos os nomes de colunas.
3. **Normalização da Matrícula**: Aplicar `normalize_matricula_series` na coluna `Matrícula`.
4. **Pré-Filtro de Exclusão (Transferência sem ônus)**:
   - Excluir qualquer linha onde `Tipo de Rescisão` contenha o texto `"Transferência sem ônus"` (busca case-insensitive ou comparação de prefixo `"5 -"`).
   - Registros restantes: exatamente **12.904 linhas**.
5. **Deduplicação Temporal (Rescisão Mais Recente)**:
   - Converter `Data Rescisão` para `pd.to_datetime` com formato estrito `'%d/%m/%Y'`.
   - Ordenar o DataFrame por data de rescisão decrescente (`ascending=False`), com `na_position='last'`.
   - Executar `drop_duplicates(subset=['Matrícula'], keep='first')`.
   - Registros únicos finais: exatamente **12.901 colaboradores**.

```python
def load_and_deduplicate_demitidos(filepath: str) -> pd.DataFrame:
    """
    Ingere e processa COLA_Demitidos.CSV:
    - Exclui 'Transferência sem ônus'
    - Ordena por Data Rescisão decrescente
    - Rete a rescisão mais recente por Matrícula
    """
    df = pd.read_csv(filepath, sep=';', encoding='latin1', dtype=str)
    df.columns = [c.strip() for c in df.columns]
    
    # 1. Normalizar Matrícula
    df['Matrícula'] = normalize_matricula_series(df['Matrícula'])
    df = df[df['Matrícula'] != ""].copy()
    
    # 2. Pré-filtro: Excluir 'Transferência sem ônus'
    mask_transferencia = df['Tipo de Rescisão'].fillna('').str.contains(
        'Transferência sem ônus', case=False, regex=False
    )
    df_filtered = df[~mask_transferencia].copy()
    
    # 3. Parsing de data para ordenação cronológica
    df_filtered['_dt_rescisao'] = pd.to_datetime(
        df_filtered['Data Rescisão'], format='%d/%m/%Y', errors='coerce'
    )
    
    # 4. Ordenação decrescente (mais recente primeiro) e deduplicação
    df_sorted = df_filtered.sort_values(
        by='_dt_rescisao', ascending=False, na_position='last'
    )
    df_dedup = df_sorted.drop_duplicates(subset=['Matrícula'], keep='first').copy()
    
    # 5. Limpeza de colunas temporárias
    df_dedup = df_dedup.drop(columns=['_dt_rescisao'])
    
    return df_dedup
```

### 3.4 Rastreamento dos Casos Reais de Duplicidade

| Matrícula | Ocorrências no Bruto | Datas de Rescisão | Tipos de Rescisão | Tratativa Aplicada | Linha Selecionada |
| :---: | :---: | :---: | :--- | :--- | :---: |
| `324474` | 2 | 01/08/2025<br>11/03/2025 | `5 - Transferência sem ônus`<br>`2 - Sem justa causa do empregador` | Linha de transferência eliminada pelo pré-filtro; mantida a rescisão real. | 11/03/2025 |
| `325212` | 2 | 14/01/2026<br>02/01/2025 | `1 - Com justa causa do empregador`<br>`4 - Sem justa causa do empregado` | Ambas rescisões válidas; ordenada por data decrescente. | **14/01/2026** |
| `327987` | 2 | 01/09/2026<br>01/07/2025 | `2 - Sem justa causa do empregador`<br>`2 - Sem justa causa do empregador` | Recontratação; ordenada por data decrescente. | **01/09/2026** |
| `333635` | 2 | 05/08/2026<br>18/02/2026 | `2 - Sem justa causa do empregador`<br>`17 - Término do contrato...` | Recontratação; ordenada por data decrescente. | **05/08/2026** |

---

## 4. Deduplicação e Pré-Filtro: `COLA_Colaboradores_Afastados`

### 4.1 Diagnóstico do Arquivo
- **Arquivo**: `Relatorio Original/COLA_Colaboradores_Afastados.CSV`
- **Total de Linhas Brutas**: 1.867 registros de dados
- **Comportamento das Datas**:
  - `Início Afastamento`: 100% preenchido (1.867 registros), formato `DD/MM/YYYY`.
  - `Término Afastamento`: 137 preenchidos (todos $\ge$ 01/09/2026), 1.730 vazios (`NaN` / afastamentos contínuos sem data de término definida).
  - `Dias Afastamento`: 100% preenchido (mínimo: 2 dias, média: 1.195 dias, máximo: 12.569 dias).

### 4.2 Regra de Negócio Oficial (Aba `Orientações`)
> *"Não deve considerar os afastamentos que foram finalizados antes da primeira data do mês a que se refere a compra."* (Ex: para setembro/2026, ignorar se `Término Afastamento < 01/09/2026`).  
> *"Verificar se existe colaboradores com mais de um período de afastamento e deve deixar somente o que possui data mais extensa."*

### 4.3 Especificação do Algoritmo

1. **Parâmetros de Auditoria**: `ref_month_start: pd.Timestamp = pd.Timestamp("2026-09-01")`.
2. **Leitura com Tipagem**: `skiprows=0`, `sep=';'`, `encoding='latin1'`, `dtype=str`.
3. **Normalização da Matrícula**: Aplicar `normalize_matricula_series`.
4. **Parsing de Datas**:
   - `_dt_inicio = pd.to_datetime(df['Início Afastamento'], format='%d/%m/%Y', errors='coerce')`
   - `_dt_fim = pd.to_datetime(df['Término Afastamento'], format='%d/%m/%Y', errors='coerce')`
5. **Pré-Filtro de Afastamentos Encerrados**:
   - Manter registros onde `_dt_fim` é nulo (`NaT` / contínuo) OU `_dt_fim >= ref_month_start`.
   - Se `_dt_fim < ref_month_start`, o afastamento já cessou antes do mês auditado e deve ser descartado.
6. **Cálculo da Duração Efetiva**:
   - Períodos fechados: `(dt_fim - dt_inicio).dt.days`.
   - Períodos abertos (`dt_fim` nulo): o colaborador continua afastado. Utiliza-se o valor contido na coluna `Dias Afastamento` (numérico) ou atribui-se prioridade máxima de afastamento ativo (`_is_open = True`).
7. **Ordenação e Deduplicação**:
   - Ordenar por:
     1. `_is_open` decrescente (afastamentos abertos têm prioridade sobre encerrados).
     2. `_duration` decrescente (maior quantidade de dias de afastamento).
     3. `_dt_inicio` decrescente (início mais recente para desempate).
   - Executar `drop_duplicates(subset=['Matrícula'], keep='first')`.

```python
def load_and_deduplicate_afastados(
    filepath: str, ref_month_start: str = "2026-09-01"
) -> pd.DataFrame:
    """
    Ingere e processa COLA_Colaboradores_Afastados.CSV:
    - Filtra afastamentos encerrados antes do início do mês de referência
    - Calcula duração de cada afastamento
    - Rete o período mais longo/abrangente por Matrícula
    """
    df = pd.read_csv(filepath, sep=';', encoding='latin1', dtype=str)
    df.columns = [c.strip() for c in df.columns]
    
    df['Matrícula'] = normalize_matricula_series(df['Matrícula'])
    df = df[df['Matrícula'] != ""].copy()
    
    ref_start = pd.to_datetime(ref_month_start)
    
    dt_inicio = pd.to_datetime(df['Início Afastamento'], format='%d/%m/%Y', errors='coerce')
    dt_fim = pd.to_datetime(df['Término Afastamento'], format='%d/%m/%Y', errors='coerce')
    
    # Pré-filtro: manter abertos (NaT) ou término >= início da competência
    mask_vigente = dt_fim.isna() | (dt_fim >= ref_start)
    df_filtered = df[mask_vigente].copy()
    
    # Cálculo de duração
    dias_col = pd.to_numeric(df_filtered['Dias Afastamento'], errors='coerce')
    calc_days = (dt_fim.loc[df_filtered.index] - dt_inicio.loc[df_filtered.index]).dt.days
    
    df_filtered['_duration'] = calc_days.fillna(dias_col).fillna(0)
    df_filtered['_is_open'] = dt_fim.loc[df_filtered.index].isna()
    df_filtered['_dt_inicio'] = dt_inicio.loc[df_filtered.index]
    
    # Ordenar por: aberto primeiro, maior duração, início mais recente
    df_sorted = df_filtered.sort_values(
        by=['_is_open', '_duration', '_dt_inicio'],
        ascending=[False, False, False]
    )
    
    df_dedup = df_sorted.drop_duplicates(subset=['Matrícula'], keep='first').copy()
    df_dedup = df_dedup.drop(columns=['_duration', '_is_open', '_dt_inicio'])
    
    return df_dedup
```

---

## 5. Deduplicação e Pré-Filtro: `COLA_Ferias_Colaboradores`

### 5.1 Diagnóstico do Arquivo
- **Arquivo**: `Relatorio Original/COLA_Ferias_Colaboradores.CSV`
- **Total de Linhas no Arquivo**: 2.143 linhas $\rightarrow$ 2.141 de dados após `skiprows=1`
- **Duplicidades Encontradas**: 8 matrículas duplicadas (16 linhas no total) decorrentes de férias fracionadas em múltiplos períodos (ex: 15 dias em agosto e 15 dias em setembro).
- **Valores da Coluna `Status`**:
  - `'Gozada'`: 1.089 ocorrências (férias já usufruídas/encerradas).
  - `'Em férias'`: 1.040 ocorrências (colaborador atualmente em período de férias).
  - `'Programada'`: 12 ocorrências (férias futuras programadas).

### 5.2 Regra de Negócio Oficial (Aba `Orientações`)
> *"Manter somente o período que compreende o mês de referência da auditoria/compra."*  
> *"Caso haja múltiplos períodos que interceptem o mês, reter o período ativo ('Em férias')."*

### 5.3 Especificação do Algoritmo

1. **Parâmetros de Auditoria**:
   - `ref_month_start = pd.Timestamp("2026-09-01")`
   - `ref_month_end = pd.Timestamp("2026-09-30")`
2. **Leitura com Tipagem**: `skiprows=1`, `sep=';'`, `encoding='latin1'`, `dtype=str`.
3. **Normalização da Matrícula**: Aplicar `normalize_matricula_series`.
4. **Parsing de Datas de Gozo**:
   - `dt_inicio = pd.to_datetime(df['Início do Período de Gozo'], format='%d/%m/%Y', errors='coerce')`
   - `dt_fim = pd.to_datetime(df['Fim do Período de Gozo'], format='%d/%m/%Y', errors='coerce')`
5. **Pré-Filtro de Interseção Temporal (Sobreposição com Setembro/2026)**:
   - Uma faixa de datas $[I, F]$ tem sobreposição com $[M_{\text{ini}}, M_{\text{fim}}]$ se, e somente se:
     $$I \le M_{\text{fim}} \quad \text{E} \quad F \ge M_{\text{ini}}$$
   - Linhas que atendem à sobreposição: **2.007 linhas** (134 linhas descartadas por serem de outros períodos sem sobreposição).
6. **Deduplicação de Período Ativo**:
   - Entre as 2.007 linhas com sobreposição, 2 matrículas (`207813` e `318937`) possuem dois períodos fracionados dentro de setembro.
   - Critério de prioridade:
     1. Priorizar o status ativo: `is_em_ferias = df['Status'].str.lower().str.contains('em f')`.
     2. Início de gozo mais recente: `dt_inicio` decrescente.
     3. Quantidade de dias: `Dias Gozo` numérico decrescente.
   - Executar `drop_duplicates(subset=['Matrícula'], keep='first')`.
   - Registros únicos finais: exatamente **2.005 colaboradores**.

```python
def load_and_deduplicate_ferias(
    filepath: str,
    ref_month_start: str = "2026-09-01",
    ref_month_end: str = "2026-09-30"
) -> pd.DataFrame:
    """
    Ingere e processa COLA_Ferias_Colaboradores.CSV:
    - Pula a linha 1 de título (skiprows=1)
    - Filtra períodos que interceptam o mês de referência
    - Prioriza o período com status 'Em férias' e início mais recente
    """
    df = pd.read_csv(filepath, sep=';', encoding='latin1', skiprows=1, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    
    df['Matrícula'] = normalize_matricula_series(df['Matrícula'])
    df = df[df['Matrícula'] != ""].copy()
    
    ref_start = pd.to_datetime(ref_month_start)
    ref_end = pd.to_datetime(ref_month_end)
    
    dt_inicio = pd.to_datetime(df['Início do Período de Gozo'], format='%d/%m/%Y', errors='coerce')
    dt_fim = pd.to_datetime(df['Fim do Período de Gozo'], format='%d/%m/%Y', errors='coerce')
    
    # 1. Filtro de sobreposição
    mask_overlap = (dt_inicio <= ref_end) & (dt_fim >= ref_start)
    df_filtered = df[mask_overlap].copy()
    
    # 2. Indicadores de ordenação
    df_filtered['_is_active'] = df_filtered['Status'].fillna('').str.lower().str.contains('em f', regex=False)
    df_filtered['_dt_inicio'] = dt_inicio.loc[df_filtered.index]
    df_filtered['_dias'] = pd.to_numeric(df_filtered['Dias Gozo'], errors='coerce').fillna(0)
    
    # 3. Ordenação e deduplicação
    df_sorted = df_filtered.sort_values(
        by=['_is_active', '_dt_inicio', '_dias'],
        ascending=[False, False, False]
    )
    
    df_dedup = df_sorted.drop_duplicates(subset=['Matrícula'], keep='first').copy()
    df_dedup = df_dedup.drop(columns=['_is_active', '_dt_inicio', '_dias'])
    
    return df_dedup
```

### 5.4 Validação Forense Contra a Base Mestre (`Conferência Good Card_Base.xlsx`)

O comportamento do algoritmo coincide 100% com os registros mantidos no modelo mestre:

| Matrícula | Período 1 | Período 2 | Status no Bruto | Seleção do Algoritmo | Registro no Modelo Base | Conformidade |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| `207813` | 19/08 a 02/09 (15d) | 16/09 a 30/09 (15d) | `Gozada`<br>`Em férias` | **16/09 a 30/09 (`Em férias`)** | 16/09 a 30/09 (`Em férias`) | **100% Exato** |
| `318937` | 17/08 a 05/09 (20d) | 14/09 a 23/09 (10d) | `Gozada`<br>`Em férias` | **14/09 a 23/09 (`Em férias`)** | 14/09 a 23/09 (`Em férias`) | **100% Exato** |
| `20357` | 03/08 a 17/08 (15d) | 01/09 a 15/09 (15d) | `Gozada`<br>`Gozada` | **01/09 a 15/09** (Período 1 fora) | 01/09 a 15/09 | **100% Exato** |

---

## 6. Parsing Monetário Brasileiro: `parse_brazilian_currency`

### 6.1 Problema Identificado
Valores monetários podem ter origens distintas:
- CSVs com strings formatadas em padrão brasileiro: `'5.582,44'`, `'317,06'`.
- CSVs com traço indicando zero: `"-"`, `" -"`.
- Excel pré-carregado: valores já são floats nativos (`301.85`). Uma conversão ingênua via `str(val).replace('.', '')` transformaria `301.85` em `30185.0` (multiplicação espúria por 100).

### 6.2 Implementação Oficial

```python
def parse_brazilian_currency(val: Any) -> float:
    """
    Converte qualquer representação monetária para float64 com 2 casas decimais.
    - None / NaN / '-' / ' -' -> 0.0
    - float / int -> round(val, 2)
    - '5.582,44' -> 5582.44
    - '301.85' -> 301.85
    """
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return round(float(val), 2)
    s = str(val).strip()
    if s in ("", "-", " -", "nan", "NaN", "None", "<NA>", "0,00", "0.00"):
        return 0.0
    # Se contém vírgula, segue o padrão brasileiro (pontos de milhar, vírgula decimal)
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    return round(float(s), 2)
```

---

## 7. Especificação de Ingestão de `load_valores_importados()`

### 7.1 Suporte a Múltiplas Fontes (Template Base vs CSV Externo)
A função deve suportar tanto a carga a partir do modelo mestre (`Conferência Good Card_Base.xlsx`, aba `Valores Importados`) quanto de um arquivo CSV externo fornecido pela operadora Ticket/Good Card (ex.: `valores_importados.csv`).

### 7.2 Implementação Oficial

```python
import os
from typing import Union
from pathlib import Path
import openpyxl
import pandas as pd

def load_valores_importados(
    source_path: Union[str, Path],
    sheet_name: str = "Valores Importados"
) -> pd.DataFrame:
    """
    Carrega o lote de valores importados da operadora:
    - Se arquivo .xlsx/.xlsm: lê a aba 'Valores Importados' via openpyxl (read_only rápido).
    - Se arquivo .csv: lê com detecção de delimitador e encoding latin1.
    - Normaliza Matrícula e converte Valor para float.
    - Retorna DataFrame padronizado com as colunas base:
      ['Emp', 'Matrícula', 'Evento', 'Mês', 'Valor', 'Obs. Conferência']
    """
    path_str = str(source_path)
    if not os.path.exists(path_str):
        raise FileNotFoundError(f"Arquivo de valores importados não encontrado: {path_str}")
    
    ext = os.path.splitext(path_str)[1].lower()
    
    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(path_str, read_only=True, data_only=True)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Aba '{sheet_name}' não encontrada no arquivo {path_str}")
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return pd.DataFrame(columns=["Emp", "Matrícula", "Evento", "Mês", "Valor", "Obs. Conferência"])
        
        headers = [str(h).strip() if h is not None else f"Col_{i}" for i, h in enumerate(rows[0])]
        df_raw = pd.DataFrame(rows[1:], columns=headers)
    else:
        # Modo CSV
        df_raw = pd.read_csv(path_str, sep=';', encoding='latin1', dtype=str)
        df_raw.columns = [c.strip() for c in df_raw.columns]
    
    # Identificar colunas correspondentes de forma resiliente
    col_emp = "Emp" if "Emp" in df_raw.columns else df_raw.columns[0]
    col_mat = "Matrícula" if "Matrícula" in df_raw.columns else df_raw.columns[1]
    col_evt = "Evento" if "Evento" in df_raw.columns else df_raw.columns[2]
    col_mes = "Mês" if "Mês" in df_raw.columns else df_raw.columns[3]
    col_val = "Valor" if "Valor" in df_raw.columns else df_raw.columns[4]
    col_obs = "Obs. Conferência" if "Obs. Conferência" in df_raw.columns else None
    
    res = pd.DataFrame()
    res["Emp"] = pd.to_numeric(df_raw[col_emp], errors="coerce").fillna(0).astype(int)
    res["Matrícula"] = normalize_matricula_series(df_raw[col_mat])
    res["Evento"] = pd.to_numeric(df_raw[col_evt], errors="coerce").fillna(29613).astype(int)
    res["Mês"] = pd.to_numeric(df_raw[col_mes], errors="coerce").fillna(0).astype(int)
    res["Valor"] = df_raw[col_val].apply(parse_brazilian_currency)
    
    if col_obs and col_obs in df_raw.columns:
        res["Obs. Conferência"] = df_raw[col_obs].fillna("").astype(str).str.strip()
    else:
        res["Obs. Conferência"] = ""
        
    # Filtrar eventuais linhas vazias
    res = res[res["Matrícula"] != ""].copy()
    
    return res
```

### 7.3 Métricas de Validação no Modelo Base
- **Total de Linhas Ingeridas**: 1.334 registros
- **Soma Acumulada do Campo `Valor`**: **R$ 482.265,63** (conferência 100% idêntica ao somatório do arquivo `Conferência Good Card_Base.xlsx`).

---

## 8. Especificação dos Testes Unitários (`test_data_loader.py`)

A suíte de testes unitários do Milestone 1 deve validar isoladamente cada transformação e caso de borda sem depender da existência dos arquivos completos em disco (utilizando fixtures e dados sintéticos):

### 8.1 Cenários de Teste da Normalização de Matrícula
1. `test_normalize_matricula_numeric()`:
   - Entradas: `305204`, `305204.0`, `'305204'`, `' 305204 '`, `'305204.0'`, `'305204.00'`.
   - Saída esperada: `'305204'` para todas.
2. `test_normalize_matricula_leading_zeros()`:
   - Entradas: `'001234'`, `'  001234  '`, `'001234.0'`.
   - Saída esperada: `'001234'` (não pode virar `'1234'`).
3. `test_normalize_matricula_nulls_and_hyphens()`:
   - Entradas: `None`, `np.nan`, `pd.NA`, `""`, `"-"`, `" -"`.
   - Saída esperada: `""`.
4. `test_normalize_matricula_series_vectorized()`:
   - Testa uma Series heterogênea contendo inteiros, floats, strings, nulos e traços.

### 8.2 Cenários de Teste do Parsing Monetário
1. `test_parse_brazilian_currency_strings()`:
   - Entradas: `'5.582,44'`, `'317,06'`, `'0,00'`, `'-'`, `' -'`.
   - Saídas esperadas: `5582.44`, `317.06`, `0.0`, `0.0`, `0.0`.
2. `test_parse_brazilian_currency_floats()`:
   - Entradas: `301.85`, `0.0`, `5582.44`.
   - Saídas esperadas: `301.85`, `0.0`, `5582.44` (garante que não multiplicou por 100).

### 8.3 Cenários de Teste de Deduplicação de Demitidos
1. `test_deduplicate_demitidos_transferencia_sem_onus()`:
   - Dados sintéticos com 1 registro normal e 1 registro de `"5 - Transferência sem ônus"`.
   - Verifica se o registro de transferência sem ônus foi removido.
2. `test_deduplicate_demitidos_most_recent_date()`:
   - Colaborador com duas rescisões: `01/07/2025` e `01/09/2026`.
   - Verifica se o registro de `01/09/2026` foi retido e o de `2025` descartado.

### 8.4 Cenários de Teste de Deduplicação de Afastados
1. `test_deduplicate_afastados_past_termination()`:
   - Registro encerrado em `15/08/2026` vs competência `01/09/2026`.
   - Verifica se foi descartado.
2. `test_deduplicate_afastados_open_ended_priority()`:
   - Colaborador com afastamento de 30 dias fechado vs afastamento aberto contínuo (`Término` vazio).
   - Verifica se o afastamento aberto foi retido.

### 8.5 Cenários de Teste de Deduplicação de Férias
1. `test_deduplicate_ferias_outside_audit_month()`:
   - Férias de `01/07/2026` a `30/07/2026` vs mês `01/09/2026` a `30/09/2026`.
   - Verifica descarte por não sobreposição.
2. `test_deduplicate_ferias_active_status_priority()`:
   - Dois períodos em setembro: um com status `'Gozada'` e outro com status `'Em férias'`.
   - Verifica se o registro com status `'Em férias'` foi mantido.

### 8.6 Cenários de Teste de Carga de Valores Importados
1. `test_load_valores_importados_from_excel()`:
   - Mock ou leitura de planilha com 5 colunas base.
   - Verifica tipagem de `Emp` (int), `Matrícula` (str limpa), `Valor` (float).
2. `test_load_valores_importados_from_csv()`:
   - Cria arquivo CSV temporário em memória (`io.StringIO`) com delimitador `;` e testa a ingestão.

---

## 9. Contrato de Interface Oficial de `data_loader.py`

Para o orquestrador `main.py` e o motor analítico `business_rules.py`, o módulo `data_loader.py` expõe a seguinte API pública:

```python
# data_loader.py API Contract

def load_fopa(filepath: str) -> pd.DataFrame:
    """Carrega folha FOPA com skiprows=3, sep=';', encoding='latin1'."""
    ...

def load_colaboradores(filepath: str) -> pd.DataFrame:
    """Carrega cadastro ativo com skiprows=1, sep=';', encoding='latin1'."""
    ...

def load_afastados(filepath: str, ref_month_start: str = "2026-09-01") -> pd.DataFrame:
    """Carrega afastados deduplicados e filtrados."""
    ...

def load_demitidos(filepath: str) -> pd.DataFrame:
    """Carrega demitidos deduplicados e filtrados."""
    ...

def load_ferias(filepath: str, ref_month_start: str = "2026-09-01", ref_month_end: str = "2026-09-30") -> pd.DataFrame:
    """Carrega férias deduplicadas e filtradas por sobreposição."""
    ...

def load_valores_importados(source_path: Union[str, Path], sheet_name: str = "Valores Importados") -> pd.DataFrame:
    """Carrega valores importados do fornecedor (Excel ou CSV)."""
    ...

def load_all_data(data_dir: str, template_path: str, ref_month: str = "2026-09") -> dict[str, pd.DataFrame]:
    """
    Função agregadora que executa a carga completa e retorna dicionário com os 6 DataFrames:
    - 'fopa': 15.913 linhas
    - 'nomina': 16.038 linhas
    - 'demitidos': 12.901 linhas
    - 'afastados': 1.867 linhas
    - 'ferias': 2.005 linhas
    - 'valores_importados': 1.334 linhas
    """
    ...
```

---
*Fim da Especificação Técnica do Milestone 1.*
