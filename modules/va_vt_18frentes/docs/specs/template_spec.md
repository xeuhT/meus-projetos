# Especificação Técnica do Template Excel: Conferência Good Card

**Arquivo Base Analisado**: `Conferência Good Card_Base.xlsx`  
**Data da Análise**: 2026-10-06  
**Autor**: spec_miner_survey_1  

---

## 1. Visão Geral das Abas do Workbook

O arquivo base original `Conferência Good Card_Base.xlsx` possui 10 abas:
1. `Orientações` (Aba oficial de instruções metodológicas e regras de conferência)
2. `Conferência` (Aba analítica principal com 38 colunas de auditoria)
3. `Valores Importados` (Aba de conciliação do evento 29613 com 11 colunas)
4. `FOPA_Eventos_por_Colaborador - ` (Base importada da folha de pagamento com eventos)
5. `Desligado` (Relatório de colaboradores desligados)
6. `Afastamento` (Relatório de colaboradores afastados)
7. `Férias` (Relatório de programação e gozo de férias)
8. `Nomina` (Relatório cadastral / situação eSocial)
9. `Evidência` (Aba de anotações de auditoria)
10. `Lista` (Tabela de domínio com os 9 status oficiais de parecer/tratativa)

### Estrutura do Arquivo Final Gerado (`.xlsx`)
Conforme definido em `ORIGINAL_REQUEST.md`, o script automatizado processa todos os CSVs em memória e exporta uma planilha **estática (sem fórmulas)** contendo exatamente as 4 abas finais:
1. **`Orientações`**: Preserva fielmente a formatação, layout, textos e blocos do modelo base.
2. **`Resumo da Auditoria`**: Painel executivo / Dashboard com KPIs consolidados, indicadores de conformidade, distribuição de pareceres e resumo da conciliação do evento 29613.
3. **`Conferência`**: As 38 colunas exatas da base, preenchidas estaticamente em memória a partir dos cruzamentos de dados.
4. **`Valores Importados`**: As 11 colunas exatas da base, conciliando valores importados do evento 29613 contra a folha.

---

## 2. Especificação Detalhada: Aba "Orientações"

### 2.1 Propriedades da Planilha
- **Título da Aba**: `Orientações`
- **Freeze Panes**: Nenhum (`None`)
- **Larguras de Coluna**:
  - Coluna A: `30.85` (ou `35.0`)
  - Coluna B: `100.0` (com texto quebrado / `wrap_text=True`)
  - Coluna C+: padrão

### 2.2 Textos, Coordenadas e Estilos

A aba é organizada em 3 seções conceituais com cabeçalhos destacados:

| Linha | Coluna A (Título/Assunto) | Coluna B (Texto de Orientação) | Formatação |
|---|---|---|---|
| **2** | **Relatórios** | **Orientações** | Cabeçalho Seção 1: Fonte Aptos Narrow 11pt Negrito, Fundo Cinza Claro (`#D9D9D9` / theme 0 tint -0.15), Borda Fina |
| 3 | Afastados | Verificar se existe colaboradores com mais de um período de afastamento e deve deixar somente o que possui data mais extensa | Fonte 11pt regular, alinhamento superior esquerdo, wrap_text=True |
| 4 | Afastados | Não deve considerar os afastamentos que foram finalizados antes da primeira data do mês a que se refere a compra | Fonte 11pt regular, alinhamento superior esquerdo, wrap_text=True |
| 5 | Demitidos | Não deve considerar os colaboradores com tipo de desligamento "Transferência sem ônus" | Fonte 11pt regular, alinhamento superior esquerdo, wrap_text=True |
| 6 | Férias | Verificar se existe claboradores com férias quebradas que comprendem os dois meses de referência do relatório e deve manter somente o período que compreende o mês da compra | Fonte 11pt regular, alinhamento superior esquerdo, wrap_text=True |
| 7 | Férias | Verificar colaboradores que iniciaram e finalizaram as férias no mês que não compete o período da compra. | Fonte 11pt regular, alinhamento superior esquerdo, wrap_text=True |
| 8 | Eventos Selecionados - Coluna | Em gerador de relatórios deve pesquisar em Parâmetros o filtro com o nome " Conferência Good Card" alterar para o mês da folha correspondente e emitir. | Fonte 11pt regular, alinhamento superior esquerdo, wrap_text=True |
| 9 | Eventos Selecionados - Coluna | Nas colunas relacionadas a eventos deve realizar a substituição dos "-" por "0". | Fonte 11pt regular, alinhamento superior esquerdo, wrap_text=True |
| 10 | Nomina | - | Fonte 11pt regular |
| 11 | *(vazia)* | *(vazia)* | - |
| **12** | **Aba** | **Orientações** | Cabeçalho Seção 2: Fonte Aptos Narrow 11pt Negrito, Fundo Cinza Claro (`#D9D9D9`), Borda Fina |
| 13 | Conferência | Na aba conferência o relatório base será o de Eventos Selecionados - Coluna. | Fonte 11pt regular |
| 14 | Conferência | Nessa aba irá colar as informações das coluna "Códido Empresa" até "Cargo". | Fonte 11pt regular |
| 15 | Conferência | Cuidado ao colar as informações para não sobrepor as colunas que possuem fórmula. | Fonte 11pt regular |
| 16 | Conferência | As colunas que possuem fórmula estão pintadas com a cor cinza claro | Fonte 11pt regular |
| 17 | *(vazia)* | *(vazia)* | - |
| **18** | **Conferências a serem realizadas** | **Orientações** | Cabeçalho Seção 3: Fonte Aptos Narrow 11pt Negrito, Fundo Cinza Claro (`#D9D9D9`), Borda Fina |
| 19 | Cargos | Os colaboradores nos cargos de aprendiz e diretor não devem ter Good Card então deve conferi se existe algum colaborador em um desses cargos | Fonte 11pt regular |
| 20 | Custo Empresa | No Giga, o parâmetro de desconto é diferente das demais bandeiras, sendo fixado em 5% por definição da Diretoria da bandeira.<br>Como em SP não possuímos lojas como GBarbosa, Mercantil ou Perini, e a única possibilidade de compras seria nas lojas do Giga, o estabelecimento 400 precisou seguir o mesmo critério de desconto do Giga. | Altura da linha: `30.75`. Fonte 11pt regular, quebra de linha |
| 21 | Custo Empresa | Validar se o valor levado para o custo empresa correspondente realmente a 5% do saldo devedor | Fonte 11pt regular |
| 22 | Desconto | Deve validar se existe algum desconto ocorrendo fora da margem dos 20% | Fonte 11pt regular |
| 23 | Saldo devedor | Colaboradores com saldo devedor zerado se tem desconto | Fonte 11pt regular |
| 24 | Saldo devedor | Colaboradores com saldo devedor e que não tiveram desconto | Fonte 11pt regular |
| 25 | Saldo devedor | Saldo devedor " Dentro da Margem" deve ter o saldo totalmente liquidado sem apresentar valores no evento 29621 | Fonte 11pt regular |
| 26 | Saldo devedor | Salde devedor "Fora da Margem" deve descontar dentro da margem dos 20% e o restante levar para o evento 29621 | Fonte 11pt regular |

---

## 3. Especificação Detalhada: Aba "Resumo da Auditoria"

### 3.1 Finalidade e Estrutura
A aba **Resumo da Auditoria** opera como o painel executivo (dashboard de KPIs). Ela consolida os números totais da auditoria, permitindo que a liderança visualize em segundos a taxa de conformidade da folha e a distribuição das inconsistências.

### 3.2 Layout e Estilos Visuais
- **Visualização**: `showGridLines = True`
- **Banner de Título**:
  - Intervalo mesclado: `B2:F2`
  - Valor: `PAINEL DE AUDITORIA - CONFERÊNCIA GOOD CARD`
  - Altura da linha 2: `35`
  - Preenchimento: Azul Institucional Cencosud (`1F497D`)
  - Fonte: Calibri/Aptos 14pt, Negrito, Branca (`FFFFFF`)
  - Alinhamento: Centralizado horizontal e verticalmente

### 3.3 Tabela 1: KPIs Executivos Principais (Linhas 4 a 11)
| Linha | Célula (Rótulo - Coluna B) | Célula (Valor - Coluna D) | Formato Numérico | Descrição / Fórmula de Negócio |
|---|---|---|---|---|
| 4 | Total de Colaboradores Auditados | `len(df_conferencia)` | `#,##0` | Total de colaboradores presentes na base FOPA auditada |
| 5 | Colaboradores com Saldo Devedor Good Card | `(df['Total Saldo Devedor'] > 0).sum()` | `#,##0` | Colaboradores com saldo a pagar/descontar |
| 6 | Colaboradores com Desconto em Folha (6280) | `(df['6280 - Desc Cartão Benefícios Especiais'] > 0).sum()` | `#,##0` | Colaboradores que sofreram débito em folha |
| 7 | Colaboradores Dentro da Margem de 20% | `(df['Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"'] == 'Dentro da Margem').sum()` | `#,##0` | Saldo inferior ou igual a 20% do salário |
| 8 | Colaboradores Fora da Margem de 20% | `(df['Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"'] == 'Fora da Margem').sum()` | `#,##0` | Saldo excedente aos 20% do salário |
| 9 | Saldo Devedor Transferido (Evt 29621) | `(df['29621 - Ben Esp - Saldo Devedor Beneficios Especiais'] > 0).sum()` | `#,##0` | Colaboradores com saldo levado p/ mês seguinte |
| 10 | Registros em Plena Conformidade (OK) | `total_ok` | `#,##0` | Registros sem nenhuma divergência apontada |
| 11 | Divergências / Inconsistências Apontadas | `total_divergencias` | `#,##0` | Registros que exigem análise/tratativa |
| 12 | Índice Geral de Conformidade da Folha | `(total_ok / total_auditados) * 100` | `0.0%` | Percentual de assertividade da folha |

- **Estilos das Linhas de KPI**:
  - Rótulos (Coluna B): Fonte 11pt, Negrito
  - Valores (Coluna D): Fonte 11pt, Negrito, Alinhamento à Direita

### 3.4 Tabela 2: Distribuição dos Status da Auditoria (Linha 14+)
- **Título (B14)**: `DISTRIBUIÇÃO DOS STATUS DA AUDITORIA (Obs. Conferência)` (Fonte 11pt Negrito, Cor `1F497D`)
- **Cabeçalho (B15)**: `Parecer da Auditoria (Coluna 37)` | **(D15)**: `Quantidade` (Alinhado à direita)
- **Status da Tabela de Domínio (da aba `Lista`)**:
  1. `Ok - Sem saldo devedor e sem desconto`
  2. `Ok - Liquidou todo o saldo devedor `
  3. `Ok - Suspensão de Contrato`
  4. `Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621`
  5. `Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)`
  6. `Afastamento - Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 `
  7. `Férias- Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 `
  8. `Verificar com quality - Não estava em situação de afastamento e não houve o desconto`
  9. `Aprendiz - Verificar`
- **Destaque Visual de Alerta**:
  - Para linhas de parecer que **NÃO** contenham "Ok", aplicar:
    - Preenchimento: Salmão claro (`FFC7CE`)
    - Fonte: Vermelho escuro (`9C0006`), Negrito

### 3.5 Tabela 3: Conciliação de Valores Importados (Evento 29613)
- Linhas abaixo da distribuição de pareceres:
  - `Total de Registros Importados (Evt 29613)`
  - `Registros Conciliados com a Folha (Valor Importado == Folha)`
  - `Divergências de Valor (Importado != Folha)`
  - `Colaboradores Desligados com Evento Importado`

### 3.6 Larguras das Colunas do Dashboard
- Coluna B: `48`
- Coluna C: `5`
- Coluna D: `20`
- Coluna E: `5`
- Coluna F: `15`

---

## 4. Especificação Detalhada: Aba "Conferência"

### 4.1 Propriedades Gerais
- **Título da Aba**: `Conferência`
- **Total de Colunas**: Exatamente **38** colunas
- **Freeze Panes**: **`A2`** (Congela a primeira linha, mantendo o cabeçalho sempre visível na rolagem)
- **Altura da Linha de Cabeçalho (Linha 1)**: `28`

### 4.2 Tabela Oficial das 38 Colunas (Ordem Estrita 1 a 38)

| # | Letra | Nome Exato da Coluna | Tipo de Dado Original | Alinhamento Requerido | Formato Numérico | Largura Padrão | Regra / Origem dos Dados (Cálculo em Memória) |
|---|---|---|---|---|---|---|---|
| 1 | A | `Código da Empresa` | Inteiro | Centralizado | `General` | 20.14 | FOPA: Coluna 1 |
| 2 | B | `Nome da Empresa` | Texto | À Esquerda | `General` | 36.00 | FOPA: Coluna 2 |
| 3 | C | `Código Estabelecimento` | Inteiro | Centralizado | `General` | 24.28 | FOPA: Coluna 3 |
| 4 | D | `Estabelecimento` | Texto | À Esquerda | `General` | 50.43 | FOPA: Coluna 4 |
| 5 | E | `Código Centro Custo` | Texto | Centralizado | `General` | 20.85 | FOPA: Coluna 5 |
| 6 | F | `Centro Custo` | Texto | À Esquerda | `General` | 43.71 | FOPA: Coluna 6 |
| 7 | G | `Matrícula` | Texto/ID | Centralizado | `General` / `@` | 11.57 | FOPA: Coluna 7 (Chave mestre de cruzamento) |
| 8 | H | `Nome do Colaborador` | Texto | À Esquerda | `General` | 56.43 | FOPA: Coluna 8 |
| 9 | I | `ID Colaborador` | Inteiro | Centralizado | `General` | 16.57 | FOPA: Coluna 9 |
| 10 | J | `Data Admissão` | Data | Centralizado | `dd/mm/yyyy` | 13.00 | FOPA: Coluna 10 |
| 11 | K | `Código Cargo` | Inteiro | Centralizado | `General` | 14.85 | FOPA: Coluna 11 |
| 12 | L | `Cargo` | Texto | À Esquerda | `General` | 81.14 | FOPA: Coluna 12 |
| 13 | M | `1 - Salario` | Decimal | Centralizado | `R$ #,##0.00` | 14.43 | FOPA: Valor do Evento 1 |
| 14 | N | `11010 - Salario Mensal` | Decimal | Centralizado | `R$ #,##0.00` | 22.85 | FOPA: Valor do Evento 11010 |
| 15 | O | `10990 - Liquido a Receber` | Decimal | Centralizado | `R$ #,##0.00` | 25.28 | FOPA: Valor do Evento 10990 |
| 16 | P | `29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ` | Decimal | Centralizado | `R$ #,##0.00` | 56.43 | FOPA: Valor do Evento 29622 *(atenção ao espaço final)* |
| 17 | Q | `29613 - Ben Esp - Valor Informado para Desconto` | Decimal | Centralizado | `R$ #,##0.00` | 45.85 | FOPA: Valor do Evento 29613 |
| 18 | R | `Total Saldo Devedor` | Decimal | Centralizado | `R$ #,##0.00` | 21.14 | Soma: `Evt 29622 (Col 16) + Evt 29613 (Col 17)` |
| 19 | S | `Margem 20% do Salário` | Decimal | Centralizado | `R$ #,##0.00` | 26.43 | Cálculo: `Evt 11010 (Col 14) * 0.20` |
| 20 | T | `Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"` | Texto | Centralizado | `General` | 63.00 | Se `round(Total Saldo Devedor, 2) <= round(Margem 20%, 2)`: "Dentro da Margem", senão "Fora da Margem" |
| 21 | U | `6280 - Desc Cartão Benefícios Especiais` | Decimal | Centralizado | `R$ #,##0.00` | 37.71 | FOPA: Valor do Evento 6280 |
| 22 | V | `Desconto x Margem de 20%` | Texto | Centralizado | `General` | 27.57 | Se `round(Evt 6280, 2) <= round(Margem 20%, 2)`: "Dentro da Margem", senão "Fora da Margem" |
| 23 | W | `Saldo Devedor x Evt 6280` | Booleano | Centralizado | `General` | 52.00 | Validação: `(Evt 29624 + Evt 6280) == Total Saldo Devedor` |
| 24 | X | `Total Saldo Devedor - Evt 6280 x Evt 29621` | Booleano | Centralizado | `General` | 47.14 | Validação: `(Evt 29624 + Evt 6280) == Total Saldo Devedor` |
| 25 | Y | `29621 - Ben Esp - Saldo Devedor Beneficios Especiais` | Decimal | Centralizado | `R$ #,##0.00` | 49.14 | FOPA: Valor do Evento 29621 |
| 26 | Z | `29624 - Ben Esp - Custo Empresa Beneficios Especiais` | Decimal | Centralizado | `R$ #,##0.00` | 49.71 | FOPA: Valor do Evento 29624 |
| 27 | AA | `5% Saldo devedor` | Decimal | Centralizado | `R$ #,##0.00` | 18.57 | Cálculo: `Total Saldo Devedor (Col 18) * 0.05` |
| 28 | AB | `Cálculo 5% x 29624` | Decimal | Centralizado | `R$ #,##0.00` | 19.57 | Diferença: `Evt 29624 (Col 26) - 5% Saldo Devedor (Col 27)` |
| 29 | AC | `Início Afastamento` | Data | Centralizado | `dd/mm/yyyy` | 19.71 | Cruzamento `COLA_Colaboradores_Afastados` (Data Início) |
| 30 | AD | `Final Afastamento` | Data | Centralizado | `dd/mm/yyyy` | 19.14 | Cruzamento `COLA_Colaboradores_Afastados` (Data Fim) |
| 31 | AE | `Tipo do Afastamento` | Texto | Centralizado | `General` | 13.00 | Cruzamento `COLA_Colaboradores_Afastados` (Motivo/Tipo) |
| 32 | AF | `Início Férias` | Data | Centralizado | `dd/mm/yyyy` | 14.00 | Cruzamento `COLA_Ferias_Colaboradores` (Data Início Gozo) |
| 33 | AG | `Final Férias` | Data | Centralizado | `dd/mm/yyyy` | 13.43 | Cruzamento `COLA_Ferias_Colaboradores` (Data Fim Gozo) |
| 34 | AH | `Qtd de dias` | Inteiro | Centralizado | `General` | 13.14 | Cruzamento `COLA_Ferias_Colaboradores` (Dias de férias) |
| 35 | AI | `Desligamento` | Data | Centralizado | `dd/mm/yyyy` | 13.00 | Cruzamento `COLA_Demitidos` (Data Demissão) |
| 36 | AJ | `Situação` | Texto | Centralizado | `General` | 29.00 | Cruzamento `COLA_Colaboradores` / Nomina (Situação funcional) |
| 37 | AK | `Obs. Conferência` | Texto | Centralizado | `General` | 123.00 | Parecer da auditoria calculado conforme árvore de regras |
| 38 | AL | `Tratativa` | Texto | À Esquerda | `General` | 15.00 | Ações operacionais / comentários de correção |

### 4.3 Destaques de Cores do Cabeçalho
- **Colunas 1 a 19 (A a S)**: Fundo Padrão / Neutro
- **Coluna 20 (T - `Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"`)**: Fundo Vermelho Sólido (`FFFF0000`), Texto Branco Negrito (`FFFFFF`)
- **Coluna 21 (U - `6280 - Desc Cartão Benefícios Especiais`)**: Fundo Azul Claro Acento (`theme:5 tint=0.4` ou `#B8CCE4`)
- **Colunas 29 a 36 (AC a AJ - Afastamentos, Férias, Demissões, Situação)**: Fundo Laranja/Ouro Claro Acento (`theme:8 tint=0.4` ou `#FDE9D9`)
- **Coluna 37 (AK - `Obs. Conferência`)**: Fundo Laranja Muito Claro (`theme:8 tint=0.8` ou `#FFF2E6`), indicando a coluna síntese da auditoria.

### 4.4 Formatação das Células Calculadas (Cinza Claro)
Conforme orientações oficiais da empresa (linha 16 de `Orientações`):
*"As colunas que possuem fórmula estão pintadas com a cor cinza claro"*.
Nas linhas de dados, as colunas calculadas/derivadas de cruzamento (colunas 13 a 36) recebem preenchimento cinza sutil:
- **Cor**: `theme:0 tint=-0.05` ou `#F2F2F2` (cinza claríssimo de proteção visual).

### 4.5 Regras de Destaque de Divergência (Linhas e Células)
Quando um registro auditado apresentar inconsistência (qualquer parecer que **não** seja `Ok...`):
- **Preenchimento da Linha Divergente**: Fundo de alerta suave (`#FFF0F2`).
- **Células Críticas da Divergência** (em especial Coluna 37 `Obs. Conferência` e Coluna 38 `Tratativa`, ou colunas com incompatibilidade de margem/desconto):
  - Preenchimento: `PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")`
  - Fonte: `Font(name="Calibri", size=10, bold=True, color="9C0006")`

---

## 5. Especificação Detalhada: Aba "Valores Importados"

### 5.1 Finalidade
Conciliar o relatório de importação de débitos do cartão (evento 29613) contra os valores efetivamente lançados na folha de pagamento (`FOPA`), identificando divergências de valor, descontos duplicados ou lançamentos indevidos para colaboradores demitidos.

### 5.2 Propriedades da Planilha
- **Total de Colunas**: Exatamente **11** colunas
- **Freeze Panes**: **`A2`** (Linha 1 congelada)
- **Altura do Cabeçalho**: `26` a `28`

### 5.3 Tabela Oficial das 11 Colunas

| # | Letra | Nome Exato da Coluna | Tipo de Dado | Alinhamento | Formato Numérico | Largura | Descrição / Cálculo |
|---|---|---|---|---|---|---|---|
| 1 | A | `Emp` | Inteiro | Centralizado | `General` | 9.43 | Código da Empresa |
| 2 | B | `Matrícula` | Texto/ID | Centralizado | `General` / `@` | 14.00 | Matrícula do Colaborador |
| 3 | C | `Evento` | Inteiro | Centralizado | `General` | 11.29 | Código do Evento (`29613`) |
| 4 | D | `Mês` | Texto/Int | Centralizado | `General` | 9.14 | Competência da Folha (ex: `92026`) |
| 5 | E | `Valor` | Decimal | Centralizado | `R$ #,##0.00` | 10.29 | Valor importado do arquivo Good Card |
| 6 | F | `Folha` | Decimal | Centralizado | `R$ #,##0.00` | 12.00 | Valor lançado na folha (`FOPA`, evento 29613) |
| 7 | G | `Importado x Folha` | Booleano | Centralizado | `General` | 17.29 | Validação: `Valor == Folha` (`True`/`False`) |
| 8 | H | `Desligado` | Data | Centralizado | `dd/mm/yyyy` | 17.29 | Data de demissão (se houver, via `COLA_Demitidos`) |
| 9 | I | `Tipo de Rescisão` | Texto | À Esquerda | `General` | 32.86 | Motivo/Tipo do desligamento |
| 10 | J | `Obs. Conferência` | Texto | À Esquerda | `General` | 62.29 | Anotações (ex: "Transferência de empresa...", "Demitido antes da compra") |
| 11 | K | `Tratativa` | Texto | À Esquerda | `General` | 80.00 | Ação requerida de ajuste operacional |

### 5.4 Regras de Destaque em "Valores Importados"
- Se `Importado x Folha` for `False` (divergência de centavos ou valor ausente na folha):
  - Preenchimento da célula e da linha: Salmão claro (`FFC7CE`), fonte vermelha negrito (`9C0006`).
- Se o colaborador estiver `Desligado` e houver valor importado:
  - Destacar coluna `Desligado` e `Tipo de Rescisão` em vermelho de alerta.

---

## 6. Arquitetura Estática vs. Fórmulas Dinâmicas

### 6.1 Problema Identificado no Modelo Base Original
Na planilha base `Conferência Good Card_Base.xlsx`, o arquivo possui fórmulas dinâmicas (`VLOOKUP`, `=P2+Q2`, `=E2=F2`) espalhadas em 15.914 linhas e 38 colunas. Isso totaliza mais de **380.000 fórmulas no workbook**, causando:
1. Tamanho inflado do arquivo (23 MB).
2. Travamento completo do Excel ao abrir e recalcular referências.
3. Necessidade de abas intermediárias pesadas (`Nomina`, `Desligado`, `Afastamento`, etc.) mantidas apenas para permitir `VLOOKUP`.

### 6.2 Solução Arquitetural de Alta Performance
O módulo `excel_exporter.py`:
1. **Zero Fórmulas no Arquivo Exportado**: Todas as 38 colunas de `Conferência` e 11 colunas de `Valores Importados` recebem valores escalares computados previamente em memória pelo Pandas e pelas regras de negócio.
2. **Formatação Nativa Aberta**: Valores numéricos são inseridos como float/int e estilizados com `cell.number_format = 'R$ #,##0.00'` ou `dd/mm/yyyy`, garantindo que o usuário possa fazer somas e filtros normalmente no Excel.
3. **Economia de Recursos**: A planilha final conterá apenas as 4 abas relevantes (`Orientações`, `Resumo da Auditoria`, `Conferência`, `Valores Importados`), reduzindo o tamanho de 23 MB para menos de 3 MB e abrindo instantaneamente no Excel.

---

## 7. Tabelas de Descoberta e Casos de Borda

### Features Discovered
| # | Categoria | Feature | Descrição | Inputs | Outputs | Comportamento de Erro | Descoberto Via |
|---|---|---|---|---|---|---|---|
| 1 | Estrutura | 4 Abas Oficiais | Geração estrita das 4 abas na ordem oficial: Orientações, Resumo da Auditoria, Conferência, Valores Importados | Dataframes processados | Workbook Excel (.xlsx) | Abas vazias tratadas sem quebra | `ORIGINAL_REQUEST.md` e Base Excel |
| 2 | Layout | Orientações Verbatim | Manter os 3 blocos originais (Relatórios, Aba, Conferências a serem realizadas) com suas 24 orientações | Textos estáticos oficiais | Aba Orientações formatada | Fallback seguro com dimensões pré-ajustadas | Inspeção `Conferência Good Card_Base.xlsx` |
| 3 | Dashboard | Painel de KPIs Executivos | Cabeçalho `1F497D`, indicadores de volume, margem 20%, desvios e taxa de conformidade | `df_conferencia` | Resumo da Auditoria B2:F12 | Evitar divisão por zero caso base venha vazia | Projeto de referência `auditoria_va_vt` |
| 4 | Dashboard | Tabela de Status de Parecer | Distribuição exata dos 9 status oficiais da folha com destaque vermelho para desvios | Série `Obs. Conferência` | Tabela dinâmica em B14:D25 | Agrupar status inesperados em "Outros" | Sheet `Lista` da base |
| 5 | Conferência | 38 Colunas Oficiais | Grid analítico completo contendo eventos, margens de 20%, custos e status funcional | Base FOPA + Cadastros | 38 colunas estáticas em ordem A-AL | Colunas ausentes preenchidas com nulo/vazio | Inspeção direta em `Conferência` linha 1 e 2 |
| 6 | Conferência | Proteção Cinza Claro | Preenchimento cinza sutil (`#F2F2F2`) nas colunas calculadas (13 a 36) | Índices de colunas 13 a 36 | Fill aplicado nas células de dados | N/A | Regra declarada na aba Orientações linha 16 |
| 7 | Usabilidade | Freeze Panes A2 | Congelamento rigoroso da linha 1 nas abas analíticas | Propriedade `ws.freeze_panes = 'A2'` | Cabeçalho fixo na rolagem vertical | N/A | Exigência explícita do usuário |
| 8 | Usabilidade | Alinhamento Seletivo | Centralização de códigos, matrículas, datas, eventos e valores; alinhamento à esquerda de textos | Regras de tipagem por coluna | Alinhamento de célula | Strings não conversíveis mantidas à esquerda | `ORIGINAL_REQUEST.md` |
| 9 | Valores Importados | Conciliação Evt 29613 | 11 colunas comparando valor importado vs lançado na folha com alerta para divergências e demitidos | CSV Importado + FOPA + Demitidos | 11 colunas na aba Valores Importados | Tratamento para matrículas não localizadas | Inspeção direta na aba `Valores Importados` |
| 10 | Destaque Visual | Marcação Vermelha | Preenchimento `FFC7CE` e fonte `9C0006` negrito para pareceres não conformes | Colunas `Obs. Conferência`, `Tratativa` | Estilo OpenPyXL | Registros OK mantêm fundo normal | Inspeção do projeto irmão e regras de UX |

### Edge Cases
| # | Feature | Input / Cenário | Comportamento Observado / Esperado |
|---|---|---|---|
| 1 | Tratamento de Eventos Nulos | Colunas de eventos no FOPA preenchidas com `-` ou vazias | Devem ser convertidas para `0.00` antes dos cálculos, conforme orientação linha 9 de `Orientações`. |
| 2 | Nome de Coluna com Espaço | Coluna 16 possui espaço extra no final (`...Mes Ant `) | Manter rigorosamente o nome com espaço para compatibilidade exata com o modelo. |
| 3 | Colaborador Giga / Est. 400 | Colaborador lotado na bandeira Giga ou estabelecimento 400 | Desconto com parâmetro de custo empresa de 5% sobre o saldo devedor (linhas 20 e 21 de `Orientações`). |
| 4 | Afastamento com Múltiplos Registros | Colaborador com mais de uma linha no relatório de afastamento | Filtrar apenas o período com data mais extensa vigente no mês de competência (linha 3 de `Orientações`). |
| 5 | Férias Quebradas | Férias que iniciam em um mês e terminam em outro | Considerar apenas o período de gozo que coincide com o mês auditado (linha 6 de `Orientações`). |
| 6 | Demitidos "Transferência sem ônus" | Colaborador demitido com tipo de desligamento "Transferência sem ônus" | Desconsiderar bloqueio indevido de Good Card para esses casos (linha 5 de `Orientações`). |
| 7 | Cargos Inelegíveis | Colaboradores aprendizes ou diretores com lançamentos Good Card | Sinalizar parecer `Aprendiz - Verificar` e destacar linha com alerta vermelho (linha 19 de `Orientações`). |
| 8 | Formatação de Matrícula | Matrículas com zeros à esquerda (ex: `000935`) | Manter representação em string pura para evitar perda de zeros e falhas de join. |
