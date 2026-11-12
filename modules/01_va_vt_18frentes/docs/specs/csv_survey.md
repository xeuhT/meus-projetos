# Relatório Técnico de Mapeamento e Análise dos Arquivos CSV
**Auditoria de Benefícios "04 - GOOD CARD" — Cencosud Brasil**  
**Data da Análise:** 06/10/2026  
**Investigador:** teamwork_preview_explorer (`explorer_survey_csv_2`)  
**Diretório de Trabalho:** `C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\04 - GOOD CARD`

---

## 1. Visão Geral dos Arquivos e Estrutura do Workspace

O repositório do projeto contém os arquivos originais em formato CSV na pasta `Relatorio Original`, versões tratadas em formato Excel na pasta `Relatório Tratado`, e a planilha modelo de referência `Conferência Good Card_Base.xlsx` na raiz:

| Caminho do Arquivo | Tipo / Formato | Tamanho (Bytes) | Tamanho Aprox. | Função no Fluxo de Negócio |
| :--- | :---: | :---: | :---: | :--- |
| `Relatorio Original\COLA_Colaboradores.CSV` | CSV | 24.908.301 | 24,9 MB | Cadastro ativo de colaboradores (Nomina), dados contratuais, cargos, salários. |
| `Relatorio Original\COLA_Colaboradores_Afastados.CSV` | CSV | 822.460 | 822 KB | Registro de afastamentos previdenciários e atestados médicos (>15 dias). |
| `Relatorio Original\COLA_Demitidos.CSV` | CSV | 13.711.091 | 13,7 MB | Histórico de desligamentos, rescisões e tipos de rescisão. |
| `Relatorio Original\COLA_Ferias_Colaboradores.CSV` | CSV | 561.339 | 561 KB | Escala e períodos de gozo de férias dos colaboradores. |
| `Relatorio Original\FOPA_Eventos_por_Colaborador - 22092026_091703_796705.csv` | CSV | 4.245.243 | 4,2 MB | Folha de pagamento (competência set/26) com eventos financeiros pivotados. |
| `Conferência Good Card_Base.xlsx` | XLSX | 23.067.555 | 23,0 MB | Modelo Excel mestre de referência (contém abas Orientações, Conferência, Valores Importados, etc.). |
| `Relatório Tratado\COLA_Colaboradores.xlsx` | XLSX | 10.014.768 | 10,0 MB | Cópia Excel de suporte do cadastro. |
| `Relatório Tratado\COLA_Colaboradores_Afastados.xlsx` | XLSX | 288.304 | 288 KB | Cópia Excel de suporte dos afastados. |
| `Relatório Tratado\COLA_Demitidos.xlsx` | XLSX | 3.356.043 | 3,3 MB | Cópia Excel de suporte dos demitidos. |
| `Relatório Tratado\COLA_Ferias_Colaboradores.xlsx` | XLSX | 276.712 | 276 KB | Cópia Excel de suporte das férias. |
| `Relatório Tratado\FOPA_Eventos_por_Colaborador.xlsx` | XLSX | 2.095.984 | 2,1 MB | Cópia Excel de suporte da folha de pagamento. |

---

## 2. Ficha Técnica Detalhada por Arquivo CSV

### 2.1. `COLA_Colaboradores.CSV`
* **Localização Exata:** `Relatorio Original\COLA_Colaboradores.CSV`
* **Tamanho do Arquivo:** 24.908.301 bytes (~24,9 MB)
* **Total de Linhas no Arquivo:** 16.040 linhas
* **Linhas de Metadados / Cabeçalho Superior:** 1 linha.
  * Linha 1 contém o título do relatório: `"Colaboradores"`
  * Linha 2 contém os nomes das 109 colunas
* **Configuração Pandas:** `skiprows=1` (ou `header=1`), `sep=';'`, `encoding='latin1'` (ou `'cp1252'`)
* **Delimitador:** Ponto e vírgula (`;`)
* **Qualificador de Texto:** Aspas duplas (`"`)
* **Encoding Comprovado:** `cp1252` / `iso-8859-1` / `latin1` (bytes como `0xf3` = `ó`, `0xe3` = `ã`, `0xed` = `í`). Decodificação em UTF-8 falha com `UnicodeDecodeError`.
* **Dimensões dos Dados:** 16.038 linhas de dados × 109 colunas
* **Chave Primária (`Matrícula`):**
  * Coluna oficial: `Matrícula` (Índice 3 / 4ª coluna)
  * Valores únicos: 16.038 (100% único, sem duplicatas, sem nulos, sem vazios)
  * Comprimento dos dígitos: 3 dígitos (1 caso), 4 dígitos (7 casos), 5 dígitos (2.335 casos), 6 dígitos (13.695 casos)
  * Zeros à esquerda: **0 casos** com zeros à esquerda no arquivo bruto
  * Natureza: 100% numérico puro (somente dígitos `0-9`)
  * **Alerta Crítico:** Existe também a coluna `Matrícula eSocial` (Índice 2), contendo caracteres de controle (ex: `\x1d`), traços e formatos compostos como `3-289-00277064`. **NUNCA** utilizar `Matrícula eSocial` para joins; utilizar estritamente `Matrícula`.
* **Colunas de Negócio Relevantes:**
  * `Data Admissão` (Índice 7): 16.038 registros não-nulos. 100% no formato `DD/MM/YYYY`. Mínimo: 14/01/1985; Máximo: 22/09/2026.
  * `Cargo` (Índice 71): 334 colaboradores identificados com termos de "APRENDIZ". 0 diretores identificados.
  * `Categoria` (Índice 61): 334 registros com valor exato `'Empregado - Aprendiz'`.
  * `Salário Atual` (Índice 73): Formato brasileiro (`1.800,00`).
  * `Situação` (Índice 107) e `Tipo de Situação` (Índice 106): 13.476 em 'Atividade Normal', 1.510 em 'Afastamento', 1.052 em 'Férias'.
  * `Data de Rescisão` (Índice 108): 100% vazia (colaboradores ativos no cadastro).

<details>
<summary><b>Clique para expandir a lista completa das 109 colunas de COLA_Colaboradores.CSV</b></summary>

1. `Código da Empresa`
2. `Razão Social`
3. `Matrícula eSocial`
4. `Matrícula`
5. `Colaborador`
6. `Nome Social`
7. `ID Pessoa`
8. `Data Admissão`
9. `Modalidade Contratual`
10. `Data de Vencimento ASO`
11. `Data de Vencimento Contrato de Trabalho`
12. `Data de Vencimento Prorrogação de Contrato`
13. `Data Nascimento`
14. `Idade`
15. `Exibir Data de Aniversário`
16. `Sexo`
17. `Gênero`
18. `Estado Civil`
19. `Raça/Cor`
20. `País de Nascimento`
21. `País de Nacionalidade`
22. `Naturalidade`
23. `Tipo da Primeira Filiação`
24. `Primeira Filiação`
25. `Tipo da Segunda Filiação`
26. `Segunda Filiação`
27. `PCD`
28. `Tipo de Deficiência`
29. `Grau de Instrução`
30. `Tipo de Logradouro`
31. `Logradouro`
32. `Número`
33. `Complemento`
34. `Bairro`
35. `CEP`
36. `Código do Município`
37. `Município`
38. `Estado`
39. `UF`
40. `País`
41. `Código Sindicato`
42. `Sindicato`
43. `Primeiro Emprego`
44. `CNPJ`
45. `Código Estabelecimento`
46. `Estabelecimento`
47. `Código Lotação Tributária`
48. `Lotação Tributária`
49. `Código da Unidade Organizacional`
50. `Unidade Organizacional`
51. `Código Nível Organizacional`
52. `Nível Organizacional`
53. `Código do Centro de Custo`
54. `Centro de Custo`
55. `Código Área de Atuação`
56. `Área de Atuação`
57. `Código do Órgão Responsável`
58. `Órgão Responsável`
59. `Código da Natureza Profissional`
60. `Natureza Profissional`
61. `Código Categoria`
62. `Categoria`
63. `Código Vínculo`
64. `Vínculo`
65. `Codigo Jornada`
66. `Jornada`
67. `Carga Horária Mensal`
68. `Carga Horária Semanal`
69. `Código Posição`
70. `Posição`
71. `Código do Cargo`
72. `Cargo`
73. `Tipo de Salário`
74. `Salário Atual`
75. `Código da Função`
76. `Função`
77. `Código do CBO`
78. `CBO`
79. `Telefone`
80. `Celular`
81. `E-mail Corporativo`
82. `E-mail Particular`
83. `CPF/CNPJ`
84. `PIS/PASEP`
85. `RG - Expedidor  - UF`
86. `RG - Data Emissão`
87. `CTPS - Serie - UF/ CTPS Digital`
88. `Título de Eleitor - Número`
89. `Título de Eleitor - Zona`
90. `Título de Eleitor - UF`
91. `Título de Eleitor - Data de Emissão`
92. `Título de Eleitor - Seção`
93. `Habilitação - Número`
94. `Habilitação - Categoria`
95. `Habilitação - Órgão Emissor`
96. `Habilitação - UF`
97. `Habilitação - Data da 1ª habilitação`
98. `Habilitação - Data de Emissão`
99. `Habilitação - Data de Validade`
100. `Certificado Militar - Documento`
101. `Certificado Militar - Categoria`
102. `Certificado Militar - RM / DN / Comar`
103. `Certificado Militar - Número`
104. `Certificado Militar - CSM / OAM`
105. `Certificado Militar - Série`
106. `Tipo de Situação`
107. `Situação`
108. `Data de Rescisão`
109. `Motivo Desligamento`

</details>

---

### 2.2. `COLA_Colaboradores_Afastados.CSV`
* **Localização Exata:** `Relatorio Original\COLA_Colaboradores_Afastados.CSV`
* **Tamanho do Arquivo:** 822.460 bytes (~822 KB)
* **Total de Linhas no Arquivo:** 1.871 linhas (1 cabeçalho + 1.867 dados + quebras)
* **Linhas de Metadados / Cabeçalho Superior:** 0 linhas a pular.
  * Linha 1 contém diretamente os nomes das 27 colunas
* **Configuração Pandas:** `skiprows=0` (ou omitir), `sep=';'`, `encoding='latin1'`
* **Delimitador:** Ponto e vírgula (`;`)
* **Qualificador de Texto:** Aspas duplas (`"`)
* **Encoding Comprovado:** `cp1252` / `latin1`
* **Dimensões dos Dados:** 1.867 linhas de dados × 27 colunas
* **Chave Primária (`Matrícula`):**
  * Coluna: `Matrícula` (Índice 4 / 5ª coluna)
  * Valores únicos: 1.867 (100% único neste arquivo, 0 duplicatas, 0 vazios)
  * Comprimentos: 4 dígitos (1), 5 dígitos (563), 6 dígitos (1.303)
  * Sem zeros à esquerda, numérico puro
* **Colunas de Negócio Relevantes:**
  * `Admissão` (Índice 6): Formato `DD/MM/YYYY`.
  * `Início Afastamento` (Índice 11): 1.867 preenchidos (100%), formato `DD/MM/YYYY`.
  * `Término Afastamento` (Índice 12): 137 preenchidos, 1.730 vazios (afastamentos contínuos/em aberto sem término previsto).
  * `Dias Afastamento` (Índice 13): Total de dias de afastamento (ex: '897', '120', '1999', etc.).
  * `Status Afastamento` (Índice 14): 100% com valor `'Ativo'`.
  * `Tipo Afastamento` (Índice 15):
    - `5- Motivo de Doença, per. superior a 15 dias`
    - `19- Motivo de Doença, per. igual ou inferior a 15 dias`
    - `2- Por Acidente de Trabalho, per. superior a 15 dias`
    - `4- Licença Gestante`
    - `18- Suspensão de Contrato de Trabalho`
    - `1-Sem Remuneração,mantendo direitos de recolhimento`
    - `12- Prorrogação da Licença Maternidade`
    - `76-Aposentadoria por invalidez exceto doença profissional ou acidente de trabalho`
    - `73-Aposentadoria por ivalidez decorrente de acidente de trabalho`
    - `74-Aposentadoria por invalidez decorrente de doença profissional`
  * `Motivo Afastamento` (Índice 16): Causa do afastamento (ex: `'3- Acidente/Doença não relacionada ao trabalho'`, etc.).

<details>
<summary><b>Lista completa das 27 colunas de COLA_Colaboradores_Afastados.CSV</b></summary>

1. `Código Empresa`
2. `Empresa`
3. `Código Estabelecimento`
4. `Estabelecimento`
5. `Matrícula`
6. `Colaborador`
7. `Admissão`
8. `Código Centro de Custo`
9. `Centro de Custo`
10. `Código do Cargo`
11. `Cargo`
12. `Início Afastamento`
13. `Término Afastamento`
14. `Dias Afastamento`
15. `Status Afastamento`
16. `Tipo Afastamento`
17. `Motivo Afastamento`
18. `CID Afastamento`
19. `Benefício de auxílio-doença concedido pela Previdência`
20. `Observação`
21. `Início Atestado`
22. `Término Atestado`
23. `CID Atestado`
24. `Dias Atestado`
25. `Nome do Emitente`
26. `Órgão Emissor`
27. `Inscrição Órgão de Classe`

</details>

---

### 2.3. `COLA_Demitidos.CSV`
* **Localização Exata:** `Relatorio Original\COLA_Demitidos.CSV`
* **Tamanho do Arquivo:** 13.711.091 bytes (~13,7 MB)
* **Total de Linhas no Arquivo:** 16.309 linhas
* **Linhas de Metadados / Cabeçalho Superior:** 0 linhas a pular.
  * Linha 1 contém diretamente os nomes das 42 colunas
* **Configuração Pandas:** `skiprows=0`, `sep=';'`, `encoding='latin1'`
* **Delimitador:** Ponto e vírgula (`;`)
* **Qualificador de Texto:** Aspas duplas (`"`)
* **Encoding Comprovado:** `cp1252` / `latin1`
* **Dimensões dos Dados:** 16.044 linhas de dados × 42 colunas
* **Chave Primária (`Matrícula`):**
  * Coluna do Colaborador: `Matrícula` (Índice 5 / 6ª coluna)
  * Valores únicos: 16.040
  * **Duplicatas de Matrícula:** 4 matrículas duplicadas (8 linhas no total): `324474`, `325212`, `327987`, `333635`.
    - Causa: Múltiplos registros rescisórios históricos, transferências ou recontratações.
    - Exemplo: Matrícula `324474` possui 1 linha com `5 - Transferência sem ônus` e 1 linha com `2 - Sem justa causa do empregador`.
    - Tratativa de Negócio Oficial (Aba Orientações): *Desconsiderar colaboradores com tipo de desligamento 'Transferência sem ônus'*.
  * Comprimento dos dígitos: 2 dígitos (1 caso: '78'), 4 dígitos (2 casos), 5 dígitos (1.686 casos), 6 dígitos (14.355 casos)
  * Sem zeros à esquerda, numérico puro
  * **Atenção:** Existe também a coluna `Matrícula do Gestor` (Índice 38), que contém matrículas compostas com pipes (ex: `'36172 | 330516'`). Não utilizar essa coluna como chave.
* **Colunas de Negócio Relevantes:**
  * `Data Admissão` (Índice 12): 100% no formato `DD/MM/YYYY`.
  * `Data Rescisão` (Índice 13): 100% preenchida (16.044 linhas), formato `DD/MM/YYYY`.
  * `Tipo de Rescisão` (Índice 14): Motivo formal (ex: `'2 - Sem justa causa do empregador'`, `'5 - Transferência sem ônus'`, `'1 - Com justa causa'`, etc.).
  * `Salário` (Índice 11): Salário no momento do desligamento.

<details>
<summary><b>Lista completa das 42 colunas de COLA_Demitidos.CSV</b></summary>

1. `Empresa`
2. `Razão Social`
3. `Código Estabelecimento`
4. `Estabelecimento`
5. `CNPJ`
6. `Matrícula`
7. `Colaborador`
8. `Número do PIS`
9. `UF`
10. `Código Cargo`
11. `Cargo`
12. `Salário`
13. `Data Admissão`
14. `Data Rescisão`
15. `Tipo de Rescisão`
16. `Motivo do Desligamento`
17. `Justificativa do Desligamento - AutoAtendimento`
18. `Data do Aviso Prévio`
19. `Projeção do Aviso Prévio`
20. `Código Posição`
21. `Posição`
22. `Código Centro de Custo`
23. `Centro de Custo`
24. `Código Unidade Organizacional`
25. `Unidade Organizacional`
26. `Código Natureza Profissional`
27. `Natureza Profissional`
28. `Situação`
29. `CPF`
30. `Sexo`
31. `Data Nascimento`
32. `Escolaridade`
33. `Raça/Cor`
34. `Código Vínculo`
35. `Vínculo`
36. `Código da Categoria`
37. `Categoria`
38. `PCD`
39. `Matrícula do Gestor`
40. `Nome do Gestor`
41. `Código do Sindicato `
42. `Sindicato`

</details>

---

### 2.4. `COLA_Ferias_Colaboradores.CSV`
* **Localização Exata:** `Relatorio Original\COLA_Ferias_Colaboradores.CSV`
* **Tamanho do Arquivo:** 561.339 bytes (~561 KB)
* **Total de Linhas no Arquivo:** 2.143 linhas
* **Linhas de Metadados / Cabeçalho Superior:** 1 linha a pular.
  * Linha 1 contém o título: `"Férias de Colaboradores"`
  * Linha 2 contém os nomes das 19 colunas
* **Configuração Pandas:** `skiprows=1`, `sep=';'`, `encoding='latin1'`
* **Delimitador:** Ponto e vírgula (`;`)
* **Qualificador de Texto:** Aspas duplas (`"`)
* **Encoding Comprovado:** `cp1252` / `latin1`
* **Dimensões dos Dados:** 2.141 linhas de dados × 19 colunas
* **Chave Primária (`Matrícula`):**
  * Coluna: `Matrícula` (Índice 2 / 3ª coluna)
  * Valores únicos: 2.133
  * **Duplicatas de Matrícula:** 8 matrículas duplicadas (16 linhas no total) devido a férias fracionadas em 2 períodos de gozo (ex: 15 dias em agosto e 15 dias em setembro).
    - Exemplo: Matrícula `20357` (período 1: 03/08/2026 a 17/08/2026; período 2: 01/09/2026 a 15/09/2026).
    - Tratativa de Negócio Oficial (Aba Orientações): *Manter somente o período que compreende o mês de referência da auditoria/compra*.
  * Comprimento dos dígitos: 5 dígitos (320 casos), 6 dígitos (1.821 casos)
  * Sem zeros à esquerda, numérico puro
* **Colunas de Negócio Relevantes:**
  * `Início do Período de Gozo` (Índice 9): 100% preenchida (2.141 linhas), formato `DD/MM/YYYY`.
  * `Fim do Período de Gozo` (Índice 10): 100% preenchida (2.141 linhas), formato `DD/MM/YYYY`.
  * `Dias Gozo` (Índice 11): Quantidade de dias de férias (ex: '15', '30', '10').
  * `Retorno` (Índice 16): Data de retorno ao trabalho (`DD/MM/YYYY`).
  * `Status` (Índice 18): Situação das férias.

<details>
<summary><b>Lista completa das 19 colunas de COLA_Ferias_Colaboradores.CSV</b></summary>

1. `Código Empresa`
2. `Razão Social`
3. `Matrícula`
4. `Colaborador`
5. `CPF`
6. `Unidade Organizacional`
7. `Início do Período Aquisitivo`
8. `Fim do Período Aquisitivo`
9. `Data Limite`
10. `Início do Período de Gozo`
11. `Fim do Período de Gozo`
12. `Dias Gozo`
13. `Dias a desconsiderar`
14. `Abono`
15. `Parcela`
16. `1ª Parc. 13º Salário`
17. `Retorno`
18. `Data de Pagamento`
19. `Status`

</details>

---

### 2.5. `FOPA_Eventos_por_Colaborador - 22092026_091703_796705.csv`
* **Localização Exata:** `Relatorio Original\FOPA_Eventos_por_Colaborador - 22092026_091703_796705.csv`
* **Padrão Glob de Correspondência:** `Relatorio Original\FOPA_Eventos_por_*.csv`
* **Tamanho do Arquivo:** 4.245.243 bytes (~4,2 MB)
* **Total de Linhas no Arquivo:** 15.917 linhas
* **Linhas de Metadados no Topo:** 3 linhas completas a pular:
  * Linha 1: `Eventos Selecionados - Coluna;;;;;;;;;;;;;;;;;;;;;;`
  * Linha 2: `Folha ;1;;;;;;;;;;;;;;;;;;;;;`
  * Linha 3: `Competência;set/26;;;;;;;;;;;;;;;;;;;;;`
  * Linha 4: Cabeçalho com os nomes das 23 colunas
* **Configuração Pandas:** `skiprows=3`, `sep=';'`, `encoding='latin1'`
* **Delimitador:** Ponto e vírgula (`;`)
* **Encoding Comprovado:** `cp1252` / `latin1`
* **Dimensões dos Dados:** 15.913 linhas de dados × 23 colunas
* **Chave Primária (`Matrícula`):**
  * Coluna: `Matrícula` (Índice 6 / 7ª coluna)
  * Valores únicos: 15.913 (100% único, sem duplicatas, sem nulos)
  * Comprimento dos dígitos: 3 dígitos (1), 4 dígitos (7), 5 dígitos (2.332), 6 dígitos (13.573)
  * Sem zeros à esquerda, numérico puro
  * **Interseção:** 100% dos colaboradores da FOPA (15.913) constam no cadastro `COLA_Colaboradores.CSV`.
* **Estrutura Fundamental dos Eventos (Descoberta Crítica):**
  * Ao contrário de arquivos transacionais em linhas, o arquivo FOPA já se encontra **PIVOTADO EM COLUNAS**.
  * Cada colaborador possui exatamente 1 linha, e cada evento da folha é uma coluna própria com o código e a descrição no nome da coluna.
  * Colunas de eventos mapeadas e verificadas:
    1. `1 - Salario` (Índice 15): 13.609 com valores > 0, 2.173 com traço `'-'` (afastados/licenciados).
    2. `6280 - Desc Cartão Benefícios Especiais` (Índice 16): 1.088 colaboradores com desconto em folha; 14.649 com traço `'-'`.
    3. `10990 - Liquido a Receber` (Índice 17): 13.595 colaboradores com valor positivo; 2.268 com traço `'-'`.
    4. `11010 - Salario Mensal` (Índice 18): 15.899 colaboradores com salário contratual fixo; apenas 8 com `'-'`. Esta coluna é a base para a **Margem de 20%**.
    5. `29613 - Ben Esp - Valor Informado para Desconto` (Índice 19): 1.081 colaboradores com valor informado para desconto; 14.648 com `'-'`.
    6. `29621 - Ben Esp - Saldo Devedor Beneficios Especiais` (Índice 20): 75 colaboradores com saldo devedor do mês atual; 15.834 com `'-'`.
    7. `29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ` (Índice 21): 72 colaboradores com saldo do mês anterior (atenção ao espaço final no nome da coluna!); 15.837 com `'-'`.
    8. `29624 - Ben Esp - Custo Empresa Beneficios Especiais` (Índice 22): 211 colaboradores com custo empresa (5% Giga / loja 400); 15.695 com `'-'`.
* **Formatação Monetária e Numérica:**
  * Padrão brasileiro estrito: ponto (`.`) como separador de milhar e vírgula (`,`) como separador decimal (ex: `'5.582,44'`, `'317,06'`).
  * O traço literal (`'-'`) é utilizado uniformemente para indicar valor zero (`0,00`) ou ausência de lançamento.
  * **Não existem valores negativos** (não há sinal de menos `-` precedendo números).
  * Regra de conversão para float em Python:
    ```python
    def parse_br_currency(val):
        if val is None or str(val).strip() in ('', '-'):
            return 0.0
        return float(str(val).strip().replace('.', '').replace(',', '.'))
    ```

<details>
<summary><b>Lista completa das 23 colunas de FOPA_Eventos_por_Colaborador</b></summary>

1. `Código da Empresa`
2. `Nome da Empresa`
3. `Código Estabelecimento`
4. `Estabelecimento`
5. `Código Centro Custo`
6. `Centro Custo`
7. `Matrícula`
8. `Nome do Colaborador`
9. `ID Colaborador`
10. `Data Admissão`
11. `Código Cargo`
12. `Cargo`
13. `Situação`
14. `Tipo Situação`
15. `Data Rescisão`
16. `1 - Salario`
17. `6280 - Desc Cartão Benefícios Especiais`
18. `10990 - Liquido a Receber`
19. `11010 - Salario Mensal`
20. `29613 - Ben Esp - Valor Informado para Desconto`
21. `29621 - Ben Esp - Saldo Devedor Beneficios Especiais`
22. `29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ` *(com espaço no final)*
23. `29624 - Ben Esp - Custo Empresa Beneficios Especiais`

</details>

---

## 3. Modelo Excel de Referência (`Conferência Good Card_Base.xlsx`)

A inspeção do arquivo modelo mestre de 23 MB revelou sua exata composição de abas e colunas:

### 3.1. Abas Existentes
1. `Orientações` (Instruções detalhadas de negócio e regras de conferência)
2. `Conferência` (Tabela principal com 38 colunas e 15.913 linhas de colaboradores)
3. `Valores Importados` (Conciliação do evento 29613 com 11 colunas e 1.334 registros)
4. `FOPA_Eventos_por_Colaborador - ` (Aba espelho dos dados da folha)
5. `Desligado` (Aba espelho dos demitidos)
6. `Afastamento` (Aba espelho dos afastados)
7. `Férias` (Aba espelho das férias)
8. `Nomina` (Aba espelho do cadastro ativo)
9. `Evidência`
10. `Lista` (Domínio oficial dos textos da coluna `Tratativa`)

### 3.2. Estrutura Exata das 38 Colunas da Aba `Conferência`
As 38 colunas correspondem às seguintes fontes e fórmulas extraídas diretamente do modelo:

| Nº | Nome Exato da Coluna | Origem dos Dados / Regra de Cálculo |
| :---: | :--- | :--- |
| 1 | `Código da Empresa` | FOPA (Col 0) |
| 2 | `Nome da Empresa` | FOPA (Col 1) |
| 3 | `Código Estabelecimento` | FOPA (Col 2) |
| 4 | `Estabelecimento` | FOPA (Col 3) |
| 5 | `Código Centro Custo` | FOPA (Col 4) |
| 6 | `Centro Custo` | FOPA (Col 5) |
| 7 | `Matrícula` | FOPA (Col 6) — Chave Universal |
| 8 | `Nome do Colaborador` | FOPA (Col 7) |
| 9 | `ID Colaborador` | FOPA (Col 8) |
| 10 | `Data Admissão` | FOPA (Col 9) |
| 11 | `Código Cargo` | FOPA (Col 10) |
| 12 | `Cargo` | FOPA (Col 11) |
| 13 | `1 - Salario` | FOPA (Col 15) |
| 14 | `11010 - Salario Mensal` | FOPA (Col 18) |
| 15 | `10990 - Liquido a Receber` | FOPA (Col 17) |
| 16 | `29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ` | FOPA (Col 21) |
| 17 | `29613 - Ben Esp - Valor Informado para Desconto` | FOPA (Col 19) |
| 18 | `Total Saldo Devedor` | Fórmula: `Col 16 (29622) + Col 17 (29613)` |
| 19 | `Margem 20% do Salário` | Fórmula: `Col 14 (11010 Salario Mensal) * 0.20` |
| 20 | `Total Saldo Devedor " Dentro da Margem" ou "Fora da Margem"` | Se `round(Total Saldo Devedor, 2) <= round(Margem 20%, 2)`: `"Dentro da Margem"`; senão: `"Fora da Margem"` |
| 21 | `6280 - Desc Cartão Benefícios Especiais` | FOPA (Col 16) |
| 22 | `Desconto x Margem de 20%` | Se `round(6280, 2) <= round(Margem 20%, 2)`: `"Dentro da Margem"`; senão: `"Fora da Margem"` |
| 23 | `Saldo Devedor x Evt 6280` | Comparação: `round(29624 + 6280, 2) == round(Total Saldo Devedor, 2)` (ou `29621 + 6280 == Total Saldo Devedor`) |
| 24 | `Total Saldo Devedor - Evt 6280 x Evt 29621` | Comparação de batimento entre saldo devedor e desconto realizado |
| 25 | `29621 - Ben Esp - Saldo Devedor Beneficios Especiais` | FOPA (Col 20) |
| 26 | `29624 - Ben Esp - Custo Empresa Beneficios Especiais` | FOPA (Col 22) |
| 27 | `5% Saldo devedor` | Fórmula: `Col 18 (Total Saldo Devedor) * 0.05` |
| 28 | `Cálculo 5% x 29624` | Fórmula: `Col 26 (29624) - Col 27 (5% Saldo devedor)` |
| 29 | `Início Afastamento` | Procv em `COLA_Colaboradores_Afastados` pela `Matrícula` |
| 30 | `Final Afastamento` | Procv em `COLA_Colaboradores_Afastados` pela `Matrícula` (`Término Afastamento`) |
| 31 | `Tipo do Afastamento` | Procv em `COLA_Colaboradores_Afastados` pela `Matrícula` |
| 32 | `Início Férias` | Procv em `COLA_Ferias_Colaboradores` pela `Matrícula` (`Início do Período de Gozo`) |
| 33 | `Final Férias` | Procv em `COLA_Ferias_Colaboradores` pela `Matrícula` (`Fim do Período de Gozo`) |
| 34 | `Qtd de dias` | Procv em `COLA_Ferias_Colaboradores` pela `Matrícula` (`Dias Gozo`) |
| 35 | `Desligamento` | Procv em `COLA_Demitidos` pela `Matrícula` (`Data Rescisão`) |
| 36 | `Situação` | Procv em `COLA_Colaboradores` pela `Matrícula` (`Situação`) |
| 37 | `Obs. Conferência` | Observação descritiva da regra |
| 38 | `Tratativa` | Classificação formal padronizada (domínio da aba `Lista`) |

### 3.3. Domínio Oficial da Coluna `Tratativa` (Aba `Lista`)
Valores válidos e exatos para a coluna `Tratativa`:
1. `Ok - Sem saldo devedor e sem desconto`
2. `Ok - Liquidou todo o saldo devedor `
3. `Ok - Suspensão de Contrato`
4. `Afastamento - Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 `
5. `Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621`
6. `Férias- Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 `
7. `Verificar com quality - Não estava em situação de afastamento e não houve o desconto`
8. `Aprendiz - Verificar`
9. `Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)`

### 3.4. Estrutura Exata das 11 Colunas da Aba `Valores Importados`
1. `Emp` (Código da Empresa)
2. `Matrícula` (Chave de conciliação)
3. `Evento` (Fixado em 29613)
4. `Mês` (92026 = 09/2026)
5. `Valor` (Valor importado pelo arquivo Good Card)
6. `Folha` (Valor apurado na FOPA para o evento 29613)
7. `Importado x Folha` (Validação booleana: `Valor == Folha`)
8. `Desligado` (Data de rescisão via Demitidos, caso conste)
9. `Tipo de Rescisão` (Tipo de rescisão via Demitidos)
10. `Obs. Conferência` (Observações adicionais, ex: transferências)
11. `Tratativa` (Tratativa da divergência)

---

## 4. Análise de Relacionamento (Joins), Inconsistências e Casos de Borda

### 4.1. Cruzamento Universal por `Matrícula`
* **Tipo de Dado:** `str` (string).
* **Ausência de Zeros à Esquerda:** Em todos os CSVs fornecidos, as matrículas foram extraídas como números sem zeros à esquerda (ex: `'935'`, `'1000'`, `'11388'`, `'331372'`). Na planilha Excel de base, a coluna `Matrícula` também foi formatada/armazenada sem zeros à esquerda.
* **Cobertura:** 
  * 100% das 15.913 matrículas da FOPA existem em `COLA_Colaboradores.CSV`.
  * 1.492 matrículas da FOPA estão em `COLA_Colaboradores_Afastados.CSV`.
  * 2.117 matrículas da FOPA estão em `COLA_Ferias_Colaboradores.CSV`.
  * 0 matrículas da FOPA estão em `COLA_Demitidos.CSV` (pois a folha FOPA de set/26 já não processa desligados anteriores).

### 4.2. Tratamento de Duplicidades nas Tabelas Secundárias
1. **Demitidos (`COLA_Demitidos.CSV`):**
   * Existem 4 matrículas duplicadas no arquivo bruto (8 registros).
   * **Regra de Deduplicação:** Desconsiderar registros cujo `Tipo de Rescisão` contenha `'Transferência sem ônus'`. Se persistir duplicidade, ordenar por `Data Rescisão` decrescente e reter a rescisão mais recente.
2. **Férias (`COLA_Ferias_Colaboradores.CSV`):**
   * Existem 8 matrículas com múltiplos períodos de gozo fracionados.
   * **Regra de Deduplicação:** Filtrar para manter apenas o período cujo intervalo (`Início do Período de Gozo` a `Fim do Período de Gozo`) coincida ou intercepte o mês de competência da auditoria (setembro/2026).
3. **Afastados (`COLA_Colaboradores_Afastados.CSV`):**
   * **Regra de Deduplicação:** Conforme a aba Orientações, caso existam múltiplos períodos de afastamento para um mesmo colaborador, reter aquele com a data mais extensa. Além disso, desconsiderar afastamentos finalizados antes do primeiro dia do mês de competência (01/09/2026).

### 4.3. Regras de Negócio Específicas
1. **Regra 1 - Demitidos:**
   * Cruzamento com `COLA_Demitidos.CSV`. Todo colaborador presente em `Valores Importados` que possua data de desligamento deve ser sinalizado para cancelamento de cartão e conferência de verbas rescisórias.
2. **Regra 2 - Afastados > 15 dias:**
   * No arquivo `COLA_Colaboradores_Afastados.CSV`, avaliar `Dias Afastamento > 15` e tipos de afastamento por doença superior a 15 dias. Quando o colaborador não possuir saldo suficiente na folha, o saldo devedor deve ser encaminhado para o evento 29621.
3. **Regra 3 - Novos Admitidos > 90 dias e Exclusões:**
   * Filtrar admissões com base em `Data Admissão` e data base da compra.
   * Excluir Aprendizes (`Categoria == 'Empregado - Aprendiz'` ou `Cargo` contendo `'APRENDIZ'`) e Diretores.
4. **Regra 4 - Desconto em Folha e Margem de 20%:**
   * Margem calculada sobre `11010 - Salario Mensal` (20%).
   * Saldo Devedor Total = Evento 29622 (mês anterior) + Evento 29613 (informado do mês).
   * Se `Saldo Devedor Total <= Margem 20%`: deve ser descontado integralmente no Evento 6280, e o Evento 29621 deve ficar zerado (`'-'`).
   * Se `Saldo Devedor Total > Margem 20%`: deve descontar no Evento 6280 até o limite da Margem de 20%, e a diferença excedente deve constar no Evento 29621 para transporte ao mês seguinte.
   * Exceção Giga / Estabelecimento 400: Parâmetro especial de Custo Empresa de 5% sobre o saldo devedor (Evento 29624).

---

## 5. Resumo das Recomendações Técnicas para o Desenvolvimento

1. **Leitura e Ingestão de Dados:**
   * Sempre usar `encoding='latin1'` e `sep=';'`.
   * Tratar explicitamente `skiprows=1` para `COLA_Colaboradores.CSV` e `COLA_Ferias_Colaboradores.CSV`.
   * Tratar explicitamente `skiprows=3` para arquivos no padrão `FOPA_Eventos_por_*.csv`.
   * Sempre carregar `Matrícula` com `dtype=str` e aplicar `.str.strip()`.
2. **Normalização de Nomes de Coluna:**
   * A coluna `29622 - Ben Esp - Bsc Saldo Dev Beneficios Especiais Mes Ant ` possui um espaço no final no arquivo original. Recomenda-se aplicar `.str.strip()` em todos os nomes de colunas após a leitura.
3. **Conversão de Valores:**
   * Tratar a string `'-'` como `0.0`.
   * Converter moedas brasileiras substituindo ponto por nada e vírgula por ponto antes de converter para float.
4. **Exportação Excel Sem Fórmulas:**
   * Escrever todos os valores calculados como valores estáticos no OpenPyXL.
   * Freeze Panes fixado estritamente em `'A2'`.
   * Centralizar colunas de códigos, datas e eventos; alinhar textos e tratativas à esquerda; alinhar valores monetários à direita (ou centralizados conforme padrão Cencosud).
   * Destacar linhas divergentes (casos fora da margem, não conformidades de saldo ou cancelamentos pendentes) com estilo visual vermelho.
