# Project: Automação da Auditoria "04 - GOOD CARD" (Cencosud)

## Architecture
- **Paradigma**: Processamento analítico 100% em memória utilizando Pandas e geração estática via OpenPyXL.
- **Módulos Principais**:
  - `config.py`: Definições globais, mapeamento oficial das 38 colunas da Conferência e 11 colunas de Valores Importados, paleta corporativa de cores Cencosud, códigos de eventos e tolerâncias financeiras.
  - `data_loader.py`: Ingestão resiliente dos 5 CSVs originais do RH e da base importada do fornecedor, detecção de encoding (`latin1`/`cp1252`), separador (`;`), tratamento de cabeçalhos (`skiprows`), normalização estrita da `Matrícula` (string sem perda de zeros/sem sufixo `.0`) e parsing numérico de valores monetários brasileiros (`5.582,44` e `"-"` -> `0.00`).
  - `business_rules.py`: Cruzamentos relacionais (left joins por `Matrícula`), aplicação das 4 Regras de Negócio do Good Card (Desligados, Afastados > 15 dias, Admitidos > 90 dias / Inelegibilidade de Aprendiz e Diretor, Desconto em Folha e Margem de 20% com regra subsidiada de 5% para Giga/Estab 400), conciliação da aba Valores Importados (evento 29613) e apuração dos KPIs consolidados.
  - `excel_exporter.py`: Geração do workbook Excel `.xlsx` estático (sem fórmulas) contendo exatamente as 4 abas requeridas (`Orientações`, `Resumo da Auditoria`, `Conferência`, `Valores Importados`), estilos corporativos, freeze panes em `'A2'`, alinhamento centralizado para códigos/datas/números, alinhamento à esquerda para nomes/tratativas e destaque visual em vermelho suave (`#FFC7CE` / `#9C0006`) para linhas/células com divergência apontada.
  - `main.py`: Ponto de entrada CLI orquestrador, controle de fluxo, telemetria de execução, validação de integridade e tratamento robusto de exceções.
- **Data Flow**:
  1. CSVs em `Relatorio Original/` + Modelo Base -> `data_loader.py` -> DFs em Memória
  2. DFs em Memória -> `business_rules.py` -> `df_conferencia` (38 colunas) + `df_valores_importados` (11 colunas) + `kpis_dict`
  3. DFs Auditados -> `excel_exporter.py` -> `Conferência Good Card_Gerado.xlsx` (4 abas estáticas, formatadas)

## Code Layout
```
C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\04 - GOOD CARD/
├── config.py
├── data_loader.py
├── business_rules.py
├── excel_exporter.py
├── main.py
├── tests/
│   ├── e2e/
│   │   ├── run_all_e2e_tests.py
│   │   ├── test_tier1_features.py
│   │   ├── test_tier2_boundaries.py
│   │   ├── test_tier3_combinations.py
│   │   └── test_tier4_workloads.py
│   └── unit/
│       ├── test_data_loader.py
│       ├── test_business_rules.py
│       └── test_excel_exporter.py
├── Relatorio Original/          (CSVs brutos de entrada)
├── Conferência Good Card_Base.xlsx (Modelo mestre)
└── Relatório Tratado/           (Destino da planilha gerada)
```

## Feature Inventory
| # | Feature | Description | Milestone | Source |
|---|---|---|---|---|
| 1 | Ingestão Resiliente de CSVs | Leitura com `latin1`, separador `;`, `skiprows` parametrizado (FOPA: 3, Férias: 1, Colaboradores: 1, Demitidos/Afastados: 0) | M1 | ORIGINAL_REQUEST § R1, survey |
| 2 | Normalização Universal de Matrícula | Garantir `Matrícula` como string limpa sem perda de zeros e sem sufixo `.0` em todas as bases | M1 | ORIGINAL_REQUEST § R1, survey |
| 3 | Parsing Numérico / Moeda BR | Conversão de strings financeiras (`5.582,44`, `"-"`, `" -"` para float `0.00`), remoção de NaNs | M1 | survey csv |
| 4 | Deduplicação e Pré-filtros de RH | Desconsiderar "Transferência sem ônus" em Demitidos; filtrar afastamentos findos antes do mês; férias no mês | M1 | survey spec & csv |
| 5 | Chave de Join Universal | Left join partindo de `FOPA` associando cadastral, afastamentos, férias e demitidos | M2 | ORIGINAL_REQUEST § R1 |
| 6 | Regra 1: Desligados | Cancelamento de cartão p/ demitidos, apuração de saldo devedor remanescente e tratativa | M2 | ORIGINAL_REQUEST § R1 |
| 7 | Regra 2: Afastados > 15 dias | Solicitação de bloqueio p/ afastamento > 15 dias, validação do evento 29621 na folha | M2 | ORIGINAL_REQUEST § R1 |
| 8 | Regra 3: Novos Admitidos > 90 dias | Elegibilidade pós-experiência (90 dias), bloqueio estrito de cargos de Aprendiz e Diretor | M2 | ORIGINAL_REQUEST § R1 |
| 9 | Regra 4: Desconto em Folha e Margem 20% | Apuração de Total Saldo Devedor (`29622 + 29613`), margem 20% do salário (`11010`), status de margem e eventos | M2 | ORIGINAL_REQUEST § R1 |
| 10 | Regra 4: Custo Empresa 5% Giga/Estab 400 | Desconto subsidiado de 5% empresa (`29624`) e 95% colaborador (`6280`) no Estab 400 e rede Giga | M2 | survey spec |
| 11 | Conciliação Valores Importados | Conciliar lote do fornecedor (11 colunas) com evento 29613 da folha e apontar desligados | M2 | ORIGINAL_REQUEST § R2, survey |
| 12 | Apuração de KPIs Executivos | Total auditados, dentro/fora da margem, liquidados, carregados para 29621, índice de conformidade | M2 | ORIGINAL_REQUEST § R2, survey |
| 13 | Aba "Orientações" | Geração fiel da aba de instruções em 26 linhas com larguras, fontes e cores do modelo base | M3 | ORIGINAL_REQUEST § R2, survey |
| 14 | Aba "Resumo da Auditoria" | Dashboard executivo com banner azul (`1F497D`), cartões de KPIs e tabela de distribuição de pareceres | M3 | ORIGINAL_REQUEST § R2, survey |
| 15 | Aba "Conferência" (38 Colunas) | Geração estática das 38 colunas exatas, preenchimento ordenado, formatações de moeda e datas | M3 | ORIGINAL_REQUEST § R2, survey |
| 16 | Aba "Valores Importados" (11 Colunas) | Geração estática das 11 colunas exatas da conciliação do evento 29613 | M3 | ORIGINAL_REQUEST § R2, survey |
| 17 | Usabilidade Excel (Freeze Panes 'A2') | Fixação de painéis estritamente na linha 1 / célula `'A2'` nas abas de dados | M3 | ORIGINAL_REQUEST § R2 |
| 18 | Destaque Visual de Divergências | Aplicação de estilo suave vermelho (`#FFC7CE` / `#9C0006`) nas linhas com inconsistência em Tratativa | M3 | ORIGINAL_REQUEST § R2, survey |
| 19 | Saída Estática Sem Fórmulas | Eliminação de fórmulas (`=VLOOKUP`, `=IF`), exportando dados calculados 100% estáticos | M3 | ORIGINAL_REQUEST § R2, survey |
| 20 | CLI Runner & Orquestração em Memória | Script `main.py` com telemetria, zero CSVs temporários em disco e tratamento de exceções | M3 | ORIGINAL_REQUEST § R2, Acceptance Criteria |
| 21 | Suíte de Testes E2E (Tiers 1-4) | Suíte de testes independente cobrindo isolamento, limites, combinações e cenários reais | M4 | Dual Track Testing |
| 22 | Hardening Adversarial (Tier 5) | Testes estressantes de casos de borda, dados corrompidos, ausências e tolerâncias | M5 | Dual Track Testing |

## Milestones
| # | Name | Scope | Dependencies | Status |
|---|---|---|---|---|
| M1 | Ingestão e Normalização (`data_loader.py` + `config.py`) | Configurações oficiais, ingestão dos 5 CSVs e template base, tratamento de encoding, separador, skiprows, normalização de matrícula e parsing monetário | Survey | DONE |
| M2 | Motor de Regras de Negócio (`business_rules.py`) | Joins relacionais, Regras 1 a 4, cálculo de margem 20%, regra de 5% Giga/Estab 400, conciliação de Valores Importados e agregador de KPIs | M1 | DONE |
| M3 | Exportador Excel e Runner CLI (`excel_exporter.py` + `main.py`) | OpenPyXL com 4 abas estáticas, 38 colunas de Conferência, 11 de Valores Importados, dashboard de Resumo, freeze panes 'A2', destaque vermelho, CLI runner | M2 | PLANNED |
| M4 | Validação E2E Test Suite (Tiers 1-4) | Execução integral da suíte E2E do track de testes, aprovação em 100% dos testes e conformidade funcional | M3, TEST_READY | PLANNED |
| M5 | Adversarial Coverage Hardening (Tier 5) | Testes adversariais com dados sintéticos extremos, validação estrita de integridade forense | M4 | PLANNED |

## Interface Contracts
### `data_loader.py` ↔ `business_rules.py`
- `load_all_data(data_dir: str, template_path: str) -> dict[str, pd.DataFrame]`
  - Retorna dicionário contendo:
    - `'fopa'`: DataFrame com 15.913 linhas (ou total do mês), eventos numéricos parseados (`float64`), `Matrícula` string limpa.
    - `'nomina'`: DataFrame cadastral com `Matrícula`, `Cargo`, `Situação`, `Data Admissão`.
    - `'afastados'`: DataFrame deduplicado (período mais longo, início pós-mês auditado), com `Matrícula`, `Início Afastamento`, `Término Afastamento`, `Tipo Afastamento`.
    - `'demitidos'`: DataFrame deduplicado (rescisão mais recente, excluindo transferência sem ônus), com `Matrícula`, `Data Rescisão`, `Tipo de Rescisão`.
    - `'ferias'`: DataFrame deduplicado (período do mês), com `Matrícula`, `Início Férias`, `Final Férias`, `Qtd de dias`.
    - `'valores_importados'`: DataFrame com `Emp`, `Matrícula`, `Evento`, `Mês`, `Valor`.

### `business_rules.py` ↔ `excel_exporter.py`
- `execute_audit(data_dict: dict[str, pd.DataFrame]) -> tuple[pd.DataFrame, pd.DataFrame, dict]`
  - Retorna:
    - `df_conferencia`: DataFrame estruturado com exatamente as 38 colunas na ordem oficial (Colunas A até AL), incluindo todas as colunas de dados, cálculos matemáticos, `Obs. Conferência` e `Tratativa`.
    - `df_valores_importados`: DataFrame com exatamente as 11 colunas da aba Valores Importados (Colunas A até K), incluindo `Importado x Folha` e tratativas de conciliação.
    - `kpis_dict`: Dicionário com KPIs executivos, métricas de conformidade e tabela de distribuição de status para a aba `Resumo da Auditoria`.

### `excel_exporter.py` ↔ `main.py`
- `export_audit_excel(df_conferencia: pd.DataFrame, df_valores_importados: pd.DataFrame, kpis: dict, output_path: str) -> str`
  - Gera arquivo `.xlsx` sem fórmulas com as 4 abas formatadas, salva em disco e retorna o caminho absoluto verificado.
