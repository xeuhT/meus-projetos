# Test Infrastructure Specification: "04 - GOOD CARD"

**Document Version:** 1.0.0  
**Status:** Approved & Implemented  
**Environment:** Python 3.14.2 | Pandas 3.0.5 | OpenPyXL 3.1.5 | Windows 11  
**Project Root:** `C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\04 - GOOD CARD`

---

## 1. Executive Summary & Testing Philosophy

The test suite for the **04 - GOOD CARD** audit automation project implements an **opaque-box end-to-end (E2E) testing framework** designed to validate every business rule, edge case, multi-rule combination, and real-world workload against authoritative specifications.

The testing architecture follows the **Dual-Track Delivery Model**:
1. **Implementation Track**: Developing `config.py`, `data_loader.py` (M1), `business_rules.py` (M2), and `excel_exporter.py` / `main.py` (M3).
2. **E2E Testing Track**: Independent opaque-box test suite organized across 4 hierarchical tiers with >=5 test cases per feature, ensuring progressive testability and authoritative expected output derivation.

---

## 2. Test Architecture & Tier Structure

```
tests/e2e/
├── run_all_e2e_tests.py           # Unified test runner (Exit Code 0 on success)
├── test_tier1_features.py         # Tier 1: Functional coverage (>=5 tests per feature)
├── test_tier2_boundaries.py       # Tier 2: Boundary, corner, and threshold cases
├── test_tier3_combinations.py     # Tier 3: Multi-rule interactions and edge overlaps
├── test_tier4_workloads.py        # Tier 4: E2E pipeline workloads and workbook audit
└── oracle.py                      # Authoritative Reference Oracle & Contract Validator
```

### 2.1 Tier Definitions

| Tier | File | Target Scope | Methodology |
|:---:|---|---|---|
| **Tier 1** | `test_tier1_features.py` | Individual functional features (Features 1 to 20 in `PROJECT.md`) | Minimum 5 isolated test cases per feature verifying happy path, invalid inputs, and error handling. |
| **Tier 2** | `test_tier2_boundaries.py` | Mathematical and temporal boundaries | Threshold tests: exactly 20% vs 20% + 0.01 centavo, 15 vs 16 days afastamento, 90 vs 91 days admissão, zeros, `-`, NaNs, matricula padding. |
| **Tier 3** | `test_tier3_combinations.py` | Cross-rule interactions | Combinatorial scenarios: demitido + afastado, ferias + estouro de margem, aprendiz com saldo devedor, estab 400 + afastamento, rescisão sem ônus. |
| **Tier 4** | `test_tier4_workloads.py` | End-to-end data pipelines & Excel audit | Execution over real CSVs in `Relatorio Original/`, static workbook generation, 4 sheets, 38/11 columns, freeze panes 'A2', red divergence styling. |

---

## 3. Feature Inventory Mapping (Features 1 to 20)

Every feature defined in `PROJECT.md § Feature Inventory` is mapped to its primary test suite, tier, and specific test cases:

| Feature ID | Feature Name | Milestone | Tier | Primary Test Class / Module | Test Count |
|:---:|---|:---:|:---:|---|:---:|
| **F01** | Ingestão Resiliente de CSVs | M1 | Tier 1 | `TestFeature01CsvLoading` | >= 5 |
| **F02** | Normalização Universal de Matrícula | M1 | Tier 1 | `TestFeature02MatriculaTypeSafety` | >= 5 |
| **F03** | Parsing Numérico / Moeda BR | M1 | Tier 1 | `TestFeature03CurrencyParsing` | >= 5 |
| **F04** | Deduplicação e Pré-filtros de RH | M1 | Tier 1 | `TestFeature04RhDeduplication` | >= 5 |
| **F05** | Chave de Join Universal | M2 | Tier 1 | `TestFeature05UniversalJoin` | >= 5 |
| **F06** | Regra 1: Desligados | M2 | Tier 1 | `TestFeature06Rule1Desligados` | >= 5 |
| **F07** | Regra 2: Afastados > 15 dias | M2 | Tier 1 | `TestFeature07Rule2Afastados` | >= 5 |
| **F08** | Regra 3: Novos Admitidos > 90 dias / Inelegíveis | M2 | Tier 1 | `TestFeature08Rule3AdmitidosInelegiveis` | >= 5 |
| **F09** | Regra 4: Desconto em Folha e Margem 20% | M2 | Tier 1 | `TestFeature09Rule4Margem20SaldoDevedor` | >= 5 |
| **F10** | Regra 4: Custo Empresa 5% Giga / Estab 400 | M2 | Tier 1 | `TestFeature10Rule4CustoEmpresa5Percent` | >= 5 |
| **F11** | Conciliação Valores Importados (Evento 29613) | M2 | Tier 1 | `TestFeature11ConciliacaoValoresImportados` | >= 5 |
| **F12** | Apuração de KPIs Executivos | M2 | Tier 1 | `TestFeature12KpiExecutiveMetrics` | >= 5 |
| **F13** | Aba "Orientações" | M3 | Tier 1 | `TestFeature13AbaOrientacoes` | >= 5 |
| **F14** | Aba "Resumo da Auditoria" | M3 | Tier 1 | `TestFeature14AbaResumoAuditoria` | >= 5 |
| **F15** | Aba "Conferência" (38 Colunas) | M3 | Tier 1 | `TestFeature15AbaConferencia38Cols` | >= 5 |
| **F16** | Aba "Valores Importados" (11 Colunas) | M3 | Tier 1 | `TestFeature16AbaValoresImportados11Cols` | >= 5 |
| **F17** | Usabilidade Excel (Freeze Panes 'A2') | M3 | Tier 1 | `TestFeature17FreezePanesA2` | >= 5 |
| **F18** | Destaque Visual de Divergências (#FFC7CE/#9C0006) | M3 | Tier 1 | `TestFeature18DivergenceVisualHighlights` | >= 5 |
| **F19** | Saída Estática Sem Fórmulas | M3 | Tier 1 | `TestFeature19StaticValuesNoFormulas` | >= 5 |
| **F20** | CLI Runner & Orquestração em Memória | M3 | Tier 1 | `TestFeature20CliRunnerInMemory` | >= 5 |

---

## 4. Expected Output Derivation & Authoritative Sources

All test assertions derive expected values strictly from authoritative, documented specifications:

1. **Master Excel Oracle**: `Conferência Good Card_Base.xlsx` (15.913 linhas na folha de setembro/2026, 38 colunas na aba `Conferência`, 11 colunas em `Valores Importados`, 9 status em `Lista`, e 26 linhas em `Orientações`).
2. **Business Rules Specification**: `.agents/teamwork/explorer_survey_rules_3/business_rules_architecture.md`.
3. **Template Specification**: `.agents/teamwork/spec_miner_survey_1/template_spec.md`.
4. **CSV & Data Mapping Specification**: `.agents/teamwork/explorer_survey_csv_2/csv_survey.md`.
5. **Normalization & Deduplication Rules**: `.agents/teamwork/explorer_m1_3/normalization_spec.md`.

---

## 5. Progressive Testability & Milestone Awareness

To support parallel track execution:
- Tests inspect the presence of implementation modules (`config`, `data_loader`, `business_rules`, `excel_exporter`, `main`).
- Tests validate both the implementation code (when landed) and verify against the authoritative reference oracle (`tests/e2e/oracle.py`).
- If an implementation milestone is in progress, the test runner cleanly exercises available milestone features and oracle contracts without failing unbuilt downstream dependencies.
- Once all milestones are landed, the test runner validates the entire integrated system end-to-end.

---

## 6. Test Runner Commands

### Standard Runner Command
To execute the complete E2E test suite with comprehensive progress reporting and return code 0 on success:
```powershell
python tests/e2e/run_all_e2e_tests.py
```

### Python Unittest Discovery
Alternatively, using the standard Python test discovery runner:
```powershell
python -m unittest discover -s tests/e2e -p "test_*.py" -v
```

### Individual Tier Execution
```powershell
python -m unittest tests/e2e/test_tier1_features.py -v
python -m unittest tests/e2e/test_tier2_boundaries.py -v
python -m unittest tests/e2e/test_tier3_combinations.py -v
python -m unittest tests/e2e/test_tier4_workloads.py -v
```
