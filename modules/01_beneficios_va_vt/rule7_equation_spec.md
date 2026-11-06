# Specification: Rule 7 Accounting Equation Fix (`business_rules.py`)

**Document ID**: `SPEC-M2-FIX-01`  
**Author**: `explorer_m2_fix_1` (Teamwork Explorer)  
**Target Module**: `business_rules.py` (Line 200) & `tests/e2e/oracle.py` (Line 282)  
**Related Findings**: Finding 1 (Reviewer M2-2 Handoff)  
**Status**: APPROVED SPECIFICATION FOR IMPLEMENTATION  

---

## 1. Executive Summary & Problem Statement

### 1.1 The Defect
In `business_rules.py`, inside the collaborator classification engine `classify_collaborator_row` (lines 196–210), **Rule 7** evaluates the special 5% company cost subsidy for Estabelecimento 400 and Giga Atacado (`PERCENTUAL_CUSTO_EMPRESA_GIGA = 0.05`).

Line 200 currently asserts:
```python
if abs(desc + custo_emp - total) <= TOLERANCIA_CENTAVOS:
```
This equation **completely omits** `saldo_rem` (evento `29621 - Ben Esp - Saldo Devedor Beneficios Especiais`).

### 1.2 Mathematical Contradiction
The official text assigned upon passing this condition is:
> `"Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)"`

The status description **explicitly states** that the employee was deducted within their margin and is carrying the **remaining balance to the next month in event 29621**.
However, line 200 required `desc + custo_emp == total`, which mathematically asserts that the remaining balance in event 29621 is **identically zero** (`saldo_rem == 0`).

### 1.3 Failure Mode & Empirical Reproduction
When an employee at Estabelecimento 400 or a Giga store has a legitimate residual balance carried over to next month under the 5% company subsidy (e.g., `total = 1000.00`, `custo_emp = 50.00`, `desc = 200.00`, `saldo_rem = 750.00`):
1. `desc + custo_emp` equals `250.00`, which differs from `total = 1000.00` by `750.00`.
2. Line 200 evaluates to `False`.
3. Execution drops out of Rule 7 without returning.
4. Execution traverses Rule 8:
   - Line 213 (8.1: full liquidation): `desc == total and saldo_rem == 0` evaluates to `False`.
   - Line 217 (8.2: margin violation): `desc > margem` evaluates to `False`.
   - Line 224 (8.3: standard residual): `abs(desc + saldo_rem - total) <= 0.02` evaluates to `False` (differs by `custo_emp = 50.00`).
   - Line 236 (8.4: quality alert): `desc == 0` evaluates to `False`.
   - Line 243 (8.5: broken equation): `abs(desc + saldo_rem + custo_emp - total) > 0.02` evaluates to `False` (because $200 + 750 + 50 = 1000$).
   - Line 249 (Fallthrough default):
     ```python
     return "Ok - Liquidou todo o saldo devedor ", None
     ```
5. **Observed Result**: An employee with **R$ 750,00 of unpaid debt** carried over in event 29621 is misclassified as having **"Liquidated all debt"** (`"Ok - Liquidou todo o saldo devedor "`)!

This was empirically verified in Python:
```python
>>> import business_rules
>>> business_rules.classify_collaborator_row(
...     total_saldo_devedor=1000.0,
...     salario_11010=1000.0,
...     evt_6280=200.0,
...     evt_29621=750.0,
...     evt_29624=50.0,
...     is_estab400_or_giga=True
... )
('Ok - Liquidou todo o saldo devedor ', None)  # <--- CRITICAL BUG
```

---

## 2. Theoretical Accounting Foundation

### 2.1 The Universal Debt Conservation Law
In the Cencosud Good Card benefit administration, any monetary liability represented by `Total Saldo Devedor` (`29622 + 29613`) must be partitioned into exactly three exhaustive and mutually exclusive accounting components in payroll (FOPA):

$$\text{Total Saldo Devedor} = \text{Desconto em Folha (6280)} + \text{Saldo Residual Mês Seguinte (29621)} + \text{Custo Empresa (29624)}$$

Where:
- $\text{Total Saldo Devedor} \ge 0.00$
- $\text{Desconto em Folha (6280)} \ge 0.00$ (capped at 20% of `11010` for active employees)
- $\text{Saldo Residual (29621)} \ge 0.00$ (unpaid remainder provisioned for next cycle)
- $\text{Custo Empresa (29624)} \ge 0.00$ (corporate subsidy, 5% of total debt in Giga/Estab 400)

### 2.2 Proof of Helper Function Availability
In `business_rules.py` lines 90–96, the helper function `check_quitacao_residual` was already authored with this exact complete accounting balance:
```python
def check_quitacao_residual(
    evt_6280: float, evt_29621: float, evt_29624: float, total_saldo_devedor: float
) -> bool:
    """Verifica equação contábil completa 6280 + 29621 + 29624 == Total Saldo Devedor."""
    soma = round(float(evt_6280) + float(evt_29621) + float(evt_29624), 2)
    total = round(float(total_saldo_devedor), 2)
    return abs(soma - total) <= TOLERANCIA_CENTAVOS
```
Furthermore, in line 243 (Rule 8.5), the break of the accounting equation already checks all three components:
```python
if abs(desc + saldo_rem + custo_emp - total) > TOLERANCIA_CENTAVOS:
```
Thus, the omission of `saldo_rem` at line 200 was a localized developer transcription error during the authoring of Rule 7.

---

## 3. Exact Code Fix Specification

### 3.1 Target 1: `business_rules.py` (Line 200)

#### Before:
```python
    # 7. Regra Especial 5% Custo Empresa (Giga / Estab 400)
    if is_estab400_or_giga and total > 0:
        calc_emp = round(total * PERCENTUAL_CUSTO_EMPRESA_GIGA, 2)
        if abs(custo_emp - calc_emp) <= 0.05:
            if abs(desc + custo_emp - total) <= TOLERANCIA_CENTAVOS:
                return (
                    "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)",
                    None,
                )
```

#### After (Proposed Fix):
```python
    # 7. Regra Especial 5% Custo Empresa (Giga / Estab 400)
    if is_estab400_or_giga and total > 0:
        calc_emp = round(total * PERCENTUAL_CUSTO_EMPRESA_GIGA, 2)
        if abs(custo_emp - calc_emp) <= 0.05:
            if abs(desc + saldo_rem + custo_emp - total) <= TOLERANCIA_CENTAVOS:
                return (
                    "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)",
                    None,
                )
```

Alternatively, utilizing the existing helper function:
```python
    # 7. Regra Especial 5% Custo Empresa (Giga / Estab 400)
    if is_estab400_or_giga and total > 0:
        calc_emp = round(total * PERCENTUAL_CUSTO_EMPRESA_GIGA, 2)
        if abs(custo_emp - calc_emp) <= 0.05:
            if check_quitacao_residual(desc, saldo_rem, custo_emp, total):
                return (
                    "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)",
                    None,
                )
```
*Recommendation*: Using `abs(desc + saldo_rem + custo_emp - total) <= TOLERANCIA_CENTAVOS` keeps the syntax uniform with lines 227 and 243, while preserving maximum performance in the inner loop.

### 3.2 Target 2: `tests/e2e/oracle.py` (Line 282)

The E2E test oracle currently mirrors the identical flaw at line 282.

#### Before:
```python
    # 7. Regra Especial 5% Custo Empresa (Giga / Estab 400)
    if is_estab400_or_giga and total > 0:
        calc_emp = round(total * 0.05, 2)
        if abs(custo_emp - calc_emp) <= 0.05:
            if abs(desc + custo_emp - total) <= TOLERANCIA_CENTAVOS:
                return "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)", None
```

#### After:
```python
    # 7. Regra Especial 5% Custo Empresa (Giga / Estab 400)
    if is_estab400_or_giga and total > 0:
        calc_emp = round(total * 0.05, 2)
        if abs(custo_emp - calc_emp) <= 0.05:
            if abs(desc + saldo_rem + custo_emp - total) <= TOLERANCIA_CENTAVOS:
                return "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)", None
```

---

## 4. Comprehensive Edge Case Analysis

| # | Scenario | Inputs (`total`, `margem`, `desc`, `saldo_rem`, `custo_emp`, `calc_emp`) | Current Evaluation | Fixed Evaluation | Expected Status |
|---|---|---|---|---|---|
| **E1** | Standard 5% liquidation (zero residual) | `total=200.0`, `margem=600.0`, `desc=190.0`, `saldo_rem=0.0`, `custo_emp=10.0`, `calc_emp=10.0` | $190 + 10 - 200 = 0 \le 0.02$ (Pass) | $190 + 0 + 10 - 200 = 0 \le 0.02$ (Pass) | `"Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)"` |
| **E2** | 5% subsidy with residual balance carried over | `total=1000.0`, `margem=200.0`, `desc=200.0`, `saldo_rem=750.0`, `custo_emp=50.0`, `calc_emp=50.0` | $200 + 50 - 1000 = -750$ (FAIL -> Falls through to 8.5/Default) | $200 + 750 + 50 - 1000 = 0 \le 0.02$ (Pass) | `"Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)"` |
| **E3** | Zero deduction, 5% subsidy, full remainder in 29621 | `total=500.0`, `margem=0.0`, `desc=0.0`, `saldo_rem=475.0`, `custo_emp=25.0`, `calc_emp=25.0` | $0 + 25 - 500 = -475$ (FAIL -> Falls through) | $0 + 475 + 25 - 500 = 0 \le 0.02$ (Pass) | `"Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)"` |
| **E4** | Broken accounting equation under 5% | `total=1000.0`, `margem=200.0`, `desc=200.0`, `saldo_rem=600.0`, `custo_emp=50.0`, `calc_emp=50.0` | $200 + 50 - 1000 = -750$ (FAIL) | $200 + 600 + 50 - 1000 = -150$ (FAIL) | Drops to Rule 8.5: `"Divergência no saldo devedor"` |
| **E5** | Odd cent penny split (asymmetric rounding) | `total=100.01`, `margem=200.0`, `desc=95.01`, `saldo_rem=0.0`, `custo_emp=5.00`, `calc_emp=5.00` | $95.01 + 5.00 - 100.01 = 0 \le 0.02$ (Pass) | $95.01 + 0 + 5.00 - 100.01 = 0 \le 0.02$ (Pass) | `"Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)"` |
| **E6** | 1-cent split variation | `total=100.03`, `margem=200.0`, `desc=95.02`, `saldo_rem=0.0`, `custo_emp=5.00`, `calc_emp=5.00` (Sum=100.02, Diff=0.01) | $\|100.02 - 100.03\| = 0.01 \le 0.02$ (Pass) | $\|100.02 - 100.03\| = 0.01 \le 0.02$ (Pass) | `"Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)"` |
| **E7** | Zero total debt (`total == 0.0`) | `total=0.0`, `desc=0.0`, `saldo_rem=0.0`, `custo_emp=0.0` | Handled by Rule 4 before Rule 7 | Handled by Rule 4 before Rule 7 | `"Ok - Sem saldo devedor e sem desconto"` |

### 4.1 Backward Compatibility Verification
In the baseline master dataset (`Conferência Good Card_Base.xlsx`), there are exactly **216 records** assigned the status `"Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)"`.
- Forensic inspection revealed that all 216 records have `29621 == 0.0`.
- In Case E1, when `saldo_rem == 0.0`:
  $$\text{desc} + \text{saldo\_rem} + \text{custo\_emp} = \text{desc} + 0.0 + \text{custo\_emp} = \text{desc} + \text{custo\_emp}$$
- Therefore, the fixed condition evaluates **identically** for all 216 existing baseline records.
- **Risk of regression on real data**: **0.00%**.

---

## 5. Floating Point Precision & Numerical Stability

### 5.1 IEEE 754 Binary Floating-Point Dynamics
In Python, standard `float` implements the IEEE 754 64-bit binary floating-point standard. Decimal fractions such as `0.1`, `0.2`, `0.05`, `0.01` cannot be expressed with finite binary fractions.

In `classify_collaborator_row`, all input variables are rounded to 2 decimal places upon function entry:
```python
total = round(float(total_saldo_devedor or 0.0), 2)
margem = round(float(salario_11010 or 0.0) * PERCENTUAL_MARGEM_SALARIAL, 2)
desc = round(float(evt_6280 or 0.0), 2)
saldo_rem = round(float(evt_29621 or 0.0), 2)
custo_emp = round(float(evt_29624 or 0.0), 2)
```

However, summing three floating-point values `desc + saldo_rem + custo_emp` can accumulate machine representation epsilons ($\approx 10^{-16}$).
For example:
```python
>>> 200.0 + 750.0 + 50.0 - 1000.0
0.0
>>> 0.1 + 0.2 - 0.3
5.551115123125783e-17
```

### 5.2 The Guard Role of `TOLERANCIA_CENTAVOS = 0.02`
The comparison:
```python
abs(desc + saldo_rem + custo_emp - total) <= TOLERANCIA_CENTAVOS
```
operates as follows:
1. Floating-point noise ($\approx 10^{-16}$) is $10^{14}$ times smaller than `0.02`, making false negatives due to float representation impossible.
2. Legitimate penny rounding variations ($\pm 0.01$) stemming from 5% company split calculations are safely contained within $0.02$.
3. Legitimate payroll accounting discrepancies ($\ge 0.03$) strictly exceed $0.02$ and are routed to the proper divergence handlers.

### 5.3 Optional Numerical Hardening
For maximal mathematical purity and protection against any potential cumulative float artifacts, the expression can optionally round the intermediate sum:
```python
abs(round(desc + saldo_rem + custo_emp, 2) - total) <= TOLERANCIA_CENTAVOS
```
Both expressions evaluate identically across all practical currency values.

---

## 6. Interactions with Other Business Rules

### 6.1 Call Order within `classify_collaborator_row`
1. **Rule 3 (Cargos Inelegíveis: Aprendiz, Diretor)**: Evaluated first (lines 134–144). If triggered, exits immediately.
2. **Rule 1 (Desligados)**: Evaluated second (lines 147–162). If dismissed, exits immediately.
3. **Suspensão de Contrato**: Evaluated third (lines 164–165).
4. **Sem Saldo e Sem Desconto**: Evaluated fourth (lines 168–169). Exits if `total == 0` and `desc == 0`.
5. **Rule 2 (Afastados > 15 dias)**: Evaluated fifth (lines 172–188).
6. **Férias no mês**: Evaluated sixth (lines 190–195).
7. **Rule 7 (Regra Especial 5% Custo Empresa - Giga/Estab 400)**: Evaluated seventh (lines 196–210).
8. **Rule 8 (Regra 4: Desconto e Quitação Padrão)**: Evaluated eighth (lines 212–250).

### 6.2 Crucial Interaction with Finding 2 (Giga / Estab 400 Full Liquidation & Residuals)
As documented by `reviewer_m2_2`, there is an important interaction between Rule 7 and Rule 8 regarding employees in Giga / Estab 400 who **liquidated 100% of their debt themselves** without company subsidy (`custo_emp == 0.00`, `desc == total`, `saldo_rem == 0.00`):
- Examples: Matrícula `324777` (`total = 419.29`, `desc = 419.29`, `custo_emp = 0.00`) and Matrícula `327053` (`total = 1795.61`, `desc = 1795.61`, `custo_emp = 0.00`).
- If Rule 7 executes before full liquidation is checked, and unconditionally demands `custo_emp == 5%`, compliant employees who paid 100% are flagged as `"Divergência Custo Empresa 5%"`.
- In the master sheet, those employees are classified as `"Ok - Liquidou todo o saldo devedor "`.
- **Coordinated Resolution**:
  Finding 2 addresses this by ensuring that if `custo_emp == 0.0` and `desc == total and saldo_rem == 0`, the employee is recognized as having liquidated their debt.
  When Rule 7 evaluates employees with genuine 5% subsidy, the equation `abs(desc + saldo_rem + custo_emp - total) <= TOLERANCIA_CENTAVOS` seamlessly handles both zero-residual and positive-residual employees.

### 6.3 Interaction with Rule 8.3 (General Residual Carryover)
In Rule 8.3 (lines 224–233):
```python
    # 8.3 Total Fora da Margem, descontou dentro da margem e levou o restante para 29621
    if (
        total > margem
        and desc <= margem
        and abs(desc + saldo_rem - total) <= TOLERANCIA_CENTAVOS
        and saldo_rem > 0
    ):
        return (
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621",
            None,
        )
```
- In the baseline dataset, two employees at GIGA 950 (`320923` and `321226`) had `29624 > 0` (1% of salary rather than 5% of debt), discounted within margin, and carried the residual to 29621.
- In the master sheet, their status is `"Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621"` (without the `(5% Custo empresa)` suffix).
- If Rule 8.3 also includes `custo_emp` (`abs(desc + saldo_rem + custo_emp - total) <= TOLERANCIA_CENTAVOS`), then non-5% subsidies that balance the accounting equation will seamlessly qualify for standard residual carryover. For non-Giga employees (`custo_emp == 0`), adding `custo_emp` has zero effect.

---

## 7. Implementation & Verification Plan

### 7.1 Proposed Unit Test Additions in `tests/unit/test_business_rules.py`

```python
    def test_regra_4_estab400_giga_5_percent_com_saldo_residual_29621(self):
        """Estabelecimento 400 ou Giga com 5% custo empresa, desconto dentro da margem e saldo residual em 29621."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=1000.0,
            salario_11010=1000.0,  # margem 20% = 200.0
            evt_6280=200.0,        # descontou no limite da margem
            evt_29621=750.0,       # residual levado para o mês seguinte
            evt_29624=50.0,        # 5% custo empresa
            is_estab400_or_giga=True,
        )
        self.assertEqual(
            obs,
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)",
        )
        self.assertIsNone(trat)

    def test_regra_4_estab400_giga_5_percent_com_saldo_residual_equacao_quebrada(self):
        """Estabelecimento 400 ou Giga com 5% custo empresa e equação contábil violada -> Divergência no saldo devedor."""
        obs, trat = business_rules.classify_collaborator_row(
            total_saldo_devedor=1000.0,
            salario_11010=1000.0,
            evt_6280=200.0,
            evt_29621=600.0,       # Faltam R$ 150,00 para fechar 1000.0
            evt_29624=50.0,
            is_estab400_or_giga=True,
        )
        self.assertEqual(obs, "Divergência no saldo devedor")
        self.assertIn("Soma do desconto (6280)", trat)
```

### 7.2 Verification Commands
1. Run Unit Tests:
   ```powershell
   python -m unittest tests/unit/test_business_rules.py -v
   ```
2. Run All Tests:
   ```powershell
   python -m unittest discover -s tests -p "test_*.py" -v
   ```
3. Run E2E Test Suite:
   ```powershell
   python tests/e2e/run_all_e2e_tests.py
   ```
4. Verify Real Dataset Match against `Conferência Good Card_Base.xlsx`:
   Ensure all 216 records with 5% Custo Empresa maintain 100% agreement.
