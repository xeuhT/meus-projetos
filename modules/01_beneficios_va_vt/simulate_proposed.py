import sys
from pathlib import Path
import pandas as pd
import openpyxl

root_dir = Path.cwd()
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

import business_rules
from config import PERCENTUAL_CUSTO_EMPRESA_GIGA, PERCENTUAL_MARGEM_SALARIAL, TOLERANCIA_CENTAVOS

# Load base Excel
wb = openpyxl.load_workbook('Conferência Good Card_Base.xlsx', read_only=True, data_only=True)
ws = wb['Conferência']
headers = [cell for cell in next(ws.iter_rows(max_row=1, values_only=True))]
rows = [r for r in ws.iter_rows(min_row=2, values_only=True)]
df_base = pd.DataFrame(rows, columns=headers)

# Define proposed classify_collaborator_row
def classify_collaborator_row_proposed(
    total_saldo_devedor: float,
    salario_11010: float,
    evt_6280: float,
    evt_29621: float,
    evt_29624: float,
    cargo: str = "",
    dias_afastamento: int = 0,
    has_afastamento: bool = False,
    is_ferias: bool = False,
    is_demitido: bool = False,
    tipo_rescisao: str = "",
    is_estab400_or_giga: bool = False,
    situacao: str = "Atividade Normal",
):
    total = round(float(total_saldo_devedor or 0.0), 2)
    margem = round(float(salario_11010 or 0.0) * PERCENTUAL_MARGEM_SALARIAL, 2)
    desc = round(float(evt_6280 or 0.0), 2)
    saldo_rem = round(float(evt_29621 or 0.0), 2)
    custo_emp = round(float(evt_29624 or 0.0), 2)

    # 1. Cargos Inelegíveis
    if business_rules.is_aprendiz(cargo) and (total > 0 or desc > 0):
        return ("Aprendiz - Verificar", "Concessão indevida de Good Card para cargo de Aprendiz. Solicitar cancelamento do cartão junto à operadora.")
    if business_rules.is_diretor(cargo) and (total > 0 or desc > 0):
        return ("Diretor - Verificar", "Concessão indevida de Good Card para cargo de Diretor. Cargo inelegível conforme política de benefícios.")

    # 2. Desligados
    if is_demitido:
        if "TRANSFERÊNCIA SEM ÔNUS" in str(tipo_rescisao).upper():
            return ("Transferência sem ônus", "Transferência de empresa: verificar migração do cartão e acerto entre filiais.")
        if total == 0 and desc == 0:
            return "Ok - Sem saldo devedor e sem desconto", None
        if total > 0 and desc == total and saldo_rem == 0:
            return "Ok - Liquidou todo o saldo devedor ", None
        if saldo_rem > 0 or (total > desc):
            return ("Demitido - Saldo devedor pendente", "Colaborador desligado com saldo devedor não liquidado. Verificar inclusão no TRCT e solicitar cancelamento do cartão.")

    # 3. Suspensão
    if "SUSPENS" in str(situacao).upper():
        return "Ok - Suspensão de Contrato", None

    # 4. Sem Saldo e Sem Desconto
    if total == 0 and desc == 0:
        return "Ok - Sem saldo devedor e sem desconto", None

    # 5. Afastados > 15 dias
    if has_afastamento and (dias_afastamento > 15 or "AFAST" in str(situacao).upper()):
        if desc == 0 and saldo_rem == total and total > 0:
            return ("Afastamento - Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ", None)
        if desc > 0 and salario_11010 == 0:
            return ("Afastado - Desconto indevido em folha", "Colaborador afastado > 15 dias com desconto em folha. Validar margem líquida e solicitar estorno/regularização.")
        if total > 0 and desc == 0 and saldo_rem == 0:
            return ("Afastado - Saldo devedor não provisionado", "Saldo devedor não transferido para o evento 29621 na folha. Acionar equipe de folha para provisionamento.")

    # 6. Férias no mês
    if is_ferias and desc == 0 and saldo_rem == total and total > 0:
        return ("Férias- Não teve desconto mas o saldo devedor está sendo apresentado no evento 29621 ", None)

    # --- NOVO FLUXO REGRAS 7 E 8 ---
    # 7. Quitação Integral (Válido universalmente, inclusive Giga/400 se o colaborador quitou 100% e custo_emp == 0)
    if total <= margem and abs(desc - total) <= TOLERANCIA_CENTAVOS and saldo_rem == 0 and custo_emp == 0:
        return "Ok - Liquidou todo o saldo devedor ", None

    # 8. Regra Especial 5% Custo Empresa (Giga / Estab 400)
    if is_estab400_or_giga and total > 0:
        calc_emp = round(total * PERCENTUAL_CUSTO_EMPRESA_GIGA, 2)
        # Se houve subsídio da empresa (ou se saldo_rem == 0 e desc < total)
        if abs(custo_emp - calc_emp) <= 0.05:
            # Equação contábil completa
            if abs(desc + saldo_rem + custo_emp - total) <= TOLERANCIA_CENTAVOS:
                return (
                    "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621 (5% Custo empresa)",
                    None,
                )
        elif custo_emp > 0 and saldo_rem == 0:
            # Divergência se custo_emp foi informado mas diverge dos 5% regulamentares
            return (
                "Divergência Custo Empresa 5%",
                "Valor levado para o custo empresa (29624) diverge dos 5% regulamentares para o estabelecimento 400/Giga.",
            )

    # 9. Regra 4: Desconto e Quitação Padrão
    # 9.1 Desconto fora da margem dos 20%
    if desc > margem:
        return (
            "Desconto fora da margem dos 20%",
            "Desconto no evento 6280 superior à margem consignável de 20% do salário mensal (11010). Ajustar valor descontado.",
        )

    # 9.2 Descontou dentro da margem e levou restante para 29621
    if (
        total > margem
        and desc <= margem
        and abs(desc + saldo_rem + custo_emp - total) <= TOLERANCIA_CENTAVOS
        and saldo_rem > 0
    ):
        return (
            "Ok - Descontou dentro da margem e está levando o restante do saldo devedor para o mês seguinte no evento 29621",
            None,
        )

    # 9.3 Ativo sem desconto
    if not has_afastamento and not is_ferias and total > 0 and desc == 0:
        return (
            "Verificar com quality - Não estava em situação de afastamento e não houve o desconto",
            "Verificar com Quality/Folha: Colaborador com saldo devedor ativo e sem desconto em folha, não estando afastado nem em férias.",
        )

    # 9.4 Quebra da equação contábil
    if abs(desc + saldo_rem + custo_emp - total) > TOLERANCIA_CENTAVOS:
        return (
            "Divergência no saldo devedor",
            "Soma do desconto (6280) + saldo residual (29621) + custo empresa (29624) diverge do Total Saldo Devedor. Auditar folha.",
        )

    return "Ok - Liquidou todo o saldo devedor ", None

# Test against the 4 target rows
targets = ['324777', '327053', '320923', '321226']
for m in targets:
    r = df_base[df_base['Matrícula'].astype(str) == m].iloc[0]
    sal = float(r['11010 - Salario Mensal'] or 0.0)
    tot = float(r['Total Saldo Devedor'] or 0.0)
    d = float(r['6280 - Desc Cartão Benefícios Especiais'] or 0.0)
    s_rem = float(r['29621 - Ben Esp - Saldo Devedor Beneficios Especiais'] or 0.0)
    c_emp = float(r['29624 - Ben Esp - Custo Empresa Beneficios Especiais'] or 0.0)
    obs_prop, trat_prop = classify_collaborator_row_proposed(
        total_saldo_devedor=tot,
        salario_11010=sal,
        evt_6280=d,
        evt_29621=s_rem,
        evt_29624=c_emp,
        cargo=str(r['Cargo']),
        is_estab400_or_giga=True,
    )
    print(f"=== Matr {m} ===")
    print(f"  Base Obs:     {repr(r['Obs. Conferência'])}")
    print(f"  Proposed Obs: {repr(obs_prop)}")
    print(f"  Base Trat:     {repr(r['Tratativa'])}")
    print(f"  Proposed Trat: {repr(trat_prop)}")
    assert obs_prop == r['Obs. Conferência'], f"Mismatch Obs for {m}"
    assert trat_prop == r['Tratativa'], f"Mismatch Trat for {m}"

print("\nALL 4 TARGET ROWS MATCH PERFECTLY!")
