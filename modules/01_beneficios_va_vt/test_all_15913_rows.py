import sys
from pathlib import Path
import pandas as pd
import openpyxl

root_dir = Path.cwd()
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

from simulate_proposed import classify_collaborator_row_proposed
import business_rules

print("Loading base Excel...")
wb = openpyxl.load_workbook('Conferência Good Card_Base.xlsx', read_only=True, data_only=True)
ws = wb['Conferência']
headers = [cell for cell in next(ws.iter_rows(max_row=1, values_only=True))]
rows = [r for r in ws.iter_rows(min_row=2, values_only=True)]
df_base = pd.DataFrame(rows, columns=headers)
print(f"Total rows in base: {len(df_base)}")

mismatches_obs = []
mismatches_trat = []

for idx, r in df_base.iterrows():
    mat = str(r['Matrícula'])
    sal = float(r['11010 - Salario Mensal'] or 0.0)
    tot = float(r['Total Saldo Devedor'] or 0.0)
    d = float(r['6280 - Desc Cartão Benefícios Especiais'] or 0.0)
    s_rem = float(r['29621 - Ben Esp - Saldo Devedor Beneficios Especiais'] or 0.0)
    c_emp = float(r['29624 - Ben Esp - Custo Empresa Beneficios Especiais'] or 0.0)
    cargo = str(r['Cargo'] or '')
    
    estab_cod = str(r['Código Estabelecimento'] or '')
    estab_nome = str(r['Estabelecimento'] or '')
    is_giga_400 = business_rules.is_giga_ou_estab400(estab_cod, estab_nome)
    
    has_afast = str(r['Início Afastamento'] or '') != '#N/A' and str(r['Início Afastamento'] or '') != 'None' and str(r['Início Afastamento'] or '') != ''
    dias_afast = 30 if has_afast else 0
    is_fer = str(r['Início Férias'] or '') != '#N/A' and str(r['Início Férias'] or '') != 'None' and str(r['Início Férias'] or '') != ''
    is_dem = str(r['Desligamento'] or '') != '#N/A' and str(r['Desligamento'] or '') != 'None' and str(r['Desligamento'] or '') != ''
    sit = str(r['Situação'] or 'Atividade Normal')
    
    obs_prop, trat_prop = classify_collaborator_row_proposed(
        total_saldo_devedor=tot,
        salario_11010=sal,
        evt_6280=d,
        evt_29621=s_rem,
        evt_29624=c_emp,
        cargo=cargo,
        dias_afastamento=dias_afast,
        has_afastamento=has_afast,
        is_ferias=is_fer,
        is_demitido=is_dem,
        is_estab400_or_giga=is_giga_400,
        situacao=sit,
    )
    
    base_obs = r['Obs. Conferência']
    base_trat = r['Tratativa']
    
    if obs_prop != base_obs:
        mismatches_obs.append((mat, base_obs, obs_prop, tot, d, s_rem, c_emp, is_giga_400))
    if (trat_prop or None) != (base_trat or None):
        mismatches_trat.append((mat, base_trat, trat_prop))

print(f"\nTotal Obs mismatches: {len(mismatches_obs)}")
if mismatches_obs:
    print("Sample Obs mismatches:")
    for m in mismatches_obs[:10]:
        print(f"  Matr {m[0]} | Base: {repr(m[1])} | Proposed: {repr(m[2])} | Saldo: {m[3]}, 6280: {m[4]}, 29621: {m[5]}, 29624: {m[6]}, Giga: {m[7]}")

print(f"Total Tratativa mismatches: {len(mismatches_trat)}")
if mismatches_trat:
    print("Sample Tratativa mismatches:")
    for m in mismatches_trat[:10]:
        print(f"  Matr {m[0]} | Base: {repr(m[1])} | Proposed: {repr(m[2])}")
