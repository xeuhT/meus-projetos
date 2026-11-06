import sys
from pathlib import Path
import pandas as pd
import openpyxl

root_dir = Path.cwd()
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

import business_rules
from simulate_proposed import classify_collaborator_row_proposed
from data_loader import load_all_data

# Monkey-patch business_rules.classify_collaborator_row
business_rules.classify_collaborator_row = classify_collaborator_row_proposed

print("Loading data via data_loader...")
data = load_all_data(data_dir=Path("Relatorio Original"), template_path=Path("Conferência Good Card_Base.xlsx"))

print("Executing audit with proposed classification...")
df_conf, df_vi, kpis = business_rules.execute_audit(data)

# Compare with base Excel
wb = openpyxl.load_workbook('Conferência Good Card_Base.xlsx', read_only=True, data_only=True)
ws = wb['Conferência']
headers = [cell for cell in next(ws.iter_rows(max_row=1, values_only=True))]
rows = [r for r in ws.iter_rows(min_row=2, values_only=True)]
df_base = pd.DataFrame(rows, columns=headers)

df_conf['Matrícula'] = df_conf['Matrícula'].astype(str)
df_base['Matrícula'] = df_base['Matrícula'].astype(str)

merged = pd.merge(df_conf, df_base, on='Matrícula', suffixes=('_gen', '_base'))

obs_diff = merged[merged['Obs. Conferência_gen'] != merged['Obs. Conferência_base']]
print(f"\nTotal rows with differing Obs. Conferência: {len(obs_diff)}")
for idx, r in obs_diff.iterrows():
    print(f"Matr: {r['Matrícula']} | Estab: {r.get('Estabelecimento_gen', '')} | Saldo: {r.get('Total Saldo Devedor_gen', 0)} | 6280: {r.get('6280 - Desc Cartão Benefícios Especiais_gen', 0)} | 29621: {r.get('29621 - Ben Esp - Saldo Devedor Beneficios Especiais_gen', 0)} | 29624: {r.get('29624 - Ben Esp - Custo Empresa Beneficios Especiais_gen', 0)}")
    print(f"   Generated: {repr(r['Obs. Conferência_gen'])} | Trat: {repr(r['Tratativa_gen'])}")
    print(f"   Base:      {repr(r['Obs. Conferência_base'])} | Trat: {repr(r['Tratativa_base'])}")

trat_diff = merged[merged['Tratativa_gen'].fillna('') != merged['Tratativa_base'].fillna('')]
print(f"\nTotal rows with differing Tratativa: {len(trat_diff)}")
if len(trat_diff) > 0:
    for idx, r in trat_diff.iterrows():
        print(f"Matr {r['Matrícula']}: gen={repr(r['Tratativa_gen'])}, base={repr(r['Tratativa_base'])}")

print(f"\nKPIs Conformes: {kpis.get('total_conformes')} / {kpis.get('total_auditados')}")
print(f"KPIs Divergências: {kpis.get('total_divergencias')}")
