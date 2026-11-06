import openpyxl
import pandas as pd

wb = openpyxl.load_workbook('Conferência Good Card_Base.xlsx', read_only=True, data_only=True)
ws = wb['Conferência']

headers = [cell for cell in next(ws.iter_rows(max_row=1, values_only=True))]
rows = []
for row in ws.iter_rows(min_row=2, values_only=True):
    rows.append(row)
df = pd.DataFrame(rows, columns=headers)

# Group by Obs. Conferência and see conditions:
for obs_name, group in df.groupby('Obs. Conferência'):
    print(f"\n================ OBS: {repr(obs_name)} (count={len(group)}) ================")
    print("Summary of columns:")
    for col in ['Total Saldo Devedor', '6280 - Desc Cartão Benefícios Especiais', 
                '29621 - Ben Esp - Saldo Devedor Beneficios Especiais', 
                '29624 - Ben Esp - Custo Empresa Beneficios Especiais',
                'Margem 20% do Salário']:
        s = group[col].dropna()
        print(f"  {col}: min={s.min():.2f}, max={s.max():.2f}, mean={s.mean():.2f}, (zeros={(s==0).sum()}/{len(s)})")
    
    # Check is_estab400_or_giga in this group
    giga_mask = (
        group['Estabelecimento'].astype(str).str.contains('GIGA|400', case=False, na=False) |
        group['Código Estabelecimento'].astype(str).str.contains('400', na=False)
    )
    print(f"  Giga/400 rows in this obs: {giga_mask.sum()} / {len(group)}")
