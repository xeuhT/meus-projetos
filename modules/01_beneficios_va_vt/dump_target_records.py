import openpyxl
import pandas as pd

wb = openpyxl.load_workbook('Conferência Good Card_Base.xlsx', read_only=True, data_only=True)
ws = wb['Conferência']

headers = [cell for cell in next(ws.iter_rows(max_row=1, values_only=True))]
mat_idx = headers.index('Matrícula')

target_mats = {'324777', '327053', '320923', '321226'}

rows = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if str(row[mat_idx]) in target_mats:
        rows.append(row)

df = pd.DataFrame(rows, columns=headers)
for idx, r in df.iterrows():
    print(f"\n================ MATRÍCULA: {r['Matrícula']} ================")
    for col in headers:
        val = r[col]
        if val is not None and val != '' and val != 0 and val != 0.0:
            print(f"  {col}: {val}")
