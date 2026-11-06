import pandas as pd
df = pd.read_excel(r'C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\rescisao\Auditoria_Rescisao_Agosto26.xlsx')

print("=== 335714 ===")
emp1 = df[df['MATRICULA COLABORADOR'] == 335714]
for _, row in emp1.iterrows():
    if row['VALOR EVENTO'] > 0:
        print(f"{row['CÓDIGO DO EVENTO']} - {row['DESCRIÇÃO DO EVENTO']}: {row['VALOR EVENTO']} (Ref: {row['VALOR EVENTO REFERÊNCIA']})")

print("\n=== 298063 ===")
emp2 = df[df['MATRICULA COLABORADOR'] == 298063]
for _, row in emp2.iterrows():
    if row['VALOR EVENTO'] > 0:
        print(f"{row['CÓDIGO DO EVENTO']} - {row['DESCRIÇÃO DO EVENTO']}: {row['VALOR EVENTO']} (Ref: {row['VALOR EVENTO REFERÊNCIA']})")

print("\n=== 336002 ===")
emp3 = df[df['MATRICULA COLABORADOR'] == 336002]
for _, row in emp3.iterrows():
    if row['VALOR EVENTO'] > 0:
        print(f"{row['CÓDIGO DO EVENTO']} - {row['DESCRIÇÃO DO EVENTO']}: {row['VALOR EVENTO']} (Ref: {row['VALOR EVENTO REFERÊNCIA']})")
