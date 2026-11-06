import pandas as pd
df = pd.read_excel(r'C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\rescisao\Auditoria_Rescisao_Agosto26.xlsx')
emp = df[(df['MATRICULA COLABORADOR'] == 292620) & (df['CÓDIGO DO EVENTO'].isin([11000, 11220, 721]))]
for _, row in emp.iterrows():
    print(f"{row['CÓDIGO DO EVENTO']} - {row['DESCRIÇÃO DO EVENTO']}: {row['VALOR EVENTO']} (Ref: {row['VALOR EVENTO REFERÊNCIA']})")
