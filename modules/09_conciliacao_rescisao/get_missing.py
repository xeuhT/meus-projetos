import pandas as pd
import re

caminho = r'C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\rescisao\eventos_unicos.xlsx'
df = pd.read_excel(caminho)

# Load auditor_massivo.py to find all implemented codes
with open(r'C:\Users\matcaval\OneDrive - Cencosud\Documentos\matheus\codigos\codigo rescisao\auditor_massivo.py', 'r', encoding='utf-8') as f:
    codigo = f.read()

# Extract implemented codes (look for 'elif cod == XYZ:')
implemented = set()
matches = re.findall(r'cod == (\d+):', codigo)
for m in matches:
    implemented.add(int(m))

# We also know some codes are references/bases that shouldn't be audited directly, or were handled manually
bases_and_refs = {499, 10100, 10130, 10140, 10150, 10200, 10202, 10210, 10230, 10290, 10295, 10301, 10316, 17910, 11000, 11220, 20208, 20305, 20417, 31906, 5015, 137, 138, 152}

missing = []
for idx, row in df.iterrows():
    c = int(row['CÓDIGO DO EVENTO'])
    if c not in implemented and c not in bases_and_refs:
        missing.append(f"{c} - {row['DESCRIÇÃO DO EVENTO']} ({row['TIPO DE EVENTO']})")

print("Faltam:", len(missing))
for m in missing:
    print(m)
