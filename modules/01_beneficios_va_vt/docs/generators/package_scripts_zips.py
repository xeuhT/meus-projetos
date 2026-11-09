"""
Empacota os scripts oficiais de cálculo em arquivos ZIP para cada frente de auditoria.
Salva o .zip na pasta de documentação correspondente.
"""

import os
import zipfile
from pathlib import Path

BASE_DOC_DIR = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\documentacao")
ENGINES_DIR = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\app_central_beneficios\engines")
ROOT_VA_VT = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT")

FRONT_MAP = [
    {
        "doc_dir": BASE_DOC_DIR / "01_Good_Card_Evt_6280",
        "zip_name": "Scripts_Good_Card_Evt_6280.zip",
        "sources": [
            ROOT_VA_VT / "04 - GOOD CARD",
            ENGINES_DIR / "good_card"
        ]
    },
    {
        "doc_dir": BASE_DOC_DIR / "02_Compra_Mensal_VT",
        "zip_name": "Scripts_Compra_Mensal_VT.zip",
        "sources": [
            ROOT_VA_VT / "Compra Mensal VT" / "scripts",
            ENGINES_DIR / "compra_mensal_vt"
        ]
    },
    {
        "doc_dir": BASE_DOC_DIR / "03_Valores_Pre_Folha_VT",
        "zip_name": "Scripts_Valores_Pre_Folha_VT.zip",
        "sources": [
            ROOT_VA_VT / "Valores Pré-Folha" / "scripts",
            ENGINES_DIR / "valores_pre_folha"
        ]
    },
    {
        "doc_dir": BASE_DOC_DIR / "04_Bloqueio_Good_Card",
        "zip_name": "Scripts_Bloqueio_Good_Card.zip",
        "sources": [
            ROOT_VA_VT / "Conferência Bloqueio Good Card" / "scripts",
            ENGINES_DIR / "bloqueio_good_card"
        ]
    }
]

for item in FRONT_MAP:
    doc_dir = item["doc_dir"]
    zip_path = doc_dir / item["zip_name"]
    doc_dir.mkdir(parents=True, exist_ok=True)
    
    # Prioridade para o pacote engine oficial
    selected_source = None
    for src in reversed(item["sources"]):
        if src.exists():
            selected_source = src
            break
            
    if not selected_source:
        print(f"[ALERTA] Nenhuma fonte encontrada para {item['zip_name']}")
        continue
        
    print(f"\nEmpacotando: {item['zip_name']} a partir de {selected_source}...")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(selected_source):
            # Ignorar pycache, tests, etc.
            if "__pycache__" in root or ".agents" in root or "tests" in root:
                continue
            for f in files:
                if f.endswith(".py"):
                    fp = Path(root) / f
                    arcname = f
                    zf.write(fp, arcname=arcname)
                    print(f"  + Adicionado: {arcname} ({fp.stat().st_size:,} bytes)")
                    
    print(f"[SUCESSO] ZIP criado: {zip_path.name} ({zip_path.stat().st_size:,} bytes)")

print("\n>>> TODOS OS ZIPS DA ETAPA 1 FORAM CRIADOS COM SUCESSO! <<<")
