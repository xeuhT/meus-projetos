"""
Script de Compilação do Executável Único (.EXE Standalone — OneFile)
Conferência e Carga de Seguro de Vida — Sistemas RH Cencosud Brasil.

Boas práticas aplicadas:
- Compilação isolada em pasta temporária única fora do OneDrive (evita travas de sync).
- Zero exclusão massiva de arquivos (compatibilidade total com TrendAI Security Agent).
- Inclusão completa de CustomTkinter, Pillow (PIL), OpenPyXL, Pandas, Numpy, Assets visuais
  Cencosud e Catálogos JSON embutidos.
- Publicação de um único executável oficial pronto para envio à equipe de RH.
"""

import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path


def build() -> None:
    print("=" * 80)
    print("COMPILAÇÃO DO EXECUTÁVEL ÚNICO (.EXE STANDALONE — ONEFILE)")
    print("CONFERÊNCIA E CARGA DE SEGURO DE VIDA — CENCOSUD BRASIL")
    print("=" * 80)

    base_dir = Path(__file__).resolve().parent
    main_script = base_dir / "app_seguro_vida.py"
    exe_name = "Conferencia_Seguro_Vida_Cencosud.exe"

    # Diretório oficial de publicação do executável único para o RH
    sistemas_rh_seguro = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\seguro de vida")
    target_exe = (sistemas_rh_seguro if sistemas_rh_seguro.exists() else base_dir) / exe_name

    # Pasta temporária exclusiva fora do OneDrive (sem exclusão em massa para não acionar TrendAI)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_build_root = Path(tempfile.gettempdir()) / f"cencosud_seguro_vida_pyi_{stamp}"
    temp_work = temp_build_root / "work"
    temp_dist = temp_build_root / "dist"
    temp_spec = temp_build_root / "spec"
    for d in (temp_work, temp_dist, temp_spec):
        d.mkdir(parents=True, exist_ok=True)

    print(f"[*] Diretório isolado de compilação: {temp_build_root}")

    icon_path = base_dir / "assets" / "cencosud.ico"
    assets_dir = base_dir / "assets"
    optantes_json = base_dir / "optantes_complementar.json"
    diretores_json = base_dir / "diretores_cadastro.json"

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--name",
        "Conferencia_Seguro_Vida_Cencosud",
        "--workpath",
        str(temp_work),
        "--distpath",
        str(temp_dist),
        "--specpath",
        str(temp_spec),
        "--paths",
        str(base_dir),
        "--collect-all",
        "customtkinter",
        "--collect-all",
        "PIL",
        "--collect-all",
        "openpyxl",
        "--hidden-import",
        "config",
        "--hidden-import",
        "file_detector",
        "--hidden-import",
        "data_loader",
        "--hidden-import",
        "business_rules",
        "--hidden-import",
        "excel_exporter",
        "--hidden-import",
        "main",
        "--add-data",
        f"{assets_dir};assets",
        "--add-data",
        f"{optantes_json};.",
        "--add-data",
        f"{diretores_json};.",
        "--icon",
        str(icon_path),
        str(main_script),
    ]

    print("\n[*] Executando PyInstaller (--onefile --windowed)...")
    result = subprocess.run(cmd, cwd=str(base_dir))

    if result.returncode != 0:
        print("\n[!] FALHA NA COMPILAÇÃO DO EXECUTÁVEL.")
        sys.exit(result.returncode)

    compiled_exe = temp_dist / exe_name
    if not compiled_exe.exists():
        print(f"[!] Erro: Executável compilado não encontrado em {compiled_exe}")
        sys.exit(1)

    print(f"\n[*] Publicando executável oficial em: {target_exe}")
    shutil.copy2(compiled_exe, target_exe)

    size_mb = target_exe.stat().st_size / (1024 * 1024)
    print("\n" + "=" * 80)
    print("[+] SUCESSO! EXECUTÁVEL STANDALONE GERADO E PRONTO PARA A EQUIPE:")
    print(f"    -> {target_exe}")
    print(f"    -> Tamanho total: {size_mb:.2f} MB (Arquivo .exe único com tudo embutido)")
    print("=" * 80)


if __name__ == "__main__":
    build()
