"""
Script de Compilação do Executável Único (.EXE OneFile) - Central de Benefícios Cencosud.

Utiliza diretório temporário isolado fora do OneDrive para evitar bloqueios de sincronização
durante o empacotamento via PyInstaller e publica o executável pronto na raiz de VA e VT
e na pasta do Good Card.
"""

import os
import shutil
import subprocess
import sys
import tempfile
import time


def build():
    print("=" * 80)
    print("COMPILAÇÃO DO EXECUTÁVEL ÚNICO (.EXE INDEPENDENTE - ONEFILE)")
    print("CENTRAL EXECUTIVA DE AUDITORIA DE BENEFÍCIOS CENCOSUD")
    print("=" * 80)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    va_vt_root = os.path.dirname(base_dir)
    main_script = os.path.join(base_dir, "app_central_beneficios.py")
    exe_name = "Central_Beneficios_Cencosud.exe"

    # Local único oficial de publicação do executável
    targets = [
        os.path.join(va_vt_root, exe_name),
    ]

    # Pasta temporária isolada fora do OneDrive
    temp_build_root = os.path.join(tempfile.gettempdir(), "cencosud_beneficios_pyi_build")
    temp_work = os.path.join(temp_build_root, "work")
    temp_dist = os.path.join(temp_build_root, "dist")

    print(f"\n>>> Diretório de compilação isolado: {temp_build_root}")
    if os.path.exists(temp_build_root):
        shutil.rmtree(temp_build_root, ignore_errors=True)
    os.makedirs(temp_work, exist_ok=True)
    os.makedirs(temp_dist, exist_ok=True)

    for t in targets:
        if os.path.exists(t):
            try:
                os.remove(t)
            except Exception as e:
                print(f"Aviso ao remover executável antigo ({t}): {e}")

    cmd = [
        sys.executable,
        "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name", "Central_Beneficios_Cencosud",
        "--workpath", temp_work,
        "--distpath", temp_dist,
        "--collect-all", "customtkinter",
        "--collect-all", "PIL",
        "--collect-all", "openpyxl",
        "--collect-all", "odf",
        "--collect-all", "defusedxml",
        "--collect-all", "engines",
        "--add-data", f"{os.path.join(base_dir, 'assets')};assets",
        "--icon", os.path.join(base_dir, "assets", "cencosud.ico"),
        main_script,
    ]

    print("\n>>> Executando PyInstaller (--onefile)...")
    print(" ".join(cmd))
    print("\nAguarde, empacotando o aplicativo executável único...")

    result = subprocess.run(cmd, cwd=base_dir)

    if result.returncode == 0:
        compiled_exe = os.path.join(temp_dist, exe_name)
        if os.path.exists(compiled_exe):
            for t in targets:
                if os.path.exists(os.path.dirname(t)):
                    print(f">>> Publicando executável em: {t}")
                    shutil.copy2(compiled_exe, t)

            print(">>> Limpando arquivos temporários de compilação...")
            time.sleep(1)
            shutil.rmtree(temp_build_root, ignore_errors=True)

            spec_file = os.path.join(base_dir, "Central_Beneficios_Cencosud.spec")
            if os.path.exists(spec_file):
                try:
                    os.remove(spec_file)
                except Exception:
                    pass

            size_mb = os.path.getsize(targets[0]) / (1024 * 1024)
            print("\n" + "=" * 80)
            print("SUCESSO ABSOLUTO! EXECUTÁVEL ÚNICO GERADO E PRONTO PARA USO")
            for t in targets:
                if os.path.exists(t):
                    print(f"  -> {t}")
            print(f"Tamanho total: {size_mb:.2f} MB (Tudo embutido em um único arquivo .exe)")
            print("=" * 80)
        else:
            print(f"Erro: Arquivo compilado não localizado em {compiled_exe}")
            sys.exit(1)
    else:
        print("\n" + "=" * 80)
        print("FALHA NA COMPILAÇÃO DO EXECUTÁVEL.")
        print("=" * 80)
        sys.exit(1)


if __name__ == "__main__":
    build()
