import os
import sys
import subprocess
import shutil
import tempfile
import time

def build():
    print("=" * 80)
    print("COMPILAÇÃO DO EXECUTÁVEL ÚNICO (.EXE INDEPENDENTE - ONEFILE)")
    print("=" * 80)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    main_script = os.path.join(base_dir, "app_conferencia_ferias.py")
    final_root_exe = os.path.join(base_dir, "Conferencia_Ferias_Cencosud.exe")

    # Usar pasta temporária do sistema para compilação (evita bloqueios de sincronização do OneDrive)
    temp_build_root = os.path.join(tempfile.gettempdir(), "cencosud_pyi_build")
    temp_work = os.path.join(temp_build_root, "work")
    temp_dist = os.path.join(temp_build_root, "dist")

    print(f"\n>>> Diretório de compilação isolado: {temp_build_root}")
    if os.path.exists(temp_build_root):
        shutil.rmtree(temp_build_root, ignore_errors=True)
    os.makedirs(temp_work, exist_ok=True)
    os.makedirs(temp_dist, exist_ok=True)

    if os.path.exists(final_root_exe):
        try:
            os.remove(final_root_exe)
        except Exception as e:
            print(f"Aviso ao remover executável antigo na raiz: {e}")

    # Comando PyInstaller com --onefile para arquivo 100% único
    cmd = [
        sys.executable,
        "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name", "Conferencia_Ferias_Cencosud",
        "--workpath", temp_work,
        "--distpath", temp_dist,
        "--collect-all", "customtkinter",
        "--collect-all", "openpyxl",
        "--add-data", f"{os.path.join(base_dir, 'assets')};assets",
        "--icon", os.path.join(base_dir, "assets", "cencosud.ico"),
        main_script
    ]

    print("\n>>> Executando PyInstaller (--onefile)...")
    print(" ".join(cmd))
    print("\nAguarde, gerando o arquivo executável único...")

    result = subprocess.run(cmd, cwd=base_dir)

    if result.returncode == 0:
        compiled_exe = os.path.join(temp_dist, "Conferencia_Ferias_Cencosud.exe")
        if os.path.exists(compiled_exe):
            print(f"\n>>> Movendo executável final para a raiz: {final_root_exe}")
            shutil.copy2(compiled_exe, final_root_exe)

            # Limpar diretório temporário
            print(">>> Limpando arquivos de compilação temporários...")
            time.sleep(1)
            shutil.rmtree(temp_build_root, ignore_errors=True)

            spec_file = os.path.join(base_dir, "Conferencia_Ferias_Cencosud.spec")
            if os.path.exists(spec_file):
                try: os.remove(spec_file)
                except: pass

            size_mb = os.path.getsize(final_root_exe) / (1024 * 1024)
            print("\n" + "=" * 80)
            print("SUCESSO ABSOLUTO! EXECUTÁVEL ÚNICO GERADO E PRONTO PARA USO")
            print(f"Executável: {final_root_exe}")
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

if __name__ == '__main__':
    build()
