# -*- coding: utf-8 -*-
"""
anonimizar_planilhas.py  —  Anonimizador determinístico para planilhas (Excel)
e arquivos (CSV) de faturamento (Ativos) e movimentação (Fopag) de RH.
"""

from __future__ import annotations
import argparse
import hashlib
import os
import random
import re
import sys
import csv
import unicodedata
from copy import copy
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import pandas as pd

# ---------------------------------------------------------------------------
# CONFIG — colunas sensíveis (header normalizado -> tipo de anonimização)
# ---------------------------------------------------------------------------
COLUNAS_SENSIVEIS = {
    # CPF do beneficiário
    "cpf": "cpf",
    "cpf beneficiario": "cpf",
    "cpf do beneficiario": "cpf",
    "nr cpf": "cpf",
    "numero cpf": "cpf",
    "n cpf": "cpf",
    # CPF do titular
    "cpf titular": "cpf",
    "cpf do titular": "cpf",
    # Nome
    "nome": "nome",
    "nome beneficiario": "nome",
    "nome do beneficiario": "nome",
    "nome completo": "nome",
    "nome do titular": "nome",
    "beneficiario": "nome",
    "segurado": "nome",
    # Matrícula
    "matricula": "matricula",
    "matrícula": "matricula",
    "matricula funcional": "matricula",
    "matricula colaborador": "matricula",
    "chapa": "matricula",
    "registro": "matricula",
    # Carteirinha / cartão
    "carteirinha": "carteirinha",
    "carteira": "carteirinha",
    "cartao": "carteirinha",
    "cartão": "carteirinha",
    "numero carteirinha": "carteirinha",
    "n carteirinha": "carteirinha",
    "codigo beneficiario": "carteirinha",
    "id beneficiario": "carteirinha",
    # Nascimento
    "data nascimento": "nascimento",
    "data de nascimento": "nascimento",
    "nascimento": "nascimento",
    "dt nascimento": "nascimento",
    "dt nasc": "nascimento",
    # E-mail / telefone / RG / endereço
    "email": "email",
    "e-mail": "email",
    "telefone": "telefone",
    "celular": "telefone",
    "rg": "rg",
    "endereco": "endereco",
    "endereço": "endereco",
    "logradouro": "endereco",
}

PRIMEIROS_NOMES = [
    "ANA", "BRUNO", "CARLOS", "DANIELA", "EDUARDO", "FERNANDA", "GUSTAVO", "HELENA",
    "IGOR", "JULIANA", "KARINA", "LUCAS", "MARIANA", "NATALIA", "OTAVIO", "PAULA",
    "QUEILA", "RAFAEL", "SANDRA", "THIAGO", "ULISSES", "VANESSA", "WAGNER", "XENIA",
    "YASMIN", "ZULMA", "ANDRE", "BEATRIZ", "CESAR", "DIANA", "EMERSON", "FLAVIA",
    "GABRIEL", "HUGO", "ISABELA", "JOAO", "KAUE", "LARISSA", "MARCOS", "NICOLE",
    "OSCAR", "PRISCILA", "RENATO", "SOFIA", "TATIANE", "VITOR", "WESLEY", "YURI",
]
SOBRENOMES = [
    "SILVA", "SANTOS", "OLIVEIRA", "SOUZA", "RODRIGUES", "FERREIRA", "ALVES", "PEREIRA",
    "LIMA", "GOMES", "COSTA", "RIBEIRO", "MARTINS", "CARVALHO", "ALMEIDA", "LOPES",
    "SOARES", "FERNANDES", "VIEIRA", "BARBOSA", "ROCHA", "DIAS", "NASCIMENTO", "ARAUJO",
    "MOREIRA", "NUNES", "MARQUES", "MACHADO", "MENDES", "FREITAS", "CARDOSO", "RAMOS",
    "GONCALVES", "SANTANA", "TEIXEIRA", "CORREIA", "CUNHA", "PINTO", "MELO", "BATISTA",
]


# ---------------------------------------------------------------------------
# Utilitários
# ---------------------------------------------------------------------------
def normaliza_header(s) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s


def hash_int(seed: str, valor) -> int:
    h = hashlib.sha256(f"{seed}|{valor}".encode("utf-8")).hexdigest()
    return int(h[:16], 16)


def calc_dv_cpf(nove_digitos: str) -> str:
    def dv(base):
        s = sum(int(d) * p for d, p in zip(base, range(len(base) + 1, 1, -1)))
        r = (s * 10) % 11
        return "0" if r == 10 else str(r)

    d1 = dv(nove_digitos)
    d2 = dv(nove_digitos + d1)
    return d1 + d2


# ---------------------------------------------------------------------------
# Geradores fake
# ---------------------------------------------------------------------------
class Anonimizador:
    def __init__(self, seed: str = "lovable-anon-v1"):
        self.seed = seed
        self.cache_cpf: dict[str, str] = {}
        self.cache_nome: dict[str, str] = {}
        self.cache_matricula: dict[str, str] = {}
        self.cache_carteirinha: dict[str, str] = {}

    def fake_cpf(self, original) -> str:
        if original is None or (isinstance(original, float) and original != original):
            return original
        chave = re.sub(r"\D", "", str(original))
        if not chave:
            return original
        if chave in self.cache_cpf:
            return self.cache_cpf[chave]
        rng = random.Random(hash_int(self.seed + "::cpf", chave))
        nove = "".join(str(rng.randint(0, 9)) for _ in range(9))
        if len(set(nove)) == 1:
            nove = nove[:-1] + ("1" if nove[-1] != "1" else "2")
        cpf = nove + calc_dv_cpf(nove)
        masc = f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
        out = masc if ("." in str(original) or "-" in str(original)) else cpf
        self.cache_cpf[chave] = out
        return out

    def fake_nome(self, original) -> str:
        if original is None:
            return original
        chave = str(original).strip().upper()
        if not chave:
            return original
        if chave in self.cache_nome:
            return self.cache_nome[chave]
        rng = random.Random(hash_int(self.seed + "::nome", chave))
        partes = chave.split()
        n_sob = max(1, min(3, len(partes) - 1)) if len(partes) > 1 else 1
        nome = rng.choice(PRIMEIROS_NOMES)
        sobs = [rng.choice(SOBRENOMES) for _ in range(n_sob)]
        out = " ".join([nome] + sobs)
        self.cache_nome[chave] = out
        return out

    def fake_matricula(self, original):
        if original is None or (isinstance(original, float) and original != original):
            return original
        s = str(original).strip()
        if not s:
            return original
        if s in self.cache_matricula:
            out = self.cache_matricula[s]
        else:
            rng = random.Random(hash_int(self.seed + "::mat", s))
            so_num = re.sub(r"\D", "", s)
            n = len(so_num) if so_num else 6
            n = max(4, min(n, 10))
            primeiro = rng.randint(1, 9)
            resto = "".join(str(rng.randint(0, 9)) for _ in range(n - 1))
            out = f"{primeiro}{resto}"
            self.cache_matricula[s] = out
        if isinstance(original, int):
            try:
                return int(out)
            except:
                return out
        return out

    def fake_carteirinha(self, original):
        if original is None or (isinstance(original, float) and original != original):
            return original
        s = str(original).strip()
        if not s:
            return original
        if s in self.cache_carteirinha:
            out = self.cache_carteirinha[s]
        else:
            rng = random.Random(hash_int(self.seed + "::cart", s))
            so_num = re.sub(r"\D", "", s)
            if so_num:
                n = len(so_num)
                out = "".join(str(rng.randint(0, 9)) for _ in range(n))
            else:
                out = "".join(rng.choice("ABCDEFGHJKLMNPQRSTUVWXYZ23456789") for _ in range(len(s)))
            self.cache_carteirinha[s] = out
        if isinstance(original, int):
            try:
                return int(out)
            except:
                return out
        return out

    def fake_nascimento(self, original, jitter_dias=15):
        if original is None:
            return original
        if isinstance(original, datetime):
            base = original
        else:
            try:
                base = datetime.strptime(str(original)[:10], "%Y-%m-%d")
            except Exception:
                try:
                    base = datetime.strptime(str(original)[:10], "%d/%m/%Y")
                except Exception:
                    return original
        rng = random.Random(hash_int(self.seed + "::nasc", base.isoformat()))
        delta = rng.randint(-jitter_dias, jitter_dias)
        nova = base + timedelta(days=delta)
        return nova if isinstance(original, datetime) else nova.strftime("%d/%m/%Y")

    def fake_email(self, original):
        if not original: return original
        rng = random.Random(hash_int(self.seed + "::email", str(original)))
        return f"user{rng.randint(1000, 99999)}@exemplo.com"

    def fake_telefone(self, original):
        if not original: return original
        rng = random.Random(hash_int(self.seed + "::tel", str(original)))
        return f"(11) 9{rng.randint(1000, 9999)}-{rng.randint(1000, 9999)}"

    def fake_rg(self, original):
        if not original: return original
        rng = random.Random(hash_int(self.seed + "::rg", str(original)))
        return f"{rng.randint(10, 99)}.{rng.randint(100, 999)}.{rng.randint(100, 999)}-{rng.randint(0, 9)}"

    def fake_endereco(self, original):
        if not original: return original
        rng = random.Random(hash_int(self.seed + "::end", str(original)))
        ruas = ["RUA DAS FLORES", "AV BRASIL", "RUA XV DE NOVEMBRO", "AV PAULISTA", "RUA DOS PINHEIROS"]
        return f"{rng.choice(ruas)}, {rng.randint(1, 9999)}"

    def aplica(self, tipo, valor):
        return {
            "cpf": self.fake_cpf,
            "nome": self.fake_nome,
            "matricula": self.fake_matricula,
            "carteirinha": self.fake_carteirinha,
            "nascimento": self.fake_nascimento,
            "email": self.fake_email,
            "telefone": self.fake_telefone,
            "rg": self.fake_rg,
            "endereco": self.fake_endereco,
        }[tipo](valor)


# ---------------------------------------------------------------------------
# Processamento de Arquivos CSV
# ---------------------------------------------------------------------------
def processar_csv(caminho: Path, anon: Anonimizador, amostra=None):
    print(f"\n📂 {caminho.name} [CSV]")

    encodings_para_testar = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
    linhas_originais = []
    encoding_usado = None
    dialect_usado = None

    for enc in encodings_para_testar:
        try:
            with open(caminho, 'r', encoding=enc, newline='') as f:
                amostra_txt = f.read(8192)
                if not amostra_txt:
                    return

                f.seek(0)
                try:
                    # ✅ Aviso de Syntax corrigido retirando o escape desnecessário do pipe '|'
                    dialect = csv.Sniffer().sniff(amostra_txt, delimiters=';,|\t')
                except csv.Error:
                    if amostra_txt.count(';') > amostra_txt.count(','):
                        class ExcelSemicolon(csv.excel):
                            delimiter = ';'

                        dialect = ExcelSemicolon
                    else:
                        dialect = csv.excel

                leitor = csv.reader(f, dialect=dialect)
                linhas_originais = list(leitor)

            encoding_usado = enc
            dialect_usado = dialect
            break
        except UnicodeDecodeError:
            continue

    if not linhas_originais:
        print("   ❌ falha ao abrir: não foi possível ler o arquivo CSV ou ele está vazio.")
        return

    melhor_linha, melhor_score = 0, -1
    for i, row in enumerate(linhas_originais[:50]):
        score = sum(1 for cell in row if normaliza_header(cell) in COLUNAS_SENSIVEIS)
        if score > melhor_score:
            melhor_linha, melhor_score = i, score

    if melhor_score == 0:
        print("   — ⚠️ Nenhuma coluna sensível encontrada.")
        return

    mapa = {}
    for col_idx, cell in enumerate(linhas_originais[melhor_linha]):
        tipo = COLUNAS_SENSIVEIS.get(normaliza_header(cell))
        if tipo:
            mapa[col_idx] = tipo

    cells_alteradas = 0
    linhas_processadas = 0
    max_row = len(linhas_originais)

    if amostra:
        max_row = min(max_row, melhor_linha + 1 + amostra)

    for i in range(melhor_linha + 1, max_row):
        row = linhas_originais[i]
        for col, tipo in mapa.items():
            if col < len(row):
                valor_atual = row[col]
                if not valor_atual:
                    continue

                novo = str(anon.aplica(tipo, valor_atual))
                if novo != valor_atual:
                    row[col] = novo
                    cells_alteradas += 1
        linhas_processadas += 1

    saida = caminho.with_name(caminho.stem + "_ANON.csv")
    with open(saida, 'w', encoding=encoding_usado, newline='') as f:
        escritor = csv.writer(f, dialect=dialect_usado)
        escritor.writerows(linhas_originais[:max_row])

    marca = "✓" if cells_alteradas else "—"
    print(f"   {marca} Base CSV: {linhas_processadas} linhas, {cells_alteradas} células anonimizadas")
    print(f"   💾 salvo: {saida.name}")


# ---------------------------------------------------------------------------
# Processamento de Arquivos XLS (Legado / HTML Disfarçado)
# ---------------------------------------------------------------------------
def processar_xls_legado(caminho: Path, anon: Anonimizador, amostra=None):
    print(f"\n📂 {caminho.name} [XLS LEGADO]")

    dfs_dict = {}

    # Tentativa 1: Ler como arquivo .xls real
    try:
        xls = pd.ExcelFile(caminho, engine='xlrd')
        for nome_aba in xls.sheet_names:
            dfs_dict[nome_aba] = pd.read_excel(xls, sheet_name=nome_aba, header=None)

    except Exception as e_xls:
        # Tentativa 2: Ler como HTML disfarçado (comum em arquivos da SulAmérica, SAP, etc)
        try:
            # Pega as tabelas que estão dentro do HTML
            tabelas_html = pd.read_html(caminho, encoding='utf-8')
            for i, df_html in enumerate(tabelas_html):
                dfs_dict[f"Tabela_{i + 1}"] = df_html
            print("   ⚠️ O arquivo estava disfarçado de XLS, mas era HTML. Leitura HTML bem-sucedida.")
        except Exception as e_html:
            print(f"   ❌ Erro ao abrir: Não é um .xls válido nem HTML legível.")
            print(f"      (Info Erro XLS: {repr(e_xls)})")
            print(f"      (Info Erro HTML: {repr(e_html)})")
            return

    # Processar o dicionário de abas extraído
    try:
        saida = caminho.with_name(caminho.stem + "_ANON.xlsx")

        with pd.ExcelWriter(saida, engine='openpyxl') as writer:
            abas_processadas = 0
            for nome_aba, df_full in dfs_dict.items():
                if df_full.empty or df_full.shape[0] < 2:
                    continue

                # Procura a linha de cabeçalho nas 50 primeiras linhas
                melhor_linha, melhor_score = 0, -1
                for idx, row in df_full.head(50).iterrows():
                    score = sum(1 for cell in row if not pd.isna(cell) and normaliza_header(cell) in COLUNAS_SENSIVEIS)
                    if score > melhor_score:
                        melhor_linha, melhor_score = idx, score

                header_idx = melhor_linha if melhor_score > 0 else 0

                # Ajusta o DataFrame pra considerar o cabeçalho encontrado
                df = df_full.iloc[header_idx + 1:].copy()
                df.columns = df_full.iloc[header_idx].values
                # Limpeza: remove colunas com nome nulo (comum no parsing de HTML pra Excel)
                df = df.loc[:, df.columns.notna()]
                df.reset_index(drop=True, inplace=True)

                if amostra:
                    df = df.head(amostra)

                cells_alteradas = 0

                # Aplica Anonimização
                for col in df.columns:
                    tipo = COLUNAS_SENSIVEIS.get(normaliza_header(col))
                    if tipo:
                        def aplica_anon(val):
                            if pd.isna(val) or str(val).strip() == "":
                                return val
                            return anon.aplica(tipo, val)

                        cells_alteradas += df[col].notna().sum()
                        df[col] = df[col].apply(aplica_anon)

                # Excel/openpyxl não suportam nomes de abas maiores que 31 caracteres
                nome_aba_limpo = str(nome_aba)[:31]
                df.to_excel(writer, sheet_name=nome_aba_limpo, index=False)
                abas_processadas += 1

                marca = "✓" if cells_alteradas else "—"
                print(f"   {marca} aba '{nome_aba_limpo}': {len(df)} linhas, {cells_alteradas} células anonimizadas")

            if abas_processadas > 0:
                print(f"   💾 salvo (convertido para XLSX): {saida.name}")
            else:
                print(f"   ⚠️ Nenhuma aba processada ou vazia.")

    except Exception as e:
        print(f"   ❌ erro ao salvar o arquivo: {repr(e)}")


# ---------------------------------------------------------------------------
# Processamento por aba (EXCEL XLSX/XLSM)
# ---------------------------------------------------------------------------
def detectar_linha_header(ws, max_scan=50):
    melhor_linha, melhor_score = 1, -1
    for row in range(1, min(max_scan, ws.max_row) + 1):
        score = 0
        for cell in ws[row]:
            if normaliza_header(cell.value) in COLUNAS_SENSIVEIS:
                score += 1
        if score > melhor_score:
            melhor_linha, melhor_score = row, score
    return melhor_linha if melhor_score > 0 else 1


def mapear_colunas_sensiveis(ws, linha_header):
    mapa = {}
    for cell in ws[linha_header]:
        tipo = COLUNAS_SENSIVEIS.get(normaliza_header(cell.value))
        if tipo:
            mapa[cell.column] = tipo
    return mapa


def processar_aba(ws, anon: Anonimizador, amostra=None):
    linha_header = detectar_linha_header(ws)
    mapa = mapear_colunas_sensiveis(ws, linha_header)
    if not mapa:
        return 0, 0
    cells_alteradas = 0
    linhas_processadas = 0
    max_row = ws.max_row
    if amostra:
        max_row = min(max_row, linha_header + amostra)
    for row in range(linha_header + 1, max_row + 1):
        for col, tipo in mapa.items():
            cell = ws.cell(row=row, column=col)
            if cell.value is None or cell.value == "":
                continue
            novo = anon.aplica(tipo, cell.value)
            if novo != cell.value:
                cell.value = novo
                cells_alteradas += 1
        linhas_processadas += 1
    return linhas_processadas, cells_alteradas


def processar_arquivo_excel(caminho: Path, anon: Anonimizador, amostra=None):
    print(f"\n📂 {caminho.name} [EXCEL XLSX]")
    try:
        wb = load_workbook(caminho)
    except Exception as e:
        print(f"   ❌ falha ao abrir: {e}")
        return
    for ws in wb.worksheets:
        linhas, alt = processar_aba(ws, anon, amostra)
        marca = "✓" if alt else "—"
        print(f"   {marca} aba '{ws.title[:31]}': {linhas} linhas, {alt} células anonimizadas")
    saida = caminho.with_name(caminho.stem + "_ANON.xlsx")
    wb.save(saida)
    print(f"   💾 salvo: {saida.name}")


def coletar_arquivos(alvo: Path):
    if alvo.is_file():
        return [alvo] if alvo.suffix.lower() in (".xlsx", ".xlsm", ".xls", ".csv") else []

    arquivos = []
    for ext in ("*.xlsx", "*.xls", "*.csv"):
        arquivos.extend(
            p for p in alvo.rglob(ext)
            if not p.name.endswith("_ANON.xlsx")
            and not p.name.endswith("_ANON.csv")
            and not p.name.startswith("~$")
        )
    return sorted(arquivos)


# ---------------------------------------------------------------------------
def main():
    # ============================================================
    # 👇 EDITE AQUI: caminho da pasta OU arquivo
    # ============================================================
    CAMINHO = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\anonimizar planilhas"
    AMOSTRA = None  # ou um número, ex: 30, pra limitar linhas
    SEED = "lovable-anon-v1"
    # ============================================================

    alvo = Path(CAMINHO).expanduser().resolve()
    if not alvo.exists():
        print(f"❌ caminho não existe: {alvo}")
        sys.exit(1)

    arquivos = coletar_arquivos(alvo)
    if not arquivos:
        print("❌ nenhum .xlsx, .xls ou .csv encontrado.")
        sys.exit(1)

    anon = Anonimizador(seed=SEED)
    print(f"🔐 seed: {SEED!r}   |   {len(arquivos)} arquivo(s)")

    for f in arquivos:
        if f.suffix.lower() == ".csv":
            processar_csv(f, anon, amostra=AMOSTRA)
        elif f.suffix.lower() == ".xls":
            processar_xls_legado(f, anon, amostra=AMOSTRA)
        else:
            processar_arquivo_excel(f, anon, amostra=AMOSTRA)

    print("\n✅ pronto. Arquivos *_ANON.* prontos pra compartilhar.")
    print(f"   CPFs únicos mapeados:        {len(anon.cache_cpf)}")
    print(f"   Nomes únicos mapeados:       {len(anon.cache_nome)}")
    print(f"   Matrículas únicas mapeadas:  {len(anon.cache_matricula)}")
    print(f"   Carteirinhas únicas mapeadas:{len(anon.cache_carteirinha)}")


if __name__ == "__main__":
    main()