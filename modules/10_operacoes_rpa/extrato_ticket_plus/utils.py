"""
Módulo utilitário para o Robô de Extração de Extratos GoodManager (Ticket Plus).
Responsável por:
- Carregamento de variáveis de ambiente (.env)
- Configuração de logging estruturado (console e arquivo)
- Gestão de pastas de saída e sanitização de caminhos no Windows
- Leitura e padronização da lista de matrículas e nomes (TXT, CSV, Excel na raiz ou em 'dados e matriculas p consulta')
"""

import os
import re
import csv
import shutil
import logging
from pathlib import Path
from typing import List, Optional, Tuple, Dict, Any

# Tenta carregar python-dotenv, com fallback manual para .env
try:
    from dotenv import load_dotenv
    _HAS_DOTENV = True
except ImportError:
    _HAS_DOTENV = False

# Tenta importar pandas para leitura de planilhas
try:
    import pandas as pd
    _HAS_PANDAS = True
except ImportError:
    _HAS_PANDAS = False


# Caminhos base do projeto
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "Relatórios extraídos"
DADOS_DIR = BASE_DIR / "dados e matriculas p consulta"
LOG_FILE = BASE_DIR / "execucao.log"
ENV_FILE = BASE_DIR / ".env"

# Carrega variáveis de ambiente
if _HAS_DOTENV and ENV_FILE.exists():
    load_dotenv(ENV_FILE)
elif ENV_FILE.exists():
    with open(ENV_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                os.environ.setdefault(key.strip(), val.strip().strip('"').strip("'"))

GOODMANAGER_COD_USUARIO = os.getenv("GOODMANAGER_COD_USUARIO", "")
GOODMANAGER_NOME_USUARIO = os.getenv("GOODMANAGER_NOME_USUARIO", "")
GOODMANAGER_SENHA = os.getenv("GOODMANAGER_SENHA", "")


def get_logger(name: str = "TicketExtratoRobo") -> logging.Logger:
    """
    Configura e retorna um logger com saída simultânea no console e em arquivo (execucao.log).
    """
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # Handler do console
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # Handler de arquivo
    try:
        file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8", mode="a")
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except Exception as e:
        logger.warning(f"Não foi possível criar o FileHandler em {LOG_FILE}: {e}")

    return logger


def sanitize_filename(name: str) -> str:
    """
    Sanitiza uma string para que possa ser usada com segurança como nome de pasta/arquivo no Windows.
    Remove caracteres proibidos: <>:"/\\|?* e caracteres de controle, aparando espaços e pontos finais.
    """
    if not name:
        return "SEM_NOME"
    sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', str(name))
    sanitized = re.sub(r'\s+', ' ', sanitized)
    sanitized = sanitized.strip('. ')
    return sanitized or "SEM_NOME"


def get_employee_output_dir(employee_identifier: str) -> Path:
    """
    Cria e retorna o diretório individual do colaborador dentro de 'Relatórios extraídos'.
    Exemplo: 'Relatórios extraídos/7887 - NOME DO COLABORADOR'
    """
    safe_name = sanitize_filename(employee_identifier)
    target_dir = OUTPUT_DIR / safe_name
    target_dir.mkdir(parents=True, exist_ok=True)
    return target_dir


def get_report_pdf_path(employee_identifier: str, period_label: str) -> Path:
    """
    Retorna o caminho completo seguro para o PDF de um período específico.
    Exemplo: 'Relatórios extraídos/7887 - NOME/Extrato_Outubro_2026.pdf'
    """
    emp_dir = get_employee_output_dir(employee_identifier)
    safe_period = sanitize_filename(period_label)
    filename = f"Extrato_{safe_period}.pdf"
    return emp_dir / filename


def verificar_cache_colaborador(matricula: str) -> Optional[Path]:
    """
    Verifica se a matrícula já possui os 2 PDFs extraídos (Mês Atual e Compras Próximo Período)
    em alguma subpasta dentro de 'Relatórios extraídos'.
    Retorna o Path da pasta se ambos existirem e forem válidos (> 1KB), ou None se precisar extrair.
    """
    if not OUTPUT_DIR.exists():
        return None

    matr_str = str(matricula).strip()
    for pasta in OUTPUT_DIR.iterdir():
        if not pasta.is_dir():
            continue
        # Pasta segue formato 'MATRICULA - NOME' ou 'MATRICULA'
        nome_pasta = pasta.name
        partes = nome_pasta.split(" - ", 1)
        if partes[0].strip() == matr_str or nome_pasta.strip() == matr_str:
            pdfs = list(pasta.glob("*.pdf"))
            tem_mes_atual = any("Mes_Atual" in p.name and p.stat().st_size > 1024 for p in pdfs)
            tem_proximo_periodo = any("Compras_Proximo_Periodo" in p.name and p.stat().st_size > 1024 for p in pdfs)
            if tem_mes_atual and tem_proximo_periodo:
                return pasta

    return None


def _copiar_se_bloqueado(origem: Path) -> Path:
    """
    Se o arquivo Excel estiver aberto pelo usuário ou travado pelo OneDrive,
    cria uma cópia temporária e retorna o caminho da cópia para leitura limpa.
    """
    try:
        # Testa abertura
        with open(origem, "rb"):
            pass
        return origem
    except (PermissionError, IOError):
        temp_dest = origem.parent / f"_temp_leitura_{origem.stem}{origem.suffix}"
        try:
            shutil.copy2(origem, temp_dest)
            return temp_dest
        except Exception:
            return origem


def load_matriculas_e_nomes(filepath: Optional[str or Path] = None) -> Tuple[List[str], Dict[str, str]]:
    """
    Carrega a lista de matrículas e o dicionário de mapeamento {matricula: nome}
    a partir de planilhas Excel ou arquivos de texto.
    
    Ordem de busca automática se filepath for None:
      1. Planilhas Excel (.xlsx/.xls) em 'dados e matriculas p consulta/'
      2. Planilhas Excel (.xlsx/.xls) na pasta raiz do robô
      3. matriculas.csv ou matriculas.txt na raiz do robô
    """
    logger = get_logger()
    target_path: Optional[Path] = None

    raw_pairs: List[Tuple[str, str]] = []

    def _processar_dataframe_excel(df: Any, nome_arquivo: str):
        col_matr = None
        col_nome = None

        # 1. Procura nas colunas padrão
        for col in df.columns:
            c_clean = str(col).lower().strip()
            if not col_matr and any(k in c_clean for k in ["matrícula", "matricula", "chapa"]):
                col_matr = col
            if not col_nome and any(k in c_clean for k in ["nome", "colaborador", "funcionario"]):
                col_nome = col

        # 2. Se não achou na linha 0, verifica se a primeira linha (row 0) contém os nomes das colunas
        if not col_matr and len(df) > 0:
            for c in df.columns:
                val = str(df.iloc[0][c]).lower().strip()
                if not col_matr and any(k in val for k in ["matrícula", "matricula", "chapa"]):
                    col_matr = c
                if not col_nome and any(k in val for k in ["nome", "colaborador", "funcionario"]):
                    col_nome = c
            if col_matr is not None:
                df_dados = df.iloc[1:].copy()
            else:
                df_dados = df
        else:
            df_dados = df

        if col_matr is None and len(df_dados.columns) > 0:
            col_matr = df_dados.columns[0]

        logger.info(f"[{nome_arquivo}] Coluna Matrícula='{col_matr}', Coluna Nome='{col_nome}', Linhas={len(df_dados)}")

        for _, row in df_dados.iterrows():
            m_raw = row[col_matr]
            if pd.isna(m_raw) or str(m_raw).strip() in ["", "nan", "None"]:
                continue

            m_str = str(m_raw).strip()
            if m_str.endswith(".0"):
                m_str = m_str[:-2]
            m_clean = "".join(c for c in m_str if c.isdigit())

            n_clean = ""
            if col_nome and pd.notna(row[col_nome]):
                n_clean = str(row[col_nome]).strip()

            if m_clean:
                raw_pairs.append((m_clean, n_clean))

    if filepath:
        target_paths = [Path(filepath)]
    else:
        target_paths = []
        # 1. Prioridade: subpasta 'dados e matriculas p consulta' (todas as planilhas)
        if DADOS_DIR.exists():
            planilhas_dados = [p for p in DADOS_DIR.glob("*") if p.suffix.lower() in [".xlsx", ".xls", ".xlsb"] and not p.name.startswith(("~", "_temp_"))]
            if planilhas_dados:
                target_paths = sorted(planilhas_dados, key=lambda x: x.name)

        # 2. Arquivos na raiz do robô
        if not target_paths:
            for candidate in ["matriculas.xlsx", "matriculas.xls", "matriculas.xlsb", "matriculas.csv", "matriculas.txt"]:
                p = BASE_DIR / candidate
                if p.exists() and p.stat().st_size > 0:
                    target_paths = [p]
                    break

    if not target_paths:
        logger.info("Nenhum arquivo de matrículas especificado ou encontrado automaticamente.")
        return [], {}

    for target_path in target_paths:
        logger.info(f"Lendo base de dados/matrículas de: {target_path.name}")
        ext = target_path.suffix.lower()
        try:
            if ext in [".xlsx", ".xls", ".xlsb"]:
                if not _HAS_PANDAS:
                    logger.error("Pandas não está disponível para ler arquivos Excel.")
                    continue

                arquivo_leitura = _copiar_se_bloqueado(target_path)
                try:
                    engine = "pyxlsb" if ext == ".xlsb" else None
                    df = pd.read_excel(arquivo_leitura, engine=engine, dtype=str)
                    _processar_dataframe_excel(df, target_path.name)
                finally:
                    if arquivo_leitura != target_path and arquivo_leitura.exists():
                        try:
                            arquivo_leitura.unlink()
                        except Exception:
                            pass

            elif ext == ".csv":
                with open(target_path, "r", encoding="utf-8", errors="ignore") as f:
                    sample = f.read(2048)
                    f.seek(0)
                    delimiter = ";" if sample.count(";") > sample.count(",") else ","
                    reader = csv.reader(f, delimiter=delimiter)
                    header = next(reader, None)

                    col_matr_idx = 0
                    col_nome_idx = None
                    if header:
                        for idx, col in enumerate(header):
                            c_clean = col.lower().strip()
                            if any(k in c_clean for k in ["matricula", "matrícula", "chapa", "id", "cod"]):
                                col_matr_idx = idx
                            if any(k in c_clean for k in ["nome", "colaborador"]):
                                col_nome_idx = idx

                    for row in reader:
                        if row and len(row) > col_matr_idx:
                            m_str = row[col_matr_idx].strip()
                            if m_str.endswith(".0"):
                                m_str = m_str[:-2]
                            m_clean = "".join(c for c in m_str if c.isdigit())
                            n_clean = row[col_nome_idx].strip() if (col_nome_idx and len(row) > col_nome_idx) else ""
                            if m_clean:
                                raw_pairs.append((m_clean, n_clean))

            elif ext == ".txt":
                with open(target_path, "r", encoding="utf-8", errors="ignore") as f:
                    for line in f:
                        item = line.strip()
                        if item and not item.startswith("#"):
                            m_str = item.split(";")[0].split(",")[0].strip()
                            if m_str.endswith(".0"):
                                m_str = m_str[:-2]
                            m_clean = "".join(c for c in m_str if c.isdigit())
                            if m_clean:
                                raw_pairs.append((m_clean, ""))

        except Exception as e:
            logger.error(f"Erro ao carregar matrículas/nomes de {target_path}: {e}")
            continue

    # Deduplica preservando ordem
    cleaned_matriculas: List[str] = []
    mapa_nomes: Dict[str, str] = {}
    seen = set()

    for m, n in raw_pairs:
        if m not in seen:
            seen.add(m)
            cleaned_matriculas.append(m)
            if n:
                mapa_nomes[m] = n
        elif n and m not in mapa_nomes:
            mapa_nomes[m] = n

    logger.info(f"Total de {len(cleaned_matriculas)} matrículas únicas carregadas (com {len(mapa_nomes)} nomes mapeados).")
    return cleaned_matriculas, mapa_nomes


def load_matriculas(filepath: Optional[str or Path] = None) -> List[str]:
    """
    Função legada mantendo compatibilidade direta: retorna apenas a lista de matrículas.
    """
    matriculas, _ = load_matriculas_e_nomes(filepath)
    return matriculas
