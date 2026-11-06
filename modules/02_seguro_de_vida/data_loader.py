"""
Módulo de Ingestão e Tratamento de Dados para a Conferência de Seguro de Vida.
Lê Nomina, Relatório de Oposições, FOPA (Descontos), Catálogo de Optantes Complementares,
Cadastro de Diretores e Fatura da Seguradora,
garantindo tipagem precisa, preservação de zeros à esquerda em CPFs/matrículas e suporte a encodings.
"""

import os
import glob
import json
import re
import unicodedata
from pathlib import Path
from typing import Dict, Any, List, Optional, Callable
import openpyxl
import pandas as pd

from config import (
    RAW_DIR,
    BASE_DIR,
    CADASTRO_DIRETORES_JSON,
    OPTANTES_COMPLEMENTAR_JSON,
    EVENTO_SEGURO_NORMAL,
    EVENTO_SEGURO_COMP
)


def normalize_str(s: Any) -> str:
    """Remove acentos, espaços extras e converte para minúsculas para buscas resilientes."""
    if s is None or pd.isna(s):
        return ""
    txt = str(s).strip()
    return "".join(c for c in unicodedata.normalize("NFD", txt) if unicodedata.category(c) != "Mn").lower()


def format_cpf(val: Any) -> str:
    """
    Higieniza e formata o CPF no padrão 000.000.000-00,
    garantindo a preservação absoluta de zeros à esquerda (11 dígitos).
    """
    if val is None or pd.isna(val):
        return ""
    digits = re.sub(r"\D", "", str(val))
    if not digits:
        return ""
    digits = digits.zfill(11)
    if len(digits) == 11:
        return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"
    return digits


def format_matricula(val: Any) -> int:
    """Converte matrícula para número inteiro limpo."""
    if val is None or pd.isna(val):
        return 0
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        return 0


def parse_numeric_currency(val: Any) -> float:
    """
    Higieniza e converte strings/números de valores monetários para float com 2 casas decimais.
    Preserva rigorosamente:
    - Floats e inteiros nativos
    - Strings com ponto decimal: '2271.41' -> 2271.41
    - Strings com padrão brasileiro: '2.271,41' -> 2271.41 ou '2271,41' -> 2271.41
    - Strings nulas, hífens ou inválidas -> 0.0
    """
    if val is None or pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace("-", "0")
    if not val_str or val_str.lower() in ["nan", "none", "null"]:
        return 0.0
    # Se contiver tanto ponto quanto vírgula (ex: 1.234,56 ou 1,234.56)
    if "." in val_str and "," in val_str:
        if val_str.rfind(",") > val_str.rfind("."):
            # Padrão brasileiro: 1.234,56 -> 1234.56
            val_str = val_str.replace(".", "").replace(",", ".")
        else:
            # Padrão internacional: 1,234.56 -> 1234.56
            val_str = val_str.replace(",", "")
    elif "," in val_str:
        # Padrão brasileiro sem separador de milhar: 1234,56 -> 1234.56
        val_str = val_str.replace(",", ".")
    # Se apenas contiver ponto, já é o separador decimal padrão (ex: 2271.41) -> preserva o ponto!
    try:
        return round(float(val_str), 2)
    except (ValueError, TypeError):
        return 0.0


class DataLoader:
    """Classe responsável por carregar e sanitizar todas as bases de entrada."""

    def __init__(
        self,
        raw_dir: Optional[Path] = None,
        raw_map: Optional[Dict[str, Optional[str]]] = None,
        log_cb: Optional[Callable[[str], None]] = None,
        progress_cb: Optional[Callable[[float, str], None]] = None,
    ):
        self._custom_raw = raw_dir is not None
        self.raw_dir = Path(raw_dir) if raw_dir else RAW_DIR
        if not self._custom_raw and not self.raw_dir.exists() and BASE_DIR.exists():
            for sub in BASE_DIR.iterdir():
                if sub.is_dir() and "bruto" in normalize_str(sub.name):
                    self.raw_dir = sub
                    break
        self.raw_map = raw_map or {}
        self.log_cb = log_cb
        self.progress_cb = progress_cb
        self._diretores_com_gestor: List[Dict[str, Any]] = []
        self._aba_com_gestor_encontrada: bool = False

    def _log(self, msg: str) -> None:
        print(msg)
        if self.log_cb:
            self.log_cb(msg)

    def _progress(self, pct: float, msg: str) -> None:
        if self.progress_cb:
            self.progress_cb(pct, msg)

    def _find_files(self, patterns: List[str]) -> List[Path]:
        """Busca arquivos pelos padrões ignorando temporários de bloqueio do Excel (~$)."""
        found: List[Path] = []
        if self.raw_dir.exists():
            for pat in patterns:
                found.extend(f for f in self.raw_dir.glob(pat) if not f.name.startswith("~$"))
        if not found and not self._custom_raw and BASE_DIR.exists():
            for pat in patterns:
                found.extend(f for f in BASE_DIR.glob(f"**/{pat}") if not f.name.startswith("~$"))
        found.sort(key=lambda p: p.stat().st_mtime if p.exists() else 0.0, reverse=True)
        return found

    def load_all(self) -> Dict[str, Any]:
        """Carrega e retorna todas as bases estruturadas com batimento trilateral."""
        self._log(f"[*] Diretório de origem: {self.raw_dir}")
        self._progress(0.05, "Lendo planilha Nomina e aba 'Com gestor'...")
        df_nomina = self.load_nomina()
        self._progress(0.18, "Carregando Relatório de Oposições ao Seguro...")
        df_oposicao = self.load_oposicao()
        self._progress(0.24, "Carregando descontos em folha (FOPA)...")
        df_fopa = self.load_fopa()
        self._progress(0.28, "Carregando catálogo de Optantes Complementares e Diretores...")
        optantes_comp = self.load_optantes_complementar()
        diretores = self.load_diretores()
        self._progress(0.32, "Preparando base de conciliação da Fatura da Seguradora...")
        df_fatura = self.load_fatura_seguradora(df_nomina)

        return {
            "nomina": df_nomina,
            "oposicao": df_oposicao,
            "fopa": df_fopa,
            "optantes_complementar": optantes_comp,
            "diretores": diretores,
            "fatura_seguradora": df_fatura
        }

    @staticmethod
    def _clean_str_series(series: pd.Series, default: str = "") -> pd.Series:
        """Higieniza séries textuais evitando conversões literais de None/NaN para 'None'/'nan'."""
        s = series.fillna(default).astype(str).str.strip()
        return s.replace({"None": default, "nan": default, "NaN": default, "<NA>": default})

    @staticmethod
    def _read_sheet_fast(wb: openpyxl.Workbook, sheet_name: str) -> pd.DataFrame:
        """Lê uma aba do Excel via streaming read_only parando em blocos vazios para alta performance."""
        ws = wb[sheet_name]
        rows = []
        header = None
        empty_streak = 0

        for idx, r in enumerate(ws.iter_rows(values_only=True)):
            if idx == 0:
                header = r
                continue
            if not r or all(v is None for v in r[:8]):
                empty_streak += 1
                if empty_streak > 30:
                    break
                continue
            empty_streak = 0
            rows.append(r)

        if not header:
            return pd.DataFrame()

        norm_cols = [normalize_str(c) if c is not None else f"col_{i}" for i, c in enumerate(header)]
        df_raw = pd.DataFrame(rows, columns=norm_cols)
        # Remove colunas duplicadas preservando a primeira ocorrência (ex.: 'colaborador', 'natureza profissional')
        df_raw = df_raw.loc[:, ~df_raw.columns.duplicated()]
        return df_raw

    def _standardize_nomina_df(self, df_renamed: pd.DataFrame, sheet_label: str) -> pd.DataFrame:
        """Valida colunas obrigatórias e padroniza o DataFrame extraído da Nomina ou da aba Com Gestor."""
        cols = list(df_renamed.columns)

        def find_col(candidates_fn, nome_amigavel: str, required: bool = True) -> Optional[str]:
            matches = [c for c in cols if candidates_fn(c)]
            if not matches and required:
                raise KeyError(
                    f"Coluna obrigatória '{nome_amigavel}' não encontrada na aba '{sheet_label}' da Nomina. "
                    f"Verifique a extração do relatório no sistema LG."
                )
            return matches[0] if matches else None

        col_mat = find_col(lambda c: c in ("matricula", "matricula_colaborador"), "Matrícula")
        col_nome = find_col(lambda c: c in ("colaborador", "nome do colaborador"), "Colaborador")
        col_nasc = find_col(lambda c: "nascimento" in c and "data" in c, "Data Nascimento")
        col_cpf = find_col(lambda c: c in ("cpf", "cpf/cnpj") or "cpf" in c, "CPF / CPF/CNPJ")
        col_adm = find_col(lambda c: "admissao" in c and "data" in c, "Data Admissão")
        col_idade = find_col(lambda c: c == "idade", "Idade")
        col_cargo = find_col(lambda c: c == "cargo", "Cargo")
        col_estab = find_col(lambda c: c == "estabelecimento", "Estabelecimento")
        col_natureza = find_col(lambda c: c == "natureza profissional", "Natureza Profissional")
        col_sindicato = find_col(lambda c: c == "sindicato", "Sindicato")
        col_salario = find_col(lambda c: "salario atual" in c or c == "salario", "Salário Atual")
        col_cat = find_col(lambda c: c == "categoria", "Categoria")
        col_vinc = find_col(lambda c: c == "vinculo", "Vínculo")
        col_tipo_sal = find_col(lambda c: "tipo de salario" in c, "Tipo de Salário")

        # Colunas opcionais/complementares de situação especial, rescisão e hierarquia
        col_tipo_sit = find_col(lambda c: c in ("tipo de situacao", "tipo situacao"), "Tipo de Situação", required=False)
        col_sit = find_col(lambda c: c == "situacao", "Situação", required=False)
        col_resc = find_col(lambda c: "rescisao" in c and "data" in c, "Data de Rescisão", required=False)
        col_motivo_deslig = find_col(lambda c: "motivo desligamento" in c, "Motivo Desligamento", required=False)
        col_nivel_hier = find_col(lambda c: c == "nivel hierarquico", "Nível Hierárquico", required=False)

        clean_df = pd.DataFrame()
        clean_df["matricula"] = df_renamed[col_mat].apply(format_matricula)
        clean_df["colaborador"] = self._clean_str_series(df_renamed[col_nome]).str.upper()
        clean_df["data_nascimento"] = pd.to_datetime(df_renamed[col_nasc], errors="coerce")
        clean_df["cpf"] = df_renamed[col_cpf].apply(format_cpf)
        clean_df["data_admissao"] = pd.to_datetime(df_renamed[col_adm], errors="coerce")
        clean_df["idade"] = pd.to_numeric(df_renamed[col_idade], errors="coerce").fillna(0).astype(int)
        clean_df["cargo"] = self._clean_str_series(df_renamed[col_cargo])
        clean_df["estabelecimento"] = self._clean_str_series(df_renamed[col_estab])
        clean_df["natureza_profissional"] = self._clean_str_series(df_renamed[col_natureza])
        clean_df["sindicato"] = self._clean_str_series(df_renamed[col_sindicato])
        clean_df["salario_atual"] = df_renamed[col_salario].apply(parse_numeric_currency)
        clean_df["categoria"] = self._clean_str_series(df_renamed[col_cat])
        clean_df["vinculo"] = self._clean_str_series(df_renamed[col_vinc])
        clean_df["tipo_salario"] = self._clean_str_series(df_renamed[col_tipo_sal])

        clean_df["tipo_situacao"] = self._clean_str_series(df_renamed[col_tipo_sit], default="Atividade Normal") if col_tipo_sit else "Atividade Normal"
        clean_df["situacao"] = self._clean_str_series(df_renamed[col_sit], default="Atividade Normal") if col_sit else "Atividade Normal"
        clean_df["data_rescisao"] = pd.to_datetime(df_renamed[col_resc], errors="coerce") if col_resc else pd.NaT
        clean_df["motivo_desligamento"] = self._clean_str_series(df_renamed[col_motivo_deslig]) if col_motivo_deslig else ""
        clean_df["nivel_hierarquico"] = self._clean_str_series(df_renamed[col_nivel_hier]) if col_nivel_hier else ""

        clean_df = clean_df[clean_df["matricula"] > 0].copy()
        return clean_df

    def load_nomina(self) -> pd.DataFrame:
        """
        Carrega a base da Nomina de Ativos (aba 'Nomina') E localiza dinamicamente
        todos os Diretores na aba 'Com Gestor' (onde Natureza Profissional == 'Diretores'
        ou Nível Hierárquico == 'Executivo'), unificando-os com rastreabilidade total.
        """
        mapped = self.raw_map.get("nomina")
        if mapped and Path(mapped).exists():
            file_path = Path(mapped)
        else:
            nomina_files = self._find_files(["*Nomina*.xlsx", "*nomina*.xlsx", "*Nomina*.xls"])
            if not nomina_files:
                raise FileNotFoundError(f"Arquivo da Nomina não encontrado em {self.raw_dir}")
            file_path = nomina_files[0]

        self._log(f"[*] Carregando Nomina: {file_path.name}")

        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        try:
            sheet_target = "Nomina" if "Nomina" in wb.sheetnames else wb.sheetnames[0]
            df_raw_nom = self._read_sheet_fast(wb, sheet_target)
            clean_nom = self._standardize_nomina_df(df_raw_nom, sheet_label=sheet_target)
            clean_nom["is_diretor"] = clean_nom["natureza_profissional"].apply(normalize_str) == "diretores"
            self._log(f"    -> {len(clean_nom)} colaboradores carregados da aba '{sheet_target}'.")

            # Busca dinâmica da aba 'Com Gestor' para extração oficial dos Diretores
            cg_candidates = [s for s in wb.sheetnames if "com gestor" in normalize_str(s) or "gestor" in normalize_str(s)]
            if cg_candidates:
                cg_sheet = cg_candidates[0]
                self._aba_com_gestor_encontrada = True
                df_raw_cg = self._read_sheet_fast(wb, cg_sheet)
                clean_cg = self._standardize_nomina_df(df_raw_cg, sheet_label=cg_sheet)

                mask_diretores = (
                    (clean_cg["natureza_profissional"].apply(normalize_str) == "diretores") |
                    (clean_cg["nivel_hierarquico"].apply(normalize_str) == "executivo")
                )
                df_diretores_cg = clean_cg[mask_diretores].copy()
                df_diretores_cg["is_diretor"] = True
                df_diretores_cg = df_diretores_cg.sort_values(by="colaborador").reset_index(drop=True)

                self._diretores_com_gestor = [
                    {
                        "matricula": int(r.matricula),
                        "nome": str(r.colaborador),
                        "nascimento": r.data_nascimento.strftime("%d/%m/%Y") if pd.notna(r.data_nascimento) else "",
                        "cpf": str(r.cpf),
                        "cargo": str(r.cargo),
                        "bandeira_normal": "GBARBOSA",
                        "plano_normal": "PLANO 01"
                    }
                    for r in df_diretores_cg.itertuples(index=False)
                ]

                self._log(f"    -> {len(df_diretores_cg)} Diretores localizados automaticamente na aba '{cg_sheet}'.")

                # Concatena Diretores no topo da base unificada (evitando duplicatas caso algum conste nas duas abas)
                clean_df = pd.concat([df_diretores_cg, clean_nom], ignore_index=True).drop_duplicates(subset=["matricula"], keep="first")
                return clean_df
            else:
                self._log("    -> [AVISO] Aba 'Com Gestor' não encontrada na Nomina. Será utilizado o cadastro contingencial de diretores.")
                self._aba_com_gestor_encontrada = False
                return clean_nom
        finally:
            wb.close()

    def load_oposicao(self) -> pd.DataFrame:
        """Carrega o Relatório de Oposições ao Desconto (CSV/Excel)."""
        mapped = self.raw_map.get("oposicao")
        if mapped and Path(mapped).exists():
            file_path = Path(mapped)
        else:
            opo_files = self._find_files(["*Oposi*.CSV", "*Oposi*.csv", "*Oposi*.xlsx", "*oposi*.csv", "*oposi*.xlsx"])
            if not opo_files:
                self._log("[!] Aviso: Arquivo de Oposições não encontrado. Base vazia assumida.")
                return pd.DataFrame(columns=["matricula", "oposicao_ativa", "motivo"])
            file_path = opo_files[0]

        self._log(f"[*] Carregando Relatório de Oposições: {file_path.name}")

        if file_path.suffix.lower() == ".csv":
            df = self._read_csv_with_encodings(file_path, sep=";")
        else:
            df = pd.read_excel(file_path, dtype=str)

        norm_cols = {c: normalize_str(c) for c in df.columns}
        df = df.rename(columns=norm_cols)

        mat_candidates = [c for c in df.columns if "matricula" in c]
        if not mat_candidates:
            raise KeyError(f"Coluna 'Matrícula' não encontrada no Relatório de Oposições ({file_path.name}).")
        col_mat = mat_candidates[0]
        col_val_adic = [c for c in df.columns if "valor inf adicional" in c or "valor" in c]

        clean_df = pd.DataFrame()
        clean_df["matricula"] = df[col_mat].apply(format_matricula)

        if col_val_adic:
            val_col = col_val_adic[0]
            clean_df["oposicao_ativa"] = df[val_col].astype(str).str.contains("SIM", case=False, na=False)
        else:
            clean_df["oposicao_ativa"] = True

        clean_df = clean_df[clean_df["oposicao_ativa"]].drop_duplicates(subset=["matricula"])
        self._log(f"    -> {len(clean_df)} oposições ativas encontradas.")
        return clean_df

    def load_fopa(self) -> pd.DataFrame:
        """
        Carrega o Relatório FOPA de Descontos em Folha.
        Valida explicitamente a presença dos eventos de seguro (6266/6267).
        Se outro arquivo for fornecido (ex: odonto), emite aviso transparente de não conformidade.
        """
        mapped = self.raw_map.get("fopa")
        if mapped and Path(mapped).exists():
            file_path = Path(mapped)
        else:
            fopa_files = self._find_files(["*Eventos*.CSV", "*FOPA*.CSV", "*Eventos*.csv", "*FOPA*.csv", "*FOPA*.xlsx"])
            if not fopa_files:
                self._log("[!] Aviso: Arquivo FOPA não encontrado. Descontos em folha não serão confrontados.")
                return pd.DataFrame(columns=["matricula", "desconto_6266", "desconto_6267", "fopa_valida"])
            file_path = fopa_files[0]

        self._log(f"[*] Carregando FOPA: {file_path.name}")

        if file_path.suffix.lower() == ".csv":
            skip_rows = self._detect_header_row(file_path)
            df = self._read_csv_with_encodings(file_path, sep=";", skiprows=skip_rows)
        else:
            df = pd.read_excel(file_path, dtype=str)

        col_mat = None
        col_6266 = None
        col_6267 = None

        for col in df.columns:
            c_norm = normalize_str(col)
            if "matricula" in c_norm and not col_mat:
                col_mat = col
            if EVENTO_SEGURO_NORMAL in str(col):
                col_6266 = col
            if EVENTO_SEGURO_COMP in str(col):
                col_6267 = col

        clean_df = pd.DataFrame()
        if col_mat:
            clean_df["matricula"] = df[col_mat].apply(format_matricula)
        else:
            clean_df["matricula"] = []

        # Validação estrita do cabeçalho da FOPA (Critério 3 da Auditoria)
        if col_6266 or col_6267:
            clean_df["desconto_6266"] = self._clean_currency_col(df[col_6266]) if col_6266 else 0.0
            clean_df["desconto_6267"] = self._clean_currency_col(df[col_6267]) if col_6267 else 0.0
            clean_df["fopa_valida"] = True
            self._log(f"    -> [CONFORME] Eventos de Seguro mapeados com sucesso na FOPA: 6266={bool(col_6266)}, 6267={bool(col_6267)}")
        else:
            clean_df["desconto_6266"] = 0.0
            clean_df["desconto_6267"] = 0.0
            clean_df["fopa_valida"] = False
            detected_events = [c for c in df.columns if any(char.isdigit() for char in str(c))][:3]
            self._log(f"    -> [ALERTA AUDITORIA] O arquivo FOPA fornecido '{file_path.name}' NÃO contém os eventos 6266 e 6267 de Seguro de Vida!")
            self._log(f"       Colunas detectadas: {detected_events}. Operando com batimento histórico e cadastral.")

        return clean_df

    def load_optantes_complementar(self) -> Dict[int, Dict[str, Any]]:
        """Carrega o catálogo de optantes conhecidos do Seguro Complementar (12x e 24x)."""
        if OPTANTES_COMPLEMENTAR_JSON.exists():
            with open(OPTANTES_COMPLEMENTAR_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
            optantes = {int(k): v for k, v in data.items()}
            self._log(f"[*] {len(optantes)} Optantes Complementares carregados do catálogo.")
            return optantes
        return {}

    def load_diretores(self) -> List[Dict[str, Any]]:
        """
        Retorna os Diretores extraídos dinamicamente da aba 'Com Gestor' da Nomina.
        Usa o cadastro JSON estático apenas como contingência caso a aba 'Com Gestor' não exista.
        """
        if self._aba_com_gestor_encontrada:
            return self._diretores_com_gestor
        if CADASTRO_DIRETORES_JSON.exists():
            with open(CADASTRO_DIRETORES_JSON, "r", encoding="utf-8") as f:
                diretores = json.load(f)
            self._log(f"[*] {len(diretores)} Diretores carregados do cadastro contingencial (JSON).")
            return diretores
        return []

    def load_fatura_seguradora(self, df_nomina: pd.DataFrame) -> pd.DataFrame:
        """
        Carrega ou gera o espelho de faturamento da seguradora/corretora (Zurich/Bradesco)
        para permitir a conciliação trilateral (Nomina x FOPA x Fatura Seguradora).
        """
        mapped = self.raw_map.get("fatura")
        if mapped and Path(mapped).exists():
            fatura_files = [Path(mapped)]
        else:
            fatura_files = self._find_files(["*Fatura*.xlsx", "*Seguradora*.xlsx", "*Zurich*.xlsx", "*Fatura*.csv"])

        if fatura_files:
            file_path = fatura_files[0]
            self._log(f"[*] Carregando Fatura da Seguradora: {file_path.name}")
            if file_path.suffix.lower() == ".csv":
                df_fat = self._read_csv_with_encodings(file_path, sep=";")
            else:
                df_fat = pd.read_excel(file_path, dtype=str)
            norm_cols = {c: normalize_str(c) for c in df_fat.columns}
            df_fat = df_fat.rename(columns=norm_cols)

            col_mat = [c for c in df_fat.columns if "matricula" in c][0] if any("matricula" in c for c in df_fat.columns) else None
            col_premio = [c for c in df_fat.columns if "premio" in c or "valor" in c][0] if any("premio" in c or "valor" in c for c in df_fat.columns) else None

            clean_fat = pd.DataFrame()
            if col_mat:
                clean_fat["matricula"] = df_fat[col_mat].apply(format_matricula)
                clean_fat["premio_faturado"] = self._clean_currency_col(df_fat[col_premio]) if col_premio else 5.50
                clean_fat["status_fatura"] = "FATURADO"
                self._log(f"    -> {len(clean_fat)} vidas faturadas pela seguradora.")
                return clean_fat

        # Se a fatura ainda não estiver no diretório para o mês corrente,
        # montamos a base de referência esperada com base na apólice para auditoria preventiva
        self._log("[!] Fatura da Seguradora não encontrada no diretório bruto. Ativando módulo de conciliação preventiva com vidas ativas esperadas.")
        expected_fat = pd.DataFrame({
            "matricula": df_nomina["matricula"],
            "premio_faturado": 5.50,  # Custo médio de referência da apólice
            "status_fatura": "ESPERADO_FATURAMENTO"
        })
        return expected_fat

    def _read_csv_with_encodings(self, path: Path, sep: str = ";", skiprows: int = 0) -> pd.DataFrame:
        """Tenta abrir CSV testando múltiplos encodings de forma resiliente."""
        encodings = ["latin-1", "cp1252", "utf-8-sig", "utf-8"]
        for enc in encodings:
            try:
                df = pd.read_csv(path, sep=sep, skiprows=skiprows, encoding=enc, dtype=str)
                return df
            except Exception:
                continue
        raise ValueError(f"Não foi possível decodificar o arquivo {path.name} com nenhum dos encodings conhecidos.")

    def _detect_header_row(self, path: Path) -> int:
        """Localiza em qual linha se encontra o cabeçalho real do relatório Gen.te/LG."""
        encodings = ["latin-1", "cp1252", "utf-8"]
        for enc in encodings:
            try:
                with open(path, "r", encoding=enc) as f:
                    for idx, line in enumerate(f):
                        if "matricula" in normalize_str(line):
                            return idx
                        if idx > 15:
                            break
            except Exception:
                pass
        return 0

    def _clean_currency_col(self, series: pd.Series) -> pd.Series:
        """Converte coluna de valores monetários para float com suporte a múltiplos formatos."""
        return series.apply(parse_numeric_currency)
