"""
Módulo de Carregamento Ultra-Rápido via Streaming XML para Conferência de Valores Pré-Folha VT.
Carrega a aba principal FOPA_Eventos_por_Colaborador (1.183 colaboradores) e as abas auxiliares
de importação dos eventos 31935 (387 registros) e 6292 (30 registros) em menos de 1 segundo.
"""

import glob
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
if str(SCRIPT_DIR.parent) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR.parent))

try:
    from config import GLOB_CONFERENCIA, INPUT_DIR
except ImportError:
    from scripts.config import GLOB_CONFERENCIA, INPUT_DIR


def parse_excel_date(val: Any) -> Any:
    """Converte serial numérico do Excel ou string em datetime.date nativo."""
    if val is None or pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s in ("#N/A", "None", "nan", "-", "0"):
        return None
    try:
        serial = float(s)
        if 15000 < serial < 80000:
            return (datetime(1899, 12, 30) + timedelta(days=int(serial))).date()
    except (ValueError, TypeError):
        pass
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s[:19], fmt).date()
        except ValueError:
            continue
    return s


def parse_clean_int(val: Any) -> Optional[int]:
    """Converte valor em int nativo limpo (sem .0), retornando None se inválido."""
    if val is None or pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s in ("#N/A", "None", "nan", "-"):
        return None
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return None


def parse_clean_float(val: Any, default: float = 0.0) -> float:
    """Converte valor em float nativo arredondado em 2 casas decimais."""
    if val is None or pd.isna(val):
        return default
    s = str(val).strip()
    if not s or s in ("#N/A", "None", "nan", "-"):
        return default
    try:
        return round(float(s), 2)
    except (ValueError, TypeError):
        return default


def parse_clean_str(val: Any, default: Optional[str] = None) -> Optional[str]:
    """Higieniza strings removendo #N/A e espaços extras."""
    if val is None or pd.isna(val):
        return default
    s = str(val).strip()
    if not s or s in ("#N/A", "None", "nan"):
        return default
    return s


class DataLoader:
    """Carregador otimizado de Valores Pré-Folha VT (Eventos 5750, 31935 e 6292)."""

    def __init__(self, input_dir: Path = INPUT_DIR) -> None:
        self.input_dir = Path(input_dir)

    def _stream_xlsx_sheet(
        self, xlsx_path: Path, target_sheet_name: str, max_empty_streak: int = 50
    ) -> List[Dict[str, Any]]:
        """
        Realiza streaming XML direto do arquivo .xlsx via zipfile + ElementTree.iterparse,
        interrompendo após encontrar linhas vazias consecutivas.
        """
        rows_data: List[Dict[str, Any]] = []
        with zipfile.ZipFile(xlsx_path, "r") as zf:
            wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
            ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            rel_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

            rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            rel_map = {r.attrib["Id"]: "xl/" + r.attrib["Target"] for r in rels_xml}

            sheet_xml_path: Optional[str] = None
            for sh in wb_xml.findall(".//main:sheet", ns):
                if sh.attrib.get("name") == target_sheet_name:
                    sheet_xml_path = rel_map.get(sh.attrib.get(rel_ns, ""))
                    break

            if not sheet_xml_path:
                raise ValueError(f"Aba '{target_sheet_name}' não encontrada em {xlsx_path.name}")

            shared_strings: List[str] = []
            if "xl/sharedStrings.xml" in zf.namelist():
                with zf.open("xl/sharedStrings.xml") as f_ss:
                    for _, elem in ET.iterparse(f_ss, events=("end",)):
                        if elem.tag.endswith("}si"):
                            shared_strings.append("".join(elem.itertext()))
                            elem.clear()

            empty_streak = 0
            with zf.open(sheet_xml_path) as f_sh:
                for _, elem in ET.iterparse(f_sh, events=("end",)):
                    if elem.tag.endswith("}row"):
                        row_dict: Dict[str, Any] = {}
                        for cell in elem:
                            if cell.tag.endswith("}c"):
                                ref = cell.attrib.get("r", "")
                                col_let = "".join(ch for ch in ref if ch.isalpha())
                                c_type = cell.attrib.get("t", "")
                                v_text: Optional[str] = None
                                for child in cell:
                                    if child.tag.endswith("}v"):
                                        v_text = child.text
                                        break
                                if v_text is not None:
                                    if c_type == "s":
                                        idx = int(v_text)
                                        row_dict[col_let] = (
                                            shared_strings[idx]
                                            if idx < len(shared_strings)
                                            else v_text
                                        )
                                    else:
                                        row_dict[col_let] = v_text
                        elem.clear()
                        if row_dict:
                            rows_data.append(row_dict)
                            empty_streak = 0
                        else:
                            empty_streak += 1
                            if empty_streak > max_empty_streak:
                                break
        return rows_data

    def load_fopa_prefolha(self, conf_path: Path) -> pd.DataFrame:
        """Carrega os 1.183 colaboradores da aba 'FOPA_Eventos_por_Colaborador - '."""
        raw_rows = self._stream_xlsx_sheet(conf_path, "FOPA_Eventos_por_Colaborador - ")
        if len(raw_rows) < 2:
            raise RuntimeError("Aba 'FOPA_Eventos_por_Colaborador - ' está vazia.")

        records: List[Dict[str, Any]] = []
        for r in raw_rows[1:]:
            matricula = parse_clean_int(r.get("G"))
            if matricula is None:
                continue
            records.append(
                {
                    "Código da Empresa": parse_clean_int(r.get("A")),
                    "Nome da Empresa": parse_clean_str(r.get("B"), "-"),
                    "Código Estabelecimento": parse_clean_int(r.get("C")),
                    "Estabelecimento": parse_clean_str(r.get("D"), "-"),
                    "Código Centro Custo": parse_clean_str(r.get("E"), "-"),
                    "Centro Custo": parse_clean_str(r.get("F"), "-"),
                    "Matrícula": matricula,
                    "Nome do Colaborador": parse_clean_str(r.get("H"), "-"),
                    "ID Colaborador": parse_clean_int(r.get("I")),
                    "Data Admissão": parse_excel_date(r.get("J")),
                    "Código Cargo": parse_clean_int(r.get("K")),
                    "Cargo": parse_clean_str(r.get("L"), "-"),
                    "Situação": parse_clean_str(r.get("M"), "-"),
                    "Tipo Situação": parse_clean_str(r.get("N"), "-"),
                    "Data Rescisão": parse_excel_date(r.get("O")),
                    "118 - Credito Vale Transporte": parse_clean_float(r.get("P")),
                    "147 - Credito Vale Transporte Dev Operadora FL 40": parse_clean_float(r.get("Q")),
                    "148 - Credito Vale Transporte Dev Operadora FL 41": parse_clean_float(r.get("R")),
                    "159 - Deslocamento Atividades Externa": parse_clean_float(r.get("S")),
                    "490 - Credito Vale Transporte": parse_clean_float(r.get("T")),
                    "5731 - Saldo Vale Transporte Rescisao": parse_clean_float(r.get("U")),
                    "5750 - Adiantamento VT Atraves Alcadas": parse_clean_float(r.get("V")),
                    "6292 - Desconto Cartao VTE": parse_clean_float(r.get("W")),
                    "9290 - Dev Vale Transporte": parse_clean_float(r.get("X")),
                    "9421 - Reembolso VT Sem Dedução - Folha 40": parse_clean_float(r.get("Y")),
                    "9422 - Reembolso VT Sem Dedução - Folha 41": parse_clean_float(r.get("Z")),
                    "9423 - Reembolso VT Sem Dedução - Folha 42": parse_clean_float(r.get("AA")),
                    "9424 - Credito Vale Transporte - Folha 41": parse_clean_float(r.get("AB")),
                    "9425 - Credito Vale Transporte - Folha 42": parse_clean_float(r.get("AC")),
                    "9426 - Dev Vale Transporte - Folha 41": parse_clean_float(r.get("AD")),
                    "9427 - Dev Vale Transporte - Folha 42": parse_clean_float(r.get("AE")),
                    "9428 - Deslocamento Atividades Externa - Folha 41": parse_clean_float(r.get("AF")),
                    "9429 - Deslocamento Atividades Externa - Folha 42": parse_clean_float(r.get("AG")),
                    "31935 - Valor Vale Transporte Alternativo Mes Seguinte": parse_clean_float(r.get("AH")),
                    "Obs_Original": parse_clean_str(r.get("AJ"), "Ok"),
                }
            )
        return pd.DataFrame(records)

    def load_import_31935(self, conf_path: Path) -> pd.DataFrame:
        """Carrega os 387 registros da aba '31935' (Importação VT Alternativo Mês Seguinte)."""
        raw_rows = self._stream_xlsx_sheet(conf_path, "31935")
        records: List[Dict[str, Any]] = []
        for r in raw_rows[1:]:
            mat = parse_clean_int(r.get("B"))
            if mat is None:
                continue
            fopa_raw = r.get("D")
            fopa_val = None if str(fopa_raw).strip() in ("#N/A", "None", "") else parse_clean_float(fopa_raw)
            records.append(
                {
                    "Emp": parse_clean_int(r.get("A")),
                    "Matrícula": mat,
                    "Valor Importado 31935": parse_clean_float(r.get("C")),
                    "Fopa Mensal 31935": fopa_val,
                    "Obs 31935": parse_clean_str(r.get("F"), "Ok"),
                }
            )
        return pd.DataFrame(records)

    def load_import_6292(self, conf_path: Path) -> pd.DataFrame:
        """Carrega os 30 registros da aba '6292' (Importação Desconto Cartão VTE)."""
        raw_rows = self._stream_xlsx_sheet(conf_path, "6292")
        records: List[Dict[str, Any]] = []
        for r in raw_rows[1:]:
            mat = parse_clean_int(r.get("C"))
            if mat is None:
                continue
            fopa_raw = r.get("G")
            fopa_val = None if str(fopa_raw).strip() in ("#N/A", "None", "") else parse_clean_float(fopa_raw)
            records.append(
                {
                    "Emp": parse_clean_int(r.get("A")),
                    "Loja": parse_clean_int(r.get("B")),
                    "Matrícula": mat,
                    "Colaborador": parse_clean_str(r.get("D"), "-"),
                    "Evento": parse_clean_int(r.get("E")) or 6292,
                    "Valor Importado 6292": parse_clean_float(r.get("F")),
                    "Fopa Mensal 6292": fopa_val,
                    "Obs 6292": parse_clean_str(r.get("I"), "Ok"),
                }
            )
        return pd.DataFrame(records)

    def load_all(self) -> Dict[str, pd.DataFrame]:
        """Executa o carregamento consolidado das abas de Valores Pré-Folha."""
        matches = sorted(glob.glob(str(self.input_dir / GLOB_CONFERENCIA)))
        if not matches:
            raise FileNotFoundError(
                f"Arquivo '{GLOB_CONFERENCIA}' não encontrado em {self.input_dir}"
            )
        conf_path = Path(matches[0])
        print(f" [LOADER] Streaming XML de alta performance: {conf_path.name}")

        df_fopa = self.load_fopa_prefolha(conf_path)
        df_31935 = self.load_import_31935(conf_path)
        df_6292 = self.load_import_6292(conf_path)

        print(f" [LOADER] FOPA Pré-Folha carregada: {len(df_fopa):,} colaboradores")
        print(f" [LOADER] Base Importação 31935: {len(df_31935):,} colaboradores")
        print(f" [LOADER] Base Importação 6292: {len(df_6292):,} colaboradores")

        return {
            "df_fopa": df_fopa,
            "df_31935": df_31935,
            "df_6292": df_6292,
        }
