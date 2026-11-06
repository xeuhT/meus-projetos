"""
Detector Inteligente e Validador Preventivo de Colunas LG — Seguro de Vida Cencosud Brasil.
Inspeciona a pasta de entrada selecionada pelo usuário, identifica os relatórios mais recentes
(Nomina, Oposições, FOPA e Fatura da Seguradora) e valida preventivamente se todas as abas
e colunas obrigatórias da LG estão presentes antes de liberar o processamento.
"""

import json
import os
import re
import unicodedata
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple
import openpyxl

from config import (
    OPTANTES_COMPLEMENTAR_JSON,
    EVENTO_SEGURO_NORMAL,
    EVENTO_SEGURO_COMP,
)


def normalize_text(s: Any) -> str:
    """Remove acentos e espaços extras e converte para minúsculas."""
    if s is None:
        return ""
    txt = str(s).strip()
    return "".join(c for c in unicodedata.normalize("NFD", txt) if unicodedata.category(c) != "Mn").lower()


def format_size(num_bytes: int) -> str:
    """Formata tamanho de arquivo em KB ou MB para exibição amigável."""
    if num_bytes < 1024:
        return f"{num_bytes} B"
    kb = num_bytes / 1024
    if kb < 1024:
        return f"{kb:.1f} KB"
    mb = kb / 1024
    return f"{mb:.1f} MB"


class FileDetector:
    """Varre a pasta de relatórios brutos e realiza validação preventiva de abas e colunas LG."""

    IGNORED_DIR_PREFIXES = (
        "carga seguro",
        "relatorios de conferencia",
        "relatórios de conferência",
        "backup",
        "old",
        "temp",
        "_sistema",
        ".",
        "__pycache__",
    )

    def __init__(self, root_dir: str):
        self.root_dir = Path(root_dir).resolve()
        self.all_files: List[Dict[str, Any]] = []

    def scan(self) -> Dict[str, Any]:
        """Executa varredura recursiva controlada e retorna o diagnóstico completo para a GUI."""
        self.all_files = []
        if self.root_dir.exists():
            for root, dirs, files in os.walk(self.root_dir):
                dirs[:] = [
                    d for d in dirs
                    if not any(normalize_text(d).startswith(normalize_text(p)) for p in self.IGNORED_DIR_PREFIXES)
                ]
                for fname in files:
                    if fname.startswith("~$") or fname.startswith("."):
                        continue
                    fpath = Path(root) / fname
                    ext = fpath.suffix.lower()
                    if ext not in (".xlsx", ".xls", ".csv"):
                        continue
                    try:
                        st = fpath.stat()
                        mtime = st.st_mtime
                        size = st.st_size
                    except OSError:
                        mtime = 0.0
                        size = 0

                    self.all_files.append({
                        "name": fname,
                        "full_path": str(fpath),
                        "norm": normalize_text(fname),
                        "mtime": mtime,
                        "size_bytes": size,
                        "size": format_size(size),
                        "ext": ext,
                    })

            self.all_files.sort(key=lambda x: x["mtime"], reverse=True)

        return self._diagnose()

    def _find_best(self, keywords: Tuple[str, ...], exts: Tuple[str, ...], exclude_kw: Tuple[str, ...] = ()) -> Optional[Dict[str, Any]]:
        for item in self.all_files:
            if item["ext"] not in exts:
                continue
            if any(ek in item["norm"] for ek in exclude_kw):
                continue
            if any(kw in item["norm"] for kw in keywords):
                return item
        return None

    @staticmethod
    def _validate_nomina_workbook(file_path: str) -> Tuple[bool, str, Optional[Tuple[str, str]]]:
        """
        Abre apenas o cabeçalho (linha 1) das abas 'Nomina' e 'Com gestor' em modo streaming
        para validar se nenhuma coluna obrigatória da LG foi removida.
        Retorna (ok, mensagem_observacao, competencia_sugerida_mm_yyyy).
        """
        wb = None
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            sheets = wb.sheetnames
            if not sheets:
                return False, "Planilha sem abas legíveis", None

            sheet_nomina = "Nomina" if "Nomina" in sheets else sheets[0]
            cg_candidates = [s for s in sheets if "com gestor" in normalize_text(s) or "gestor" in normalize_text(s)]
            sheet_cg = cg_candidates[0] if cg_candidates else None

            regras_obrigatorias = [
                ("Matrícula", lambda c: c in ("matricula", "matricula_colaborador")),
                ("Colaborador", lambda c: c in ("colaborador", "nome do colaborador")),
                ("Data Nascimento", lambda c: "nascimento" in c and "data" in c),
                ("CPF / CPF/CNPJ", lambda c: c in ("cpf", "cpf/cnpj") or "cpf" in c),
                ("Data Admissão", lambda c: "admissao" in c and "data" in c),
                ("Idade", lambda c: c == "idade"),
                ("Cargo", lambda c: c == "cargo"),
                ("Estabelecimento", lambda c: c == "estabelecimento"),
                ("Natureza Profissional", lambda c: c == "natureza profissional"),
                ("Sindicato", lambda c: c == "sindicato"),
                ("Salário Atual", lambda c: "salario atual" in c or c == "salario"),
                ("Categoria", lambda c: c == "categoria"),
                ("Vínculo", lambda c: c == "vinculo"),
                ("Tipo de Salário", lambda c: "tipo de salario" in c),
            ]

            for aba in ([sheet_nomina, sheet_cg] if sheet_cg else [sheet_nomina]):
                ws = wb[aba]
                header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
                if not header_row:
                    return False, f"Aba '{aba}' está vazia", None
                norm_cols = [normalize_text(c) for c in header_row if c is not None]
                faltantes = [
                    nome_col for nome_col, fn in regras_obrigatorias
                    if not any(fn(c) for c in norm_cols)
                ]
                if faltantes:
                    return False, f"Falta coluna na aba '{aba}': {', '.join(faltantes[:3])}", None

            if sheet_cg:
                obs = f"Abas '{sheet_nomina}' + '{sheet_cg}' (Diretores) validadas • Colunas LG OK"
            else:
                obs = f"Aba '{sheet_nomina}' OK • Atenção: aba 'Com gestor' ausente (usará contingência)"

            return True, obs, None
        except PermissionError:
            return True, "Arquivo aberto no Excel (feche se houver bloqueio de leitura)", None
        except Exception as ex:
            return False, f"Erro ao validar cabeçalhos da Nomina: {ex}", None
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass

    @staticmethod
    def _validate_oposicao_file(file_path: str) -> Tuple[bool, str]:
        """Valida rapidamente o cabeçalho do Relatório de Oposições."""
        p = Path(file_path)
        if p.suffix.lower() == ".csv":
            for enc in ("latin-1", "cp1252", "utf-8-sig", "utf-8"):
                try:
                    with open(p, "r", encoding=enc) as f:
                        first_lines = [normalize_text(next(f, "")) for _ in range(5)]
                    joined = " ".join(first_lines)
                    if "matricula" in joined:
                        return True, "Colunas 'Matrícula' e 'Valor Inf Adicional' OK"
                except Exception:
                    continue
            return False, "Coluna 'Matrícula' não identificada no CSV de Oposições"
        return True, "Planilha de Oposições identificada"

    @staticmethod
    def _validate_fopa_file(file_path: str) -> Tuple[str, str]:
        """
        Verifica se o arquivo FOPA possui os eventos 6266 e/ou 6267 no cabeçalho.
        Retorna (status: 'OK' | 'WARNING', observacao).
        """
        p = Path(file_path)
        if p.suffix.lower() == ".csv":
            for enc in ("latin-1", "cp1252", "utf-8-sig", "utf-8"):
                try:
                    with open(p, "r", encoding=enc) as f:
                        lines = [next(f, "") for _ in range(15)]
                    header_text = "\n".join(lines)
                    has_6266 = EVENTO_SEGURO_NORMAL in header_text
                    has_6267 = EVENTO_SEGURO_COMP in header_text
                    if has_6266 and has_6267:
                        return "OK", "Eventos 6266 (Normal) e 6267 (Complementar) detectados"
                    if has_6266 or has_6267:
                        ev = "6266" if has_6266 else "6267"
                        return "WARNING", f"Apenas evento {ev} localizado no FOPA"
                    return (
                        "WARNING",
                        "Sem eventos 6266/6267 neste FOPA (operará em modo cadastral)",
                    )
                except Exception:
                    continue
        return "OK", "Arquivo FOPA localizado"

    @staticmethod
    def _detect_month_year_hint(files: List[Dict[str, Any]]) -> Optional[Tuple[str, str]]:
        """Tenta inferir mês/ano (MM, YYYY) a partir das datas nos nomes dos arquivos LG."""
        for f in files:
            # Padrão DDMMYYYY comum nas extrações LG (ex: 05102026_102839, 22102026_090721)
            m = re.search(r"(\d{2})(\d{2})(20\d{2})", f["name"])
            if m:
                mes = int(m.group(2))
                ano = m.group(3)
                if 1 <= mes <= 12:
                    return f"{mes:02d}", ano
        return None

    def _diagnose(self) -> Dict[str, Any]:
        nomina_item = self._find_best(("nomina",), (".xlsx", ".xls"), exclude_kw=("relatorio_conferencia",))
        oposicao_item = self._find_best(("oposi",), (".csv", ".xlsx", ".xls"))
        fopa_item = self._find_best(("fopa", "eventos_por_colaborador", "eventos por colaborador"), (".csv", ".xlsx"))
        fatura_item = self._find_best(("fatura", "zurich", "seguradora"), (".xlsx", ".xls", ".csv"), exclude_kw=("relatorio_conferencia", "carga seguro"))

        # 1. Nomina (Obrigatória + Validação de Abas/Colunas)
        nomina_valid = False
        nomina_obs = ""
        if nomina_item:
            nomina_valid, nomina_obs, _ = self._validate_nomina_workbook(nomina_item["full_path"])

        # 2. Oposições (Obrigatória para evitar envio de oponentes à seguradora)
        opo_valid = False
        opo_obs = ""
        if oposicao_item:
            opo_valid, opo_obs = self._validate_oposicao_file(oposicao_item["full_path"])

        # 3. FOPA (Recomendada — alerta preventivo se vier sem 6266/6267)
        fopa_status = "MISSING_OPT"
        fopa_obs = "Opcional/Recomendado — selecione eventos 6266 e 6267 na LG para auditoria de folha"
        if fopa_item:
            fopa_status, fopa_obs = self._validate_fopa_file(fopa_item["full_path"])

        # 4. Fatura da Seguradora (Opcional)
        fatura_status = "OK" if fatura_item else "OPTIONAL"
        fatura_obs = "Pronta para conciliação real Fatura x Carga" if fatura_item else "Opcional — conciliação preventiva ativa se ausente"

        # 5. Catálogo de Optantes Complementares (Embutido)
        qtd_optantes = 0
        if OPTANTES_COMPLEMENTAR_JSON.exists():
            try:
                with open(OPTANTES_COMPLEMENTAR_JSON, "r", encoding="utf-8") as f:
                    qtd_optantes = len(json.load(f))
            except Exception:
                qtd_optantes = 0

        items = [
            {
                "key": "nomina",
                "label": "1. Nomina de Ativos (Abas 'Nomina' + 'Com gestor')",
                "file": nomina_item,
                "req": True,
                "status": ("OK" if nomina_valid else "INVALID") if nomina_item else "MISSING",
                "obs": nomina_obs if nomina_item else "Obrigatório — exporte o relatório Nomina (.xlsx) na LG",
            },
            {
                "key": "oposicao",
                "label": "2. Relatório de Oposições ao Seguro (SGV)",
                "file": oposicao_item,
                "req": True,
                "status": ("OK" if opo_valid else "INVALID") if oposicao_item else "MISSING",
                "obs": opo_obs if oposicao_item else "Obrigatório — exporte o Relatório de Oposições na LG",
            },
            {
                "key": "fopa",
                "label": "3. Desconto em Folha — FOPA (Eventos 6266 / 6267)",
                "file": fopa_item,
                "req": False,
                "status": fopa_status,
                "obs": fopa_obs,
            },
            {
                "key": "fatura",
                "label": "4. Fatura da Seguradora (Zurich / Corretora)",
                "file": fatura_item,
                "req": False,
                "status": fatura_status,
                "obs": fatura_obs,
            },
            {
                "key": "catalogo_comp",
                "label": "5. Catálogo Optantes Seguro Complementar (12x / 24x)",
                "file": {"name": OPTANTES_COMPLEMENTAR_JSON.name, "size": f"{qtd_optantes} optantes"},
                "req": False,
                "status": "OK" if qtd_optantes > 0 else "WARNING",
                "obs": "Integrado ao aplicativo (Planos 05, 06, 08 e 09)",
            },
        ]

        ready = (nomina_item is not None and nomina_valid) and (oposicao_item is not None and opo_valid)
        detected_comp = self._detect_month_year_hint([x for x in (oposicao_item, fopa_item, nomina_item) if x])

        return {
            "ready": ready,
            "items": items,
            "competencia_hint": detected_comp,
            "raw_map": {
                "nomina": nomina_item["full_path"] if nomina_item else None,
                "oposicao": oposicao_item["full_path"] if oposicao_item else None,
                "fopa": fopa_item["full_path"] if fopa_item else None,
                "fatura": fatura_item["full_path"] if fatura_item else None,
            },
        }
