"""
Detector Inteligente de Arquivos de Entrada - Central de Benefícios Cencosud Brasil.

Realiza varredura resiliente da pasta selecionada pelo usuário (suportando tanto a pasta
direta de relatórios quanto a pasta raiz do módulo) e monta os cartões de diagnóstico
visual (🟢 OK, 🟡 Opcional, 🔴 Pendente) para a interface gráfica.
"""

import fnmatch
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional


def resource_path(relative_path: str) -> str:
    """Obtém o caminho absoluto para recursos internos (dev ou PyInstaller --onefile)."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


class FileDetector:
    """Varre e valida os arquivos de entrada para os módulos da Central de Benefícios."""

    def __init__(self, root_dir: str, output_dir: str):
        self.root_dir = Path(root_dir) if root_dir else Path.cwd()
        self.output_dir = Path(output_dir) if output_dir else Path.cwd()

    def _resolve_candidate_dirs(self, subfolder_names: List[str]) -> List[Path]:
        """Retorna lista ordenada de diretórios candidatos para busca de relatórios."""
        candidates: List[Path] = []
        base = self.root_dir

        # Base direta
        for p in [base, base.parent]:
            if p.exists() and p.is_dir() and p not in candidates:
                candidates.append(p)

        # Subpastas relativas
        for sub in subfolder_names:
            for parent_dir in [base, base.parent, base.parent.parent if base.parent else base]:
                target = parent_dir / sub
                if target.exists() and target.is_dir() and target not in candidates:
                    candidates.append(target)
                target_sub = parent_dir / sub / "Relatórios Originais"
                if target_sub.exists() and target_sub.is_dir() and target_sub not in candidates:
                    candidates.append(target_sub)
                target_orig = parent_dir / sub / "Relatorio Original"
                if target_orig.exists() and target_orig.is_dir() and target_orig not in candidates:
                    candidates.append(target_orig)

        return candidates

    def _find_first_match(
        self,
        dirs: List[Path],
        patterns: List[str],
        exclude_substrings: Optional[List[str]] = None,
    ) -> Optional[Path]:
        """Busca o primeiro arquivo que corresponda aos padrões informados nos diretórios candidatos."""
        exclude_substrings = [s.lower() for s in (exclude_substrings or [])]
        for d in dirs:
            try:
                all_files = [f for f in d.iterdir() if f.is_file() and not f.name.startswith("~$")]
            except Exception:
                continue

            for pat in patterns:
                pat_low = pat.lower()
                for f in all_files:
                    fname_low = f.name.lower()
                    if any(ex in fname_low for ex in exclude_substrings):
                        continue
                    if fnmatch.fnmatch(fname_low, pat_low):
                        return f
        return None

    def _fmt_info(self, p: Optional[Path]) -> Optional[Dict[str, str]]:
        if p is None or not p.exists():
            return None
        size_kb = p.stat().st_size / 1024.0
        size_str = f"{size_kb:,.0f} KB" if size_kb < 1024 else f"{size_kb/1024.0:.2f} MB"
        return {
            "name": p.name,
            "path": str(p.resolve()),
            "size": size_str,
        }

    def _infer_competencia_from_fopa(self, fopa_path: Optional[Path]) -> tuple[int, int]:
        """Infere mês e ano da competência a partir da linha de Competência do FOPA."""
        meses_abrev = {
            "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
            "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
        }
        if fopa_path and fopa_path.exists():
            try:
                with open(fopa_path, "r", encoding="latin1", errors="ignore") as f:
                    for _ in range(5):
                        line = f.readline()
                        if "compet" in line.lower():
                            m_text = re.search(r"\b([a-zA-Z]{3})/(\d{2,4})\b", line)
                            if m_text:
                                mes_str = m_text.group(1).lower()
                                ano_str = m_text.group(2)
                                mes = meses_abrev.get(mes_str, 9)
                                ano = int(ano_str)
                                if ano < 100:
                                    ano += 2000
                                return mes, ano

                            m_num = re.search(r"\b(0?[1-9]|1[0-2])/(\d{2,4})\b", line)
                            if m_num:
                                mes = int(m_num.group(1))
                                ano = int(m_num.group(2))
                                if ano < 100:
                                    ano += 2000
                                return mes, ano
            except Exception:
                pass
        return 9, 2026

    # =========================================================================
    # 1. SCAN GOOD CARD (01)
    # =========================================================================
    def scan_good_card(self) -> Dict[str, Any]:
        """Diagnóstico de entradas do módulo 01 - Good Card."""
        search_dirs = self._resolve_candidate_dirs(["04 - GOOD CARD", "Relatorio Original", "Relatórios Originais"])

        fopa_file = self._find_first_match(
            search_dirs,
            ["fopa_eventos_por_colaborador.csv", "*fopa*eventos*.csv"],
        )
        nomina_file = self._find_first_match(
            search_dirs,
            ["cola_colaboradores.csv", "*colaboradores*.csv"],
            exclude_substrings=["afastados", "ferias", "férias", "demitidos"],
        )
        afastados_file = self._find_first_match(
            search_dirs,
            ["cola_colaboradores_afastados.csv", "*afastados*.csv"],
        )
        demitidos_file = self._find_first_match(
            search_dirs,
            ["cola_demitidos.csv", "*demitidos*.csv"],
        )
        ferias_file = self._find_first_match(
            search_dirs,
            ["cola_ferias_colaboradores.csv", "*ferias*.csv", "*férias*.csv"],
        )
        vi_file = self._find_first_match(
            search_dirs,
            [
                "conferência good card_base.xlsx",
                "conferencia good card_base.xlsx",
                "*valores*importados*.xlsx",
                "*valores*importados*.csv",
                "*good card_base*.xlsx",
            ],
            exclude_substrings=["gerado"],
        )

        if vi_file is None:
            bundled = Path(resource_path(os.path.join("assets", "Conferência Good Card_Base.xlsx")))
            if bundled.exists():
                vi_file = bundled

        mes, ano = self._infer_competencia_from_fopa(fopa_file)

        items = [
            {
                "key": "fopa",
                "label": "1. Folha de Pagamento (FOPA Eventos por Colaborador)",
                "file": self._fmt_info(fopa_file),
                "req": True,
                "obs": f"Competência detectada: {mes:02d}/{ano}" if fopa_file else "",
            },
            {
                "key": "nomina",
                "label": "2. Cadastro Geral Ativo — Nomina (COLA_Colaboradores)",
                "file": self._fmt_info(nomina_file),
                "req": True,
                "obs": "Vínculo cadastral e Situação em folha",
            },
            {
                "key": "afastados",
                "label": "3. Colaboradores Afastados (COLA_Colaboradores_Afastados)",
                "file": self._fmt_info(afastados_file),
                "req": True,
                "obs": "Filtro automático de afastamentos vigentes",
            },
            {
                "key": "demitidos",
                "label": "4. Colaboradores Desligados (COLA_Demitidos)",
                "file": self._fmt_info(demitidos_file),
                "req": True,
                "obs": "Exclui 'Transferência sem ônus' automaticamente",
            },
            {
                "key": "ferias",
                "label": "5. Férias de Colaboradores (COLA_Ferias_Colaboradores)",
                "file": self._fmt_info(ferias_file),
                "req": True,
                "obs": f"Períodos de gozo em {mes:02d}/{ano}",
            },
            {
                "key": "valores_importados",
                "label": "6. Base Valores Importados — Evt 29613 (Good Card Base)",
                "file": self._fmt_info(vi_file),
                "req": True,
                "obs": "Conciliação do Lote Importado x Folha (29613)",
            },
        ]

        ready = all(item["file"] is not None for item in items if item["req"])
        effective_data_dir = str(fopa_file.parent) if fopa_file else str(self.root_dir)

        raw_map = {
            "data_dir": effective_data_dir,
            "fopa": str(fopa_file) if fopa_file else None,
            "nomina": str(nomina_file) if nomina_file else None,
            "afastados": str(afastados_file) if afastados_file else None,
            "demitidos": str(demitidos_file) if demitidos_file else None,
            "ferias": str(ferias_file) if ferias_file else None,
            "valores_importados": str(vi_file) if vi_file else None,
            "month": mes,
            "year": ano,
        }

        return {
            "ready": ready,
            "items": items,
            "raw_map": raw_map,
            "month": mes,
            "year": ano,
        }

    # =========================================================================
    # 2. SCAN COMPRA MENSAL VT (03)
    # =========================================================================
    def scan_compra_mensal_vt(self) -> Dict[str, Any]:
        """Diagnóstico de entradas do módulo Compra Mensal VT."""
        search_dirs = self._resolve_candidate_dirs(["Compra Mensal VT", "Relatórios Originais"])

        compra_reg = self._find_first_match(search_dirs, ["relaod*.csv", "*compra*regular*.csv", "*compra*.csv"], exclude_substrings=["complementar"])
        compra_comp = self._find_first_match(search_dirs, ["*compra*complementar*.xlsx", "*complementar*.xlsx"])
        opt_reg = self._find_first_match(search_dirs, ["vltr_p*.csv", "*optantes*.csv", "*vltr*.csv"], exclude_substrings=["admitidos"])
        opt_adm = self._find_first_match(search_dirs, ["*optantes*admitidos*.xlsx", "*admitidos*.xlsx"])
        censo_csv = self._find_first_match(search_dirs, ["cola_c*3*.csv", "*censo*.csv", "*colaboradores*.csv"], exclude_substrings=["afastados", "cessados", "ferias", "demitidos"])
        nomina_xlsx = self._find_first_match(search_dirs, ["nomina*.xlsx", "*nomina*.xlsx"])
        afast_ativos = self._find_first_match(search_dirs, ["cola_c*1*.csv", "*afastados*ativos*.csv", "*afastados*.csv"], exclude_substrings=["cessados"])
        afast_cessados = self._find_first_match(search_dirs, ["cola_c*2*.csv", "*afastados*cessados*.csv"])
        ferias_csv = self._find_first_match(search_dirs, ["cola_f*.csv", "*ferias*.csv", "*férias*.csv"])
        demit_csv = self._find_first_match(search_dirs, ["cola_d*.csv", "*demitidos*.csv", "*desligados*.csv"])

        effective_dir = compra_reg.parent if compra_reg else (censo_csv.parent if censo_csv else self.root_dir)

        items = [
            {
                "key": "compra_reg",
                "label": "1. Relatório Compra Regular VT (RELAOD / CSV)",
                "file": self._fmt_info(compra_reg),
                "req": True,
                "obs": "Volume e valores de recarga regular das operadoras",
            },
            {
                "key": "compra_comp",
                "label": "2. Compra Complementar (Excel)",
                "file": self._fmt_info(compra_comp),
                "req": False,
                "obs": "Cargas complementares avulsas do período",
            },
            {
                "key": "opt_reg",
                "label": "3. Base de Optantes VT (VLTR_P / CSV)",
                "file": self._fmt_info(opt_reg),
                "req": True,
                "obs": "Linhas cadastradas, tarifas e fornecedores ativos",
            },
            {
                "key": "opt_adm",
                "label": "4. Optantes Admitidos no Mês (Excel)",
                "file": self._fmt_info(opt_adm),
                "req": False,
                "obs": "Novos colaboradores admitidos no período de compra",
            },
            {
                "key": "censo_csv",
                "label": "5. Censo Funcional Ativo (COLA_C3 / CSV)",
                "file": self._fmt_info(censo_csv),
                "req": True,
                "obs": "Vínculo empregatício e estabelecimento",
            },
            {
                "key": "nomina_xlsx",
                "label": "6. Nomina de Colaboradores (Nomina.xlsx)",
                "file": self._fmt_info(nomina_xlsx),
                "req": False,
                "obs": "Complemento cadastral institucional",
            },
            {
                "key": "afast_ativos",
                "label": "7. Afastados Ativos (COLA_C1 / CSV)",
                "file": self._fmt_info(afast_ativos),
                "req": True,
                "obs": "Identificação de colaboradores sem direito a recarga",
            },
            {
                "key": "afast_cessados",
                "label": "8. Afastados Cessados / Retornos (COLA_C2 / CSV)",
                "file": self._fmt_info(afast_cessados),
                "req": False,
                "obs": "Retornos de afastamento no mês",
            },
            {
                "key": "ferias_csv",
                "label": "9. Programação de Férias (COLA_F / CSV)",
                "file": self._fmt_info(ferias_csv),
                "req": True,
                "obs": "Dedução proporcional dos dias de férias em gozo",
            },
            {
                "key": "demit_csv",
                "label": "10. Desligados / Rescindidos (COLA_D / CSV)",
                "file": self._fmt_info(demit_csv),
                "req": True,
                "obs": "Prevenção de compra indevida para demitidos",
            },
        ]

        ready = all(item["file"] is not None for item in items if item["req"])

        raw_map = {
            "data_dir": str(effective_dir),
            "compra_reg": str(compra_reg) if compra_reg else None,
            "compra_comp": str(compra_comp) if compra_comp else None,
            "opt_reg": str(opt_reg) if opt_reg else None,
            "opt_adm": str(opt_adm) if opt_adm else None,
            "censo_csv": str(censo_csv) if censo_csv else None,
            "nomina_xlsx": str(nomina_xlsx) if nomina_xlsx else None,
            "afast_ativos": str(afast_ativos) if afast_ativos else None,
            "afast_cessados": str(afast_cessados) if afast_cessados else None,
            "ferias_csv": str(ferias_csv) if ferias_csv else None,
            "demit_csv": str(demit_csv) if demit_csv else None,
        }

        return {
            "ready": ready,
            "items": items,
            "raw_map": raw_map,
            "month": 10,
            "year": 2026,
        }

    # =========================================================================
    # 3. SCAN VALORES PRÉ-FOLHA (06)
    # =========================================================================
    def scan_valores_pre_folha(self) -> Dict[str, Any]:
        """Diagnóstico de entradas do módulo Valores Pré-Folha."""
        search_dirs = self._resolve_candidate_dirs(["Valores Pré-Folha", "Relatórios Originais"])

        base_xlsx = self._find_first_match(
            search_dirs,
            ["conferência 09.2026.xlsx", "conferencia 09.2026.xlsx", "*onfer*.xlsx"],
            exclude_substrings=["auditada", "gerado", "tratado"],
        )
        fopa_csv = self._find_first_match(
            search_dirs,
            ["fopa_eventos_por_colaborador*.csv", "*fopa*.csv"],
        )

        effective_dir = base_xlsx.parent if base_xlsx else self.root_dir

        items = [
            {
                "key": "conferencia_base",
                "label": "1. Planilha Base Pré-Folha (FOPA + Lotes 31935 e 6292)",
                "file": self._fmt_info(base_xlsx),
                "req": True,
                "obs": "Contém as abas oficiais: FOPA_Eventos_por_Colaborador, 31935 e 6292",
            },
            {
                "key": "fopa_csv",
                "label": "2. Extração FOPA Complementar (CSV)",
                "file": self._fmt_info(fopa_csv),
                "req": False,
                "obs": "Arquivo bruto emitido pelo gerador de relatórios",
            },
        ]

        ready = base_xlsx is not None

        raw_map = {
            "data_dir": str(effective_dir),
            "conferencia_base": str(base_xlsx) if base_xlsx else None,
            "fopa_csv": str(fopa_csv) if fopa_csv else None,
        }

        return {
            "ready": ready,
            "items": items,
            "raw_map": raw_map,
            "month": 9,
            "year": 2026,
        }

    # =========================================================================
    # 4. SCAN BLOQUEIO GOOD CARD (10)
    # =========================================================================
    def scan_bloqueio_good_card(self) -> Dict[str, Any]:
        """Diagnóstico de entradas do módulo Bloqueio Good Card."""
        search_dirs = self._resolve_candidate_dirs(["Conferência Bloqueio Good Card", "Relatórios Originais"])

        bloqueio_xlsx = self._find_first_match(
            search_dirs,
            ["bloqueio*afatamento*cencosud*.xlsx", "bloqueio*cencosud*atacado*perini*.xlsx", "*bloqueio*.xlsx"],
            exclude_substrings=["gerado", "tratado"],
        )
        afast_csv = self._find_first_match(
            search_dirs,
            ["cola_colaboradores_afastados*.csv", "*afastados*.csv"],
        )

        effective_dir = bloqueio_xlsx.parent if bloqueio_xlsx else self.root_dir

        items = [
            {
                "key": "bloqueio_base",
                "label": "1. Planilha Base de Bloqueio Ticket (Bandeiras + Geral + Afastados)",
                "file": self._fmt_info(bloqueio_xlsx),
                "req": True,
                "obs": "Abas Cencosud, Atacado, Perini, Geral e Afastados",
            },
            {
                "key": "afast_csv",
                "label": "2. Base de Afastados Complementar (CSV)",
                "file": self._fmt_info(afast_csv),
                "req": False,
                "obs": "Base de afastamentos emitida pelo gerador de relatórios",
            },
        ]

        ready = bloqueio_xlsx is not None

        raw_map = {
            "data_dir": str(effective_dir),
            "bloqueio_base": str(bloqueio_xlsx) if bloqueio_xlsx else None,
            "afast_csv": str(afast_csv) if afast_csv else None,
        }

        return {
            "ready": ready,
            "items": items,
            "raw_map": raw_map,
            "month": 9,
            "year": 2026,
        }

    # =========================================================================
    # 5. SCAN UNIDADE DE ENTREGA (12)
    # =========================================================================
    def scan_unidade_entrega(self) -> Dict[str, Any]:
        search_dirs = self._resolve_candidate_dirs(["Unidade de Entrega", "Relatórios Originais"])
        relaod_csv = self._find_first_match(search_dirs, ["*relaod*.csv", "*compra*.csv"])
        optantes_xlsx = self._find_first_match(search_dirs, ["*optante*.xlsx"])
        effective_dir = relaod_csv.parent if relaod_csv else self.root_dir

        items = [
            {
                "key": "compra_csv",
                "label": "1. Relação de Compra de Vale Transporte com Saldo (RELAOD)",
                "file": self._fmt_info(relaod_csv),
                "req": True,
                "obs": "Relatório emitido pela operadora ou sistema de compra de VT",
            },
            {
                "key": "optantes_xlsx",
                "label": "2. Parametrização de Unidades de Entrega (Optantes.xlsx)",
                "file": self._fmt_info(optantes_xlsx),
                "req": True,
                "obs": "Cadastro de colaboradores com unidade de entrega personalizada",
            },
        ]
        ready = (relaod_csv is not None) and (optantes_xlsx is not None)
        raw_map = {
            "data_dir": str(effective_dir),
            "compra_csv": str(relaod_csv) if relaod_csv else None,
            "optantes_xlsx": str(optantes_xlsx) if optantes_xlsx else None,
        }
        return {"ready": ready, "items": items, "raw_map": raw_map, "month": 10, "year": 2026}

    # =========================================================================
    # 6. SCAN CENCOPAY (13)
    # =========================================================================
    def scan_cencopay(self) -> Dict[str, Any]:
        search_dirs = self._resolve_candidate_dirs(["Cencopay", "Relatório Original", "Relatórios Originais"])
        base_xlsx = self._find_first_match(search_dirs, ["*cencopay*.xlsx"], exclude_substrings=["auditada", "gerado", "tratado"])
        fopa_csv = self._find_first_match(search_dirs, ["fopa_eventos_por_colaborador*.csv", "*fopa*.csv"])
        afast_csv = self._find_first_match(search_dirs, ["cola_colaboradores_afastados*.csv", "*afastados*.csv"])
        ferias_csv = self._find_first_match(search_dirs, ["cola_ferias_colaboradores*.csv", "*ferias*.csv"])
        dem_csv = self._find_first_match(search_dirs, ["cola_demitidos*.csv", "*demitidos*.csv"])

        effective_dir = base_xlsx.parent if base_xlsx else self.root_dir

        items = [
            {
                "key": "cencopay_base",
                "label": "1. Planilha Base Cencopay (Conferência + Valores Importados)",
                "file": self._fmt_info(base_xlsx),
                "req": True,
                "obs": "Base de saldos e deduções Cencopay",
            },
            {
                "key": "fopa_csv",
                "label": "2. Extração FOPA (Eventos 10990, 5081, 14001..14004)",
                "file": self._fmt_info(fopa_csv),
                "req": True,
                "obs": "Folha de pagamento com eventos Cencopay",
            },
            {
                "key": "afast_csv",
                "label": "3. Base de Afastados Complementar",
                "file": self._fmt_info(afast_csv),
                "req": False,
                "obs": "Verificação de líquido zerado por afastamento",
            },
            {
                "key": "ferias_csv",
                "label": "4. Programação de Férias Complementar",
                "file": self._fmt_info(ferias_csv),
                "req": False,
                "obs": "Verificação de líquido zerado por férias",
            },
            {
                "key": "dem_csv",
                "label": "5. Relação de Demitidos Complementar",
                "file": self._fmt_info(dem_csv),
                "req": False,
                "obs": "Identificação de desligados com saldo",
            },
        ]
        ready = (base_xlsx is not None) and (fopa_csv is not None)
        raw_map = {
            "data_dir": str(effective_dir),
            "cencopay_base": str(base_xlsx) if base_xlsx else None,
            "fopa_csv": str(fopa_csv) if fopa_csv else None,
            "afast_csv": str(afast_csv) if afast_csv else None,
            "ferias_csv": str(ferias_csv) if ferias_csv else None,
            "dem_csv": str(dem_csv) if dem_csv else None,
        }
        return {"ready": ready, "items": items, "raw_map": raw_map, "month": 9, "year": 2026}

    # =========================================================================
    # 7. SCAN ARQUIVO DE RETORNO LG (14)
    # =========================================================================
    def scan_retorno_lg(self) -> Dict[str, Any]:
        search_dirs = self._resolve_candidate_dirs(["Arquivo de Retorno", "Relatórios Originais"])
        gb_ods = self._find_first_match(search_dirs, ["*gbarbosa*.ods", "*gbarbosa*.xlsx"])
        mr_ods = self._find_first_match(search_dirs, ["*mercantil*.ods", "*mercantil*.xlsx"])
        compra_tratada = self._find_first_match(search_dirs, ["*compra*tratado*.xlsx", "*relação de compra*.xlsx"])

        effective_dir = compra_tratada.parent if compra_tratada else self.root_dir

        items = [
            {
                "key": "gb_ods",
                "label": "1. Arquivo de Retorno Alelo — Gbarbosa (.ods / .xlsx)",
                "file": self._fmt_info(gb_ods),
                "req": True,
                "obs": "Retorno de recarga Gbarbosa com gestão de saldo",
            },
            {
                "key": "mr_ods",
                "label": "2. Arquivo de Retorno Alelo — Mercantil (.ods / .xlsx)",
                "file": self._fmt_info(mr_ods),
                "req": True,
                "obs": "Retorno de recarga Mercantil com gestão de saldo",
            },
            {
                "key": "compra_tratada",
                "label": "3. Relação de Compra de Vale Transporte Tratado (Excel)",
                "file": self._fmt_info(compra_tratada),
                "req": True,
                "obs": "Base de conciliação para geração dos layouts LG",
            },
        ]
        ready = (gb_ods is not None or mr_ods is not None) and (compra_tratada is not None)
        raw_map = {
            "data_dir": str(effective_dir),
            "gb_ods": str(gb_ods) if gb_ods else None,
            "mr_ods": str(mr_ods) if mr_ods else None,
            "compra_tratada": str(compra_tratada) if compra_tratada else None,
        }
        return {"ready": ready, "items": items, "raw_map": raw_map, "month": 10, "year": 2026}

    # =========================================================================
    # 8. SCAN COMPRA MENSAL TICKET BRETAS (15)
    # =========================================================================
    def scan_compra_ticket_bretas(self) -> Dict[str, Any]:
        search_dirs = self._resolve_candidate_dirs(["Compra Mensal Ticket", "Relatórios Originais"])
        base_xlsx = self._find_first_match(search_dirs, ["*bretas*.xlsx", "*ticket*.xlsx"], exclude_substrings=["auditada", "gerado", "tratado"])
        effective_dir = base_xlsx.parent if base_xlsx else self.root_dir

        items = [
            {
                "key": "ticket_base",
                "label": "1. Base de Conferência Ticket Bretas (Excel)",
                "file": self._fmt_info(base_xlsx),
                "req": True,
                "obs": "Abas FOPA, Conferencia, Compra Manual e Parametrizações",
            },
        ]
        ready = base_xlsx is not None
        raw_map = {
            "data_dir": str(effective_dir),
            "ticket_base": str(base_xlsx) if base_xlsx else None,
        }
        return {"ready": ready, "items": items, "raw_map": raw_map, "month": 9, "year": 2026}

    # =========================================================================
    # 9. SCAN CESTA PRÊMIO BRETAS (16)
    # =========================================================================
    def scan_cesta_premio_bretas(self) -> Dict[str, Any]:
        search_dirs = self._resolve_candidate_dirs(["Compra Mensal Cesta Prêmio", "Relatórios Originais"])
        base_xlsx = self._find_first_match(search_dirs, ["*premio*.xlsx", "*cesta*.xlsx"], exclude_substrings=["auditada", "gerado", "tratado"])
        effective_dir = base_xlsx.parent if base_xlsx else self.root_dir

        items = [
            {
                "key": "cesta_base",
                "label": "1. Base de Conferência Cesta Prêmio Bretas (Excel)",
                "file": self._fmt_info(base_xlsx),
                "req": True,
                "obs": "Abas Conferencia, Frequência, CC, Cargos e Manual",
            },
        ]
        ready = base_xlsx is not None
        raw_map = {
            "data_dir": str(effective_dir),
            "cesta_base": str(base_xlsx) if base_xlsx else None,
        }
        return {"ready": ready, "items": items, "raw_map": raw_map, "month": 10, "year": 2026}

    # =========================================================================
    # DISPATCHER CENTRAL
    # =========================================================================
    def scan_module(self, mod_id: str) -> Dict[str, Any]:
        """Executa a varredura adequada para o módulo selecionado."""
        if mod_id == "good_card":
            return self.scan_good_card()
        elif mod_id == "compra_mensal_vt":
            return self.scan_compra_mensal_vt()
        elif mod_id == "valores_pre_folha":
            return self.scan_valores_pre_folha()
        elif mod_id == "bloqueio_good_card":
            return self.scan_bloqueio_good_card()
        elif mod_id == "unidade_entrega":
            return self.scan_unidade_entrega()
        elif mod_id == "cencopay":
            return self.scan_cencopay()
        elif mod_id == "retorno_lg":
            return self.scan_retorno_lg()
        elif mod_id == "compra_ticket_bretas":
            return self.scan_compra_ticket_bretas()
        elif mod_id == "cesta_premio_bretas":
            return self.scan_cesta_premio_bretas()
        else:
            return {
                "ready": False,
                "items": [],
                "raw_map": {"data_dir": str(self.root_dir)},
                "month": 9,
                "year": 2026,
            }
