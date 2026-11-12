"""
Aplicativo Corporativo Standalone (GUI) — Conferência e Carga de Seguro de Vida Cencosud Brasil.
Padrão Visual Institucional Sistemas RH Cencosud:
- Identidade Oficial Cencosud: Azul Real (#095AA5), Laranja Cencosud (#FF690A) e Logo Oficial em Card Branco.
- Diagnóstico Preventivo de Entradas LG (valida abas 'Nomina' e 'Com gestor', colunas obrigatórias e eventos 6266/6267).
- Guia Integrado de Colunas Obrigatórias LG com cópia rápida para orientação da equipe.
- Execução Assíncrona (Multi-thread), Barra de Progresso Real, Cards Executivos de KPIs, Console de Logs
  e Proteção contra Planilha Aberta no Excel (FileLockModal).
"""

import io
import os
import sys
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any
import tkinter as tk
from tkinter import filedialog, messagebox

# Em modo PyInstaller --windowed no Windows, sys.stdout e sys.stderr podem ser None.
# Redirecionamos para um buffer em memória para evitar erros em chamadas print() internas.
if sys.stdout is None:
    sys.stdout = io.StringIO()
if sys.stderr is None:
    sys.stderr = io.StringIO()

CURRENT_DIR = Path(__file__).resolve().parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

import customtkinter as ctk

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    Image = None
    HAS_PIL = False

from config import (
    BASE_DIR,
    RAW_DIR,
    resource_path,
    CENCOSUD_BLUE,
    CENCOSUD_BLUE_HOVER,
    CENCOSUD_DARK_BLUE,
    CENCOSUD_ORANGE,
    CENCOSUD_ORANGE_HOVER,
    SUCCESS_GREEN,
    SUCCESS_GREEN_HOVER,
)
from file_detector import FileDetector
from main import run_conferencia

# Aparência Corporativa Padrão dos Apps de Sistemas RH Cencosud
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

CARD_BG = "#1E293B"
SUBCARD_BG = "#0F172A"

CHECKLIST_LG_TEXTO = """CHECKLIST DE COLUNAS OBRIGATÓRIAS NAS EXTRAÇÕES DO SISTEMA LG (GEN.TE)
Frente: Conferência e Carga Mensal de Seguro de Vida — Cencosud Brasil
================================================================================

1. RELATÓRIO NOMINA (*Nomina*.xlsx)
   IMPORTANTE: Manter obrigatoriamente as duas abas: "Nomina" e "Com gestor".
   Colunas obrigatórias que NÃO podem ser removidas:
   • Matrícula              (Chave única de cruzamento entre bases)
   • Colaborador            (Nome completo enviado às seguradoras)
   • Data Nascimento        (Enviada na carga e validação etária)
   • CPF/CNPJ ou CPF        (Preservação de 11 dígitos com zeros à esquerda)
   • Data Admissão          (Corte de contratados pós-2015 acima de 65 anos)
   • Idade                  (Validação etária > 65 anos)
   • Categoria              (Identifica Estagiários, Aprendizes, Intermitentes e Autônomos)
   • Vínculo                (Apoio na elegibilidade contratual)
   • Tipo de Salário        (Exclui colaboradores Horistas)
   • Tipo de Situação / Situação (Identifica Afastados INSS, Férias e Desligados)
   • Data de Rescisão       (Exclui colaboradores desligados da carga ativa)
   • Natureza Profissional  (Define Bandeira + identifica 'Diretores' na aba 'Com gestor')
   • Nível Hierárquico      (Apoio 'Executivo' na aba 'Com gestor')
   • Estabelecimento        (Separa Bretas Lojas PLANO 03 x Bretas Postos PLANO 04)
   • Sindicato              (Identifica localidade Jaguaquara PLANO 07)
   • Cargo                  (Auditoria e apoio em Estagiários/Aprendizes)
   • Salário Atual          (Obrigatória para Seguro Complementar 12x e 24x)

2. RELATÓRIO DE OPOSIÇÕES (*Oposi*.CSV ou .xlsx)
   • Matricula (ou Matrícula)
   • Valor Inf Adicional    (Sinalização SIM / 2 - SIM de oposição ativa ao SGV)

3. DESCONTO EM FOLHA — FOPA (*FOPA*.CSV ou *Eventos*.CSV)
   • Matrícula
   • Evento 6266            (Seguro de Vida - Normal)
   • Evento 6267            (Seguro Complementar)
   * Atenção: Selecionar os códigos 6266 e 6267 no filtro de eventos da LG antes de exportar.

4. FATURA DA SEGURADORA (Opcional — *Fatura*.xlsx / *Zurich*.xlsx)
   • Matrícula e coluna de Prêmio / Valor (para conciliação financeira contra o boleto).
"""


class FileLockModal(ctk.CTkToplevel):
    """Janela modal exibida quando uma planilha de saída está aberta/travada no Excel."""

    def __init__(self, parent: ctk.CTk, file_path: str):
        super().__init__(parent)
        self.parent = parent
        self.file_path = file_path
        self.result = "cancel"
        self.event = threading.Event()

        self.title("Aviso: Planilha em Uso no Excel")
        self.geometry("530x245")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.pack(fill="both", expand=True, padx=15, pady=15)

        ctk.CTkLabel(frame, text="⚠️", font=ctk.CTkFont(size=34)).pack(pady=(10, 0))
        ctk.CTkLabel(
            frame,
            text="Planilha Aberta no Microsoft Excel",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(pady=(4, 5))

        fname = os.path.basename(file_path)
        ctk.CTkLabel(
            frame,
            text=(
                f"O arquivo '{fname}' está aberto no Excel ou em uso por outro processo.\n"
                "Feche a planilha no Excel e clique em 'Tentar Novamente', ou salve como cópia."
            ),
            font=ctk.CTkFont(size=12),
            wraplength=470,
            justify="center",
        ).pack(pady=(0, 15))

        btn_box = ctk.CTkFrame(frame, fg_color="transparent")
        btn_box.pack(fill="x", padx=10, pady=(4, 10))

        ctk.CTkButton(
            btn_box,
            text="Tentar Novamente",
            fg_color=CENCOSUD_BLUE,
            hover_color=CENCOSUD_BLUE_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self._on_retry,
        ).pack(side="left", expand=True, padx=5)

        ctk.CTkButton(
            btn_box,
            text="Salvar como Cópia",
            fg_color=SUCCESS_GREEN,
            hover_color=SUCCESS_GREEN_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self._on_copy,
        ).pack(side="left", expand=True, padx=5)

        ctk.CTkButton(
            btn_box,
            text="Cancelar",
            fg_color="#555555",
            hover_color="#777777",
            command=self._on_cancel,
        ).pack(side="left", expand=True, padx=5)

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self._center_window()

    def _center_window(self) -> None:
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _on_retry(self) -> None:
        self.result = "retry"
        self.event.set()
        self.destroy()

    def _on_copy(self) -> None:
        self.result = "copy"
        self.event.set()
        self.destroy()

    def _on_cancel(self) -> None:
        self.result = "cancel"
        self.event.set()
        self.destroy()


class LGColumnsGuideModal(ctk.CTkToplevel):
    """Modal educativo com o checklist de abas e colunas obrigatórias na extração da LG."""

    def __init__(self, parent: ctk.CTk):
        super().__init__(parent)
        self.title("Guia Oficial — Colunas Obrigatórias nas Extrações LG (Seguro de Vida)")
        self.geometry("760x580")
        self.minsize(680, 480)
        self.transient(parent)
        self.grab_set()

        header = ctk.CTkFrame(self, fg_color=CENCOSUD_BLUE, corner_radius=0, height=56)
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(
            header,
            text="📖 Checklist Obrigatório de Abas e Colunas para Geração de Relatórios na LG",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color="#FFFFFF",
        ).pack(side="left", padx=18, pady=12)

        txt = ctk.CTkTextbox(
            self,
            font=ctk.CTkFont(family="Consolas", size=12),
            corner_radius=8,
            wrap="word",
        )
        txt.pack(fill="both", expand=True, padx=16, pady=(14, 10))
        txt.insert("1.0", CHECKLIST_LG_TEXTO)
        txt.configure(state="disabled")

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=16, pady=(0, 14))

        self.lbl_copied = ctk.CTkLabel(footer, text="", font=ctk.CTkFont(size=12, weight="bold"), text_color="#34D399")
        self.lbl_copied.pack(side="left", padx=6)

        ctk.CTkButton(
            footer,
            text="Fechar",
            width=110,
            fg_color="#475569",
            hover_color="#334155",
            command=self.destroy,
        ).pack(side="right", padx=(8, 0))

        ctk.CTkButton(
            footer,
            text="📋 Copiar Checklist",
            width=170,
            fg_color=CENCOSUD_ORANGE,
            hover_color=CENCOSUD_ORANGE_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self._copy_text,
        ).pack(side="right")

    def _copy_text(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(CHECKLIST_LG_TEXTO)
        self.lbl_copied.configure(text="✔ Checklist copiado para a área de transferência!")


class SeguroVidaApp(ctk.CTk):
    """Aplicativo Executivo Cencosud para Conferência e Carga Mensal de Seguro de Vida."""

    def __init__(self):
        super().__init__()
        self.title("SISTEMAS RH CENCOSUD — Conferência e Carga de Seguro de Vida")
        self.geometry("1180x790")
        self.minsize(1020, 700)

        self._configurar_icone()
        self.logo_img = self._carregar_logo()

        # Diretórios padrão inteligentes (compatível com máquina dev e máquinas da equipe RH)
        default_raw = RAW_DIR if RAW_DIR.exists() else (BASE_DIR / "Relatórios brutos" if (BASE_DIR / "Relatórios brutos").exists() else BASE_DIR)
        default_dest = BASE_DIR

        now = datetime.now()
        self.raw_dir_var = tk.StringVar(value=str(default_raw))
        self.dest_dir_var = tk.StringVar(value=str(default_dest))
        self.month_var = tk.StringVar(value=f"{now.month:02d}")
        self.year_var = tk.StringVar(value=str(now.year))
        self._comp_auto_initialized = False

        self.diag_result: Dict[str, Any] = {}
        self.is_running = False
        self._is_scanning = False
        self.last_audit_report: Optional[Path] = None
        self.last_cargas_dir: Optional[Path] = None
        self.last_zip_file: Optional[Path] = None
        self.kpi_value_labels: Dict[str, ctk.CTkLabel] = {}

        self._setup_ui()
        # Agenda a varredura inicial logo após mapear a janela para abertura instantânea da interface
        self.after(150, self.rescan_files)

    def _configurar_icone(self) -> None:
        for ico_name in ("cencosud.ico", "cencosud_logo.ico"):
            ico_path = resource_path(os.path.join("assets", ico_name))
            if ico_path.exists():
                try:
                    self.iconbitmap(str(ico_path))
                    break
                except Exception:
                    pass

    def _carregar_logo(self) -> Optional[ctk.CTkImage]:
        if not HAS_PIL or Image is None:
            return None
        candidates = [
            resource_path(os.path.join("assets", "cencosud_logo_cropped.png")),
            resource_path(os.path.join("assets", "cencosud_logo.png")),
            Path(r"C:\Users\matcaval\OneDrive - Cencosud\Imagens\cencosud-logo-png_seeklogo-251998.png"),
        ]
        for p in candidates:
            if p.exists():
                try:
                    pil_img = Image.open(p).convert("RGBA")
                    w, h = pil_img.size
                    target_h = 44
                    target_w = max(int(w * (target_h / h)), 44)
                    return ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(target_w, target_h))
                except Exception:
                    continue
        return None

    def _setup_ui(self) -> None:
        # ------------------------------------------------------------------
        # 1. CABEÇALHO CORPORATIVO CENCOSUD (#095AA5 + Faixa Laranja #FF690A)
        # ------------------------------------------------------------------
        header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=CENCOSUD_BLUE, height=78)
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)

        header_inner = ctk.CTkFrame(header_frame, fg_color="transparent")
        header_inner.pack(fill="both", expand=True, padx=20, pady=10)

        left_box = ctk.CTkFrame(header_inner, fg_color="transparent")
        left_box.pack(side="left", fill="y", anchor="w")

        if self.logo_img:
            logo_card = ctk.CTkFrame(left_box, fg_color="#FFFFFF", corner_radius=8)
            logo_card.pack(side="left", padx=(0, 16), pady=2)
            ctk.CTkLabel(logo_card, text="", image=self.logo_img).pack(padx=10, pady=4)

        title_box = ctk.CTkFrame(left_box, fg_color="transparent")
        title_box.pack(side="left", fill="y", anchor="w")

        ctk.CTkLabel(
            title_box,
            text="CENCOSUD BRASIL • SISTEMAS RH",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#FFFFFF",
        ).pack(anchor="w", pady=(2, 0))

        ctk.CTkLabel(
            title_box,
            text="Central Executiva de Conferência e Carga de Seguro de Vida (Básico, Diretoria & Complementar 12x/24x)",
            font=ctk.CTkFont(size=12),
            text_color="#D8E8F8",
        ).pack(anchor="w", pady=(0, 2))

        right_box = ctk.CTkFrame(header_inner, fg_color="transparent")
        right_box.pack(side="right", fill="y", anchor="e")

        badge = ctk.CTkFrame(right_box, fg_color=CENCOSUD_DARK_BLUE, corner_radius=15)
        badge.pack(anchor="e", pady=12)
        ctk.CTkLabel(
            badge,
            text=" BENEFÍCIOS RH • V1.0 HOMOLOGADO ",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#BCE0FD",
        ).pack(padx=12, pady=5)

        # Faixa sutil institucional Laranja Cencosud (#FF690A)
        accent_bar = ctk.CTkFrame(self, corner_radius=0, fg_color=CENCOSUD_ORANGE, height=3)
        accent_bar.pack(fill="x", side="top")

        # ------------------------------------------------------------------
        # 2. BARRA DE CONFIGURAÇÃO DE DIRETÓRIOS E COMPETÊNCIA
        # ------------------------------------------------------------------
        config_frame = ctk.CTkFrame(self, corner_radius=10)
        config_frame.pack(fill="x", padx=16, pady=(12, 8))
        config_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            config_frame,
            text="Pasta de Entrada (Relatórios LG):",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=0, column=0, padx=(15, 10), pady=(10, 4), sticky="w")

        ctk.CTkEntry(config_frame, textvariable=self.raw_dir_var, height=32).grid(
            row=0, column=1, padx=(0, 8), pady=(10, 4), sticky="ew"
        )

        ctk.CTkButton(
            config_frame,
            text="Procurar...",
            width=95,
            height=32,
            fg_color=CENCOSUD_BLUE,
            hover_color=CENCOSUD_BLUE_HOVER,
            command=self._browse_raw_dir,
        ).grid(row=0, column=2, padx=(0, 12), pady=(10, 4))

        ctk.CTkLabel(
            config_frame,
            text="Pasta de Destino (Saída / Cargas):",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=1, column=0, padx=(15, 10), pady=(4, 10), sticky="w")

        ctk.CTkEntry(config_frame, textvariable=self.dest_dir_var, height=32).grid(
            row=1, column=1, padx=(0, 8), pady=(4, 10), sticky="ew"
        )

        ctk.CTkButton(
            config_frame,
            text="Alterar...",
            width=95,
            height=32,
            fg_color=CENCOSUD_ORANGE,
            hover_color=CENCOSUD_ORANGE_HOVER,
            command=self._browse_dest_dir,
        ).grid(row=1, column=2, padx=(0, 12), pady=(4, 10))

        # Coluna direita do topo: Competência + Botão Escanear
        right_ctrl = ctk.CTkFrame(config_frame, fg_color="transparent")
        right_ctrl.grid(row=0, column=3, rowspan=2, padx=(4, 15), pady=8, sticky="ns")

        comp_box = ctk.CTkFrame(right_ctrl, fg_color="transparent")
        comp_box.pack(fill="x", pady=(0, 6))

        ctk.CTkLabel(comp_box, text="Competência:", font=ctk.CTkFont(size=12, weight="bold")).pack(
            side="left", padx=(0, 6)
        )
        ctk.CTkOptionMenu(
            comp_box,
            values=[f"{m:02d}" for m in range(1, 13)],
            variable=self.month_var,
            width=66,
            height=28,
            fg_color=CENCOSUD_BLUE,
            button_color=CENCOSUD_DARK_BLUE,
        ).pack(side="left", padx=(0, 4))

        ctk.CTkOptionMenu(
            comp_box,
            values=["2025", "2026", "2027", "2028"],
            variable=self.year_var,
            width=78,
            height=28,
            fg_color=CENCOSUD_BLUE,
            button_color=CENCOSUD_DARK_BLUE,
        ).pack(side="left")

        self.btn_scan = ctk.CTkButton(
            right_ctrl,
            text="🔄 Escanear / Validar LG",
            height=34,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=CENCOSUD_BLUE,
            hover_color=CENCOSUD_BLUE_HOVER,
            command=self.rescan_files,
        )
        self.btn_scan.pack(fill="x")

        # ------------------------------------------------------------------
        # 3. ÁREA DE TRABALHO DIVIDIDA: CHECKLIST LG (ESQ) x EXECUÇÃO (DIR)
        # ------------------------------------------------------------------
        work_area = ctk.CTkFrame(self, fg_color="transparent")
        work_area.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        work_area.grid_columnconfigure(0, weight=4)
        work_area.grid_columnconfigure(1, weight=6)
        work_area.grid_rowconfigure(0, weight=1)

        # --- COLUNA ESQUERDA: Diagnóstico Preventivo de Entradas LG ---
        frame_left = ctk.CTkFrame(work_area, corner_radius=10)
        frame_left.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=0)
        frame_left.grid_columnconfigure(0, weight=1)
        frame_left.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            frame_left,
            text="📋 Diagnóstico Preventivo de Entradas (LG)",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=15, pady=(12, 6))

        self.scroll_checklist = ctk.CTkScrollableFrame(frame_left, fg_color="transparent")
        self.scroll_checklist.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 6))
        self.scroll_checklist.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(
            frame_left,
            text="📖 Ver Colunas Obrigatórias LG (Guia para Equipe)",
            height=36,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#334155",
            hover_color="#475569",
            command=self._open_lg_guide,
        ).grid(row=2, column=0, sticky="ew", padx=14, pady=(4, 12))

        # --- COLUNA DIREITA: Execução, KPIs e Console em Tempo Real ---
        frame_right = ctk.CTkFrame(work_area, corner_radius=10)
        frame_right.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=0)
        frame_right.grid_columnconfigure(0, weight=1)
        frame_right.grid_rowconfigure(4, weight=1)

        # Botão Principal de Disparo
        self.btn_run = ctk.CTkButton(
            frame_right,
            text="▶ GERAR CARGAS POR BANDEIRA E CONFERÊNCIA DE SEGURO DE VIDA",
            height=44,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=CENCOSUD_BLUE,
            hover_color=CENCOSUD_BLUE_HOVER,
            command=self.start_execution,
        )
        self.btn_run.grid(row=0, column=0, sticky="ew", padx=15, pady=(14, 6))

        self.lbl_progress_status = ctk.CTkLabel(
            frame_right,
            text="Status: Pronto para processar.",
            font=ctk.CTkFont(size=12, slant="italic"),
            anchor="w",
        )
        self.lbl_progress_status.grid(row=1, column=0, sticky="ew", padx=15, pady=(2, 2))

        self.progress_bar = ctk.CTkProgressBar(frame_right, height=12, progress_color=CENCOSUD_ORANGE)
        self.progress_bar.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 10))
        self.progress_bar.set(0.0)

        # Faixa de Mini-Cards Executivos de KPIs
        kpi_strip = ctk.CTkFrame(frame_right, fg_color="transparent")
        kpi_strip.grid(row=3, column=0, sticky="ew", padx=15, pady=(0, 8))
        kpi_defs = [
            ("nomina", "Nomina (+ Gestor)", "#93C5FD"),
            ("elegiveis", "Elegíveis Básico", "#34D399"),
            ("diretores", "Diretores (GB)", "#FBBF24"),
            ("excluidos", "Excluídos Regra", "#F87171"),
            ("complementar", "Comp 12x / 24x", "#C4B5FD"),
        ]
        for idx, (key, title, color) in enumerate(kpi_defs):
            kpi_strip.grid_columnconfigure(idx, weight=1)
            box = ctk.CTkFrame(kpi_strip, corner_radius=8, fg_color=SUBCARD_BG)
            box.grid(row=0, column=idx, sticky="ew", padx=3)
            ctk.CTkLabel(
                box,
                text=title,
                font=ctk.CTkFont(size=10, weight="bold"),
                text_color="#94A3B8",
            ).pack(pady=(6, 0), padx=6)
            lbl_val = ctk.CTkLabel(
                box,
                text="—",
                font=ctk.CTkFont(size=15, weight="bold"),
                text_color=color,
            )
            lbl_val.pack(pady=(0, 6), padx=6)
            self.kpi_value_labels[key] = lbl_val

        # Console de Logs Integrado
        self.txt_log = ctk.CTkTextbox(
            frame_right,
            font=ctk.CTkFont(family="Consolas", size=11),
            corner_radius=8,
            wrap="word",
        )
        self.txt_log.grid(row=4, column=0, sticky="nsew", padx=15, pady=4)
        self.append_log("Central de Seguro de Vida inicializada. Verifique o checklist de entradas ao lado.")

        # Botões Pós-Processamento
        frame_post = ctk.CTkFrame(frame_right, fg_color="transparent")
        frame_post.grid(row=5, column=0, sticky="ew", padx=15, pady=(8, 14))
        frame_post.grid_columnconfigure(0, weight=2)
        frame_post.grid_columnconfigure(1, weight=2)
        frame_post.grid_columnconfigure(2, weight=1)

        self.btn_open_audit = ctk.CTkButton(
            frame_post,
            text="📊 Abrir Relatório Conferência",
            height=38,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=SUCCESS_GREEN,
            hover_color=SUCCESS_GREEN_HOVER,
            state="disabled",
            command=self._open_audit_report,
        )
        self.btn_open_audit.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.btn_open_cargas = ctk.CTkButton(
            frame_post,
            text="📦 Abrir Pasta Cargas (8 Planilhas)",
            height=38,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=CENCOSUD_BLUE,
            hover_color=CENCOSUD_BLUE_HOVER,
            command=self._open_cargas_folder,
        )
        self.btn_open_cargas.grid(row=0, column=1, sticky="ew", padx=5)

        self.btn_open_dest = ctk.CTkButton(
            frame_post,
            text="📁 Pasta Raiz",
            height=38,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#334155",
            hover_color="#475569",
            command=self._open_dest_folder,
        )
        self.btn_open_dest.grid(row=0, column=2, sticky="ew", padx=(5, 0))

    def append_log(self, msg: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.txt_log.configure(state="normal")
        for line in str(msg).splitlines():
            if line.strip():
                self.txt_log.insert("end", f"[{timestamp}] {line}\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _browse_raw_dir(self) -> None:
        curr = self.raw_dir_var.get()
        chosen = filedialog.askdirectory(
            initialdir=curr if os.path.exists(curr) else os.getcwd(),
            title="Selecione a Pasta com os Relatórios Brutos da LG (Nomina, Oposição, FOPA)",
        )
        if chosen:
            self.raw_dir_var.set(chosen)
            self._comp_auto_initialized = False
            self.rescan_files()

    def _browse_dest_dir(self) -> None:
        curr = self.dest_dir_var.get()
        chosen = filedialog.askdirectory(
            initialdir=curr if os.path.exists(curr) else os.getcwd(),
            title="Selecione a Pasta de Destino das Cargas e do Relatório de Conferência",
        )
        if chosen:
            self.dest_dir_var.set(chosen)
            self.append_log(f"Pasta de destino definida para: {chosen}")

    def _open_lg_guide(self) -> None:
        LGColumnsGuideModal(self)

    def rescan_files(self) -> None:
        """Varre a pasta de entrada em background thread e atualiza os cartões preventivos LG sem travar a UI."""
        if self.is_running or self._is_scanning:
            return
        self._is_scanning = True
        raw_dir = self.raw_dir_var.get().strip()
        self.btn_scan.configure(state="disabled", text="Validando...")
        self.btn_run.configure(state="disabled")
        for w in self.scroll_checklist.winfo_children():
            w.destroy()
        ctk.CTkLabel(
            self.scroll_checklist,
            text="🔄 Inspecionando relatórios e colunas LG na pasta de entrada...",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#93C5FD",
        ).pack(pady=18, padx=10)

        def _scan_worker() -> None:
            try:
                res = FileDetector(raw_dir).scan()
            except Exception as exc:
                res = {"ready": False, "items": [], "competencia_hint": None, "raw_map": {}, "error": str(exc)}
            try:
                if self.winfo_exists():
                    self.after(0, lambda r=res: self._apply_scan_result(r))
            except Exception:
                pass

        threading.Thread(target=_scan_worker, daemon=True).start()

    def _apply_scan_result(self, res: Dict[str, Any]) -> None:
        self._is_scanning = False
        if not self.winfo_exists():
            return
        self.diag_result = res

        # Ajusta sugestão inicial de competência (ex: 10/2026) na primeira varredura da pasta
        hint = self.diag_result.get("competencia_hint")
        if hint and not self._comp_auto_initialized:
            self.month_var.set(hint[0])
            self.year_var.set(hint[1])
            self._comp_auto_initialized = True

        self._render_checklist(self.diag_result)
        self.btn_scan.configure(state="normal", text="🔄 Escanear / Validar LG")

    def _render_checklist(self, scan_data: Dict[str, Any]) -> None:
        for w in self.scroll_checklist.winfo_children():
            w.destroy()

        items = scan_data.get("items", [])
        ready = scan_data.get("ready", False)

        for item in items:
            file_info = item.get("file")
            status = item.get("status", "MISSING")
            label = item.get("label", "")
            obs = item.get("obs", "")

            card = ctk.CTkFrame(self.scroll_checklist, corner_radius=8, fg_color=CARD_BG)
            card.pack(fill="x", pady=4, padx=2)
            card.grid_columnconfigure(1, weight=1)

            if status == "OK" and file_info is not None:
                icon_text = "🟢"
                title_color = "#FFFFFF"
                detail_color = "#34D399"
                fname = file_info.get("name", "")
                fsize = file_info.get("size", "")
                detail_text = f"{fname} ({fsize})"
                if obs:
                    detail_text += f"\n{obs}"
            elif status in ("WARNING", "OPTIONAL", "MISSING_OPT"):
                icon_text = "🟡"
                title_color = "#FFFFFF"
                detail_color = "#FBBF24"
                if file_info is not None:
                    detail_text = f"{file_info.get('name', '')} ({file_info.get('size', '')})\n{obs}"
                else:
                    detail_text = obs
            else:
                icon_text = "🔴"
                title_color = "#FCA5A5"
                detail_color = "#F87171"
                if file_info is not None:
                    detail_text = f"{file_info.get('name', '')}\n{obs}"
                else:
                    detail_text = obs

            ctk.CTkLabel(card, text=icon_text, font=ctk.CTkFont(size=15)).grid(
                row=0, column=0, rowspan=2, padx=(10, 8), pady=8
            )

            ctk.CTkLabel(
                card,
                text=label,
                font=ctk.CTkFont(size=11, weight="bold"),
                text_color=title_color,
                anchor="w",
            ).grid(row=0, column=1, sticky="w", padx=(0, 8), pady=(6, 1))

            ctk.CTkLabel(
                card,
                text=detail_text,
                font=ctk.CTkFont(size=10),
                text_color=detail_color,
                anchor="w",
                justify="left",
                wraplength=340,
            ).grid(row=1, column=1, sticky="w", padx=(0, 8), pady=(0, 7))

        if ready:
            self.btn_run.configure(
                state="normal",
                fg_color=CENCOSUD_BLUE,
                text="▶ GERAR CARGAS POR BANDEIRA E CONFERÊNCIA DE SEGURO DE VIDA",
            )
            self.lbl_progress_status.configure(
                text="Pronto: Bases obrigatórias da LG validadas com sucesso."
            )
        else:
            self.btn_run.configure(
                state="disabled",
                fg_color="#475569",
                text="Aguardando Arquivos / Colunas Obrigatórias LG (Verifique os itens em vermelho)",
            )
            self.lbl_progress_status.configure(
                text="Atenção: Corrija os itens em vermelho no diagnóstico ao lado para liberar a geração."
            )

    def start_execution(self) -> None:
        if self.is_running:
            return
        raw_dir = Path(self.raw_dir_var.get().strip())
        dest_dir = Path(self.dest_dir_var.get().strip())
        competencia = f"{self.month_var.get()}{self.year_var.get()}"
        raw_map = dict(self.diag_result.get("raw_map", {}))

        if not raw_dir.exists():
            messagebox.showwarning("Pasta Inválida", "A pasta de entrada informada não existe.")
            return

        self.is_running = True
        self.btn_run.configure(state="disabled")
        self.btn_scan.configure(state="disabled")
        self.btn_open_audit.configure(state="disabled")
        self.progress_bar.set(0.02)
        self.lbl_progress_status.configure(text=f"Iniciando processamento da competência {competencia}...")
        self.append_log(f">>> Disparando conferência e geração de cargas — Competência {competencia}")

        def worker() -> None:
            try:
                def progress_cb(pct: float, text: str) -> None:
                    self.after(0, lambda p=pct, t=text: self._update_progress(p, t))

                def log_cb(msg: str) -> None:
                    self.after(0, lambda m=msg: self.append_log(m))

                def prompt_lock_cb(file_path: str) -> str:
                    holder = [None]

                    def open_modal() -> None:
                        holder[0] = FileLockModal(self, file_path)

                    self.after(0, open_modal)
                    while holder[0] is None or not holder[0].event.is_set():
                        time.sleep(0.1)
                    return holder[0].result

                res = run_conferencia(
                    competencia=competencia,
                    raw_dir=raw_dir,
                    dest_dir=dest_dir,
                    raw_map=raw_map,
                    progress_cb=progress_cb,
                    log_cb=log_cb,
                    prompt_lock_cb=prompt_lock_cb,
                )
                self.after(0, lambda r=res: self._on_success(r))
            except Exception as ex:
                err_trace = traceback.format_exc()
                self.after(0, lambda msg=str(ex), tr=err_trace: self._on_error(msg, tr))

        threading.Thread(target=worker, daemon=True).start()

    def _update_progress(self, pct: float, text: str) -> None:
        clamped = max(0.0, min(1.0, float(pct)))
        self.progress_bar.set(clamped)
        self.lbl_progress_status.configure(text=f"{int(clamped * 100)}% — {text}")

    def _on_success(self, res: Dict[str, Any]) -> None:
        self.is_running = False
        self.btn_run.configure(state="normal")
        self.btn_scan.configure(state="normal")
        self.progress_bar.set(1.0)
        self.lbl_progress_status.configure(
            text=f"✔ Processamento concluído em {res['tempo_execucao']:.1f}s! Todas as cargas e auditoria geradas."
        )

        kpis = res["kpis"]
        self.kpi_value_labels["nomina"].configure(text=f"{kpis['total_nomina']:,}".replace(",", "."))
        self.kpi_value_labels["elegiveis"].configure(text=f"{kpis['total_elegiveis']:,}".replace(",", "."))
        self.kpi_value_labels["diretores"].configure(text=str(kpis.get("total_diretores_elegiveis", 0)))
        self.kpi_value_labels["excluidos"].configure(text=f"{kpis['total_excluidos']:,}".replace(",", "."))
        self.kpi_value_labels["complementar"].configure(
            text=f"{kpis['total_comp_12x']} / {kpis['total_comp_24x']}"
        )

        arqs = res.get("arquivos_gerados", {})
        self.last_audit_report = arqs.get("relatorio_auditoria")
        self.last_cargas_dir = res.get("output_dir")
        self.last_zip_file = arqs.get("pacote_zip")

        if self.last_audit_report and Path(self.last_audit_report).exists():
            self.btn_open_audit.configure(state="normal")

        messagebox.showinfo(
            "Conferência de Seguro de Vida Concluída",
            f"Processamento finalizado com sucesso!\n\n"
            f"• Total na Nomina (+ Com gestor): {kpis['total_nomina']:,}\n"
            f"• Vidas Elegíveis (Seguro Básico): {kpis['total_elegiveis']:,}\n"
            f"• Diretores no Básico (GB): {kpis.get('total_diretores_elegiveis', 0)}\n"
            f"• Exclusões por Regra: {kpis['total_excluidos']:,}\n"
            f"• Complementar 12x / 24x: {kpis['total_comp_12x']} / {kpis['total_comp_24x']} vidas\n\n"
            f"As 8 planilhas de carga, o arquivo .ZIP e o Relatório Master de Conferência "
            f"já estão prontos na pasta de destino!".replace(",", "."),
        )

    def _on_error(self, err_msg: str, full_trace: str) -> None:
        self.is_running = False
        self.btn_run.configure(state="normal")
        self.btn_scan.configure(state="normal")
        self.progress_bar.set(0.0)
        self.lbl_progress_status.configure(text="❌ Erro durante o processamento (verifique o log).")
        self.append_log(f"ERRO: {err_msg}")
        self.append_log(full_trace)
        messagebox.showerror(
            "Erro na Execução",
            f"Não foi possível concluir o processamento:\n\n{err_msg}",
        )

    def _open_audit_report(self) -> None:
        if self.last_audit_report and Path(self.last_audit_report).exists():
            try:
                os.startfile(str(self.last_audit_report))
            except Exception as ex:
                messagebox.showerror("Erro ao Abrir", f"Não foi possível abrir o relatório:\n{ex}")
        else:
            messagebox.showwarning("Aviso", "Gere a conferência primeiro para abrir o relatório.")

    def _open_cargas_folder(self) -> None:
        target = self.last_cargas_dir or (Path(self.dest_dir_var.get().strip()) / "Carga Seguro de Vida Gerada")
        target.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(target))
        except Exception as ex:
            messagebox.showerror("Erro ao Abrir Pasta", f"Não foi possível abrir a pasta de cargas:\n{ex}")

    def _open_dest_folder(self) -> None:
        target = Path(self.dest_dir_var.get().strip())
        target.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(target))
        except Exception as ex:
            messagebox.showerror("Erro ao Abrir Pasta", f"Não foi possível abrir a pasta raiz:\n{ex}")


def main() -> None:
    app = SeguroVidaApp()
    app.mainloop()


if __name__ == "__main__":
    main()
