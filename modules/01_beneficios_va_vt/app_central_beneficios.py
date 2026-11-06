"""
Aplicativo Executivo Unificado - Central de Auditoria de Benefícios Cencosud Brasil.

Segue o padrão corporativo CustomTkinter dos projetos de Sistemas RH Cencosud:
- Cabeçalho institucional com logo oficial e paleta Cencosud (#095AA5).
- Arquitetura Modular com Feature Flags (Sidebar de Módulos):
  • Módulos Ativos (ex: 01 - Good Card): Checklist inteligente de entradas, validação
    preventiva, thread assíncrona, barra de progresso real, console de logs, FileLockModal
    e atalhos de abertura pós-processamento.
  • Módulos em Homologação: Painel executivo de status e escopo, prontos para ativação
    imediata via flag assim que validados pela equipe.
"""

import os
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    Image = None
    HAS_PIL = False

from file_detector import FileDetector, resource_path
from good_card_engine_adapter import exec_auditoria_good_card
from compra_mensal_vt_adapter import exec_auditoria_compra_mensal_vt
from valores_pre_folha_adapter import exec_auditoria_valores_pre_folha
from bloqueio_good_card_adapter import exec_auditoria_bloqueio_good_card
from unidade_entrega_adapter import exec_auditoria_unidade_entrega
from cencopay_adapter import exec_auditoria_cencopay
from retorno_lg_adapter import exec_auditoria_retorno_lg
from compra_ticket_bretas_adapter import exec_auditoria_compra_ticket_bretas
from cesta_premio_bretas_adapter import exec_auditoria_cesta_premio_bretas
from module_registry import MODULES_CONFIG

ENGINES_MAP = {
    "good_card": exec_auditoria_good_card,
    "compra_mensal_vt": exec_auditoria_compra_mensal_vt,
    "valores_pre_folha": exec_auditoria_valores_pre_folha,
    "bloqueio_good_card": exec_auditoria_bloqueio_good_card,
    "unidade_entrega": exec_auditoria_unidade_entrega,
    "cencopay": exec_auditoria_cencopay,
    "retorno_lg": exec_auditoria_retorno_lg,
    "compra_ticket_bretas": exec_auditoria_compra_ticket_bretas,
    "cesta_premio_bretas": exec_auditoria_cesta_premio_bretas,
}

# Configurações de Aparência Corporativa CustomTkinter
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# Paleta Oficial Cencosud
CENCOSUD_BLUE = "#095AA5"
CENCOSUD_BLUE_HOVER = "#0B6FCB"
CENCOSUD_DARK_BLUE = "#063A6B"
SUCCESS_GREEN = "#1E7B34"
SUCCESS_GREEN_HOVER = "#28A745"
SIDEBAR_BG = "#111827"
CARD_BG = "#1E293B"


class FileLockModal(ctk.CTkToplevel):
    """Janela modal amigável exibida quando o Excel trava a planilha final."""

    def __init__(self, parent, file_path: str):
        super().__init__(parent)
        self.parent = parent
        self.file_path = file_path
        self.result = "cancel"
        self.event = threading.Event()

        self.title("Aviso: Arquivo Aberto no Excel")
        self.geometry("520x240")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.pack(fill="both", expand=True, padx=15, pady=15)

        lbl_icon = ctk.CTkLabel(frame, text="⚠️", font=ctk.CTkFont(size=36))
        lbl_icon.pack(pady=(10, 0))

        lbl_title = ctk.CTkLabel(
            frame,
            text="Planilha em Uso no Excel",
            font=ctk.CTkFont(size=16, weight="bold"),
        )
        lbl_title.pack(pady=(5, 5))

        fname = os.path.basename(file_path)
        lbl_msg = ctk.CTkLabel(
            frame,
            text=(
                f"A planilha '{fname}' está aberta no Excel ou por outro usuário.\n"
                "Por favor, feche o arquivo para prosseguir com a gravação."
            ),
            font=ctk.CTkFont(size=12),
            wraplength=460,
            justify="center",
        )
        lbl_msg.pack(pady=(0, 15))

        btn_box = ctk.CTkFrame(frame, fg_color="transparent")
        btn_box.pack(fill="x", padx=10, pady=(5, 10))

        btn_retry = ctk.CTkButton(
            btn_box,
            text="Tentar Novamente",
            fg_color=CENCOSUD_BLUE,
            hover_color=CENCOSUD_BLUE_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self.on_retry,
        )
        btn_retry.pack(side="left", expand=True, padx=5)

        btn_copy = ctk.CTkButton(
            btn_box,
            text="Salvar como Cópia",
            fg_color=SUCCESS_GREEN,
            hover_color=SUCCESS_GREEN_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self.on_copy,
        )
        btn_copy.pack(side="left", expand=True, padx=5)

        btn_cancel = ctk.CTkButton(
            btn_box,
            text="Cancelar",
            fg_color="#555555",
            hover_color="#777777",
            command=self.on_cancel,
        )
        btn_cancel.pack(side="left", expand=True, padx=5)

        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.center_window()

    def center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def on_retry(self):
        self.result = "retry"
        self.event.set()
        self.destroy()

    def on_copy(self):
        self.result = "copy"
        self.event.set()
        self.destroy()

    def on_cancel(self):
        self.result = "cancel"
        self.event.set()
        self.destroy()


class ActiveModuleView(ctk.CTkFrame):
    """Painel operacional completo para módulos homologados e ativos (ex: Good Card)."""

    def __init__(self, master, module_cfg: dict, engine_fn, app_instance):
        super().__init__(master, fg_color="transparent")
        self.module_cfg = module_cfg
        self.engine_fn = engine_fn
        self.app = app_instance
        self.generated_file = None

        self.setup_ui()

    def setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # -------------------------------------------------------------
        # 1. BANNER DO MÓDULO ATIVO
        # -------------------------------------------------------------
        banner = ctk.CTkFrame(self, corner_radius=10, fg_color=CARD_BG)
        banner.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 8))
        banner.grid_columnconfigure(0, weight=1)

        top_row = ctk.CTkFrame(banner, fg_color="transparent")
        top_row.grid(row=0, column=0, sticky="ew", padx=15, pady=(10, 2))
        top_row.grid_columnconfigure(0, weight=1)

        lbl_title = ctk.CTkLabel(
            top_row,
            text=self.module_cfg["full_title"],
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color="#FFFFFF",
            anchor="w",
        )
        lbl_title.grid(row=0, column=0, sticky="w")

        badge = ctk.CTkFrame(top_row, fg_color="#065F46", corner_radius=12)
        badge.grid(row=0, column=1, sticky="e")
        ctk.CTkLabel(
            badge,
            text=f" {self.module_cfg['status_badge']} • HOMOLOGADO ",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#A7F3D0",
        ).pack(padx=10, pady=3)

        lbl_desc = ctk.CTkLabel(
            banner,
            text=self.module_cfg["description"],
            font=ctk.CTkFont(size=11),
            text_color="#94A3B8",
            anchor="w",
            justify="left",
            wraplength=860,
        )
        lbl_desc.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 10))

        # -------------------------------------------------------------
        # 2. BARRA DE DIRETÓRIOS E COMPETÊNCIA
        # -------------------------------------------------------------
        config_frame = ctk.CTkFrame(self, corner_radius=10)
        config_frame.grid(row=1, column=0, sticky="ew", padx=0, pady=(0, 8))
        config_frame.grid_columnconfigure(1, weight=1)

        lbl_root = ctk.CTkLabel(
            config_frame,
            text="Pasta de Entrada (Relatórios):",
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        lbl_root.grid(row=0, column=0, padx=(15, 10), pady=(10, 4), sticky="w")

        entry_root = ctk.CTkEntry(config_frame, textvariable=self.app.root_var, height=30)
        entry_root.grid(row=0, column=1, padx=(0, 8), pady=(10, 4), sticky="ew")

        btn_browse_root = ctk.CTkButton(
            config_frame,
            text="Procurar...",
            width=85,
            height=30,
            command=self.app.browse_root_folder,
        )
        btn_browse_root.grid(row=0, column=2, padx=(0, 10), pady=(10, 4))

        lbl_out = ctk.CTkLabel(
            config_frame,
            text="Pasta de Destino (Saída):",
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        lbl_out.grid(row=1, column=0, padx=(15, 10), pady=(4, 10), sticky="w")

        entry_out = ctk.CTkEntry(config_frame, textvariable=self.app.output_var, height=30)
        entry_out.grid(row=1, column=1, padx=(0, 8), pady=(4, 10), sticky="ew")

        btn_browse_out = ctk.CTkButton(
            config_frame,
            text="Procurar...",
            width=85,
            height=30,
            command=self.app.browse_output_folder,
        )
        btn_browse_out.grid(row=1, column=2, padx=(0, 10), pady=(4, 10))

        # Seletor de Competência + Botão Escanear
        right_ctrl = ctk.CTkFrame(config_frame, fg_color="transparent")
        right_ctrl.grid(row=0, column=3, rowspan=2, padx=(4, 15), pady=8, sticky="ns")

        comp_box = ctk.CTkFrame(right_ctrl, fg_color="transparent")
        comp_box.pack(fill="x", pady=(0, 4))

        ctk.CTkLabel(comp_box, text="Competência:", font=ctk.CTkFont(size=11, weight="bold")).pack(
            side="left", padx=(0, 6)
        )
        self.cmb_month = ctk.CTkOptionMenu(
            comp_box,
            values=[f"{m:02d}" for m in range(1, 13)],
            variable=self.app.month_var,
            width=62,
            height=26,
        )
        self.cmb_month.pack(side="left", padx=(0, 4))

        self.cmb_year = ctk.CTkOptionMenu(
            comp_box,
            values=["2025", "2026", "2027"],
            variable=self.app.year_var,
            width=74,
            height=26,
        )
        self.cmb_year.pack(side="left")

        self.btn_scan = ctk.CTkButton(
            right_ctrl,
            text="🔄 Escanear / Atualizar Arquivos",
            height=32,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=CENCOSUD_BLUE,
            hover_color=CENCOSUD_BLUE_HOVER,
            command=self.app.rescan_files,
        )
        self.btn_scan.pack(fill="x")

        # -------------------------------------------------------------
        # 3. ÁREA DIVIDIDA: CHECKLIST (ESQUERDA) E EXECUÇÃO/LOGS (DIREITA)
        # -------------------------------------------------------------
        work_area = ctk.CTkFrame(self, fg_color="transparent")
        work_area.grid(row=2, column=0, sticky="nsew")
        work_area.grid_columnconfigure(0, weight=4)
        work_area.grid_columnconfigure(1, weight=5)
        work_area.grid_rowconfigure(0, weight=1)

        # Coluna Esquerda: Diagnóstico de Entradas
        frame_left = ctk.CTkFrame(work_area, corner_radius=10)
        frame_left.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=0)
        frame_left.grid_columnconfigure(0, weight=1)
        frame_left.grid_rowconfigure(1, weight=1)

        lbl_chk_header = ctk.CTkLabel(
            frame_left,
            text=f"📋 Diagnóstico de Entradas ({self.module_cfg['short_name']})",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        lbl_chk_header.grid(row=0, column=0, sticky="w", padx=15, pady=(12, 6))

        self.scroll_checklist = ctk.CTkScrollableFrame(frame_left, fg_color="transparent")
        self.scroll_checklist.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self.scroll_checklist.grid_columnconfigure(0, weight=1)

        # Coluna Direita: Execução, Progresso e Console de Logs
        frame_right = ctk.CTkFrame(work_area, corner_radius=10)
        frame_right.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=0)
        frame_right.grid_columnconfigure(0, weight=1)
        frame_right.grid_rowconfigure(3, weight=1)

        frame_action = ctk.CTkFrame(frame_right, fg_color="transparent")
        frame_action.grid(row=0, column=0, sticky="ew", padx=15, pady=(12, 4))
        frame_action.grid_columnconfigure(0, weight=1)

        self.btn_run = ctk.CTkButton(
            frame_action,
            text=f"▶ Gerar {self.module_cfg['full_title']}",
            height=40,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=CENCOSUD_BLUE,
            hover_color=CENCOSUD_BLUE_HOVER,
            command=self.start_execution,
        )
        self.btn_run.grid(row=0, column=0, sticky="ew")

        self.lbl_progress_status = ctk.CTkLabel(
            frame_right,
            text="Status: Pronto para processar.",
            font=ctk.CTkFont(size=11, slant="italic"),
            anchor="w",
        )
        self.lbl_progress_status.grid(row=1, column=0, sticky="ew", padx=15, pady=(4, 2))

        self.progress_bar = ctk.CTkProgressBar(frame_right, height=11)
        self.progress_bar.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 8))
        self.progress_bar.set(0.0)

        self.txt_log = ctk.CTkTextbox(
            frame_right,
            font=ctk.CTkFont(family="Consolas", size=11),
            corner_radius=8,
            wrap="word",
        )
        self.txt_log.grid(row=3, column=0, sticky="nsew", padx=15, pady=4)
        self.append_log(f"Módulo '{self.module_cfg['short_name']}' inicializado e pronto.")

        frame_post = ctk.CTkFrame(frame_right, fg_color="transparent")
        frame_post.grid(row=4, column=0, sticky="ew", padx=15, pady=(8, 12))
        frame_post.grid_columnconfigure(0, weight=1)
        frame_post.grid_columnconfigure(1, weight=1)

        self.btn_open_file = ctk.CTkButton(
            frame_post,
            text="📄 Abrir Planilha Gerada",
            font=ctk.CTkFont(weight="bold"),
            fg_color=SUCCESS_GREEN,
            hover_color=SUCCESS_GREEN_HOVER,
            state="disabled",
            command=self.open_generated_file,
        )
        self.btn_open_file.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.btn_open_dir = ctk.CTkButton(
            frame_post,
            text="📁 Abrir Pasta de Destino",
            font=ctk.CTkFont(weight="bold"),
            fg_color="#333333",
            hover_color="#4F4F4F",
            command=self.open_output_folder,
        )
        self.btn_open_dir.grid(row=0, column=1, sticky="ew", padx=(5, 0))

    def append_log(self, msg: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", f"[{timestamp}] {msg}\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def update_checklist(self, scan_data: dict):
        for w in self.scroll_checklist.winfo_children():
            w.destroy()

        items = scan_data.get("items", [])
        ready = scan_data.get("ready", False)

        for item in items:
            file_info = item.get("file")
            label = item.get("label", "")
            obs = item.get("obs", "")

            card = ctk.CTkFrame(self.scroll_checklist, corner_radius=8, fg_color=CARD_BG)
            card.pack(fill="x", pady=4, padx=2)
            card.grid_columnconfigure(1, weight=1)

            if file_info is not None:
                icon_text = "🟢"
                title_color = "#FFFFFF"
                detail_color = "#34D399"
                fname = file_info.get("name", "")
                fsize = file_info.get("size", "")
                detail_text = f"{fname} ({fsize})"
                if obs:
                    detail_text += f" • {obs}"
            elif not item.get("req", True):
                icon_text = "🟡"
                title_color = "#FFFFFF"
                detail_color = "#FBBF24"
                detail_text = obs or "Arquivo opcional"
            else:
                icon_text = "🔴"
                title_color = "#FCA5A5"
                detail_color = "#F87171"
                detail_text = "Arquivo obrigatório não localizado na pasta"

            lbl_ico = ctk.CTkLabel(card, text=icon_text, font=ctk.CTkFont(size=14))
            lbl_ico.grid(row=0, column=0, rowspan=2, padx=(10, 8), pady=7)

            lbl_name = ctk.CTkLabel(
                card,
                text=label,
                font=ctk.CTkFont(size=11, weight="bold"),
                text_color=title_color,
                anchor="w",
            )
            lbl_name.grid(row=0, column=1, sticky="w", padx=0, pady=(5, 0))

            lbl_detail = ctk.CTkLabel(
                card,
                text=detail_text,
                font=ctk.CTkFont(size=10),
                text_color=detail_color,
                anchor="w",
            )
            lbl_detail.grid(row=1, column=1, sticky="w", padx=0, pady=(0, 5))

        if ready:
            self.btn_run.configure(
                state="normal",
                fg_color=CENCOSUD_BLUE,
                text=f"▶ Gerar {self.module_cfg['full_title']}",
            )
            self.lbl_progress_status.configure(
                text=f"Pronto: Arquivos obrigatórios de {self.module_cfg['short_name']} identificados."
            )
        else:
            self.btn_run.configure(
                state="disabled",
                fg_color="#444444",
                text="Aguardando Arquivos Obrigatórios (Ver Checklist ao Lado)",
            )
            self.lbl_progress_status.configure(
                text="Atenção: Selecione a pasta contendo os arquivos em vermelho para liberar a execução."
            )

    def start_execution(self):
        self.btn_run.configure(state="disabled")
        self.btn_open_file.configure(state="disabled")
        self.progress_bar.set(0.02)
        self.lbl_progress_status.configure(text="Iniciando processamento em segundo plano...")
        self.append_log(f">>> Disparando motor de auditoria para: {self.module_cfg['short_name']}")

        raw_map = self.app.diag_result.get("raw_map", {})
        output_dir = self.app.get_output_dir()
        try:
            month = int(self.app.month_var.get())
            year = int(self.app.year_var.get())
        except ValueError:
            month, year = 9, 2026

        def worker():
            try:
                def progress_cb(pct, text):
                    self.app.after(0, lambda: self._update_progress(pct, text))

                def log_cb(msg):
                    self.app.after(0, lambda: self.append_log(msg))

                def prompt_lock_cb(file_path):
                    modal = [None]

                    def open_modal():
                        modal[0] = FileLockModal(self.app, file_path)

                    self.app.after(0, open_modal)
                    while modal[0] is None or not modal[0].event.is_set():
                        time.sleep(0.1)
                    return modal[0].result

                saved = self.engine_fn(
                    raw_map=raw_map,
                    output_dir=output_dir,
                    month=month,
                    year=year,
                    progress_cb=progress_cb,
                    log_cb=log_cb,
                    prompt_lock_cb=prompt_lock_cb,
                )
                self.app.after(0, lambda: self._on_success(saved))
            except Exception as e:
                import traceback
                err = traceback.format_exc()
                self.app.after(0, lambda: self._on_error(str(e), err))

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

    def _update_progress(self, pct: float, text: str):
        self.progress_bar.set(pct)
        self.lbl_progress_status.configure(text=f"{int(pct * 100)}% - {text}")

    def _on_success(self, file_path: str):
        self.generated_file = file_path
        self.btn_run.configure(state="normal")
        self.btn_open_file.configure(state="normal")
        self.progress_bar.set(1.0)
        self.lbl_progress_status.configure(text="Concluído com sucesso!")
        self.append_log(f"Processo finalizado! Arquivo disponível: {os.path.basename(file_path)}")
        messagebox.showinfo(
            "Auditoria Concluída",
            f"Relatório {self.module_cfg['short_name']} gerado com sucesso!\n\nArquivo:\n{file_path}",
        )

    def _on_error(self, err_msg: str, full_trace: str):
        self.btn_run.configure(state="normal")
        self.progress_bar.set(0.0)
        self.lbl_progress_status.configure(text="Erro durante o processamento!")
        self.append_log(f"ERRO: {err_msg}")
        self.append_log(full_trace)
        messagebox.showerror("Erro de Execução", f"Ocorreu um erro durante a geração:\n\n{err_msg}")

    def open_generated_file(self):
        if self.generated_file and os.path.exists(self.generated_file):
            try:
                os.startfile(self.generated_file)
            except Exception as e:
                messagebox.showerror("Erro ao Abrir", f"Não foi possível abrir o arquivo:\n{e}")
        else:
            messagebox.showwarning("Aviso", "O arquivo gerado ainda não está disponível.")

    def open_output_folder(self):
        out_dir = self.app.get_output_dir()
        if os.path.exists(out_dir):
            try:
                os.startfile(out_dir)
            except Exception as e:
                messagebox.showerror("Erro ao Abrir", f"Não foi possível abrir a pasta:\n{e}")
        else:
            messagebox.showwarning("Aviso", "A pasta de destino ainda não foi criada.")


class PendingValidationView(ctk.CTkFrame):
    """Painel executivo exibido para os módulos que estão em validação pela equipe."""

    def __init__(self, master, module_cfg: dict):
        super().__init__(master, fg_color="transparent")
        self.module_cfg = module_cfg
        self.setup_ui()

    def setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        card = ctk.CTkFrame(self, corner_radius=14, fg_color=CARD_BG)
        card.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        card.grid_columnconfigure(0, weight=1)

        # Badge de Status
        badge = ctk.CTkFrame(card, fg_color="#78350F", corner_radius=16)
        badge.pack(pady=(32, 12))
        ctk.CTkLabel(
            badge,
            text=f"  {self.module_cfg['status_badge']} • AGUARDANDO HOMOLOGAÇÃO FINAL DA EQUIPE  ",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#FDE68A",
        ).pack(padx=14, pady=6)

        # Título do Módulo
        ctk.CTkLabel(
            card,
            text=f"Módulo {self.module_cfg['code']}: {self.module_cfg['full_title']}",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#FFFFFF",
        ).pack(pady=(4, 8), padx=25)

        # Status Label
        ctk.CTkLabel(
            card,
            text=self.module_cfg["status_label"],
            font=ctk.CTkFont(size=13, slant="italic"),
            text_color="#FBBF24",
        ).pack(pady=(0, 18), padx=25)

        # Escopo Técnico
        desc_frame = ctk.CTkFrame(card, corner_radius=10, fg_color="#0F172A")
        desc_frame.pack(fill="x", padx=40, pady=(0, 16))

        ctk.CTkLabel(
            desc_frame,
            text="📌 Escopo da Auditoria Automatizada:",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#38BDF8",
            anchor="w",
        ).pack(fill="x", padx=18, pady=(12, 4))

        ctk.CTkLabel(
            desc_frame,
            text=self.module_cfg["description"],
            font=ctk.CTkFont(size=12),
            text_color="#E2E8F0",
            justify="left",
            wraplength=760,
            anchor="w",
        ).pack(fill="x", padx=18, pady=(0, 12))

        # Grid de Eventos e Bases Requeridas
        info_grid = ctk.CTkFrame(card, fg_color="transparent")
        info_grid.pack(fill="x", padx=40, pady=(0, 20))
        info_grid.grid_columnconfigure(0, weight=1)
        info_grid.grid_columnconfigure(1, weight=1)

        ev_box = ctk.CTkFrame(info_grid, corner_radius=10, fg_color="#0F172A")
        ev_box.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ctk.CTkLabel(
            ev_box,
            text="📊 Eventos / Escopo Auditado:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#93C5FD",
            anchor="w",
        ).pack(fill="x", padx=15, pady=(10, 6))
        for ev in self.module_cfg.get("events", []):
            ctk.CTkLabel(
                ev_box,
                text=f"  • {ev}",
                font=ctk.CTkFont(size=11),
                text_color="#CBD5E1",
                anchor="w",
            ).pack(fill="x", padx=15, pady=2)
        ctk.CTkLabel(ev_box, text="", height=6).pack()

        req_box = ctk.CTkFrame(info_grid, corner_radius=10, fg_color="#0F172A")
        req_box.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        ctk.CTkLabel(
            req_box,
            text="📁 Bases de Entrada Mapeadas:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#93C5FD",
            anchor="w",
        ).pack(fill="x", padx=15, pady=(10, 6))
        for rf in self.module_cfg.get("required_files", []):
            ctk.CTkLabel(
                req_box,
                text=f"  • {rf}",
                font=ctk.CTkFont(size=11),
                text_color="#CBD5E1",
                anchor="w",
            ).pack(fill="x", padx=15, pady=2)
        ctk.CTkLabel(req_box, text="", height=6).pack()

        # Botão Informativo Desabilitado
        btn_locked = ctk.CTkButton(
            card,
            text="🔒 Módulo em Validação pela Equipe — Liberação Automática após Aprovação",
            height=40,
            width=540,
            state="disabled",
            fg_color="#334155",
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        btn_locked.pack(pady=(10, 24))


class CentralBeneficiosApp(ctk.CTk):
    """Portal Executivo Unificado para todas as Conferências de Benefícios (VA, VT e Convênios)."""

    def __init__(self):
        super().__init__()
        self.title("SISTEMAS RH CENCOSUD - Central Executiva de Auditoria de Benefícios (VA e VT)")
        self.geometry("1220x800")
        self.minsize(1080, 720)

        # Configurar Ícone da Janela
        ico_path = resource_path(os.path.join("assets", "cencosud.ico"))
        if os.path.exists(ico_path):
            try:
                self.iconbitmap(ico_path)
            except Exception:
                pass

        # Carregar Logo Oficial da Cencosud
        logo_path = resource_path(os.path.join("assets", "cencosud_logo_cropped.png"))
        fallback_path = r"C:\Users\matcaval\OneDrive - Cencosud\Imagens\cencosud-logo-png_seeklogo-251998.png"

        self.logo_img = None
        if HAS_PIL and Image is not None:
            for p in [logo_path, fallback_path]:
                if os.path.exists(p):
                    try:
                        pil_img = Image.open(p).convert("RGBA")
                        w, h = pil_img.size
                        target_h = 42
                        target_w = int(w * (target_h / h))
                        self.logo_img = ctk.CTkImage(
                            light_image=pil_img, dark_image=pil_img, size=(target_w, target_h)
                        )
                        break
                    except Exception:
                        pass

        # Diretórios padrão inteligentes para o Good Card
        default_root = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\04 - GOOD CARD\Relatorio Original"
        default_output = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\04 - GOOD CARD\Relatório Tratado"

        app_dir = (
            os.path.dirname(sys.executable)
            if getattr(sys, "frozen", False)
            else os.path.dirname(os.path.abspath(__file__))
        )
        if not os.path.exists(default_root):
            cand_rel = os.path.join(app_dir, "Relatorio Original")
            cand_gc = os.path.join(app_dir, "04 - GOOD CARD", "Relatorio Original")
            if os.path.exists(cand_rel):
                default_root = cand_rel
            elif os.path.exists(cand_gc):
                default_root = cand_gc
            else:
                default_root = app_dir

        if not os.path.exists(default_output):
            cand_out = os.path.join(app_dir, "Relatório Tratado")
            cand_gc_out = os.path.join(app_dir, "04 - GOOD CARD", "Relatório Tratado")
            if os.path.exists(cand_gc_out):
                default_output = cand_gc_out
            else:
                default_output = cand_out

        self.root_var = tk.StringVar(value=default_root)
        self.output_var = tk.StringVar(value=default_output)
        self.month_var = tk.StringVar(value="09")
        self.year_var = tk.StringVar(value="2026")
        self.diag_result = {}

        self.sidebar_buttons = {}
        self.module_views = {}
        self.current_module_id = None

        self.setup_ui()
        self.select_module("good_card")
        self.rescan_files()

    def setup_ui(self):
        # -------------------------------------------------------------
        # 1. CABEÇALHO CORPORATIVO CENCOSUD (#095AA5)
        # -------------------------------------------------------------
        header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=CENCOSUD_BLUE, height=76)
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)

        header_inner = ctk.CTkFrame(header_frame, fg_color="transparent")
        header_inner.pack(fill="both", expand=True, padx=20, pady=10)

        left_box = ctk.CTkFrame(header_inner, fg_color="transparent")
        left_box.pack(side="left", fill="y", anchor="w")

        if self.logo_img:
            logo_card = ctk.CTkFrame(left_box, fg_color="#FFFFFF", corner_radius=8)
            logo_card.pack(side="left", padx=(0, 16), pady=2)
            lbl_logo = ctk.CTkLabel(logo_card, text="", image=self.logo_img)
            lbl_logo.pack(padx=10, pady=4)

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
            text="Central Executiva de Auditoria Automatizada de Benefícios (Good Card, VT, Refeitório e Convênios)",
            font=ctk.CTkFont(size=12),
            text_color="#D8E8F8",
        ).pack(anchor="w", pady=(0, 2))

        right_box = ctk.CTkFrame(header_inner, fg_color="transparent")
        right_box.pack(side="right", fill="y", anchor="e")

        badge = ctk.CTkFrame(right_box, fg_color="#07447D", corner_radius=15)
        badge.pack(anchor="e", pady=10)
        ctk.CTkLabel(
            badge,
            text=" PORTAL BENEFÍCIOS • V1.0 OFICIAL ",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#BCE0FD",
        ).pack(padx=12, pady=5)

        # -------------------------------------------------------------
        # 2. CORPO PRINCIPAL: SIDEBAR ESQUERDA + CONTAINER DE MÓDULOS
        # -------------------------------------------------------------
        body_frame = ctk.CTkFrame(self, fg_color="transparent")
        body_frame.pack(fill="both", expand=True, padx=14, pady=12)
        body_frame.grid_columnconfigure(0, weight=0)
        body_frame.grid_columnconfigure(1, weight=1)
        body_frame.grid_rowconfigure(0, weight=1)

        # Sidebar de Módulos
        sidebar = ctk.CTkFrame(body_frame, width=265, corner_radius=10, fg_color=SIDEBAR_BG)
        sidebar.grid(row=0, column=0, sticky="nsew", padx=(0, 12), pady=0)
        sidebar.grid_propagate(False)

        ctk.CTkLabel(
            sidebar,
            text="MÓDULOS DE AUDITORIA",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94A3B8",
            anchor="w",
        ).pack(fill="x", padx=16, pady=(14, 6))

        for mod in MODULES_CONFIG:
            mod_id = mod["id"]
            is_enabled = mod["enabled"]
            dot = "🟢" if is_enabled else "🟡"
            btn_text = f"{dot}  {mod['code']}. {mod['short_name']}"

            btn = ctk.CTkButton(
                sidebar,
                text=btn_text,
                anchor="w",
                height=36,
                corner_radius=8,
                font=ctk.CTkFont(size=12, weight="bold" if is_enabled else "normal"),
                fg_color="transparent",
                hover_color="#1E293B",
                text_color="#FFFFFF" if is_enabled else "#94A3B8",
                command=lambda mid=mod_id: self.select_module(mid),
            )
            btn.pack(fill="x", padx=10, pady=3)
            self.sidebar_buttons[mod_id] = btn

        # Rodapé da Sidebar com Legenda de Status
        legend_frame = ctk.CTkFrame(sidebar, corner_radius=8, fg_color="#1E293B")
        legend_frame.pack(side="bottom", fill="x", padx=10, pady=12)
        ctk.CTkLabel(
            legend_frame,
            text="🟢 Homologado e Liberado\n🟡 Em Validação pela Equipe",
            font=ctk.CTkFont(size=10),
            text_color="#CBD5E1",
            justify="left",
        ).pack(padx=10, pady=8, anchor="w")

        # Container Principal à Direita
        self.content_container = ctk.CTkFrame(body_frame, fg_color="transparent")
        self.content_container.grid(row=0, column=1, sticky="nsew")
        self.content_container.grid_columnconfigure(0, weight=1)
        self.content_container.grid_rowconfigure(0, weight=1)

        # Instanciar Views para cada módulo
        for mod in MODULES_CONFIG:
            mod_id = mod["id"]
            if mod["enabled"] and mod_id in ENGINES_MAP:
                view = ActiveModuleView(
                    self.content_container,
                    module_cfg=mod,
                    engine_fn=ENGINES_MAP[mod_id],
                    app_instance=self,
                )
            else:
                view = PendingValidationView(self.content_container, module_cfg=mod)
            self.module_views[mod_id] = view

    def select_module(self, mod_id: str):
        if self.current_module_id and self.current_module_id in self.module_views:
            self.module_views[self.current_module_id].grid_forget()

        self.current_module_id = mod_id
        self.module_views[mod_id].grid(row=0, column=0, sticky="nsew")

        mod_cfg = next((m for m in MODULES_CONFIG if m["id"] == mod_id), None)
        if mod_cfg:
            va_vt_base = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT")
            candidates = [
                va_vt_base,
                Path.cwd(),
                Path.cwd().parent,
                Path(self.root_var.get()).parent if self.root_var.get() else Path.cwd(),
            ]
            for base in candidates:
                cand_in = base / mod_cfg.get("default_subfolder", "")
                if cand_in.exists():
                    self.root_var.set(str(cand_in))
                    cand_out = base / mod_cfg.get("default_out_subfolder", "")
                    self.output_var.set(str(cand_out))
                    break

        for mid, btn in self.sidebar_buttons.items():
            if mid == mod_id:
                btn.configure(fg_color=CENCOSUD_BLUE, hover_color=CENCOSUD_BLUE_HOVER, text_color="#FFFFFF")
            else:
                is_en = next((m["enabled"] for m in MODULES_CONFIG if m["id"] == mid), False)
                btn.configure(
                    fg_color="transparent",
                    hover_color="#1E293B",
                    text_color="#FFFFFF" if is_en else "#94A3B8",
                )

        self.rescan_files()

    def get_output_dir(self) -> str:
        return os.path.abspath(self.output_var.get())

    def get_root_dir(self) -> str:
        return os.path.abspath(self.root_var.get())

    def browse_root_folder(self):
        curr = self.root_var.get()
        chosen = filedialog.askdirectory(
            initialdir=curr if os.path.exists(curr) else os.getcwd(),
            title="Selecione a Pasta dos Relatórios Originais",
        )
        if chosen:
            self.root_var.set(chosen)
            self.rescan_files()

    def browse_output_folder(self):
        curr = self.output_var.get()
        chosen = filedialog.askdirectory(
            initialdir=curr if os.path.exists(curr) else os.getcwd(),
            title="Selecione a Pasta de Destino da Conferência",
        )
        if chosen:
            self.output_var.set(chosen)

    def rescan_files(self):
        root_dir = self.get_root_dir()
        out_dir = self.get_output_dir()

        detector = FileDetector(root_dir, out_dir)
        self.diag_result = detector.scan_module(self.current_module_id or "good_card")

        det_month = self.diag_result.get("month")
        det_year = self.diag_result.get("year")
        if det_month:
            self.month_var.set(f"{int(det_month):02d}")
        if det_year:
            self.year_var.set(str(det_year))

        active_view = self.module_views.get(self.current_module_id)
        if isinstance(active_view, ActiveModuleView):
            active_view.update_checklist(self.diag_result)


if __name__ == "__main__":
    app = CentralBeneficiosApp()
    app.mainloop()
