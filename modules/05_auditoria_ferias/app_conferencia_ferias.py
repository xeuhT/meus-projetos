import os
import sys
import threading
import time
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image

from file_detector import FileDetector
from audit_engine import (
    exec_conferencia_formatada,
    exec_auditoria_folha,
    exec_auditoria_calculo
)

# Helper para recursos internos do executável PyInstaller
def resource_path(relative_path):
    """Obtém o caminho absoluto para recursos, funcionando em ambiente de desenvolvimento e em bundle do PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# Configurações de Aparência Corporativa CustomTkinter
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# Cores Oficiais Cencosud
CENCOSUD_BLUE = "#095AA5"
CENCOSUD_BLUE_HOVER = "#0B6FCB"
CENCOSUD_DARK_BLUE = "#063A6B"
CENCOSUD_LIGHT_BG = "#EBF4FB"
SUCCESS_GREEN = "#1E7B34"
SUCCESS_GREEN_HOVER = "#28A745"

class FileLockModal(ctk.CTkToplevel):
    """Janela modal amigável exibida quando o Excel trava a planilha final."""
    def __init__(self, parent, file_path):
        super().__init__(parent)
        self.parent = parent
        self.file_path = file_path
        self.result = 'cancel'
        self.event = threading.Event()

        self.title("Aviso: Arquivo Aberto no Excel")
        self.geometry("520x240")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        # Layout
        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.pack(fill="both", expand=True, padx=15, pady=15)

        lbl_icon = ctk.CTkLabel(frame, text="⚠️", font=ctk.CTkFont(size=36))
        lbl_icon.pack(pady=(10, 0))

        lbl_title = ctk.CTkLabel(
            frame, 
            text="Planilha em Uso no Excel", 
            font=ctk.CTkFont(size=16, weight="bold")
        )
        lbl_title.pack(pady=(5, 5))

        fname = os.path.basename(file_path)
        lbl_msg = ctk.CTkLabel(
            frame,
            text=f"A planilha '{fname}' está aberta no Excel ou por outro usuário.\nPor favor, feche o arquivo para prosseguir com a gravação.",
            font=ctk.CTkFont(size=12),
            wraplength=460,
            justify="center"
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
            command=self.on_retry
        )
        btn_retry.pack(side="left", expand=True, padx=5)

        btn_copy = ctk.CTkButton(
            btn_box,
            text="Salvar como Cópia",
            fg_color=SUCCESS_GREEN,
            hover_color=SUCCESS_GREEN_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self.on_copy
        )
        btn_copy.pack(side="left", expand=True, padx=5)

        btn_cancel = ctk.CTkButton(
            btn_box,
            text="Cancelar",
            fg_color="#555555",
            hover_color="#777777",
            command=self.on_cancel
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
        self.result = 'retry'
        self.event.set()
        self.destroy()

    def on_copy(self):
        self.result = 'copy'
        self.event.set()
        self.destroy()

    def on_cancel(self):
        self.result = 'cancel'
        self.event.set()
        self.destroy()


class AuditTab(ctk.CTkFrame):
    """Componente modular para cada uma das 3 abas de conferência."""
    def __init__(self, master, tab_id, title_text, btn_text, engine_fn, app_instance):
        super().__init__(master, fg_color="transparent")
        self.tab_id = tab_id
        self.title_text = title_text
        self.btn_text = btn_text
        self.engine_fn = engine_fn
        self.app = app_instance
        self.generated_file = None

        self.setup_ui()

    def setup_ui(self):
        # Container com 2 colunas: Esquerda (Checklist) e Direita (Ações, Progresso e Logs)
        self.grid_columnconfigure(0, weight=4) # Checklist
        self.grid_columnconfigure(1, weight=5) # Execução e Logs
        self.grid_rowconfigure(0, weight=1)

        # -------------------------------------------------------------
        # COLUNA ESQUERDA: DIAGNÓSTICO E CHECKLIST DE ARQUIVOS
        # -------------------------------------------------------------
        frame_left = ctk.CTkFrame(self, corner_radius=10)
        frame_left.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)
        frame_left.grid_columnconfigure(0, weight=1)
        frame_left.grid_rowconfigure(1, weight=1)

        lbl_chk_header = ctk.CTkLabel(
            frame_left,
            text=f"📋 Diagnóstico de Entradas ({self.title_text})",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        lbl_chk_header.grid(row=0, column=0, sticky="w", padx=15, pady=(12, 8))

        # Scrollable Frame para o checklist de arquivos
        self.scroll_checklist = ctk.CTkScrollableFrame(frame_left, fg_color="transparent")
        self.scroll_checklist.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.scroll_checklist.grid_columnconfigure(0, weight=1)

        # -------------------------------------------------------------
        # COLUNA DIREITA: EXECUÇÃO, PROGRESSO E CONSOLE DE LOGS
        # -------------------------------------------------------------
        frame_right = ctk.CTkFrame(self, corner_radius=10)
        frame_right.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=0)
        frame_right.grid_columnconfigure(0, weight=1)
        frame_right.grid_rowconfigure(3, weight=1)

        # Barra de Ação Principal
        frame_action = ctk.CTkFrame(frame_right, fg_color="transparent")
        frame_action.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 5))
        frame_action.grid_columnconfigure(0, weight=1)

        self.btn_run = ctk.CTkButton(
            frame_action,
            text=self.btn_text,
            height=42,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=CENCOSUD_BLUE,
            hover_color=CENCOSUD_BLUE_HOVER,
            command=self.start_execution
        )
        self.btn_run.grid(row=0, column=0, sticky="ew")

        # Barra de Progresso Real
        self.lbl_progress_status = ctk.CTkLabel(
            frame_right,
            text="Status: Pronto para processar.",
            font=ctk.CTkFont(size=12, slant="italic"),
            anchor="w"
        )
        self.lbl_progress_status.grid(row=1, column=0, sticky="ew", padx=15, pady=(5, 2))

        self.progress_bar = ctk.CTkProgressBar(frame_right, height=12)
        self.progress_bar.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 10))
        self.progress_bar.set(0.0)

        # Console de Logs Integrado
        self.txt_log = ctk.CTkTextbox(
            frame_right,
            font=ctk.CTkFont(family="Consolas", size=11),
            corner_radius=8,
            wrap="word"
        )
        self.txt_log.grid(row=3, column=0, sticky="nsew", padx=15, pady=5)
        self.append_log(f"Aba inicializada. Aguardando comando para {self.title_text}.")

        # Ações Pós-Processamento (Abrir Arquivo e Abrir Pasta)
        frame_post = ctk.CTkFrame(frame_right, fg_color="transparent")
        frame_post.grid(row=4, column=0, sticky="ew", padx=15, pady=(10, 15))
        frame_post.grid_columnconfigure(0, weight=1)
        frame_post.grid_columnconfigure(1, weight=1)

        self.btn_open_file = ctk.CTkButton(
            frame_post,
            text="📄 Abrir Planilha Gerada",
            font=ctk.CTkFont(weight="bold"),
            fg_color=SUCCESS_GREEN,
            hover_color=SUCCESS_GREEN_HOVER,
            state="disabled",
            command=self.open_generated_file
        )
        self.btn_open_file.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.btn_open_dir = ctk.CTkButton(
            frame_post,
            text="📁 Abrir Pasta de Destino",
            font=ctk.CTkFont(weight="bold"),
            fg_color="#333333",
            hover_color="#4F4F4F",
            command=self.open_output_folder
        )
        self.btn_open_dir.grid(row=0, column=1, sticky="ew", padx=(5, 0))

    def append_log(self, msg):
        """Adiciona uma mensagem com carimbo de hora no console de logs."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", f"[{timestamp}] {msg}\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def update_checklist(self, tab_data):
        """Atualiza visualmente os cartões do checklist com base na varredura."""
        for w in self.scroll_checklist.winfo_children():
            w.destroy()

        items = tab_data.get('items', [])
        ready = tab_data.get('ready', False)

        for i, item in enumerate(items):
            file_info = item.get('file')
            status = item.get('status', 'OK' if file_info else 'MISSING')
            label = item.get('label', '')
            obs = item.get('obs', '')

            # Card individual
            card = ctk.CTkFrame(self.scroll_checklist, corner_radius=8, fg_color="#1E293B")
            card.pack(fill="x", pady=4, padx=2)
            card.grid_columnconfigure(1, weight=1)

            # Ícone e Cor de Status
            if file_info is not None:
                icon_text = "🟢"
                title_color = "#FFFFFF"
                detail_color = "#34D399" # Verde esmeralda corporativo
                fname = file_info.get('name', '')
                detail_text = f"{fname}"
                if obs: detail_text += f" • {obs}"
            elif not item.get('req', True):
                icon_text = "🟡"
                title_color = "#FFFFFF"
                detail_color = "#FBBF24"
                detail_text = obs or "Arquivo opcional (gerado dinamicamente se ausente)"
            else:
                icon_text = "🔴"
                title_color = "#FCA5A5"
                detail_color = "#F87171" # Vermelho somente se faltar o arquivo
                detail_text = "Arquivo obrigatório não localizado na pasta"

            lbl_ico = ctk.CTkLabel(card, text=icon_text, font=ctk.CTkFont(size=14))
            lbl_ico.grid(row=0, column=0, rowspan=2, padx=(10, 8), pady=8)

            lbl_name = ctk.CTkLabel(
                card, 
                text=label, 
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=title_color,
                anchor="w"
            )
            lbl_name.grid(row=0, column=1, sticky="w", padx=0, pady=(6, 0))

            lbl_detail = ctk.CTkLabel(
                card,
                text=detail_text,
                font=ctk.CTkFont(size=11),
                text_color=detail_color,
                anchor="w"
            )
            lbl_detail.grid(row=1, column=1, sticky="w", padx=0, pady=(0, 6))

        # Habilita ou desabilita o botão de execução com base no checklist
        if ready:
            self.btn_run.configure(
                state="normal", 
                fg_color=CENCOSUD_BLUE,
                text=self.btn_text
            )
            self.lbl_progress_status.configure(text="Pronto: Todos os arquivos obrigatórios foram identificados.")
        else:
            self.btn_run.configure(
                state="disabled", 
                fg_color="#444444",
                text="Aguardando Arquivos Obrigatórios (Ver Checklist ao Lado)"
            )
            self.lbl_progress_status.configure(text="Atenção: Localize os arquivos em vermelho para liberar a execução.")

    def start_execution(self):
        """Inicia a execução do motor em uma thread assíncrona dedicada."""
        self.btn_run.configure(state="disabled")
        self.btn_open_file.configure(state="disabled")
        self.progress_bar.set(0.02)
        self.lbl_progress_status.configure(text="Iniciando processamento em segundo plano...")
        self.append_log(f">>> Disparando motor de cálculo para: {self.title_text}")

        raw_map = self.app.diag_result.get('raw_map', {})
        output_dir = self.app.get_output_dir()

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
                    raw_map,
                    output_dir,
                    progress_cb=progress_cb,
                    log_cb=log_cb,
                    prompt_lock_cb=prompt_lock_cb
                )
                self.app.after(0, lambda: self._on_success(saved))
            except Exception as e:
                import traceback
                err = traceback.format_exc()
                self.app.after(0, lambda: self._on_error(str(e), err))

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

    def _update_progress(self, pct, text):
        self.progress_bar.set(pct)
        self.lbl_progress_status.configure(text=f"{int(pct*100)}% - {text}")

    def _on_success(self, file_path):
        self.generated_file = file_path
        self.btn_run.configure(state="normal")
        self.btn_open_file.configure(state="normal")
        self.progress_bar.set(1.0)
        self.lbl_progress_status.configure(text="Concluído com sucesso!")
        self.append_log(f"Processo finalizado com sucesso. Arquivo disponível: {os.path.basename(file_path)}")
        messagebox.showinfo("Sucesso", f"Conferência gerada com sucesso!\n\nArquivo: {os.path.basename(file_path)}")

    def _on_error(self, err_msg, full_trace):
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


class ConferenciaFeriasApp(ctk.CTk):
    """Aplicativo Executivo Unificado para Auditoria de Férias e Folha Cencosud."""
    def __init__(self):
        super().__init__()
        self.title("SISTEMAS RH CENCOSUD - Central de Auditoria de Férias e Folha")
        self.geometry("1140x780")
        self.minsize(1000, 700)

        # Configurar Ícone da Janela
        ico_path = resource_path(os.path.join("assets", "cencosud.ico"))
        if os.path.exists(ico_path):
            try:
                self.iconbitmap(ico_path)
            except Exception:
                pass

        # Carregar Logo da Cencosud
        logo_path = resource_path(os.path.join("assets", "cencosud_logo_cropped.png"))
        fallback_path = r"C:\Users\matcaval\OneDrive - Cencosud\Imagens\cencosud-logo-png_seeklogo-251998.png"
        
        self.logo_img = None
        for p in [logo_path, fallback_path]:
            if os.path.exists(p):
                try:
                    pil_img = Image.open(p).convert("RGBA")
                    # Mantém proporção e define tamanho elegante para o banner
                    w, h = pil_img.size
                    target_h = 42
                    target_w = int(w * (target_h / h))
                    self.logo_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(target_w, target_h))
                    break
                except Exception as ex:
                    print(f"Erro ao carregar logo de {p}: {ex}")

        # Caminhos padrão do ambiente corporativo
        default_root = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\férias\planilhas"
        default_output = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\férias\conferencias"

        app_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        if not os.path.exists(default_root):
            default_root = os.path.abspath(os.path.join(app_dir, "planilhas"))
        if not os.path.exists(default_output):
            default_output = os.path.abspath(os.path.join(app_dir, "conferencias"))

        self.root_var = tk.StringVar(value=default_root)
        self.output_var = tk.StringVar(value=default_output)
        self.diag_result = {}

        self.setup_ui()
        self.rescan_files()

    def setup_ui(self):
        # -------------------------------------------------------------
        # CABEÇALHO CORPORATIVO COM AS CORES DA LOGO CENCOSUD (#095AA5)
        # -------------------------------------------------------------
        header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=CENCOSUD_BLUE, height=76)
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)

        header_inner = ctk.CTkFrame(header_frame, fg_color="transparent")
        header_inner.pack(fill="both", expand=True, padx=20, pady=10)

        # Lado Esquerdo: Logo em Card Branco e Título
        left_box = ctk.CTkFrame(header_inner, fg_color="transparent")
        left_box.pack(side="left", fill="y", anchor="w")

        if self.logo_img:
            logo_card = ctk.CTkFrame(left_box, fg_color="#FFFFFF", corner_radius=8)
            logo_card.pack(side="left", padx=(0, 16), pady=2)
            lbl_logo = ctk.CTkLabel(logo_card, text="", image=self.logo_img)
            lbl_logo.pack(padx=10, pady=4)

        title_box = ctk.CTkFrame(left_box, fg_color="transparent")
        title_box.pack(side="left", fill="y", anchor="w")

        lbl_app_title = ctk.CTkLabel(
            title_box,
            text="CENCOSUD BRASIL • SISTEMAS RH",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#FFFFFF"
        )
        lbl_app_title.pack(anchor="w", pady=(2, 0))

        lbl_app_subtitle = ctk.CTkLabel(
            title_box,
            text="Central Executiva de Auditoria Automatizada: Pré-Cálculo, Férias e Folha Mensal",
            font=ctk.CTkFont(size=12),
            text_color="#D8E8F8"
        )
        lbl_app_subtitle.pack(anchor="w", pady=(0, 2))

        # Lado Direito: Badge Institucional
        right_box = ctk.CTkFrame(header_inner, fg_color="transparent")
        right_box.pack(side="right", fill="y", anchor="e")

        badge = ctk.CTkFrame(right_box, fg_color="#07447D", corner_radius=15)
        badge.pack(anchor="e", pady=10)
        lbl_badge = ctk.CTkLabel(
            badge,
            text=" SISTEMAS RH • V1.0 OFICIAL ",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#BCE0FD"
        )
        lbl_badge.pack(padx=12, pady=5)

        # -------------------------------------------------------------
        # BARRA DE CONFIGURAÇÃO DE DIRETÓRIOS
        # -------------------------------------------------------------
        config_frame = ctk.CTkFrame(self, corner_radius=10)
        config_frame.pack(fill="x", padx=15, pady=(12, 8))
        config_frame.grid_columnconfigure(1, weight=1)

        # Pasta Raiz das Planilhas
        lbl_root = ctk.CTkLabel(config_frame, text="Pasta Raiz das Planilhas:", font=ctk.CTkFont(weight="bold"))
        lbl_root.grid(row=0, column=0, padx=(15, 10), pady=(10, 4), sticky="w")

        entry_root = ctk.CTkEntry(config_frame, textvariable=self.root_var, height=32)
        entry_root.grid(row=0, column=1, padx=(0, 10), pady=(10, 4), sticky="ew")

        btn_browse_root = ctk.CTkButton(
            config_frame,
            text="Procurar...",
            width=90,
            height=32,
            command=self.browse_root_folder
        )
        btn_browse_root.grid(row=0, column=2, padx=(0, 15), pady=(10, 4))

        # Pasta de Destino
        lbl_out = ctk.CTkLabel(config_frame, text="Destino das Conferências:", font=ctk.CTkFont(weight="bold"))
        lbl_out.grid(row=1, column=0, padx=(15, 10), pady=(4, 10), sticky="w")

        entry_out = ctk.CTkEntry(config_frame, textvariable=self.output_var, height=32)
        entry_out.grid(row=1, column=1, padx=(0, 10), pady=(4, 10), sticky="ew")

        btn_browse_out = ctk.CTkButton(
            config_frame,
            text="Procurar...",
            width=90,
            height=32,
            command=self.browse_output_folder
        )
        btn_browse_out.grid(row=1, column=2, padx=(0, 15), pady=(4, 10))

        # Botão de Varredura / Atualização
        self.btn_scan = ctk.CTkButton(
            config_frame,
            text="🔄 Escanear / Atualizar Arquivos",
            height=36,
            font=ctk.CTkFont(weight="bold"),
            fg_color=CENCOSUD_BLUE,
            hover_color=CENCOSUD_BLUE_HOVER,
            command=self.rescan_files
        )
        self.btn_scan.grid(row=0, column=3, rowspan=2, padx=(5, 15), pady=10, sticky="ns")

        # -------------------------------------------------------------
        # NAVEGADOR COM ABAS (TABVIEW)
        # -------------------------------------------------------------
        self.tabview = ctk.CTkTabview(
            self, 
            corner_radius=10,
            segmented_button_selected_color=CENCOSUD_BLUE,
            segmented_button_selected_hover_color=CENCOSUD_BLUE_HOVER
        )
        self.tabview.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        tab_names = [
            "Automação Conferência Férias Formatada (Pré-Cálculo)",
            "Auditoria do Cálculo de Férias (85 Colunas)",
            "Auditoria da Folha Mensal (65 Colunas)"
        ]

        for name in tab_names:
            self.tabview.add(name)

        # Aba 1: Conferência Férias Formatada
        self.tab_aba1 = AuditTab(
            self.tabview.tab(tab_names[0]),
            tab_id='aba1',
            title_text="Conferência Férias Formatada",
            btn_text="▶ Gerar Automação Conferência Formatada",
            engine_fn=exec_conferencia_formatada,
            app_instance=self
        )
        self.tab_aba1.pack(fill="both", expand=True, padx=5, pady=5)

        # Aba 2: Auditoria Cálculo de Férias
        self.tab_aba2 = AuditTab(
            self.tabview.tab(tab_names[1]),
            tab_id='aba2',
            title_text="Cálculo de Férias",
            btn_text="▶ Gerar Auditoria do Cálculo de Férias",
            engine_fn=exec_auditoria_calculo,
            app_instance=self
        )
        self.tab_aba2.pack(fill="both", expand=True, padx=5, pady=5)

        # Aba 3: Auditoria Folha Mensal
        self.tab_aba3 = AuditTab(
            self.tabview.tab(tab_names[2]),
            tab_id='aba3',
            title_text="Folha Mensal",
            btn_text="▶ Gerar Auditoria da Folha Mensal",
            engine_fn=exec_auditoria_folha,
            app_instance=self
        )
        self.tab_aba3.pack(fill="both", expand=True, padx=5, pady=5)

    def get_output_dir(self):
        return os.path.abspath(self.output_var.get())

    def get_root_dir(self):
        return os.path.abspath(self.root_var.get())

    def browse_root_folder(self):
        curr = self.root_var.get()
        chosen = filedialog.askdirectory(initialdir=curr if os.path.exists(curr) else os.getcwd(), title="Selecione a Pasta Raiz das Planilhas")
        if chosen:
            self.root_var.set(chosen)
            self.rescan_files()

    def browse_output_folder(self):
        curr = self.output_var.get()
        chosen = filedialog.askdirectory(initialdir=curr if os.path.exists(curr) else os.getcwd(), title="Selecione a Pasta de Destino das Conferências")
        if chosen:
            self.output_var.set(chosen)

    def rescan_files(self):
        """Dispara a varredura e atualiza todas as 3 abas."""
        root_dir = self.get_root_dir()
        out_dir = self.get_output_dir()

        self.btn_scan.configure(state="disabled", text="Varrendo...")
        
        detector = FileDetector(root_dir, out_dir)
        self.diag_result = detector.scan()

        self.tab_aba1.update_checklist(self.diag_result.get('aba1', {}))
        self.tab_aba2.update_checklist(self.diag_result.get('aba2', {}))
        self.tab_aba3.update_checklist(self.diag_result.get('aba3', {}))

        self.btn_scan.configure(state="normal", text="🔄 Escanear / Atualizar Arquivos")

if __name__ == '__main__':
    app = ConferenciaFeriasApp()
    app.mainloop()
