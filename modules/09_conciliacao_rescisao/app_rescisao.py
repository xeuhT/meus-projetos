import os
import sys
import threading
from datetime import datetime
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image

# Importar o motor analítico de auditoria
from auditor_massivo_rescisao import auditar_folha_massiva

def resource_path(relative_path):
    """Obtém caminho absoluto para recursos, funcionando em dev e com PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)

class AuditorRescisaoApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configurações da Janela
        self.title("Cencosud Brasil - Auditor Analítico de Rescisão")
        self.geometry("820x760")
        self.minsize(780, 700)
        
        # Ícone da janela se disponível
        icon_path = resource_path("app_icon.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        # Tema Corporativo Cencosud
        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("blue")

        self.cencosud_blue = "#004B87"
        self.cencosud_hover = "#003560"
        self.cencosud_orange = "#F37021"
        self.cencosud_bg = "#F8FAFC"
        self.card_bg = "#FFFFFF"

        self.configure(fg_color=self.cencosud_bg)

        # Variáveis de Estado
        self.caminho_pasta = ctk.StringVar(value="")
        self.arquivos_detectados = []
        self.is_running = False

        self._criar_interface()

    def _criar_interface(self):
        # 1. HEADER CORPORATIVO
        header_frame = ctk.CTkFrame(self, fg_color=self.cencosud_blue, corner_radius=0, height=95)
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)

        header_content = ctk.CTkFrame(header_frame, fg_color="transparent")
        header_content.pack(fill="both", expand=True, padx=25, pady=12)

        # Logo
        logo_path = resource_path("logo_cropped.png")
        if not os.path.exists(logo_path):
            logo_path = resource_path("logo_cencosud.png")
            
        if os.path.exists(logo_path):
            try:
                img_pil = Image.open(logo_path)
                logo_ctk = ctk.CTkImage(light_image=img_pil, dark_image=img_pil, size=(110, 60))
                lbl_logo = ctk.CTkLabel(header_content, image=logo_ctk, text="")
                lbl_logo.pack(side="left", padx=(0, 20))
            except Exception:
                pass

        # Títulos do Header
        title_box = ctk.CTkFrame(header_content, fg_color="transparent")
        title_box.pack(side="left", fill="y", expand=True)

        lbl_title = ctk.CTkLabel(
            title_box, 
            text="Auditor de Rescisão & FOPAG", 
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color="#FFFFFF"
        )
        lbl_title.pack(anchor="w")

        lbl_sub = ctk.CTkLabel(
            title_box, 
            text="Motor Analítico de Conformidade CLT e Conciliação LG (AD003)", 
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#CBD5E1"
        )
        lbl_sub.pack(anchor="w")

        # Badge de Versão
        badge = ctk.CTkLabel(
            header_content,
            text="Versão V005 (Painel & Governança)",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color="#FFFFFF",
            fg_color="#002D54",
            corner_radius=12,
            padx=12,
            pady=4
        )
        badge.pack(side="right", padx=5)

        # CONTAINER PRINCIPAL ROLÁVEL OU FIXO
        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=25, pady=15)

        # 2. CARD: SELEÇÃO DA PASTA
        card_pasta = ctk.CTkFrame(main_container, fg_color=self.card_bg, corner_radius=10, border_width=1, border_color="#E2E8F0")
        card_pasta.pack(fill="x", pady=(0, 12), ipady=8)

        lbl_step1 = ctk.CTkLabel(
            card_pasta, 
            text="1. Seleção da Pasta de Trabalho", 
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color="#1E293B"
        )
        lbl_step1.pack(anchor="w", padx=20, pady=(10, 4))

        lbl_step1_desc = ctk.CTkLabel(
            card_pasta, 
            text="Selecione a pasta onde estão os arquivos de rescisão (.csv ou .xlsb):", 
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#64748B"
        )
        lbl_step1_desc.pack(anchor="w", padx=20, pady=(0, 10))

        # Linha do Campo + Botão
        folder_row = ctk.CTkFrame(card_pasta, fg_color="transparent")
        folder_row.pack(fill="x", padx=20, pady=(0, 10))

        self.entry_pasta = ctk.CTkEntry(
            folder_row, 
            textvariable=self.caminho_pasta, 
            placeholder_text="Clique em 'Selecionar Pasta' para escolher o diretório...",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            height=38,
            border_color="#CBD5E1"
        )
        self.entry_pasta.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_selecionar = ctk.CTkButton(
            folder_row, 
            text="📁 Selecionar Pasta", 
            command=self._selecionar_pasta,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=self.cencosud_blue,
            hover_color=self.cencosud_hover,
            height=38,
            width=150
        )
        self.btn_selecionar.pack(side="right")

        # Status dos Arquivos Detectados
        self.lbl_arquivos_status = ctk.CTkLabel(
            card_pasta, 
            text="Nenhuma pasta selecionada.", 
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#64748B"
        )
        self.lbl_arquivos_status.pack(anchor="w", padx=20, pady=(0, 5))

        self.lbl_arquivos_lista = ctk.CTkLabel(
            card_pasta, 
            text="", 
            font=ctk.CTkFont(family="Consolas", size=11),
            text_color="#475569",
            justify="left"
        )
        self.lbl_arquivos_lista.pack(anchor="w", padx=20, pady=(0, 5))

        # 3. CARD: PAINEL DE EXECUÇÃO
        card_exec = ctk.CTkFrame(main_container, fg_color=self.card_bg, corner_radius=10, border_width=1, border_color="#E2E8F0")
        card_exec.pack(fill="x", pady=(0, 12), ipady=8)

        lbl_step2 = ctk.CTkLabel(
            card_exec, 
            text="2. Execução da Auditoria", 
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color="#1E293B"
        )
        lbl_step2.pack(anchor="w", padx=20, pady=(10, 6))

        # Botões de Ação
        action_row = ctk.CTkFrame(card_exec, fg_color="transparent")
        action_row.pack(fill="x", padx=20, pady=(0, 12))

        self.btn_iniciar = ctk.CTkButton(
            action_row, 
            text="▶ Iniciar Auditoria de Rescisão", 
            command=self._iniciar_auditoria,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=self.cencosud_blue,
            hover_color=self.cencosud_hover,
            height=42,
            state="disabled"
        )
        self.btn_iniciar.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_abrir_pasta = ctk.CTkButton(
            action_row, 
            text="📂 Abrir Pasta", 
            command=self._abrir_pasta_destino,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#F1F5F9",
            hover_color="#E2E8F0",
            text_color="#1E293B",
            height=42,
            width=130,
            state="disabled"
        )
        self.btn_abrir_pasta.pack(side="right")

        # Barra de Progresso
        self.progress_bar = ctk.CTkProgressBar(
            card_exec, 
            orientation="horizontal", 
            mode="determinate",
            progress_color=self.cencosud_orange,
            height=12
        )
        self.progress_bar.pack(fill="x", padx=20, pady=(0, 6))
        self.progress_bar.set(0)

        # Label de Status
        self.lbl_status = ctk.CTkLabel(
            card_exec, 
            text="Aguardando seleção da pasta...", 
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#64748B"
        )
        self.lbl_status.pack(anchor="w", padx=20, pady=(0, 5))

        # 4. CARD: CONSOLE DE ATIVIDADES
        card_log = ctk.CTkFrame(main_container, fg_color=self.card_bg, corner_radius=10, border_width=1, border_color="#E2E8F0")
        card_log.pack(fill="both", expand=True, pady=(0, 5))

        log_header = ctk.CTkFrame(card_log, fg_color="transparent")
        log_header.pack(fill="x", padx=20, pady=(8, 4))

        lbl_log_title = ctk.CTkLabel(
            log_header, 
            text="Log de Auditoria em Tempo Real", 
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#334155"
        )
        lbl_log_title.pack(side="left")

        btn_limpar_log = ctk.CTkButton(
            log_header,
            text="Limpar Log",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            fg_color="transparent",
            text_color="#64748B",
            hover_color="#F1F5F9",
            width=70,
            height=20,
            command=self._limpar_log
        )
        btn_limpar_log.pack(side="right")

        self.txt_log = ctk.CTkTextbox(
            card_log, 
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color="#0F172A",
            text_color="#E2E8F0",
            corner_radius=6,
            wrap="word"
        )
        self.txt_log.pack(fill="both", expand=True, padx=20, pady=(0, 12))

        # 5. FOOTER
        lbl_footer = ctk.CTkLabel(
            self, 
            text="Cencosud Brasil • Automação e Engenharia de Dados de RH", 
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="#94A3B8"
        )
        lbl_footer.pack(side="bottom", pady=(0, 8))

        self._log("Aplicação iniciada com sucesso. Selecione uma pasta com as folhas de rescisão.")

    def _log(self, mensagem):
        timestamp = datetime.now().strftime("%H:%M:%S")
        linha = f"[{timestamp}] {mensagem}\n"
        
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", linha)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

    def _selecionar_pasta(self):
        caminho = filedialog.askdirectory(title="Selecione a Pasta com as Folhas de Rescisão")
        if not caminho:
            return

        self.caminho_pasta.set(caminho)
        self._varrer_pasta(caminho)

    def _varrer_pasta(self, pasta):
        extensoes_validas = ('.csv', '.xlsb')
        self.arquivos_detectados = []

        try:
            for item in os.listdir(pasta):
                full_path = os.path.join(pasta, item)
                if os.path.isfile(full_path):
                    # Não reprocessar relatórios já gerados de auditoria
                    if item.lower().startswith("auditoria_"):
                        continue
                    if item.lower().endswith(extensoes_validas):
                        self.arquivos_detectados.append(full_path)
        except Exception as e:
            self._log(f"Erro ao listar pasta: {e}")
            messagebox.showerror("Erro", f"Não foi possível acessar a pasta:\n{e}")
            return

        total = len(self.arquivos_detectados)
        if total > 0:
            self.lbl_arquivos_status.configure(
                text=f"✓ {total} arquivo(s) detectado(s) para auditoria:",
                text_color="#16A34A"
            )
            # Exibir lista formatada com nomes dos arquivos
            nomes = [f"• {os.path.basename(f)} ({os.path.getsize(f) // 1024} KB)" for f in self.arquivos_detectados]
            self.lbl_arquivos_lista.configure(text="\n".join(nomes[:6]) + ("\n..." if len(nomes) > 6 else ""))
            self.btn_iniciar.configure(state="normal")
            self.btn_abrir_pasta.configure(state="normal")
            self.lbl_status.configure(
                text=f"{total} arquivo(s) pronto(s). Clique em 'Iniciar Auditoria de Rescisão'.",
                text_color="#1E293B"
            )
            self._log(f"Pasta selecionada: {pasta}")
            self._log(f"Encontrados {total} arquivos para conferência.")
        else:
            self.lbl_arquivos_status.configure(
                text="⚠️ Nenhum arquivo de rescisão (.csv ou .xlsb) encontrado nesta pasta.",
                text_color="#DC2626"
            )
            self.lbl_arquivos_lista.configure(text="")
            self.btn_iniciar.configure(state="disabled")
            self.lbl_status.configure(
                text="Nenhum arquivo elegível encontrado. Escolha outra pasta.",
                text_color="#DC2626"
            )
            self._log(f"Pasta selecionada: {pasta} (Sem arquivos .csv ou .xlsb elegíveis)")

    def _iniciar_auditoria(self):
        if self.is_running:
            return

        if not self.arquivos_detectados:
            messagebox.showwarning("Aviso", "Nenhum arquivo encontrado para auditar.")
            return

        self.is_running = True
        self.btn_iniciar.configure(state="disabled", text="⏳ Processando Auditoria...")
        self.btn_selecionar.configure(state="disabled")
        self.progress_bar.set(0)

        # Rodar auditoria em segundo plano (Multithreading)
        thread = threading.Thread(target=self._processar_em_thread, daemon=True)
        thread.start()

    def _processar_em_thread(self):
        total_arquivos = len(self.arquivos_detectados)
        pasta_origem = self.caminho_pasta.get()

        self._log("=" * 60)
        self._log(f"INICIANDO CICLO DE AUDITORIA: {total_arquivos} arquivo(s)")
        self._log("=" * 60)

        arquivos_sucesso = 0
        arquivos_falha = 0

        for i, arq_entrada in enumerate(self.arquivos_detectados, start=1):
            nome_arquivo = os.path.basename(arq_entrada)
            base_name, _ = os.path.splitext(nome_arquivo)

            # Nome padrão corporativo para a saída
            arq_saida = os.path.join(pasta_origem, f"Auditoria_{base_name}.xlsx")

            self._atualizar_status(f"[{i}/{total_arquivos}] Processando: {nome_arquivo}...")
            self._log(f"[{i}/{total_arquivos}] Iniciando: {nome_arquivo}")

            def progress_callback(progresso_arquivo):
                # Progresso geral ponderado
                prog_geral = ((i - 1) + progresso_arquivo) / total_arquivos
                self.after(0, lambda p=prog_geral: self.progress_bar.set(p))

            def log_callback(msg):
                self.after(0, lambda m=msg: self._log(f"  -> {m}"))

            try:
                sucesso = auditar_folha_massiva(
                    caminho_entrada=arq_entrada, 
                    caminho_saida_xlsx=arq_saida,
                    logger=log_callback,
                    progress_fn=progress_callback
                )
                if sucesso:
                    arquivos_sucesso += 1
                    self._log(f"✓ Arquivo gerado com sucesso: Auditoria_{base_name}.xlsx")
                else:
                    arquivos_falha += 1
                    self._log(f"❌ Falha no processamento de: {nome_arquivo}")
            except Exception as e:
                arquivos_falha += 1
                self._log(f"❌ Erro crítico em {nome_arquivo}: {e}")

        # Finalização
        self.after(0, lambda: self._finalizar_auditoria(arquivos_sucesso, arquivos_falha, total_arquivos))

    def _atualizar_status(self, texto):
        self.after(0, lambda: self.lbl_status.configure(text=texto, text_color=self.cencosud_blue))

    def _finalizar_auditoria(self, sucesso, falhas, total):
        self.is_running = False
        self.progress_bar.set(1.0)
        self.btn_iniciar.configure(state="normal", text="▶ Iniciar Auditoria de Rescisão")
        self.btn_selecionar.configure(state="normal")
        self.btn_abrir_pasta.configure(state="normal")

        msg_final = f"Auditoria concluída! {sucesso}/{total} arquivos gerados com sucesso."
        if falhas > 0:
            msg_final += f" ({falhas} com erro)"

        self.lbl_status.configure(
            text=msg_final,
            text_color="#16A34A" if falhas == 0 else "#DC2626"
        )

        self._log("=" * 60)
        self._log(msg_final)
        self._log("=" * 60)

        # Exibir alerta de sucesso
        if falhas == 0:
            resposta = messagebox.askyesno(
                "Auditoria Concluída", 
                f"Auditoria finalizada com sucesso!\n\nForam gerados {sucesso} relatório(s) na pasta.\nDeseja abrir a pasta agora?"
            )
            if resposta:
                self._abrir_pasta_destino()
        else:
            messagebox.showwarning(
                "Auditoria Finalizada com Alertas",
                f"A auditoria terminou, mas houve falha em {falhas} arquivo(s).\nVerifique o console de log para mais detalhes."
            )

    def _abrir_pasta_destino(self):
        pasta = self.caminho_pasta.get()
        if pasta and os.path.exists(pasta):
            try:
                os.startfile(pasta)
            except Exception as e:
                self._log(f"Erro ao abrir pasta: {e}")
        else:
            messagebox.showwarning("Aviso", "A pasta selecionada não existe ou não foi definida.")

if __name__ == "__main__":
    app = AuditorRescisaoApp()
    app.mainloop()
