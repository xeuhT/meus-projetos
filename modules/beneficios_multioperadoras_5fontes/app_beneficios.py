import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import sys
from PIL import Image
import os
import queue
import logging
from pathlib import Path
from datetime import datetime

# Adjust sys.path to ensure modules are found
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import core

# Cencosud Colors
CENCOSUD_BLUE = "#005099"
CENCOSUD_ORANGE = "#F37021"
CENCOSUD_LIGHT_BG = "#F5F6F8"
CENCOSUD_TEXT_DARK = "#333333"

ctk.set_appearance_mode("Light")

class QueueHandler(logging.Handler):
    """Class to send logging records to a queue"""
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(self.format(record))

class AuditoriaApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Motor Analítico de Benefícios - Cencosud")
        self.geometry("850x700")
        self.configure(fg_color=CENCOSUD_LIGHT_BG)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # Header Frame (Logo + Title)
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        
        # Load Logo
        if getattr(sys, 'frozen', False):
            base_path = Path(sys._MEIPASS)
        else:
            base_path = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
            
        logo_path = base_path / "logo.png"
        if logo_path.exists():
            img = Image.open(logo_path)
            orig_w, orig_h = img.size
            h = 80
            w = int(orig_w * (h / float(orig_h)))

            logo_img = ctk.CTkImage(light_image=img, dark_image=img, size=(w, h))
            self.logo_label = ctk.CTkLabel(self.header_frame, image=logo_img, text="")
            self.logo_label.grid(row=0, column=0, padx=(0, 20))
            
        # Title
        self.title_label = ctk.CTkLabel(self.header_frame, text="Auditoria de Benefícios e FOPAG", 
                                        font=ctk.CTkFont(size=24, weight="bold"), text_color=CENCOSUD_BLUE)
        self.title_label.grid(row=0, column=1, sticky="w")

        # Input Frame
        self.input_frame = ctk.CTkFrame(self, fg_color="white", corner_radius=10)
        self.input_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.input_frame.grid_columnconfigure(1, weight=1)

        self.pasta_label = ctk.CTkLabel(self.input_frame, text="Pasta de Entrada:", font=ctk.CTkFont(weight="bold"), text_color=CENCOSUD_TEXT_DARK)
        self.pasta_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.pasta_entry = ctk.CTkEntry(self.input_frame, placeholder_text="Selecione a pasta com os arquivos...", border_color=CENCOSUD_BLUE)
        self.pasta_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        self.buscar_btn = ctk.CTkButton(self.input_frame, text="Buscar Pasta", command=self.buscar_pasta, 
                                        fg_color=CENCOSUD_BLUE, hover_color="#003d7a")
        self.buscar_btn.grid(row=0, column=2, padx=10, pady=10)

        # Validation Display area
        self.validation_frame = ctk.CTkScrollableFrame(self, fg_color="white", corner_radius=10, height=200)
        self.validation_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        
        self.validation_label = ctk.CTkLabel(self.validation_frame, text="Selecione a pasta para verificar os arquivos.", 
                                             text_color="gray", font=ctk.CTkFont(weight="bold"))
        self.validation_label.pack(pady=10, padx=10, anchor="w")

        self.iniciar_btn = ctk.CTkButton(self.input_frame, text="Iniciar Auditoria", command=self.iniciar_auditoria_thread, 
                                         state="disabled", fg_color=CENCOSUD_ORANGE, hover_color="#d95a10", font=ctk.CTkFont(weight="bold"))
        self.iniciar_btn.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
        from tkinter import ttk
        self.progress = ttk.Progressbar(self.input_frame, mode="indeterminate")
        self.progress.grid(row=2, column=0, columnspan=3, padx=10, pady=(0, 10), sticky="ew")

        # Log Frame
        self.log_frame = ctk.CTkFrame(self, fg_color="white", corner_radius=10)
        self.log_frame.grid(row=3, column=0, padx=20, pady=(10, 20), sticky="nsew")
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(0, weight=1)

        self.log_textbox = ctk.CTkTextbox(self.log_frame, state="disabled", font=ctk.CTkFont(family="Consolas", size=12), 
                                          fg_color="#F0F0F0", text_color="#333333")
        self.log_textbox.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Setup custom logging handler for the UI
        self.log_queue = queue.Queue()
        self.queue_handler = QueueHandler(self.log_queue)
        formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', datefmt='%H:%M:%S')
        self.queue_handler.setFormatter(formatter)
        
        # Patch core.configurar_log
        self.original_configurar_log = core.configurar_log
        core.configurar_log = self.patched_configurar_log
        
        # Add UI logger
        logging.getLogger().addHandler(self.queue_handler)
        logging.getLogger().setLevel(logging.INFO)

        self.after(100, self.poll_log_queue)

    def patched_configurar_log(self, pasta_log: Path):
        pasta_log.mkdir(parents=True, exist_ok=True)
        arquivo_log = pasta_log / f"auditoria_{datetime.now():%Y%m%d_%H%M%S}.log"
        fh = logging.FileHandler(arquivo_log, encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s"))
        logging.getLogger().addHandler(fh)
        
        sh = logging.StreamHandler(sys.stdout)
        sh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s"))
        logging.getLogger().addHandler(sh)
        return arquivo_log

    def validar_planilhas(self, pasta: Path):
        self.iniciar_btn.configure(state="disabled")
        self.progress.start(10)
        
        # Clear validation frame
        for widget in self.validation_frame.winfo_children():
            widget.destroy()
            
        header = ctk.CTkLabel(self.validation_frame, text="Relatório de Identificação de Arquivos:", font=ctk.CTkFont(weight="bold", size=14))
        header.pack(anchor="w", padx=10, pady=(5, 10))
        
        arquivos = core.descobrir_arquivos(pasta)
        itens_validacao = []
        
        # RH e Corporativo
        colab_ok = False
        if arquivos.get("colaboradores"):
            try:
                df_c_test = core.leitores.ler_colaboradores(arquivos["colaboradores"])
                colab_ok = not df_c_test.empty
            except Exception:
                colab_ok = False
        itens_validacao.append(("Base Mestre de Colaboradores (RH)", colab_ok))
        itens_validacao.append(("Base de Dependentes (Marsh)", bool(arquivos.get("dependentes"))))
        itens_validacao.append(("Base de Desligados (RH)", bool(arquivos.get("demitidos"))))
        itens_validacao.append(("Dicionário de Eventos (Eventos 1.xlsx)", (pasta / "Eventos 1.xlsx").exists()))
        itens_validacao.append(("FOPAG Financeira Aspirador", 
bool(list(pasta.glob("FOPA_Eventos_por_Colaborador*")))))
        
        # Operadoras
        operadoras_requiridas = set(core.PADROES_ATIVOS.values())
        ativos_encontrados = set(arquivos.get("ativos", {}).keys())
        fopag_encontrados = set(arquivos.get("fopag", {}).keys())
        
        for op in sorted(operadoras_requiridas):
            itens_validacao.append((f"Listagem de Ativos ({op})", op in ativos_encontrados))
            itens_validacao.append((f"Movimentações FOPAG ({op})", op in fopag_encontrados))
            
        faltantes_count = 0
        for nome, status in itens_validacao:
            cor = "green" if status else "red"
            icone = "✅" if status else "❌"
            if not status: faltantes_count += 1
            
            lbl = ctk.CTkLabel(self.validation_frame, text=f"{icone} {nome}", text_color=cor, font=ctk.CTkFont(weight="bold" if not status else "normal"))
            lbl.pack(anchor="w", padx=20, pady=2)
            
        if faltantes_count > 0:
            status_lbl = ctk.CTkLabel(self.validation_frame, text=f"\n⚠️ Faltam {faltantes_count} arquivos. O relatório será gerado SEM essas bases.", text_color="#FF8C00", font=ctk.CTkFont(weight="bold"))
            status_lbl.pack(anchor="w", padx=10, pady=10)
            self.iniciar_btn.configure(state="normal")
        else:
            status_lbl = ctk.CTkLabel(self.validation_frame, text="\n✅ Todos os arquivos identificados! Pronto para iniciar.", text_color="green", font=ctk.CTkFont(weight="bold"))
            status_lbl.pack(anchor="w", padx=10, pady=10)
            self.iniciar_btn.configure(state="normal")
        self.iniciar_btn.configure(state="normal")

    def buscar_pasta(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.pasta_entry.delete(0, ctk.END)
            self.pasta_entry.insert(0, pasta)
            self.validar_planilhas(Path(pasta))

    def poll_log_queue(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get()
            self.log_textbox.configure(state="normal")
            self.log_textbox.insert(ctk.END, msg + '\n')
            self.log_textbox.see(ctk.END)
            self.log_textbox.configure(state="disabled")
        self.after(100, self.poll_log_queue)

    def iniciar_auditoria_thread(self):
        pasta_str = self.pasta_entry.get()
        if not pasta_str:
            messagebox.showwarning("Aviso", "Por favor, selecione a pasta de entrada primeiro.")
            return

        self.iniciar_btn.configure(state="disabled")
        self.progress.start(10)
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("0.0", ctk.END)
        self.log_textbox.configure(state="disabled")
        
        logging.info(f"Iniciando auditoria. Pasta selecionada: {pasta_str}")

        t = threading.Thread(target=self.processar, args=(pasta_str,), daemon=True)
        t.start()

    def processar(self, pasta_str):
        try:
            pasta_entrada = Path(pasta_str)
            pasta_saida = pasta_entrada / f"Auditoria_Consolidada_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            
            # Setup file logging
            core.configurar_log(pasta_entrada / "logs")

            # Run execution
            core.executar(pasta_entrada, pasta_saida)
            
            logging.info("Auditoria finalizada com sucesso!")
            self.after(0, lambda: messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso em:\n{pasta_saida}"))
            
        except Exception as e:
            logging.error(f"Erro durante a execução: {e}", exc_info=True)
            self.after(0, lambda: messagebox.showerror("Erro", f"Ocorreu um erro durante a auditoria:\n{e}"))
            
        finally:
            self.after(0, self.progress.stop)
            self.after(0, lambda: self.iniciar_btn.configure(state="normal"))

if __name__ == "__main__":
    app = AuditoriaApp()
    app.mainloop()
