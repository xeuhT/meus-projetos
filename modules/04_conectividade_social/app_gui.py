"""
Interface Gráfica Corporativa (GUI) - Robô Conectividade Social V2
Cencosud Brasil - Gestão de Benefícios e Fins Rescisórios
Desenvolvido com o Padrão Visual Institucional Cencosud (#095AA5 / #FF690A).
"""
import os
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict, Any

# Garantir importações locais
sys.path.insert(0, str(Path(__file__).resolve().parent))

import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image

from config import (
    BASE_DIR,
    DIR_SAIDA,
    DIR_MODELOS,
    DIR_EXTRATOS,
    CERTIFICADOS_CENCOSUD,
    CERTIFICADO_PADRAO,
    BASE_CONTA_PADRAO,
    DELAY_ENTRE_CONSULTAS
)
from excel_handler import carregar_lista_pis, exportar_relatorio_extratos
from browser_automation import ConectividadeSocialBot

# Paleta Oficial de Cores Cencosud (extraída diretamente da identidade visual)
COR_AZUL_CENCOSUD = "#095AA5"        # Azul Real Cencosud
COR_AZUL_HOVER = "#074885"           # Azul Escuro Interativo
COR_LARANJA_CENCOSUD = "#FF690A"     # Laranja Cencosud
COR_LARANJA_HOVER = "#E05500"        # Laranja Escuro
COR_VERDE_SUCESSO = "#2E7D32"        # Verde Homologação
COR_VERDE_HOVER = "#1B5E20"
COR_VERMELHO_CANCELAR = "#C62828"

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class AppConectividadeSocial(ctk.CTk):
    """Janela Principal do Aplicativo Conectividade Social V2 Cencosud."""

    def __init__(self):
        super().__init__()

        self.title("Cencosud Brasil - Extrato FGTS & Fins Rescisórios (Conectividade Social V2)")
        self.geometry("980x820")
        self.minsize(900, 720)

        # Caminhos padrão (compatível com pacote standalone .exe e modo desenvolvimento)
        self.diretorio_base = BASE_DIR
        self.diretorio_recursos = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))

        # Aplicar ícone oficial da Cencosud na janela e barra de tarefas
        caminho_icone = self.diretorio_recursos / "assets" / "cencosud_logo.ico"
        if caminho_icone.exists():
            try:
                self.iconbitmap(str(caminho_icone))
            except Exception:
                pass

        planilha_padrao = self.diretorio_base / "teste" / "nomes e pis para teste.xlsx"
        self.caminho_planilha_selecionada: Optional[Path] = planilha_padrao if planilha_padrao.exists() else None
        self.caminho_pasta_saida = DIR_SAIDA
        self.caminho_pasta_saida.mkdir(parents=True, exist_ok=True)

        # Estado da Aplicação
        self.processo_em_execucao = False
        self.cancelamento_solicitado = False
        self.ultimo_relatorio_gerado: Optional[Path] = None
        self.bot_ativo: Optional[ConectividadeSocialBot] = None

        # Carregar Logomarca da Cencosud
        self.logo_image = self._carregar_logo()

        # Construção da Interface
        self._criar_layout()

        # Carregar dados da planilha padrão se existir
        if self.caminho_planilha_selecionada and self.caminho_planilha_selecionada.exists():
            self._atualizar_info_planilha(self.caminho_planilha_selecionada)

    def _carregar_logo(self) -> Optional[ctk.CTkImage]:
        """Carrega a logomarca institucional da Cencosud a partir dos recursos empacotados."""
        caminhos_logo = [
            self.diretorio_recursos / "assets" / "cencosud_logo.png",
            self.diretorio_recursos / "assets" / "cencosud_logo_cropped.png",
            self.diretorio_base / "assets" / "cencosud_logo.png"
        ]
        for p in caminhos_logo:
            if p.exists():
                try:
                    pil_img = Image.open(p)
                    return ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(64, 64))
                except Exception:
                    continue
        return None

    def _criar_layout(self):
        # 1. Cabeçalho Institucional com Logo e Cores da Cencosud
        header_frame = ctk.CTkFrame(self, fg_color=COR_AZUL_CENCOSUD, corner_radius=0, height=90)
        header_frame.pack(fill="x", padx=0, pady=0)

        header_content = ctk.CTkFrame(header_frame, fg_color="transparent")
        header_content.pack(fill="both", expand=True, padx=20, pady=10)

        # Inserção da Logo
        if self.logo_image:
            lbl_logo = ctk.CTkLabel(header_content, image=self.logo_image, text="")
            lbl_logo.pack(side="left", padx=(5, 18))

        # Textos do Cabeçalho
        title_box = ctk.CTkFrame(header_content, fg_color="transparent")
        title_box.pack(side="left", fill="both", expand=True)

        lbl_titulo = ctk.CTkLabel(
            title_box,
            text="CENCOSUD BRASIL - GESTÃO DE BENEFÍCIOS RH",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#FFFFFF"
        )
        lbl_titulo.pack(anchor="w", pady=(2, 2))

        lbl_subtitulo = ctk.CTkLabel(
            title_box,
            text="Robô de Consulta em Lote de Extratos FGTS & Fins Rescisórios com Exportação Oficial em PDF",
            font=ctk.CTkFont(size=12),
            text_color="#D8E8F8"
        )
        lbl_subtitulo.pack(anchor="w")

        # Container Principal Rolável
        scroll_container = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll_container.pack(fill="both", expand=True, padx=25, pady=12)

        # 2. Card 1: Arquivos e Diretórios (Planilha de Entrada & Pasta de Saída)
        card_arquivos = ctk.CTkFrame(scroll_container, corner_radius=10)
        card_arquivos.pack(fill="x", pady=8, padx=5)

        lbl_sec1 = ctk.CTkLabel(
            card_arquivos,
            text="1. Arquivos e Diretórios:",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COR_AZUL_CENCOSUD
        )
        lbl_sec1.pack(anchor="w", padx=15, pady=(12, 8))

        # Linha 1: Planilha de Entrada
        lbl_sub_ent = ctk.CTkLabel(card_arquivos, text="Planilha de Entrada (Excel com Colaboradores):", font=ctk.CTkFont(size=12, weight="bold"))
        lbl_sub_ent.pack(anchor="w", padx=15, pady=(0, 3))

        file_row = ctk.CTkFrame(card_arquivos, fg_color="transparent")
        file_row.pack(fill="x", padx=15, pady=(0, 4))

        self.txt_caminho_arquivo = ctk.CTkEntry(
            file_row,
            placeholder_text="Selecione o arquivo Excel com as colunas NOME_COLABORADOR e PIS/PASEP...",
            font=ctk.CTkFont(size=12)
        )
        self.txt_caminho_arquivo.pack(side="left", fill="x", expand=True, padx=(0, 10))

        btn_procurar_planilha = ctk.CTkButton(
            file_row,
            text="Procurar Planilha...",
            command=self._selecionar_arquivo,
            width=150,
            fg_color=COR_AZUL_CENCOSUD,
            hover_color=COR_AZUL_HOVER
        )
        btn_procurar_planilha.pack(side="right")

        self.lbl_status_planilha = ctk.CTkLabel(
            card_arquivos,
            text="Nenhum arquivo selecionado.",
            font=ctk.CTkFont(size=11, slant="italic"),
            text_color="gray"
        )
        self.lbl_status_planilha.pack(anchor="w", padx=15, pady=(0, 10))

        # Linha 2: Pasta de Saída dos Relatórios e PDFs
        lbl_sub_sai = ctk.CTkLabel(card_arquivos, text="Pasta de Saída dos Relatórios e Extratos em PDF:", font=ctk.CTkFont(size=12, weight="bold"))
        lbl_sub_sai.pack(anchor="w", padx=15, pady=(4, 3))

        saida_row = ctk.CTkFrame(card_arquivos, fg_color="transparent")
        saida_row.pack(fill="x", padx=15, pady=(0, 15))

        self.txt_caminho_saida = ctk.CTkEntry(
            saida_row,
            font=ctk.CTkFont(size=12)
        )
        self.txt_caminho_saida.insert(0, str(self.caminho_pasta_saida))
        self.txt_caminho_saida.pack(side="left", fill="x", expand=True, padx=(0, 10))

        btn_procurar_saida = ctk.CTkButton(
            saida_row,
            text="Alterar Pasta...",
            command=self._selecionar_pasta_saida,
            width=150,
            fg_color=COR_LARANJA_CENCOSUD,
            hover_color=COR_LARANJA_HOVER
        )
        btn_procurar_saida.pack(side="right")

        # 3. Card 2: Parâmetros de Consulta (Base UF + Certificado)
        card_params = ctk.CTkFrame(scroll_container, corner_radius=10)
        card_params.pack(fill="x", pady=8, padx=5)

        lbl_sec2 = ctk.CTkLabel(
            card_params,
            text="2. Configurações de Consulta e Certificado Digital:",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COR_AZUL_CENCOSUD
        )
        lbl_sec2.pack(anchor="w", padx=15, pady=(12, 10))

        params_grid = ctk.CTkFrame(card_params, fg_color="transparent")
        params_grid.pack(fill="x", padx=15, pady=(0, 15))

        # Coluna 1: Base da Conta
        col_base = ctk.CTkFrame(params_grid, fg_color="transparent")
        col_base.pack(side="left", fill="x", expand=True, padx=(0, 15))

        lbl_base = ctk.CTkLabel(col_base, text="Base da Conta (UF):", font=ctk.CTkFont(size=12, weight="bold"))
        lbl_base.pack(anchor="w", pady=(0, 3))

        self.cb_base_uf = ctk.CTkComboBox(
            col_base,
            values=[
                "SE-SERGIPE", "BA-BAHIA", "MG-MINAS GERAIS", "CE-CEARÁ",
                "GO-GOIÁS", "PE-PERNAMBUCO", "RJ-RIO DE JANEIRO", "SP-SÃO PAULO"
            ],
            width=260
        )
        self.cb_base_uf.set("SE-SERGIPE")
        self.cb_base_uf.pack(fill="x")

        # Coluna 2: Certificado Digital
        col_cert = ctk.CTkFrame(params_grid, fg_color="transparent")
        col_cert.pack(side="left", fill="x", expand=True)

        lbl_cert = ctk.CTkLabel(col_cert, text="Certificado Digital (e-CNPJ Cencosud):", font=ctk.CTkFont(size=12, weight="bold"))
        lbl_cert.pack(anchor="w", pady=(0, 3))

        self.cb_certificado = ctk.CTkComboBox(
            col_cert,
            values=[
                "CENCOSUD BRASIL COMERCIAL LTDA",
                "CENCOSUD BRASIL ATACADO LTDA",
                "PERINI COMERCIAL DE ALIMENTOS LTDA"
            ],
            width=320
        )
        self.cb_certificado.set("CENCOSUD BRASIL COMERCIAL LTDA")
        self.cb_certificado.pack(fill="x")

        # 4. Card 3: Ações de Execução e Progresso
        card_acoes = ctk.CTkFrame(scroll_container, corner_radius=10)
        card_acoes.pack(fill="x", pady=8, padx=5)

        lbl_sec3 = ctk.CTkLabel(
            card_acoes,
            text="3. Execução das Consultas:",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COR_AZUL_CENCOSUD
        )
        lbl_sec3.pack(anchor="w", padx=15, pady=(12, 10))

        btn_row = ctk.CTkFrame(card_acoes, fg_color="transparent")
        btn_row.pack(fill="x", padx=15, pady=(0, 10))

        self.btn_iniciar = ctk.CTkButton(
            btn_row,
            text="▶ INICIAR CONSULTA EM LOTE",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._iniciar_execucao,
            height=44,
            fg_color=COR_AZUL_CENCOSUD,
            hover_color=COR_AZUL_HOVER
        )
        self.btn_iniciar.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_cancelar = ctk.CTkButton(
            btn_row,
            text="⏹ Cancelar",
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self._cancelar_execucao,
            height=44,
            width=120,
            fg_color=COR_VERMELHO_CANCELAR,
            hover_color="#991B1B",
            state="disabled"
        )
        self.btn_cancelar.pack(side="right")

        self.progress_bar = ctk.CTkProgressBar(card_acoes, height=14, progress_color=COR_LARANJA_CENCOSUD)
        self.progress_bar.pack(fill="x", padx=15, pady=(5, 5))
        self.progress_bar.set(0.0)

        self.lbl_progresso_status = ctk.CTkLabel(
            card_acoes,
            text="Pronto para iniciar.",
            font=ctk.CTkFont(size=12)
        )
        self.lbl_progresso_status.pack(anchor="w", padx=15, pady=(0, 10))

        # Painel de Contadores
        badges_frame = ctk.CTkFrame(card_acoes, fg_color="transparent")
        badges_frame.pack(fill="x", padx=15, pady=(0, 12))

        self.badge_total = ctk.CTkLabel(badges_frame, text="Total: 0", font=ctk.CTkFont(size=12, weight="bold"), width=90)
        self.badge_total.pack(side="left", padx=(0, 10))

        self.badge_sucesso = ctk.CTkLabel(badges_frame, text="Sucesso: 0", font=ctk.CTkFont(size=12, weight="bold"), text_color=COR_VERDE_SUCESSO, width=100)
        self.badge_sucesso.pack(side="left", padx=(0, 10))

        self.badge_alertas = ctk.CTkLabel(badges_frame, text="Não Localizados / Erros: 0", font=ctk.CTkFont(size=12, weight="bold"), text_color=COR_LARANJA_CENCOSUD, width=180)
        self.badge_alertas.pack(side="left")

        # 5. Card 4: Console de Log em Tempo Real
        card_log = ctk.CTkFrame(scroll_container, corner_radius=10)
        card_log.pack(fill="both", expand=True, pady=8, padx=5)

        lbl_sec4 = ctk.CTkLabel(
            card_log,
            text="4. Log de Monitoramento em Tempo Real:",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COR_AZUL_CENCOSUD
        )
        lbl_sec4.pack(anchor="w", padx=15, pady=(12, 5))

        self.txt_log = ctk.CTkTextbox(
            card_log,
            height=180,
            font=ctk.CTkFont(family="Consolas", size=11),
            activate_scrollbars=True
        )
        self.txt_log.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        # 6. Rodapé com Botões de Conclusão do Relatório
        footer_card = ctk.CTkFrame(self, corner_radius=0, height=70)
        footer_card.pack(fill="x", side="bottom", padx=0, pady=0)

        self.btn_abrir_relatorio = ctk.CTkButton(
            footer_card,
            text="📊 ABRIR RELATÓRIO EXCEL GERADO",
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self._abrir_relatorio_excel,
            fg_color=COR_VERDE_SUCESSO,
            hover_color=COR_VERDE_HOVER,
            height=42,
            state="disabled"
        )
        self.btn_abrir_relatorio.pack(side="left", padx=20, pady=14)

        self.btn_abrir_pasta = ctk.CTkButton(
            footer_card,
            text="📁 Abrir Pasta de Relatórios & PDFs",
            font=ctk.CTkFont(size=12),
            command=self._abrir_pasta_relatorios,
            height=42,
            width=220,
            fg_color="gray",
            hover_color="#555555"
        )
        self.btn_abrir_pasta.pack(side="right", padx=20, pady=14)

    def _log(self, mensagem: str):
        """Adiciona uma mensagem formatada ao console de log com timestamp."""
        hora = datetime.now().strftime("%H:%M:%S")
        texto_completo = f"[{hora}] {mensagem}\n"
        self.txt_log.insert("end", texto_completo)
        self.txt_log.see("end")

    def _selecionar_arquivo(self):
        """Abre caixa de diálogo para seleção do arquivo Excel de entrada."""
        caminho = filedialog.askopenfilename(
            title="Selecione a Planilha de Colaboradores",
            filetypes=[("Planilhas Excel", "*.xlsx;*.xls"), ("Todos os Arquivos", "*.*")]
        )
        if caminho:
            caminho_path = Path(caminho)
            self._atualizar_info_planilha(caminho_path)

    def _selecionar_pasta_saida(self):
        """Permite ao usuário escolher onde salvar os relatórios e extratos em PDF."""
        pasta = filedialog.askdirectory(
            title="Selecione a Pasta de Saída dos Relatórios e Extratos em PDF",
            initialdir=str(self.caminho_pasta_saida)
        )
        if pasta:
            self.caminho_pasta_saida = Path(pasta)
            self.txt_caminho_saida.delete(0, "end")
            self.txt_caminho_saida.insert(0, str(self.caminho_pasta_saida))
            self._log(f"Pasta de saída atualizada para: '{self.caminho_pasta_saida}'")

    def _atualizar_info_planilha(self, caminho: Path):
        """Lê a planilha e atualiza as informações cadastrais na tela."""
        self.caminho_planilha_selecionada = caminho
        self.txt_caminho_arquivo.delete(0, "end")
        self.txt_caminho_arquivo.insert(0, str(caminho))

        try:
            colaboradores = carregar_lista_pis(caminho)
            qtd = len(colaboradores)
            self.lbl_status_planilha.configure(
                text=f"✔ Planilha validada com sucesso! Total de {qtd} colaboradores encontrados.",
                text_color=COR_VERDE_SUCESSO
            )
            self.badge_total.configure(text=f"Total: {qtd}")
            self.badge_sucesso.configure(text="Sucesso: 0")
            self.badge_alertas.configure(text="Não Localizados / Erros: 0")
            self._log(f"Planilha carregada: '{caminho.name}' ({qtd} colaboradores detectados).")
        except Exception as e:
            self.lbl_status_planilha.configure(
                text=f"❌ Erro ao ler planilha: {str(e)}",
                text_color=COR_VERMELHO_CANCELAR
            )

    def _obter_chave_certificado(self) -> str:
        """Mapeia o nome do certificado selecionado no combobox para a chave de config."""
        nome_cert = self.cb_certificado.get().upper()
        if "ATACADO" in nome_cert:
            return "atacado"
        elif "PERINI" in nome_cert:
            return "perini"
        return "comercial"

    def _iniciar_execucao(self):
        """Dispara o processamento em lote em uma thread secundária."""
        if not self.caminho_planilha_selecionada or not self.caminho_planilha_selecionada.exists():
            messagebox.showwarning("Atenção", "Por favor, selecione uma planilha Excel válida antes de iniciar.")
            return

        # Validar ou criar pasta de saída informada
        pasta_saida_str = self.txt_caminho_saida.get().strip()
        if pasta_saida_str:
            self.caminho_pasta_saida = Path(pasta_saida_str)
        self.caminho_pasta_saida.mkdir(parents=True, exist_ok=True)

        self.processo_em_execucao = True
        self.cancelamento_solicitado = False
        self.btn_iniciar.configure(state="disabled")
        self.btn_cancelar.configure(state="normal")
        self.btn_abrir_relatorio.configure(state="disabled")
        self.progress_bar.set(0.0)

        # Disparar thread de automação
        thread_trabalho = threading.Thread(target=self._executar_robo_worker, daemon=True)
        thread_trabalho.start()

    def _cancelar_execucao(self):
        """Sinaliza cancelamento para o loop."""
        if self.processo_em_execucao:
            self.cancelamento_solicitado = True
            self._log("⚠️ Solicitação de cancelamento enviada pelo usuário...")
            self.lbl_progresso_status.configure(text="Cancelando processo com segurança...")
            if self.bot_ativo:
                try:
                    self.bot_ativo.fechar_navegador()
                except Exception:
                    pass

    def _executar_robo_worker(self):
        """Worker que roda em segundo plano executando a automação Selenium."""
        caminho_planilha = self.caminho_planilha_selecionada
        pasta_saida = self.caminho_pasta_saida
        base_uf = self.cb_base_uf.get()
        chave_cert = self._obter_chave_certificado()

        self.after(0, self._log, "Iniciando robô Conectividade Social V2...")
        self.after(0, self._log, f"Base selecionada: {base_uf} | Certificado: {chave_cert.upper()}")
        self.after(0, self._log, f"Pasta de saída configurada: '{pasta_saida}'")

        try:
            colaboradores = carregar_lista_pis(caminho_planilha)
            total = len(colaboradores)
            if total == 0:
                self.after(0, self._log, "❌ Nenhum colaborador válido encontrado na planilha.")
                self.after(0, self._finalizar_execucao, None, 0, 0)
                return

            self.after(0, self._log, f"Total a processar: {total} colaboradores.")

            # Inicializar Bot com o diretório de saída customizado
            self.bot_ativo = ConectividadeSocialBot(tipo_certificado=chave_cert, diretorio_saida=pasta_saida)
            
            self.after(0, self._log, "Abrindo Google Chrome e navegando para a Caixa...")
            self.after(0, self._log, ">>> POR FAVOR, SELECIONE O CERTIFICADO DIGITAL NA JANELA DO WINDOWS E CLIQUE EM OK <<<")
            
            # 1. Login e Autenticação
            if not self.bot_ativo.login_conectividade():
                self.after(0, self._log, "❌ Tempo de autenticação excedido ou falha no login.")
                self.after(0, self._finalizar_execucao, None, 0, 0)
                return

            self.after(0, self._log, "✔ Autenticação confirmada com sucesso!")

            # 2. Navegar para Solicitar Extrato
            if not self.bot_ativo.navegar_para_solicitar_extrato():
                self.after(0, self._log, "❌ Falha ao acessar formulário de solicitação de extrato.")
                self.after(0, self._finalizar_execucao, None, 0, 0)
                return

            self.after(0, self._log, "✔ Formulário de pesquisa carregado. Iniciando consultas em lote...")

            resultados: List[Dict[str, Any]] = []
            sucessos = 0
            alertas = 0

            # 3. Loop das Consultas
            for idx, colab in enumerate(colaboradores, 1):
                if self.cancelamento_solicitado:
                    self.after(0, self._log, "⏹ Execução interrompida pelo usuário.")
                    break

                pis = colab["pis"]
                nome = colab.get("nome", "")
                matricula = colab.get("matricula", "")

                progresso = idx / total
                self.after(0, self._atualizar_progresso, progresso, idx, total, nome, pis)
                self.after(0, self._log, f"[{idx}/{total}] Consultando PIS {pis} - {nome}...")

                try:
                    res = self.bot_ativo.preencher_formulario_e_consultar(pis=pis, nome=nome, base_uf=base_uf)
                    res["matricula"] = matricula
                    res["nome_informado"] = nome
                    res["timestamp_consulta"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    resultados.append(res)

                    status = str(res.get("status") or "").upper()
                    if "SUCESSO" in status:
                        sucessos += 1
                        val_rescisorio = res.get("valor_base_rescisorios")
                        val_fmt = f"R$ {val_rescisorio:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if val_rescisorio else "R$ 0,00"
                        self.after(0, self._log, f"    ✔ SUCESSO! Fins Rescisórios: {val_fmt}")
                        if res.get("caminho_pdf"):
                            nome_pdf = Path(res["caminho_pdf"]).name
                            self.after(0, self._log, f"    📄 Extrato Oficial em PDF gerado: {nome_pdf}")
                    else:
                        alertas += 1
                        self.after(0, self._log, f"    ⚠️ ALERTA: {res.get('mensagem', 'Conta não localizada')}")

                    self.after(0, self._atualizar_contadores, sucessos, alertas)

                    # Retorno para o próximo colaborador
                    if idx < total and not self.cancelamento_solicitado:
                        self.bot_ativo.retornar_para_formulario()
                        time.sleep(DELAY_ENTRE_CONSULTAS)

                except Exception as ex_colab:
                    alertas += 1
                    self.after(0, self._log, f"    ❌ Erro no colaborador: {ex_colab}")
                    resultados.append({
                        "matricula": matricula,
                        "nome_informado": nome,
                        "pis_pesquisado": pis,
                        "base_uf": base_uf,
                        "status": "ERRO",
                        "mensagem": f"Exceção: {str(ex_colab)}",
                        "timestamp_consulta": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    })
                    try:
                        self.bot_ativo.retornar_para_formulario()
                    except Exception:
                        pass

            # 4. Exportação do Relatório Excel Consolidado na pasta escolhida pelo usuário
            caminho_relatorio = None
            if resultados:
                caminho_relatorio = exportar_relatorio_extratos(resultados, diretorio_saida=pasta_saida)
                self.after(0, self._log, f"✔ Relatório oficial consolidado exportado com sucesso em:\n    {caminho_relatorio}")

            self.after(0, self._finalizar_execucao, caminho_relatorio, sucessos, alertas)

        except Exception as e:
            self.after(0, self._log, f"❌ Erro fatal na execução: {str(e)}")
            self.after(0, self._finalizar_execucao, None, 0, 0)
        finally:
            if self.bot_ativo:
                try:
                    self.bot_ativo.fechar_navegador()
                except Exception:
                    pass
                self.bot_ativo = None

    def _atualizar_progresso(self, progresso: float, idx: int, total: int, nome: str, pis: str):
        """Atualiza a barra e a legenda de progresso na interface."""
        self.progress_bar.set(progresso)
        pct = int(progresso * 100)
        self.lbl_progresso_status.configure(
            text=f"Processando {idx} de {total} ({pct}%) - PIS: {pis} | {nome}"
        )

    def _atualizar_contadores(self, sucessos: int, alertas: int):
        """Atualiza as contagens de sucesso e alertas nos badges."""
        self.badge_sucesso.configure(text=f"Sucesso: {sucessos}")
        self.badge_alertas.configure(text=f"Não Localizados / Erros: {alertas}")

    def _finalizar_execucao(self, caminho_relatorio: Optional[Path], sucessos: int, alertas: int):
        """Restaura o estado da interface após conclusão do processo."""
        self.processo_em_execucao = False
        self.btn_iniciar.configure(state="normal")
        self.btn_cancelar.configure(state="disabled")

        if caminho_relatorio and caminho_relatorio.exists():
            self.ultimo_relatorio_gerado = caminho_relatorio
            self.btn_abrir_relatorio.configure(state="normal")
            self.lbl_progresso_status.configure(
                text="✔ Processamento em lote finalizado com sucesso! Relatório e PDFs prontos.",
                text_color=COR_VERDE_SUCESSO
            )
            messagebox.showinfo(
                "Processamento Concluído",
                f"Consultas finalizadas com sucesso!\n\n"
                f"• Sucesso: {sucessos}\n"
                f"• Alertas / Não Localizados: {alertas}\n\n"
                f"O relatório Excel e todos os extratos em PDF oficiais estão salvos na pasta:\n{self.caminho_pasta_saida}"
            )
        else:
            self.lbl_progresso_status.configure(
                text="Processamento encerrado sem novo relatório gerado.",
                text_color="gray"
            )

    def _abrir_relatorio_excel(self):
        """Abre o arquivo Excel gerado diretamente no programa padrão (Excel)."""
        if self.ultimo_relatorio_gerado and self.ultimo_relatorio_gerado.exists():
            try:
                os.startfile(str(self.ultimo_relatorio_gerado))
            except Exception as e:
                messagebox.showerror("Erro ao Abrir", f"Não foi possível abrir o arquivo:\n{e}")
        else:
            messagebox.showwarning("Aviso", "Nenhum relatório gerado nesta sessão.")

    def _abrir_pasta_relatorios(self):
        """Abre a pasta onde os relatórios e extratos são armazenados no Windows Explorer."""
        try:
            os.startfile(str(self.caminho_pasta_saida))
        except Exception as e:
            messagebox.showerror("Erro ao Abrir Pasta", f"Não foi possível abrir o diretório:\n{e}")


def main():
    app = AppConectividadeSocial()
    app.mainloop()


if __name__ == "__main__":
    main()
