import os
import datetime
import re
import pandas as pd
import numpy as np
import openpyxl
import unicodedata
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import time
import traceback


class MotorRH_App:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador Automático de Relatório - Admissão V003")
        self.root.geometry("850x600")
        self.root.configure(bg="#f4f4f4")

        # Variáveis de Controle
        self.caminho_pasta = ""
        self.caminhos_arquivos = {}
        self.animando = False
        self.linha_atual = None
        self.inicio_passo = 0
        self.frames_animacao = ['|', '/', '-', '\\']
        self.frame_idx = 0

        self.criar_interface()

    def criar_interface(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TButton", padding=6, relief="flat", background="#0052cc", foreground="white",
                        font=("Segoe UI", 10, "bold"))
        style.map("TButton", background=[('active', '#003d99')])

        # Cabeçalho
        frame_top = tk.Frame(self.root, bg="#f4f4f4", pady=10)
        frame_top.pack(fill=tk.X)
        lbl_titulo = tk.Label(frame_top, text="🚀 Gerador de Relatório Admissão (Censo Master)",
                              font=("Segoe UI", 16, "bold"), bg="#f4f4f4", fg="#333333")
        lbl_titulo.pack()

        # Botões Principais
        frame_botoes = tk.Frame(self.root, bg="#f4f4f4")
        frame_botoes.pack(pady=10)

        self.btn_pasta = ttk.Button(frame_botoes, text="📁 1. Selecionar Pasta e Validar", command=self.selecionar_pasta)
        self.btn_pasta.grid(row=0, column=0, padx=10)

        self.btn_gerar = ttk.Button(frame_botoes, text="⚙️ 2. Gerar Relatório", state=tk.DISABLED,
                                    command=self.iniciar_geracao)
        self.btn_gerar.grid(row=0, column=1, padx=10)

        # Terminal / Log Visual
        frame_log = tk.Frame(self.root, padx=20, pady=10, bg="#f4f4f4")
        frame_log.pack(fill=tk.BOTH, expand=True)

        self.log_widget = tk.Text(frame_log, wrap=tk.WORD, bg="#1e1e1e", fg="#00ff00", font=("Consolas", 10),
                                  state=tk.DISABLED)
        self.log_widget.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        scroll = tk.Scrollbar(frame_log, command=self.log_widget.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_widget.config(yscrollcommand=scroll.set)

        self.log_widget.tag_config("erro_vermelho", foreground="#ff4d4d")

        self.print_log("Bem-vindo ao Motor Analítico de RH. Clique no botão 1 para começar.")

    # --- FUNÇÕES DE LOG E ANIMAÇÃO ---
    def print_log(self, texto):
        self.log_widget.config(state=tk.NORMAL)
        self.log_widget.insert(tk.END, f"{texto}\n")
        self.log_widget.see(tk.END)
        self.log_widget.config(state=tk.DISABLED)

    def iniciar_passo(self, texto):
        self.root.after(0, self._iniciar_passo_gui, texto)

    def _iniciar_passo_gui(self, texto):
        self.inicio_passo = time.time()
        self.log_widget.config(state=tk.NORMAL)
        self.log_widget.insert(tk.END, f"[ ] {texto}\n")
        index = self.log_widget.index("end-1c")
        self.linha_atual = int(float(index)) - 1
        self.log_widget.see(tk.END)
        self.log_widget.config(state=tk.DISABLED)
        self.animando = True
        self.atualizar_animacao()

    def atualizar_animacao(self):
        if self.animando and self.linha_atual:
            frame = self.frames_animacao[self.frame_idx]
            self.frame_idx = (self.frame_idx + 1) % len(self.frames_animacao)
            self.log_widget.config(state=tk.NORMAL)
            self.log_widget.delete(f"{self.linha_atual}.1", f"{self.linha_atual}.2")
            self.log_widget.insert(f"{self.linha_atual}.1", frame)
            self.log_widget.config(state=tk.DISABLED)
            self.root.after(150, self.atualizar_animacao)

    def concluir_passo(self, alerta=False):
        self.root.after(0, self._concluir_passo_gui, alerta)

    def _concluir_passo_gui(self, alerta):
        self.animando = False
        tempo = time.time() - self.inicio_passo
        icone = "❌" if alerta else "✔️"

        self.log_widget.config(state=tk.NORMAL)
        self.log_widget.delete(f"{self.linha_atual}.1", f"{self.linha_atual}.2")
        self.log_widget.insert(f"{self.linha_atual}.1", icone)
        self.log_widget.insert(f"{self.linha_atual}.end", f" ({tempo:.2f}s)")

        if alerta:
            self.log_widget.tag_add("erro_vermelho", f"{self.linha_atual}.0", f"{self.linha_atual}.end")

        self.log_widget.config(state=tk.DISABLED)

    # --- LÓGICA DO NEGÓCIO ---
    def normalizar(self, texto):
        sem_acento = ''.join(c for c in unicodedata.normalize('NFD', str(texto)) if unicodedata.category(c) != 'Mn')
        return re.sub(r'[^A-Z0-9]', '', sem_acento.upper())

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com as planilhas do mês")
        if not pasta:
            return

        self.caminho_pasta = pasta
        self.log_widget.config(state=tk.NORMAL)
        self.log_widget.delete(1.0, tk.END)
        self.log_widget.config(state=tk.DISABLED)

        self.print_log(f"🔍 Buscando arquivos automaticamente na pasta:\n📁 {pasta}\n")

        # Adicionado o prefixo 'estab' da nova base
        PREFIXOS_ARQUIVOS = {
            'colab': "01.Relação de Funcionários",
            'dep': "COLA_Dependentes_por_Colaborador_CSV",
            'info_adc': "COLA_Informacao_Adicional_Colaborador_CSV",
            'perfil': "Escala_e_Perfil_de_Frequência_Atual_do_Colaborador",
            'eventos_calc': "FOPA_Eventos_Calculados_Conta_Contabil",
            'eventos': "FOPA_Eventos_por_Colaborador",
            'lojas_suspensas': "Suspensão de Refeitorio lojas",
            'bate_ponto': "VALIDACAO DE BATE PONTO OU NAO",
            'cracha': "EMISSAO_CRACHA",
            'inss_incidencia': "RELATORIO_EVENTOS_COM_INCIDENCIA_INSS",
            'estab': "ORGA_Estabelecimento_CSV"  # NOVO ARQUIVO DE ESTADOS
        }

        self.caminhos_arquivos = {}
        todos_encontrados = True

        for chave, prefixo in PREFIXOS_ARQUIVOS.items():
            self.iniciar_passo(f"Buscando {prefixo}...")
            prefixo_norm = self.normalizar(prefixo)
            encontrado = None

            for arquivo in os.listdir(pasta):
                if self.normalizar(arquivo).startswith(prefixo_norm):
                    encontrado = os.path.join(pasta, arquivo)
                    break

            if encontrado:
                self.caminhos_arquivos[chave] = encontrado
                self.concluir_passo()
            else:
                self.concluir_passo(alerta=True)
                todos_encontrados = False

        self.print_log("-" * 50)
        if todos_encontrados:
            self.print_log("✅ Todos os arquivos necessários foram localizados!")
            self.btn_gerar.config(state=tk.NORMAL)
        else:
            self.print_log("⛔ ERRO: Faltam planilhas essenciais! A geração do relatório foi bloqueada.")
            self.btn_gerar.config(state=tk.DISABLED)  # <--- Botão TRAVADO!

            self.log_widget.config(state=tk.NORMAL)
            ultima_linha = int(float(self.log_widget.index("end-1c"))) - 1
            self.log_widget.tag_add("erro_vermelho", f"{ultima_linha}.0", tk.END)
            self.log_widget.config(state=tk.DISABLED)

    def iniciar_geracao(self):
        self.btn_pasta.config(state=tk.DISABLED)
        self.btn_gerar.config(state=tk.DISABLED)
        self.print_log("\n🚀 INICIANDO O MOTOR ANALÍTICO GLOBAL DE RH...\n")
        threading.Thread(target=self.processamento_pesado, daemon=True).start()

    def processamento_pesado(self):
        try:
            # =========================================================================
            # FUNÇÕES AUXILIARES EXATAMENTE IGUAIS AO CÓDIGO HOMOLOGADO
            # =========================================================================
            def ler_arquivo_dinamico(caminho):
                if not caminho:
                    return pd.DataFrame()

                if caminho.lower().endswith('.xlsx') or caminho.lower().endswith('.xls'):
                    return pd.read_excel(caminho)

                encodings = ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']
                linhas = []
                enc_usado = 'utf-8'

                for enc in encodings:
                    try:
                        with open(caminho, 'r', encoding=enc) as f:
                            linhas = f.readlines()
                        enc_usado = enc
                        break
                    except UnicodeDecodeError:
                        continue

                if not linhas:
                    enc_usado = 'latin1'
                    with open(caminho, 'r', encoding=enc_usado, errors='ignore') as f:
                        linhas = f.readlines()

                header_idx = 0
                keywords = ['MATR', 'NOME', 'CPF', 'CÓDIGO', 'EMPRESA', 'EVENTO']
                max_matches = 0

                for i, linha in enumerate(linhas[:30]):
                    linha_upper = linha.upper()
                    matches = sum(1 for k in keywords if k in linha_upper)
                    delimitadores = max(linha.count(';'), linha.count(','))

                    if matches > max_matches and delimitadores > 3:
                        max_matches = matches
                        header_idx = i

                sep = ';' if linhas and ';' in linhas[header_idx] else ','
                return pd.read_csv(caminho, sep=sep, encoding=enc_usado, skiprows=header_idx, low_memory=False)

            def limpar_chave_matricula(coluna):
                s = coluna.fillna('').astype(str).str.strip().str.upper()
                s = s.str.replace(r'\.0$', '', regex=True)
                s = s.str.lstrip('0')
                s = s.replace(['', 'NAN', 'NONE'], '0')
                return s

            def extrair_hora_pura(val):
                if pd.isna(val): return 0.0
                val_str = str(val).upper().replace(',', '.').strip()
                if ':' in val_str: val_str = val_str.split(':')[0]
                match = re.search(r'(\d+(?:\.\d+)?)', val_str)
                if match: return float(match.group(1))
                return 0.0

            def limpar_moeda(valor):
                v_str = str(valor).strip()
                if v_str in ['-', '', 'nan', 'None', 'NaN', '<NA>']: return 0.00
                if ',' in v_str: v_str = v_str.replace('.', '').replace(',', '.')
                try:
                    return float(v_str)
                except ValueError:
                    return 0.00

            def extrair_codigo_evento(valor):
                if pd.isna(valor): return "0"
                try:
                    return str(int(float(valor)))
                except (ValueError, TypeError):
                    match = re.search(r'(\d+)', str(valor))
                    return str(int(match.group(1))) if match else "0"

            def get_col(df, palavras_chave, indice_fallback):
                for k in palavras_chave:
                    for c in df.columns:
                        if k in str(c).upper(): return c
                if indice_fallback < len(df.columns): return df.columns[indice_fallback]
                return df.columns[-1]

            def extract_safe(df, idx):
                if idx < len(df.columns): return df.iloc[:, idx].copy()
                return pd.Series('', index=df.index)

            # =========================================================================
            # 2. CARREGAMENTO DAS BASES
            # =========================================================================
            self.iniciar_passo("Carregando as bases de dados (Caçador de Cabeçalhos Ativado)...")
            df_colab = ler_arquivo_dinamico(self.caminhos_arquivos.get('colab'))
            df_dep = ler_arquivo_dinamico(self.caminhos_arquivos.get('dep'))
            df_info_adc = ler_arquivo_dinamico(self.caminhos_arquivos.get('info_adc'))
            df_perfil = ler_arquivo_dinamico(self.caminhos_arquivos.get('perfil'))
            df_eventos_calc = ler_arquivo_dinamico(self.caminhos_arquivos.get('eventos_calc'))
            df_eventos = ler_arquivo_dinamico(self.caminhos_arquivos.get('eventos'))
            df_lojas_susp = ler_arquivo_dinamico(self.caminhos_arquivos.get('lojas_suspensas'))
            df_bate_ponto = ler_arquivo_dinamico(self.caminhos_arquivos.get('bate_ponto'))
            df_cracha = ler_arquivo_dinamico(self.caminhos_arquivos.get('cracha'))
            df_inss_incidencia = ler_arquivo_dinamico(self.caminhos_arquivos.get('inss_incidencia'))
            df_estab = ler_arquivo_dinamico(self.caminhos_arquivos.get('estab'))  # NOVA BASE DE ESTABELECIMENTOS
            self.concluir_passo()

            # =========================================================================
            # 3. TRANSFORMAÇÃO E CAPTURA LITERAL BLINDADA
            # =========================================================================
            self.iniciar_passo("Iniciando Transformação Literal de Dados (A-F)...")

            # --- A. TRATAMENTO COLABORADORES ---
            df_colab['COD_EMPRESA_SEGURO'] = extract_safe(df_colab, 0)
            df_colab['MATRICULA_MASTER'] = extract_safe(df_colab, 1)
            df_colab['NOME_PESSOA_SEGURO'] = extract_safe(df_colab, 3)
            df_colab['DATA_ADMISSAO_ORIGINAL'] = extract_safe(df_colab, 4)
            df_colab['SITUACAO_SEGURO'] = extract_safe(df_colab, 6)
            df_colab['CPF_SEGURO'] = extract_safe(df_colab, 7)
            df_colab['NOME_CARGO_SEGURO'] = extract_safe(df_colab, 65)
            df_colab['COD_CC_SEGURO'] = extract_safe(df_colab, 66)
            df_colab['COD_ESTAB_SEGURO'] = extract_safe(df_colab, 68)
            df_colab['COD_UO_SEGURO'] = extract_safe(df_colab, 72).astype(str).str.replace(r'\.0$', '',
                                                                                           regex=True).str.strip() + " - " + extract_safe(
                df_colab, 73).astype(str)
            df_colab['CÓDIGO_ESCALA_SEGURO'] = extract_safe(df_colab, 79)
            df_colab['DESCRIÇÃO_ESCALA'] = extract_safe(df_colab, 80)
            df_colab['CARGA_HORÁRIA_SEMANAL_SEGURO'] = extract_safe(df_colab, 81)
            df_colab['CARGA_HORÁRIA_MENSAL_SEGURO'] = extract_safe(df_colab, 82)
            df_colab['COD_BANCO'] = extract_safe(df_colab, 85)
            df_colab['AGENCIA_BANCO_RAW'] = extract_safe(df_colab, 86)
            df_colab['CONTA_BANCO_RAW'] = extract_safe(df_colab, 88)
            df_colab['VINCULO_SEGURO'] = extract_safe(df_colab, 91)
            df_colab['CATEGORIA_SEGURO'] = extract_safe(df_colab, 93)
            df_colab['REGIME_SEGURO'] = extract_safe(df_colab, 94)
            df_colab['NATUREZA_SEGURO'] = extract_safe(df_colab, 99).astype(str).str.replace(r'\.0$', '',
                                                                                             regex=True).str.strip() + " - " + extract_safe(
                df_colab, 100).astype(str)

            col_sindicato = get_col(df_colab, ['SINDICATO', 'SINDICAL'], 96)
            col_salario = get_col(df_colab, ['SALÁRIO', 'SALARIO', 'SAL.BASE', 'REMUNERAÇÃO'], 97)
            df_colab['SINDICATO_SEGURO'] = df_colab[col_sindicato].astype(str).str.strip().replace(
                ['nan', 'None', 'NaN', ''], 'NÃO INFORMADO')
            df_colab['SALARIO_BASE_SEGURO'] = df_colab[col_salario].apply(limpar_moeda)

            df_colab['NUMERO_LOJA_SEGURO'] = extract_safe(df_colab, 68).astype(str).str.replace(r'\.0$', '',
                                                                                                regex=True).str.strip()

            # --- NOVA EXTRAÇÃO: ESTADOS DA NOVA PLANILHA (Índices 2 e 10) ---
            if not df_estab.empty:
                col_cod_estab_novo = get_col(df_estab, ['CÓDIGO ESTABELECIMENTO', 'CODIGO ESTABELECIMENTO'], 2)
                col_estado_novo = get_col(df_estab, ['ESTADO'], 10)

                df_estab['COD_ESTAB_MAP'] = df_estab[col_cod_estab_novo].fillna('').astype(str).str.replace(r'\.0$', '',
                                                                                                            regex=True).str.strip()
                df_estab['ESTADO_MAP'] = df_estab[col_estado_novo].fillna('').astype(str).str.strip().str.upper()

                mapa_estado = df_estab.groupby('COD_ESTAB_MAP')['ESTADO_MAP'].first()
                # Cruza o Nº DA LOJA com o código de estabelecimento e puxa o Estado
                df_colab['ESTADO_CORRETO'] = df_colab['NUMERO_LOJA_SEGURO'].map(mapa_estado)

                # Fallback: Se por acaso a loja não existir na planilha nova, ele usa a antiga coluna 30
                df_colab['ESTADO_CORRETO'] = df_colab['ESTADO_CORRETO'].fillna(extract_safe(df_colab, 30))
            else:
                df_colab['ESTADO_CORRETO'] = extract_safe(df_colab, 30)

            df_colab['BANDEIRA_RAW'] = extract_safe(df_colab, 100).astype(str).str.strip().replace(['nan', 'None', ''],
                                                                                                   np.nan)
            mapa_loja_bandeira = (
                df_colab[df_colab['BANDEIRA_RAW'].notna()].groupby('NUMERO_LOJA_SEGURO')['BANDEIRA_RAW'].agg(
                    lambda s: s.mode().iloc[0] if not s.mode().empty else np.nan))
            df_colab['BANDEIRA_SEGURO'] = df_colab['BANDEIRA_RAW'].fillna(
                df_colab['NUMERO_LOJA_SEGURO'].map(mapa_loja_bandeira))
            df_colab['BANDEIRA_SEGURO'] = df_colab['BANDEIRA_SEGURO'].fillna('SEM INFORMAÇÃO').astype(
                str).str.upper().str.strip()

            df_colab['DATA_ADMISSAO_DT'] = pd.to_datetime(df_colab['DATA_ADMISSAO_ORIGINAL'], errors='coerce',
                                                          dayfirst=True)
            df_colab['DATA_ADMISSAO_FORMATADA'] = df_colab['DATA_ADMISSAO_DT'].dt.strftime('%d/%m/%Y')
            df_colab = df_colab[df_colab['DATA_ADMISSAO_FORMATADA'] != '24/06/2026'].copy()

            df_colab['CONTA_BANCO'] = df_colab['CONTA_BANCO_RAW'].fillna('').astype(str).str.replace(r'\.0$', '',
                                                                                                     regex=True).str.strip()
            df_colab['AGENCIA_BANCO'] = df_colab['AGENCIA_BANCO_RAW'].fillna('').astype(str).str.replace(r'\.0$', '',
                                                                                                         regex=True).str.strip()
            banco_limpo_dup = df_colab['COD_BANCO'].fillna('').astype(str).str.replace(r'\.0$', '',
                                                                                       regex=True).str.strip()
            conta_invalida = df_colab['CONTA_BANCO'].isin(
                ['', '0', '00', '000', '0000', '1', 'NAN', 'NONE', 'NULL', '-'])
            banco_invalido = banco_limpo_dup.isin(['', '0', 'NAN', 'NONE', 'NULL', '-'])
            agencia_invalida = df_colab['AGENCIA_BANCO'].isin(['', '0', 'NAN', 'NONE', 'NULL', '-'])
            chave_conta_completa = banco_limpo_dup + "_" + df_colab['AGENCIA_BANCO'] + "_" + df_colab['CONTA_BANCO']
            chave_conta_valida = chave_conta_completa.where(~(conta_invalida | banco_invalido | agencia_invalida),
                                                            np.nan)
            cpfs_por_conta = df_colab.groupby(chave_conta_valida)['CPF_SEGURO'].transform('nunique')
            df_colab['ERRO_BANCO_DUPLICADO'] = cpfs_por_conta > 1
            df_colab['ESTA_NA_PLANILHA_BANCO'] = df_colab['COD_BANCO'].fillna('').astype(str).str.strip() != ''

            # --- B. DEPENDENTES ---
            col_mat_dep = get_col(df_dep, ['MATR'], 3)
            col_nasc_dep = get_col(df_dep, ['NASCIMENTO', 'DATA'], 21)
            col_par_dep = get_col(df_dep, ['PARENTESCO'], 24)
            col_nome_dep = next(
                (c for c in df_dep.columns if 'NOME' in str(c).upper() and 'DEPENDENTE' in str(c).upper()), None)
            if col_nome_dep is None: col_nome_dep = get_col(df_dep, ['NOME'], 5)
            df_dep[col_mat_dep] = limpar_chave_matricula(df_dep[col_mat_dep])
            df_dep['ID DEPENDENTE'] = np.arange(1, len(df_dep) + 1)
            df_dep['CÓDIGO DEPENDENTE'] = df_dep[col_mat_dep].astype(str) + "-" + df_dep['ID DEPENDENTE'].astype(str)
            df_dep['NOME DEPENDENTE'] = df_dep[col_nome_dep]
            hoje = pd.Timestamp(datetime.datetime.now().date())
            df_dep[col_nasc_dep] = pd.to_datetime(df_dep[col_nasc_dep], errors='coerce', dayfirst=True)
            aniversario_ja_ocorreu = (df_dep[col_nasc_dep].dt.month < hoje.month) | (
                        (df_dep[col_nasc_dep].dt.month == hoje.month) & (df_dep[col_nasc_dep].dt.day <= hoje.day))
            idade_exata = (hoje.year - df_dep[col_nasc_dep].dt.year) - (~aniversario_ja_ocorreu).astype(int)
            df_dep['IDADE DEPENDENTE'] = idade_exata.where(df_dep[col_nasc_dep].notna())
            df_dep['PARENTESCO_CLEAN'] = df_dep[col_par_dep].fillna('').astype(str).str.upper().str.strip()
            filhos_menores = df_dep[
                (df_dep['PARENTESCO_CLEAN'].str.contains('FILHO|FILHA|GUARDA|ENTEADO', na=False)) & (
                            df_dep['IDADE DEPENDENTE'] < 14)]
            matriculas_com_filho_menor = set(limpar_chave_matricula(filhos_menores[col_mat_dep]))

            # --- C. PERFIL DE FREQUÊNCIA ---
            col_mat_perf = get_col(df_perfil, ['MATR'], 4)
            col_esc_perf = get_col(df_perfil, ['CODIGO ESCALA', 'ESCALA'], 6)
            col_sem_perf = get_col(df_perfil, ['CH SEMANAL', 'SEMANAL'], 8)
            col_men_perf = get_col(df_perfil, ['CH MENSAL', 'MENSAL'], 9)
            col_sit_perf = get_col(df_perfil, ['PERFIL', 'FREQUENCIA'], 11)
            df_perfil[col_mat_perf] = limpar_chave_matricula(df_perfil[col_mat_perf])
            df_perf_val = df_perfil[df_perfil[col_mat_perf] != '0'].copy()
            df_perf_val[col_sem_perf] = df_perf_val[col_sem_perf].apply(extrair_hora_pura)
            df_perf_val[col_men_perf] = df_perf_val[col_men_perf].apply(extrair_hora_pura)
            df_perfil_clean = df_perf_val.groupby(col_mat_perf).agg(
                {col_sit_perf: 'first', col_esc_perf: 'first', col_sem_perf: 'first',
                 col_men_perf: 'first'}).reset_index()
            df_perfil_clean.rename(columns={col_sit_perf: 'SITUACAO DE PONTO SISTEMA', col_esc_perf: 'ESCALA_GABARITO',
                                            col_sem_perf: 'CH_SEMANAL_GABARITO', col_men_perf: 'CH_MENSAL_GABARITO',
                                            col_mat_perf: 'MATR_PERFIL'}, inplace=True)

            # --- D. BATE PONTO DINÂMICO ---
            df_bate_ponto.columns = df_bate_ponto.columns.astype(str).str.strip().str.upper()
            col_cargo_bp = df_bate_ponto.columns[0]
            col_regra_bp = df_bate_ponto.columns[1]
            df_bate_ponto[col_cargo_bp] = df_bate_ponto[col_cargo_bp].fillna('').astype(str).str.upper().str.strip()
            df_bate_ponto_clean = df_bate_ponto.groupby(col_cargo_bp)[col_regra_bp].first().reset_index()
            df_bate_ponto_clean.rename(columns={col_regra_bp: 'REGRA_BATE_PONTO', col_cargo_bp: 'CARGO_BP'},
                                       inplace=True)

            # --- E. OPÇÕES DO REFEITÓRIO ---
            col_mat_inf = get_col(df_info_adc, ['MATR'], 2)
            col_desc_inf = get_col(df_info_adc, ['DESCRI', 'INFORMA'], 5)
            col_val_inf = get_col(df_info_adc, ['VALOR', 'CONTEÚDO', 'OPÇÃO'], 6)
            df_info_adc[col_mat_inf] = limpar_chave_matricula(df_info_adc[col_mat_inf])
            df_ref_opt = \
            df_info_adc[df_info_adc[col_desc_inf].astype(str).str.upper().str.contains('REFEIT', na=False)].groupby(
                col_mat_inf)[col_val_inf].first().reset_index()
            df_ref_opt.rename(columns={col_val_inf: 'OPCAO_REFEITORIO', col_mat_inf: 'MATR_REF'}, inplace=True)

            # --- F. LISTA NEGRA DE LOJAS DO REFEITÓRIO ---
            col_loja_susp = df_lojas_susp.columns[1]
            lojas_suspensas_set = set(
                df_lojas_susp[col_loja_susp].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip())
            lojas_suspensas_set.discard('')
            self.concluir_passo()

            # --- G. EXTRAÇÃO: BASE INSS ---
            self.iniciar_passo("Somando Base de Cálculo Previdenciária (INSS)...")
            col_mat_calc = get_col(df_eventos_calc, ['MATR'], 2)
            col_ev_calc = get_col(df_eventos_calc, ['CÓDIGO EVENTO', 'COD EVENTO', 'CÓD. EVENTO', 'EVENTO', 'CÓD'], 23)
            col_val_calc = get_col(df_eventos_calc, ['VALOR', 'VLR'], 25)

            if not df_inss_incidencia.empty:
                col_cod_inss_oficial = get_col(df_inss_incidencia,
                                               ['CÓDIGO DO EVENTO', 'CODIGO DO EVENTO', 'CÓD. EVENTO', 'CÓD'], 1)
                col_flag_inss_oficial = get_col(df_inss_incidencia, ['INSS'], 12)
                df_inss_incidencia['COD_CLEAN'] = df_inss_incidencia[col_cod_inss_oficial].apply(extrair_codigo_evento)
                flag_inss_upper = df_inss_incidencia[col_flag_inss_oficial].fillna('').astype(
                    str).str.strip().str.upper()
                codigos_inss = set(df_inss_incidencia.loc[flag_inss_upper == 'S', 'COD_CLEAN'])
            else:
                codigos_inss = {str(c) for c in
                                [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 27, 28, 29, 30, 31, 32, 33,
                                 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55,
                                 56, 57, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 78, 79, 80, 81,
                                 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 11010]}

            df_eventos_calc['COD_CLEAN'] = df_eventos_calc[col_ev_calc].apply(extrair_codigo_evento)
            df_inss = df_eventos_calc[df_eventos_calc['COD_CLEAN'].isin(codigos_inss)].copy()
            df_inss['VAL'] = df_inss[col_val_calc].apply(limpar_moeda)
            df_inss[col_mat_calc] = limpar_chave_matricula(df_inss[col_mat_calc])

            df_base_inss = df_inss.groupby(col_mat_calc)['VAL'].sum().reset_index()
            df_base_inss.rename(columns={'VAL': 'BASE_INSS_CALCULADA', col_mat_calc: 'MATR_INSS'}, inplace=True)
            self.concluir_passo()

            # --- H. EXTRAÇÃO: CONTRACHEQUE REAL ---
            self.iniciar_passo("Lendo Contracheque Real (Salário Família e Teletrabalho)...")
            col_mat_ev = get_col(df_eventos, ['MATR'], 6)
            df_eventos[col_mat_ev] = limpar_chave_matricula(df_eventos[col_mat_ev])

            col_110 = next((c for c in df_eventos.columns if '110' in str(c) or 'AJUDA' in str(c).upper()), None)
            col_126 = next((c for c in df_eventos.columns if '126' in str(c) or 'FAMILIA' in str(c).upper()), None)
            col_190 = next((c for c in df_eventos.columns if '190' in str(c) or 'MOBILIA' in str(c).upper()), None)

            colunas_nomeadas_encontradas = col_110 is not None or col_126 is not None or col_190 is not None
            if not colunas_nomeadas_encontradas and (
                    'CÓDIGO' in ''.join([str(c).upper() for c in df_eventos.columns]) or 'EVENTO' in ''.join(
                    [str(c).upper() for c in df_eventos.columns])):
                c_mat_t = get_col(df_eventos, ['MATR'], 0)
                c_ev_t = get_col(df_eventos, ['EVENTO', 'CÓD'], 1)
                c_val_t = get_col(df_eventos, ['VALOR'], 2)

                df_eventos['COD_T'] = df_eventos[c_ev_t].apply(extrair_codigo_evento)
                df_eventos['VAL_T'] = df_eventos[c_val_t].apply(limpar_moeda)
                df_eventos[c_mat_t] = limpar_chave_matricula(df_eventos[c_mat_t])

                df_ev_110 = df_eventos[df_eventos['COD_T'] == '110'].groupby(c_mat_t)[
                    'VAL_T'].sum().reset_index().rename(columns={'VAL_T': 'V110', c_mat_t: 'M'})
                df_ev_126 = df_eventos[df_eventos['COD_T'].isin(['126', '14009'])].groupby(c_mat_t)[
                    'VAL_T'].sum().reset_index().rename(columns={'VAL_T': 'V126', c_mat_t: 'M'})
                df_ev_190 = df_eventos[df_eventos['COD_T'] == '190'].groupby(c_mat_t)[
                    'VAL_T'].sum().reset_index().rename(columns={'VAL_T': 'V190', c_mat_t: 'M'})

                df_ev_clean = pd.DataFrame({col_mat_ev: df_eventos[c_mat_t].unique()})
                df_ev_clean = pd.merge(df_ev_clean, df_ev_110, left_on=col_mat_ev, right_on='M', how='left').drop(
                    columns=['M'])
                df_ev_clean = pd.merge(df_ev_clean, df_ev_126, left_on=col_mat_ev, right_on='M', how='left').drop(
                    columns=['M'])
                df_ev_clean = pd.merge(df_ev_clean, df_ev_190, left_on=col_mat_ev, right_on='M', how='left').drop(
                    columns=['M'])
                df_ev_clean = df_ev_clean.fillna(0.00)
                df_ev_clean.rename(columns={col_mat_ev: 'MATR_EV'}, inplace=True)
            else:
                df_eventos['V110'] = df_eventos[col_110].apply(limpar_moeda) if col_110 else 0.0
                df_eventos['V126'] = df_eventos[col_126].apply(limpar_moeda) if col_126 else 0.0
                df_eventos['V190'] = df_eventos[col_190].apply(limpar_moeda) if col_190 else 0.0
                df_ev_clean = df_eventos.groupby(col_mat_ev).agg(
                    {'V110': 'sum', 'V126': 'sum', 'V190': 'sum'}).reset_index()
                df_ev_clean.rename(
                    columns={'V110': 'VALOR_110_TELE', 'V126': '126 - SALARIO FAMILIA', 'V190': 'VALOR_190_TELE',
                             col_mat_ev: 'MATR_EV'}, inplace=True)
            self.concluir_passo()

            # --- I. EXTRAÇÃO: CRACHÁ ---
            self.iniciar_passo("Processando Planilha Mestre de Crachás...")
            df_cracha['EMPRESA_CRACHA'] = extract_safe(df_cracha, 0)
            df_cracha['LOJA_CRACHA'] = extract_safe(df_cracha, 2)
            df_cracha['NOME_COMPLETO_CRACHA'] = extract_safe(df_cracha, 4)
            df_cracha['MATRICULA_CRACHA'] = limpar_chave_matricula(extract_safe(df_cracha, 5))
            df_cracha['NOME_CRACHA_RAW'] = extract_safe(df_cracha, 6)
            df_cracha['BANDEIRA_CRACHA'] = extract_safe(df_cracha, 7)
            df_cracha_clean = df_cracha.groupby('MATRICULA_CRACHA').first().reset_index()
            self.concluir_passo()

            # =========================================================================
            # 4. MERGE GLOBAL
            # =========================================================================
            self.iniciar_passo("Realizando Merge Global de todas as bases...")
            df_colab['MATRICULA_MASTER'] = limpar_chave_matricula(df_colab['MATRICULA_MASTER'])
            df_colab['CÓDIGO_ESCALA_LIMPO'] = limpar_chave_matricula(df_colab['CÓDIGO_ESCALA_SEGURO'])

            df_consolidado = pd.merge(df_colab, df_dep, left_on='MATRICULA_MASTER', right_on=col_mat_dep, how='left')
            df_consolidado = pd.merge(df_consolidado, df_base_inss, left_on='MATRICULA_MASTER', right_on='MATR_INSS',
                                      how='left')
            df_consolidado = pd.merge(df_consolidado, df_ev_clean, left_on='MATRICULA_MASTER', right_on='MATR_EV',
                                      how='left')
            df_consolidado = pd.merge(df_consolidado, df_perfil_clean, left_on='MATRICULA_MASTER',
                                      right_on='MATR_PERFIL', how='left')
            df_consolidado = pd.merge(df_consolidado, df_ref_opt, left_on='MATRICULA_MASTER', right_on='MATR_REF',
                                      how='left')
            df_consolidado = pd.merge(df_consolidado, df_cracha_clean, left_on='MATRICULA_MASTER',
                                      right_on='MATRICULA_CRACHA', how='left')

            df_consolidado['FLAG_FILHO_MENOR_14'] = df_consolidado['MATRICULA_MASTER'].isin(matriculas_com_filho_menor)
            df_consolidado['FLAG_TEM_REGISTRO_CRACHA'] = df_consolidado['MATRICULA_CRACHA'].notna()
            df_consolidado['126 - SALARIO FAMILIA'] = df_consolidado['126 - SALARIO FAMILIA'].fillna(0.00)
            df_consolidado['VALOR_110_TELE'] = df_consolidado['VALOR_110_TELE'].fillna(0.00)
            df_consolidado['VALOR_190_TELE'] = df_consolidado['VALOR_190_TELE'].fillna(0.00)
            df_consolidado['BASE SAL FAMILIA'] = 1980.38
            df_consolidado['PREFIXO_CC'] = df_consolidado['COD_CC_SEGURO'].fillna('').astype(str).str.replace(r'\.0$',
                                                                                                              '',
                                                                                                              regex=True).str.strip().str[
                :3]

            df_consolidado['BASE_INSS_CALCULADA'] = df_consolidado['BASE_INSS_CALCULADA'].fillna(0.00)
            df_consolidado['SALARIO_BASE_SEGURO'] = df_consolidado['SALARIO_BASE_SEGURO'].fillna(0.00)
            df_consolidado['BASE_INSS_CALCULADA'] = np.where(
                df_consolidado['BASE_INSS_CALCULADA'] == 0,
                df_consolidado['SALARIO_BASE_SEGURO'],
                df_consolidado['BASE_INSS_CALCULADA']
            )
            self.concluir_passo()

            # =========================================================================
            # 5. AUDITORIAS
            # =========================================================================
            self.iniciar_passo("Executando as Auditorias Trabalhistas (Intermitentes, Ponto, Escala...)...")
            cod_estab = df_consolidado['COD_ESTAB_SEGURO'].fillna('').astype(str).str.replace(r'\.0$', '',
                                                                                              regex=True).str.strip()

            escala_limpa = df_consolidado['CÓDIGO_ESCALA_SEGURO'].fillna('').astype(str).str.replace(r'\.0$', '',
                                                                                                     regex=True).str.strip()
            categoria_limpa = df_consolidado['CATEGORIA_SEGURO'].fillna('').astype(str).str.replace(r'\.0$', '',
                                                                                                    regex=True).str.strip()
            cargo_limpo_up = df_consolidado['NOME_CARGO_SEGURO'].fillna('').astype(str).str.upper().str.strip()

            escala_correta = escala_limpa.isin(['5005', '5006', '5007', '5008', '5009'])
            categoria_correta = (categoria_limpa == '111')
            cond_sistema_intermitente = escala_correta & categoria_correta
            contem_texto_intermitente = cargo_limpo_up.str.contains('INTERMITENTE', na=False)

            msg_erro_escala = pd.Series(
                np.where(~escala_correta, "Escala (" + escala_limpa + ") incorreta; ", "")).astype(str)
            msg_erro_cat = pd.Series(
                np.where(~categoria_correta, "Cat. (" + categoria_limpa + ") dif. de 111", "")).astype(str)
            motivo_erro_dinamico = (msg_erro_escala + msg_erro_cat).str.strip().replace(r'^\s*$',
                                                                                        'Parâmetros Incorretos',
                                                                                        regex=True)

            df_consolidado['VALIDACAO INTERMITENTE'] = np.select(
                [(cond_sistema_intermitente) & (~contem_texto_intermitente),
                 (~cond_sistema_intermitente) & (contem_texto_intermitente)],
                ['REVISAR CADASTRO (Cargo sem termo Intermitente)', 'REVISAR CADASTRO (' + motivo_erro_dinamico + ')'],
                default='OK'
            )

            nat_prof = df_consolidado['NATUREZA_SEGURO'].fillna('').astype(str).str.upper()
            estab_elegivel_gb = cod_estab.isin(['1', '6', '46'])
            estab_na_lista_negra = cod_estab.isin(lojas_suspensas_set)
            opcao_ref = df_consolidado['OPCAO_REFEITORIO'].fillna('').astype(str).str.upper().str.strip()

            def validar_refeitorio(row):
                # Utiliza o Estado Corrigido da nova planilha para validar
                estado = str(row.get('ESTADO_CORRETO', '')).strip().upper()
                cat = str(row.get('CATEGORIA_SEGURO', '')).replace('.0', '').strip()
                cargo = str(row.get('NOME_CARGO_SEGURO', '')).upper()
                opcao = str(row.get('OPCAO_REFEITORIO', '')).upper().strip()
                loja = str(row.get('COD_ESTAB_SEGURO', '')).replace('.0', '').strip()
                natureza = str(row.get('NATUREZA_SEGURO', '')).upper()

                is_intermitente = (cat == '111') or ('INTERMITENTE' in cargo)
                is_gb = 'GBARBOSA' in natureza
                loja_elegivel = loja in ['1', '6', '46']
                loja_suspensa = loja in lojas_suspensas_set

                if estado not in ['SE', 'SERGIPE']: return 'TEM DIREITO'
                if is_intermitente: return 'AJUSTAR (Não tem direito - Intermitente, mas optou SIM)' if opcao == 'SIM' else 'NÃO TEM DIREITO (INTERMITENTE)'
                if is_gb and (
                        not loja_elegivel or loja_suspensa): return 'AJUSTAR (Não tem direito - Loja, mas optou SIM)' if opcao == 'SIM' else 'NÃO TEM DIREITO (LOJA)'
                return 'TEM DIREITO'

            df_consolidado['VALIDACAO REFEITORIO'] = df_consolidado.apply(validar_refeitorio, axis=1)
            df_consolidado['OPÇÃO REFEITÓRIO'] = opcao_ref

            def validar_teletrabalho(linha):
                v110 = float(linha.get('VALOR_110_TELE', 0.0))
                v190 = float(linha.get('VALOR_190_TELE', 0.0))
                sit = str(linha.get('SITUACAO_SEGURO', '')).strip().upper()

                if v110 > 0 or v190 > 0:
                    if sit not in ['ATIVIDADE NORMAL', 'FÉRIAS',
                                   'FERIAS']: return f'REVISAR (Recebeu rubrica, mas está c/ situação: {sit})'
                    return 'OK (Possui Contrato)'
                return 'NÃO RECEBE'

            df_consolidado['VALIDACAO TELETRABALHO'] = df_consolidado.apply(validar_teletrabalho, axis=1)

            df_consolidado['NOME_CARGO_LIMPO'] = df_consolidado['NOME_CARGO_SEGURO'].fillna('').astype(
                str).str.upper().str.strip()
            df_consolidado = pd.merge(df_consolidado, df_bate_ponto_clean, left_on='NOME_CARGO_LIMPO',
                                      right_on='CARGO_BP', how='left')
            df_consolidado['PADRAO INTERNO-REGISTROXCARGO'] = df_consolidado['REGRA_BATE_PONTO'].fillna(
                'CARGO NÃO IDENTIFICADO NA PLANILHA')
            df_consolidado['SITUACAO DE PONTO SISTEMA'] = df_consolidado['SITUACAO DE PONTO SISTEMA'].fillna(
                'NÃO ENCONTRADO')

            def auditar_ponto_mestre(linha):
                situacao_colab = str(linha.get('SITUACAO_SEGURO', '')).strip().upper()
                if any(x in situacao_colab for x in ['FÉRIAS', 'FERIAS', 'AFASTAMENTO', 'AFASTADO']):
                    return "OK (Dispensado do registro por Férias/Afastamento)"

                padrao = str(linha.get('PADRAO INTERNO-REGISTROXCARGO', '')).upper()
                situacao_ponto = str(linha.get('SITUACAO DE PONTO SISTEMA', '')).upper()
                escala_func = str(linha.get('CÓDIGO_ESCALA_LIMPO', ''))
                carga_func = extrair_hora_pura(linha.get('CARGA_HORÁRIA_MENSAL_SEGURO', 0))
                carga_gab = linha.get('CH_MENSAL_GABARITO', None)

                erros = []
                isento_padrao = 'ISENTO' in padrao
                isento_situacao = 'ISENTO' in situacao_ponto

                if "NÃO ENCONTRADO" in situacao_ponto:
                    erros.append("Perfil de frequência não preenchido")
                elif "NÃO IDENTIFICADO" in padrao:
                    erros.append("Cargo não homologado na validação")
                elif isento_padrao and not isento_situacao:
                    erros.append("Deveria ser ISENTO mas está no Sindicato")
                elif not isento_padrao and isento_situacao:
                    erros.append("Deveria REGISTRAR PONTO mas está como ISENTO")

                if escala_func in ['', '0']:
                    erros.append("Código da escala em branco na folha")
                elif pd.isna(carga_gab):
                    erros.append("Colaborador ausente no gabarito do RH")
                elif carga_func != carga_gab:
                    erros.append(f"Carga Mensal diverge ({carga_func}h folha vs {carga_gab}h gabarito)")

                return "DIVERGÊNCIA: " + " | ".join(erros) if erros else "OK"

            df_consolidado['VALIDACAO PONTO'] = df_consolidado.apply(auditar_ponto_mestre, axis=1)

            def checar_malha_fina_escala(linha):
                cat_str = str(linha.get('CATEGORIA_SEGURO', '')).strip().replace('.0', '').zfill(3)
                match_cat = re.search(r'(\d+)', cat_str)
                cat = int(match_cat.group(1)) if match_cat else 0

                carga = extrair_hora_pura(linha.get('CARGA_HORÁRIA_MENSAL_SEGURO', 0))
                escala_str = str(linha.get('CÓDIGO_ESCALA_LIMPO', '')).strip().replace('.0', '')

                if cat == 101:
                    return "OK" if carga == 220 else f"AJUSTAR (Cat. 101 exige 220h, atual {carga}h)"
                elif cat == 103:
                    return "OK" if carga in [100, 110, 120,
                                             125] else f"AJUSTAR (Cat. 103 exige 100/110/120/125h, atual {carga}h)"
                elif cat == 111:
                    if escala_str == '5004': return "AJUSTAR (Escala 5004 descontinuada p/ Intermitente)"
                    if escala_str not in ['5005', '5006', '5007', '5008',
                                          '5009']: return f"AJUSTAR (Escala {escala_str} inválida p/ Cat. 111)"
                    if carga not in [180, 220]: return f"AJUSTAR (Cat. 111 exige 180h ou 220h, atual {carga}h)"
                    return "OK"
                elif cat == 901:
                    return "OK" if escala_str == '869' else f"AJUSTAR (Cat. 901 exige escala 869, atual {escala_str})"

                return f"AJUSTAR (Categoria {cat_str} sem regra de escala definida)"

            df_consolidado['VALIDACAO ESCALA'] = df_consolidado.apply(checar_malha_fina_escala, axis=1)

            def checar_cargo_vs_categoria(linha):
                cat_str = str(linha.get('CATEGORIA_SEGURO', '')).strip().replace('.0', '').zfill(3)
                match_cat = re.search(r'(\d+)', cat_str)
                cat = int(match_cat.group(1)) if match_cat else 0
                cargo = str(linha.get('NOME_CARGO_SEGURO', '')).upper().strip()

                if cat == 103 and "APRENDIZ" not in cargo:
                    return "AJUSTAR CARGO (Cat. 103 exige cargo de Aprendiz)"
                elif cat == 111 and "INTERMITENTE" not in cargo:
                    return "AJUSTAR CARGO (Cat. 111 exige cargo de Intermitente)"
                elif "APRENDIZ" in cargo and cat != 103:
                    return "AJUSTAR CATEGORIA (Cargo de Aprendiz exige Cat. 103)"
                elif "INTERMITENTE" in cargo and cat != 111:
                    return "AJUSTAR CATEGORIA (Cargo de Intermitente exige Cat. 111)"
                return "OK"

            df_consolidado['VALIDACAO CARGOXCATEGORIA'] = df_consolidado.apply(checar_cargo_vs_categoria, axis=1)

            vinculo = df_consolidado['VINCULO_SEGURO'].fillna('').astype(str).str.strip()
            regime = df_consolidado['REGIME_SEGURO'].fillna('').astype(str).str.strip()
            df_consolidado['VALIDACAO VINCULO CLT'] = np.where(
                (vinculo.isin(['', 'nan', '0'])) & (regime.isin(['', 'nan', '0'])),
                'REVISAR CADASTRO (Vínculo/Regime ausente)',
                np.where(vinculo.isin(['', 'nan', '0']), 'REVISAR CADASTRO (Vínculo ausente)',
                         np.where(regime.isin(['', 'nan', '0']), 'REVISAR CADASTRO (Regime ausente)', 'OK')))

            esta_no_banco = df_consolidado['ESTA_NA_PLANILHA_BANCO'] == True
            situacao_ativo = df_consolidado['SITUACAO_SEGURO'].fillna('').astype(
                str).str.upper().str.strip() == 'ATIVIDADE NORMAL'
            conta_banco_txt = df_consolidado['CONTA_BANCO_RAW'].fillna('').astype(str).str.strip()
            natureza_esmagada = df_consolidado['NATUREZA_SEGURO'].fillna('').astype(str).str.upper().str.replace(
                r'[^A-Z0-9]', '', regex=True)
            cod_banco_limpo = df_consolidado['COD_BANCO'].fillna('').astype(str).str.extract(r'(\d+)',
                                                                                             expand=False).str.lstrip(
                '0').fillna('')

            is_bradesco = cod_banco_limpo == '237'
            is_santander = cod_banco_limpo == '33'
            natureza_bradesco = natureza_esmagada.str.contains('BRETAS|GBABORSA|GBARBOSA|MERCANTIL|GIGA', na=False)
            natureza_santander = natureza_esmagada.str.contains('PERINI|PREZUNIC|SPID|ADM400', na=False) | (
                        cod_estab == '400')
            cond_banco_ok = (is_bradesco & natureza_bradesco) | (is_santander & natureza_santander)

            detalhe_erro_banco = np.where(is_bradesco & ~natureza_bradesco, 'REVISAR DADOS (237 incompatível)',
                                          np.where(is_santander & ~natureza_santander,
                                                   'REVISAR DADOS (033 incompatível)',
                                                   'REVISAR DADOS (Banco não homologado)'))
            df_consolidado['VALIDACAO BANCO'] = np.where(~esta_no_banco, 'ACUSAR COMO ERRO AJUSTAR (Banco ausente)',
                                                         np.where(situacao_ativo & (conta_banco_txt == ''),
                                                                  'ACUSAR COMO ERRO AJUSTAR (Conta em branco)',
                                                                  np.where(
                                                                      df_consolidado['ERRO_BANCO_DUPLICADO'].fillna(
                                                                          False),
                                                                      'VALIDAR ERRO BANCO AJUSTAR (Conta repetida p/ pessoas dif.)',
                                                                      np.where(cond_banco_ok, 'OK',
                                                                               detalhe_erro_banco))))

            # NOVIDADE: A CHAVE 5_BCR FOI INCLUÍDA NA LISTA DE CENTROS DE CUSTO VÁLIDOS!
            empresa_limpa_cc = df_consolidado['COD_EMPRESA_SEGURO'].fillna('').astype(str).str.replace(r'\.0$', '',
                                                                                                       regex=True).str.strip()
            prefixo_cc_limpo = df_consolidado['PREFIXO_CC'].fillna('').astype(str).str.strip().str.upper()
            df_consolidado['CHAVE_EMPRESA_CC'] = empresa_limpa_cc + "_" + prefixo_cc_limpo
            chaves_validas_cc = [
                '2_BCB', '2_BCG', '2_BCO', '2_BCS', '2_BFD', '2_BGO', '2_BHG',
                '2_GHL', '2_BHO', '2_BRG', '2_BSB', '2_BSG', '2_BSL', '2_BSZ',
                '2_BVO', '2_BWG', '2_BYG', '2_BZB', '2_BHL', '2_BCR', '5_BGO',
                '5_BHA', '5_BHR', '5_BSG', '5_BSL', '7_BGO', '7_BSF', '7_BSL',
                '7_BSP', '7_BYP', '5_BCR'  # <--- ADICIONADO BCR PARA EMPRESA 5 AQUI
            ]
            df_consolidado['VALIDACAO CENTRO DE CUSTO'] = np.where(
                (prefixo_cc_limpo == '') | (prefixo_cc_limpo == 'NAN'), 'REVISAR ALOCAÇÃO (Centro de custo em branco)',
                np.where(df_consolidado['CHAVE_EMPRESA_CC'].isin(chaves_validas_cc), 'CENTRO DE CUSTO OK',
                         'REVISAR ALOCAÇÃO (Prefixo incompatível)'))
            self.concluir_passo()

            # =========================================================================
            # 10. MAPEAMENTO (DE-PARA) - ORDEM FINAL
            # =========================================================================
            self.iniciar_passo("Estruturando as colunas (47 originais homologadas)...")
            df_consolidado = df_consolidado.loc[:, ~df_consolidado.columns.duplicated()]

            # DICIONÁRIO EXATO DE ACORDO COM SEU CÓDIGO FONTE HOMOLOGADO
            mapeamento_colunas = {
                'Nº COD EMPRESA': 'COD_EMPRESA_SEGURO',
                'BANDEIRA': 'BANDEIRA_SEGURO',
                'Nº DA LOJA': 'NUMERO_LOJA_SEGURO',
                'ESTADO (UF)': 'ESTADO_CORRETO',  # <--- MUDOU PARA A COLUNA COM O ESTADO DA NOVA PLANILHA
                'MATRÍCULA COLABORADOR': 'MATRICULA_MASTER',
                'NOME COLABORADOR': 'NOME_PESSOA_SEGURO',
                'CPF': 'CPF_SEGURO',
                'DATA DE ADMISSÃO': 'DATA_ADMISSAO_FORMATADA',
                'NOME_CARGO': 'NOME_CARGO_SEGURO',
                'SITUAÇÃO': 'SITUACAO_SEGURO',
                'CÓDIGO CENTRO DE CUSTO': 'COD_CC_SEGURO',
                'PREFIXO CENTRO DE CUSTO': 'PREFIXO_CC',
                'CÓDIGO UNIDADE ORGANIZACIONAL': 'COD_UO_SEGURO',
                'CÓDIGO ESTABELECIMENTO': 'COD_ESTAB_SEGURO',
                'CÓDIGO_ESCALA': 'CÓDIGO_ESCALA_SEGURO',
                'DESCRIÇÃO DA ESCALA': 'DESCRIÇÃO_ESCALA',
                'CARGA_HORÁRIA_SEMANAL': 'CARGA_HORÁRIA_SEMANAL_SEGURO',
                'CARGA_HORÁRIA_MENSAL': 'CARGA_HORÁRIA_MENSAL_SEGURO',
                'CATEGORIA': 'CATEGORIA_SEGURO',
                'VÍNCULO EMPREGATÍCIO': 'VINCULO_SEGURO',
                'REGIME TRABALHISTA': 'REGIME_SEGURO',
                'CÓDIGO BANCO': 'COD_BANCO',
                'AGÊNCIA BANCO': 'AGENCIA_BANCO_RAW',
                'CONTA BANCO': 'CONTA_BANCO_RAW',
                'CÓDIGO DEPENDENTE': 'CÓDIGO DEPENDENTE',
                'NOME DEPENDENTE': 'NOME DEPENDENTE',
                'IDADE DEPENDENTE': 'IDADE DEPENDENTE',
                'PARENTESCO': 'PARENTESCO_CLEAN',
                'EVENTOS DE SALÁRIO FAMÍLIA': '126 - SALARIO FAMILIA',
                'BASE SAL FAMILIA': 'BASE SAL FAMILIA',
                'Salário + Eventos Incide INSS': 'BASE_INSS_CALCULADA',
                'VALIDACAO SAL FAMILIA': 'VALIDACAO SAL FAMILIA',
                '110 - AJUDA CUSTO TELETRABALHO': 'VALOR_110_TELE',
                '190 - AJUDA CUSTO TELETRAB MOBILIA': 'VALOR_190_TELE',
                'VALIDACAO TELETRABALHO': 'VALIDACAO TELETRABALHO',
                'VALIDACAO INTERMITENTE': 'VALIDACAO INTERMITENTE',
                'OPÇÃO REFEITÓRIO': 'OPÇÃO REFEITÓRIO',
                'VALIDACAO REFEITORIO': 'VALIDACAO REFEITORIO',
                'VALIDACAO VINCULO CLT': 'VALIDACAO VINCULO CLT',
                'VALIDACAO BANCO': 'VALIDACAO BANCO',
                'VALIDACAO CENTRO DE CUSTO': 'VALIDACAO CENTRO DE CUSTO',
                'PADRAO INTERNO-REGISTROXCARGO': 'PADRAO INTERNO-REGISTROXCARGO',
                'SITUACAO DE PONTO SISTEMA': 'SITUACAO DE PONTO SISTEMA',
                'VALIDACAO PONTO': 'VALIDACAO PONTO',
                'VALIDACAO ESCALA': 'VALIDACAO ESCALA',
                'VALIDACAO CARGOXCATEGORIA': 'VALIDACAO CARGOXCATEGORIA',
                'NOME NO CRACHÁ': 'NOME_CRACHA_RAW',
                'VALIDACAO CRACHÁ': 'VALIDACAO CRACHÁ'
            }

            df_relatorio_final = pd.DataFrame()
            for nome_final, nome_original in mapeamento_colunas.items():
                if nome_original and nome_original in df_consolidado.columns:
                    df_relatorio_final[nome_final] = df_consolidado[nome_original]
                else:
                    df_relatorio_final[nome_final] = ""

            # --- 11. TRATAMENTO FINANCEIRO E AUDITORIA COMPLETA ---
            df_relatorio_final['FLAG_FILHO_MENOR_14'] = df_consolidado['FLAG_FILHO_MENOR_14']

            def auditar_salario_familia_completo(row):
                salario = pd.to_numeric(row["Salário + Eventos Incide INSS"], errors="coerce")
                salario_val = 0.0 if pd.isna(salario) else salario
                teto = pd.to_numeric(row["BASE SAL FAMILIA"], errors="coerce")
                teto_val = 1980.38 if pd.isna(teto) else teto
                idade = pd.to_numeric(row["IDADE DEPENDENTE"], errors="coerce")
                recebido = pd.to_numeric(row["EVENTOS DE SALÁRIO FAMÍLIA"], errors="coerce")
                recebido_val = 0.0 if pd.isna(recebido) else recebido
                parentesco = str(row["PARENTESCO"]).strip().upper()
                situacao = str(row["SITUAÇÃO"]).strip().upper()
                tem_filho_menor = row.get("FLAG_FILHO_MENOR_14", False)

                is_parentesco_elegivel = any(p in parentesco for p in ["FILHO", "FILHA", "GUARDA", "ENTEADO"])

                if parentesco != "" and not is_parentesco_elegivel: return "[4] NÃO ELEGÍVEL — Parentesco não elegível"

                if recebido_val > 0:
                    if salario_val > teto_val: return "[3] REVISAR — Pago indevidamente (Salário acima do teto)"
                    if situacao not in ["ATIVIDADE NORMAL", "FÉRIAS",
                                        "FERIAS"]: return f"[3] REVISAR — Pago indevidamente (Contrato em {situacao})"
                    if parentesco == "" and not tem_filho_menor: return "[3] REVISAR — Pago indevidamente (Colaborador sem dependentes elegíveis)"
                    if parentesco != "" and pd.notna(idade) and idade > 13:
                        if tem_filho_menor: return "[1] OK — Regular (Recebimento coberto por irmão caçula)"
                        return "[3] REVISAR — Pago indevidamente (Dependente com idade superior a 13 anos e 11 meses, sem irmão elegível)"
                    return "[1] OK — Regular"
                else:
                    if situacao not in ["ATIVIDADE NORMAL", "FÉRIAS",
                                        "FERIAS"]: return "[4] NÃO ELEGÍVEL — Contrato Suspenso"
                    if salario_val > teto_val or salario_val == 0: return "[4] NÃO ELEGÍVEL — Salário acima do teto"
                    if parentesco != "" and pd.notna(
                        idade) and idade > 13: return "[4] NÃO ELEGÍVEL — Idade superior a 13 anos e 11 meses"
                    if parentesco == "" and not tem_filho_menor: return "[4] NÃO ELEGÍVEL — Sem Requisitos Legais"
                    return "[2] AJUSTAR — Elegível, mas não pago"

            df_relatorio_final["VALIDACAO SAL FAMILIA"] = df_relatorio_final.apply(auditar_salario_familia_completo,
                                                                                   axis=1)
            df_relatorio_final.drop(columns=['FLAG_FILHO_MENOR_14'], inplace=True)

            df_relatorio_final['FLAG_TEM_REGISTRO_CRACHA'] = df_consolidado['FLAG_TEM_REGISTRO_CRACHA']

            def auditar_cracha_linha(row):
                nome_completo = str(row.get('NOME COLABORADOR', '')).strip().upper()
                nome_cracha = str(row.get('NOME NO CRACHÁ', '')).strip().upper()
                tem_registro_cracha = row.get('FLAG_TEM_REGISTRO_CRACHA', False)
                if nome_cracha == "" or nome_cracha == "NAN":
                    if tem_registro_cracha: return "[3] REVISAR — Crachá emitido/registrado sem nome definido"
                    return "[4] NÃO SOLICITADO — Sem registro de crachá"
                if re.search(r'[ÁÉÍÓÚÂÊÔÃÕÀÇÜ]', nome_cracha): return "[2] AJUSTAR — Contém Acentuação / Ç"

                palavras_completo = nome_completo.split()
                palavras_cracha = nome_cracha.split()
                if len(palavras_cracha) == 1:
                    return "[1] OK — Primeiro Nome"
                elif len(palavras_cracha) == 2:
                    def remover_acentos(txt):
                        return ''.join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')

                    comp_clean = [remover_acentos(w) for w in palavras_completo]
                    cracha_clean = [remover_acentos(w) for w in palavras_cracha]
                    if len(comp_clean) >= 2 and comp_clean[:2] == cracha_clean[:2]:
                        return "[1] OK — Nome Composto"
                    else:
                        return "[2] AJUSTAR — Contém Sobrenome"
                else:
                    return "[2] AJUSTAR — Contém Sobrenome"

            df_relatorio_final["VALIDACAO CRACHÁ"] = df_relatorio_final.apply(auditar_cracha_linha, axis=1)
            df_relatorio_final.drop(columns=['FLAG_TEM_REGISTRO_CRACHA'], inplace=True)

            def converter_para_numero_nativo(val):
                if pd.isna(val) or val == "": return ""
                v_str = str(val).split('.')[0].strip()
                return int(v_str) if v_str.isdigit() else v_str

            colunas_chaves = ['MATRÍCULA COLABORADOR', 'CONTA BANCO', 'AGÊNCIA BANCO', 'CÓDIGO_ESCALA', 'CATEGORIA',
                              'CÓDIGO BANCO', 'CÓDIGO ESTABELECIMENTO']
            for c in colunas_chaves:
                if c in df_relatorio_final.columns:
                    df_relatorio_final[c] = df_relatorio_final[c].apply(converter_para_numero_nativo)

            correcoes_texto = {'CÃ´njuge': 'Cônjuge', 'NÃ£o': 'Não', 'SEM INFORMAÃ‡ÃƒO': 'SEM INFORMAÇÃO',
                               'atÃ©': 'até', 'Ã§': 'ç', 'Ã£': 'ã', 'Ãµ': 'õ', 'Ãá': 'á', 'Ã©': 'é', 'Ã­': 'í',
                               'Ã³': 'ó', 'Ãº': 'ú', 'Ã¢': 'â', 'Ãª': 'ê', 'Ã´': 'ô', 'Ã‡': 'Ç', 'Ãƒ': 'Ã', 'Ã‰': 'É',
                               'Âº': 'º', 'â€“': '-', 'Ã\x8d': 'Í', 'Ã': 'Í', 'Ã¡': 'à'}
            for col in df_relatorio_final.select_dtypes(include=['object', 'string']).columns:
                df_relatorio_final[col] = df_relatorio_final[col].replace(correcoes_texto, regex=True)

            def colorir_celula(val):
                if not isinstance(val, str): return ''
                v_upper = val.upper()
                if any(x in v_upper for x in
                       ['REVISAR', 'ERRO', 'AJUSTAR', 'DIVERGência', 'NÃO ENCONTRADO', 'NÃO IDENTIFICADO']):
                    return 'background-color: #ff4d4d'
                return ''

            df_estilizado = df_relatorio_final.style.map(colorir_celula)
            self.concluir_passo()

            # =========================================================================
            # 12. EXPORTAÇÃO
            # =========================================================================
            self.iniciar_passo("Salvando a super-planilha estilizada no disco (Pode levar de 5 a 15 segundos)...")
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            caminho_saida = os.path.join(self.caminho_pasta, f"RELATORIO_ANALITICO_MASTER_{timestamp}.xlsx")
            df_estilizado.to_excel(caminho_saida, index=False, engine='openpyxl')
            self.concluir_passo()

            self.iniciar_passo("Aplicando formato contábil nativo do Excel...")
            wb = openpyxl.load_workbook(caminho_saida)
            ws = wb.active
            formato_contabil = '_-"R$"* #,##0.00_-;-"R$"* #,##0.00_-;_-"R$"* "-"??_-;_-@_-'

            for col in ws.iter_cols(min_row=1, max_row=ws.max_row):
                if col[0].value in ['Salário + Eventos Incide INSS', 'BASE SAL FAMILIA', 'EVENTOS DE SALÁRIO FAMÍLIA',
                                    '110 - AJUDA CUSTO TELETRABALHO', '190 - AJUDA CUSTO TELETRAB MOBILIA']:
                    for celula in col[1:]: celula.number_format = formato_contabil

            wb.save(caminho_saida)
            self.concluir_passo()

            self.print_log("-" * 50)
            self.print_log(f"✨ SUCESSO ABSOLUTO! O Censo Master gerou o relatório limpo em:\n👉 {caminho_saida}")

            self.root.after(0, lambda: self.btn_pasta.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.btn_gerar.config(state=tk.NORMAL))
            messagebox.showinfo("Sucesso", "O Censo Master foi gerado com sucesso!")

        except Exception as e:
            self.concluir_passo(alerta=True)
            erro = traceback.format_exc()
            self.print_log(f"\n❌ ERRO CRÍTICO NA EXECUÇÃO:\n{str(e)}\n\nDetalhes:\n{erro}")
            self.root.after(0, lambda: self.btn_pasta.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.btn_gerar.config(state=tk.NORMAL))


if __name__ == "__main__":
    root = tk.Tk()
    app = MotorRH_App(root)
    root.mainloop()