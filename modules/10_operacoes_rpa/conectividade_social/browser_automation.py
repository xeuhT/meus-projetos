"""
Módulo de Automação de Navegador (Chrome com Suporte a Certificado Digital mTLS)
Conectividade Social V2 - Cencosud Brasil
Suporta consulta contínua por PIS com fallback automático por Nome do Trabalhador.
"""
import json
import time
import os
import sys
import re
import base64
from typing import Dict, Any, Optional
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

sys.path.insert(0, str(Path(__file__).resolve().parent))

from config import (
    URL_CONECTIVIDADE_V2,
    CERTIFICADOS_CENCOSUD,
    CERTIFICADO_PADRAO,
    TIMEOUT_PAGINA,
    TIMEOUT_ELEMENTO,
    BASE_CONTA_PADRAO,
    DIR_EXTRATOS
)
from extractor import extrair_dados_extrato_html, detectar_mensagem_erro


class ConectividadeSocialBot:
    """Controlador Selenium para automação no portal Conectividade Social V2 da Caixa."""

    def __init__(self, tipo_certificado: str = CERTIFICADO_PADRAO, porta_depuracao: Optional[int] = None, diretorio_saida: Optional[Path] = None):
        self.tipo_certificado = tipo_certificado
        self.porta_depuracao = porta_depuracao
        self.diretorio_saida = Path(diretorio_saida) if diretorio_saida else DIR_EXTRATOS
        self.diretorio_saida.mkdir(parents=True, exist_ok=True)
        self.driver: Optional[webdriver.Chrome] = None
        self.cert_info = CERTIFICADOS_CENCOSUD.get(tipo_certificado, CERTIFICADOS_CENCOSUD["comercial"])

    def iniciar_navegador(self) -> webdriver.Chrome:
        """
        Inicializa o Google Chrome nativo configurado para mTLS com seleção automática do certificado digital.
        """
        options = Options()
        options.page_load_strategy = 'eager'

        filtro_cn = self.cert_info["filtro_cn"]
        cert_rule = {
            "pattern": "https://[*.]caixa.gov.br",
            "filter": {
                "SUBJECT": {"CN": filtro_cn}
            }
        }
        cert_rule_json = json.dumps(cert_rule)

        options.add_argument(f"--auto-select-certificate-for-urls={cert_rule_json}")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-notifications")
        options.add_argument("--ignore-certificate-errors")

        prefs = {
            "security.auto_select_certificate_for_urls": [cert_rule_json],
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False
        }
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)

        caminhos_chrome = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe")
        ]
        for caminho in caminhos_chrome:
            if Path(caminho).exists():
                options.binary_location = caminho
                break

        print(f"[*] Iniciando driver do Chrome para certificado: {self.cert_info['nome']}...")
        self.driver = webdriver.Chrome(options=options)
        self.driver.set_page_load_timeout(60)
        return self.driver

    def fechar_navegador(self):
        """Finaliza a sessão do navegador liberando recursos."""
        if self.driver:
            try:
                self.driver.quit()
                print("[*] Navegador finalizado com sucesso.")
            except Exception as e:
                print(f"[!] Erro ao fechar navegador: {e}")
            finally:
                self.driver = None

    def aguardar_elemento(self, localizadores: list, timeout: int = TIMEOUT_ELEMENTO):
        """Tenta encontrar um elemento utilizando uma lista de possíveis localizadores."""
        wait = WebDriverWait(self.driver, timeout)
        ultimo_erro = None
        for by, sel in localizadores:
            try:
                return wait.until(EC.presence_of_element_located((by, sel)))
            except TimeoutException as e:
                ultimo_erro = e
                continue
        raise ultimo_erro or TimeoutException("Elemento não encontrado na lista de seletores.")

    def clicar_elemento_seguro(self, localizadores: list, timeout: int = TIMEOUT_ELEMENTO):
        """Localiza e clica em um elemento com scroll e tratamento de interceptação."""
        wait = WebDriverWait(self.driver, timeout)
        ultimo_erro = None
        for by, sel in localizadores:
            try:
                elem = wait.until(EC.element_to_be_clickable((by, sel)))
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                time.sleep(0.3)
                elem.click()
                return True
            except Exception as e:
                ultimo_erro = e
                continue
        raise ultimo_erro or NoSuchElementException("Não foi possível clicar em nenhum dos seletores fornecidos.")

    def _aguardar_carregamento(self, timeout: int = 10):
        """Aguarda spinners de carregamento do Angular da Caixa desaparecerem."""
        try:
            wait = WebDriverWait(self.driver, timeout)
            wait.until(EC.invisibility_of_element_located((
                By.XPATH, "//*[contains(@class, 'loading') or contains(@class, 'spinner') or contains(@class, 'overlay') or contains(@class, 'backdrop')]"
            )))
        except Exception:
            pass

    def _esta_autenticado_ou_no_portal(self) -> bool:
        """Verifica de forma ampla e resiliente se a sessão já está logada no portal do empregador."""
        try:
            src = (self.driver.page_source or "").lower()
            indicadores = [
                "serviço desejado",
                "servico desejado",
                "solicitar extrato",
                "localização do trabalhador",
                "localizacao do trabalhador",
                "base da conta",
                "dados do trabalhador",
                "relação de trabalhadores",
                "relacao de trabalhadores",
                "extrato fgts",
                "portal do empregador",
                "serviços ao empregador",
                "servicos ao empregador",
                "conectividade social empregador"
            ]
            if any(ind in src for ind in indicadores):
                return True

            # Verificar se já existe um dropdown select na tela
            selects = self.driver.find_elements(By.TAG_NAME, "select")
            if len(selects) > 0:
                for sel in selects:
                    if sel.is_displayed():
                        return True
        except Exception:
            pass
        return False

    def login_conectividade(self) -> bool:
        """
        Acessa a Home do Conectividade Social V2 e autentica clicando em 'Empregador'.
        Permite janela de tempo para o usuário selecionar o certificado caso a Caixa exija confirmação.
        """
        if not self.driver:
            self.iniciar_navegador()

        print(f"[*] Navegando para {URL_CONECTIVIDADE_V2}...")
        try:
            self.driver.get(URL_CONECTIVIDADE_V2)
        except TimeoutException:
            print("[*] Carregamento inicial continuou com sucesso.")
        time.sleep(2)
        self._aguardar_carregamento(timeout=8)

        if self._esta_autenticado_ou_no_portal():
            print("[+] Já autenticado no Portal do Empregador!")
            return True

        print("[*] Localizando botão 'Empregador' no portal da Caixa...")
        seletores_empregador = [
            (By.XPATH, "//button[contains(translate(., 'EMPREGADOR', 'empregador'), 'empregador')]"),
            (By.XPATH, "//a[contains(translate(., 'EMPREGADOR', 'empregador'), 'empregador')]"),
            (By.ID, "btnEmpregador"),
            (By.XPATH, "//*[@id='btnEmpregador']")
        ]

        # Tentar localizar o botão de Empregador
        btn_clicado = False
        for by, sel in seletores_empregador:
            try:
                elem = self.driver.find_element(by, sel)
                if elem:
                    # Rolar até o botão e clicar via JavaScript para evitar interceptação pelo spinner
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                    time.sleep(0.3)
                    self.driver.execute_script("arguments[0].click();", elem)
                    print("[*] Botão 'Empregador' acionado!")
                    btn_clicado = True
                    break
            except Exception:
                continue

        if not btn_clicado:
            print("[*] Botão 'Empregador' não precisou ser clicado ou já estava acionado.")

        print("\n" + "=" * 70)
        print(">>> AGUARDANDO AUTENTICAÇÃO COM CERTIFICADO DIGITAL <<<")
        print("Selecione o certificado CENCOSUD BRASIL COMERCIAL LTDA e clique em OK.")
        print("(O robô detectará a tela automaticamente assim que você autenticar)")
        print("=" * 70 + "\n")

        inicio_espera = time.time()
        while time.time() - inicio_espera < 120:
            if self._esta_autenticado_ou_no_portal():
                print("[+] Autenticação no Portal do Empregador confirmada com sucesso!")
                time.sleep(2)
                return True
            time.sleep(1)

        print("[!] Tempo limite de espera para autenticação excedido.")
        return False

    def navegar_para_solicitar_extrato(self) -> bool:
        """
        Garante que estamos na tela do formulário de pesquisa ('Localização do Trabalhador').
        Se já estiver na tela do formulário ou com a opção selecionada, não falha.
        """
        self._aguardar_carregamento(timeout=5)

        # Se já estiver na tela de busca com os campos visíveis, não precisa fazer nada!
        if self._esta_no_formulario_busca():
            print("[+] Já posicionado no formulário de pesquisa!")
            return True

        print("[*] Selecionando serviço 'Solicitar Extrato do Trabalhador'...")
        try:
            seletores_select = [
                (By.XPATH, "//select[contains(@name, 'servico') or contains(@id, 'servico')]"),
                (By.XPATH, "//select[preceding::*[contains(text(), 'serviço desejado') or contains(text(), 'servico desejado')]]"),
                (By.XPATH, "//select")
            ]
            elem_select = self.aguardar_elemento(seletores_select, timeout=TIMEOUT_ELEMENTO)
            select_obj = Select(elem_select)

            # Verificar se a opção já está selecionada atualmente
            opcao_ja_selecionada = False
            try:
                if "Extrato do Trabalhador" in select_obj.first_selected_option.text:
                    opcao_ja_selecionada = True
            except Exception:
                pass

            if not opcao_ja_selecionada:
                for opt in select_obj.options:
                    texto_opt = opt.text.strip()
                    if "Solicitar Extrato do Trabalhador" in texto_opt or "Extrato do Trabalhador" in texto_opt:
                        select_obj.select_by_visible_text(texto_opt)
                        break

            time.sleep(2)
            self._aguardar_carregamento(timeout=8)
            WebDriverWait(self.driver, TIMEOUT_ELEMENTO).until(
                lambda d: self._esta_no_formulario_busca()
            )
            print("[+] Formulário de solicitação de extrato carregado com sucesso!")
            return True
        except Exception as e:
            if self._esta_no_formulario_busca():
                print("[+] Formulário de pesquisa identificado com sucesso.")
                return True
            print(f"[!] Aviso na navegação: {e}")
            return False

    def _esta_no_formulario_busca(self) -> bool:
        """
        Verifica se a tela atual é o formulário de localização do trabalhador.
        Checa estritamente a presença do campo de entrada txtNumPisPasep e ausência do extrato.
        """
        try:
            # Se a tela atual contiver o título de extrato, NÃO é o formulário de busca
            src = (self.driver.page_source or "").lower()
            if "extrato fgts do trabalhador" in src or "fins rescisórios" in src or "fins rescisorios" in src:
                return False

            # Verifica pelo campo oficial de input de PIS da Caixa (único do formulário)
            inputs_pis = self.driver.find_elements(By.NAME, "txtNumPisPasep")
            if len(inputs_pis) > 0 and inputs_pis[0].is_displayed():
                return True

            return "txtnumpispasep" in src and "localiza" in src
        except Exception:
            return False

    def _selecionar_base_conta(self, base_uf: str = BASE_CONTA_PADRAO):
        """Seleciona a Base da Conta no dropdown."""
        try:
            seletores_base = [
                (By.XPATH, "//select[contains(@name, 'base') or contains(@id, 'base')]"),
                (By.XPATH, "//select[preceding::*[contains(text(), 'Base da Conta')]]"),
                (By.XPATH, "//select[1]")
            ]
            elem_base = self.aguardar_elemento(seletores_base, timeout=TIMEOUT_ELEMENTO)
            select_base = Select(elem_base)

            uf_prefix = base_uf.split("-")[0].strip().upper()
            base_selecionada = False
            for opt in select_base.options:
                if uf_prefix in opt.text.upper():
                    select_base.select_by_visible_text(opt.text)
                    base_selecionada = True
                    break

            if not base_selecionada:
                select_base.select_by_index(1)
        except Exception as e:
            print(f"[!] Aviso ao selecionar Base da Conta: {e}")

    def _clicar_continuar(self):
        """Clica no botão CONTINUAR da tela de pesquisa acionando a função nativa ou a imagem."""
        self._aguardar_carregamento(timeout=3)

        # 1. Tentar acionar diretamente a função JavaScript nativa da página da Caixa
        try:
            self.driver.execute_script("if (typeof subm_localizar_trabalhador === 'function') { subm_localizar_trabalhador(); }")
            time.sleep(1.5)
            return
        except Exception:
            pass

        # 2. Localizar por seletores específicos da imagem e link da Caixa
        seletores_continuar = [
            (By.XPATH, "//a[contains(@href, 'subm_localizar_trabalhador')]"),
            (By.XPATH, "//img[contains(@src, 'botao_continuar')]/parent::a"),
            (By.XPATH, "//img[contains(@src, 'botao_continuar')]"),
            (By.XPATH, "//a[contains(@href, 'subm_localizar')]"),
            (By.XPATH, "//*[contains(@src, 'continuar')]"),
            (By.XPATH, "//button[contains(translate(text(), 'CONTINUAR', 'continuar'), 'continuar')]"),
            (By.XPATH, "//input[@type='button' or @type='submit'][contains(translate(@value, 'CONTINUAR', 'continuar'), 'continuar')]")
        ]
        try:
            self.clicar_elemento_seguro(seletores_continuar, timeout=TIMEOUT_ELEMENTO)
        except Exception:
            for by, sel in seletores_continuar:
                try:
                    elem = self.driver.find_element(by, sel)
                    if elem:
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                        time.sleep(0.3)
                        self.driver.execute_script("arguments[0].click();", elem)
                        return
                except Exception:
                    continue

    def _executar_pesquisa_formulario(self, pis: str = "", nome: str = "", base_uf: str = BASE_CONTA_PADRAO) -> Dict[str, Any]:
        """Preenche o formulário por PIS ou por Nome e clica em continuar."""
        # 1. Base da Conta (sltRegiao)
        self._selecionar_base_conta(base_uf)

        # 2. Preencher Nome (se fornecido, campo txtNomeTrab)
        try:
            seletores_nome = [
                (By.NAME, "txtNomeTrab"),
                (By.XPATH, "//input[@name='txtNomeTrab']"),
                (By.XPATH, "//input[contains(@name, 'nome') or contains(@id, 'nome')]"),
                (By.XPATH, "//input[preceding::*[contains(text(), 'Nome do Trabalhador')]]"),
                (By.XPATH, "//input[@type='text'][1]")
            ]
            elem_nome = self.aguardar_elemento(seletores_nome, timeout=5)
            elem_nome.clear()
            if nome:
                elem_nome.send_keys(nome.strip().upper())
        except Exception:
            pass

        # 3. Preencher PIS (se fornecido, campo txtNumPisPasep)
        try:
            pis_limpo = re.sub(r"\D", "", str(pis or ""))
            if pis_limpo:
                pis_limpo = pis_limpo.zfill(11)

            seletores_pis = [
                (By.NAME, "txtNumPisPasep"),
                (By.XPATH, "//input[@name='txtNumPisPasep']"),
                (By.XPATH, "//input[contains(@name, 'pis') or contains(@id, 'pis')]"),
                (By.XPATH, "//input[preceding::*[contains(text(), 'PIS/PASEP/NIT')]]"),
                (By.XPATH, "//input[@type='text'][2]")
            ]
            elem_pis = self.aguardar_elemento(seletores_pis, timeout=5)
            elem_pis.clear()
            if pis_limpo:
                elem_pis.send_keys(pis_limpo)
        except Exception as e:
            return {
                "pis_pesquisado": re.sub(r"\D", "", str(pis or "")),
                "base_uf": base_uf,
                "status": "ERRO",
                "mensagem": f"Falha ao localizar campo de formulário: {e}"
            }

        # 4. Clicar em CONTINUAR
        try:
            self._clicar_continuar()
        except Exception as e:
            return {
                "pis_pesquisado": pis,
                "base_uf": base_uf,
                "status": "ERRO",
                "mensagem": f"Falha ao clicar em CONTINUAR: {e}"
            }

        # 5. Aguardar resposta da Caixa
        time.sleep(2)
        try:
            WebDriverWait(self.driver, TIMEOUT_ELEMENTO).until(
                lambda d: "Extrato FGTS do Trabalhador" in d.page_source or "Fins Rescisórios" in d.page_source or "Relação de Trabalhadores" in d.page_source or detectar_mensagem_erro(d.page_source) is not None
            )
        except TimeoutException:
            pass

        # 6. Se caiu na lista de múltiplos trabalhadores ou conta única
        html_atual = self.driver.page_source
        
        if "Relação de Trabalhadores" in html_atual and "Extrato FGTS do Trabalhador" not in html_atual:
            try:
                link_conta = self.driver.find_element(By.XPATH, "//a[contains(@href, 'extrato') or contains(@href, 'conta') or contains(@href, 'trabalhador')]")
                link_conta.click()
                time.sleep(2)
                html_atual = self.driver.page_source
            except Exception:
                pass

        dados = extrair_dados_extrato_html(html_atual)
        dados["pis_pesquisado"] = pis
        dados["base_uf"] = base_uf
        return dados

    def preencher_formulario_e_consultar(self, pis: str, nome: str = "", base_uf: str = BASE_CONTA_PADRAO) -> Dict[str, Any]:
        """
        Consulta pelo PIS. Caso o trabalhador não seja localizado e o Nome esteja presente,
        aciona automaticamente o fallback de pesquisa pelo Nome do Trabalhador.
        Remove pontos, traços e espaços do PIS antes de preencher o formulário.
        """
        pis_limpo = re.sub(r"\D", "", str(pis or ""))
        if pis_limpo:
            pis_limpo = pis_limpo.zfill(11)
        nome_limpo = " ".join(str(nome or "").split()).strip().upper()

        print(f"[*] Consultando PIS: {pis_limpo} | Base: {base_uf}...")
        dados = self._executar_pesquisa_formulario(pis=pis_limpo, nome="", base_uf=base_uf)

        status = str(dados.get("status") or "").upper()
        if "SUCESSO" not in status and nome_limpo:
            print(f"[!] PIS {pis_limpo} não localizado diretamente. Acionando Fallback de pesquisa por Nome: '{nome_limpo}'...")
            self.retornar_para_formulario()
            time.sleep(1)
            dados_nome = self._executar_pesquisa_formulario(pis="", nome=nome_limpo, base_uf=base_uf)
            if "SUCESSO" in str(dados_nome.get("status") or "").upper():
                print(f"[+] Sucesso no Fallback por Nome para '{nome_limpo}'! Conta localizada.")
                dados = dados_nome
                dados["pis_pesquisado"] = pis_limpo
                dados["mensagem"] = "Localizado com sucesso via Fallback por Nome"
            else:
                dados["mensagem"] = f"Não localizado por PIS ({dados.get('mensagem')}) nem por Nome ({dados_nome.get('mensagem')})"

        # Se a consulta obteve sucesso, gerar o PDF Oficial timbrado da Caixa
        status = str(dados.get("status") or "").upper()
        if "SUCESSO" in status:
            nome_colab = dados.get("nome_caixa") or nome_limpo
            caminho_pdf = self.gerar_extrato_pdf_oficial(pis=pis_limpo, base_uf=base_uf, nome=nome_colab)
            if caminho_pdf:
                dados["caminho_pdf"] = caminho_pdf

        return dados

    def gerar_extrato_pdf_oficial(self, pis: str, base_uf: str, nome: str = "") -> Optional[str]:
        """
        Dispara a funcionalidade oficial 'Visualizar Impressão' da Caixa (impr_solicitar_extrato_fgts),
        captura o documento timbrado oficial e converte em arquivo PDF real via CDP do Chrome.
        Retorna o caminho absoluto do arquivo PDF gerado.
        """
        janela_original = self.driver.current_window_handle
        nome_saneado = re.sub(r"[^A-Za-z0-9]", "_", nome.strip()).strip("_") if nome else ""
        sufixo = f"_{nome_saneado}" if nome_saneado else ""
        caminho_pdf = self.diretorio_saida / f"Extrato_FGTS_{pis}_{base_uf}{sufixo}.pdf"

        try:
            janelas_antes = self.driver.window_handles

            # 1. Acionar a função oficial de impressão da Caixa
            self.driver.execute_script("""
                if (typeof impr_solicitar_extrato_fgts === 'function') {
                    impr_solicitar_extrato_fgts();
                }
            """)

            # Aguardar abertura da nova janela/aba de impressão
            try:
                WebDriverWait(self.driver, 6).until(lambda d: len(d.window_handles) > len(janelas_antes))
            except TimeoutException:
                try:
                    elem_img = self.driver.find_element(By.XPATH, "//a[contains(@href, 'impr_solicitar_extrato_fgts')] | //img[contains(@src, 'visualizar_impressao')]")
                    self.driver.execute_script("arguments[0].click();", elem_img)
                    WebDriverWait(self.driver, 5).until(lambda d: len(d.window_handles) > len(janelas_antes))
                except Exception:
                    pass

            # Identificar e mudar para a aba de impressão
            janela_impressao = None
            for h in self.driver.window_handles:
                if h != janela_original:
                    janela_impressao = h
                    break

            if not janela_impressao:
                print(f"[!] Aba de impressão não detectada para PIS {pis}.")
                return None

            self.driver.switch_to.window(janela_impressao)
            time.sleep(1.5)  # Aguardar renderização completa do extrato timbrado

            # 2. Gerar PDF nativo via Chrome DevTools Protocol (CDP)
            resultado_cdp = self.driver.execute_cdp_cmd("Page.printToPDF", {
                "printBackground": True,
                "preferCSSPageSize": True,
                "landscape": False,
                "marginTop": 0.4,
                "marginBottom": 0.4,
                "marginLeft": 0.4,
                "marginRight": 0.4
            })

            bytes_pdf = base64.b64decode(resultado_cdp["data"])
            with open(caminho_pdf, "wb") as f:
                f.write(bytes_pdf)

            print(f"[+] Documento Oficial PDF exportado com sucesso: {caminho_pdf.name}")

            # 3. Fechar aba de impressão e retornar para a tela da consulta
            self.driver.close()
            self.driver.switch_to.window(janela_original)
            return str(caminho_pdf.resolve())

        except Exception as e:
            print(f"[!] Erro ao gerar PDF oficial para PIS {pis}: {e}")
            try:
                if self.driver.current_window_handle != janela_original:
                    self.driver.close()
                self.driver.switch_to.window(janela_original)
            except Exception:
                pass
            return None

    def retornar_para_formulario(self):
        """
        Retorna da tela de extrato para o formulário de pesquisa ('Localização do Trabalhador').
        Executa a função oficial retr_solicitar_extrato_fgts() da Caixa.
        """
        # Se já estiver no formulário de busca, apenas limpa os campos de texto
        if self._esta_no_formulario_busca():
            try:
                for inp in self.driver.find_elements(By.NAME, "txtNumPisPasep"):
                    inp.clear()
                for inp in self.driver.find_elements(By.NAME, "txtNomeTrab"):
                    inp.clear()
                return
            except Exception:
                pass

        print("[*] Retornando para o formulário de pesquisa...")

        # 1. Executar a função oficial da tela de extrato da Caixa: retr_solicitar_extrato_fgts()
        try:
            self.driver.execute_script("""
                if (typeof retr_solicitar_extrato_fgts === 'function') {
                    retr_solicitar_extrato_fgts();
                } else if (typeof go === 'function') {
                    go('ExtratoFgts.AcessarBaseConta');
                }
            """)
            time.sleep(2)
        except Exception as e:
            print(f"[!] Tentativa JS retornar: {e}")

        # 2. Se ainda não voltou, clicar no link da imagem botao_retornar
        if not self._esta_no_formulario_busca():
            seletores_retornar = [
                (By.XPATH, "//a[contains(@href, 'retr_solicitar_extrato_fgts')]"),
                (By.XPATH, "//img[contains(@src, 'botao_retornar')]/parent::a"),
                (By.XPATH, "//img[contains(@src, 'botao_retornar')]"),
                (By.XPATH, "//a[contains(@href, 'AcessarBaseConta')]"),
                (By.XPATH, "//a[contains(@href, 'Principal.Visualizar')]")
            ]
            for by, sel in seletores_retornar:
                try:
                    elem = self.driver.find_element(by, sel)
                    if elem:
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                        time.sleep(0.3)
                        self.driver.execute_script("arguments[0].click();", elem)
                        time.sleep(2)
                        break
                except Exception:
                    continue

        # 3. Aguardar o formulário reaparecer e limpar os campos
        try:
            WebDriverWait(self.driver, 8).until(
                lambda d: len(d.find_elements(By.NAME, "txtNumPisPasep")) > 0
            )
            for inp in self.driver.find_elements(By.NAME, "txtNumPisPasep"):
                inp.clear()
            for inp in self.driver.find_elements(By.NAME, "txtNomeTrab"):
                inp.clear()
            print("[+] Retorno ao formulário concluído com sucesso!")
        except Exception:
            self.navegar_para_solicitar_extrato()
