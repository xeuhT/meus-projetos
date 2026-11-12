"""
Orquestrador Principal RPA - Extração de Extrato de Compras do Usuário (GoodManager / Ticket Plus)

Fluxo Automatizado:
1. Acesso direto a ExtratoForm.cfm e autenticação automática caso redirecionado.
2. Consulta por matrícula garantindo os filtros 'Usuários = Todos (T)' e 'Situação = Todos (T)'.
3. Fallback automático para 'Opções de Acesso' -> 'PLUS CENCOSUD CHIP' caso a matrícula não seja localizada na base atual.
4. Seleção exclusiva do cartão que contenha simultaneamente '(A)' (Ativo) e '(T)' (Titular).
5. Extração em PDF (Mês Vigente e Compras Próximo Período) em pastas individuais por colaborador.
"""

import argparse
import base64
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from playwright.sync_api import BrowserContext, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

from utils import (
    GOODMANAGER_COD_USUARIO,
    GOODMANAGER_NOME_USUARIO,
    GOODMANAGER_SENHA,
    OUTPUT_DIR,
    get_employee_output_dir,
    get_logger,
    get_report_pdf_path,
    load_matriculas,
    load_matriculas_e_nomes,
    verificar_cache_colaborador,
)
from consolidar_excel import gerar_consolidado_excel

EXTRATO_FORM_URL = "https://www.goodmanager.com.br/GoodManagerSSL/Beneficio/ExtratoForm.cfm"
TROCA_FILTROS_URL = "https://www.goodmanager.com.br/GoodManagerSSL/TrocaFiltros.cfm"
CODIGO_CLIENTE_PLUS_CHIP = "163631"  # 163631 - PLUS CENCOSUD CHIP

logger = get_logger()


def salvar_pagina_como_pdf(page: Page, caminho_pdf: Path) -> None:
    """
    Salva a página aberta (relatório) como PDF A4.
    Funciona tanto em modo headless quanto em modo visível (headed) via CDP Page.printToPDF.
    """
    caminho_pdf.parent.mkdir(parents=True, exist_ok=True)
    try:
        # Tenta método nativo do Playwright (funciona em headless)
        page.pdf(
            path=str(caminho_pdf),
            format="A4",
            print_background=True,
            margin={"top": "10mm", "bottom": "10mm", "left": "10mm", "right": "10mm"},
        )
    except Exception:
        # Fallback via Chrome DevTools Protocol (funciona também com navegador visível headless=False)
        cdp = page.context.new_cdp_session(page)
        pdf_data = cdp.send(
            "Page.printToPDF",
            {
                "printBackground": True,
                "paperWidth": 8.27,   # A4 width in inches
                "paperHeight": 11.69, # A4 height in inches
                "marginTop": 0.39,
                "marginBottom": 0.39,
                "marginLeft": 0.39,
                "marginRight": 0.39,
            },
        )
        with open(caminho_pdf, "wb") as f:
            f.write(base64.b64decode(pdf_data["data"]))


def autenticar_se_necessario(page: Page) -> None:
    """
    Acessa a URL direta do ExtratoForm.cfm.
    Se redirecionar para a tela de login, realiza a autenticação com os 3 campos obrigatórios
    e garante que ao final esteja posicionado em ExtratoForm.cfm.
    """
    if not all([GOODMANAGER_COD_USUARIO, GOODMANAGER_NOME_USUARIO, GOODMANAGER_SENHA]):
        raise ValueError("Credenciais incompletas no arquivo .env (GOODMANAGER_COD_USUARIO, GOODMANAGER_NOME_USUARIO, GOODMANAGER_SENHA).")

    logger.info(f"Acessando {EXTRATO_FORM_URL} ...")
    page.goto(EXTRATO_FORM_URL, wait_until="domcontentloaded", timeout=60000)

    # Verifica se caiu na tela de login
    if "autenticacao" in page.url.lower() or page.locator("input#codigo").count() > 0:
        logger.info("Sessão não ativa detectada. Realizando login no portal GoodManager...")
        page.locator("input#codigo").fill(GOODMANAGER_COD_USUARIO)
        page.locator("input#usuario").fill(GOODMANAGER_NOME_USUARIO)
        page.locator("input#senha").fill(GOODMANAGER_SENHA)

        with page.expect_navigation(wait_until="domcontentloaded", timeout=60000):
            page.locator("input#aem-login").click()

        logger.info(f"Login efetuado com sucesso. URL pós-login: {page.url}")

    # Pós-Login: Caso não tenha redirecionado automaticamente para ExtratoForm.cfm, navega até lá
    if "extratoform.cfm" not in page.url.lower():
        logger.info("Redirecionando para a tela de Extrato de Compras do Usuário (ExtratoForm.cfm)...")
        page.goto(EXTRATO_FORM_URL, wait_until="domcontentloaded", timeout=60000)


def trocar_base_cliente(page: Page, codigo_cliente: str, nome_base: str) -> bool:
    """
    Acessa 'Opções de Acesso' (barra superior / TrocaFiltros.cfm),
    seleciona o cliente indicado (ex: '163631' -> PLUS CENCOSUD CHIP ou '200602' -> CENCOSUD ATACADO)
    e confirma em Continuar.
    """
    logger.info(f"Acessando 'Opções de Acesso' para alterar base para '{nome_base}' ({codigo_cliente})...")
    link_opcoes = page.locator("a.menu-opcoes-acesso")
    if link_opcoes.count() > 0 and link_opcoes.is_visible():
        with page.expect_navigation(wait_until="domcontentloaded", timeout=60000):
            link_opcoes.first.click()
    else:
        page.goto(TROCA_FILTROS_URL, wait_until="domcontentloaded", timeout=60000)

    # Localiza o item da base na lista de clientes
    item_base = page.locator(f"li[data-value='{codigo_cliente}']")
    if item_base.count() == 0:
        item_base = page.locator("ul#lista-clientes li", has_text=nome_base)

    if item_base.count() == 0:
        logger.warning(f"Não foi possível localizar a opção '{nome_base}' ({codigo_cliente}) em Opções de Acesso.")
        return False

    item_base.first.click()

    with page.expect_navigation(wait_until="domcontentloaded", timeout=60000):
        page.locator("input#btContinuarTC").click()

    logger.info(f"Base alterada com sucesso para '{nome_base}'.")
    return True


def alterar_opcoes_acesso_para_plus_chip(page: Page) -> None:
    """Mantém compatibilidade com chamadas legadas."""
    trocar_base_cliente(page, CODIGO_CLIENTE_PLUS_CHIP, "PLUS CENCOSUD CHIP")


def executar_busca_matricula(page: Page, matricula: str) -> bool:
    """
    Etapa 2: Na tela ExtratoForm.cfm, preenche a matrícula, valida obrigatoriamente se
    'Usuários' e 'Situação' estão em 'Todos' (value='T'), clica em Continuar e retorna
    True se o colaborador foi localizado, ou False se não houver itens.
    """
    if "extratoform.cfm" not in page.url.lower() or page.locator("input#nr_matricula").count() == 0:
        page.goto(EXTRATO_FORM_URL, wait_until="domcontentloaded", timeout=60000)

    campo_matricula = page.locator("input#nr_matricula")
    campo_matricula.wait_for(state="visible", timeout=30000)
    campo_matricula.fill("")
    campo_matricula.fill(str(matricula).strip())

    # Validação Obrigatória: Usuários = Todos ('T') e Situação = Todos ('T')
    select_usuarios = page.locator("select#fl_usuarios")
    val_usuarios = select_usuarios.input_value()
    if val_usuarios != "T":
        logger.info(f"Ajustando filtro 'Usuários' de '{val_usuarios}' para 'Todos' (T).")
        select_usuarios.select_option("T")

    select_situacao = page.locator("select#fl_situacao")
    val_situacao = select_situacao.input_value()
    if val_situacao != "T":
        logger.info(f"Ajustando filtro 'Situação' de '{val_situacao}' para 'Todos' (T).")
        select_situacao.select_option("T")

    # Clica em Continuar
    btn_continuar = page.locator("input[type='submit'][value='Continuar']").first
    with page.expect_navigation(wait_until="domcontentloaded", timeout=60000):
        btn_continuar.click()

    # Verifica se encontrou resultados (presença de form[name='Form'] e ausência de 'Não há itens para exibição')
    conteudo_pagina = page.content()
    if "Não há itens para exibição" in conteudo_pagina or page.locator("form[name='Form']").count() == 0:
        return False

    return True


def selecionar_cartao_ativo_titular(page: Page, matricula: str) -> Tuple[Optional[int], Optional[str]]:
    """
    Etapa 4.1: Analisa as linhas de cartão na tela ExtratoLista.cfm.
    Seleciona SEMPRE o cartão que contenha simultaneamente '(A)' (Ativo) e '(T)' (Titular),
    desmarcando eventuais outros.
    Retorna (indice_nth_do_checkbox, nome_do_colaborador) ou (None, None) se não encontrar.
    """
    form = page.locator("form[name='Form']")
    checkboxes = form.locator("input[type='checkbox'][name^='cd']")
    total_cb = checkboxes.count()

    cartao_alvo_idx: Optional[int] = None
    nome_colaborador: Optional[str] = None
    cartoes_encontrados_debug: List[str] = []

    for i in range(total_cb):
        cb = checkboxes.nth(i)
        # Obtém a linha (tr) correspondente a este checkbox
        linha = cb.locator("xpath=ancestor::tr[1]")
        texto_linha = " ".join(linha.inner_text().split())
        cartoes_encontrados_debug.append(texto_linha)

        # Extrai nome do colaborador (geralmente 3ª coluna td da linha)
        tds = linha.locator("td")
        if tds.count() >= 3 and not nome_colaborador:
            possivel_nome = tds.nth(2).inner_text().strip()
            if possivel_nome:
                nome_colaborador = possivel_nome

        # Regra: conter simultaneamente (A) e (T)
        if "(A)" in texto_linha and "(T)" in texto_linha and cartao_alvo_idx is None:
            cartao_alvo_idx = i
            if not cb.is_checked():
                cb.check()
            logger.info(f"[Matrícula {matricula}] Cartão (A) e (T) selecionado: {texto_linha}")
        else:
            if cb.is_checked():
                cb.uncheck()

    if cartao_alvo_idx is None:
        logger.warning(
            f"[Matrícula {matricula}] Nenhum cartão com '(A)' e '(T)' simultâneos foi encontrado. "
            f"Opções listadas: {cartoes_encontrados_debug}"
        )
        return None, nome_colaborador

    return cartao_alvo_idx, nome_colaborador


def obter_opcoes_periodo(page: Page) -> List[Tuple[str, str]]:
    """
    Etapa 4.2: Identifica no dropdown select#dt_inicial as duas opções requeridas:
    1. O mês vigente (ex: Outubro/2026 - opção selecionada por padrão / mais recente '|S')
    2. 'Compras Próximo Período' (opção '|N')
    Retorna lista de tuplas (option_value, label_para_arquivo).
    """
    select_periodo = page.locator("select#dt_inicial")
    valor_padrao = (select_periodo.input_value() or "").strip()
    options = select_periodo.locator("option")
    total_opt = options.count()

    opcoes_s: List[Tuple[str, str]] = []
    opcao_proximo_periodo: Optional[Tuple[str, str]] = None

    for i in range(total_opt):
        opt = options.nth(i)
        val = (opt.get_attribute("value") or "").strip()
        txt = " ".join(opt.inner_text().split()).strip()

        if "próximo período" in txt.lower() or "proximo periodo" in txt.lower() or val.endswith("|N"):
            if not opcao_proximo_periodo:
                opcao_proximo_periodo = (val, "Compras_Proximo_Periodo")
        elif val.endswith("|S"):
            label_mes = txt.replace("/", "_").replace(" ", "_")
            opcoes_s.append((val, f"Mes_Atual_{label_mes}"))

    # Prioriza a opção |S que já vem selecionada por padrão no sistema (mês vigente)
    # ou a mais recente (última da lista ordenada por YYYYMM01|S)
    opcao_mes_vigente: Optional[Tuple[str, str]] = None
    for val, label in opcoes_s:
        if val == valor_padrao:
            opcao_mes_vigente = (val, label)
            break
    if not opcao_mes_vigente and opcoes_s:
        opcao_mes_vigente = sorted(opcoes_s, key=lambda x: x[0])[-1]

    resultado: List[Tuple[str, str]] = []
    if opcao_mes_vigente:
        resultado.append(opcao_mes_vigente)
    if opcao_proximo_periodo:
        resultado.append(opcao_proximo_periodo)

    return resultado


def processar_matricula(
    context: BrowserContext,
    page: Page,
    matricula: str,
    estado_sessao: Dict[str, str],
) -> bool:
    """
    Executa o fluxo completo (Etapas 2 a 5) para uma matrícula específica.
    Retorna True se extraiu os relatórios com sucesso, False caso contrário.
    """
    # 0. Otimização de Cache: se já tiver os 2 PDFs na pasta, pula re-extração
    pasta_cache = verificar_cache_colaborador(matricula)
    if pasta_cache:
        logger.info(f"[Matrícula {matricula}] Já possui os 2 PDFs extraídos na pasta '{pasta_cache.name}'. Pulando re-extração (cache ativo)...")
        return True

    logger.info(f"--- Iniciando processamento da matrícula: {matricula} (Base atual: {estado_sessao['cliente']}) ---")

    encontrou = executar_busca_matricula(page, matricula)

    # Etapa 3: Tratamento de Exceção com Fallback Bidirecional de Base
    bases_alternativas = [
        ("163631", "PLUS CENCOSUD CHIP"),
        ("200602", "CENCOSUD ATACADO"),
        ("200600", "PERINI COMERCIAL CHIP"),
    ]

    if not encontrou:
        for cod_base, nome_base in bases_alternativas:
            if estado_sessao["cliente"] != nome_base:
                logger.warning(
                    f"[Matrícula {matricula}] Não localizada na base '{estado_sessao['cliente']}'. "
                    f"Tentando fallback para '{nome_base}' ({cod_base})..."
                )
                ok_troca = trocar_base_cliente(page, cod_base, nome_base)
                if ok_troca:
                    estado_sessao["cliente"] = nome_base
                    encontrou = executar_busca_matricula(page, matricula)
                    if encontrou:
                        logger.info(f"[Matrícula {matricula}] Localizada com sucesso na base '{nome_base}'!")
                        break

    if not encontrou:
        logger.warning(
            f"[Matrícula {matricula}] Não localizada em nenhuma das bases disponíveis. Pulando..."
        )
        if estado_sessao["cliente"] != "PLUS CENCOSUD CHIP":
            trocar_base_cliente(page, "163631", "PLUS CENCOSUD CHIP")
            estado_sessao["cliente"] = "PLUS CENCOSUD CHIP"
        return False

    # Etapa 4: Seleção de Cartão (A) e (T)
    cartao_cb_idx, nome_colaborador = selecionar_cartao_ativo_titular(page, matricula)
    if cartao_cb_idx is None:
        return False

    # Define identificador da pasta do colaborador
    if nome_colaborador:
        identificador_pasta = f"{matricula} - {nome_colaborador}"
    else:
        identificador_pasta = str(matricula)

    pasta_colaborador = get_employee_output_dir(identificador_pasta)
    logger.info(f"[Matrícula {matricula}] Pasta de destino: {pasta_colaborador}")

    # Obtém os 2 períodos (Mês Vigente e Compras Próximo Período)
    periodos = obter_opcoes_periodo(page)
    if not periodos:
        logger.error(f"[Matrícula {matricula}] Não foi possível identificar os períodos em select#dt_inicial.")
        return False

    # Etapa 5: Para cada período, seleciona, abre a nova aba/popup e salva o PDF
    for valor_periodo, label_periodo in periodos:
        logger.info(f"[Matrícula {matricula}] Extraindo período '{label_periodo}' (valor={valor_periodo})...")

        # Garante que apenas o checkbox do cartão (A)+(T) continua marcado
        checkboxes = page.locator("form[name='Form'] input[type='checkbox'][name^='cd']")
        cb_alvo = checkboxes.nth(cartao_cb_idx)
        if not cb_alvo.is_checked():
            cb_alvo.check()

        page.locator("select#dt_inicial").select_option(valor_periodo)

        btn_continuar_lista = page.locator("form[name='Form'] input[type='submit'][value='Continuar']").first

        with context.expect_page(timeout=60000) as nova_aba_info:
            btn_continuar_lista.click()

        aba_relatorio = nova_aba_info.value
        try:
            aba_relatorio.wait_for_load_state("domcontentloaded", timeout=60000)
            # Aguarda estabilização visual da tabela do extrato
            try:
                aba_relatorio.wait_for_load_state("networkidle", timeout=15000)
            except PlaywrightTimeoutError:
                pass

            caminho_pdf = get_report_pdf_path(identificador_pasta, f"{matricula}_{label_periodo}")
            salvar_pagina_como_pdf(aba_relatorio, caminho_pdf)
            logger.info(f"[Matrícula {matricula}] PDF salvo com sucesso: {caminho_pdf.name}")
        finally:
            if not aba_relatorio.is_closed():
                aba_relatorio.close()

    return True


def executar_robo(matriculas: List[str], headless: bool = True) -> Dict[str, List[str]]:
    """
    Inicializa o navegador Playwright, realiza login e processa toda a fila de matrículas.
    """
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    resumo = {"sucesso": [], "falha": []}
    estado_sessao = {"cliente": "PADRAO"}

    logger.info("=" * 70)
    logger.info(f"Iniciando Robô Ticket Plus (GoodManager) | Total de matrículas: {len(matriculas)} | Headless: {headless}")
    logger.info(f"Diretório de saída: {OUTPUT_DIR}")
    logger.info("=" * 70)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(
            viewport={"width": 1366, "height": 768},
            ignore_https_errors=True,
        )

        # Bloqueia scripts de tour/overlay de terceiros (UserPilot) que podem interceptar cliques
        context.route("**/*userpilot*", lambda route: route.abort())

        page = context.new_page()

        try:
            # Etapa 1: Acesso e Autenticação
            autenticar_se_necessario(page)
            alterar_opcoes_acesso_para_plus_chip(page)
            estado_sessao["cliente"] = "PLUS CENCOSUD CHIP"

            # Processa cada matrícula
            for idx, matricula in enumerate(matriculas, start=1):
                logger.info(f"[{idx}/{len(matriculas)}] Processando matrícula {matricula}...")
                try:
                    ok = processar_matricula(context, page, matricula, estado_sessao)
                    if ok:
                        resumo["sucesso"].append(matricula)
                    else:
                        resumo["falha"].append(matricula)
                except Exception as e:
                    logger.error(f"Erro inesperado ao processar matrícula {matricula}: {e}", exc_info=True)
                    resumo["falha"].append(matricula)

        finally:
            context.close()
            browser.close()

    logger.info("=" * 70)
    logger.info(
        f"Execução finalizada! Sucessos ({len(resumo['sucesso'])}): {resumo['sucesso']} | "
        f"Falhas/Não encontradas ({len(resumo['falha'])}): {resumo['falha']}"
    )
    logger.info("=" * 70)

    # Consolidação Automática em Planilha Excel
    try:
        logger.info("Iniciando consolidação automática dos extratos em planilha Excel...")
        caminho_excel = gerar_consolidado_excel()
        logger.info(f"Planilha consolidada atualizada com sucesso: {caminho_excel.name}")
    except Exception as e:
        logger.error(f"Erro ao gerar consolidação em Excel ao final da execução: {e}", exc_info=True)

    return resumo


def main() -> None:
    parser = argparse.ArgumentParser(description="Robô de Extração de Extrato de Compras GoodManager (Ticket Plus)")
    parser.add_argument(
        "-m", "--matriculas",
        nargs="+",
        help="Lista de matrículas separadas por espaço (ex: -m 7887 10733 11482)",
    )
    parser.add_argument(
        "-f", "--arquivo",
        type=str,
        help="Caminho para arquivo .txt, .csv ou .xlsx contendo as matrículas",
    )
    parser.add_argument(
        "--visivel",
        action="store_true",
        help="Executa o navegador em modo visível (headed) em vez de segundo plano (headless)",
    )
    args = parser.parse_args()

    if args.matriculas:
        lista_matriculas = [str(m).strip() for m in args.matriculas if str(m).strip()]
    else:
        lista_matriculas = load_matriculas(args.arquivo)

    if not lista_matriculas:
        logger.error("Nenhuma matrícula informada ou encontrada em matriculas.txt/.csv/.xlsx.")
        sys.exit(1)

    executar_robo(matriculas=lista_matriculas, headless=not args.visivel)


if __name__ == "__main__":
    main()
