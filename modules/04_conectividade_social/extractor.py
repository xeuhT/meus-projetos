"""
Módulo Extrator e Parser de Dados do Extrato FGTS
Conectividade Social V2 - Cencosud Brasil
Extrai campos cadastrais, Valor Base para Fins Rescisórios, Saldo, Afastamentos/Movimentações e histórico.
"""
import re
from typing import Dict, Any, Optional
from html.parser import HTMLParser


def limpar_texto(valor: Optional[str]) -> str:
    """Remove quebras de linha, entidades HTML e múltiplos espaços."""
    if not valor:
        return ""
    texto = valor.replace("\xa0", " ").replace("&nbsp;", " ").replace("&NBSP;", " ")
    return " ".join(texto.split()).strip()


def converter_valor_monetario(valor_str: Optional[str]) -> Optional[float]:
    """
    Converte strings monetárias no padrão brasileiro ('R$ 206,05' ou '1.250,50') para float.
    Retorna None se não for um valor válido.
    """
    if not valor_str:
        return None
    texto = limpar_texto(valor_str)
    match = re.search(r"([\d\.,-]+)", texto)
    if not match:
        return None
    raw = match.group(1).replace(".", "").replace(",", ".")
    try:
        return float(raw)
    except ValueError:
        return None


class SimpleHTMLTextExtractor(HTMLParser):
    """Extrai texto ordenado de HTML preservando quebras de linha de blocos estruturais."""
    def __init__(self):
        super().__init__()
        self.chunks = []
        self._block_tags = {'div', 'p', 'tr', 'td', 'th', 'br', 'li', 'h1', 'h2', 'h3', 'table'}

    def handle_data(self, data):
        cleaned = data.strip()
        if cleaned:
            self.chunks.append(cleaned)

    def handle_starttag(self, tag, attrs):
        if tag.lower() in self._block_tags:
            self.chunks.append("\n")

    def handle_endtag(self, tag):
        if tag.lower() in self._block_tags:
            self.chunks.append("\n")

    def get_text(self) -> str:
        return " ".join(self.chunks)


def extrair_pares_tabela(html: str) -> Dict[str, str]:
    """
    Extrai pares chave: valor de tabelas HTML (<tr><td>Rótulo</td><td>Valor</td></tr>).
    """
    pares = {}
    linhas = re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.DOTALL | re.IGNORECASE)
    for linha in linhas:
        celulas = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", linha, re.DOTALL | re.IGNORECASE)
        celulas_limpas = [limpar_texto(re.sub(r"<[^>]+>", " ", c)) for c in celulas if limpar_texto(re.sub(r"<[^>]+>", " ", c))]
        
        for i in range(0, len(celulas_limpas) - 1, 2):
            chave = celulas_limpas[i].rstrip(":").strip()
            valor = celulas_limpas[i + 1].strip()
            if chave and valor:
                pares[chave.upper()] = valor

    return pares


def extrair_dados_extrato_html(html_conteudo: str) -> Dict[str, Any]:
    """
    Parser robusto do HTML do Extrato FGTS do Trabalhador.
    Combina extração direta de células da tabela HTML com busca regex no texto plano.
    """
    dados: Dict[str, Any] = {
        "nome_caixa": "",
        "pis_caixa": "",
        "empresa": "",
        "cnpj": "",
        "cod_estab": "",
        "categoria": "",
        "conta_fgts": "",
        "data_admissao": "",
        "data_opcao": "",
        "data_movimentacao": "",
        "cod_movimentacao": "",
        "taxa_juros": "",
        "tipo_conta": "",
        "valor_base_rescisorios": None,
        "valor_base_rescisorios_fmt": "",
        "saldo_fgts": None,
        "saldo_fgts_fmt": "",
        "base_caixa": "",
        "data_atualizacao": "",
        "status": "SUCESSO",
        "mensagem": "Extrato extraído com sucesso"
    }

    # Estratégia 1: Extrair pares chave-valor das células de tabelas
    pares = extrair_pares_tabela(html_conteudo)

    mapa_chaves = {
        "NOME": "nome_caixa",
        "PIS/PASEP/NIT": "pis_caixa",
        "PIS": "pis_caixa",
        "EMPRESA": "empresa",
        "CNPJ/CEI/CPF": "cnpj",
        "CNPJ": "cnpj",
        "CÓD. ESTAB.": "cod_estab",
        "COD. ESTAB.": "cod_estab",
        "CÓD. ESTAB": "cod_estab",
        "COD. ESTAB": "cod_estab",
        "CATEGORIA": "categoria",
        "Nº CONTA FGTS": "conta_fgts",
        "NO CONTA FGTS": "conta_fgts",
        "CONTA FGTS": "conta_fgts",
        "DATA ADMISSÃO": "data_admissao",
        "DATA ADMISSAO": "data_admissao",
        "DATA OPÇÃO": "data_opcao",
        "DATA OPCAO": "data_opcao",
        "DATA/CÓD. MOVIMENTAÇÃO": "data_cod_movimentacao",
        "DATA/COD. MOVIMENTACAO": "data_cod_movimentacao",
        "DATA/CÓD MOVIMENTAÇÃO": "data_cod_movimentacao",
        "DATA/COD MOVIMENTACAO": "data_cod_movimentacao",
        "TAXA JUROS": "taxa_juros",
        "TIPO CONTA": "tipo_conta",
        "VALOR BASE PARA FINS RESCISÓRIOS": "valor_base_rescisorios",
        "VALOR BASE PARA FINS RESCISORIOS": "valor_base_rescisorios",
        "BASE": "base_caixa",
        "SALDO": "saldo_fgts",
        "ATUALIZADO EM": "data_atualizacao"
    }

    for chave_tab, valor in pares.items():
        chave_normalizada = chave_tab.replace(":", "").strip()
        campo_destino = mapa_chaves.get(chave_normalizada)
        if campo_destino:
            if campo_destino in ["valor_base_rescisorios", "saldo_fgts"]:
                dados[f"{campo_destino}_fmt"] = valor
                dados[campo_destino] = converter_valor_monetario(valor)
            elif campo_destino == "data_cod_movimentacao":
                # Tratar campo composto Data/Cód. Movimentação (ex: "22/09/2026 - J" ou "01/10/2026 - I1" ou "-")
                valor_limpo = limpar_texto(valor)
                match_mov = re.search(r"(\d{2}/\d{2}/\d{4})\s*[-/]?\s*([A-Za-z0-9]+)", valor_limpo)
                if match_mov:
                    dados["data_movimentacao"] = match_mov.group(1).strip()
                    dados["cod_movimentacao"] = match_mov.group(2).strip().upper()
                else:
                    codigo_puro = re.sub(r"[^A-Za-z0-9]", "", valor_limpo)
                    dados["cod_movimentacao"] = codigo_puro.upper() if codigo_puro else "-"
            else:
                dados[campo_destino] = valor

    # Estratégia 2: Regex sobre texto plano (fallback ou reforço)
    parser = SimpleHTMLTextExtractor()
    try:
        parser.feed(html_conteudo)
        texto_plano = parser.get_text()
    except Exception:
        texto_plano = re.sub(r"<[^>]+>", " ", html_conteudo)

    padroes_texto = {
        "nome_caixa": r"Nome:\s*([A-ZÀ-Ú\s]+?)(?:\s+PIS|\n|$)",
        "pis_caixa": r"PIS(?:/PASEP/NIT)?:\s*([\d\.-]+)",
        "empresa": r"Empresa:\s*([A-ZÀ-Ú0-9\s\.-]+?)(?:\s+CNPJ|\n|$)",
        "cnpj": r"CNPJ(?:/CEI/CPF)?:\s*([\d\./-]+)",
        "cod_estab": r"C[oó]d\.?\s*Estab\.?:?\s*([\d]+)",
        "categoria": r"Categoria:\s*([\d]+)",
        "conta_fgts": r"N[ºo°]\s*Conta\s*FGTS:\s*([\d]+)",
        "data_admissao": r"Data\s*Admiss[aã]o:\s*([\d]{2}/[\d]{2}/[\d]{4})",
        "data_opcao": r"Data\s*Op[cç][aã]o:\s*([\d]{2}/[\d]{2}/[\d]{4})",
        "taxa_juros": r"Taxa\s*Juros:\s*([\d\s%]+)",
        "tipo_conta": r"Tipo\s*Conta:\s*([A-Z]+)",
        "valor_base_rescisorios": r"Valor\s*Base\s*para\s*Fins\s*Rescis[oó]rios:\s*(?:R\$\s*)?([\d\.,]+)",
        "base_caixa": r"Base:\s*([A-Z]{2})",
        "saldo_fgts": r"SALDO:\s*(?:R\$\s*)?([\d\.,-]+)",
        "data_atualizacao": r"Atualizado\s*em:\s*([\d]{2}/[\d]{2}/[\d]{4})"
    }

    # Tratamento de regex para Data/Cód. Movimentação
    if not dados["cod_movimentacao"]:
        match_mov = re.search(r"Data/C[oó]d\.?\s*Movimenta[cç][aã]o:\s*([\d]{2}/[\d]{2}/[\d]{4})?\s*(?:/)?\s*([A-Z0-9-]+)?", texto_plano, re.IGNORECASE)
        if match_mov:
            dados["data_movimentacao"] = match_mov.group(1) or ""
            dados["cod_movimentacao"] = match_mov.group(2) or "-"

    # Correção do Ponto 8: Validação estrita para não sobrescrever valor zerado (0.0)
    for campo, regex in padroes_texto.items():
        if dados[campo] is None or dados[campo] == "":
            match = re.search(regex, texto_plano, re.IGNORECASE)
            if match:
                val = limpar_texto(match.group(1))
                if campo in ["valor_base_rescisorios", "saldo_fgts"]:
                    dados[f"{campo}_fmt"] = f"R$ {val}"
                    dados[campo] = converter_valor_monetario(val)
                else:
                    dados[campo] = val

    # Verificar mensagens de erro ou conta inexistente
    if dados["valor_base_rescisorios"] is None and "Fins Rescisórios" not in texto_plano:
        msg_erro = detectar_mensagem_erro(texto_plano)
        if msg_erro:
            dados["status"] = "NÃO LOCALIZADO"
            dados["mensagem"] = msg_erro
        elif "Localização do Trabalhador" in texto_plano:
            dados["status"] = "NÃO LOCALIZADO"
            dados["mensagem"] = "Trabalhador não localizado na base informada"

    return dados


def detectar_mensagem_erro(texto_ou_html: str) -> Optional[str]:
    """Detecta alertas e mensagens de retorno do portal Conectividade Social."""
    mensagens_conhecidas = [
        r"Conta n[aã]o localizada",
        r"Trabalhador n[aã]o localizado",
        r"N[aã]o existem contas para o trabalhador",
        r"PIS/PASEP/NIT inv[aá]lido",
        r"Base da conta inv[aá]lida",
        r"Nenhum registro encontrado",
        r"Ocorreu um erro no processamento",
        r"Certificado inv[aá]lido ou expirado"
    ]

    for padrao in mensagens_conhecidas:
        match = re.search(padrao, texto_ou_html, re.IGNORECASE)
        if match:
            return match.group(0).strip()
    return None
