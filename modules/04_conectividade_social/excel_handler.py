"""
Módulo de Tratamento e Exportação de Planilhas Excel
Conectividade Social V2 - Cencosud Brasil
Garante conformidade com as regras de integridade técnica e padrão visual institucional.
"""
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from typing import List, Dict, Any, Optional
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config import DIR_SAIDA, CORES_LAYOUT, BASE_CONTA_PADRAO
from business_rules import auditar_situacao_colaborador


def remover_acentos(texto: str) -> str:
    """Normaliza texto removendo acentos e caracteres especiais para comparação segura."""
    if not texto:
        return ""
    nfkd = unicodedata.normalize('NFKD', str(texto))
    return "".join(c for c in nfkd if not unicodedata.combining(c)).strip().upper()


def normalizar_pis(valor: Any) -> str:
    """Normaliza o PIS para string de 11 dígitos mantendo zeros à esquerda."""
    if valor is None:
        return ""
    texto = str(valor).strip()
    apenas_digitos = "".join(ch for ch in texto if ch.isdigit())
    if not apenas_digitos:
        return ""
    return apenas_digitos.zfill(11)


def normalizar_matricula(valor: Any) -> str:
    """
    Normaliza matrícula preservando formato texto estrito.
    Se for puramente numérica com menos de 8 dígitos, aplica .zfill(8)
    para garantir a integridade do padrão corporativo Cencosud.
    """
    if valor is None:
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    if texto.isdigit() and len(texto) < 8:
        return texto.zfill(8)
    return texto


def carregar_lista_pis(caminho_arquivo: Any) -> List[Dict[str, str]]:
    """
    Carrega a lista de colaboradores a partir de planilha Excel (.xlsx),
    garantindo tipagem de texto pura sem perda de zeros à esquerda e suporte
    a cabeçalhos com ou sem acentuação ("MATRÍCULA" vs "MATRICULA").
    """
    caminho_arquivo = Path(caminho_arquivo)
    if not caminho_arquivo.exists():
        raise FileNotFoundError(f"Arquivo de entrada não encontrado: {caminho_arquivo}")

    wb = openpyxl.load_workbook(caminho_arquivo, data_only=True)
    try:
        ws = wb.active
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            val = ws.cell(row=1, column=col_idx).value
            if val:
                chave_limpa = remover_acentos(str(val))
                headers[chave_limpa] = col_idx

        # Localizar índices de colunas com suporte a variações
        col_pis = headers.get("PIS") or headers.get("PIS/PASEP") or headers.get("PIS PASEP") or headers.get("NIS") or 3
        col_matr = headers.get("MATRICULA") or headers.get("MATRICULA DO EMPREGADO") or None
        col_nome = headers.get("NOME") or headers.get("NOME DO TRABALHADOR") or headers.get("NOME_COLABORADOR") or headers.get("NOME COLABORADOR") or 2
        col_base = headers.get("BASE_UF") or headers.get("BASE UF") or headers.get("BASE") or None

        registros = []
        for row_idx in range(2, ws.max_row + 1):
            pis_raw = ws.cell(row=row_idx, column=col_pis).value
            pis = normalizar_pis(pis_raw)
            if not pis:
                continue

            matricula_raw = ws.cell(row=row_idx, column=col_matr).value if col_matr else ""
            nome_raw = ws.cell(row=row_idx, column=col_nome).value if col_nome else ""
            base_uf_raw = ws.cell(row=row_idx, column=col_base).value if col_base else ""

            registros.append({
                "matricula": normalizar_matricula(matricula_raw),
                "nome": str(nome_raw or "").strip(),
                "pis": pis,
                "base_uf": str(base_uf_raw or BASE_CONTA_PADRAO).strip()
            })

        return registros
    finally:
        wb.close()


def exportar_relatorio_extratos(resultados: List[Dict[str, Any]], nome_arquivo: str = None, diretorio_saida: Optional[Path] = None) -> Path:
    """
    Exporta os resultados das consultas para planilha Excel seguindo rigorosamente
    o Padrão Visual Institucional Cencosud:
    - Colunas Cadastrais: Fundo Azul Corporativo (#1F497D) e fonte branca.
    - Colunas de Conferência/Valores: Fundo Amarelo (#FFFF00) e fonte vermelha negrito (#9C0006).
    - Células monetárias formatadas em R$ #,##0.00.
    - PIS e matrículas preservados como texto puro sem perda de zeros.
    - Colunas de auditoria de regras de negócio incluídas explicitamente.
    """
    if diretorio_saida is None:
        diretorio_saida = DIR_SAIDA
    else:
        diretorio_saida = Path(diretorio_saida)
    diretorio_saida.mkdir(parents=True, exist_ok=True)

    if nome_arquivo is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"Relatorio_Extratos_FGTS_Cencosud_{timestamp}.xlsx"

    caminho_saida = diretorio_saida / nome_arquivo

    wb = openpyxl.Workbook()
    try:
        ws = wb.active
        ws.title = "Extratos_FGTS"
        ws.views.sheetView[0].showGridLines = True

        # Estilos dos Blocos
        fill_cadastral = PatternFill(
            start_color=CORES_LAYOUT["cadastral_fundo"],
            end_color=CORES_LAYOUT["cadastral_fundo"],
            fill_type="solid"
        )
        font_cadastral = Font(name="Calibri", size=10, bold=True, color=CORES_LAYOUT["cadastral_fonte"])

        fill_conferencia = PatternFill(
            start_color=CORES_LAYOUT["conferencia_fundo"],
            end_color=CORES_LAYOUT["conferencia_fundo"],
            fill_type="solid"
        )
        font_conferencia = Font(name="Calibri", size=10, bold=True, color=CORES_LAYOUT["conferencia_fonte"])

        thin_border = Border(
            left=Side(style='thin', color=CORES_LAYOUT["borda_cinza"]),
            right=Side(style='thin', color=CORES_LAYOUT["borda_cinza"]),
            top=Side(style='thin', color=CORES_LAYOUT["borda_cinza"]),
            bottom=Side(style='thin', color=CORES_LAYOUT["borda_cinza"])
        )

        # Definição das Colunas do Relatório Oficial
        # Colunas 1 a 16: Bloco Cadastral (#1F497D / Branco)
        # Colunas 17 a 22: Bloco de Conferência e Auditoria (#FFFF00 / #9C0006)
        colunas_definicao = [
            # Bloco Cadastral (Azul Corporativo #1F497D / Branco)
            ("MATRÍCULA", "matricula", fill_cadastral, font_cadastral, "@", "center"),
            ("NOME INFORMADO", "nome_informado", fill_cadastral, font_cadastral, None, "left"),
            ("NOME CAIXA", "nome_caixa", fill_cadastral, font_cadastral, None, "left"),
            ("PIS PESQUISADO", "pis_pesquisado", fill_cadastral, font_cadastral, "@", "center"),
            ("BASE CONTA", "base_uf", fill_cadastral, font_cadastral, None, "center"),
            ("EMPRESA", "empresa", fill_cadastral, font_cadastral, None, "left"),
            ("CNPJ / CEI", "cnpj", fill_cadastral, font_cadastral, "@", "center"),
            ("CÓD. ESTABELECIMENTO", "cod_estab", fill_cadastral, font_cadastral, "@", "center"),
            ("CATEGORIA", "categoria", fill_cadastral, font_cadastral, "@", "center"),
            ("DESCRIÇÃO CATEGORIA", "categoria_descricao", fill_cadastral, font_cadastral, None, "left"),
            ("Nº CONTA FGTS", "conta_fgts", fill_cadastral, font_cadastral, "@", "center"),
            ("DATA ADMISSÃO", "data_admissao", fill_cadastral, font_cadastral, "DD/MM/YYYY", "center"),
            ("DATA OPÇÃO", "data_opcao", fill_cadastral, font_cadastral, "DD/MM/YYYY", "center"),
            ("TAXA JUROS", "taxa_juros", fill_cadastral, font_cadastral, None, "center"),
            ("TIPO CONTA", "tipo_conta", fill_cadastral, font_cadastral, None, "center"),
            ("DATA ATUALIZAÇÃO", "data_atualizacao", fill_cadastral, font_cadastral, "DD/MM/YYYY", "center"),

            # Bloco de Conferência / Financeiro / Auditoria (Amarelo #FFFF00 / Vermelho #9C0006)
            ("VALOR BASE FINS RESCISÓRIOS", "valor_base_rescisorios", fill_conferencia, font_conferencia, "R$ #,##0.00", "right"),
            ("SALDO FGTS", "saldo_fgts", fill_conferencia, font_conferencia, "R$ #,##0.00", "right"),
            ("REGRA / ALÍQUOTA FGTS", "regra_fgts", fill_conferencia, font_conferencia, None, "left"),
            ("MOVIMENTAÇÃO / RESCISÃO", "movimentacao_descricao", fill_conferencia, font_conferencia, None, "left"),
            ("DIVERGÊNCIAS AUDITORIA", "divergencias_auditoria", fill_conferencia, font_conferencia, None, "left"),
            ("STATUS CONSULTA", "status", fill_conferencia, font_conferencia, None, "center"),
            ("EXTRATO OFICIAL (PDF)", "caminho_pdf", fill_conferencia, font_conferencia, None, "center"),
            ("DATA/HORA DA CONSULTA", "timestamp_consulta", fill_conferencia, font_conferencia, "DD/MM/YYYY HH:MM:SS", "center")
        ]

        # Escrever Cabeçalhos
        ws.row_dimensions[1].height = 28
        for col_idx, (titulo, _, fill, font, _, _) in enumerate(colunas_definicao, 1):
            cell = ws.cell(row=1, column=col_idx, value=titulo)
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        # Escrever Linhas de Dados
        for row_idx, res in enumerate(resultados, 2):
            # Aplicar regras de auditoria e consistência cadastral
            auditar_situacao_colaborador(res)

            ws.row_dimensions[row_idx].height = 20
            # Efeito zebra suave apenas no bloco cadastral (col_idx <= 17)
            fill_linha = PatternFill(start_color=CORES_LAYOUT["zebra_par"], end_color=CORES_LAYOUT["zebra_par"], fill_type="solid") if row_idx % 2 == 0 else None

            for col_idx, (_, campo_chave, _, _, num_format, align_h) in enumerate(colunas_definicao, 1):
                valor = res.get(campo_chave, "")
                cell = ws.cell(row=row_idx, column=col_idx, value=valor)
                cell.border = thin_border

                # Aplicar zebrado exclusivamente no bloco cadastral (colunas 1 a 16)
                if fill_linha and col_idx <= 16:
                    cell.fill = fill_linha

                if num_format:
                    cell.number_format = num_format

                cell.alignment = Alignment(horizontal=align_h, vertical="center")

                # Destaque de status na célula
                if campo_chave == "status":
                    status_upper = str(valor).upper()
                    if "SUCESSO" in status_upper:
                        cell.font = Font(name="Calibri", size=10, bold=True, color=CORES_LAYOUT["sucesso_fonte"])
                    else:
                        cell.font = Font(name="Calibri", size=10, bold=True, color=CORES_LAYOUT["alerta_fonte"])

                # Link clicável para o arquivo PDF oficial do extrato
                if campo_chave == "caminho_pdf":
                    caminho_str = str(valor or "").strip()
                    if caminho_str and Path(caminho_str).exists():
                        cell.value = "Visualizar PDF Oficial"
                        cell.hyperlink = caminho_str
                        cell.font = Font(name="Calibri", size=10, bold=True, color="0000EE", underline="single")
                    else:
                        cell.value = "Não disponível"

        # Ajuste de Largura das Colunas
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

        wb.save(caminho_saida)
        return caminho_saida
    finally:
        wb.close()
