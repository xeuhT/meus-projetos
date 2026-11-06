"""
Script para Geração da Planilha Modelo de Consulta PIS
Conectividade Social V2 - Cencosud Brasil
Garante conformidade com o Padrão Visual Institucional (#1F497D para cadastral).
"""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path

from config import DIR_MODELOS, CORES_LAYOUT, BASE_CONTA_PADRAO


def criar_modelo_entrada(caminho_arquivo: Path = None) -> Path:
    """
    Gera a planilha modelo Excel formatada no padrão institucional Cencosud,
    com células tipadas estritamente como texto para PIS e Matrícula.
    Todas as colunas cadastrais utilizam Azul Corporativo (#1F497D) e fonte branca.
    """
    if caminho_arquivo is None:
        caminho_arquivo = DIR_MODELOS / "modelo_consulta_pis.xlsx"

    wb = openpyxl.Workbook()
    try:
        ws = wb.active
        ws.title = "Colaboradores_Consulta"
        ws.views.sheetView[0].showGridLines = True

        # Padrão Visual Institucional: Cabeçalho Cadastral em Azul Corporativo (#1F497D)
        fill_cadastral = PatternFill(
            start_color=CORES_LAYOUT["cadastral_fundo"],
            end_color=CORES_LAYOUT["cadastral_fundo"],
            fill_type="solid"
        )
        font_cadastral = Font(name="Calibri", size=11, bold=True, color=CORES_LAYOUT["cadastral_fonte"])

        thin_border = Border(
            left=Side(style='thin', color=CORES_LAYOUT["borda_cinza"]),
            right=Side(style='thin', color=CORES_LAYOUT["borda_cinza"]),
            top=Side(style='thin', color=CORES_LAYOUT["borda_cinza"]),
            bottom=Side(style='thin', color=CORES_LAYOUT["borda_cinza"])
        )

        # Definição das colunas com estilo institucional estrito
        cabecalhos = [
            ("MATRÍCULA", fill_cadastral, font_cadastral),
            ("NOME DO TRABALHADOR", fill_cadastral, font_cadastral),
            ("PIS", fill_cadastral, font_cadastral),
            ("BASE_UF", fill_cadastral, font_cadastral),
            ("OBSERVAÇÕES", fill_cadastral, font_cadastral)
        ]

        for col_idx, (titulo, fill, font) in enumerate(cabecalhos, 1):
            cell = ws.cell(row=1, column=col_idx, value=titulo)
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        
        ws.row_dimensions[1].height = 28

        # Dados de Exemplo (Piloto homologado)
        linhas_exemplo = [
            ("09941000", "MATHEUS CAVALCANTI SANTOS", "26984717122", BASE_CONTA_PADRAO, "Piloto Cencosud Homologado"),
            ("09941001", "COLABORADOR EXEMPLO 2", "12345678901", BASE_CONTA_PADRAO, "Preencher com PIS de 11 dígitos")
        ]

        for row_idx, dados in enumerate(linhas_exemplo, 2):
            ws.row_dimensions[row_idx].height = 20
            for col_idx, valor in enumerate(dados, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=valor)
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")
                
                # Formatação estrita como texto (@) para impedir perda de zeros à esquerda
                cell.number_format = "@"
                if col_idx in [1, 3]:  # Matrícula e PIS centralizados
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        # Ajuste automático da largura das colunas
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 18)

        wb.save(caminho_arquivo)
        print(f"[OK] Planilha modelo regenerada com sucesso em: {caminho_arquivo}")
        return caminho_arquivo
    finally:
        wb.close()


if __name__ == "__main__":
    criar_modelo_entrada()
