"""
Script de Geração Oficial dos Manuais Técnicos e Executivos - Etapa 1
Frentes Homologadas e Ativas:
1. Good Card (Evt 6280)
2. Compra Mensal VT
3. Valores Pré-Folha VT
4. Bloqueio Good Card
"""

import os
import subprocess
from pathlib import Path

BASE_DOC_DIR = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\documentacao")
EDGE_EXE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

# ==============================================================================
# TEMPLATE CSS PROFISSIONAL CENCOSUD (IDÊNTICO AO AD004)
# ==============================================================================
CSS_TEMPLATE = """
@page {
    size: A4;
    margin: 12mm 14mm 12mm 14mm;
    @bottom-right {
        content: "Página " counter(page);
        font-family: Arial, Helvetica, sans-serif;
        font-size: 7.5pt;
        color: #7F7F7F;
    }
}

* {
    box-sizing: border-box;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
}

body {
    font-family: Arial, Helvetica, sans-serif;
    color: #262626;
    margin: 0;
    padding: 0;
    font-size: 8.5pt;
    line-height: 1.4;
}

.header-container {
    border-bottom: 2pt solid #095AA5;
    padding-bottom: 4px;
    margin-bottom: 10px;
}

.brand-title {
    font-size: 12pt;
    font-weight: 800;
    color: #095AA5;
    letter-spacing: 0.5px;
    text-transform: uppercase;
    margin: 0;
}

.brand-subtitle {
    font-size: 7.5pt;
    font-weight: 600;
    color: #595959;
    margin: 1px 0 4px 0;
}

.doc-main-title {
    font-size: 10pt;
    font-weight: 800;
    color: #1F497D;
    margin: 2px 0 4px 0;
}

.badge-bar {
    background-color: #F2F5F9;
    border: 1px solid #DCE6F1;
    border-radius: 3px;
    padding: 3px 6px;
    font-size: 7pt;
    display: flex;
    justify-content: space-between;
    margin-top: 3px;
    color: #1F497D;
    font-weight: 600;
}

.badge-bar span {
    margin-right: 12px;
}

.badge-label {
    color: #595959;
    font-weight: normal;
}

.kpi-grid {
    display: table;
    width: 100%;
    margin: 8px 0 10px 0;
    table-layout: fixed;
}

.kpi-card {
    display: table-cell;
    background-color: #F8F9FA;
    border: 1px solid #D9D9D9;
    border-top: 2.5pt solid #095AA5;
    padding: 5px 6px;
    text-align: center;
    width: 25%;
    vertical-align: middle;
}

.kpi-card:not(:last-child) {
    border-right: none;
}

.kpi-title {
    font-size: 6.5pt;
    font-weight: 700;
    color: #595959;
    text-transform: uppercase;
    margin-bottom: 2px;
}

.kpi-value {
    font-size: 11pt;
    font-weight: 800;
    color: #095AA5;
    line-height: 1.1;
}

.kpi-sub {
    font-size: 6pt;
    color: #7F7F7F;
    margin-top: 2px;
}

h2.section-title {
    font-size: 9.5pt;
    font-weight: 800;
    color: #095AA5;
    text-transform: uppercase;
    border-left: 3pt solid #095AA5;
    padding-left: 5px;
    margin: 12px 0 6px 0;
    page-break-after: avoid;
}

h3.subsection-title {
    font-size: 8.5pt;
    font-weight: 700;
    color: #1F497D;
    margin: 8px 0 4px 0;
    page-break-after: avoid;
}

p {
    margin: 3px 0 5px 0;
    text-align: justify;
}

ul {
    margin: 3px 0 5px 0;
    padding-left: 16px;
}

li {
    margin-bottom: 2px;
}

table.report-table {
    width: 100%;
    border-collapse: collapse;
    margin: 6px 0 8px 0;
    font-size: 7.5pt;
    page-break-inside: avoid;
}

table.report-table th {
    background-color: #1F497D;
    color: #FFFFFF;
    font-weight: 700;
    padding: 4px 5px;
    border: 1px solid #1F497D;
    text-align: left;
    font-size: 7pt;
}

table.report-table th.th-audit {
    background-color: #FFFF00 !important;
    color: #9C0006 !important;
    border: 1px solid #D9D9D9;
    font-weight: 800;
}

table.report-table td {
    padding: 3.5px 5px;
    border: 1px solid #D9D9D9;
    vertical-align: middle;
}

table.report-table tr:nth-child(even) {
    background-color: #F8F9FA;
}

.badge-rn {
    background-color: #095AA5;
    color: #FFFFFF;
    font-weight: 800;
    padding: 1.5px 3.5px;
    border-radius: 2px;
    font-size: 6.5pt;
    display: inline-block;
}

.code-box {
    background-color: #1E222D;
    color: #D4D4D4;
    font-family: "Consolas", "Courier New", monospace;
    font-size: 7pt;
    padding: 6px 8px;
    border-radius: 3px;
    margin: 6px 0;
    line-height: 1.3;
    border-left: 2.5pt solid #095AA5;
    page-break-inside: avoid;
    white-space: pre-wrap;
}

.code-box .kw { color: #569CD6; font-weight: bold; }
.code-box .str { color: #CE9178; }
.code-box .com { color: #6A9955; font-style: italic; }
.code-box .num { color: #B5CEA8; }
.code-box .fn { color: #DCDCAA; }

.callout {
    background-color: #F2F5F9;
    border-left: 2.5pt solid #095AA5;
    padding: 5px 8px;
    margin: 6px 0;
    font-size: 7.5pt;
    page-break-inside: avoid;
}

.callout-alert {
    background-color: #FFF8E6;
    border-left: 2.5pt solid #FFB800;
    color: #8A5B00;
}

.footer-signatures {
    margin-top: 18px;
    display: table;
    width: 100%;
    page-break-inside: avoid;
}

.sig-box {
    display: table-cell;
    width: 50%;
    text-align: center;
    padding: 8px 16px;
}

.sig-line {
    border-top: 1px solid #595959;
    margin: 20px 24px 5px 24px;
}

.sig-title {
    font-size: 8pt;
    font-weight: 700;
    color: #1F497D;
}

.sig-role {
    font-size: 7pt;
    color: #595959;
}

.page-break {
    page-break-before: always;
}

.footer-bar {
    border-top: 1px solid #D9D9D9;
    padding-top: 3px;
    margin-top: 15px;
    font-size: 6.5pt;
    color: #7F7F7F;
    display: flex;
    justify-content: space-between;
}
"""

def compile_pdf(html_content: str, output_html: Path, output_pdf: Path):
    """Grava o HTML e compila para PDF via Edge headless."""
    output_html.write_text(html_content, encoding="utf-8")
    cmd = [
        EDGE_EXE,
        "--headless=new",
        "--disable-gpu",
        "--no-pdf-header-footer",
        f"--print-to-pdf={output_pdf}",
        str(output_html)
    ]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if not output_pdf.exists():
        raise RuntimeError(f"Erro ao compilar {output_pdf.name}: {res.stderr}")
    print(f"  [OK] PDF gerado: {output_pdf.name} ({output_pdf.stat().st_size:,} bytes)")

print("Ambiente de compilação pronto.")
