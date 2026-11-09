"""
Gerador de Manuais Técnicos e Executivos Oficiais - Cencosud Brasil
Central de Benefícios (VA, VT, Cartões e Refeição)
"""

import os
import subprocess
from pathlib import Path

BASE_DOC_DIR = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\documentacao")
EDGE_EXE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

CSS_STYLE = """
@page {
    size: A4;
    margin: 14mm 15mm 14mm 15mm;
    @bottom-right {
        content: "Página " counter(page) " de " counter(pages);
        font-family: Arial, Helvetica, sans-serif;
        font-size: 8pt;
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
    font-size: 9.5pt;
    line-height: 1.45;
}

.header-container {
    border-bottom: 2.5pt solid #095AA5;
    padding-bottom: 6px;
    margin-bottom: 12px;
}

.brand-title {
    font-size: 13pt;
    font-weight: 800;
    color: #095AA5;
    letter-spacing: 0.5px;
    text-transform: uppercase;
    margin: 0;
}

.brand-subtitle {
    font-size: 8pt;
    font-weight: 600;
    color: #595959;
    margin: 2px 0 6px 0;
}

.doc-main-title {
    font-size: 11pt;
    font-weight: 800;
    color: #1F497D;
    margin: 4px 0;
}

.badge-bar {
    background-color: #F2F5F9;
    border: 1px solid #DCE6F1;
    border-radius: 4px;
    padding: 4px 8px;
    font-size: 7.5pt;
    display: flex;
    justify-content: space-between;
    margin-top: 5px;
    color: #1F497D;
    font-weight: 600;
}

.badge-item {
    display: inline-block;
}

.badge-label {
    color: #595959;
    font-weight: normal;
}

.kpi-grid {
    display: table;
    width: 100%;
    margin: 12px 0;
    table-layout: fixed;
}

.kpi-card {
    display: table-cell;
    background-color: #F8F9FA;
    border: 1px solid #D9D9D9;
    border-top: 3pt solid #095AA5;
    padding: 6px 8px;
    text-align: center;
    width: 25%;
    vertical-align: middle;
}

.kpi-card:not(:last-child) {
    margin-right: 8px;
}

.kpi-title {
    font-size: 7pt;
    font-weight: 700;
    color: #595959;
    text-transform: uppercase;
    margin-bottom: 2px;
}

.kpi-value {
    font-size: 12pt;
    font-weight: 800;
    color: #095AA5;
    line-height: 1.1;
}

.kpi-sub {
    font-size: 6.5pt;
    color: #7F7F7F;
    margin-top: 2px;
}

h2.section-title {
    font-size: 10.5pt;
    font-weight: 800;
    color: #095AA5;
    text-transform: uppercase;
    border-left: 3.5pt solid #095AA5;
    padding-left: 6px;
    margin: 14px 0 8px 0;
    page-break-after: avoid;
}

h3.subsection-title {
    font-size: 9pt;
    font-weight: 700;
    color: #1F497D;
    margin: 10px 0 4px 0;
    page-break-after: avoid;
}

p {
    margin: 4px 0 6px 0;
    text-align: justify;
}

ul {
    margin: 4px 0 6px 0;
    padding-left: 18px;
}

li {
    margin-bottom: 3px;
}

table.report-table {
    width: 100%;
    border-collapse: collapse;
    margin: 8px 0;
    font-size: 8pt;
    page-break-inside: avoid;
}

table.report-table th {
    background-color: #1F497D;
    color: #FFFFFF;
    font-weight: 700;
    padding: 5px 6px;
    border: 1px solid #1F497D;
    text-align: left;
    font-size: 7.5pt;
}

table.report-table th.th-audit {
    background-color: #FFFF00 !important;
    color: #9C0006 !important;
    border: 1px solid #E6D800;
    font-weight: 800;
}

table.report-table td {
    padding: 4px 6px;
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
    padding: 2px 4px;
    border-radius: 3px;
    font-size: 7pt;
    display: inline-block;
}

.code-box {
    background-color: #1E222D;
    color: #D4D4D4;
    font-family: "Consolas", "Courier New", monospace;
    font-size: 7.5pt;
    padding: 8px 10px;
    border-radius: 4px;
    margin: 8px 0;
    line-height: 1.35;
    border-left: 3pt solid #095AA5;
    page-break-inside: avoid;
}

.code-box .keyword { color: #569CD6; font-weight: bold; }
.code-box .string { color: #CE9178; }
.code-box .comment { color: #6A9955; font-style: italic; }
.code-box .number { color: #B5CEA8; }
.code-box .func { color: #DCDCAA; }

.callout {
    background-color: #F2F5F9;
    border-left: 3pt solid #095AA5;
    padding: 6px 10px;
    margin: 8px 0;
    font-size: 8pt;
    page-break-inside: avoid;
}

.callout-alert {
    background-color: #FFF8E6;
    border-left: 3pt solid #FFB800;
    color: #8A5B00;
}

.callout-success {
    background-color: #EBF8F2;
    border-left: 3pt solid #00A859;
    color: #006636;
}

.footer-signatures {
    margin-top: 25px;
    display: table;
    width: 100%;
    page-break-inside: avoid;
}

.sig-box {
    display: table-cell;
    width: 50%;
    text-align: center;
    padding: 10px 20px;
}

.sig-line {
    border-top: 1px solid #595959;
    margin: 25px 30px 6px 30px;
}

.sig-title {
    font-size: 8.5pt;
    font-weight: 700;
    color: #1F497D;
}

.sig-role {
    font-size: 7.5pt;
    color: #595959;
}

.page-break {
    page-break-before: always;
}
"""

def generate_pdf_from_html(html_path: Path, pdf_path: Path):
    """Executa Microsoft Edge Headless para converter HTML em PDF."""
    cmd = [
        EDGE_EXE,
        "--headless=new",
        "--disable-gpu",
        "--no-pdf-header-footer",
        f"--print-to-pdf={pdf_path}",
        str(html_path)
    ]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if not pdf_path.exists():
        raise RuntimeError(f"Falha ao gerar PDF {pdf_path}: {res.stderr}")
    print(f"[PDF GERADO] {pdf_path.name} ({pdf_path.stat().st_size:,} bytes)")

print("Módulo gerador base pronto.")
