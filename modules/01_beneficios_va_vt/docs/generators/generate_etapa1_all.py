"""
Gerador Completo de Manuais Técnicos e Executivos Oficiais - Etapa 1
Gera simultaneamente:
1. Markdown Oficial (.md)
2. PDF Oficial Executivo (.pdf) via Microsoft Edge Headless
"""

import sys
import subprocess
from pathlib import Path

BASE_DOC_DIR = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\documentacao")
EDGE_EXE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

# ==============================================================================
# CSS CORPORATIVO OFICIAL CENCOSUD (IDÊNTICO AO AD004)
# ==============================================================================
CSS = """
@page {
    size: A4;
    margin: 10mm 12mm 10mm 12mm;
    @bottom-right {
        content: "Página " counter(page);
        font-family: Arial, Helvetica, sans-serif;
        font-size: 7pt;
        color: #7F7F7F;
    }
}
* { box-sizing: border-box; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
body { font-family: Arial, Helvetica, sans-serif; color: #262626; margin: 0; padding: 0; font-size: 8pt; line-height: 1.35; }
.header-container { border-bottom: 2pt solid #095AA5; padding-bottom: 3px; margin-bottom: 8px; }
.brand-title { font-size: 11pt; font-weight: 800; color: #095AA5; letter-spacing: 0.5px; text-transform: uppercase; margin: 0; }
.brand-subtitle { font-size: 7pt; font-weight: 600; color: #595959; margin: 1px 0 3px 0; }
.doc-main-title { font-size: 9.5pt; font-weight: 800; color: #1F497D; margin: 2px 0 3px 0; }
.badge-bar { background-color: #F2F5F9; border: 1px solid #DCE6F1; border-radius: 3px; padding: 2.5px 5px; font-size: 6.5pt; display: flex; justify-content: space-between; margin-top: 2px; color: #1F497D; font-weight: 600; }
.badge-bar span { margin-right: 8px; }
.badge-label { color: #595959; font-weight: normal; }
.kpi-grid { display: table; width: 100%; margin: 6px 0 8px 0; table-layout: fixed; }
.kpi-card { display: table-cell; background-color: #F8F9FA; border: 1px solid #D9D9D9; border-top: 2.5pt solid #095AA5; padding: 4px 5px; text-align: center; width: 25%; vertical-align: middle; }
.kpi-card:not(:last-child) { border-right: none; }
.kpi-title { font-size: 6pt; font-weight: 700; color: #595959; text-transform: uppercase; margin-bottom: 1px; }
.kpi-value { font-size: 10pt; font-weight: 800; color: #095AA5; line-height: 1.1; }
.kpi-sub { font-size: 5.5pt; color: #7F7F7F; margin-top: 1px; }
h2.section-title { font-size: 8.5pt; font-weight: 800; color: #095AA5; text-transform: uppercase; border-left: 2.5pt solid #095AA5; padding-left: 4px; margin: 8px 0 4px 0; page-break-after: avoid; }
h3.subsection-title { font-size: 7.5pt; font-weight: 700; color: #1F497D; margin: 6px 0 3px 0; page-break-after: avoid; }
p { margin: 2px 0 4px 0; text-align: justify; }
ul { margin: 2px 0 4px 0; padding-left: 14px; }
li { margin-bottom: 1.5px; }
table.report-table { width: 100%; border-collapse: collapse; margin: 4px 0 6px 0; font-size: 7pt; page-break-inside: avoid; }
table.report-table th { background-color: #1F497D; color: #FFFFFF; font-weight: 700; padding: 3px 4px; border: 1px solid #1F497D; text-align: left; font-size: 6.5pt; }
table.report-table th.th-audit { background-color: #FFFF00 !important; color: #9C0006 !important; border: 1px solid #D9D9D9; font-weight: 800; }
table.report-table td { padding: 2.5px 4px; border: 1px solid #D9D9D9; vertical-align: middle; }
table.report-table tr:nth-child(even) { background-color: #F8F9FA; }
.badge-rn { background-color: #095AA5; color: #FFFFFF; font-weight: 800; padding: 1px 3px; border-radius: 2px; font-size: 6pt; display: inline-block; }
.code-box { background-color: #1E222D; color: #D4D4D4; font-family: "Consolas", monospace; font-size: 6.5pt; padding: 5px 6px; border-radius: 3px; margin: 4px 0; line-height: 1.25; border-left: 2pt solid #095AA5; page-break-inside: avoid; white-space: pre-wrap; }
.callout { background-color: #F2F5F9; border-left: 2.5pt solid #095AA5; padding: 4px 6px; margin: 4px 0; font-size: 7pt; page-break-inside: avoid; }
.callout-alert { background-color: #FFF8E6; border-left: 2.5pt solid #FFB800; color: #8A5B00; }
.footer-signatures { margin-top: 14px; display: table; width: 100%; page-break-inside: avoid; }
.sig-box { display: table-cell; width: 50%; text-align: center; padding: 6px 12px; }
.sig-line { border-top: 1px solid #595959; margin: 16px 20px 4px 20px; }
.sig-title { font-size: 7.5pt; font-weight: 700; color: #1F497D; }
.sig-role { font-size: 6.5pt; color: #595959; }
.page-break { page-break-before: always; }
"""

def compile_pdf(html_content: str, out_html: Path, out_pdf: Path):
    out_html.write_text(html_content, encoding="utf-8")
    cmd = [
        EDGE_EXE,
        "--headless=new",
        "--disable-gpu",
        "--no-pdf-header-footer",
        f"--print-to-pdf={out_pdf}",
        str(out_html)
    ]
    subprocess.run(cmd, capture_output=True, text=True)
    if out_pdf.exists():
        print(f"  [OK] PDF: {out_pdf.name} ({out_pdf.stat().st_size:,} bytes)")
    else:
        print(f"  [ERRO] Falha ao gerar {out_pdf.name}")

print("Carregado gerador oficial.")
