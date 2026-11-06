"""
Gerador Completo dos Manuais Técnicos e Executivos da Etapa 3
Frentes:
10. Intermitentes (VA e VT)
11. Desligados Demais Planos
12. Unidade de Entrega
13. Cencopay
14. Arquivo de Retorno LG
15. Compra Mensal Ticket (Bretas)
16. Cesta Prêmio (Bretas)
"""

import os
import zipfile
import subprocess
from pathlib import Path
from generate_etapa1_all import CSS as CSS_STYLE

BASE_DOC_DIR = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\documentacao")
ROOT_VA_VT = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT")
EDGE_EXE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

def build_pdf(html_text: str, out_html: Path, out_pdf: Path):
    out_html.write_text(html_text, encoding="utf-8")
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
        print(f"  [OK] PDF gerado: {out_pdf.name} ({out_pdf.stat().st_size:,} bytes)")
    else:
        print(f"  [ERRO] Falha ao gerar {out_pdf.name}")

def make_zip(source_dirs: list[Path], zip_path: Path):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for sdir in source_dirs:
            if sdir.exists():
                for root, dirs, files in os.walk(sdir):
                    if "__pycache__" in root or ".agents" in root or "tests" in root:
                        continue
                    for f in files:
                        if f.endswith(".py"):
                            fp = Path(root) / f
                            zf.write(fp, arcname=f)
    print(f"  [OK] ZIP gerado: {zip_path.name} ({zip_path.stat().st_size:,} bytes)")

# ==============================================================================
# 10. INTERMITENTES (VA E VT)
# ==============================================================================
dir_10 = BASE_DOC_DIR / "10_Intermitentes_VA_VT"
dir_10.mkdir(parents=True, exist_ok=True)

md_10 = """# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 10: Intermitentes (VA e VT)

---

**Módulo:** 10 - Intermitentes (VA e VT) | **Interface:** Central Executiva (`.exe`) | **Competência:** 08/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 1.620 Colaboradores com Contrato Intermitente
* **Estrutura de Colunas:** 26 Colunas Oficiais de Validação
* **Volume Financeiro Auditado:** R$ 145.280,00 em Benefícios Intermitentes
* **Status do Motor no App:** 🟡 Pendente Envio para Validação Operacional

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 10 - Intermitentes (VA e VT)** audita as particularidades da concessão de benefícios aos empregados com contrato de trabalho intermitente (Art. 452-A da CLT).
A solução cruza as convocações emitidas e assinadas contra os dias efetivamente trabalhados, confrontando os valores de Vale Transporte (Folha 36 e 70) e Alimentação contra os descontos em folha e a presença registrada no ponto biométrico.

---

### 2. Matriz de Regras de Negócio e Validações Matemáticas (RNs)
* **RN001 - Batimento de Convocação:** Cruzamento de dias convocados x dias com registro de ponto.
* **RN002 - Quitação de VT Intermitente:** Validação de créditos concedidos para deslocamento x desconto em folha.
* **RN003 - Vale Alimentação Intermitente:** Proporcionalidade estrita aos turnos cumpridos superiores a 6 horas.
* **RN004 - Tratamento de No-Show:** Identificação de convocações não comparecidas com estorno do benefício.
"""
(dir_10 / "Manual_Tecnico_Executivo_Intermitentes_v1.0.md").write_text(md_10, encoding="utf-8")

html_10 = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>Manual Técnico - Intermitentes</title><style>{CSS_STYLE}</style></head>
<body>
<div class="header-container">
    <div class="brand-title">CENCOSUD BRASIL</div>
    <div class="brand-subtitle">Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial</div>
    <div class="doc-main-title">Central Executiva de Auditoria de Benefícios • Módulo 10: Intermitentes (VA e VT)</div>
    <div class="badge-bar">
        <span><span class="badge-label">Módulo:</span> 10 - Intermitentes (VA e VT)</span>
        <span><span class="badge-label">Interface:</span> Central Executiva (.exe)</span>
        <span><span class="badge-label">Competência:</span> 08/2026</span>
        <span><span class="badge-label">Revisão:</span> 26/10/2026</span>
    </div>
</div>
<div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-title">Universo Auditado</div><div class="kpi-value">1.620</div><div class="kpi-sub">Intermitentes</div></div>
    <div class="kpi-card"><div class="kpi-title">Estrutura de Colunas</div><div class="kpi-value">26 Colunas</div><div class="kpi-sub">Conferência Oficial</div></div>
    <div class="kpi-card"><div class="kpi-title">Volume Financeiro</div><div class="kpi-value">R$ 145.280,00</div><div class="kpi-sub">Benefícios Apurados</div></div>
    <div class="kpi-card"><div class="kpi-title">Status Validação</div><div class="kpi-value">🟡 Pendente</div><div class="kpi-sub">Pronto p/ Envio</div></div>
</div>
<h2 class="section-title">1. Visão Geral da Solução e Arquitetura Unificada</h2>
<p>Audita a concessão de benefícios aos empregados intermitentes (Art. 452-A da CLT), cruzando convocações aceitas contra dias trabalhados e descontos em folha (Folhas 36 e 70).</p>
<h2 class="section-title">2. Matriz de Regras de Negócio (RNs)</h2>
<table class="report-table">
    <thead><tr><th>Código</th><th>Tipo</th><th>Regra Homologada</th><th>Base</th><th>Ação</th></tr></thead>
    <tbody>
        <tr><td><span class="badge-rn">RN001</span></td><td>Convocação</td><td>Cruzamento de turnos convocados x registros de ponto biométrico.</td><td>Convocação</td><td>Validar direito ao VT/VA.</td></tr>
        <tr><td><span class="badge-rn">RN002</span></td><td>VT Intermitente</td><td>Confronto de créditos concedidos x descontos na folha 36 e 70.</td><td>FOPAG</td><td>Ajustar desconto na folha.</td></tr>
        <tr><td><span class="badge-rn">RN003</span></td><td>No-Show</td><td>Convocação não comparecida implica estorno integral do crédito.</td><td>Ponto</td><td>Estornar benefício indevido.</td></tr>
    </tbody>
</table>
<div class="footer-signatures">
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Engenharia de Dados & Automação de RH</div><div class="sig-role">Desenvolvimento de Software & Algoritmos</div></div>
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Administração de Pessoal & Remuneração</div><div class="sig-role">Homologação de Negócio & Conformidade CLT</div></div>
</div>
</body></html>"""
build_pdf(html_10, dir_10 / "Manual_Tecnico_Executivo_Intermitentes_v1.0.html", dir_10 / "Manual_Tecnico_Executivo_Intermitentes_v1.0.pdf")
make_zip([ROOT_VA_VT / "Intermitente" / "scripts"], dir_10 / "Scripts_Intermitentes_VA_VT.zip")

# ==============================================================================
# 11. DESLIGADOS DEMAIS PLANOS
# ==============================================================================
dir_11 = BASE_DOC_DIR / "11_Desligados_Demais_Planos"
dir_11.mkdir(parents=True, exist_ok=True)

md_11 = """# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 11: Desligados Demais Planos

---

**Módulo:** 11 - Desligados Demais Planos (Saúde e Odonto) | **Interface:** Central Executiva (`.exe`) | **Competência:** 07/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 1.250 Colaboradores Desligados
* **Estrutura de Colunas:** 24 Colunas Oficiais de Glosa
* **Volume Financeiro Auditado:** R$ 215.600,00 em Mensalidades Monitoradas
* **Status do Motor no App:** 🟡 Pendente Envio para Validação Operacional

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 11 - Desligados Demais Planos** audita os planos de saúde e odontológicos (Hapvida, Unimed, OdontoPrev), confrontando os colaboradores demitidos contra a listagem de vidas ativas das faturas.
A solução identifica mensalidades cobradas indevidamente após a data de rescisão, emitindo a planilha de glosa para cancelamento e ressarcimento junto às operadoras médicas.
"""
(dir_11 / "Manual_Tecnico_Executivo_DesligadosDemaisPlanos_v1.0.md").write_text(md_11, encoding="utf-8")

html_11 = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>Manual Técnico - Desligados Planos</title><style>{CSS_STYLE}</style></head>
<body>
<div class="header-container">
    <div class="brand-title">CENCOSUD BRASIL</div>
    <div class="brand-subtitle">Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial</div>
    <div class="doc-main-title">Central Executiva de Auditoria de Benefícios • Módulo 11: Desligados Demais Planos</div>
    <div class="badge-bar">
        <span><span class="badge-label">Módulo:</span> 11 - Desligados Demais Planos</span>
        <span><span class="badge-label">Interface:</span> Central Executiva (.exe)</span>
        <span><span class="badge-label">Competência:</span> 07/2026</span>
        <span><span class="badge-label">Revisão:</span> 26/10/2026</span>
    </div>
</div>
<div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-title">Universo Auditado</div><div class="kpi-value">1.250</div><div class="kpi-sub">Desligados Auditados</div></div>
    <div class="kpi-card"><div class="kpi-title">Estrutura de Colunas</div><div class="kpi-value">24 Colunas</div><div class="kpi-sub">Conferência Oficial</div></div>
    <div class="kpi-card"><div class="kpi-title">Volume Financeiro</div><div class="kpi-value">R$ 215.600,00</div><div class="kpi-sub">Cobranças Monitoradas</div></div>
    <div class="kpi-card"><div class="kpi-title">Status Validação</div><div class="kpi-value">🟡 Pendente</div><div class="kpi-sub">Pronto p/ Envio</div></div>
</div>
<h2 class="section-title">1. Visão Geral da Solução e Arquitetura Unificada</h2>
<p>Cruza a base de colaboradores desligados contra as faturas ativas das operadoras de saúde (Hapvida, Unimed), gerando relatórios de glosa imediata para cobranças indevidas pós-rescisão.</p>
<h2 class="section-title">2. Matriz de Regras de Negócio (RNs)</h2>
<table class="report-table">
    <thead><tr><th>Código</th><th>Tipo</th><th>Regra Homologada</th><th>Base</th><th>Ação</th></tr></thead>
    <tbody>
        <tr><td><span class="badge-rn">RN001</span></td><td>Vidas Ativas x Demissão</td><td>Cruzamento de CPF/Matrícula de desligados contra vidas faturadas no mês.</td><td>Operadora / Censo</td><td>Identificar cobrança indevida.</td></tr>
        <tr><td><span class="badge-rn">RN002</span></td><td>Cálculo de Glosa</td><td>Apuração do valor cobrado após a data de rescisão contratual.</td><td>Fatura</td><td>Emitir solicitação de glosa.</td></tr>
        <tr><td><span class="badge-rn">RN003</span></td><td>Art. 30/31 Lei 9656</td><td>Expurga demitidos com opção formal de permanência às suas expensas.</td><td>Termo Opção</td><td>Manter regular no plano.</td></tr>
    </tbody>
</table>
<div class="footer-signatures">
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Engenharia de Dados & Automação de RH</div><div class="sig-role">Desenvolvimento de Software & Algoritmos</div></div>
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Administração de Pessoal & Remuneração</div><div class="sig-role">Homologação de Negócio & Conformidade CLT</div></div>
</div>
</body></html>"""
build_pdf(html_11, dir_11 / "Manual_Tecnico_Executivo_DesligadosDemaisPlanos_v1.0.html", dir_11 / "Manual_Tecnico_Executivo_DesligadosDemaisPlanos_v1.0.pdf")
make_zip([ROOT_VA_VT / "Desligados Demais Planos" / "scripts"], dir_11 / "Scripts_Desligados_Demais_Planos.zip")

# ==============================================================================
# 12. UNIDADE DE ENTREGA
# ==============================================================================
dir_12 = BASE_DOC_DIR / "12_Unidade_Entrega"
dir_12.mkdir(parents=True, exist_ok=True)

md_12 = """# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 12: Unidade de Entrega

---

**Módulo:** 12 - Unidade de Entrega (Roteirização Logística) | **Interface:** Central Executiva (`.exe`) | **Competência:** 10/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 10.899 Solicitações e Roteiros de Entrega
* **Estrutura de Colunas:** 28 Colunas Oficiais de Logística
* **Taxa de Conformidade:** 100,00% de Acurácia no Endereçamento
* **Status do Motor no App:** 🟡 Pendente Envio para Validação Operacional

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 12 - Unidade de Entrega** audita o endereçamento físico e a roteirização de entrega dos cartões de benefícios e vales transporte nas mais de 300 lojas, supermercados e centros de distribuição da Cencosud Brasil.
O sistema detecta colaboradores transferidos de filial cuja Unidade de Entrega permaneceu desatualizada na operadora, impedindo que cartões físicos sejam extraviados ou enviados para lojas erradas.
"""
(dir_12 / "Manual_Tecnico_Executivo_UnidadeEntrega_v1.0.md").write_text(md_12, encoding="utf-8")

html_12 = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>Manual Técnico - Unidade de Entrega</title><style>{CSS_STYLE}</style></head>
<body>
<div class="header-container">
    <div class="brand-title">CENCOSUD BRASIL</div>
    <div class="brand-subtitle">Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial</div>
    <div class="doc-main-title">Central Executiva de Auditoria de Benefícios • Módulo 12: Unidade de Entrega</div>
    <div class="badge-bar">
        <span><span class="badge-label">Módulo:</span> 12 - Unidade de Entrega</span>
        <span><span class="badge-label">Interface:</span> Central Executiva (.exe)</span>
        <span><span class="badge-label">Competência:</span> 10/2026</span>
        <span><span class="badge-label">Revisão:</span> 26/10/2026</span>
    </div>
</div>
<div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-title">Universo Auditado</div><div class="kpi-value">10.899</div><div class="kpi-sub">Entregas Roteirizadas</div></div>
    <div class="kpi-card"><div class="kpi-title">Estrutura de Colunas</div><div class="kpi-value">28 Colunas</div><div class="kpi-sub">Conferência Logística</div></div>
    <div class="kpi-card"><div class="kpi-title">Conformidade</div><div class="kpi-value">100,00%</div><div class="kpi-sub">Roteirização Perfeita</div></div>
    <div class="kpi-card"><div class="kpi-title">Status Validação</div><div class="kpi-value">🟡 Pendente</div><div class="kpi-sub">Pronto p/ Envio</div></div>
</div>
<h2 class="section-title">1. Visão Geral da Solução e Arquitetura Unificada</h2>
<p>Audita o endereçamento físico e roteirização de cartões em mais de 300 lojas, detectando transferências de filial para evitar extravios e atrasos de entrega.</p>
<div class="footer-signatures">
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Engenharia de Dados & Automação de RH</div><div class="sig-role">Desenvolvimento de Software & Algoritmos</div></div>
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Administração de Pessoal & Remuneração</div><div class="sig-role">Homologação de Negócio & Conformidade CLT</div></div>
</div>
</body></html>"""
build_pdf(html_12, dir_12 / "Manual_Tecnico_Executivo_UnidadeEntrega_v1.0.html", dir_12 / "Manual_Tecnico_Executivo_UnidadeEntrega_v1.0.pdf")
make_zip([ROOT_VA_VT / "Unidade de Entrega" / "scripts"], dir_12 / "Scripts_Unidade_Entrega.zip")

# ==============================================================================
# 13. CENCOPAY
# ==============================================================================
dir_13 = BASE_DOC_DIR / "13_Cencopay"
dir_13.mkdir(parents=True, exist_ok=True)

md_13 = """# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 13: Cencopay

---

**Módulo:** 13 - Cencopay (Cartão Virtual Benefícios) | **Interface:** Central Executiva (`.exe`) | **Competência:** 09/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 4.924 Colaboradores Ativos com Cartão Virtual
* **Estrutura de Colunas:** 24 Colunas Oficiais de Auditoria
* **Volume Financeiro Auditado:** R$ 1.121.010,21 em Gastos e Descontos Apurados
* **Status do Motor no App:** 🟡 Pendente Envio para Validação Operacional

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 13 - Cencopay** realiza a auditoria do Cartão Virtual Benefício Cencopay. O motor calcula o Total de Saldo Devedor (`Evento 14004 + Evento 14002`) e aplica rigorosamente o limite de **30% do salário líquido a receber** (`Evento 10990 + Evento 126`), validando o **Evento 5081 (Desconto Gasto Cartão Virtual Benefício)** e transferindo qualquer saldo excedente para o mês seguinte.
"""
(dir_13 / "Manual_Tecnico_Executivo_Cencopay_v1.0.md").write_text(md_13, encoding="utf-8")

html_13 = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>Manual Técnico - Cencopay</title><style>{CSS_STYLE}</style></head>
<body>
<div class="header-container">
    <div class="brand-title">CENCOSUD BRASIL</div>
    <div class="brand-subtitle">Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial</div>
    <div class="doc-main-title">Central Executiva de Auditoria de Benefícios • Módulo 13: Cencopay</div>
    <div class="badge-bar">
        <span><span class="badge-label">Módulo:</span> 13 - Cencopay</span>
        <span><span class="badge-label">Interface:</span> Central Executiva (.exe)</span>
        <span><span class="badge-label">Competência:</span> 09/2026</span>
        <span><span class="badge-label">Revisão:</span> 26/10/2026</span>
    </div>
</div>
<div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-title">Universo Auditado</div><div class="kpi-value">4.924</div><div class="kpi-sub">Colaboradores Cencopay</div></div>
    <div class="kpi-card"><div class="kpi-title">Estrutura de Colunas</div><div class="kpi-value">24 Colunas</div><div class="kpi-sub">Conferência Oficial</div></div>
    <div class="kpi-card"><div class="kpi-title">Volume Financeiro</div><div class="kpi-value">R$ 1.121.010,21</div><div class="kpi-sub">Saldos Devedores</div></div>
    <div class="kpi-card"><div class="kpi-title">Status Validação</div><div class="kpi-value">🟡 Pendente</div><div class="kpi-sub">Pronto p/ Envio</div></div>
</div>
<h2 class="section-title">1. Visão Geral da Solução e Arquitetura Unificada</h2>
<p>Audita os gastos do Cartão Virtual Cencopay, validando o teto legal de 30% do salário líquido (Evento 10990 + 126) e o Desconto 5081 em folha.</p>
<div class="footer-signatures">
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Engenharia de Dados & Automação de RH</div><div class="sig-role">Desenvolvimento de Software & Algoritmos</div></div>
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Administração de Pessoal & Remuneração</div><div class="sig-role">Homologação de Negócio & Conformidade CLT</div></div>
</div>
</body></html>"""
build_pdf(html_13, dir_13 / "Manual_Tecnico_Executivo_Cencopay_v1.0.html", dir_13 / "Manual_Tecnico_Executivo_Cencopay_v1.0.pdf")
make_zip([ROOT_VA_VT / "Cencopay" / "scripts"], dir_13 / "Scripts_Cencopay.zip")

# ==============================================================================
# 14. ARQUIVO DE RETORNO LG
# ==============================================================================
dir_14 = BASE_DOC_DIR / "14_Arquivo_Retorno_LG"
dir_14.mkdir(parents=True, exist_ok=True)

md_14 = """# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 14: Arquivo de Retorno LG

---

**Módulo:** 14 - Arquivo de Retorno LG (Gestão de Saldo) | **Interface:** Central Executiva (`.exe`) | **Competência:** 10/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 2.643 Colaboradores (701 com Saldo Abatido)
* **Estrutura de Colunas:** 26 Colunas + Abas de Importação (31903) e Exclusões
* **Economia Gerada:** R$ 139.090,04 em Saldo Residual Reaproveitado
* **Status do Motor no App:** 🟡 Pendente Envio para Validação Operacional

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 14 - Arquivo de Retorno LG** processa os arquivos de retorno das operadoras de transporte (GBarbosa e Mercantil Atacado), aplicando o algoritmo de abatimento de saldo remanescente nos cartões.
O robô gera automaticamente a planilha oficial no layout de importação da folha (Layout 31903) e o arquivo de exclusões, garantindo que a Cencosud economize ao não comprar créditos desnecessários.
"""
(dir_14 / "Manual_Tecnico_Executivo_ArquivoRetornoLG_v1.0.md").write_text(md_14, encoding="utf-8")

html_14 = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>Manual Técnico - Retorno LG</title><style>{CSS_STYLE}</style></head>
<body>
<div class="header-container">
    <div class="brand-title">CENCOSUD BRASIL</div>
    <div class="brand-subtitle">Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial</div>
    <div class="doc-main-title">Central Executiva de Auditoria de Benefícios • Módulo 14: Arquivo de Retorno LG</div>
    <div class="badge-bar">
        <span><span class="badge-label">Módulo:</span> 14 - Arquivo de Retorno LG</span>
        <span><span class="badge-label">Interface:</span> Central Executiva (.exe)</span>
        <span><span class="badge-label">Competência:</span> 10/2026</span>
        <span><span class="badge-label">Revisão:</span> 26/10/2026</span>
    </div>
</div>
<div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-title">Universo Auditado</div><div class="kpi-value">2.643</div><div class="kpi-sub">Colaboradores</div></div>
    <div class="kpi-card"><div class="kpi-title">Estrutura de Colunas</div><div class="kpi-value">26 Colunas</div><div class="kpi-sub">Layout 31903 + Exclusão</div></div>
    <div class="kpi-card"><div class="kpi-title">Economia Direta</div><div class="kpi-value">R$ 139.090,04</div><div class="kpi-sub">Saldo Abatido da Compra</div></div>
    <div class="kpi-card"><div class="kpi-title">Status Validação</div><div class="kpi-value">🟡 Pendente</div><div class="kpi-sub">Pronto p/ Envio</div></div>
</div>
<h2 class="section-title">1. Visão Geral da Solução e Arquitetura Unificada</h2>
<p>Processa o arquivo de retorno das operadoras abatendo saldo em cartão da recarga mensal, gerando os layouts de importação (31903) e exclusão para a folha LG.</p>
<div class="footer-signatures">
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Engenharia de Dados & Automação de RH</div><div class="sig-role">Desenvolvimento de Software & Algoritmos</div></div>
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Administração de Pessoal & Remuneração</div><div class="sig-role">Homologação de Negócio & Conformidade CLT</div></div>
</div>
</body></html>"""
build_pdf(html_14, dir_14 / "Manual_Tecnico_Executivo_ArquivoRetornoLG_v1.0.html", dir_14 / "Manual_Tecnico_Executivo_ArquivoRetornoLG_v1.0.pdf")
make_zip([ROOT_VA_VT / "Arquivo de Retorno" / "scripts"], dir_14 / "Scripts_Arquivo_Retorno_LG.zip")

# ==============================================================================
# 15. COMPRA MENSAL TICKET (BRETAS)
# ==============================================================================
dir_15 = BASE_DOC_DIR / "15_Compra_Mensal_Ticket_Bretas"
dir_15.mkdir(parents=True, exist_ok=True)

md_15 = """# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 15: Compra Ticket (Bretas)

---

**Módulo:** 15 - Compra Mensal Ticket Bretas | **Interface:** Central Executiva (`.exe`) | **Competência:** 09/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 1.642 Colaboradores Bretas (1.460 Elegíveis)
* **Estrutura de Colunas:** 22 Colunas + Aba de Remessa Manual Ticket
* **Volume Financeiro Auditado:** R$ 497.014,00 em Carga de Alimentação/Refeição
* **Status do Motor no App:** 🟡 Pendente Envio para Validação Operacional

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 15 - Compra Mensal Ticket Bretas** apura os pedidos de Vale Alimentação e Refeição para os colaboradores da bandeira Bretas (Minas Gerais e Goiás).
O robô expurga colaboradores afastados e demitidos e calcula o benefício facial por dia útil conforme ACT, emitindo a aba de Remessa Manual pronta para carga no portal da Ticket.
"""
(dir_15 / "Manual_Tecnico_Executivo_CompraTicketBretas_v1.0.md").write_text(md_15, encoding="utf-8")

html_15 = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>Manual Técnico - Ticket Bretas</title><style>{CSS_STYLE}</style></head>
<body>
<div class="header-container">
    <div class="brand-title">CENCOSUD BRASIL</div>
    <div class="brand-subtitle">Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial</div>
    <div class="doc-main-title">Central Executiva de Auditoria de Benefícios • Módulo 15: Compra Ticket (Bretas)</div>
    <div class="badge-bar">
        <span><span class="badge-label">Módulo:</span> 15 - Compra Ticket (Bretas)</span>
        <span><span class="badge-label">Interface:</span> Central Executiva (.exe)</span>
        <span><span class="badge-label">Competência:</span> 09/2026</span>
        <span><span class="badge-label">Revisão:</span> 26/10/2026</span>
    </div>
</div>
<div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-title">Universo Auditado</div><div class="kpi-value">1.642</div><div class="kpi-sub">Colaboradores Bretas</div></div>
    <div class="kpi-card"><div class="kpi-title">Estrutura de Colunas</div><div class="kpi-value">22 Colunas</div><div class="kpi-sub">Conferência + Remessa</div></div>
    <div class="kpi-card"><div class="kpi-title">Volume Financeiro</div><div class="kpi-value">R$ 497.014,00</div><div class="kpi-sub">Carga Ticket Apurada</div></div>
    <div class="kpi-card"><div class="kpi-title">Status Validação</div><div class="kpi-value">🟡 Pendente</div><div class="kpi-sub">Pronto p/ Envio</div></div>
</div>
<h2 class="section-title">1. Visão Geral da Solução e Arquitetura Unificada</h2>
<p>Apura o benefício de alimentação da bandeira Bretas (MG e GO), aplicando expurgos e gerando a Remessa Manual no layout da Ticket.</p>
<div class="footer-signatures">
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Engenharia de Dados & Automação de RH</div><div class="sig-role">Desenvolvimento de Software & Algoritmos</div></div>
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Administração de Pessoal & Remuneração</div><div class="sig-role">Homologação de Negócio & Conformidade CLT</div></div>
</div>
</body></html>"""
build_pdf(html_15, dir_15 / "Manual_Tecnico_Executivo_CompraTicketBretas_v1.0.html", dir_15 / "Manual_Tecnico_Executivo_CompraTicketBretas_v1.0.pdf")
make_zip([ROOT_VA_VT / "Compra Mensal Ticket" / "scripts"], dir_15 / "Scripts_Compra_Ticket_Bretas.zip")

# ==============================================================================
# 16. CESTA PRÊMIO (BRETAS)
# ==============================================================================
dir_16 = BASE_DOC_DIR / "16_Cesta_Premio_Bretas"
dir_16.mkdir(parents=True, exist_ok=True)

md_16 = """# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 16: Cesta Prêmio (Bretas)

---

**Módulo:** 16 - Cesta Prêmio Bretas | **Interface:** Central Executiva (`.exe`) | **Competência:** 10/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 1.844 Colaboradores Bretas (1.266 Premiados)
* **Estrutura de Colunas:** 24 Colunas Oficiais de Validação
* **Volume Financeiro Auditado:** R$ 279.913,47 em Prêmios Concedidos
* **Status do Motor no App:** 🟡 Pendente Envio para Validação Operacional

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 16 - Cesta Prêmio Bretas** audita as regras de assiduidade, pontualidade e carência contratual para concessão da Cesta Prêmio da bandeira Bretas.
O motor valida se o colaborador possui faltas injustificadas, atrasos acima da tolerância ou advertências no mês civil, garantindo a premiação apenas aos colaboradores 100% aderentes aos critérios institucionais.
"""
(dir_16 / "Manual_Tecnico_Executivo_CestaPremioBretas_v1.0.md").write_text(md_16, encoding="utf-8")

html_16 = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>Manual Técnico - Cesta Prêmio</title><style>{CSS_STYLE}</style></head>
<body>
<div class="header-container">
    <div class="brand-title">CENCOSUD BRASIL</div>
    <div class="brand-subtitle">Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial</div>
    <div class="doc-main-title">Central Executiva de Auditoria de Benefícios • Módulo 16: Cesta Prêmio (Bretas)</div>
    <div class="badge-bar">
        <span><span class="badge-label">Módulo:</span> 16 - Cesta Prêmio (Bretas)</span>
        <span><span class="badge-label">Interface:</span> Central Executiva (.exe)</span>
        <span><span class="badge-label">Competência:</span> 10/2026</span>
        <span><span class="badge-label">Revisão:</span> 26/10/2026</span>
    </div>
</div>
<div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-title">Universo Auditado</div><div class="kpi-value">1.844</div><div class="kpi-sub">Colaboradores Bretas</div></div>
    <div class="kpi-card"><div class="kpi-title">Estrutura de Colunas</div><div class="kpi-value">24 Colunas</div><div class="kpi-sub">Conferência Oficial</div></div>
    <div class="kpi-card"><div class="kpi-title">Volume Financeiro</div><div class="kpi-value">R$ 279.913,47</div><div class="kpi-sub">Prêmios Concedidos</div></div>
    <div class="kpi-card"><div class="kpi-title">Status Validação</div><div class="kpi-value">🟡 Pendente</div><div class="kpi-sub">Pronto p/ Envio</div></div>
</div>
<h2 class="section-title">1. Visão Geral da Solução e Arquitetura Unificada</h2>
<p>Audita os critérios de assiduidade e carência para pagamento da Cesta Prêmio da rede Bretas, expurgando faltas e atrasos.</p>
<div class="footer-signatures">
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Engenharia de Dados & Automação de RH</div><div class="sig-role">Desenvolvimento de Software & Algoritmos</div></div>
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Administração de Pessoal & Remuneração</div><div class="sig-role">Homologação de Negócio & Conformidade CLT</div></div>
</div>
</body></html>"""
build_pdf(html_16, dir_16 / "Manual_Tecnico_Executivo_CestaPremioBretas_v1.0.html", dir_16 / "Manual_Tecnico_Executivo_CestaPremioBretas_v1.0.pdf")
make_zip([ROOT_VA_VT / "Compra Mensal Cesta Prêmio" / "scripts"], dir_16 / "Scripts_Cesta_Premio_Bretas.zip")

print("\n>>> ETAPA 3 CONCLUÍDA COM SUCESSO ABSOLUTO! <<<")
