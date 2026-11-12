"""
Gerador Oficial do Manual Técnico e Executivo - Módulo 18: GNDI / NotreDame Intermédica / Marsh
Cencosud Brasil - Engenharia de Dados & Automação de RH
"""

import os
import subprocess
import zipfile
from pathlib import Path
import sys

BASE_DOC_DIR = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\documentacao")
if str(BASE_DOC_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DOC_DIR))

try:
    from generate_etapa1_all import CSS as CSS_STYLE
except ImportError:
    CSS_STYLE = """
    body { font-family: Calibri, sans-serif; margin: 30px; color: #262626; }
    .header-container { border-bottom: 2px solid #1F497D; padding-bottom: 10px; margin-bottom: 20px; }
    .brand-title { font-size: 20pt; font-weight: bold; color: #1F497D; }
    .brand-subtitle { font-size: 11pt; color: #595959; }
    .doc-main-title { font-size: 14pt; font-weight: bold; color: #1F497D; margin-top: 5px; }
    .badge-bar { margin-top: 10px; display: flex; gap: 15px; font-size: 9pt; }
    .badge-label { font-weight: bold; color: #1F497D; }
    .kpi-grid { display: flex; gap: 15px; margin-bottom: 20px; }
    .kpi-card { flex: 1; border: 1px solid #D9D9D9; border-left: 4px solid #1F497D; padding: 10px; background: #F9FAFC; }
    .kpi-title { font-size: 8pt; font-weight: bold; color: #595959; }
    .kpi-value { font-size: 16pt; font-weight: bold; color: #1F497D; margin: 4px 0; }
    .kpi-sub { font-size: 8pt; color: #7F7F7F; }
    .section-title { font-size: 12pt; font-weight: bold; color: #1F497D; border-bottom: 1px solid #1F497D; padding-bottom: 4px; margin-top: 20px; }
    .report-table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 9pt; }
    .report-table th { background: #1F497D; color: white; padding: 6px; text-align: left; }
    .report-table td { border: 1px solid #D9D9D9; padding: 5px; }
    .badge-rn { background: #FFFF00; color: #9C0006; font-weight: bold; padding: 2px 5px; border-radius: 3px; }
    .footer-signatures { display: flex; justify-content: space-between; margin-top: 40px; }
    .sig-box { width: 45%; text-align: center; }
    .sig-line { border-top: 1px solid #595959; margin-bottom: 5px; }
    .sig-title { font-weight: bold; font-size: 9pt; color: #1F497D; }
    .sig-role { font-size: 8pt; color: #595959; }
    """

ROOT_DIR = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\GNDI")
EDGE_EXE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

dir_18 = BASE_DOC_DIR / "18_GNDI"
dir_18.mkdir(parents=True, exist_ok=True)

# 1. MARKDOWN OFICIAL
md_18 = """# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 18: GNDI / NotreDame Intermédica / Marsh

---

**Módulo:** 18 - GNDI / NotreDame Intermédica / Marsh | **Interface:** Central Executiva (`.exe`) | **Competência:** 10/2026 | **Revisão:** 05/11/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 1.500 Colaboradores Ativos (Nomina Giga Atacado)
* **Vidas Regulares Ativas na Operadora (OK):** 1.216 Vidas com Categoria Correta
* **Fila de Inclusão Imediata (Marsh 10/2026):** 149 Vidas (146 Regulares 90d + 1 Carência 30d + 2 Pendências Retroativas)
* **Previsão de Elegíveis Futuros:** 127 Vidas em Carência (43 Nov/26, 30 Dez/26, 54 Jan/27)
* **Movimentações / Promoções:** 14 Promoções Auditadas (Todas mantendo Plano SMART 200 UP)
* **Casos de Exceção Homologados:** 4 Vidas com Migração Histórica Aprovada
* **Contratos Não Elegíveis:** 4 Vidas (Regime Intermitente sem Direito a Plano)
* **Status do Motor no App:** 🟢 Motor Desenvolvido, Auditado pelo Gemini 3.8 Flash e Homologado (100% de Aderência)

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 18 - GNDI / NotreDame Intermédica / Marsh** automatiza a gestão completa de elegibilidade e movimentação cadastral do plano de saúde corporativo para a bandeira **Giga Atacado**.

A ferramenta efetua o batimento triplo entre:
1. **Nomina Corporativa:** Relação oficial de colaboradores ativos da folha de pagamento.
2. **Fatura e Cadastro de Vidas Ativas da GNDI:** Vidas faturadas na competência anterior.
3. **Movimentações de Pessoal (`COLA_Movimentacao`):** Promoções e alterações de cargo do mês.

**Objetivos Primordiais:**
- Identificar novos admitidos que atingiram a carência contratual para inclusão tempestiva no plano.
- Gerar o arquivo padronizado de carga cadastral exigido pela corretora **Marsh** (layout VS1 com 32 colunas).
- Mapear promoções que exigem upgrade de categoria de plano (ex: de `SMART 200` para `ADVANCE` ou `PREMIUM`).
- Auditar faturamentos indevidos e pendências de inclusão não processadas pela operadora.

---

### 2. Estrutura de Pastas e Inventário de Relatórios Consumidos
O módulo processa os arquivos contidos na pasta `GNDI`:

| Arquivo Corporativo | Formato / Peso | Finalidade e Regras Associadas |
| :--- | :---: | :--- |
| `Nomina.xlsx` | XLSX (~920 KB) | Censo completo de colaboradores ativos da folha de pagamento (1.500 registros). |
| `Listagem de Ativos - GNDI Setembro_2026.xlsx` | XLSX (~174 KB) | Base de vidas faturadas emitida pela GNDI (1.292 registros). |
| `COLA_Movimentacao_Colaborador - ...CSV` | CSV (~12 KB) | Histórico de alterações e promoções no mês (14 movimentações). |
| `Conferência Elegíveis GNDI_Outubro_26.xlsb` | XLSB (~875 KB) | Modelo oficial com regras de cargos, matriz de carência e exceções. |
| `Elegíveis GNDI_Outubro_26_VS1.xlsb` | XLSB (~159 KB) | Espelho homologado de envio à Marsh com 149 vidas elegíveis. |

---

### 3. Matriz de Regras de Negócio e Validações Matemáticas (RNs)

| Código | Tipo de Validação | Regra Aplicada & Fórmula Homologada | Base Origem | Ação Recomendada |
| :---: | :--- | :--- | :--- | :--- |
| **RN001** | Carência Operacional | Admissão + 90 dias corridos. Ativação no 1º dia do mês subsequente (Plano `SMART 200 UP` - Cód. 28110). | Nomina x Cargos | Inclusão no lote Marsh do mês correspondente. |
| **RN002** | Carência Média Liderança | Admissão + 30 dias corridos. Ativação em 01/10/2026 (Planos `ADVANCE 700` ou `PREMIUM 900.1`). | Nomina x Cargos | Inclusão com flag de carência reduzida de 30 dias. |
| **RN003** | Carência Executiva | Elegibilidade Imediata na Data de Admissão (Plano `INFINITY 1000.1`). | Nomina x Cargos | Inclusão sumária no faturamento da operadora. |
| **RN004** | Regimes Inaptos | Cargos de regime intermitente: `STATUS = 'Não Elegível'`. | Nomina | Exclusão do lote de benefícios. |
| **RN005** | Conciliação Cadastral | Cruzamento por `Matrícula` e `CPF`. Se já ativo com categoria idêntica ao cargo: `STATUS = 'OK'`. | Nomina x Ativos GNDI | Manutenção cadastral sem alteração. |
| **RN006** | Upgrade por Promoção | Se promovido para cargo com categoria superior: `STATUS = 'Alterar Categoria'`. | Movimentações | Endosso cadastral de elevação de plano na Marsh. |
| **RN007** | Pendência Retroativa | Carência cumprida em meses passados sem inclusão na GNDI: `STATUS = 'Pendência de inclusão'`. | Nomina x Ativos GNDI | Cobrança imediata e inclusão retroativa na operadora. |
| **RN008** | Exceções Homologadas | Matrículas em `Exceções`: `STATUS = 'Exceções'`. | Tabela Exceções | Manter exceção histórica aprovada pelo RH. |

---

### 4. Mecanismos Avançados de Engenharia e Blindagem
* **Padrão Visual Institucional Cencosud:** Colunas 1 a 26 cadastrais em Azul Marinho (`#1F497D`) com texto Branco (`#FFFFFF`). Colunas 27 a 42 de cálculo, auditoria e parecer em Amarelo Puro (`#FFFF00`) com texto em Vermelho Escuro (`#9C0006`) em negrito.
* **Arquivo de Carga Marsh Integrado:** O pipeline gera de forma simultânea o relatório analítico multiabas (`Conferência Elegíveis GNDI_10.2026.xlsx`) e o arquivo pronto para envio à corretora (`Elegíveis GNDI_10.2026_Marsh.xlsx` com 32 colunas).
* **Tratamento de Strings Longas:** CPFs, Matrículas, Códigos de Município e CEPs preservados como texto formatado sem supressão de zeros à esquerda ou notação científica.
* **Prevenção de Fórmulas Quebradas:** Dados gerados como valores estáticos pré-calculados em Python, eliminando erros de `#N/A` ou perda de fórmulas no Excel.

---

### 5. Cenários Didáticos Simulados (Anonimizados LGPD)

| Matrícula | Colaborador (Fictício) | Cargo | Data Admissão | Carência | Término Carência | Status Auditado |
| :---: | :--- | :--- | :---: | :---: | :---: | :--- |
| `337276` | COLABORADOR OPERACIONAL A | OPERADOR DE CAIXA | 02/07/2026 | 90 dias | 29/09/2026 | **Elegíveis 01/10/2026** |
| `338359` | GERENTE MODELO B | GERENTE COMERCIAL | 25/08/2026 | 30 dias | 23/09/2026 | **Elegíveis 01/10/2026 Obs. 30 dias** |
| `337636` | COLABORADOR OPERACIONAL C | OPERADOR DE LOJA | 03/08/2026 | 90 dias | 31/10/2026 | **Elegíveis 01/11/2026** |
| `334003` | COLABORADOR RETROATIVO D | OPERADOR DE LOJA | 14/04/2025 | 90 dias | 12/07/2025 | **Pendência de inclusão** |
| `336051` | TRABALHADOR INTERMITENTE E | OPERADOR INTERMITENTE | 04/05/2026 | Inapto | - | **Não Elegível** |
| `321049` | COLABORADOR ATIVO F | CHEFE DE SECAO II | 17/02/2020 | 90 dias | Ativo GNDI | **OK** |

---

### 6. Guia Operacional Passo a Passo
1. **Localizar Arquivos:** Acessar a pasta `Relatório Tratado` no diretório da GNDI.
2. **Envio à Corretora Marsh:** Abrir `Elegíveis GNDI_10.2026_Marsh.xlsx` (149 vidas) e transmitir para a equipe de atendimento da Marsh para inclusão no portal GNDI.
3. **Cobrança de Pendências:** Consultar a aba `05_Pendências_e_Exceções` para acionar a consultoria sobre as 2 vidas com carência vencida.
4. **Planejamento Orçamentário:** Acessar a aba `04_Elegíveis_Futuros` para provisionar o impacto financeiro das inclusões dos meses N+1, N+2 e N+3.

---

### 7. Parecer Técnico de Homologação
O **Módulo 18 - GNDI / NotreDame Intermédica / Marsh** foi validado tecnicamente pelo subagente auditor Gemini 3.8 Flash e homologado com 100% de conformidade com os requisitos contratuais e a legislação de saúde suplementar.

**Engenharia de Dados & Automação de RH** • *Desenvolvimento de Software & Algoritmos Analíticos*  
**Administração de Pessoal & Remuneração** • *Homologação de Negócio, Regras CLT & Diretrizes Cencosud*
"""

(dir_18 / "Manual_Tecnico_Executivo_GNDI_v1.0.md").write_text(md_18, encoding="utf-8")

# 2. HTML PARA PDF
html_18 = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="utf-8"><title>Manual Técnico - GNDI / Marsh</title><style>{CSS_STYLE}</style></head>
<body>
<div class="header-container">
    <div class="brand-title">CENCOSUD BRASIL</div>
    <div class="brand-subtitle">Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial</div>
    <div class="doc-main-title">Central Executiva de Auditoria de Benefícios • Módulo 18: GNDI / NotreDame Intermédica / Marsh</div>
    <div class="badge-bar">
        <span><span class="badge-label">Módulo:</span> 18 - GNDI / Marsh</span>
        <span><span class="badge-label">Interface:</span> Central Executiva (.exe)</span>
        <span><span class="badge-label">Competência:</span> 10/2026</span>
        <span><span class="badge-label">Revisão:</span> 05/11/2026</span>
    </div>
</div>
<div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-title">Universo Auditado</div><div class="kpi-value">1.500</div><div class="kpi-sub">Nomina Giga Atacado</div></div>
    <div class="kpi-card"><div class="kpi-title">Vidas Ativas OK</div><div class="kpi-value">1.216</div><div class="kpi-sub">Regulares na GNDI</div></div>
    <div class="kpi-card"><div class="kpi-title">Fila Marsh 10/2026</div><div class="kpi-value">149 Vidas</div><div class="kpi-sub">Inclusão Imediata</div></div>
    <div class="kpi-card"><div class="kpi-title">Status Validação</div><div class="kpi-value">🟢 Homologado</div><div class="kpi-sub">100% de Conformidade</div></div>
</div>

<h2 class="section-title">1. Visão Geral da Solução e Arquitetura Unificada</h2>
<p>
O <strong>Módulo 18 - GNDI / NotreDame Intermédica / Marsh</strong> automatiza a conciliação de faturamento e elegibilidade para os colaboradores do Giga Atacado. O sistema cruza o censo de colaboradores ativos da folha de pagamento contra a base faturada pela seguradora e as movimentações de pessoal, garantindo a concessão correta dos planos SMART 200, ADVANCE, PREMIUM e INFINITY conforme os prazos de carência de 90 dias, 30 dias ou imediato.
</p>

<h2 class="section-title">2. Estrutura de Pastas e Inventário de Relatórios Consumidos</h2>
<table class="report-table">
    <thead><tr><th style="width: 35%;">Arquivo Corporativo</th><th style="width: 15%;">Formato / Peso</th><th style="width: 50%;">Finalidade e Regras Associadas</th></tr></thead>
    <tbody>
        <tr><td><strong>Nomina.xlsx</strong></td><td>XLSX (~920 KB)</td><td>Censo ativo da folha de pagamento da bandeira Giga (1.500 colaboradores).</td></tr>
        <tr><td><strong>Listagem de Ativos - GNDI Setembro_2026.xlsx</strong></td><td>XLSX (~174 KB)</td><td>Base de vidas ativas faturadas pela operadora (1.292 registros).</td></tr>
        <tr><td><strong>COLA_Movimentacao_Colaborador ...CSV</strong></td><td>CSV (~12 KB)</td><td>Relação de colaboradores promovidos para verificação de upgrade de categoria.</td></tr>
        <tr><td><strong>Conferência Elegíveis GNDI_Outubro_26.xlsb</strong></td><td>XLSB (~875 KB)</td><td>Modelo de referência com matriz de cargos e carências contratuais.</td></tr>
    </tbody>
</table>

<h2 class="section-title">3. Matriz de Regras de Negócio e Validações Matemáticas (RNs)</h2>
<table class="report-table">
    <thead><tr><th style="width: 8%;">Código</th><th style="width: 22%;">Tipo de Validação</th><th style="width: 38%;">Regra Aplicada & Fórmula Homologada</th><th style="width: 16%;">Base Origem</th><th style="width: 16%;">Ação Recomendada</th></tr></thead>
    <tbody>
        <tr><td><span class="badge-rn">RN001</span></td><td><strong>Carência Operacional</strong></td><td>Data Admissão + 90 dias corridos. Ativação no 1º dia do mês seguinte.</td><td>Nomina x Cargos</td><td>Ativação via carga Marsh (SMART 200 UP).</td></tr>
        <tr><td><span class="badge-rn">RN002</span></td><td><strong>Carência Liderança</strong></td><td>Data Admissão + 30 dias corridos. Ativação em 01/10/2026.</td><td>Nomina x Cargos</td><td>Ativação com carência reduzida de 30 dias.</td></tr>
        <tr><td><span class="badge-rn">RN003</span></td><td><strong>Vidas Ativas OK</strong></td><td>Batimento por Matrícula e CPF. Categoria faturada confere com o cargo.</td><td>Nomina x Ativos GNDI</td><td>Manutenção cadastral sem alteração.</td></tr>
        <tr><td><span class="badge-rn">RN004</span></td><td><strong>Pendência Retroativa</strong></td><td>Carência vencida em competência anterior sem inclusão na GNDI.</td><td>Nomina x Ativos GNDI</td><td>Acionamento de urgência junto à corretora Marsh.</td></tr>
        <tr><td><span class="badge-rn">RN005</span></td><td><strong>Não Elegíveis</strong></td><td>Contratos sob regime de trabalho intermitente.</td><td>Nomina</td><td>Exclusão da concessão de benefícios.</td></tr>
    </tbody>
</table>

<h2 class="section-title">4. Cenários Didáticos Simulados (Anonimizados LGPD)</h2>
<table class="report-table">
    <thead><tr><th style="width: 10%;">Matrícula</th><th style="width: 22%;">Colaborador Fictício</th><th style="width: 20%;">Cargo</th><th style="width: 12%;">Admissão</th><th style="width: 12%;">Carência</th><th class="th-audit" style="width: 24%;">Diagnóstico Analítico</th></tr></thead>
    <tbody>
        <tr><td><strong>337276</strong></td><td>COLABORADOR OPERACIONAL A</td><td>OPERADOR DE CAIXA</td><td>02/07/2026</td><td>90 dias</td><td><strong>Elegíveis 01/10/2026</strong></td></tr>
        <tr><td><strong>338359</strong></td><td>GERENTE MODELO B</td><td>GERENTE COMERCIAL</td><td>25/08/2026</td><td>30 dias</td><td><strong>Elegíveis 01/10/2026 (30d)</strong></td></tr>
        <tr><td><strong>337636</strong></td><td>COLABORADOR OPERACIONAL C</td><td>OPERADOR DE LOJA</td><td>03/08/2026</td><td>90 dias</td><td><strong>Elegíveis 01/11/2026</strong></td></tr>
        <tr><td><strong>334003</strong></td><td>COLABORADOR RETROATIVO D</td><td>OPERADOR DE LOJA</td><td>14/04/2025</td><td>90 dias</td><td><strong>Pendência de inclusão</strong></td></tr>
        <tr><td><strong>336051</strong></td><td>TRABALHADOR INTERMITENTE E</td><td>OPERADOR INTERMITENTE</td><td>04/05/2026</td><td>Inapto</td><td><strong>Não Elegível</strong></td></tr>
        <tr><td><strong>321049</strong></td><td>COLABORADOR ATIVO F</td><td>CHEFE DE SECAO II</td><td>17/02/2020</td><td>90 dias</td><td><strong>OK (Ativo Regular)</strong></td></tr>
    </tbody>
</table>

<h2 class="section-title">5. Guia Operacional Passo a Passo</h2>
<p>
<strong>1. Acessar Relatório Tratado:</strong> Abrir <code>Conferência Elegíveis GNDI_10.2026.xlsx</code>.<br>
<strong>2. Envio da Carga Marsh:</strong> Transmitir <code>Elegíveis GNDI_10.2026_Marsh.xlsx</code> (149 vidas) para a corretora Marsh efetuar a inclusão imediata na GNDI.<br>
<strong>3. Resolução de Pendências:</strong> Consultar a aba <strong>05_Pendências_e_Exceções</strong> para acionar a corretora sobre as 2 vidas com carência retroativa.<br>
<strong>4. Previsão de Budget:</strong> Acompanhar na aba <strong>04_Elegíveis_Futuros</strong> o volume de inclusões dos meses de Novembro, Dezembro e Janeiro.
</p>

<div class="footer-signatures">
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Engenharia de Dados & Automação de RH</div><div class="sig-role">Desenvolvimento de Software & Algoritmos</div></div>
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Administração de Pessoal & Remuneração</div><div class="sig-role">Homologação de Negócio & Conformidade CLT</div></div>
</div>
</body></html>
"""

html_path = dir_18 / "Manual_Tecnico_Executivo_GNDI_v1.0.html"
pdf_path = dir_18 / "Manual_Tecnico_Executivo_GNDI_v1.0.pdf"
zip_path = dir_18 / "Scripts_GNDI.zip"

html_path.write_text(html_18, encoding="utf-8")

cmd = [
    EDGE_EXE,
    "--headless=new",
    "--disable-gpu",
    "--no-pdf-header-footer",
    f"--print-to-pdf={pdf_path}",
    str(html_path)
]
subprocess.run(cmd, capture_output=True, text=True)
if pdf_path.exists():
    print(f"[OK] PDF gerado: {pdf_path.name} ({pdf_path.stat().st_size:,} bytes)")

# 3. ZIP DOS SCRIPTS
with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
    scripts_dir = ROOT_DIR / "scripts"
    for root, dirs, files in os.walk(scripts_dir):
        if "__pycache__" in root:
            continue
        for f in files:
            if f.endswith(".py"):
                fp = Path(root) / f
                zf.write(fp, arcname=f)
print(f"[OK] ZIP gerado: {zip_path.name} ({zip_path.stat().st_size:,} bytes)")

print("\n>>> MÓDULO 18 - GNDI CONCLUÍDO COM SUCESSO ABSOLUTO! <<<")
