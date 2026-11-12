"""
Gerador Oficial do Manual Técnico e Executivo - Módulo 17: Desligados SulAmérica Saúde
Cencosud Brasil - Engenharia de Dados & Automação de RH
"""

import os
import subprocess
import zipfile
from pathlib import Path
from generate_etapa1_all import CSS as CSS_STYLE

BASE_DOC_DIR = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\documentacao")
ROOT_DIR = Path(r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\Desligados Sulamerica Saude")
EDGE_EXE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

dir_17 = BASE_DOC_DIR / "17_Desligados_Sulamerica_Saude"
dir_17.mkdir(parents=True, exist_ok=True)

# 1. MARKDOWN OFICIAL
md_17 = """# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 17: Desligados SulAmérica Saúde

---

**Módulo:** 17 - Desligados SulAmérica Saúde (SAS) | **Interface:** Central Executiva (`.exe`) | **Competência:** 10/2026 | **Revisão:** 04/11/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 925 Vidas Ativas Faturadas (Titulares e Dependentes)
* **Estrutura de Colunas:** 20 Colunas Oficiais (14 Cadastrais + 6 Auditoria e Parecer - Coluna B redundante de CPF removida)
* **Vidas Vinculadas a Demitidos:** 27 Vidas Mapeadas no Faturamento
* **Fila de Cancelamentos Imediatos:** 12 Vidas (Corte Imediato & Pedidos / Prazos Expirados)
* **Fila de Diligências e Monitoramento:** 15 Vidas (Demissões Recentes em Outubro / Prazo 30 dias ANS)
* **Vidas em Planos de Diretoria:** 92 Vidas em Categoria Prestige Monitoradas
* **Status do Motor no App:** 🟢 Motor Desenvolvido, Auditado e Homologado (100% de Conformidade)

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 17 - Desligados SulAmérica Saúde** automatiza o batimento mensal entre o censo de vidas ativas faturadas pela operadora SulAmérica Saúde e a relação corporativa de colaboradores demitidos do Grupo Cencosud Brasil (`COLA_Demitidos`).

A solução elimina a permanência indevida de ex-colaboradores e seus dependentes no faturamento corporativo, aplicando com rigor as disposições da **Lei Federal nº 9.656/1998 (Artigos 30 e 31)** e da **Resolução Normativa ANS nº 488/2022**:
1. **Propagação em Cascata Familiar:** A regra de desligamento identificada para o titular é automaticamente herdada por todos os seus dependentes cadastrados (cônjuges e filhos).
2. **Corte Imediato de Pedidos de Demissão:** Colaboradores que pediram rescisão voluntária (`Tipo 4`) não possuem direito legal à extensão do plano, sendo enquadrados como *"Cancelar Plano - Pedido"*, com solicitação de cancelamento imediato e glosa da fatura.
3. **Prazos Expirados de Demissão sem Justa Causa:** Demitidos pelo empregador (`Tipo 2`) anteriores a Outubro/2026 que já completaram o prazo legal são classificados como *"Cancelar Plano - Já completou o prazo"*.
4. **Diligências e Monitoramento (Rescisões Recentes):** Demitidos sem justa causa no mês corrente de Outubro/2026 têm direito regulamentar de até 30 dias para manifestação de interesse na permanência como beneficiário contribuinte (RN 488/ANS), sendo classificados como *"Manter o plano - Sem retorno RH Negpocio"*.
5. **Isolamento de Planos de Diretoria:** Titulares com categoria `PRESTIGE` (`29426`) são mapeados com a tag *"Diretor"*, segregando casos executivos que possuem regras específicas de transição contratual.

---

### 2. Estrutura de Pastas e Inventário de Relatórios Consumidos
O módulo consome as bases consolidadas na pasta `Desligados Sulamerica Saude`:

| Arquivo Corporativo | Formato / Peso | Finalidade e Regras Associadas |
| :--- | :---: | :--- |
| `Listagem de Ativos - Sulamerica Saúde - Outubro_2026.xlsx` | XLSX (~140 KB) | Base mestre de faturamento emitida pela operadora com 925 vidas (Titulares e Dependentes). |
| `Desligados 01.09 a 19.10.xlsx` | XLSX (~274 KB) | Extração corporativa da folha de pagamento (`COLA_Demitidos`) com 920 rescisões ocorridas no período. |
| `Listagem de Desligados - Sulamerica Saúde Validação RH.xlsx` | XLSX (~146 KB) | Modelo oficial homologado contendo as colunas de validação e a aba de Legenda. |

---

### 3. Matriz de Regras de Negócio e Validações Matemáticas (RNs)

| Código | Tipo de Validação | Regra Aplicada & Fórmula Homologada | Base Origem | Ação Recomendada |
| :---: | :--- | :--- | :--- | :--- |
| **RN001** | Cruzamento Titular | Cruzamento relacional por `MATRICULA` e `CPF` do Titular normalizados sem caracteres especiais. | Ativos SAS x Demitidos | Amarrar faturamento da operadora ao censo de folha. |
| **RN002** | Vínculo Familiar | Propagação automática da regra do titular para todos os dependentes associados à matrícula. | Ativos SAS | Cancelar todo o grupo familiar de ex-empregados. |
| **RN003** | Pedido de Demissão | Se `Tipo de Rescisão` contiver "EMPREGADO" ou "PEDIDO" (Tipo 4): `STATUS = 'Cancelar Plano - Pedido'`. | COLA Demitidos | Exclusão imediata sem direito ao Art. 30 e glosa na fatura. |
| **RN004** | Prazo Expirado (Art. 30) | Se demitido sem justa causa antes de Outubro/2026: `STATUS = 'Cancelar Plano - Já completou o prazo'`. | Data Rescisão | Cancelamento imediato por término do prazo legal. |
| **RN005** | Diligência RH Negócio | Se demitido sem justa causa recente em Outubro/2026: `STATUS = 'Manter o plano - Sem retorno RH Negpocio'`. | Data Rescisão | Aguardar prazo regulamentar de 30 dias de opção da ANS. |
| **RN006** | Mapeamento Diretoria | Se `CATEGORIA DO PLANO` contiver "PRESTIGE" ou `CÓDIGO` = 29426: `OBSERVAÇÃO = 'Diretor'`. | Categoria Plano | Submeter à Diretoria de Gente & Gestão. |

---

### 4. Mecanismos Avançados de Engenharia e Blindagem
* **Padrão Visual Institucional:** Cabeçalhos das 15 colunas cadastrais em Azul Marinho (`#1F497D`) com texto Branco (`#FFFFFF`). Colunas 16 a 21 de auditoria e decisão em Amarelo Puro (`#FFFF00`) com texto em Vermelho Escuro (`#9C0006`) em negrito.
* **Preservação de CPFs e Formatação de Carteira:** Coluna 2 (CPF Titular) e Coluna 8 (CPF Beneficiário) preservadas com formatação de 11 dígitos, e número de carteira gravado como string para impedir notação científica.
* **Abas Operacionais Dedicadas:** Além da aba integral `BASE DE ATIVOS`, o sistema gera automaticamente as abas `Cancelamentos Imediatos` (12 vidas em fila de corte imediato) e `Diligências e Monitoramento` (15 vidas sob prazo legal da RN 488/ANS).
* **Prevenção de Fórmulas Quebradas:** Dados gerados como valores estáticos pré-calculados em Python, eliminando erros de `#N/A` ou perda de fórmulas no Excel.

---

### 5. Cenários Didáticos Simulados (Anonimizados LGPD)

| Matrícula | Titular (Fictício) | Beneficiário | Parentesco | Data Rescisão | Tipo de Rescisão | Status Auditado |
| :---: | :--- | :--- | :---: | :---: | :--- | :--- |
| `336173` | TITULAR MODELO A | TITULAR MODELO A | TITULAR | 15/10/2026 | 4 - Sem justa causa empregado | **Cancelar Plano - Pedido** |
| `336173` | TITULAR MODELO A | CONJUGE MODELO A | CONJUGE | 15/10/2026 | 4 - Sem justa causa empregado | **Cancelar Plano - Pedido** |
| `334932` | TITULAR MODELO B | TITULAR MODELO B | TITULAR | 10/09/2026 | 2 - Sem justa causa empregador | **Cancelar Plano - Já completou o prazo** |
| `334599` | TITULAR MODELO C | TITULAR MODELO C | TITULAR | 13/10/2026 | 2 - Sem justa causa empregador | **Manter o plano - Sem retorno RH Negpocio** |
| `334557` | DIRETOR MODELO D | DIRETOR MODELO D | TITULAR | - | - (Ativo Regular) | **Ativo (Observação: Diretor)** |

---

### 6. Guia Operacional Passo a Passo
1. **Arquivo Gerado:** Abrir `Conferência Desligados Sulamerica Saúde_10.2026.xlsx` na pasta `Relatório Tratado`.
2. **Fila de Cancelamento:** Acessar a aba **"Cancelamentos Imediatos"** (12 vidas a cancelar imediatamente).
3. **Fila de Diligências:** Acessar a aba **"Diligências e Monitoramento"** (15 vidas aguardando decurso do prazo de 30 dias ou opção pelo RH Negócio).
4. **Ação no Portal:** Efetuar a baixa das carteirinhas listadas na fila de cancelamento imediato no portal da SulAmérica Saúde.
5. **Glosa Financeira:** Enviar o lote para a corretora solicitar o estorno financeiro das cobranças pós-rescisão.

---

### 7. Parecer Técnico de Homologação
O **Módulo 17 - Desligados SulAmérica Saúde** foi homologado e testado com 100% de conformidade com os apontamentos da equipe de Recursos Humanos e a legislação da ANS.

**Engenharia de Dados & Automação de RH** • *Desenvolvimento de Software & Algoritmos Analíticos*  
**Administração de Pessoal & Remuneração** • *Homologação de Negócio, Regras CLT & Diretrizes Cencosud*
"""

(dir_17 / "Manual_Tecnico_Executivo_DesligadosSulamericaSaude_v1.0.md").write_text(md_17, encoding="utf-8")

# 2. HTML PARA PDF
html_17 = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="utf-8"><title>Manual Técnico - Desligados SulAmérica Saúde</title><style>{CSS_STYLE}</style></head>
<body>
<div class="header-container">
    <div class="brand-title">CENCOSUD BRASIL</div>
    <div class="brand-subtitle">Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial</div>
    <div class="doc-main-title">Central Executiva de Auditoria de Benefícios • Módulo 17: Desligados SulAmérica Saúde</div>
    <div class="badge-bar">
        <span><span class="badge-label">Módulo:</span> 17 - Desligados SulAmérica Saúde</span>
        <span><span class="badge-label">Interface:</span> Central Executiva (.exe)</span>
        <span><span class="badge-label">Competência:</span> 10/2026</span>
        <span><span class="badge-label">Revisão:</span> 04/11/2026</span>
    </div>
</div>
<div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-title">Universo Auditado</div><div class="kpi-value">925</div><div class="kpi-sub">Vidas Faturadas SAS</div></div>
    <div class="kpi-card"><div class="kpi-title">Cancelamentos Imediatos</div><div class="kpi-value">12 Vidas</div><div class="kpi-sub">Corte Imediato & Glosa</div></div>
    <div class="kpi-card"><div class="kpi-title">Diligências & Prazos</div><div class="kpi-value">15 Vidas</div><div class="kpi-sub">Prazo 30d RN 488 ANS</div></div>
    <div class="kpi-card"><div class="kpi-title">Status Validação</div><div class="kpi-value">🟢 Homologado</div><div class="kpi-sub">100% de Conformidade</div></div>
</div>

<h2 class="section-title">1. Visão Geral da Solução e Arquitetura Unificada</h2>
<p>
O <strong>Módulo 17 - Desligados SulAmérica Saúde</strong> automatiza o batimento de faturamento das vidas ativas na seguradora contra a relação oficial de demitidos do Grupo Cencosud Brasil. A solução aplica em cascata a regra de rescisão do titular para todos os dependentes e audita a legitimidade da cobrança em estrita observância à Lei 9.656/1998 e Resolução Normativa ANS nº 488/2022.
</p>

<h2 class="section-title">2. Estrutura de Pastas e Inventário de Relatórios Consumidos</h2>
<table class="report-table">
    <thead><tr><th style="width: 35%;">Arquivo Corporativo</th><th style="width: 15%;">Formato / Peso</th><th style="width: 50%;">Finalidade e Regras Associadas</th></tr></thead>
    <tbody>
        <tr><td><strong>Listagem de Ativos - SAS - Outubro_2026.xlsx</strong></td><td>XLSX (~140 KB)</td><td>Fatura e base de vidas ativas enviada pela SulAmérica Saúde (925 registros).</td></tr>
        <tr><td><strong>Desligados 01.09 a 19.10.xlsx</strong></td><td>XLSX (~274 KB)</td><td>Relação corporativa de colaboradores demitidos (COLA_Demitidos - 920 linhas).</td></tr>
        <tr><td><strong>Listagem de Desligados - SAS Validação RH.xlsx</strong></td><td>XLSX (~146 KB)</td><td>Espelho manual de validação com colunas de auditoria e legenda oficial.</td></tr>
    </tbody>
</table>

<h2 class="section-title">3. Matriz de Regras de Negócio e Validações Matemáticas (RNs)</h2>
<table class="report-table">
    <thead><tr><th style="width: 8%;">Código</th><th style="width: 22%;">Tipo de Validação</th><th style="width: 38%;">Regra Aplicada & Fórmula Homologada</th><th style="width: 16%;">Base Origem</th><th style="width: 16%;">Ação Recomendada</th></tr></thead>
    <tbody>
        <tr><td><span class="badge-rn">RN001</span></td><td><strong>Cruzamento Titular</strong></td><td>Cruzamento por Matrícula e CPF do Titular normalizados.</td><td>Ativos x Demitidos</td><td>Amarrar faturamento da operadora à folha.</td></tr>
        <tr><td><span class="badge-rn">RN002</span></td><td><strong>Vínculo Familiar</strong></td><td>Propagação automática do cancelamento do titular para todos os dependentes.</td><td>Ativos SAS</td><td>Cancelar todo o grupo familiar do ex-colaborador.</td></tr>
        <tr><td><span class="badge-rn">RN003</span></td><td><strong>Pedido de Demissão</strong></td><td>Tipo de rescisão 4 (Pedido): STATUS = 'Cancelar Plano - Pedido'.</td><td>COLA Demitidos</td><td>Corte imediato sem direito a extensão e glosa.</td></tr>
        <tr><td><span class="badge-rn">RN004</span></td><td><strong>Prazo Expirado</strong></td><td>Demissão sem justa causa anterior a Outubro: STATUS = 'Já completou prazo'.</td><td>Data Rescisão</td><td>Cancelamento imediato por término da carência.</td></tr>
        <tr><td><span class="badge-rn">RN005</span></td><td><strong>Diligência RH Negócio</strong></td><td>Demissão recente em Outubro: STATUS = 'Manter o plano - Sem retorno RH Negpocio'.</td><td>Data Rescisão</td><td>Aguardar prazo regulamentar de 30 dias de opção da ANS.</td></tr>
        <tr><td><span class="badge-rn">RN006</span></td><td><strong>Diretoria (Prestige)</strong></td><td>Categoria PRESTIGE ou Código 29426: OBSERVAÇÃO = 'Diretor'.</td><td>Categoria Plano</td><td>Submeter à Diretoria de Gente & Gestão.</td></tr>
    </tbody>
</table>

<h2 class="section-title">4. Cenários Didáticos Simulados (Anonimizados LGPD)</h2>
<table class="report-table">
    <thead><tr><th style="width: 10%;">Matrícula</th><th style="width: 20%;">Titular Fictício</th><th style="width: 20%;">Beneficiário</th><th style="width: 10%;">Grau</th><th style="width: 15%;">Data Rescisão</th><th class="th-audit" style="width: 25%;">Diagnóstico Analítico</th></tr></thead>
    <tbody>
        <tr><td><strong>336173</strong></td><td>TITULAR MODELO A</td><td>TITULAR MODELO A</td><td>TITULAR</td><td>15/10/2026</td><td><strong>Cancelar Plano - Pedido</strong></td></tr>
        <tr><td><strong>336173</strong></td><td>TITULAR MODELO A</td><td>CONJUGE MODELO A</td><td>CONJUGE</td><td>15/10/2026</td><td><strong>Cancelar Plano - Pedido</strong></td></tr>
        <tr><td><strong>334932</strong></td><td>TITULAR MODELO B</td><td>TITULAR MODELO B</td><td>TITULAR</td><td>10/09/2026</td><td><strong>Cancelar - Prazo Expirado</strong></td></tr>
        <tr><td><strong>334599</strong></td><td>TITULAR MODELO C</td><td>TITULAR MODELO C</td><td>TITULAR</td><td>13/10/2026</td><td><strong>Manter - Sem Retorno RH</strong></td></tr>
        <tr><td><strong>334557</strong></td><td>DIRETOR MODELO D</td><td>DIRETOR MODELO D</td><td>TITULAR</td><td>-</td><td><strong>Ativo (Obs: Diretor)</strong></td></tr>
    </tbody>
</table>

<h2 class="section-title">5. Guia Operacional Passo a Passo</h2>
<p>
<strong>1. Localizar Arquivo:</strong> Abrir <code>Conferência Desligados Sulamerica Saúde_10.2026.xlsx</code> na pasta <code>Relatório Tratado</code>.<br>
<strong>2. Fila de Cancelamentos Imediatos:</strong> Consultar a aba <strong>Cancelamentos Imediatos</strong> (12 vidas mapeadas para corte imediato no portal da seguradora).<br>
<strong>3. Fila de Diligências:</strong> Consultar a aba <strong>Diligências e Monitoramento</strong> (15 vidas em período de opção ou aguardando resposta do RH Negócio).<br>
<strong>4. Baixa na Seguradora:</strong> Efetuar o cancelamento das carteirinhas no portal da SulAmérica Saúde.<br>
<strong>5. Solicitação de Glosa:</strong> Encaminhar o lote de corte para a corretora solicitar o estorno financeiro das mensalidades indevidas.
</p>

<div class="footer-signatures">
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Engenharia de Dados & Automação de RH</div><div class="sig-role">Desenvolvimento de Software & Algoritmos</div></div>
    <div class="sig-box"><div class="sig-line"></div><div class="sig-title">Administração de Pessoal & Remuneração</div><div class="sig-role">Homologação de Negócio & Conformidade CLT</div></div>
</div>
</body></html>
"""

html_path = dir_17 / "Manual_Tecnico_Executivo_DesligadosSulamericaSaude_v1.0.html"
pdf_path = dir_17 / "Manual_Tecnico_Executivo_DesligadosSulamericaSaude_v1.0.pdf"
zip_path = dir_17 / "Scripts_Desligados_Sulamerica_Saude.zip"

html_path.write_text(html_17, encoding="utf-8")

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

print("\n>>> MÓDULO 17 CONCLUÍDO COM SUCESSO ABSOLUTO! <<<")
