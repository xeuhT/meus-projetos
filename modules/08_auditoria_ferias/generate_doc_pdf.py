import os
import subprocess
import pymupdf

this_dir = os.path.dirname(os.path.abspath(__file__))
if os.path.basename(this_dir) == "documentacao":
    output_dir = os.path.dirname(this_dir)
    doc_dir = this_dir
else:
    output_dir = this_dir
    doc_dir = os.path.join(output_dir, "documentacao")

os.makedirs(doc_dir, exist_ok=True)
html_path = os.path.join(doc_dir, "documentacao_auditoria_ferias.html")
html_root_path = os.path.join(output_dir, "documentacao_auditoria_ferias.html")
pdf_path = os.path.join(output_dir, "Manual_Tecnico_e_Executivo_Auditoria_Ferias_Cencosud.pdf")
pdf_doc_path = os.path.join(doc_dir, "Manual_Tecnico_e_Executivo_Auditoria_Ferias_Cencosud.pdf")

html_content = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Manual Técnico e Executivo - Auditoria de Férias e Folha | Cencosud Brasil (v004)</title>
<style>
  @page {
    size: A4 portrait;
    margin: 8mm 11mm 8mm 11mm;
  }
  * {
    box-sizing: border-box;
    margin: 0;
    padding: 0;
  }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
    color: #1e293b;
    background: #ffffff;
    font-size: 7.6pt;
    line-height: 1.24;
  }
  .page {
    width: 100%;
    height: 278mm;
    max-height: 278mm;
    page-break-after: always;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    overflow: hidden;
  }
  .page:last-child {
    page-break-after: avoid;
  }
  .page-content {
    flex: 1;
  }
  .footer {
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-top: 1px solid #cbd5e1;
    padding-top: 3px;
    font-size: 6.9pt;
    color: #64748b;
    margin-top: 3px;
  }

  /* Header */
  .brand-header {
    border-bottom: 2px solid #095AA5;
    padding-bottom: 3px;
    margin-bottom: 4px;
  }
  .brand-title {
    font-size: 13.5pt;
    font-weight: 900;
    color: #095AA5;
    letter-spacing: -0.5px;
    text-transform: uppercase;
  }
  .brand-subtitle {
    font-size: 7.3pt;
    color: #475569;
    font-weight: 600;
    margin-top: 1px;
  }
  .project-title {
    font-size: 9.6pt;
    font-weight: 800;
    color: #0f172a;
    margin-top: 2px;
  }
  .meta-banner {
    background: #095AA5;
    color: #ffffff;
    padding: 3px 6.5px;
    border-radius: 3px;
    font-size: 6.9pt;
    display: flex;
    justify-content: space-between;
    margin-top: 3px;
    font-weight: 500;
  }
  .meta-banner span strong {
    color: #fde047;
  }

  /* KPI Cards */
  .kpi-container {
    display: grid;
    grid-template-columns: 1fr 1.2fr 1.2fr 1.15fr;
    gap: 5px;
    margin: 4px 0 5px 0;
  }
  .kpi-card {
    border: 1px solid #e2e8f0;
    border-radius: 4px;
    padding: 3.5px 5.5px;
    background: #f8fafc;
    border-top: 3px solid #095AA5;
  }
  .kpi-card.kpi-blue { border-top-color: #0284c7; }
  .kpi-card.kpi-amber { border-top-color: #d97706; }
  .kpi-card.kpi-emerald { border-top-color: #059669; }
  .kpi-label {
    font-size: 5.7pt;
    font-weight: 700;
    color: #64748b;
    text-transform: uppercase;
    letter-spacing: 0.3px;
  }
  .kpi-value {
    font-size: 11.5pt;
    font-weight: 800;
    color: #0f172a;
    margin: 1px 0;
    letter-spacing: -0.3px;
  }
  .kpi-subtext {
    font-size: 5.7pt;
    color: #475569;
    line-height: 1.12;
  }

  /* Headings */
  .section-title {
    font-size: 8.5pt;
    font-weight: 800;
    color: #095AA5;
    border-left: 3.5px solid #0284c7;
    padding-left: 5px;
    margin: 4px 0 2.5px 0;
    text-transform: uppercase;
    letter-spacing: -0.2px;
  }
  .subsection-title {
    font-size: 7.5pt;
    font-weight: 700;
    color: #0f172a;
    margin: 3.5px 0 1.5px 0;
  }

  p {
    margin-bottom: 3px;
    text-align: justify;
  }
  .highlight-box {
    background: #f0fdf4;
    border-left: 3px solid #16a34a;
    padding: 3.5px 6.5px;
    border-radius: 0 3px 3px 0;
    font-size: 7.1pt;
    color: #14532d;
    margin: 3px 0;
  }
  .alert-box {
    background: #fffbeb;
    border-left: 3px solid #d97706;
    padding: 3.5px 6.5px;
    border-radius: 0 3px 3px 0;
    font-size: 7.1pt;
    color: #78350f;
    margin: 3px 0;
  }
  .blue-box {
    background: #eff6ff;
    border-left: 3px solid #0284c7;
    padding: 3.5px 6.5px;
    border-radius: 0 3px 3px 0;
    font-size: 7.1pt;
    color: #0369a1;
    margin: 3px 0;
  }

  /* Tables */
  table {
    width: 100%;
    border-collapse: collapse;
    margin: 2.5px 0 3.5px 0;
    font-size: 6.5pt;
  }
  th {
    background: #095AA5;
    color: #ffffff;
    font-weight: 700;
    text-align: left;
    padding: 2.2px 3.5px;
    border: 1px solid #095AA5;
  }
  td {
    padding: 1.6px 3.5px;
    border: 1px solid #cbd5e1;
    vertical-align: middle;
  }
  tr:nth-child(even) td {
    background: #f8fafc;
  }

  /* Badges */
  .badge {
    display: inline-block;
    padding: 1px 3.5px;
    border-radius: 2px;
    font-weight: 700;
    font-size: 5.7pt;
    text-align: center;
    color: #ffffff;
  }
  .badge-rn { background: #dc2626; }
  .badge-ok { background: #16a34a; }
  .badge-info { background: #0284c7; }
  .badge-warn { background: #d97706; }

  ul {
    margin-left: 11px;
    margin-bottom: 3px;
  }
  li {
    margin-bottom: 1px;
  }
  code {
    background: #f1f5f9;
    padding: 0.5px 2px;
    border-radius: 2px;
    font-family: Consolas, monospace;
    font-size: 6.3pt;
    color: #0f172a;
  }
  .code-block {
    background: #0f172a;
    color: #e2e8f0;
    font-family: Consolas, monospace;
    font-size: 6.0pt;
    padding: 3.5px 5.5px;
    border-radius: 3px;
    margin: 2.5px 0;
    white-space: pre-wrap;
    line-height: 1.18;
  }
</style>
</head>
<body>

<!-- ==================== PÁGINA 1 ==================== -->
<div class="page">
  <div class="page-content">
    <div class="brand-header">
      <div class="brand-title">CENCOSUD BRASIL</div>
      <div class="brand-subtitle">Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial</div>
      <div class="project-title">Central Executiva de Auditoria de Férias, Cálculo e Folha Mensal (Projeto AD004 - Versão v004)</div>
      <div class="meta-banner">
        <span>Módulo: <strong>Auditoria Integral 360° Férias & FOPAG</strong></span>
        <span>Interface: <strong>Aplicativo Executivo (.exe)</strong></span>
        <span>Competência: <strong>Rollover Dinâmico Automático</strong></span>
        <span>Revisão: <strong>26/10/2026</strong></span>
      </div>
    </div>

    <!-- KPI CARDS ESPECÍFICOS POR AUDITORIA -->
    <div class="kpi-container">
      <div class="kpi-card">
        <div class="kpi-label">Universo Auditado</div>
        <div class="kpi-value">16.000+</div>
        <div class="kpi-subtext">Folha Mensal & Férias Consolidadas</div>
      </div>
      <div class="kpi-card kpi-blue">
        <div class="kpi-label">Cálculo de Férias</div>
        <div class="kpi-value">86 Colunas</div>
        <div class="kpi-subtext">Eventos 501-546, INSS 2026 & Insalubridade</div>
      </div>
      <div class="kpi-card kpi-amber">
        <div class="kpi-label">Folha Mensal</div>
        <div class="kpi-value">65 Colunas</div>
        <div class="kpi-subtext">Competência Dinâmica, FOPAG Completa & Afastados</div>
      </div>
      <div class="kpi-card kpi-emerald">
        <div class="kpi-label">App Executivo</div>
        <div class="kpi-value">3 Abas / GUI</div>
        <div class="kpi-subtext">Checklist Inteligente, FileLockModal & One-File</div>
      </div>
    </div>

    <div class="section-title">1. Visão Geral da Solução e Arquitetura Unificada</div>
    <p>O <strong>Projeto AD004 v004</strong> consolida a evolução da Engenharia de Dados aplicada ao RH do Grupo Cencosud Brasil (GBarbosa, Bretas, Perini, Mercantil Atacado e Eletroshow). A solução foi totalmente empacotada em uma <strong>Central Executiva Desktop Independente (App Executivo .exe)</strong>, permitindo que a equipe de Recursos Humanos execute com 1 clique a conciliação automatizada de todo o ciclo de férias e folha mensal, sem intervenção em linhas de comando e com <strong>adaptação automática a qualquer competência mensal</strong>.</p>

    <div class="highlight-box">
      <strong>Ciclo Completo em 3 Frentes Analíticas Homologadas:</strong>
      <br><strong>(1) Conferência Férias (Formatada - 2 Abas):</strong> Análise prévia de Períodos Aquisitivos (PA), controle legal de dobra (Art. 137 CLT), faltas injustificadas (Art. 130 CLT), vedação de início pré-DSR/feriados geolocalizados por loja e <strong>trava de adiantamento indevido de 13º salário em Nov/Dez/Jan (RN014)</strong>.
      <br><strong>(2) Auditoria de Cálculo de Férias (86 Colunas):</strong> Batimento do terço constitucional, abonos pecuniários e farmácia (545), pensionistas, vendedores/comissionistas, adicionais de insalubridade/periculosidade e simulação do INSS Progressivo 2026 com competência dinâmica.
      <br><strong>(3) Auditoria de Folha Mensal (65 Colunas):</strong> Confronto da soma de proventos salariais contra a Remuneração Fixa Mensal em universo de mais de 15.800 colaboradores (Opção B: <code>=BJ-BE</code>), transição dinâmica entre meses, destaque em vermelho de retorno no mesmo mês e filtro estrito de colaboradores afastados ativos.
    </div>

    <div class="subsection-title">Arquitetura de Software do Aplicativo Executivo Unificado (App Desktop)</div>
    <p>O aplicativo (<code>app_conferencia_ferias.py</code> / <code>Conferencia_Ferias_Cencosud.exe</code>) foi projetado segundo os mais modernos padrões de engenharia de software corporativo:</p>
    <ul>
      <li><strong>Interface Moderna Windows 11 (CustomTkinter):</strong> Tema escuro corporativo com identidade visual oficial Cencosud Blue (<code>#095AA5</code>), estruturado em 3 abas independentes estilo Google Chrome.</li>
      <li><strong>Motor de Varredura e Checklist Preemptivo (<code>file_detector.py</code>):</strong> Analisa a pasta selecionada em tempo real e exibe um checklist com cartões verdes/vermelhos, alertando se faltar algum relatório antes da execução. Na v004, prioriza sempre a base mais recente por <code>mtime</code>, operando perfeitamente na raiz ou em subpastas de meses específicos.</li>
      <li><strong>Execução Assíncrona Não-Bloqueante (Threading):</strong> O processamento pesado roda em threads secundárias dedicadas, garantindo barra de progresso fluida, console de logs em tempo real e interface responsiva sem congelamento.</li>
      <li><strong>Tratamento Amigável de Concorrência (<code>FileLockModal</code>):</strong> Se o arquivo final de destino estiver aberto no Microsoft Excel, o app não aborta com erro; ele abre um modal oferecendo "Tentar Novamente", "Salvar como Cópia" ou "Cancelar".</li>
      <li><strong>Orquestrador de Motores Homologados (<code>audit_engine.py</code>):</strong> Faz a ponte direta e segura com os scripts oficiais de cálculo (<code>gerar_conferencia_ferias_formatado_v4.py</code> e <code>atualizar_todas_auditorias.py</code>).</li>
    </ul>

    <div class="blue-box">
      <strong>Garantia de Fidelidade Absoluta e Rollover Dinâmico:</strong> A versão v004 elimina qualquer mês fixo do código. O motor extrai a competência diretamente dos cabeçalhos dos FOPAs (ex.: 10/2026, 11/2026), gerando relatórios com nomenclatura e colunas autocalibradas para qualquer período do ano com <strong>100,00% de conformidade matemática e estrutural</strong>.
    </div>
  </div>
  <div class="footer">
    <span>Cencosud Brasil • AD004 v004 - Manual Técnico & Auditoria de Férias</span>
    <span>Página 1 de 6</span>
  </div>
</div>

<!-- ==================== PÁGINA 2 ==================== -->
<div class="page">
  <div class="page-content">
    <div class="section-title">2. Estrutura de Pastas e Inteligência de Varredura de Arquivos</div>
    <p>Para conferir total flexibilidade à rotina mensal da equipe de RH, o sistema adota um <strong>Mecanismo Híbrido de Ingestão de Dados</strong>. A equipe pode trabalhar em uma pasta única consolidada de competência (ex.: <code>.../planilhas/conferir inconsistencia</code>) ou na estrutura modular em subpastas:</p>

    <table>
      <thead>
        <tr>
          <th style="width: 28%;">Diretório Sugerido</th>
          <th style="width: 32%;">Conteúdo / Arquivos</th>
          <th style="width: 40%;">Finalidade e Compartilhamento</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td><code>📁 bases compartilhadas</code></td>
          <td><code>COLA_Colaboradores_CSV</code><br><code>Cencosud_Ferias_Colaboradores</code></td>
          <td>Bases mestres corporativas extraídas mensalmente. Fornecem salários contratuais, cargos, sindicatos e histórico de gozo para as 3 auditorias.</td>
        </tr>
        <tr>
          <td><code>📁 relatorios folha mensal</code></td>
          <td><code>FOPA_Eventos (Folha Mensal Completa)</code><br><code>COLA_Colaboradores_Afastados</code></td>
          <td>Extrato master FOPAG (>30 cols com salário e eventos) com ~16.000 colaboradores e base de afastamentos previdenciários ativos.</td>
        </tr>
        <tr>
          <td><code>📁 relatorios calculo ferias</code></td>
          <td><code>FOPA_Eventos (Cálculo Folha 2)</code><br><code>COLA_Beneficiario_Pensionista</code><br><code>Tabelas INSS (xlsb e csv)</code></td>
          <td>Extrato bruto de eventos financeiros de férias, encargos, cadastro de beneficiários judiciais de pensão ativa e tabelas de alíquotas.</td>
        </tr>
        <tr>
          <td><code>📁 relatorios conferencia ferias formatada</code></td>
          <td><code>COLA_Falta_Injustifica_PA</code><br><code>Férias Programadas em CSV</code><br><code>Feriados_Cadastrados</code> / <code>ORGA</code></td>
          <td>Bases cadastrais de suporte ao pré-cálculo: faltas injustificadas no PA (Art. 130 CLT), programação de descanso e calendário de feriados locais.</td>
        </tr>
        <tr>
          <td><code>📁 pastas de fechamento mensal</code></td>
          <td>Todas as bases da competência agrupadas (ex.: <code>conferir inconsistencia</code>)</td>
          <td>O <code>FileDetector v004</code> reconhece bases consolidadas na mesma pasta, associando automaticamente cada relatório à respectiva aba.</td>
        </tr>
      </tbody>
    </table>

    <div class="subsection-title">2.1. Inventário Detalhado dos Relatórios Consumidos pelas Auditorias</div>
    <p>O quadro abaixo lista os relatórios corporativos processados e as validações estritas introduzidas na v004:</p>

    <table>
      <thead>
        <tr>
          <th style="width: 33%;">Nome do Arquivo / Relatório</th>
          <th style="width: 14%;">Formato / Peso</th>
          <th style="width: 18%;">Auditorias Vinculadas</th>
          <th style="width: 35%;">Finalidade e Regras Associadas (v004)</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td><code>COLA_Colaboradores_CSV - *.CSV</code></td>
          <td>CSV (~25MB)</td>
          <td>Cálculo e Folha</td>
          <td>Salário Contratual, Cargo, Sindicato, Admissão e elegibilidade 545.</td>
        </tr>
        <tr>
          <td><code>Cencosud_Ferias_Colaboradores_CSV*.CSV</code></td>
          <td>CSV (~300KB a 2.2MB)</td>
          <td>Folha e Formatada</td>
          <td>Histórico de períodos de gozo, frações, intervalo e opção de 13º salário.</td>
        </tr>
        <tr>
          <td><code>FOPA_Eventos_por_Colaborador (Folha Completa).CSV</code></td>
          <td>CSV (~7.5MB / 58 cols)</td>
          <td>Folha Mensal</td>
          <td><strong>Blindagem v004:</strong> Exige > 30 colunas e presença de salário, impedindo o uso de extratos parciais de insalubridade.</td>
        </tr>
        <tr>
          <td><code>COLA_Colaboradores_Afastados - *.CSV</code></td>
          <td>CSV (~680KB)</td>
          <td>Folha Mensal</td>
          <td><strong>Filtro v004:</strong> Prioriza e filtra estritamente <code>Status == 'Ativo'</code>, descartando bases históricas de cessados (evita 7.500+ alarmes falsos).</td>
        </tr>
        <tr>
          <td><code>FOPA_Eventos_por_Colaborador (Cálculo).CSV</code></td>
          <td>CSV (~450KB a 630KB)</td>
          <td>Cálculo Férias</td>
          <td>Eventos de férias (501 a 546, 5501 a 5505, bases de cálculo e dias na Folha 2).</td>
        </tr>
        <tr>
          <td><code>COLA_Beneficiario_Pensionista - *.CSV</code></td>
          <td>CSV (~80KB)</td>
          <td>Cálculo Férias</td>
          <td>Cadastro de pensionistas ativos para batimento do 5503/4/5.</td>
        </tr>
        <tr>
          <td><code>Confere INSS Mês 012026 / inss-empresa.csv</code></td>
          <td>XLSB / CSV</td>
          <td>Cálculo Férias</td>
          <td>Tabelas oficiais de faixas e teto previdenciário do INSS 2026.</td>
        </tr>
        <tr>
          <td><code>COLA_Falta_Injustifica_PA - *.CSV</code></td>
          <td>CSV (~120KB)</td>
          <td>Formatada</td>
          <td>Faltas injustificadas no PA para dedução de escala (Art. 130 CLT).</td>
        </tr>
        <tr>
          <td><code>Férias Programadas ou Gozadas em CSV</code></td>
          <td>CSV (~400KB cada)</td>
          <td>Formatada e Folha</td>
          <td>Múltiplos arquivos mensais; mapeamento automático por mês predominante.</td>
        </tr>
        <tr>
          <td><code>Feriados_Cadastrados / ORGA_Estabelecimento</code></td>
          <td>CSV / XLSX</td>
          <td>Formatada</td>
          <td>Calendário de feriados nacionais e municipais por CNPJ/Município.</td>
        </tr>
      </tbody>
    </table>

    <div class="highlight-box">
      <strong>Varredura Recursiva Inteligente (Zero Duplicidade e Livre de Travas):</strong>
      Na versão v004, o <code>FileDetector</code> realiza a varredura priorizando arquivos pela data de modificação real (<code>mtime</code>), eliminando filtros rígidos de subpastas que forçavam a leitura de relatórios antigos de meses passados. Bases compartilhadas são reaproveitadas automaticamente entre os motores.
    </div>
  </div>
  <div class="footer">
    <span>Cencosud Brasil • AD004 v004 - Manual Técnico & Auditoria de Férias</span>
    <span>Página 2 de 6</span>
  </div>
</div>

<!-- ==================== PÁGINA 3 ==================== -->
<div class="page">
  <div class="page-content">
    <div class="section-title">3. Matriz de Regras de Negócio e Validações Matemáticas (RNs) - Parte 1</div>
    <p>Cada linha processada nos motores analíticos é submetida a um conjunto estrito de validações matemáticas e de conformidade legal, classificadas sob a matriz de Regras de Negócio corporativas:</p>

    <table>
      <thead>
        <tr>
          <th style="width: 8%;">Código</th>
          <th style="width: 20%;">Tipo de Validação</th>
          <th style="width: 36%;">Regra Aplicada & Fórmula Homologada</th>
          <th style="width: 18%;">Base Origem</th>
          <th style="width: 18%;">Ação Recomendada</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td><span class="badge badge-rn">RN001</span></td>
          <td><strong>Terço Constitucional de Férias</strong></td>
          <td><code>Conf 502 = Evento 501 / 3</code><br>Validação: <code>Validação 502 = Conf 502 - Evento 502</code> (Repete para 504, 506, 510, 514). Tolerância centesimal de arredondamento &le; 0.02.</td>
          <td>FOPAG Cálculo</td>
          <td>Corrigir recálculo de terço quando diferença &ne; 0.</td>
        </tr>
        <tr>
          <td><span class="badge badge-rn">RN002</span></td>
          <td><strong>INSS Progressivo Oficial 2026</strong></td>
          <td>Simulação algorítmica progressiva oficial com alíquotas de 7.5% a 14% sobre Evento 10120 com dedução de parcela e teto de R$ 988,09.<br><code>Validação 5501 = BM - BO</code>.</td>
          <td>FOPAG Cálculo / Tabela INSS 2026</td>
          <td>Ajustar retenção de INSS na FOPAG de férias (0 divergências registradas).</td>
        </tr>
        <tr>
          <td><span class="badge badge-rn">RN003</span></td>
          <td><strong>Proporcionalidade Férias Fracionadas</strong></td>
          <td><code>Validação Fracionadas = (AB + AP) - Y</code><br>Soma do evento 501 (AB) com 509 (AP) confrontada com a proporção Y. Em branco para 30 dias e vendedores.</td>
          <td>FOPAG Cálculo / Faltas PA</td>
          <td>Revisar saldo de dias debitados no PA em períodos fracionados.</td>
        </tr>
        <tr>
          <td><span class="badge badge-rn">RN004</span></td>
          <td><strong>Férias em Dobro (Art. 137 CLT)</strong></td>
          <td><code>Fim do Gozo &le; Data Limite Concessivo</code><br>Se gozo ultrapassar o limite concessivo, apurar obrigatoriedade do pagamento dobrado (Evento 522).</td>
          <td>Férias Programadas / PA Aberto</td>
          <td>Implantar pagamento em dobro por concessão tardia.</td>
        </tr>
        <tr>
          <td><span class="badge badge-rn">RN005</span></td>
          <td><strong>Abono Férias Farmácia (Evento 545)</strong></td>
          <td>Se Cargo contiver <em>FARMACEUTICO</em> E Sindicato = <em>1</em> E Tempo de Casa &ge; 15 anos:<br><code>Confere 545 (BI) = (Salário Contratual / 30) * 10</code>.<br>Validação: <code>Validação 545 (BJ) = BI - BK</code>.</td>
          <td>COLA_Colaboradores / FOPAG Cálculo</td>
          <td>Conceder abono de 10 dias de salário aos farmacêuticos elegíveis.</td>
        </tr>
        <tr>
          <td><span class="badge badge-rn">RN006</span></td>
          <td><strong>Pensão Alimentícia em Férias</strong></td>
          <td>Confronta o cadastro de pensionistas ativos contra deduções judiciais nos Eventos 5503, 5504 e 5505. Exibe <code>OK</code> ou <code>Verificar</code> na Coluna BQ.</td>
          <td>COLA_Beneficiario_Pensionista</td>
          <td>Validar percentual judicial sobre a base líquida de proventos.</td>
        </tr>
      </tbody>
    </table>

    <div class="subsection-title">Detalhe do Motor Previdenciário Oficial INSS 2026 (RN002)</div>
    <p>O algoritmo homologado aplica rigorosamente as faixas oficiais vigentes e o teto previdenciário, sem qualquer desvio em relação ao sistema de folha da Cencosud:</p>
    <div class="code-block">
def calc_inss_2026_official(base):
    if not base or base <= 0: return 0.0
    if base > 8475.55: return 988.0914 # Teto Previdenciário Oficial 2026
    if base <= 1621.00: return base * 0.075
    elif base <= 2902.84: return base * 0.09 - 24.3150
    elif base <= 4354.27: return base * 0.12 - 111.4002
    else: return base * 0.14 - 198.4856
    </div>

    <div class="alert-box">
      <strong>Tolerância Centesimal Homologada:</strong> Divergências inferiores ou iguais a R$ 0,02 são tratadas como arredondamento computacional legal, mantendo o status em conformidade (<code>0,00</code>) e evitando alarmes falsos para os analistas de folha.
    </div>
  </div>
  <div class="footer">
    <span>Cencosud Brasil • AD004 v004 - Manual Técnico & Auditoria de Férias</span>
    <span>Página 3 de 6</span>
  </div>
</div>

<!-- ==================== PÁGINA 4 ==================== -->
<div class="page">
  <div class="page-content">
    <div class="section-title">3. Matriz de Regras de Negócio e Validações Matemáticas (RNs) - Parte 2</div>
    <table>
      <thead>
        <tr>
          <th style="width: 8%;">Código</th>
          <th style="width: 20%;">Tipo de Validação</th>
          <th style="width: 36%;">Regra Aplicada & Fórmula Homologada</th>
          <th style="width: 18%;">Base Origem</th>
          <th style="width: 18%;">Ação Recomendada</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td><span class="badge badge-rn">RN007</span></td>
          <td><strong>Batimento Folha Mensal (Opção B)</strong></td>
          <td><code>TT de Proventos (BJ) = SUM(R:AS)</code> (Soma proventos salariais)<br><code>Validação (BL) = BJ - BE</code> (Diferença numérica direta).<br>Coluna <code>Obs.:</code> alerta <code>Verificar</code> quando diferença &ne; 0.</td>
          <td>FOPAG Folha Mensal</td>
          <td>Auditar diferenças salariais e apontar justificativas corporativas.</td>
        </tr>
        <tr>
          <td><span class="badge badge-rn">RN008</span></td>
          <td><strong>Conferência de Vendedores / Comissionistas</strong></td>
          <td>Folha: <code>Confere vendedor (BK) = SUM(R:AR)</code>. Se Validação > 0.02, marca <code>Vendedor; Diferença de valor Positiva - Conferir</code>.<br>Cálculo: <code>Validar Vendedor II (AA) = (Z=T)</code> validando médias de comissões.</td>
          <td>COLA_Colaboradores / FOPAG Folha e Cálculo</td>
          <td>Garantir batimento integral das médias de comissões variáveis.</td>
        </tr>
        <tr>
          <td><span class="badge badge-rn">RN009</span></td>
          <td><strong>Fracionadas & Retorno no Mês (Folha)</strong></td>
          <td>Carga dinâmica de múltiplos meses consecutivos. Intervalo na Col Q. Se <em>"Retorno e novo gozo no mesmo mês"</em>, aplica <strong>Destaque em Vermelho de Alerta</strong> (<code>FFC7CE</code>) nas Colunas O, P e Q.</td>
          <td>Cencosud_Ferias_Colaboradores / Programadas</td>
          <td>Assegurar a correta quitação dos dias contratuais entre meses.</td>
        </tr>
        <tr>
          <td><span class="badge badge-rn">RN010</span></td>
          <td><strong>Vedação de Início Pré-DSR e Feriados Locais</strong></td>
          <td>Art. 134, § 3º da CLT: <code>Início Gozo &minus; DSR/Feriado &gt; 2 dias</code>.<br><strong>Geolocalização Loja a Loja:</strong> Cruzamento com <code>ORGA_Estabelecimento</code> elimina falsos feriados municipais de outras cidades.</td>
          <td>Férias Programadas / Feriados / ORGA</td>
          <td>Reprogramar a data de início do descanso para dia útil válido.</td>
        </tr>
        <tr>
          <td><span class="badge badge-rn">RN011</span></td>
          <td><strong>Totalizador de Dias de Férias</strong></td>
          <td>Soma dinâmica dos eventos de dias de gozo e abonos:<br><code>TT dias Férias (CF) = SUM(BX:BY)</code> com exibição em inteiros limpos.</td>
          <td>FOPAG Cálculo</td>
          <td>Conferir quitação integral dos dias devidos no PA.</td>
        </tr>
        <tr>
          <td><span class="badge badge-rn">RN012</span></td>
          <td><strong>Validação de Insalubridade e Periculosidade</strong></td>
          <td>Remuneração Fixa (Col U) contempla adicionais. Se <code>Col U > Col V</code> (Salário Contratual), considera <strong>OK</strong>; se <code>Col U &le; Col V</code>, aponta <em>"Divergência em Validação Sal Insalub/Pericul"</em>. Coluna 20 grava <code>round(v_u - v_v, 2)</code>.</td>
          <td>FOPAG Cálculo / Cadastro</td>
          <td>Validar integração dos adicionais de insalubridade/periculosidade.</td>
        </tr>
        <tr>
          <td><span class="badge badge-rn">RN013</span></td>
          <td><strong>Resgate Dinâmico do Período Aquisitivo (PA)</strong></td>
          <td>Preenchimento da Coluna Q na Auditoria de Cálculo via varredura na base de Férias Programadas/Gozadas, eliminando períodos aquisitivos em branco.</td>
          <td>Férias Programadas CSV</td>
          <td>Garantir rastreabilidade integral do ciclo aquisitivo do colaborador.</td>
        </tr>
        <tr>
          <td><span class="badge badge-rn">RN014</span></td>
          <td><strong>Trava de Adiantamento 13º Salário (Aba 2)</strong></td>
          <td><strong>Regra v004:</strong> Se <code>1ª Parc. 13º Salário == 'Sim'</code> e o início do gozo for em <strong>Novembro (11), Dezembro (12) ou Janeiro (1)</strong>, sinaliza na Coluna <em>Ação</em>: <strong>"Não permitido adiantamento de 13º em Nov/Dez/Jan"</strong> (Alerta Amarelo).</td>
          <td>Cencosud_Ferias_Colaboradores</td>
          <td>Remover opção de adiantamento de 13º em períodos com quitação compulsória ou virada de ano.</td>
        </tr>
      </tbody>
    </table>

    <div class="highlight-box">
      <strong>Novidades Homologadas na Versão v004:</strong>
      <br>• <strong>RN009 (Destaque Vermelho):</strong> Quando há retorno e novo gozo no mesmo mês civil, as <strong>Colunas O, P e Q</strong> recebem preenchimento vermelho suave (<code>#FFC7CE</code>) com texto vinho (<code>#9C0006</code>).
      <br>• <strong>RN014 (Trava de 13º Salário):</strong> Bloqueio preventivo na Aba 2 de férias para solicitações indevidas da 1ª parcela de 13º em meses nos quais a empresa já efetua o pagamento integral regulamentar.
    </div>

    <div class="blue-box">
      <strong>Competência Dinâmica e Evento 83:</strong>
      O motor da Folha identifica dinamicamente o mês anterior e o mês corrente através do FOPAG, nomeando as colunas automaticamente (ex.: <code>Férias Setembro</code> e <code>Férias Outubro</code>) e adaptando os cálculos de transição entre recibos.
    </div>
  </div>
  <div class="footer">
    <span>Cencosud Brasil • AD004 v004 - Manual Técnico & Auditoria de Férias</span>
    <span>Página 4 de 6</span>
  </div>
</div>

<!-- ==================== PÁGINA 5 ==================== -->
<div class="page">
  <div class="page-content">
    <div class="section-title">4. Mecanismos Avançados de Engenharia e Blindagem de Dados</div>
    <p>Para assegurar confiabilidade de nível corporativo e eliminar vícios estruturais de relatórios legados, a versão <strong>v004</strong> incorporou proteções algorítmicas de missão crítica:</p>

    <div class="subsection-title">A. O Parser Numérico Resiliente e Eliminação do Bug de Strings no SUM()</div>
    <p>Nos relatórios brutos emitidos pelo sistema LG, valores &ge; R$ 1.000,00 possuem ponto de milhar e vírgula decimal (ex: <code>1.560,00</code>). O parser de dupla checagem sanitiza 100% das rubricas, convertendo-as para <code>float</code> puro antes da injeção nas matrizes, impedindo que o Excel ignore valores dentro da função <code>=SUM()</code>.</p>

    <div class="subsection-title">B. Blindagem Contra Notação Científica e Resíduos Flutuantes IEEE-754</div>
    <p>Operações aritméticas em memória podem gerar resíduos infinitesimais como <code>-2,27374E-13</code>. O robô aplica a rotina <code>safe_round(val)</code>, garantindo que qualquer resíduo inferior a R$ 0,0001 seja truncado para zero absoluto (<code>0.0</code>).</p>

    <div class="subsection-title">C. Máscara Contábil Explícita e Aparição do "0,00"</div>
    <p>Padronizada a máscara <code>#,##0.00;-#,##0.00;0.00</code>. Casos regulares exibem explicitamente <strong>0,00</strong>, confirmando ao auditor que a linha foi validada com sucesso.</p>

    <div class="subsection-title">D. Validador Estrutural de Folha Completa (Blindagem contra Extratos Parciais)</div>
    <p>O <code>file_detector.py</code> na v004 implementa o validador <code>is_csv_fopa_folha_completa</code>, que exige mais de 30 colunas e presença de rubricas salariais. Isso impede que extratos parciais de adicionais (17 colunas e 963 linhas) sejam selecionados por engano como folha principal, evitando o sumiço de colaboradores ou colunas em branco.</p>

    <div class="subsection-title">E. Filtro Rigoroso de Colaboradores Afastados Ativos</div>
    <p>Com a rotina <code>is_active_afastados</code> e filtro em tempo de execução <code>Status == 'Ativo'</code>, o motor descarta automaticamente bases históricas de ex-colaboradores (cessados), mantendo a taxa real de afastamentos da empresa em ~9,8% e eliminando mais de 7.500 alarmes falsos de afastamento na Folha.</p>

    <div class="subsection-title">F. Resiliência de Concorrência e FileLockModal (Controle de Travamento do Excel)</div>
    <p>Se no momento da gravação a planilha final estiver aberta no Excel, o <code>FileLockModal</code> oferece "Tentar Novamente", "Salvar como Cópia Datada" ou "Cancelar", sem abortar o processamento.</p>

    <div class="section-title">5. Exemplo Prático de Diagnóstico Analítico (Cenários Didáticos Simulados)</div>
    <p style="font-size: 7.0pt; color: #475569; font-style: italic; margin-bottom: 2px;">Nota de Conformidade (LGPD): Registros <strong>100% fictícios, simulados e anonimizados</strong> para demonstração do comportamento do motor:</p>
    <div class="code-block" style="background: #1e293b; font-size: 5.9pt;">
[MATRÍCULA] | [COLABORADOR (FICTÍCIO)] | [REMUN. FIXA] | [TT PROVENTOS] | [VALIDAÇÃO] | [DIAGNÓSTICO ANALÍTICO DO MOTOR (v004)]
----------------------------------------------------------------------------------------------------------------------------
100001      | COLABORADOR MODELO A    | R$ 6.500,00   | R$ 6.500,00    | R$ 0,00     | REGULAR - Salário (5.500,00) + Triênio (1.000,00) OK
100002      | COLABORADOR MODELO B    | R$ 2.500,00   | R$ 500,00      | -R$ 2.000,00| AFASTAMENTO INSS ATIVO - Licença Médica Válida (Status Ativo)
100003      | COLABORADOR MODELO C    | R$ 3.000,00   | R$ 4.000,00    | R$ 0,00     | CÁLCULO FÉRIAS OK - Férias Mês (3.000,00) + 1/3 Constitucional OK
100004      | COLABORADOR MODELO D    | R$ 12.000,00  | R$ 8.000,00    | R$ 0,00     | FRACIONADO 15 DIAS - 501 Proporcional + 1/3 OK
100005      | COLABORADOR MODELO E    | R$ 2.200,00   | R$ 2.200,00    | R$ 0,00     | VENDEDOR / COMISSIONISTA - Validação (Z=T) OK
100006      | COLABORADOR MODELO F    | R$ 3.500,00   | R$ 3.500,00    | R$ 0,00     | RETORNO E NOVO GOZO NO MESMO MÊS (DESTAQUE VERMELHO)
100007      | COLABORADOR MODELO G    | R$ 4.200,00   | R$ 4.200,00    | R$ 0,00     | ABA 2: "NÃO PERMITIDO ADIANTAMENTO DE 13º EM NOV/DEZ/JAN" (AMARELO)
    </div>
  </div>
  <div class="footer">
    <span>Cencosud Brasil • AD004 v004 - Manual Técnico & Auditoria de Férias</span>
    <span>Página 5 de 6</span>
  </div>
</div>

<!-- ==================== PÁGINA 6 ==================== -->
<div class="page">
  <div class="page-content">
    <div class="section-title">6. Guia Operacional do Aplicativo Executivo Unificado (Passo a Passo)</div>
    <p>A operação mensal foi simplificada para permitir que qualquer colaborador da equipe de RH gere as auditorias com facilidade e autonomia:</p>

    <div class="subsection-title">6.1. Roteiro de Uso em 4 Passos</div>
    <ul>
      <li><strong>Passo 1 (Abrir o App e Definir Pastas):</strong> Abrir o executável <code>Conferencia_Ferias_Cencosud.exe</code>. Os caminhos da <em>Pasta Raiz das Planilhas</em> e <em>Destino das Conferências</em> já vêm preenchidos com o padrão corporativo. O motor v004 localiza automaticamente as bases mais recentes da competência mesmo que estejam na raiz ou em pastas de conferência.</li>
      <li><strong>Passo 2 (Selecionar a Aba e Conferir o Checklist):</strong> Clique na aba da conferência desejada:
        <br>• <em>Aba 1:</em> Conferência Férias (Formatada) | • <em>Aba 2:</em> Cálculo de Férias | • <em>Aba 3:</em> Folha Mensal.
        <br>Observe a coluna esquerda: itens com <strong>check verde</strong> indicam relatórios validados. Se algum item estiver em <strong>vermelho</strong>, adicione o arquivo faltante na pasta e clique em <em>"🔄 Escanear / Atualizar Arquivos"</em>.
      </li>
      <li><strong>Passo 3 (Disparar a Geração):</strong> Clique no botão azul <strong>"▶ Gerar Auditoria..."</strong>. Acompanhe a barra de progresso e o log em tempo real. O processamento integral leva poucos segundos.</li>
      <li><strong>Passo 4 (Acessar a Planilha Final):</strong> Ao concluir com sucesso, clique em <strong>"📄 Abrir Planilha Gerada"</strong> para abrir diretamente no Excel ou em <strong>"📁 Abrir Pasta de Destino"</strong>.</li>
    </ul>

    <div class="subsection-title">6.2. Estrutura dos Relatórios Gerados na Pasta de Conferências</div>
    <table>
      <thead>
        <tr>
          <th style="width: 33%;">Arquivo XLSX Gerado (Dinâmico)</th>
          <th style="width: 14%;">Competência</th>
          <th style="width: 17%;">Dimensão</th>
          <th style="width: 36%;">Como Analisar Divergências no Excel</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td><code>Auditoria_Folha_Mensal_<MesAno>_v6.xlsx</code><br><span style="font-size: 5.7pt; color: #64748b;">(ex.: Auditoria_Folha_Mensal_Out2026_v6.xlsx)</span></td>
          <td>Mês Atual da Folha</td>
          <td>~16.000 linhas &times; 65 colunas</td>
          <td>Filtrar a coluna <code>Obs.:</code> por <em>"Verificar"</em> ou <em>"Diferença"</em>. Conferir as linhas destacadas em vermelho nas colunas de férias O, P e Q.</td>
        </tr>
        <tr>
          <td><code>Auditoria_Calculo_Ferias_<MesAno>_v6.xlsx</code><br><span style="font-size: 5.7pt; color: #64748b;">(ex.: Auditoria_Calculo_Ferias_Nov2026_v6.xlsx)</span></td>
          <td>Mês do Cálculo</td>
          <td>Total FOPAG 2 &times; 86 colunas</td>
          <td>Filtrar colunas <code>Validação 501, 5501, 502, 545</code> e <code>Validação Sal Insalub/Pericul</code> por valores &ne; 0,00.</td>
        </tr>
        <tr>
          <td><code>Automacao_Conferencia_Ferias_Formatada.xlsx</code></td>
          <td>Programação Férias</td>
          <td>Abas PA Aberto & Férias</td>
          <td>Na Aba 1 (PA Aberto): verificar dobro legal e feriados. Na Aba 2 (Férias): verificar alertas amarelos de <em>"Não permitido adiantamento de 13º em Nov/Dez/Jan"</em>.</td>
        </tr>
      </tbody>
    </table>

    <div class="section-title">7. Conclusão da Homologação Técnica e Institucional</div>
    <div class="highlight-box">
      <strong>Parecer Técnico de Liberação Operacional:</strong><br>
      A <strong>Central Executiva de Auditoria de Férias e Folha (Projeto AD004 - Versão v004)</strong> foi integralmente homologada pela Administração de Pessoal e Engenharia de Dados da Cencosud Brasil. A solução garante 100% de acurácia matemática, resiliência operacional contra falhas de concorrência, suporte dinâmico a todas as competências mensais do ano, proteção contra extratos parciais e facilidade de uso via executável único. <strong>Solução homologada e liberada para operação definitiva.</strong>
    </div>

    <div style="margin-top: 8px; border-top: 1px solid #cbd5e1; padding-top: 7px; display: grid; grid-template-columns: 1fr 1fr; gap: 20px; text-align: center;">
      <div>
        <div style="font-weight: 700; font-size: 7.5pt; color: #095AA5;">Engenharia de Dados & Automação de RH</div>
        <div style="font-size: 6.5pt; color: #64748b; margin-top: 1px;">Desenvolvimento de Software, Algoritmos & App Desktop (v004)</div>
      </div>
      <div>
        <div style="font-weight: 700; font-size: 7.5pt; color: #095AA5;">Administração de Pessoal & Remuneração</div>
        <div style="font-size: 6.5pt; color: #64748b; margin-top: 1px;">Homologação de Negócio, Regras CLT & Conformidade Trabalhista</div>
      </div>
    </div>
  </div>
  <div class="footer">
    <span>Cencosud Brasil • AD004 v004 - Manual Técnico & Auditoria de Férias</span>
    <span>Página 6 de 6</span>
  </div>
</div>

</body>
</html>
"""

with open(html_path, "w", encoding="utf-8") as f:
    f.write(html_content)

if html_root_path != html_path:
    with open(html_root_path, "w", encoding="utf-8") as f:
        f.write(html_content)

print(f"HTML gerado com sucesso: {html_path}")

chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"
if not os.path.exists(chrome_path):
    chrome_path = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"

file_url = "file:///" + html_path.replace("\\", "/")

import tempfile
import shutil

temp_profile = os.path.join(tempfile.gettempdir(), "cencosud_chrome_pdf_profile")
temp_html = os.path.join(tempfile.gettempdir(), "cencosud_manual_doc.html")
temp_pdf = os.path.join(tempfile.gettempdir(), "cencosud_manual_doc.pdf")

with open(temp_html, "w", encoding="utf-8") as f:
    f.write(html_content)

if os.path.exists(temp_pdf):
    try: os.remove(temp_pdf)
    except: pass

file_url = "file:///" + temp_html.replace("\\", "/")

cmd = [
    chrome_path,
    "--headless",
    "--disable-gpu",
    "--no-pdf-header-footer",
    f"--user-data-dir={temp_profile}",
    f"--print-to-pdf={temp_pdf}",
    file_url
]

print("Convertendo para PDF via Chrome Headless...")
subprocess.run(cmd, check=True)

if os.path.exists(temp_pdf):
    shutil.copyfile(temp_pdf, pdf_path)
    shutil.copyfile(temp_pdf, pdf_doc_path)
    
    doc = pymupdf.open(pdf_path)
    print(f"PDF gerado com sucesso: {pdf_path}")
    print(f"PDF copiado para documentacao: {pdf_doc_path}")
    print(f"Total de paginas no PDF: {len(doc)}")
else:
    print("Erro: PDF nao foi gerado.")
