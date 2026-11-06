"""
Registro Central de Módulos e Feature Flags - Central de Benefícios Cencosud Brasil.

Permite habilitar/desabilitar cada conferência individualmente (Feature Flags)
sem necessidade de recriar a interface gráfica à medida que as validações
da equipe de Benefícios avançam.
"""

from typing import Dict, List, Any

MODULES_CONFIG: List[Dict[str, Any]] = [
    # =========================================================================
    # 🟢 MÓDULOS HOMOLOGADOS E ATIVOS
    # =========================================================================
    {
        "id": "good_card",
        "code": "01",
        "short_name": "Good Card (Evt 6280)",
        "full_title": "Auditoria 04 - Good Card (Saldos, Margem 20% e Desconto 6280)",
        "enabled": True,
        "status_badge": "🟢 ATIVO",
        "status_label": "Módulo Homologado e Ativo para Uso em Produção",
        "default_subfolder": r"04 - GOOD CARD\Relatorio Original",
        "default_out_subfolder": r"04 - GOOD CARD\Relatório Tratado",
        "output_filename": "Conferência Good Card_Gerado.xlsx",
        "description": (
            "Realiza o cruzamento integral da Folha (FOPA), Cadastro Ativo (Nomina), Afastados, "
            "Demitidos, Férias e Valores Importados (Evento 29613), validando o teto da margem "
            "consignável de 20%, liquidação de saldo devedor (6280), transferência de resíduo (29621) "
            "e subsídio empresa 5% (29624 - Giga / Estab. 400)."
        ),
        "events": ["6280 - Desc Cartão Benefícios", "29613 - Valor Informado", "29621 - Saldo Devedor", "29622 - Saldo Mês Ant.", "29624 - Custo Empresa 5%"],
        "required_files": [
            "FOPA_Eventos_por_Colaborador.CSV",
            "COLA_Colaboradores.CSV",
            "COLA_Colaboradores_Afastados.CSV",
            "COLA_Demitidos.CSV",
            "COLA_Ferias_Colaboradores.CSV",
            "Valores Importados (Conferência Good Card_Base.xlsx / CSV)",
        ],
    },
    {
        "id": "compra_mensal_vt",
        "code": "02",
        "short_name": "Compra Mensal VT",
        "full_title": "Auditoria de Compra Mensal de Vale Transporte",
        "enabled": True,
        "status_badge": "🟢 ATIVO",
        "status_label": "Módulo Homologado e Ativo para Uso em Produção",
        "default_subfolder": r"Compra Mensal VT\Relatórios Originais",
        "default_out_subfolder": r"Compra Mensal VT\Relatório Tratado",
        "output_filename": "Conferência Compra Mensal VT_10.2026.xlsx",
        "description": (
            "Valida a elegibilidade e os valores de recarga mensal de Vale Transporte por colaborador, "
            "cruzando dias úteis, tarifas cadastradas, férias, afastamentos ativos/cessados e desligamentos."
        ),
        "events": ["Recarga Regular Operadoras", "Recarga Complementar", "31903 - Custo Total VT"],
        "required_files": [
            "RELAOD (Compra Regular VT / CSV)",
            "VLTR_P (Base de Optantes VT / CSV)",
            "COLA_C3 (Censo Funcional Ativo / CSV)",
            "COLA_C1 (Afastados Ativos / CSV)",
            "COLA_F (Programação de Férias / CSV)",
            "COLA_D (Desligados / CSV)",
        ],
    },
    {
        "id": "valores_pre_folha",
        "code": "03",
        "short_name": "Valores Pré-Folha",
        "full_title": "Auditoria de Importação de Valores Pré-Folha VT",
        "enabled": True,
        "status_badge": "🟢 ATIVO",
        "status_label": "Módulo Homologado e Ativo para Uso em Produção",
        "default_subfolder": r"Valores Pré-Folha",
        "default_out_subfolder": r"Valores Pré-Folha\Relatório Tratado",
        "output_filename": "Conferência Valores Pré-Folha VT_09.2026.xlsx",
        "description": (
            "Audita a soma de 10 eventos de pagamento contra o Evento 5750 (Adiantamento VT Através Alçadas) "
            "e valida os lotes de importação dos eventos 31935 (VT Alternativo) e 6292 (Desconto VTE)."
        ),
        "events": ["5750 - Adiantamento VT Alçadas", "31935 - VT Alternativo Mês Seg.", "6292 - Desconto Cartão VTE"],
        "required_files": [
            "Conferência 09.2026.xlsx (Abas: FOPA_Eventos_por_Colaborador, 31935 e 6292)",
        ],
    },
    {
        "id": "bloqueio_good_card",
        "code": "04",
        "short_name": "Bloqueio Good Card",
        "full_title": "Auditoria de Bloqueio Good Card (Afastados > 15 Dias)",
        "enabled": True,
        "status_badge": "🟢 ATIVO",
        "status_label": "Módulo Homologado e Ativo para Uso em Produção",
        "default_subfolder": r"Conferência Bloqueio Good Card",
        "default_out_subfolder": r"Conferência Bloqueio Good Card\Relatório Tratado",
        "output_filename": "Conferência Bloqueio Good Card_09.2026.xlsx",
        "description": (
            "Audita o status de cartões Good Card para colaboradores afastados há mais de 15 dias, "
            "identificando colaboradores com cartão ativo que exigem bloqueio preventivo na operadora."
        ),
        "events": ["Cartões Ativos Ticket Good Card", "Limiar de 15 Dias de Afastamento"],
        "required_files": [
            "Bloqueio Afatamento Cencosud_Atacado_Perini.xlsx (Bandeiras + Geral + Afastados)",
        ],
    },

    # =========================================================================
    # 🟡 MÓDULOS EM VALIDAÇÃO / HOMOLOGAÇÃO PELA EQUIPE
    # =========================================================================
    {
        "id": "vt_6pct",
        "code": "05",
        "short_name": "Vale Transporte 6%",
        "full_title": "Auditoria de Desconto Legal 6% Vale Transporte (Evento 6290)",
        "enabled": False,
        "status_badge": "🟡 EM VALIDAÇÃO",
        "status_label": "Planilha com 11 comentários atendidos enviada — Aguardando retorno da equipe",
        "default_subfolder": r"6%\Relatórios Originais",
        "default_out_subfolder": r"6%",
        "output_filename": "Conferência 6% _Auditada.xlsx",
        "description": (
            "Audita o desconto legal de 6% de Vale Transporte (Evento 6290) frente ao Salário Mensal (11010), "
            "Soma Base (1 + 137 + 138), Custo Total de VT (31903 + 31905) e Custo Empresa (31906)."
        ),
        "events": ["6290 - Vale Transporte", "31903 - Custo Total VT", "31905 - Aux Crédito VT", "31906 - Custo Empresa VT"],
        "required_files": ["FOPA_Eventos_por_Colaborador.CSV", "COLA_Colaboradores.CSV", "Relação Compra VT.CSV"],
    },
    {
        "id": "devolucao_valores",
        "code": "06",
        "short_name": "Devolução de Valores",
        "full_title": "Auditoria de Devolução e Reembolso de Benefícios",
        "enabled": False,
        "status_badge": "🟡 EM VALIDAÇÃO",
        "status_label": "Módulo construído — Em fase de validação pela equipe de Benefícios",
        "default_subfolder": r"Devolução de Valores\Relatórios Originais",
        "default_out_subfolder": r"Devolução de Valores",
        "output_filename": "Conferência_Devolucao_Valores_Auditada.xlsx",
        "description": (
            "Concilia as solicitações de devolução e reembolso de benefícios com os lançamentos "
            "efetivados em folha de pagamento e histórico de descontos anteriores."
        ),
        "events": ["Eventos de Devolução / Reembolso VT e VA"],
        "required_files": ["Planilha de Solicitações de Devolução", "FOPA_Eventos_por_Colaborador.CSV"],
    },
    {
        "id": "refeitorio",
        "code": "07",
        "short_name": "Conferência Refeitório",
        "full_title": "Auditoria de Desconto de Refeitório e Alimentação",
        "enabled": False,
        "status_badge": "🟡 EM VALIDAÇÃO",
        "status_label": "Módulo construído — Em fase de validação pela equipe de Benefícios",
        "default_subfolder": r"Conferencia Refeitorio\Relatórios Originais",
        "default_out_subfolder": r"Conferencia Refeitorio",
        "output_filename": "Conferência Refeitório_Gerado.xlsx",
        "description": (
            "Audita os descontos de refeitório em folha frente às faixas salariais, tetos sindicais, "
            "proporcionalidade de férias, afastamentos e admissões no mês."
        ),
        "events": ["Desconto Refeitório / Alimentação em Folha"],
        "required_files": ["FOPA_Eventos_por_Colaborador.CSV", "COLA_Colaboradores.CSV"],
    },
    {
        "id": "rateio_cve",
        "code": "08",
        "short_name": "Rateio CVE",
        "full_title": "Auditoria e Conciliação de Rateio CVE por Centro de Custo",
        "enabled": False,
        "status_badge": "🟡 EM VALIDAÇÃO",
        "status_label": "Módulo construído — Em fase de validação pela equipe de Benefícios",
        "default_subfolder": r"Rateio CVE\Relatórios Originais",
        "default_out_subfolder": r"Rateio CVE",
        "output_filename": "Conferência Rateio CVE_Gerado.xlsx",
        "description": (
            "Consolida e audita o rateio contábil de custos e taxas de benefícios por Empresa, "
            "Estabelecimento e Centro de Custo para contabilização financeira."
        ),
        "events": ["Rateio Contábil de Benefícios por CC"],
        "required_files": ["Faturas / Extratos CVE", "COLA_Colaboradores.CSV"],
    },
    {
        "id": "intermitente",
        "code": "09",
        "short_name": "Intermitentes (VA/VT)",
        "full_title": "Auditoria de Benefícios — Colaboradores Intermitentes",
        "enabled": False,
        "status_badge": "🟡 PENDENTE ENVIO",
        "status_label": "Módulo construído — Pendente envio para homologação da equipe",
        "default_subfolder": r"Intermitente\Relatórios Originais",
        "default_out_subfolder": r"Intermitente",
        "output_filename": "Conferência Intermitente_Gerado.xlsx",
        "description": (
            "Valida o pagamento e desconto proporcional de Vale Alimentação e Vale Transporte "
            "para colaboradores em regime de contrato intermitente conforme convocações."
        ),
        "events": ["Eventos de VA/VT Intermitentes"],
        "required_files": ["Relatório de Convocações", "FOPA_Eventos_por_Colaborador.CSV"],
    },
    {
        "id": "desligados_planos",
        "code": "10",
        "short_name": "Desligados Demais Planos",
        "full_title": "Auditoria de Exclusão de Desligados — Demais Planos",
        "enabled": False,
        "status_badge": "🟡 PENDENTE ENVIO",
        "status_label": "Módulo construído — Pendente envio para homologação da equipe",
        "default_subfolder": r"Desligados Demais Planos\Relatórios Originais",
        "default_out_subfolder": r"Desligados Demais Planos",
        "output_filename": "Conferência Desligados Demais Planos_Gerado.xlsx",
        "description": (
            "Cruza a base de colaboradores desligados contra as bases ativas das operadoras de planos "
            "e benefícios para identificar cobranças indevidas pós-rescisão."
        ),
        "events": ["Cobranças Pós-Desligamento em Operadoras"],
        "required_files": ["COLA_Demitidos.CSV", "Bases Ativas das Operadoras"],
    },
    {
        "id": "conferencia_escritorio",
        "code": "11",
        "short_name": "Conferência Escritório",
        "full_title": "Auditoria de Desconto Escritório (Evento 6251)",
        "enabled": False,
        "status_badge": "🟡 EM VALIDAÇÃO",
        "status_label": "Módulo gerado dia 16 — Em fase de validação pela equipe",
        "default_subfolder": r"Conferencia escritorio\Relatórios Originais",
        "default_out_subfolder": r"Conferencia escritorio\Relatório Tratado",
        "output_filename": "Conferência Desconto Escritório 6251_09.2026.xlsx",
        "description": (
            "Audita o desconto de escritório (Evento 6251) aplicado aos colaboradores administrativos."
        ),
        "events": ["6251 - Desconto Escritório"],
        "required_files": ["FOPA_Eventos_por_Colaborador.CSV", "COLA_Colaboradores.CSV"],
    },

    # =========================================================================
    # 🟡 MÓDULOS ETAPA B (ESTRUTURADOS — PENDENTE ENVIO PARA VALIDAÇÃO)
    # =========================================================================
    {
        "id": "unidade_entrega",
        "code": "12",
        "short_name": "Unidade de Entrega",
        "full_title": "Auditoria de Unidades de Entrega VT (RELAOD x Optantes)",
        "enabled": False,
        "status_badge": "🟡 PENDENTE ENVIO",
        "status_label": "Módulo estruturado — Pendente envio para validação da equipe",
        "default_subfolder": r"Unidade de Entrega",
        "default_out_subfolder": r"Unidade de Entrega\Relatório Tratado",
        "output_filename": "Conferência Unidade de Entrega_10.2026.xlsx",
        "description": (
            "Concilia o código da unidade de entrega dos pedidos de VT contra a loja de lotação (Estabelecimento) "
            "e parametrizações de optantes, tratando exceções homologadas (Loja 338 e operações SPID)."
        ),
        "events": ["Roteirização de Entrega de Benefícios", "Validação Est x UE", "Validação UE x Optantes"],
        "required_files": ["RELAOD~1.CSV (Relação de Compra VT)", "Optantes.xlsx (Cadastro de Optantes)"],
    },
    {
        "id": "cencopay",
        "code": "13",
        "short_name": "Cencopay",
        "full_title": "Auditoria de Benefícios Cencopay (Margem 30% e Desconto 5081)",
        "enabled": False,
        "status_badge": "🟡 PENDENTE ENVIO",
        "status_label": "Módulo estruturado — Pendente envio para validação da equipe",
        "default_subfolder": r"Cencopay",
        "default_out_subfolder": r"Cencopay\Relatório Tratado",
        "output_filename": "Conferência Cencopay_Auditada_09.2026.xlsx",
        "description": (
            "Audita saldos devedores (14004 + 14002) frente ao teto da margem consignável de 30% do líquido (10990), "
            "desconto em folha (5081) e transferência de resíduo para o mês seguinte (14003)."
        ),
        "events": ["10990 - Líquido a Receber", "5081 - Desconto Gasto Cencopay", "14001 - Limite 30%", "14003 - Resíduo Mês Seg."],
        "required_files": ["Conferência Cencopay_Base.xlsx", "FOPA_Eventos_por_Colaborador.csv", "Bases Funcionais Complementares"],
    },
    {
        "id": "retorno_lg",
        "code": "14",
        "short_name": "Arquivo de Retorno LG",
        "full_title": "Auditoria de Arquivo de Retorno LG e Gestão de Saldo",
        "enabled": False,
        "status_badge": "🟡 PENDENTE ENVIO",
        "status_label": "Módulo estruturado — Pendente envio para validação da equipe",
        "default_subfolder": r"Arquivo de Retorno",
        "default_out_subfolder": r"Arquivo de Retorno\Relatório Tratado",
        "output_filename": "Conferência Arquivo Retorno LG_Auditada_10.2026.xlsx",
        "description": (
            "Concilia relatórios Alelo com gestão de saldo (Sem RC x Com RC) contra a compra efetiva de VT, "
            "gerando automaticamente os lotes de importação (31903) e exclusão de zerados para o sistema LG."
        ),
        "events": ["31903 - Custo Total VT", "Gestão de Saldo Operadora (Sem RC x Com RC)", "Arquivo Importação / Exclusão LG"],
        "required_files": ["Relatorios Alelo (.ods / .xlsx)", "Relação de Compra de Vale Transporte Tratado.xlsx"],
    },
    {
        "id": "compra_ticket_bretas",
        "code": "15",
        "short_name": "Compra Ticket (Bretas)",
        "full_title": "Auditoria de Compra Mensal Ticket — Regional Bretas",
        "enabled": False,
        "status_badge": "🟡 PENDENTE ENVIO",
        "status_label": "Módulo estruturado — Pendente envio para validação da equipe",
        "default_subfolder": r"Compra Mensal Ticket",
        "default_out_subfolder": r"Compra Mensal Ticket\Relatório Tratado",
        "output_filename": "Conferência Compra Ticket Bretas_Auditada_09.2026.xlsx",
        "description": (
            "Audita a carga mensal dos cartões Ticket da regional Bretas segundo acordos sindicais, jornadas, "
            "proporcionalidade de férias e suspensão por afastamento/inelegibilidade, gerando remessa de compra manual."
        ),
        "events": ["480 - Compra Vale Alimentação", "31500 - Valor Facial VR/VA", "Carga Mensal Fixa / Proporcional"],
        "required_files": ["Base de Conferencia Bretas_09.26.xlsx"],
    },
    {
        "id": "cesta_premio_bretas",
        "code": "16",
        "short_name": "Cesta Prêmio (Bretas)",
        "full_title": "Auditoria de Cesta Prêmio — Regional Bretas",
        "enabled": False,
        "status_badge": "🟡 PENDENTE ENVIO",
        "status_label": "Módulo estruturado — Pendente envio para validação da equipe",
        "default_subfolder": r"Compra Mensal Cesta Prêmio",
        "default_out_subfolder": r"Compra Mensal Cesta Prêmio\Relatório Tratado",
        "output_filename": "Conferência Cesta Prêmio Bretas_Auditada_10.2026.xlsx",
        "description": (
            "Valida a elegibilidade da Cesta Prêmio da bandeira Bretas baseando-se em data de admissão, centro de custo, "
            "cargo homologado e assiduidade (ausência de faltas e atestados), calculando valores cheios e proporcionais."
        ),
        "events": ["825 - Compra Vale Alimentação", "31638 - Dias Afastamento Cesta Prêmio", "Remessa Cesta Prêmio Manual"],
        "required_files": ["Base Conf Premio BT Outubro 26.xlsx"],
    },
]
