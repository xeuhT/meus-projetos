"""
Módulo de Regras de Negócio e Validações de Elegibilidade / FGTS
Conectividade Social V2 - Cencosud Brasil
Trata categorias de colaboradores, situações especiais (afastamentos, rescisões) e conciliação cadastral.
"""
from typing import Dict, Any, Optional

# Mapeamento oficial de categorias FGTS (Caixa Econômica Federal)
CATEGORIAS_FGTS = {
    "01": "Empregado com contrato de trabalho regido pela CLT (Geral)",
    "02": "Empregado doméstico",
    "03": "Trabalhador avulso",
    "04": "Empregado sob contrato de trabalho por prazo determinado (Lei 9.601/98)",
    "05": "Diretor não empregado com FGTS",
    "06": "Empregado sob contrato de trabalho por prazo determinado (Lei 6.019/74)",
    "07": "Menor Aprendiz (Lei 10.097/2000)",
    "12": "Contrato Verde e Amarelo"
}

# Códigos de Movimentação FGTS mais frequentes (Rescisões e Afastamentos)
CODIGOS_MOVIMENTACAO = {
    "I1": "Rescisão sem justa causa pelo empregador",
    "I2": "Rescisão por culpa recíproca ou força maior",
    "I3": "Rescisão por término de contrato a termo",
    "I4": "Rescisão sem justa causa do contrato de trabalho por prazo determinado",
    "J": "Rescisão com justa causa pelo empregador",
    "L": "Outros motivos de rescisão de contrato",
    "M": "Mudança de regime trabalhista",
    "N1": "Transferência de empregado para outro estabelecimento",
    "O1": "Afastamento temporário por motivo de acidente de trabalho",
    "P1": "Afastamento temporário por motivo de doença (> 15 dias)",
    "Q1": "Afastamento temporário por licença maternidade",
    "R": "Afastamento temporário por serviço militar",
    "W": "Afastamento para exercício de mandato sindical"
}


def auditar_situacao_colaborador(dados_extrato: Dict[str, Any]) -> Dict[str, Any]:
    """
    Audita as regras de negócio e situações especiais do colaborador:
    1. Validação de Categoria (ex: Categoria 07 Menor Aprendiz alíquota 2% vs Categoria 01 CLT alíquota 8%).
    2. Identificação de movimentação / rescisão ou afastamento (INSS/Acidente/Licença).
    3. Conciliação entre PIS pesquisado e PIS retornado pela Caixa.
    4. Tratamento de consultas sem sucesso ou contas não localizadas.
    """
    status = str(dados_extrato.get("status") or "").upper()
    mensagem = str(dados_extrato.get("mensagem") or "")
    conta_fgts = str(dados_extrato.get("conta_fgts") or "").strip()

    # Ponto 1 da Auditoria: Tratamento de Consultas Não Localizadas ou com Falha
    if "SUCESSO" not in status or not conta_fgts:
        dados_extrato["categoria_descricao"] = "Não aplicável / Não localizado"
        dados_extrato["regra_fgts"] = "Consulta não realizada / Conta inexistente"
        dados_extrato["divergencias_auditoria"] = f"Alerta Caixa: {mensagem if mensagem else 'Conta não localizada'}"
        return dados_extrato

    alertas = []

    # Regra 1: Validação e Descrição de Categoria FGTS
    categoria = str(dados_extrato.get("categoria") or "").strip().zfill(2)
    descricao_cat = CATEGORIAS_FGTS.get(categoria, f"Categoria {categoria}")
    dados_extrato["categoria_descricao"] = descricao_cat

    if categoria == "07":
        dados_extrato["regra_fgts"] = "Menor Aprendiz (Alíquota reduzida 2%)"
    elif categoria == "01":
        dados_extrato["regra_fgts"] = "CLT Geral (Alíquota padrão 8%)"
    else:
        dados_extrato["regra_fgts"] = f"Específica ({descricao_cat})"

    # Regra 2: Movimentação, Rescisão e Afastamentos (INSS / Acidente)
    cod_mov = str(dados_extrato.get("cod_movimentacao") or "").strip().upper()
    data_mov = str(dados_extrato.get("data_movimentacao") or "").strip()
    
    if cod_mov and cod_mov != "-":
        desc_mov = CODIGOS_MOVIMENTACAO.get(cod_mov, f"Movimentação cód. {cod_mov}")
        if data_mov:
            dados_extrato["movimentacao_descricao"] = f"{data_mov} - Cód. {cod_mov} ({desc_mov})"
        else:
            dados_extrato["movimentacao_descricao"] = f"Cód. {cod_mov} ({desc_mov})"

        # Se for rescisão sem justa causa (I1, I2, I4), exige valor base rescisório calculado
        if cod_mov in ["I1", "I2", "I4"]:
            val_rescisorio = dados_extrato.get("valor_base_rescisorios")
            if val_rescisorio is None or val_rescisorio <= 0:
                alertas.append(f"Atenção Rescisão ({cod_mov}): Valor base para fins rescisórios zerado ou ausente.")
        # Se for afastamento temporário (INSS/Acidente P1, O1)
        elif cod_mov in ["O1", "P1"]:
            alertas.append(f"Colaborador Afastado ({desc_mov} em {data_mov})")
    else:
        dados_extrato["movimentacao_descricao"] = "Sem movimentação registrada (Ativo)"

    # Regra 3: Conciliação de Identificação Cadastral (PIS Pesquisado vs Retornado)
    pis_pesquisado = "".join(ch for ch in str(dados_extrato.get("pis_pesquisado") or "") if ch.isdigit())
    pis_retornado = "".join(ch for ch in str(dados_extrato.get("pis_caixa") or "") if ch.isdigit())
    if pis_pesquisado and pis_retornado and pis_pesquisado != pis_retornado:
        alertas.append(f"Divergência de PIS: Informado {pis_pesquisado} vs Caixa {pis_retornado}")

    # Regra 4: Análise do Valor Base para Fins Rescisórios vs Saldo
    val_rescisorio = dados_extrato.get("valor_base_rescisorios")
    saldo = dados_extrato.get("saldo_fgts")

    if val_rescisorio is not None and saldo is not None:
        if val_rescisorio < saldo:
            alertas.append("Alerta: Valor para Fins Rescisórios menor que o Saldo atual.")

    if alertas:
        dados_extrato["divergencias_auditoria"] = " | ".join(alertas)
    else:
        dados_extrato["divergencias_auditoria"] = "Sem divergências cadastrais"

    return dados_extrato
