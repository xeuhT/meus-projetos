# -*- coding: utf-8 -*-
"""
CENCOSUD BRASIL - ENGENHARIA DE DADOS & AUDITORIA DE RH
Módulo 6: Super Módulo de Malha Fina de Folha de Pagamento x Ponto (Prioridade 10)
Competência: 09/2026
"""

import os
import time
import datetime
import unicodedata
from collections import defaultdict
import pyxlsb
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 1. CAMINHOS E CONFIGURAÇÕES GERAIS
# ==============================================================================
BASE_DIR = r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\ponto"
FILE_FOLHA = os.path.join(BASE_DIR, "2 - Conferencia folha 09.2026.xlsx")
FILE_SIND_XLSB = os.path.join(
    BASE_DIR,
    "Perfil de Frequência x Sindicato",
    "COLA_Colaboradores_CSV_Gestor -.xlsb",
)
FILE_SOBREAVISO_REF = os.path.join(
    BASE_DIR,
    "Sobreaviso",
    "BASE CONSULTA GERAL.xlsx",
)
OUT_DIR = os.path.join(BASE_DIR, "conferencias geradas pelo script")
OUT_FILE = os.path.join(
    OUT_DIR, "Relatorio_Conferencia_Malha_Fina_Folha_09.2026.xlsx"
)

# Grupos de Eventos FOPA (Aba Contabil)
GRP_SALARIO_BASE = {1, 7, 421}
GRP_HORAS_EXTRAS = {
    27, 28, 29, 30, 31, 32, 33, 37, 153, 156, 1656,
    1000, 1002, 1003, 1007, 1015, 1016
}
GRP_BANCO_HORAS_PROV = {63, 64, 67, 68, 139, 142, 146, 1009, 1010, 1011, 1563}
GRP_ADIC_NOTURNO = {69, 71}
GRP_DSR_PROVENTO = {60, 78, 150, 8, 66, 82, 90, 1050, 1051, 1508, 1560, 1590, 1650}
GRP_GRAT_FERIADO = {220, 227}
GRP_SOBREAVISO = {65, 66}
GRP_DESC_FALTAS_5001 = {5001}
GRP_DESC_DSR_5007 = {5007, 5112}
GRP_DESC_ATRASOS_5008 = {5008}
GRP_DESC_BH_5111 = {5111}

REGIONAIS_ORDEM = [
    "Bretas",
    "GBarbosa",
    "Giga",
    "Mercantil",
    "Perini",
    "Prezunic",
    "Spid",
    "Corporativo",
]

REGRAS_ORDEM = [
    ("1. Isento recebendo HE / BH / Adic. Noturno (Art. 62 CLT)", "Crítico"),
    ("2. Aprendiz com Pagamento de HE / BH / Noturno em Folha (Ilegal Art. 432 CLT)", "Crítico"),
    ("3. Estagiário com Evento Indevido ou Perfil Não Isento", "Crítico"),
    ("4. Escala 12x36 fora do Sind. 200 recebendo Grat. Feriado", "Crítico"),
    ("5. Adicional Noturno em Turno Diurno ou Valor Elevado (>= R$ 400)", "Crítico"),
    ("6. Inconsistência Faltas (5001) x Desconto DSR (5007)", "Alerta"),
    ("7. Desconto Elevado de Banco de Horas - Evento 5111 (>= R$ 300)", "Alerta"),
    ("8. Desconto Elevado de Faltas/Atrasos - Eventos 5001/5008", "Alerta"),
    ("9. Pico Financeiro de HE ou Banco de Horas Pago (>= R$ 800)", "Alerta"),
    ("10. Alerta Cadastral (Sem Escala no Ponto ou Sobreaviso Atípico)", "Alerta"),
    ("Conforme Malha Fina Folha x Ponto", "Regular"),
]


def norm_str(val):
    """Normaliza string para maiúsculas sem acentos."""
    if val is None:
        return ""
    s = str(val).strip().upper()
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )


def map_regional(cod_emp, emp, loja, estab, unid_org, sind_nome):
    """Mapeia a Regional/Bandeira a partir dos dados contábeis e sindicais."""
    txt = norm_str(f"{emp} {estab} {unid_org}")
    sind_txt = norm_str(sind_nome)

    if "SPID" in txt or (1700 <= loja <= 1799):
        return "Spid"
    if cod_emp == 7 or "PERINI" in txt or (850 <= loja <= 899):
        return "Perini"
    if cod_emp == 5 or "GIGA" in txt or "ATACADO" in txt or (900 <= loja <= 999):
        return "Giga"
    if "PREZUNIC" in txt or (700 <= loja <= 799) or any(
        k in sind_txt for k in ["RIO DE JANEIRO", "NITEROI", "DUQUE DE CAXIAS", "NILOPOLIS", "SAO GONCALO"]
    ):
        return "Prezunic"
    if "BRETAS" in txt or (500 <= loja <= 699) or any(
        k in sind_txt for k in ["EST GO", "MINAS", "SETE LAGOAS", "JUIZ DE FORA", "UBERLANDIA", "GOIANIA"]
    ):
        return "Bretas"
    if "MERCANTIL" in txt or (800 <= loja <= 849) or any(
        k in sind_txt for k in ["FORTALEZA", "CEARA"]
    ):
        return "Mercantil"
    if any(k in txt for k in ["CSC", "CORPORATIVO", "HOLDING", "DIRETORIA", "ADMINISTRACAO CENTRAL"]):
        return "Corporativo"
    if any(k in txt for k in ["GBARBOSA", "ELETROSHOW", "HIPER", "FARMACIA"]) or (1 <= loja <= 499):
        return "GBarbosa"
    return "Corporativo"


def parse_date(val):
    """Converte valor de data para datetime.date."""
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    if isinstance(val, (int, float)):
        try:
            return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(val))).date()
        except Exception:
            return None
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None


def evaluate_malha_fina(rec, matriz_sobreaviso):
    """Avalia as 10 regras de precedência da Malha Fina Folha x Ponto."""
    perfil_norm = norm_str(rec["perfil"])
    cargo_norm = norm_str(rec["cargo"])
    cc_norm = norm_str(rec.get("centro_custo", ""))
    escala_norm = norm_str(rec["escala"])
    situacao_norm = norm_str(rec["situacao"])

    he = rec["he"]
    bh_prov = rec["bh_prov"]
    adic_not = rec["adic_not"]
    grat_fer = rec["grat_fer"]
    sobreaviso = rec["sobreaviso"]
    desc_5001 = rec["desc_5001"]
    desc_5007 = rec["desc_5007"]
    desc_5008 = rec["desc_5008"]
    desc_5111 = rec["desc_5111"]
    ev_421 = rec["ev_421"]
    cod_sind = rec["cod_sind"]

    soma_jornada = round(he + bh_prov + adic_not, 2)

    # 1. ISENTO COM PAGAMENTO INDEVIDO DE JORNADA (Art. 62 CLT)
    if "ISENTO" in perfil_norm and soma_jornada > 0:
        return (
            "1. Isento recebendo HE / BH / Adic. Noturno (Art. 62 CLT)",
            "Crítico",
            "Estornar pagamento de horas suplementares/noturnas. Colaboradores com perfil Isento não fazem jus a horas extras.",
            soma_jornada,
        )

    # 2. APRENDIZ COM PAGAMENTO DE HE / BH / NOTURNO (Art. 432 CLT)
    if "APRENDIZ" in cargo_norm and soma_jornada > 0:
        return (
            "2. Aprendiz com Pagamento de HE / BH / Noturno em Folha (Ilegal Art. 432 CLT)",
            "Crítico",
            "Auditar e bloquear imediatamente. A CLT veda prorrogação/compensação de jornada e trabalho noturno para aprendizes.",
            soma_jornada,
        )

    # 3. ESTAGIÁRIO COM ANOMALIA DE PONTO / FOLHA (Lei 11.788/08)
    if ("ESTAGIARIO" in cargo_norm or ev_421 > 0) and (
        soma_jornada > 0 or (perfil_norm != "" and "ISENTO" not in perfil_norm)
    ):
        impacto = soma_jornada if soma_jornada > 0 else rec["sal_base"]
        return (
            "3. Estagiário com Evento Indevido ou Perfil Não Isento",
            "Crítico",
            "Ajustar perfil do estagiário para Isento e bloquear verbas salariais restritas a CLT.",
            impacto,
        )

    # 4. ESCALA 12X36 PREZUNIC FORA DO SINDICATO 200 COM GRATIFICAÇÃO
    if (
        ("12X36" in escala_norm or "12X36" in perfil_norm)
        and grat_fer > 0
        and cod_sind is not None
        and cod_sind != 200
    ):
        return (
            "4. Escala 12x36 fora do Sind. 200 recebendo Grat. Feriado",
            "Crítico",
            "Verificar pagamento indevido de gratificação de feriado (exclusivo Sindicato 200 RJ na escala 12x36).",
            grat_fer,
        )

    # 5. ADICIONAL NOTURNO EM TURNO DIURNO OU PICO CRÍTICO
    has_diurno = any(h in escala_norm for h in ["07:00", "07:30", "08:00"])
    has_noturno = any(
        n in escala_norm for n in ["22:", "23:", "00:", "01:", "02:", "03:", "04:", "05:"]
    )
    if adic_not > 0 and ((has_diurno and not has_noturno and adic_not >= 100.0) or adic_not >= 400.0):
        return (
            "5. Adicional Noturno em Turno Diurno ou Valor Elevado (>= R$ 400)",
            "Crítico",
            "Conferir espelho de ponto e escala do colaborador. Pagamento elevado de adicional noturno em turno diurno.",
            adic_not,
        )

    # 6. DESCASAMENTO FALTAS (5001) X DESCONTO DSR (5007)
    if (
        (desc_5007 > 0 and desc_5001 == 0)
        or (desc_5001 > 0 and desc_5007 == 0)
        or (desc_5007 > desc_5001 and desc_5001 > 0)
    ):
        return (
            "6. Inconsistência Faltas (5001) x Desconto DSR (5007)",
            "Alerta",
            "Validar proporção entre faltas injustificadas (5001) e desconto de DSR (5007/5112) no fechamento.",
            round(desc_5001 + desc_5007, 2),
        )

    # 7. DESCONTO ELEVADO DE DÉBITO BANCO DE HORAS (5111)
    if desc_5111 >= 300.0:
        return (
            "7. Desconto Elevado de Banco de Horas - Evento 5111 (>= R$ 300)",
            "Alerta",
            "Auditar em ordem decrescente se o desconto de débito de BH decorre de falta de justificativa no ponto.",
            desc_5111,
        )

    # 8. DESCONTO ELEVADO DE FALTAS / ATRASOS (5001 / 5008)
    if desc_5008 >= 200.0 or desc_5001 >= 500.0:
        return (
            "8. Desconto Elevado de Faltas/Atrasos - Eventos 5001/5008",
            "Alerta",
            "Confirmar se os descontos severos de falta/atraso não possuem atestado ou abono pendente na loja.",
            round(desc_5001 + desc_5008, 2),
        )

    # 9. PICO FINANCEIRO DE HORAS EXTRAS OU BANCO DE HORAS EM FOLHA
    if he >= 800.0 or bh_prov >= 800.0:
        return (
            "9. Pico Financeiro de HE ou Banco de Horas Pago (>= R$ 800)",
            "Alerta",
            "Auditar teto financeiro de horas extras / quitação de banco de horas na folha.",
            round(he + bh_prov, 2),
        )

    # 10. SOBREAVISO (CONSULTA À MATRIZ OFICIAL 'BASE CONSULTA GERAL.xlsx') / SEM ESCALA CADASTRADA
    if sobreaviso > 0:
        cargos_proibidos = [
            "APRENDIZ", "ESTAGIARIO", "OPERADOR DE CAIXA", "OPERADOR DE SUPERMERCADO",
            "OPERADOR DE LOJA", "OPERADOR DE PERECIVEIS", "EMPACOTADOR", "REPOSITOR",
            "ACOUGUEIRO", "PADEIRO", "BALCONISTA", "AUXILIAR DE LOJA",
        ]
        status_matriz = matriz_sobreaviso["pair_map"].get(
            (cc_norm, cargo_norm),
            matriz_sobreaviso["cargo_map"].get(cargo_norm, None),
        )
        is_proibido = any(p in cargo_norm for p in cargos_proibidos) or status_matriz == "NAO"
        is_autorizado = (
            status_matriz == "REALIZA"
            or any(
                k in cargo_norm
                for k in [
                    "TECNICO", "ENGENHEIRO", "ENGINNER", "ENGINEER", "ANALISTA",
                    "ANALYST", "COORDENADOR", "GERENTE", "ESPECIALISTA",
                    "SUPERVISOR", "MANUTENCAO", "TI", "SUPPORT", "DEVELOPER",
                    "ARCHITECT", "LEAD", "LEADER", "MANAGER",
                ]
            )
        )
        if is_proibido or not is_autorizado:
            return (
                "10. Alerta Cadastral (Sem Escala no Ponto ou Sobreaviso Atípico)",
                "Crítico" if is_proibido else "Alerta",
                "Alerta Crítico de Sobreaviso Indevido: Cargo/Centro de Custo consta como 'NÃO' ou proibido na Matriz Geral de Sobreaviso (Eventos 65/66).",
                sobreaviso,
            )

    if escala_norm == "" and situacao_norm == "ATIVIDADE NORMAL":
        return (
            "10. Alerta Cadastral (Sem Escala no Ponto ou Sobreaviso Atípico)",
            "Alerta",
            "Regularizar vínculo de escala/perfil de frequência no sistema de ponto.",
            rec["sal_base"],
        )

    # 11. REGULAR
    return (
        "Conforme Malha Fina Folha x Ponto",
        "Regular",
        "Conforme",
        round(rec["sal_base"] + soma_jornada + rec["dsr_prov"] + grat_fer + sobreaviso, 2),
    )


def main():
    t_start = time.time()
    print("=" * 85)
    print("CENCOSUD BRASIL - ENGENHARIA DE DADOS & AUDITORIA DE RH")
    print("MÓDULO 6: SUPER MÓDULO MALHA FINA DE FOLHA DE PAGAMENTO X PONTO (09/2026)")
    print("=" * 85)

    os.makedirs(OUT_DIR, exist_ok=True)

    # --------------------------------------------------------------------------
    # 1. Carregar Sindicatos (XLSB) e Matriz Oficial de Sobreaviso (XLSX)
    # --------------------------------------------------------------------------
    t0 = time.time()
    print("-> [1/4] Carregando base de Sindicatos (XLSB) e Matriz Oficial de Sobreaviso...")
    sind_map = {}
    with pyxlsb.open_workbook(FILE_SIND_XLSB) as wb_x:
        with wb_x.get_sheet("VALIDAÇÃO") as sh:
            for i, row in enumerate(sh.rows()):
                if i < 2:
                    continue
                vals = [c.v for c in row]
                if len(vals) >= 6 and vals[1] is not None:
                    try:
                        mat = int(float(vals[1]))
                        cod_s = int(float(vals[3])) if vals[3] is not None else None
                        nome_s = str(vals[4]).strip() if vals[4] is not None else ""
                        sind_map[mat] = (cod_s, nome_s)
                    except Exception:
                        pass

    matriz_sobreaviso = {"pair_map": {}, "cargo_map": {}}
    if os.path.exists(FILE_SOBREAVISO_REF):
        wb_sob = openpyxl.load_workbook(FILE_SOBREAVISO_REF, read_only=True, data_only=True)
        if "SOBREAVISO" in wb_sob.sheetnames:
            ws_sob = wb_sob["SOBREAVISO"]
            for i, r in enumerate(ws_sob.iter_rows(values_only=True)):
                if i == 0 or not r or r[1] is None:
                    continue
                cc_s = norm_str(r[0])
                cg_s = norm_str(r[1])
                st_s = norm_str(r[2]) if len(r) > 2 else ""
                val_norm = "REALIZA" if "REALIZA" in st_s else "NAO"
                matriz_sobreaviso["pair_map"][(cc_s, cg_s)] = val_norm
                if cg_s not in matriz_sobreaviso["cargo_map"] or val_norm == "NAO":
                    matriz_sobreaviso["cargo_map"][cg_s] = val_norm
        wb_sob.close()

    print(
        f"   Bases carregadas em {time.time() - t0:.2f}s "
        f"({len(sind_map):,} sindicatos | {len(matriz_sobreaviso['pair_map'])} regras de sobreaviso)."
    )

    # --------------------------------------------------------------------------
    # 2. Carregar Arquivo Mestre de Folha (PERFIL DE FREQUENCIA + Contabil)
    # --------------------------------------------------------------------------
    t1 = time.time()
    print("-> [2/4] Carregando arquivo mestre '2 - Conferencia folha 09.2026.xlsx'...")
    wb_in = openpyxl.load_workbook(FILE_FOLHA, read_only=True, data_only=True)

    ws_perf = wb_in["PERFIL DE FREQUENCIA"]
    perf_map = {}
    for i, row in enumerate(ws_perf.iter_rows(values_only=True)):
        if i == 0:
            continue
        if len(row) >= 12 and row[4] is not None:
            try:
                mat = int(float(row[4]))
                escala = str(row[7]).strip() if row[7] is not None else ""
                perfil = str(row[11]).strip() if row[11] is not None else ""
                perf_map[mat] = (escala, perfil)
            except Exception:
                pass
    print(f"   Aba 'PERFIL DE FREQUENCIA' indexada ({len(perf_map):,} matrículas).")

    ws_cont = wb_in["Contabil"]
    colab_data = {}
    total_lancamentos = 0

    for i, row in enumerate(ws_cont.iter_rows(values_only=True)):
        if i < 3:
            continue
        if len(row) < 26 or row[2] is None:
            continue

        total_lancamentos += 1
        try:
            mat = int(float(row[2]))
        except Exception:
            continue

        if mat not in colab_data:
            cod_emp = int(float(row[0])) if row[0] is not None else 0
            emp = str(row[1]).strip() if row[1] is not None else ""
            nome = str(row[3]).strip() if row[3] is not None else ""
            dt_adm = parse_date(row[6])
            situacao = str(row[7]).strip() if row[7] is not None else ""
            try:
                loja = int(float(row[10])) if row[10] is not None else 0
            except Exception:
                loja = 0
            estab = str(row[11]).strip() if row[11] is not None else ""
            unid_org = str(row[13]).strip() if len(row) > 13 and row[13] is not None else ""
            centro_custo = str(row[17]).strip() if len(row) > 17 and row[17] is not None else ""
            cargo = str(row[19]).strip() if row[19] is not None else ""

            escala, perfil = perf_map.get(mat, ("", ""))
            cod_sind, nome_sind = sind_map.get(mat, (None, ""))
            if cod_sind is not None and nome_sind:
                sind_label = f"{cod_sind} - {nome_sind}"
            elif nome_sind:
                sind_label = nome_sind
            else:
                sind_label = "NÃO IDENTIFICADO NA BASE SINDICAL"

            regional = map_regional(cod_emp, emp, loja, estab, unid_org, nome_sind)

            colab_data[mat] = {
                "mat": mat,
                "nome": nome,
                "regional": regional,
                "loja": loja,
                "dt_adm": dt_adm,
                "situacao": situacao,
                "centro_custo": centro_custo,
                "cargo": cargo,
                "escala": escala,
                "perfil": perfil,
                "cod_sind": cod_sind,
                "sind_label": sind_label,
                "sal_base": 0.0,
                "he": 0.0,
                "bh_prov": 0.0,
                "adic_not": 0.0,
                "dsr_prov": 0.0,
                "grat_fer": 0.0,
                "sobreaviso": 0.0,
                "grat_fer_sobreaviso": 0.0,
                "desc_5001": 0.0,
                "desc_5007": 0.0,
                "desc_5008": 0.0,
                "desc_5111": 0.0,
                "ev_421": 0.0,
            }

        try:
            cod_ev = int(float(row[23])) if row[23] is not None else -1
            val = float(row[25]) if row[25] is not None else 0.0
        except Exception:
            continue

        rec = colab_data[mat]
        if cod_ev in GRP_SALARIO_BASE:
            rec["sal_base"] += val
        if cod_ev == 421:
            rec["ev_421"] += val
        if cod_ev in GRP_HORAS_EXTRAS:
            rec["he"] += val
        if cod_ev in GRP_BANCO_HORAS_PROV:
            rec["bh_prov"] += val
        if cod_ev in GRP_ADIC_NOTURNO:
            rec["adic_not"] += val
        if cod_ev in GRP_DSR_PROVENTO:
            rec["dsr_prov"] += val
        if cod_ev in GRP_GRAT_FERIADO:
            rec["grat_fer"] += val
        if cod_ev in GRP_SOBREAVISO:
            rec["sobreaviso"] += val
        if cod_ev in GRP_GRAT_FERIADO or cod_ev in GRP_SOBREAVISO:
            rec["grat_fer_sobreaviso"] += val
        if cod_ev in GRP_DESC_FALTAS_5001:
            rec["desc_5001"] += val
        if cod_ev in GRP_DESC_DSR_5007:
            rec["desc_5007"] += val
        if cod_ev in GRP_DESC_ATRASOS_5008:
            rec["desc_5008"] += val
        if cod_ev in GRP_DESC_BH_5111:
            rec["desc_5111"] += val

    wb_in.close()
    print(
        f"   Aba 'Contabil' processada em {time.time() - t1:.2f}s: "
        f"{total_lancamentos:,} lançamentos | {len(colab_data):,} colaboradores únicos."
    )

    # --------------------------------------------------------------------------
    # 3. Avaliar Regras de Auditoria da Malha Fina (Prioridade 10)
    # --------------------------------------------------------------------------
    print("-> [3/4] Aplicando as 10 Regras de Precedência da Malha Fina...")
    records_all = []
    records_div = []

    reg_stats = {
        r: {
            "total": 0,
            "regular": 0,
            "critico": 0,
            "alerta": 0,
            "proventos_jornada": 0.0,
            "descontos_ponto": 0.0,
        }
        for r in REGIONAIS_ORDEM
    }

    rule_stats = {
        regra: {"sev": sev, "count": 0, "montante": 0.0}
        for regra, sev in REGRAS_ORDEM
    }

    total_he_bh_not = 0.0
    total_descontos = 0.0
    total_criticos = 0
    total_alertas = 0
    total_regulares = 0

    for mat, rec in colab_data.items():
        for k in (
            "sal_base", "he", "bh_prov", "adic_not", "dsr_prov",
            "grat_fer", "sobreaviso", "grat_fer_sobreaviso",
            "desc_5001", "desc_5007", "desc_5008", "desc_5111", "ev_421"
        ):
            rec[k] = round(rec[k], 2)

        regra, sev, acao, impacto = evaluate_malha_fina(rec, matriz_sobreaviso)
        rec["regra"] = regra
        rec["sev"] = sev
        rec["acao"] = acao
        rec["impacto"] = round(impacto, 2)

        prov_j = round(rec["he"] + rec["bh_prov"] + rec["adic_not"], 2)
        desc_p = round(rec["desc_5001"] + rec["desc_5007"] + rec["desc_5008"] + rec["desc_5111"], 2)

        total_he_bh_not += prov_j
        total_descontos += desc_p

        r_name = rec["regional"]
        if r_name not in reg_stats:
            r_name = "Corporativo"
            rec["regional"] = r_name

        reg_stats[r_name]["total"] += 1
        reg_stats[r_name]["proventos_jornada"] += prov_j
        reg_stats[r_name]["descontos_ponto"] += desc_p

        if sev == "Crítico":
            reg_stats[r_name]["critico"] += 1
            total_criticos += 1
            records_div.append(rec)
        elif sev == "Alerta":
            reg_stats[r_name]["alerta"] += 1
            total_alertas += 1
            records_div.append(rec)
        else:
            reg_stats[r_name]["regular"] += 1
            total_regulares += 1

        rule_stats[regra]["count"] += 1
        rule_stats[regra]["montante"] += rec["impacto"]

        records_all.append(rec)

    # Ordenação:
    # Conferência Geral: Regional, Loja, Nome Colaborador
    records_all.sort(key=lambda x: (x["regional"], x["loja"], x["nome"]))

    # Divergências: Primeiro Crítico, depois Alerta, e em ordem decrescente de impacto financeiro
    sev_rank = {"Crítico": 0, "Alerta": 1, "Regular": 2}
    records_div.sort(key=lambda x: (sev_rank.get(x["sev"], 9), -x["impacto"], x["regional"], x["loja"], x["nome"]))

    total_colabs = len(records_all)
    taxa_conformidade = (total_regulares / total_colabs * 100.0) if total_colabs else 0.0

    print("=" * 85)
    print(f"RESUMO DA AUDITORIA MALHA FINA (09/2026):")
    print(f"  - Total Colaboradores Auditados : {total_colabs:,}")
    print(f"  - Lançamentos Contábeis Lidos   : {total_lancamentos:,}")
    print(f"  - Conformes ('Regular')         : {total_regulares:,} ({taxa_conformidade:.2f}%)")
    print(f"  - Divergências Críticas         : {total_criticos:,}")
    print(f"  - Alertas de Malha Fina         : {total_alertas:,}")
    print(f"  - Total Divergências (Aba 4)    : {len(records_div):,}")
    print(f"  - Total Pago HE/BH/Noturno      : R$ {total_he_bh_not:,.2f}")
    print(f"  - Total Descontos Ponto         : R$ {total_descontos:,.2f}")
    print("-" * 85)
    for regra, sev in REGRAS_ORDEM:
        st = rule_stats[regra]
        print(f"  [{sev:7s}] {regra:66s}: {st['count']:5,d} colabs | R$ {st['montante']:12,.2f}")
    print("=" * 85)

    # --------------------------------------------------------------------------
    # 4. Construção da Planilha Oficial (Padrão Cencosud de 4 Abas)
    # --------------------------------------------------------------------------
    t2 = time.time()
    print("-> [4/4] Gerando Workbook Excel Oficial no Padrão Cencosud...")

    wb = openpyxl.Workbook()
    ws_ori = wb.active
    ws_ori.title = "Orientações"
    ws_res = wb.create_sheet(title="Resumo da Auditoria")
    ws_ger = wb.create_sheet(title="Conferência Geral")
    ws_div = wb.create_sheet(title="Divergências")

    # Paleta e Estilos Cencosud
    NAVY_HEADER = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    SUB_HEADER = PatternFill(start_color="2C5282", end_color="2C5282", fill_type="solid")
    ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    TOTAL_FILL = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    CARD_BG = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    FILL_CRIT = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    FONT_CRIT = Font(name="Calibri", size=10, bold=True, color="9C0006")

    FILL_ALER = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    FONT_ALER = Font(name="Calibri", size=10, bold=True, color="9C6500")

    FILL_REG = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    FONT_REG = Font(name="Calibri", size=10, bold=True, color="006100")

    FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    FONT_SEC_HDR = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    FONT_TBL_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    FONT_DATA = Font(name="Calibri", size=10, bold=False, color="1E293B")
    FONT_DATA_BOLD = Font(name="Calibri", size=10, bold=True, color="1E293B")
    FONT_CARD_LBL = Font(name="Calibri", size=9, bold=True, color="475569")
    FONT_CARD_VAL = Font(name="Calibri", size=16, bold=True, color="1B365D")
    FONT_CARD_SUB = Font(name="Calibri", size=8, italic=True, color="64748B")

    THIN_SIDE = Side(style="thin", color="CBD5E1")
    MED_NAVY = Side(style="medium", color="1B365D")
    BORDER_DATA = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
    BORDER_TOTAL = Border(top=THIN_SIDE, bottom=Side(style="double", color="1B365D"), left=THIN_SIDE, right=THIN_SIDE)
    BORDER_CARD = Border(left=MED_NAVY, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

    ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=False)
    ALIGN_LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
    ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

    FMT_CURRENCY = "R$ #,##0.00"
    FMT_INT = "#,##0"
    FMT_PCT = "0.00%"
    FMT_DATE = "DD/MM/YYYY"

    # ==========================================================================
    # ABA 1: ORIENTAÇÕES
    # ==========================================================================
    ws_ori.views.sheetView[0].showGridLines = True
    ws_ori.merge_cells("A1:H1")
    ws_ori["A1"] = "SUPER MÓDULO: MALHA FINA DE FOLHA DE PAGAMENTO X PONTO (PRIORIDADE 10) — COMPETÊNCIA 09/2026"
    ws_ori["A1"].font = FONT_TITLE
    ws_ori["A1"].fill = NAVY_HEADER
    ws_ori["A1"].alignment = ALIGN_LEFT
    ws_ori.row_dimensions[1].height = 34

    orientacoes_rows = [
        ("1. OBJETIVO EXECUTIVO DO MÓDULO", ""),
        (
            "Este relatório consolida a Malha Fina Final entre os Eventos Calculados da Folha de Pagamento (FOPA - 198.387 lançamentos), "
            "a base ativa de Escalas e Perfis de Frequência do Ponto e o Enquadramento Sindical da Suite LG para os 15.843 colaboradores processados.",
            "",
        ),
        ("2. EMBASAMENTO LEGAL E NORMATIVO AUDITADO", ""),
        ("• Art. 62 da CLT (Exceções de Jornada - Isentos):", "Gerentes e cargos de gestão isentos de marcação não fazem jus ao recebimento de Horas Extras, Banco de Horas ou Adicional Noturno."),
        ("• Art. 432 da CLT (Contrato de Aprendizagem):", "É expressamente vedada a prorrogação e a compensação de jornada de trabalho aos aprendizes, bem como o trabalho noturno."),
        ("• Lei nº 11.788/2008 (Lei do Estágio):", "Estagiários possuem carga horária especial, perfil obrigatoriamente Isento no ponto eletrônico e vedação a verbas salariais exclusivas de CLT."),
        ("• CCT Sindicato 200 RJ (Escala 12x36 Prezunic):", "O pagamento de Gratificação de Feriado (Eventos 220/227) para escala 12x36 é restrito exclusivamente ao Sindicato 200 (Rio de Janeiro)."),
        ("• Art. 73 da CLT (Trabalho Noturno - 22h às 05h):", "Auditoria de pagamentos de adicional noturno (Eventos 69/71) para colaboradores alocados em escalas estritamente diurnas (07h/08h) ou picos >= R$ 400,00."),
        ("• Lei nº 605/1949 e Art. 462 da CLT (Faltas, DSR e Descontos):", "Conciliação cruzada entre Faltas Injustificadas (5001) e Desconto de DSR (5007/5112), além de tetos de descontos de atrasos (5008) e débito de BH (5111)."),
        ("3. MATRIZ DE SEVERIDADE E PRECEDÊNCIA (10 REGRAS)", ""),
        ("• CRÍTICO (Vermelho):", "Violações diretas de preceito legal ou normativo sindical (Regras 1 a 5: Isentos, Aprendizes, Estagiários, 12x36 fora do Sind. 200 e Adic. Noturno Diurno). Exige bloqueio/estorno imediato antes do fechamento da folha."),
        ("• ALERTA (Amarelo):", "Inconsistências financeiras de fechamento ou tetos operacionais (Regras 6 a 10: Descasamento Faltas x DSR, Descontos Severos 5111/5001/5008, Picos de HE/BH >= R$ 800 e Sobreaviso/Sem Escala). Exige validação com RH/Gestor."),
        ("• REGULAR (Verde):", "Lançamentos contábeis, perfil de frequência, escala e sindicato 100% aderentes às regras de negócio da Cencosud Brasil."),
        ("4. GRUPOS DE EVENTOS AUDITADOS NA FOLHA (ABA CONTABIL)", ""),
        ("• SALARIO_BASE: {1, 7, 421}", "Salário Mensal, Horas Trabalhadas Horista e Bolsa de Estágio."),
        ("• HORAS_EXTRAS: {27..37, 153, 156, 1656, 1000..1016}", "Horas Extras a 50%, 60%, 70%, 100% e reflexos de suplementares."),
        ("• BANCO_HORAS_PROVENTO: {63, 64, 67, 68, 139, 142, 146, 1009..1563}", "Quitação de saldo positivo de Banco de Horas em folha."),
        ("• ADICIONAL_NOTURNO: {69, 71} | SOBREAVISO: {65, 66} | GRAT_FERIADO: {220, 227}", "Adicional Noturno, Horas de Sobreaviso e Gratificações de Domingos/Feriados."),
        ("• DESCONTOS DE PONTO: {5001, 5007, 5112, 5008, 5111}", "Faltas Injustificadas (5001), DSR s/ Faltas (5007/5112), Atrasos (5008) e Débito Banco de Horas (5111)."),
    ]

    curr_row = 3
    for col1, col2 in orientacoes_rows:
        ws_ori.row_dimensions[curr_row].height = 24
        if col2 == "" and col1.startswith(("1.", "2.", "3.", "4.")):
            ws_ori.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=8)
            c = ws_ori.cell(curr_row, 1, col1)
            c.font = FONT_SEC_HDR
            c.fill = SUB_HEADER
            c.alignment = ALIGN_LEFT
        elif col2 == "":
            ws_ori.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=8)
            c = ws_ori.cell(curr_row, 1, col1)
            c.font = FONT_DATA
            c.fill = ZEBRA_FILL
            c.alignment = ALIGN_LEFT_WRAP
            c.border = BORDER_DATA
            ws_ori.row_dimensions[curr_row].height = 36
        else:
            ws_ori.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=2)
            ws_ori.merge_cells(start_row=curr_row, start_column=3, end_row=curr_row, end_column=8)
            c1 = ws_ori.cell(curr_row, 1, col1)
            c2 = ws_ori.cell(curr_row, 3, col2)
            c1.font = FONT_DATA_BOLD
            c2.font = FONT_DATA
            c1.fill = WHITE_FILL
            c2.fill = WHITE_FILL
            c1.alignment = ALIGN_LEFT
            c2.alignment = ALIGN_LEFT_WRAP
            for col_i in range(1, 9):
                ws_ori.cell(curr_row, col_i).border = BORDER_DATA
        curr_row += 1

    for col_i, width in enumerate([22, 24, 22, 22, 22, 22, 22, 25], start=1):
        ws_ori.column_dimensions[get_column_letter(col_i)].width = width

    # ==========================================================================
    # ABA 2: RESUMO DA AUDITORIA
    # ==========================================================================
    ws_res.views.sheetView[0].showGridLines = True
    ws_res.merge_cells("A1:I1")
    ws_res["A1"] = "CENCOSUD BRASIL — DASHBOARD EXECUTIVO: MALHA FINA DE FOLHA X PONTO (09/2026)"
    ws_res["A1"].font = FONT_TITLE
    ws_res["A1"].fill = NAVY_HEADER
    ws_res["A1"].alignment = ALIGN_LEFT
    ws_res.row_dimensions[1].height = 34

    # 7 Cards Executivos de KPIs (Linhas 3 a 9)
    cards_row1 = [
        (1, 2, "TOTAL COLABORADORES NA FOLHA", f"{total_colabs:,}", "Universo Único Contábil (09/2026)"),
        (3, 4, "LANÇAMENTOS CONTÁBEIS LIDOS", f"{total_lancamentos:,}", "Registros Processados na Aba Contábil"),
        (5, 6, "TAXA DE CONFORMIDADE GERAL", f"{taxa_conformidade:.2f}%", f"{total_regulares:,} Colaboradores Regulares"),
        (7, 9, "DIVERGÊNCIAS CRÍTICAS (CLT/CCT)", f"{total_criticos:,}", "Bloqueio / Estorno Imediato antes do Fechamento"),
    ]
    cards_row2 = [
        (1, 3, "ALERTAS DE MALHA FINA", f"{total_alertas:,}", "Faltas x DSR, Descontos Severos, Picos HE/BH e Escala"),
        (4, 6, "TOTAL PAGO HE / BH / NOTURNO (R$)", f"R$ {total_he_bh_not:,.2f}", "Soma dos Proventos de Jornada Auditados"),
        (7, 9, "TOTAL DESCONTOS PONTO (R$)", f"R$ {total_descontos:,.2f}", "Eventos 5001, 5007/5112, 5008 e 5111"),
    ]

    for r_idx in (3, 4, 5, 7, 8, 9):
        ws_res.row_dimensions[r_idx].height = 20
    ws_res.row_dimensions[4].height = 28
    ws_res.row_dimensions[8].height = 28

    for c_start, c_end, lbl, val_str, sub in cards_row1:
        ws_res.merge_cells(start_row=3, start_column=c_start, end_row=3, end_column=c_end)
        ws_res.merge_cells(start_row=4, start_column=c_start, end_row=4, end_column=c_end)
        ws_res.merge_cells(start_row=5, start_column=c_start, end_row=5, end_column=c_end)
        ws_res.cell(3, c_start, lbl).font = FONT_CARD_LBL
        ws_res.cell(4, c_start, val_str).font = FONT_CARD_VAL
        ws_res.cell(5, c_start, sub).font = FONT_CARD_SUB
        for r_i in (3, 4, 5):
            for c_i in range(c_start, c_end + 1):
                cell = ws_res.cell(r_i, c_i)
                cell.fill = CARD_BG
                cell.alignment = ALIGN_CENTER
                cell.border = BORDER_CARD

    for c_start, c_end, lbl, val_str, sub in cards_row2:
        ws_res.merge_cells(start_row=7, start_column=c_start, end_row=7, end_column=c_end)
        ws_res.merge_cells(start_row=8, start_column=c_start, end_row=8, end_column=c_end)
        ws_res.merge_cells(start_row=9, start_column=c_start, end_row=9, end_column=c_end)
        ws_res.cell(7, c_start, lbl).font = FONT_CARD_LBL
        ws_res.cell(8, c_start, val_str).font = FONT_CARD_VAL
        ws_res.cell(9, c_start, sub).font = FONT_CARD_SUB
        for r_i in (7, 8, 9):
            for c_i in range(c_start, c_end + 1):
                cell = ws_res.cell(r_i, c_i)
                cell.fill = CARD_BG
                cell.alignment = ALIGN_CENTER
                cell.border = BORDER_CARD

    # Tabela 1: Consolidação por Regional / Bandeira
    ws_res.merge_cells("A11:I11")
    ws_res["A11"] = "1. CONSOLIDAÇÃO DE MALHA FINA POR REGIONAL / BANDEIRA"
    ws_res["A11"].font = FONT_SEC_HDR
    ws_res["A11"].fill = SUB_HEADER
    ws_res["A11"].alignment = ALIGN_LEFT
    ws_res.row_dimensions[11].height = 25

    t1_headers = [
        "Regional / Bandeira",
        "Total Colaboradores",
        "Regulares (Conformes)",
        "Críticos (Violação CLT/CCT)",
        "Alertas (Malha Fina)",
        "Total Divergências",
        "Taxa Conformidade (%)",
        "Proventos HE/BH/Noturno (R$)",
        "Descontos Ponto (R$)",
    ]
    ws_res.row_dimensions[12].height = 25
    for idx, h_text in enumerate(t1_headers, start=1):
        c = ws_res.cell(12, idx, h_text)
        c.font = FONT_TBL_HDR
        c.fill = NAVY_HEADER
        c.alignment = ALIGN_CENTER
        c.border = BORDER_DATA

    row_cursor = 13
    for i, reg_name in enumerate(REGIONAIS_ORDEM):
        st = reg_stats[reg_name]
        tot_r = st["total"]
        reg_r = st["regular"]
        crit_r = st["critico"]
        aler_r = st["alerta"]
        div_r = crit_r + aler_r
        taxa_r = (reg_r / tot_r) if tot_r > 0 else 1.0

        vals = [
            reg_name,
            tot_r,
            reg_r,
            crit_r,
            aler_r,
            div_r,
            taxa_r,
            round(st["proventos_jornada"], 2),
            round(st["descontos_ponto"], 2),
        ]
        fill_row = ZEBRA_FILL if i % 2 == 1 else WHITE_FILL
        ws_res.row_dimensions[row_cursor].height = 21
        for c_i, v in enumerate(vals, start=1):
            cell = ws_res.cell(row_cursor, c_i, v)
            cell.font = FONT_DATA_BOLD if c_i == 1 else FONT_DATA
            cell.fill = fill_row
            cell.border = BORDER_DATA
            if c_i == 1:
                cell.alignment = ALIGN_LEFT
            elif c_i in (2, 3, 4, 5, 6):
                cell.alignment = ALIGN_CENTER
                cell.number_format = FMT_INT
            elif c_i == 7:
                cell.alignment = ALIGN_CENTER
                cell.number_format = FMT_PCT
            else:
                cell.alignment = ALIGN_RIGHT
                cell.number_format = FMT_CURRENCY
        row_cursor += 1

    # Linha Total Tabela 1
    ws_res.row_dimensions[row_cursor].height = 24
    tot_vals = [
        "TOTAL CORPORATIVO",
        total_colabs,
        total_regulares,
        total_criticos,
        total_alertas,
        total_criticos + total_alertas,
        (total_regulares / total_colabs) if total_colabs else 0.0,
        round(total_he_bh_not, 2),
        round(total_descontos, 2),
    ]
    for c_i, v in enumerate(tot_vals, start=1):
        cell = ws_res.cell(row_cursor, c_i, v)
        cell.font = FONT_DATA_BOLD
        cell.fill = TOTAL_FILL
        cell.border = BORDER_TOTAL
        if c_i == 1:
            cell.alignment = ALIGN_LEFT
        elif c_i in (2, 3, 4, 5, 6):
            cell.alignment = ALIGN_CENTER
            cell.number_format = FMT_INT
        elif c_i == 7:
            cell.alignment = ALIGN_CENTER
            cell.number_format = FMT_PCT
        else:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = FMT_CURRENCY

    # Tabela 2: Consolidação por Regra da Malha Fina (Prioridade 10)
    r_t2_title = row_cursor + 2
    ws_res.merge_cells(start_row=r_t2_title, start_column=1, end_row=r_t2_title, end_column=9)
    ws_res.cell(r_t2_title, 1, "2. CONSOLIDAÇÃO POR REGRA DA MALHA FINA (PRIORIDADE 10 — ORDEM DE PRECEDÊNCIA)")
    ws_res.cell(r_t2_title, 1).font = FONT_SEC_HDR
    ws_res.cell(r_t2_title, 1).fill = SUB_HEADER
    ws_res.cell(r_t2_title, 1).alignment = ALIGN_LEFT
    ws_res.row_dimensions[r_t2_title].height = 25

    r_t2_hdr = r_t2_title + 1
    ws_res.row_dimensions[r_t2_hdr].height = 25
    ws_res.merge_cells(start_row=r_t2_hdr, start_column=1, end_row=r_t2_hdr, end_column=4)
    ws_res.merge_cells(start_row=r_t2_hdr, start_column=8, end_row=r_t2_hdr, end_column=9)

    t2_hdr_cells = [
        (1, "Regra de Auditoria da Malha Fina (Precedência 1 a 11)"),
        (5, "Severidade"),
        (6, "Qtd. Colaboradores"),
        (7, "% do Total"),
        (8, "Montante Financeiro Envolvido (R$)"),
    ]
    for col_i in range(1, 10):
        c = ws_res.cell(r_t2_hdr, col_i)
        c.font = FONT_TBL_HDR
        c.fill = NAVY_HEADER
        c.alignment = ALIGN_CENTER
        c.border = BORDER_DATA
    for col_i, text in t2_hdr_cells:
        ws_res.cell(r_t2_hdr, col_i, text)

    r_curr = r_t2_hdr + 1
    soma_montante_div = 0.0
    for idx, (regra_nome, sev_label) in enumerate(REGRAS_ORDEM):
        st = rule_stats[regra_nome]
        qtd = st["count"]
        pct = (qtd / total_colabs) if total_colabs else 0.0
        mont = round(st["montante"], 2)
        if sev_label in ("Crítico", "Alerta"):
            soma_montante_div += mont

        ws_res.row_dimensions[r_curr].height = 21
        ws_res.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=4)
        ws_res.merge_cells(start_row=r_curr, start_column=8, end_row=r_curr, end_column=9)

        fill_row = ZEBRA_FILL if idx % 2 == 1 else WHITE_FILL
        for col_i in range(1, 10):
            cell = ws_res.cell(r_curr, col_i)
            cell.fill = fill_row
            cell.border = BORDER_DATA
            cell.font = FONT_DATA

        c_regra = ws_res.cell(r_curr, 1, regra_nome)
        c_regra.alignment = ALIGN_LEFT
        c_regra.font = FONT_DATA_BOLD if sev_label != "Regular" else FONT_DATA

        c_sev = ws_res.cell(r_curr, 5, sev_label)
        c_sev.alignment = ALIGN_CENTER
        if sev_label == "Crítico":
            c_sev.fill = FILL_CRIT
            c_sev.font = FONT_CRIT
        elif sev_label == "Alerta":
            c_sev.fill = FILL_ALER
            c_sev.font = FONT_ALER
        else:
            c_sev.fill = FILL_REG
            c_sev.font = FONT_REG

        c_qtd = ws_res.cell(r_curr, 6, qtd)
        c_qtd.alignment = ALIGN_CENTER
        c_qtd.number_format = FMT_INT

        c_pct = ws_res.cell(r_curr, 7, pct)
        c_pct.alignment = ALIGN_CENTER
        c_pct.number_format = FMT_PCT

        c_mont = ws_res.cell(r_curr, 8, mont)
        c_mont.alignment = ALIGN_RIGHT
        c_mont.number_format = FMT_CURRENCY

        r_curr += 1

    # Linha Subtotal de Divergências (Crítico + Alerta) na Tabela 2
    ws_res.row_dimensions[r_curr].height = 24
    ws_res.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=4)
    ws_res.merge_cells(start_row=r_curr, start_column=8, end_row=r_curr, end_column=9)
    for col_i in range(1, 10):
        c = ws_res.cell(r_curr, col_i)
        c.fill = TOTAL_FILL
        c.border = BORDER_TOTAL
        c.font = FONT_DATA_BOLD

    ws_res.cell(r_curr, 1, "SUBTOTAL EM DIVERGÊNCIA / ALERTA (REGRAS 1 A 10)").alignment = ALIGN_LEFT
    ws_res.cell(r_curr, 5, "Crítico + Alerta").alignment = ALIGN_CENTER
    c_sub_q = ws_res.cell(r_curr, 6, total_criticos + total_alertas)
    c_sub_q.alignment = ALIGN_CENTER
    c_sub_q.number_format = FMT_INT
    c_sub_p = ws_res.cell(r_curr, 7, (total_criticos + total_alertas) / total_colabs if total_colabs else 0.0)
    c_sub_p.alignment = ALIGN_CENTER
    c_sub_p.number_format = FMT_PCT
    c_sub_m = ws_res.cell(r_curr, 8, round(soma_montante_div, 2))
    c_sub_m.alignment = ALIGN_RIGHT
    c_sub_m.number_format = FMT_CURRENCY

    res_widths = [24, 20, 22, 24, 20, 20, 22, 26, 24]
    for col_i, w in enumerate(res_widths, start=1):
        ws_res.column_dimensions[get_column_letter(col_i)].width = w

    # ==========================================================================
    # FUNÇÃO AUXILIAR PARA ABAS 3 (CONFERÊNCIA GERAL) E 4 (DIVERGÊNCIAS)
    # ==========================================================================
    COLS_ANALITICAS = [
        ("Regional", 16, ALIGN_CENTER, None),
        ("Loja", 10, ALIGN_CENTER, FMT_INT),
        ("Matrícula", 13, ALIGN_CENTER, FMT_INT),
        ("Nome Colaborador", 34, ALIGN_LEFT, None),
        ("Data Admissão", 15, ALIGN_CENTER, FMT_DATE),
        ("Situação", 20, ALIGN_LEFT, None),
        ("Cargo", 30, ALIGN_LEFT, None),
        ("Escala", 42, ALIGN_LEFT, None),
        ("Perfil de Frequência", 34, ALIGN_LEFT, None),
        ("Sindicato LG", 36, ALIGN_LEFT, None),
        ("Salário Base (R$)", 18, ALIGN_RIGHT, FMT_CURRENCY),
        ("Horas Extras Pagas (R$)", 20, ALIGN_RIGHT, FMT_CURRENCY),
        ("Banco de Horas Pago (R$)", 21, ALIGN_RIGHT, FMT_CURRENCY),
        ("Adicional Noturno (R$)", 20, ALIGN_RIGHT, FMT_CURRENCY),
        ("DSR Variáveis (R$)", 18, ALIGN_RIGHT, FMT_CURRENCY),
        ("Grat. Feriado / Sobreaviso (R$)", 25, ALIGN_RIGHT, FMT_CURRENCY),
        ("Desc. Faltas 5001 (R$)", 19, ALIGN_RIGHT, FMT_CURRENCY),
        ("Desc. DSR 5007 (R$)", 18, ALIGN_RIGHT, FMT_CURRENCY),
        ("Desc. Atrasos 5008 (R$)", 19, ALIGN_RIGHT, FMT_CURRENCY),
        ("Desc. Banco Horas 5111 (R$)", 22, ALIGN_RIGHT, FMT_CURRENCY),
        ("Regra Apontada Malha Fina", 48, ALIGN_LEFT, None),
        ("Severidade", 15, ALIGN_CENTER, None),
        ("Ação Recomendada RH", 58, ALIGN_LEFT, None),
    ]

    def build_analytic_sheet(ws, title_banner, dataset):
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = "A3"
        num_cols = len(COLS_ANALITICAS)
        last_col_letter = get_column_letter(num_cols)

        ws.merge_cells(f"A1:{last_col_letter}1")
        ws["A1"] = title_banner
        ws["A1"].font = FONT_TITLE
        ws["A1"].fill = NAVY_HEADER
        ws["A1"].alignment = ALIGN_LEFT
        ws.row_dimensions[1].height = 32

        ws.row_dimensions[2].height = 26
        for c_i, (col_title, col_w, _, _) in enumerate(COLS_ANALITICAS, start=1):
            cell = ws.cell(2, c_i, col_title)
            cell.font = FONT_TBL_HDR
            cell.fill = NAVY_HEADER
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_DATA
            ws.column_dimensions[get_column_letter(c_i)].width = col_w

        for idx, rec in enumerate(dataset):
            r_num = idx + 3
            row_vals = [
                rec["regional"],
                rec["loja"],
                rec["mat"],
                rec["nome"],
                rec["dt_adm"],
                rec["situacao"],
                rec["cargo"],
                rec["escala"],
                rec["perfil"],
                rec["sind_label"],
                rec["sal_base"],
                rec["he"],
                rec["bh_prov"],
                rec["adic_not"],
                rec["dsr_prov"],
                rec["grat_fer_sobreaviso"],
                rec["desc_5001"],
                rec["desc_5007"],
                rec["desc_5008"],
                rec["desc_5111"],
                rec["regra"],
                rec["sev"],
                rec["acao"],
            ]
            ws.append(row_vals)
            fill_row = ZEBRA_FILL if idx % 2 == 1 else WHITE_FILL

            for c_i, (_, _, col_align, col_fmt) in enumerate(COLS_ANALITICAS, start=1):
                cell = ws.cell(r_num, c_i)
                cell.font = FONT_DATA
                cell.fill = fill_row
                cell.alignment = col_align
                cell.border = BORDER_DATA
                if col_fmt:
                    cell.number_format = col_fmt

            # Badge de Severidade (Coluna 22)
            c_sev = ws.cell(r_num, 22)
            sev_val = rec["sev"]
            if sev_val == "Crítico":
                c_sev.fill = FILL_CRIT
                c_sev.font = FONT_CRIT
            elif sev_val == "Alerta":
                c_sev.fill = FILL_ALER
                c_sev.font = FONT_ALER
            else:
                c_sev.fill = FILL_REG
                c_sev.font = FONT_REG

        ws.auto_filter.ref = f"A2:{last_col_letter}{len(dataset) + 2}"

    # Construir Aba 3 e Aba 4
    t3 = time.time()
    print(f"   Gravando Aba 3 'Conferência Geral' ({len(records_all):,} colaboradores)...")
    build_analytic_sheet(
        ws_ger,
        f"CONFERÊNCIA GERAL — MALHA FINA DE FOLHA X PONTO (09/2026) | {len(records_all):,} COLABORADORES AUDITADOS",
        records_all,
    )
    print(f"   Aba 3 concluída em {time.time() - t3:.2f}s.")

    t4 = time.time()
    print(f"   Gravando Aba 4 'Divergências' ({len(records_div):,} ocorrências críticas/alerta)...")
    build_analytic_sheet(
        ws_div,
        f"DIVERGÊNCIAS E ALERTAS DA MALHA FINA DE FOLHA X PONTO (09/2026) | {len(records_div):,} CASOS PRIORIZADOS",
        records_div,
    )
    print(f"   Aba 4 concluída em {time.time() - t4:.2f}s.")

    # Salvar Workbook
    t5 = time.time()
    print(f"-> Salvando arquivo final em:\n   {OUT_FILE}")
    wb.save(OUT_FILE)
    file_mb = os.path.getsize(OUT_FILE) / (1024 * 1024)
    print(f"-> Planilha salva com sucesso em {time.time() - t5:.2f}s ({file_mb:.2f} MB)!")
    print("=" * 85)
    print(f"PROCESSAMENTO TOTAL CONCLUÍDO EM {time.time() - t_start:.2f} SEGUNDOS!")
    print("=" * 85)


if __name__ == "__main__":
    main()
