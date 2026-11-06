import pdfplumber
import re
import os
import glob
import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta

def extract_pdf_data(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        p1 = pdf.pages[0]
        words_p1 = p1.extract_words()
        
        # Organize words by line
        rows_p1 = []
        for w in sorted(words_p1, key=lambda x: x['top']):
            matched = False
            for r in rows_p1:
                if abs(r['top'] - w['top']) < 3.0:
                    r['words'].append(w)
                    matched = True
                    break
            if not matched:
                rows_p1.append({'top': w['top'], 'words': [w]})
                
        def get_line_text(words):
            return " ".join([w['text'] for w in sorted(words, key=lambda x: x['x0'])])

        # Header fields
        empresa = ""
        mes = ""
        folha = "Rescisao"
        matricula = ""
        colaborador = ""
        admissao = ""
        situacao = ""
        rescisao = ""
        pagamento = ""
        lotacao = ""
        centro_custo = ""
        cargo = ""
        
        for r in rows_p1:
            line = get_line_text(r['words'])
            y = r['top']
            
            if "CENCOSUD BRASIL COMERCIAL LTDA." in line and not empresa:
                empresa = line
            elif "Mês:" in line or "Ms:" in line:
                m = re.search(r"M[ê]s:\s*([^\s]+)", line)
                if m: mes = m.group(1)
            elif 65 <= y <= 75 and "-" in line and ("202" in line or "201" in line):
                # Matricula, nome, admissao, situacao, rescisao
                # Word-based partitioning
                w_mat = [w['text'] for w in r['words'] if w['x0'] < 350]
                w_adm = [w['text'] for w in r['words'] if 350 <= w['x0'] < 420]
                w_sit = [w['text'] for w in r['words'] if 420 <= w['x0'] < 600]
                w_res = [w['text'] for w in r['words'] if w['x0'] >= 600]
                
                full_mat = " ".join(w_mat)
                m_mat = re.match(r"^(\d+)\s*-\s*(.+)$", full_mat)
                if m_mat:
                    matricula = m_mat.group(1).lstrip('0')
                    colaborador = m_mat.group(2).strip()
                admissao = " ".join(w_adm).strip()
                situacao = " ".join(w_sit).strip()
                rescisao = " ".join(w_res).strip()
            elif 80 <= y <= 92:
                # Pagamento e Folha
                w_pag = [w['text'] for w in r['words'] if w['x0'] < 200]
                w_fol = [w['text'] for w in r['words'] if w['x0'] >= 600]
                if w_pag: pagamento = " ".join(w_pag).strip()
                if w_fol: folha = " ".join(w_fol).strip()
            elif 98 <= y <= 112:
                # Lotacao, centro de custo, cargo
                w_lot = [w['text'] for w in r['words'] if w['x0'] < 275]
                w_cc = [w['text'] for w in r['words'] if 275 <= w['x0'] < 530]
                w_car = [w['text'] for w in r['words'] if w['x0'] >= 530]
                lotacao = " ".join(w_lot).strip()
                centro_custo = " ".join(w_cc).strip()
                cargo = " ".join(w_car).strip()

        header = {
            'EMPRESA': empresa,
            'LOTAÇÃO TRIBUTÁRIA': lotacao,
            'MATRICULA COLABORADOR': matricula,
            'COLABORADOR': colaborador,
            'DATA DE ADMISSÃO': admissao,
            'SITUAÇÃO': situacao,
            'DATA INÍCIO SITUAÇÃO': rescisao,
            'CENTRO DE CUSTOS': centro_custo,
            'UNIDADE ORGANIZACIONAL': "",
            'CARGO': cargo,
            'CODIGO FOLHA': 5,
            'FOLHA': folha,
            'PERÍODO DA FOLHA DE PAGAMENTO': mes,
            'DATA DE PAGAMENTO': pagamento
        }
        
        # Events across pages
        events = []
        totais = {}
        
        for page_num, page in enumerate(pdf.pages):
            words = page.extract_words()
            max_y = 565 if page_num == 0 else 285
            table_words = [w for w in words if 128 <= w['top'] <= max_y]
            
            # Group by row
            rows = []
            for w in sorted(table_words, key=lambda x: x['top']):
                matched = False
                for r in rows:
                    if abs(r['top'] - w['top']) < 3.5:
                        r['words'].append(w)
                        matched = True
                        break
                if not matched:
                    rows.append({'top': w['top'], 'words': [w]})
                    
            for r in rows:
                r_words = sorted(r['words'], key=lambda x: x['x0'])
                prov_w = [w for w in r_words if w['x0'] < 266]
                desc_w = [w for w in r_words if 266 <= w['x0'] < 510]
                res_w = [w for w in r_words if w['x0'] >= 510]
                
                def parse_cell(cell_words, tipo):
                    if not cell_words: return None
                    line_txt = " ".join([w['text'] for w in cell_words])
                    
                    # Consertar espaçamentos incorretos dentro de números, ex: "18,0 0" -> "18,00"
                    line_txt = re.sub(r"(\d,\d+)\s+(\d+)", r"\1\2", line_txt)
                    
                    m = re.match(r"^(\d{5})\s*(.*)$", line_txt)
                    if not m: return None
                    cod = int(m.group(1))
                    rest = m.group(2).strip()
                    tokens = rest.split()
                    if not tokens: return None
                    val_str = tokens[-1]
                    if not re.match(r"^-?\d{1,3}(\.\d{3})*,\d{2}$|^-?\d+,\d{2}$", val_str):
                        return None
                    val = float(val_str.replace('.', '').replace(',', '.'))
                    
                    desc_tokens = tokens[:-1]
                    ref = np.nan
                    if tipo in ['PROVENTO', 'DESCONTO'] and desc_tokens:
                        if re.match(r"^\d+,\d{2}$|^\d+$", desc_tokens[-1]):
                            ref = float(desc_tokens[-1].replace(',', '.'))
                            desc_tokens = desc_tokens[:-1]
                            
                    desc = " ".join(desc_tokens).strip()
                    return {
                        'TIPO DE EVENTO': tipo,
                        'CÓDIGO DO EVENTO': cod,
                        'DESCRIÇÃO DO EVENTO': desc,
                        'VALOR EVENTO': val,
                        'VALOR EVENTO REFERÊNCIA': ref
                    }

                p = parse_cell(prov_w, 'PROVENTO')
                if p: events.append(p)
                d = parse_cell(desc_w, 'DESCONTO')
                if d: events.append(d)
                r_item = parse_cell(res_w, 'RESULTADO')
                if r_item: events.append(r_item)

            # Check for totals on page 2
            if page_num == len(pdf.pages) - 1:
                tot_words = [w for w in words if 285 <= w['top'] <= 305]
                # Look for numbers
                for tw in tot_words:
                    txt = tw['text'].replace('.', '').replace(',', '.')
                    if re.match(r"^\d+\.\d{2}$", txt):
                        val = float(txt)
                        if tw['x0'] < 260:
                            totais['TOTAL_PROVENTOS'] = val
                        elif tw['x0'] < 500:
                            totais['TOTAL_DESCONTOS'] = val
                        else:
                            totais['TOTAL_LIQUIDO'] = val

        return header, events, totais

if __name__ == '__main__':
    pdf_path = r"C:\Users\Matheus\.gemini\antigravity\brain\ae1f34f0-dbd0-4654-a5ee-8c025ff614ce\.user_uploaded\media_1788529355685.pdf"
    h, ev, tot = extract_pdf_data(pdf_path)
    print("Header:", h)
    print("Total events:", len(ev))
    print("Totais:", tot)
