import pdfplumber
import re
import pandas as pd

def parse_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        # Extract header from page 1
        p1 = pdf.pages[0]
        text_p1 = p1.extract_text()
        
        # Header info
        empresa_m = re.search(r"(\d{3}-CENCOSUD[^\n]+)", text_p1)
        empresa = empresa_m.group(1).strip() if empresa_m else ""
        
        mes_m = re.search(r"Mês:\s*([^\n]+)", text_p1, re.IGNORECASE)
        mes = mes_m.group(1).strip() if mes_m else ""
        
        func_m = re.search(r"FUNCIONÁRIO\s+ADMISSÃO\s+SITUAÇÃO\s+RESCISÃO\s*\n(\d+)\s*-\s*([^\n\d]+?)\s+(\d{2}/\d{2}/\d{4})\s+([^\n\d]+?)\s+(\d{2}/\d{2}/\d{4})", text_p1)
        if not func_m:
            # Fallback regex
            func_m = re.search(r"(\d{6,9})\s*-\s*([A-Z\s]+?)\s+(\d{2}/\d{2}/\d{4})\s+([A-Za-z0-9/\s]+?)\s+(\d{2}/\d{2}/\d{4})", text_p1)
            
        if func_m:
            matricula = func_m.group(1).lstrip('0')
            colaborador = func_m.group(2).strip()
            admissao = func_m.group(3).strip()
            situacao = func_m.group(4).strip()
            rescisao = func_m.group(5).strip()
        else:
            matricula, colaborador, admissao, situacao, rescisao = "", "", "", "", ""
            
        lotacao_m = re.search(r"LOTAÇÃO TRIBUTÁRIA[^\n]*\n([^\n]+?)\s{2,}", text_p1)
        lotacao = lotacao_m.group(1).strip() if lotacao_m else ""
        
        cc_m = re.search(r"CENTRO DE CUSTO\s+CARGO\s*\n[^\n]*?\s{2,}([A-Z0-9]+ - [^\n]+?)\s{2,}([A-Z0-9/\- ]+)", text_p1)
        if cc_m:
            centro_custo = cc_m.group(1).strip()
            cargo = cc_m.group(2).strip()
        else:
            centro_custo, cargo = "", ""
            
        pagto_m = re.search(r"DATA DE PAGAMENTO\s*\n(\d{2}/\d{2}/\d{4})", text_p1)
        data_pagamento = pagto_m.group(1).strip() if pagto_m else ""
        
        header_info = {
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
            'FOLHA': "Rescisao",
            'PERÍODO DA FOLHA DE PAGAMENTO': mes,
            'DATA DE PAGAMENTO': data_pagamento
        }
        
        print("--- HEADER EXTRACTED ---")
        for k, v in header_info.items():
            print(f"  {k}: {v}")
            
        # Extract events across all pages
        events = []
        for page_num, page in enumerate(pdf.pages):
            words = page.extract_words()
            # Filter table area
            # Y range: 128 to 565 on p1, 128 to 285 on p2
            max_y = 565 if page_num == 0 else 285
            table_words = [w for w in words if 128 <= w['top'] <= max_y]
            
            # Group words by row (similar top)
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
                # Partition into Proventos (<266), Descontos (266..510), Resultados (>510)
                prov_w = [w for w in r_words if w['x0'] < 266]
                desc_w = [w for w in r_words if 266 <= w['x0'] < 510]
                res_w = [w for w in r_words if w['x0'] >= 510]
                
                def parse_cell(cell_words, tipo):
                    if not cell_words:
                        return None
                    # First word usually has code or code+desc
                    text_line = " ".join([w['text'] for w in cell_words])
                    # Pattern: COD (5 digits) DESC REF VALOR or COD DESC VALOR
                    # Example proventos: '00005 Adic Insalubridade 194,52' or '00060 DSR Horas Extras 5,00 10,28'
                    # Example resultados: '10100Base Inss Mes 1.393,82'
                    
                    # Split code at start
                    m = re.match(r"^(\d{5})\s*(.*)$", text_line)
                    if not m:
                        return None
                    cod = int(m.group(1))
                    rest = m.group(2).strip()
                    
                    # Last token is value
                    tokens = rest.split()
                    if not tokens:
                        return None
                    val_str = tokens[-1]
                    # Check if val_str is numeric like 1.234,56 or 123,45
                    if not re.match(r"^-?\d{1,3}(\.\d{3})*,\d{2}$|^-?\d+,\d{2}$", val_str):
                        return None
                    val = float(val_str.replace('.', '').replace(',', '.'))
                    
                    # If tipo is provento/desconto, there might be a REF before val
                    desc_tokens = tokens[:-1]
                    ref = None
                    if tipo in ['PROVENTO', 'DESCONTO'] and desc_tokens:
                        # Check if last token of desc_tokens is a ref (e.g. '5,00', '18,00', '1,95')
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
                
                p_item = parse_cell(prov_w, 'PROVENTO')
                if p_item: events.append(p_item)
                d_item = parse_cell(desc_w, 'DESCONTO')
                if d_item: events.append(d_item)
                r_item = parse_cell(res_w, 'RESULTADO')
                if r_item: events.append(r_item)

        print(f"\n--- TOTAL EVENTS PARSED: {len(events)} ---")
        df_events = pd.DataFrame(events)
        print(df_events.head(15))
        print(df_events.tail(15))
        return header_info, events

if __name__ == '__main__':
    pdf_path = r"C:\Users\Matheus\.gemini\antigravity\brain\ae1f34f0-dbd0-4654-a5ee-8c025ff614ce\.user_uploaded\media_1788529355685.pdf"
    parse_pdf(pdf_path)
