import pdfplumber

pdf_path = r"C:\Users\Matheus\.gemini\antigravity\brain\ae1f34f0-dbd0-4654-a5ee-8c025ff614ce\.user_uploaded\media_1788529355685.pdf"

with pdfplumber.open(pdf_path) as pdf:
    p1 = pdf.pages[0]
    words = [w for w in p1.extract_words() if w['top'] < 120]
    # Group by line
    rows = []
    for w in sorted(words, key=lambda x: x['top']):
        matched = False
        for r in rows:
            if abs(r['top'] - w['top']) < 3.0:
                r['words'].append(w)
                matched = True
                break
        if not matched:
            rows.append({'top': w['top'], 'words': [w]})
            
    for r in sorted(rows, key=lambda x: x['top']):
        line_str = " | ".join([f"{w['text']} (x={w['x0']:.1f})" for w in sorted(r['words'], key=lambda x: x['x0'])])
        print(f"y={r['top']:.1f}: {line_str}")
