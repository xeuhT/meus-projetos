import pdfplumber

pdf_path = r"C:\Users\Matheus\.gemini\antigravity\brain\ae1f34f0-dbd0-4654-a5ee-8c025ff614ce\.user_uploaded\media_1788529355685.pdf"

with pdfplumber.open(pdf_path) as pdf:
    for idx, p in enumerate(pdf.pages):
        print(f"\n--- PAGE {idx+1} bottom words (y > 450) ---")
        words = [w for w in p.extract_words() if w['top'] > 450]
        for w in sorted(words, key=lambda x: (x['top'], x['x0']))[:30]:
            print(f"y={w['top']:.1f}, x0={w['x0']:.1f}, text='{w['text']}'")
