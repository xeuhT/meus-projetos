import pdfplumber

pdf_path = r"C:\Users\Matheus\.gemini\antigravity\brain\ae1f34f0-dbd0-4654-a5ee-8c025ff614ce\.user_uploaded\media_1788529355685.pdf"

with pdfplumber.open(pdf_path) as pdf:
    p1 = pdf.pages[0]
    # words around y=120..150
    words = [w for w in p1.extract_words() if 120 <= w['top'] <= 160]
    for w in sorted(words, key=lambda x: (x['top'], x['x0'])):
        print(f"y={w['top']:.1f}, x0={w['x0']:.1f}, text='{w['text']}'")
