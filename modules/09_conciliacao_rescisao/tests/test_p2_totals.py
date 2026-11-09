import pdfplumber

pdf_path = r"C:\Users\Matheus\.gemini\antigravity\brain\ae1f34f0-dbd0-4654-a5ee-8c025ff614ce\.user_uploaded\media_1788529355685.pdf"

with pdfplumber.open(pdf_path) as pdf:
    p2 = pdf.pages[1]
    words = [w for w in p2.extract_words() if 400 <= w['top'] <= 565]
    for w in sorted(words, key=lambda x: (x['top'], x['x0'])):
        print(f"y={w['top']:.1f}, x0={w['x0']:.1f}, text='{w['text']}'")
