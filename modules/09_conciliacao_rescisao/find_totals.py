import pdfplumber

pdf_path = r"C:\Users\Matheus\.gemini\antigravity\brain\ae1f34f0-dbd0-4654-a5ee-8c025ff614ce\.user_uploaded\media_1788529355685.pdf"

with pdfplumber.open(pdf_path) as pdf:
    p2 = pdf.pages[1]
    words = [w for w in p2.extract_words() if 'TOTAL' in w['text'].upper() or '5.662' in w['text'] or '9.673' in w['text']]
    for w in words:
        print(f"y={w['top']:.1f}, x0={w['x0']:.1f}, text='{w['text']}'")
