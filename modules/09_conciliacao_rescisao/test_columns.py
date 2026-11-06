import pdfplumber

pdf_path = r"C:\Users\Matheus\.gemini\antigravity\brain\ae1f34f0-dbd0-4654-a5ee-8c025ff614ce\.user_uploaded\media_1788529355685.pdf"

with pdfplumber.open(pdf_path) as pdf:
    p1 = pdf.pages[0]
    
    # Let's crop into the 3 columns
    # Table top starts around y = 110 to 125, bottom is around y = 560
    # Let's check words between top 110 and 130
    words_header = [w for w in p1.extract_words() if 110 <= w['top'] <= 130]
    for w in words_header:
        print(f"{w['text']} at x0={w['x0']:.1f}, x1={w['x1']:.1f}, top={w['top']:.1f}")

    print("\n--- Column Boundaries Exploration ---")
    # Proventos is column 1
    # Descontos is column 2
    # Resultados is column 3
