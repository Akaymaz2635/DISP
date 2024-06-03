from docx import Document

# Word belgesinin dosya yolunu belirtin
doc_path = "C:/Users/alika/Desktop/Yeni Microsoft Word Belgesi.docx"

# Word belgesini aç
doc = Document(doc_path)

# Tüm tabloları al
tables = doc.tables

# Her bir tabloyu döngüyle kontrol et
for table in tables:
    # Tablodaki her bir satırı döngüyle kontrol et
    for row in table.rows:
        # Satırın 4. sütunu "ATOS" ise
        if row.cells[4].text.strip() == "ATOS":
            # Satırın 2. sütununu güncelle
            row.cells[2].text = "[" + row.cells[0].text + "]"
        
        # Satırın metni içinde "S/N:" varsa
        if "S/N:" in row.cells[4].text and row.cells[5].text =="":
            # "S/N:"'i "S/N:[Element]" olarak güncelle
            row.cells[5].text = row.cells[5].text.replace("", "[Element]")
        else:
            print("İlgili kolonun boş olduğundan emin olun.")        
        # Satırın metni içinde "DATE:" varsa
        if "DATE:" in row.cells[6].text and row.cells[7].text =="":
            # "DATE:"'i "DATE:[Date]" olarak güncelle
            row.cells[7].text = row.cells[7].text.replace("", "[Date]")
        else:
            print("İlgili kolonun boş olduğundan emin olun.")
        

# Dosyayı kaydet
doc.save(doc_path)

print("Word belgesi başarıyla güncellendi.")
