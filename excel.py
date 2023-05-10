import openpyxl

def indirimsiz_fiyat_hesapla(workbook_path, sheet_name, fiyat_column):
    # Excel dosyasını yükle
    workbook = openpyxl.load_workbook(workbook_path)
    
    # Çalışma sayfasını seç
    sheet = workbook[sheet_name]
   
    
    # Fiyatları indirimsiz hale getir
    for row in range(1 , 999):
        
        fiyat = sheet["I" + str(row)].value
        
        
       
        if fiyat and type(fiyat) != str:
            indirimsiz_fiyat = fiyat / 0.7  # %30 indirimli fiyatı hesapla

            
            # İndirimsiz fiyatı yaz
            #sheet.cell(row=row[0].row, column=fiyat_column + 1, value=indirimsiz_fiyat)
            print("I" + str(row))
            sheet["I" + str(row)] = indirimsiz_fiyat
        
    # Değişiklikleri kaydet
    workbook.save(workbook_path)

# Kullanım örneği
workbook_path = "satis1.xlsx"  # Excel dosyasının yolu
sheet_name = "Kazakistan satis faturalari"  # Çalışma sayfasının adı
fiyat_column = 9  # Fiyatların bulunduğu sütun numarası (örneğin B sütunu için 2)

indirimsiz_fiyat_hesapla(workbook_path, sheet_name, fiyat_column)
