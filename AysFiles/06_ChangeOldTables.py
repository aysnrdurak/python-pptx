from pptx import Presentation

# PowerPoint dosyasını yükle
presentation = Presentation('Cam_details_template_(Draft_V01).pptx')

# Tablo sayacını başlangıçta sıfıra ayarla
table_count = 0

# Her slayttaki şekilleri kontrol et
for slide in presentation.slides:
    for shape in slide.shapes:
        if shape.has_table:
            
            table = shape.table
            # Sadece 1. tabloyu kontrol et
            if table_count == 1:
                if table and table.cell(0, 0).text:
                    print("{} tablonun 1. hücresi değeri: {}".format(table_count, table.cell(0, 0).text))
                    table.cell(0, 0).text = "Yeni Değer"  # 1. hücreye yeni değeri atar
                    print("{} tablonun 1. hücresi değeri: {}".format(table_count, table.cell(0, 0).text))
                table_count += 1
                break  # 1. tabloyu bulduktan sonra diğer tablolara bakmamak için döngüyü kırar
            else: 
                table_count +=1

# Sunuyu kaydet
presentation.save("06_ChangeOldTables.pptx")
