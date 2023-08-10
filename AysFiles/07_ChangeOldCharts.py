from pptx import Presentation
import openpyxl

# PowerPoint dosyasını yükle
presentation = Presentation('Deneme.pptx')

# Excel dosyasından yeni verileri yükle
wb = openpyxl.load_workbook('Deneme.xlsx')
sheet = wb.active
new_categories = [cell.value for cell in sheet['A'][1:]]
new_values = [cell.value for cell in sheet['B'][1:]]

# Hedef grafik başlığı
target_chart_title = "Second"

# Her slayttaki şekilleri kontrol et
for slide in presentation.slides:
    for shape in slide.shapes:
        if shape.has_chart:
            chart = shape.chart
            chart_title = chart.chart_title.text_frame.text
            
            # Belirli bir başlığa sahip grafikse
            if chart_title == target_chart_title:
                chart_data = chart.chart_data
                chart_data.categories = new_categories
                
                # İlk seriyi al (diğer serileri istediğiniz gibi değiştirebilirsiniz)
                series = chart_data.series[0]
                series.values = new_values
                
                # Güncellenmiş verilerle grafik verilerini değiştir
                chart.replace_data(chart_data)
                print(f"Updated data for the chart with title: {target_chart_title} on slide {slide.slide_id}")

# Güncellenmiş PowerPoint dosyasını kaydet
presentation.save('updated_presentation.pptx')
