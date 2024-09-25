import pandas as pd
import barcode
from barcode.writer import ImageWriter
import os
import fitz  # PyMuPDF
from barcode import Code39
# ตั้งค่ามั่วไป

# 1. ไฟล์ Excel เก็บ barcode
df = pd.read_excel('barcodexcel.xlsx') 

# 2. ไฟล์ pdf ต้นฉบับที่จะเขียน barcode
original_pdf_path = 'inputpdf.pdf'
a4_width = 595  # A4 width in points
a4_height = 842  # A4 height in points

# 3. ความกว้าง barcode 
barcode_width = 100  

# 4. ระยะห่าง barcode จากขอบ
margin_bottom = 0 

barcode_dir = "barcodes"
if not os.path.exists(barcode_dir):
    os.makedirs(barcode_dir)

data_column = df['Data Column'].astype(str)  

pdf_document = fitz.open(original_pdf_path)

num_pages = len(pdf_document)
if len(data_column) > num_pages:
    print("Warning: More data entries than pages in the PDF. Extra entries will be ignored.")

for page_number in range(num_pages):
    pdf_page = pdf_document[page_number]

    if page_number < len(data_column):
        code = data_column[page_number]

        # Sanitize the code for Code 39
        sanitized_code = "".join(c for c in code if c.isalnum() or c in '-. $/+%')  # Allow specific special characters

        # Generate Code 39 barcode
        my_barcode = Code39(sanitized_code , writer=ImageWriter(), add_checksum=False)
        
        barcode_image_path = os.path.join(barcode_dir, f"{sanitized_code}.png")
        
        my_barcode.save(barcode_image_path)

        center_x = (a4_width - barcode_width) / 2 
        barcode_y = pdf_page.rect.height - margin_bottom - 100  

        pdf_page.insert_image(fitz.Rect(center_x, barcode_y, center_x + barcode_width, barcode_y + 100), 
                               filename=barcode_image_path+".png")

        text_x = (a4_width - len(code) * 6) / 2  
        text_y = barcode_y + 110  

# Save the updated PDF
pdf_document.save('output_with_barcodes.pdf')
pdf_document.close()

print("PDF สร้างเสร็จละ")
