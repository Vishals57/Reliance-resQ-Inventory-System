import qrcode
import os

def generate_article_qr(art_no, name):
    # We store the Article No in the QR. 
    # When scanned, the system pulls Name, CP, and SP from the Excel Master sheet.
    data = f"ART_NO:{art_no}"
    
    # Create folder for QRs if it doesn't exist
    if not os.path.exists("Article_QRs"):
        os.makedirs("Article_QRs")
        
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(data)
    qr.make(fit=True)
    
    img = qr.make_image(fill_color="black", back_color="white")
    file_path = f"Article_QRs/QR_{art_no}.png"
    img.save(file_path)
    print(f"🖨️ QR Generated and saved at: {file_path}")