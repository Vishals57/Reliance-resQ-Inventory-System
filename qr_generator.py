import qrcode
import cv2
from pyzbar.pyzbar import decode
import datetime
import pandas as pd

# 1. FUNCTION TO GENERATE QR CODE
def generate_part_qr(part_id, part_name):
    # Data to store inside the QR
    data = f"ID:{part_id}|Name:{part_name}"
    qr = qrcode.make(data)
    file_name = f"QR_{part_id}.png"
    qr.save(file_name)
    print(f"✅ QR Code saved as {file_name}")

# 2. FUNCTION TO SCAN AND UPDATE (Simulated In/Out)
def scan_inventory(action="IN"):
    cap = cv2.VideoCapture(0) # Opens your webcam
    print(f"📷 Scanning for Part {action}... Press 'q' to stop.")
    
    while True:
        ret, frame = cap.read()
        for barcode in decode(frame):
            data = barcode.data.decode('utf-8')
            print(f"✨ Scanned Data: {data}")
            
            # Log the movement
            log_entry = {
                "Timestamp": datetime.datetime.now(),
                "Data": data,
                "Type": action
            }
            print(f"📝 Recorded: {log_entry}")
            
            cap.release()
            cv2.destroyAllWindows()
            return log_entry
            
        cv2.imshow('Inventory Scanner', frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    cap.release()
    cv2.destroyAllWindows()

# --- EXAMPLE USAGE ---
# generate_part_qr("RESQ-99", "Samsung-Display-Panel")
# scan_inventory(action="OUT")