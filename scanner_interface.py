import cv2
from pyzbar.pyzbar import decode

def activate_scanner():
    cap = cv2.VideoCapture(0)
    print("📷 Scanner Active. Align QR code. Press 'q' to cancel.")
    
    while True:
        ret, frame = cap.read()
        if not ret:
            continue

        for barcode in decode(frame):
            # 1. Get raw text from QR
            raw_data = barcode.data.decode('utf-8').strip()
            print(f"DEBUG: Scanned raw text -> '{raw_data}'") 
            
            cap.release()
            cv2.destroyAllWindows()
            
            # 2. Logic to clean the ID
            try:
                if "ART_NO:" in raw_data:
                    # If QR is 'ART_NO:001', this gives us '001'
                    clean_id = raw_data.split(":")[1].strip()
                    print(f"✅ Cleaned Article ID: {clean_id}")
                    return clean_id, "Linked_Part"
                
                elif "|" in raw_data:
                    # Logic for old format 'ID:123|Name:panel'
                    parts = raw_data.split("|")
                    clean_id = parts[0].split(":")[1].strip()
                    return clean_id, "Linked_Part"
                
                else:
                    # If QR is just raw text like '001'
                    return raw_data, "Direct_ID"
                    
            except Exception as e:
                print(f"❌ Error decoding QR format: {e}")
                return None, None
            
        cv2.imshow('resQ Smart Scanner', frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
            
    cap.release()
    cv2.destroyAllWindows()
    return None, None